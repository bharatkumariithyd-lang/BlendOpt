"""Run the whole job: build -> solve -> (diagnose if infeasible) -> report.

This is what ``run.py`` calls, and the easiest entry point from a notebook::

    from blend_optimizer.inputs import load_inputs
    from blend_optimizer.pipeline import optimise
    outcome = optimise(load_inputs("my_product.xlsx"), "output/my_product.xlsx")

When the problem is infeasible, :func:`diagnose_infeasible` runs
automatically: pre-checks (``diagnose``), the elastic model (``model`` +
``solve_and_report``), the reduction of the conflicting limits to those
really needed, and a search for the best achievable value of each of them
(both here, by re-solving).
"""

from dataclasses import dataclass, replace
from pathlib import Path

import pandas as pd

from blend_optimizer.data import BlendData, UsageLimit
from blend_optimizer.diagnose import Conflict, Diagnosis, build_diagnosis, conflict_set, format_diagnosis, precheck
from blend_optimizer.linear_rows import MIN, LinearRow, build_linear_rows
from blend_optimizer.model import build_model
from blend_optimizer.solve_and_report import (
    Section,
    SolveResult,
    build_report,
    print_summary,
    report_notes,
    report_row_decimals,
    solve,
    write_excel,
)

SHEET_DIAGNOSIS = "Diagnosis"
# Statuses that mean "no blend meets the limits" (HiGHS may report either)
INFEASIBLE_STATUSES = ("Infeasible", "Undefined")
# A Max target this large never restricts anything (used as the "loose" end of a search)
_LOOSE_MAX = 1e6


@dataclass
class RunOutcome:
    """Everything one run produced.

    Attributes:
        data: The input data.
        result: The solver result.
        tables: Report tables, sheet name -> table (full precision).
        excel_path: Report workbook written, or None.
        lp_path: .lp file written, or None.
        diagnosis: Explanation of an infeasible problem, or None.
    """

    data: BlendData
    result: SolveResult
    tables: dict[str, pd.DataFrame]
    excel_path: Path | None = None
    lp_path: Path | None = None
    diagnosis: Diagnosis | None = None


def solve_blend(data: BlendData, lp_path: str | Path | None = None) -> SolveResult:
    """Build and solve the LP for ``data``, without any report.

    Used by ``what_if`` and by the diagnosis for their many re-solves.
    """
    return solve(build_model(data), lp_path=lp_path)


# ===========================================================================
# Diagnosis of infeasible problems
# ===========================================================================

def _with_target(data: BlendData, row: LinearRow, target: float) -> BlendData:
    """Copy of ``data`` with one spec bound moved so that its internal target is ``target``.

    The margin is kept, so the customer limit becomes target - margin (Min)
    or target + margin (Max). The other bound of the same spec is unchanged.
    """
    specs = []
    for spec in data.specs:
        if spec.name == row.spec.name and row.bound == MIN:
            spec = replace(spec, min_value=target - spec.margin)
        elif spec.name == row.spec.name:
            spec = replace(spec, max_value=target + spec.margin)
        specs.append(spec)
    return replace(data, specs=specs)


def best_achievable(data: BlendData, row: LinearRow, tolerance: float = 1e-6) -> float | None:
    """Best value of one limit's quantity while every other limit is held.

    For a Min limit this is the highest target that still leaves a feasible
    blend; for a Max limit, the lowest. Found by bisection on the target:
    each trial is one solve, keeping one end feasible and the other not.

    * The current target is infeasible (that is why we are here).
    * A target that restricts nothing (0 for a Min, since blend values are
      never negative; a very large value for a Max) is feasible, unless the
      other limits conflict among themselves; then None is returned.

    Args:
        data: The (infeasible) input data.
        row: The spec bound to search on.
        tolerance: Accuracy, relative to the size of the limit.
    """
    loose = 0.0 if row.bound == MIN else _LOOSE_MAX
    if not solve_blend(_with_target(data, row, loose)).is_optimal:
        return None
    feasible, infeasible = loose, row.limit
    while abs(infeasible - feasible) > tolerance * max(1.0, abs(row.limit)):
        middle = (feasible + infeasible) / 2.0
        if solve_blend(_with_target(data, row, middle)).is_optimal:
            feasible = middle
        else:
            infeasible = middle
    return feasible


def _without_usage_limit(data: BlendData, material: str, side: str) -> BlendData:
    """Copy of ``data`` with one usage limit lifted: Max back to 100 %, or Min back to 0 %.

    Both the plant-wide limit and the product limit are lifted, so the
    effective limit on that side is gone.
    """
    def lifted(limit: UsageLimit) -> UsageLimit:
        return UsageLimit(limit.min_pct, 100.0) if side == "max" else UsageLimit(0.0, limit.max_pct)

    raw = dict(data.raw_materials)
    raw[material] = replace(raw[material], plant_limit=lifted(raw[material].plant_limit))
    product_limits = dict(data.product_limits)
    if material in product_limits:
        product_limits[material] = lifted(product_limits[material])
    return replace(data, raw_materials=raw, product_limits=product_limits)


def minimal_conflict(data: BlendData, conflict: Conflict) -> Conflict:
    """Keep only the limits the conflict really needs.

    The elastic model can name limits that push against the answer without
    being part of the conflict. Each limit in the set is tested in turn by
    removing it (a spec bound made unrestrictive, a usage limit lifted): if
    the problem is still infeasible without it, it is not needed and is
    dropped for good. Every limit that is kept is one whose removal makes a
    blend possible. (This is the standard "deletion filter".)

    Args:
        data: The infeasible input data.
        conflict: From ``diagnose.conflict_set``.

    Returns:
        The reduced set. If dropping went wrong (nothing would be left), the
        original set is returned.
    """
    rows = {row.name: row for row in build_linear_rows(data)}
    trial = data                         # the data with every limit dropped so far removed
    kept_rows: list[str] = []
    kept_usage: list[tuple[str, str, float]] = []
    for name in conflict.spec_rows:
        row = rows[name]
        loose = 0.0 if row.bound == MIN else _LOOSE_MAX    # a target that restricts nothing
        candidate = _with_target(trial, row, loose)
        if solve_blend(candidate).is_optimal:
            kept_rows.append(name)       # needed: without it a blend exists
        else:
            trial = candidate            # not needed: drop it for good
    for usage in conflict.usage_limits:
        material, side, _ = usage
        candidate = _without_usage_limit(trial, material, side)
        if solve_blend(candidate).is_optimal:
            kept_usage.append(usage)
        else:
            trial = candidate
    if not kept_rows and not kept_usage:
        return conflict
    return Conflict(kept_rows, kept_usage)


def diagnose_infeasible(data: BlendData, lp_path: str | Path | None = None) -> Diagnosis:
    """Explain why no blend meets the limits.

    1. Pre-checks: each limit on its own (``diagnose.precheck``).
    2. Elastic model: which limits conflict (``diagnose.conflict_set``),
       reduced to the limits really needed (:func:`minimal_conflict`).
    3. For each remaining spec limit: the best value achievable while all
       other limits are held (:func:`best_achievable`).

    Args:
        data: The infeasible input data.
        lp_path: Optional path for the elastic model's .lp file.
    """
    findings = precheck(data)
    elastic = solve(build_model(data, elastic=True), lp_path=lp_path)
    conflict = conflict_set(data, elastic) if elastic.is_optimal else Conflict()
    conflict = minimal_conflict(data, conflict)
    achievable = {name: best_achievable(data, elastic.rows[name]) for name in conflict.spec_rows}
    return build_diagnosis(data, findings, elastic, conflict, achievable)


# ===========================================================================
# The whole run
# ===========================================================================

def optimise(
    data: BlendData,
    output_path: str | Path | None = None,
    show: bool = True,
    extra_sheets: dict[str, pd.DataFrame | list[Section]] | None = None,
    extra_notes: dict[str, list[str]] | None = None,
) -> RunOutcome:
    """Find the least-cost blend and report it; diagnose it if it is infeasible.

    Args:
        data: Validated input data (from ``inputs.load_inputs``).
        output_path: Report workbook to write (.xlsx). The .lp file is written
            next to it with the same name (and ``<name>_elastic.lp`` for the
            elastic model of an infeasible problem). None writes no files.
        show: Print the console summary (and the diagnosis, if any).
        extra_sheets: More sheets to add to the workbook after the report
            (``run.py`` adds the price ranges and what-if scenario sheets here).
        extra_notes: Notes for single-table extra sheets, sheet name -> lines.

    Returns:
        The outcome, including the report tables and any diagnosis.
    """
    excel_path = Path(output_path) if output_path is not None else None
    lp_path = excel_path.with_suffix(".lp") if excel_path is not None else None

    result = solve_blend(data, lp_path=lp_path)
    tables = build_report(data, result)
    notes = report_notes(data, result)

    diagnosis = None
    if not result.is_optimal and result.status in INFEASIBLE_STATUSES:
        elastic_lp = lp_path.with_name(lp_path.stem + "_elastic.lp") if lp_path is not None else None
        diagnosis = diagnose_infeasible(data, lp_path=elastic_lp)
        tables[SHEET_DIAGNOSIS] = diagnosis.table
        notes[SHEET_DIAGNOSIS] = diagnosis.messages

    if excel_path is not None:
        write_excel({**tables, **(extra_sheets or {})}, excel_path, notes={**notes, **(extra_notes or {})},
                    scale_resolution_kg=data.settings.scale_resolution_kg,
                    row_decimals=report_row_decimals(data, result))
    if show:
        print_summary(data, result)
        if diagnosis is not None:
            print()
            print(format_diagnosis(diagnosis))
    return RunOutcome(data=data, result=result, tables=tables, excel_path=excel_path, lp_path=lp_path,
                      diagnosis=diagnosis)

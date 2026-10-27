"""Explain, in plain language, why no blend meets the limits (an infeasible LP).

Two kinds of finding:

1. **Pre-checks**: plain arithmetic, no solver. Each limit is looked at on
   its own:
   * Do the usage limits allow a blend at all? (Minimums must not add up to
     more than 100 %, maximums must reach 100 %.)
   * For each spec limit: can ANY mix of the allowed raw materials reach it,
     ignoring every other spec limit? E.g. if Fluorspar is banned and no
     other material contains CaF2, CaF2_min can never be met.

2. **Elastic model** (built in ``model``, solved by ``solve_and_report``,
   run by ``pipeline``): every spec limit may be missed, and the solver finds
   the smallest total relative miss. The limits that answer "pushes against"
   (non-zero dual values or reduced costs) are a set of limits that cannot
   all hold together. For each of them ``pipeline`` then finds the best
   value achievable while all the other limits are held, and this module
   turns it all into messages.

This module never solves an LP: it does arithmetic and writes messages.
"""

from dataclasses import dataclass, field

import pandas as pd

from blend_optimizer import config
from blend_optimizer import solve_and_report as sr
from blend_optimizer.data import BlendData, Term
from blend_optimizer.rounding import format_fixed
from blend_optimizer.linear_rows import MIN, LinearRow, achieved_value, build_linear_rows, linearise
from blend_optimizer.solve_and_report import SolveResult

_TOL = 1e-9             # duals, reduced costs and violations smaller than this count as zero
_AT_BOUND_PCT = 1e-7    # a share this close to a usage limit sits on it

DIAGNOSIS_COLUMNS = [
    "Limit", "Limits", "Found by", "Customer limit", "Internal target", "Best achievable",
    "Change needed", "While holding", "Elastic model target",
]


@dataclass(frozen=True)
class PrecheckFinding:
    """One problem found by the pre-checks.

    Attributes:
        message: Plain-language explanation.
        row: The spec limit concerned (None for a usage-limit finding).
        best: Best value achievable for that limit with only the usage limits applied.
    """

    message: str
    row: LinearRow | None = None
    best: float | None = None


@dataclass(frozen=True)
class Conflict:
    """Limits that the elastic model shows cannot all hold together.

    Attributes:
        spec_rows: Spec row names, e.g. ["CaO_min", "Al2O3_min"].
        usage_limits: (raw material, "min" or "max", limit wt %) for usage
            limits that take part, e.g. ("Fluorspar", "max", 0.0).
    """

    spec_rows: list[str] = field(default_factory=list)
    usage_limits: list[tuple[str, str, float]] = field(default_factory=list)


@dataclass
class Diagnosis:
    """The explanation of an infeasible problem.

    Attributes:
        messages: Plain-language explanation, in reading order.
        table: One line per limit that has to move (see DIAGNOSIS_COLUMNS).
        findings: The pre-check findings.
        conflict: The conflicting limits found by the elastic model.
    """

    messages: list[str]
    table: pd.DataFrame
    findings: list[PrecheckFinding] = field(default_factory=list)
    conflict: Conflict = field(default_factory=Conflict)


# ===========================================================================
# Wording helpers
# ===========================================================================

def _decimals(row: LinearRow) -> int:
    """Same display rule as the report: one more decimal than the typed limits."""
    spec = row.spec
    return sr.limit_decimals([spec.min_value, spec.max_value, spec.margin])


def _value(row: LinearRow, value: float) -> str:
    """A value of the limited quantity, e.g. '39.77 %' or '11.24' for a ratio."""
    unit = "" if row.spec.is_ratio else " %"
    return sr.format_number(value, sr.PCT, decimals=_decimals(row)) + unit


def _condition(row: LinearRow, target: float | None = None) -> str:
    """The limit written out, e.g. 'CaO >= 50.00 %'."""
    sign = ">=" if row.bound == MIN else "<="
    return f"{row.spec.expression} {sign} {_value(row, row.limit if target is None else target)}"


def _components_text(terms: tuple[Term, ...]) -> str:
    """The components of an expression, e.g. 'Al2O3 or Al' for Al2O3+1.8895*Al."""
    return " or ".join(term.component for term in terms)


def _usage_limit_source(data: BlendData, name: str, side: str) -> str:
    """Where a raw material's effective usage limit comes from.

    "plant-wide limit in RawMaterials", "UsageLimits", or the label of a
    what-if change such as "WhatIf (No fluorspar)".
    """
    plant = data.raw_materials[name].plant_limit
    product = data.product_limits.get(name)
    if product is None:
        return "plant-wide limit in RawMaterials"
    product_value = product.max_pct if side == "max" else product.min_pct
    plant_value = plant.max_pct if side == "max" else plant.min_pct
    tighter = product_value <= plant_value if side == "max" else product_value >= plant_value
    if not tighter:
        return "plant-wide limit in RawMaterials"
    return data.product_limit_sources.get(name, config.SHEET_USAGE_LIMITS)


# ===========================================================================
# Pre-checks (arithmetic only)
# ===========================================================================

def usage_bounds(data: BlendData) -> dict[str, tuple[float, float]]:
    """Raw material -> (lowest, highest) fraction allowed by the effective usage limits."""
    return {name: (data.effective_limit(name).min_pct / 100.0, data.effective_limit(name).max_pct / 100.0)
            for name in data.material_names}


def best_mix(coefficients: dict[str, float], bounds: dict[str, tuple[float, float]],
             highest: bool = True) -> float | None:
    """Highest (or lowest) value of sum(a_i x_i) over all blends the usage limits allow.

    With only sum(x_i) = 1 and the usage limits, the best blend is found by
    hand: start every material at its minimum, then fill the rest of the
    blend with the material that helps most, up to its maximum, then the
    next best, and so on. (This greedy fill is exactly optimal here because
    every material competes only for the same "room" in the blend.)

    Returns:
        The best value, or None if the usage limits allow no blend at all.
    """
    total_min = sum(low for low, _ in bounds.values())
    total_max = sum(high for _, high in bounds.values())
    if total_min > 1.0 + 1e-9 or total_max < 1.0 - 1e-9:
        return None
    shares = {name: low for name, (low, _) in bounds.items()}
    room = 1.0 - total_min
    for name in sorted(coefficients, key=coefficients.get, reverse=highest):
        add = min(bounds[name][1] - bounds[name][0], room)
        shares[name] += add
        room -= add
        if room <= 0:
            break
    return sum(coefficients[name] * share for name, share in shares.items())


def best_value_alone(row: LinearRow, bounds: dict[str, tuple[float, float]], highest: bool) -> float | None:
    """Highest (or lowest) blend value of a row's quantity, with only the usage limits applied.

    A blend value sum(N x)/sum(g x) can reach a level L exactly when some
    allowed blend has sum((N - L g) x) >= 0 (for "highest"), which
    :func:`best_mix` answers. So the best value is the largest such L, found
    by bisection between 0 and the best single material's value (every
    blend value is a weighted average of the materials' values).
    """
    allowed = [name for name, (_, high) in bounds.items() if high > 0]
    values = [row.numerator[name] / row.denominator[name] for name in allowed if row.denominator[name] > 0]
    if not values:
        return None

    def reachable(level: float) -> bool:
        best = best_mix(linearise(row.numerator, row.denominator, level), bounds, highest)
        return best is not None and (best >= -_TOL if highest else best <= _TOL)

    low, high = 0.0, max(values)
    if highest:
        if reachable(high):
            return high
        good, bad = low, high    # every blend value is >= 0, so 0 is always reachable
    else:
        if reachable(low):
            return low
        good, bad = high, low
    for _ in range(200):
        if abs(good - bad) <= 1e-10 * max(1.0, abs(good)):
            break
        middle = (good + bad) / 2.0
        if reachable(middle):
            good = middle
        else:
            bad = middle
    return good


def _usage_findings(data: BlendData) -> list[PrecheckFinding]:
    """Do the usage limits allow any blend at all?"""
    findings = []
    limits = {name: data.effective_limit(name) for name in data.material_names}
    total_min = sum(limit.min_pct for limit in limits.values())
    total_max = sum(limit.max_pct for limit in limits.values())
    if total_min > 100.0 + 1e-7:
        parts = ", ".join(f"{name} min {limit.min_pct:g} %" for name, limit in limits.items() if limit.min_pct > 0)
        findings.append(PrecheckFinding(
            f"The minimum usage limits add up to {format_fixed(total_min, 2)} % of the blend, more than 100 % "
            f"({parts})."))
    if total_max < 100.0 - 1e-7:
        parts = ", ".join(f"{name} max {limit.max_pct:g} %" for name, limit in limits.items())
        findings.append(PrecheckFinding(
            f"The maximum usage limits add up to only {format_fixed(total_max, 2)} %, so the allowed raw materials "
            f"cannot make up a whole blend ({parts})."))
    return findings


def _held_back(data: BlendData, row: LinearRow, bounds: dict[str, tuple[float, float]]) -> str:
    """Raw materials that would help reach the limit but are restricted by a usage limit."""
    is_min = row.bound == MIN
    parts = []
    for name in data.material_names:
        top, bottom = row.numerator[name], row.denominator[name]
        if bottom > 0:
            value = top / bottom
            helps = value > row.limit if is_min else value < row.limit
        else:
            value, helps = None, is_min and top > 0
        high = bounds[name][1]
        if not helps or high >= 1.0:
            continue
        restriction = (f"not allowed (Max 0 %, {_usage_limit_source(data, name, 'max')})" if high == 0
                       else f"capped at {100 * high:g} % ({_usage_limit_source(data, name, 'max')})")
        content = f"{row.spec.expression} {_value(row, value)}" if value is not None else row.spec.expression
        parts.append(f"{name} ({content}) is {restriction}")
    return ("Raw materials that would help are restricted: " + "; ".join(parts) + ".") if parts else ""


def _row_finding(data: BlendData, row: LinearRow, bounds: dict[str, tuple[float, float]]) -> PrecheckFinding:
    """Explain a spec limit that no allowed blend can reach."""
    is_min = row.bound == MIN
    best = best_value_alone(row, bounds, highest=is_min)
    allowed = [name for name, (_, high) in bounds.items() if high > 0]
    if is_min and all(row.numerator[name] == 0 for name in allowed):
        message = (f"No allowed raw material contains {_components_text(row.spec.numerator)}, "
                   f"so {row.name} ({_condition(row)}) cannot be met.")
    else:
        word = "highest" if is_min else "lowest"
        achievable = "none" if best is None else _value(row, best)
        message = (f"{row.name} ({_condition(row)}) cannot be met by any mix of the allowed raw materials, "
                   f"even ignoring all other spec limits: the {word} achievable {row.spec.expression} "
                   f"is {achievable}.")
    held_back = _held_back(data, row, bounds)
    if held_back:
        message += " " + held_back
    return PrecheckFinding(message, row, best)


def precheck(data: BlendData) -> list[PrecheckFinding]:
    """Look at each limit on its own: can it be met at all? (No solver.)

    Returns:
        The problems found; an empty list if every limit can be met on its
        own (then any infeasibility is a conflict between limits).
    """
    usage = _usage_findings(data)
    if usage:
        return usage  # with no possible blend at all, the spec checks would mean nothing
    bounds = usage_bounds(data)
    findings = []
    for row in build_linear_rows(data):
        best = best_mix(row.coefficients, bounds, highest=row.bound == MIN)
        tolerance = 1e-9 * max(1.0, abs(row.rhs))
        impossible = best < row.rhs - tolerance if row.bound == MIN else best > row.rhs + tolerance
        if impossible:
            findings.append(_row_finding(data, row, bounds))
    return findings


# ===========================================================================
# Reading the elastic model's answer
# ===========================================================================

def conflict_set(data: BlendData, elastic: SolveResult) -> Conflict:
    """The limits the elastic answer pushes against: they cannot all hold together.

    In the elastic LP a limit with a non-zero dual value (or a violation)
    is one the minimum violation depends on: loosening it would reduce the
    violation. Those limits, with the usage limits whose reduced cost is
    non-zero, form a set that cannot all be met at once. (It is one such
    set; a problem can contain several.)
    """
    spec_rows = [name for name in elastic.rows
                 if abs(elastic.row_dual.get(name, 0.0)) > _TOL or elastic.violations.get(name, 0.0) > _TOL]
    usage = []
    for name in data.material_names:
        limit = data.effective_limit(name)
        share = 100.0 * elastic.fractions[name]
        dj = elastic.reduced_costs[name]
        if limit.max_pct < 100.0 and share >= limit.max_pct - _AT_BOUND_PCT and dj < -_TOL:
            usage.append((name, "max", limit.max_pct))
        elif limit.min_pct > 0.0 and share <= limit.min_pct + _AT_BOUND_PCT and dj > _TOL:
            usage.append((name, "min", limit.min_pct))
    return Conflict(spec_rows, usage)


def _usage_condition(name: str, side: str, value: float) -> str:
    return f"{name} {'<=' if side == 'max' else '>='} {value:g} % (usage limit)"


# ===========================================================================
# Putting the diagnosis together
# ===========================================================================

def build_diagnosis(
    data: BlendData,
    findings: list[PrecheckFinding],
    elastic: SolveResult | None,
    conflict: Conflict,
    achievable: dict[str, float | None],
) -> Diagnosis:
    """Turn the pre-check findings and the elastic results into messages and a table.

    Args:
        data: The input data.
        findings: From :func:`precheck`.
        elastic: The elastic model's result (None if it was not run).
        conflict: From :func:`conflict_set`.
        achievable: Spec row name -> best value achievable while all other
            limits are held (None if the others conflict even without it),
            worked out by ``pipeline.best_achievable``.
    """
    messages = ["No blend meets every limit (the problem is infeasible)."]
    records = []

    # --- Pre-checks ---------------------------------------------------------
    if findings:
        messages.append("Pre-check (each limit on its own, with only the usage limits applied):")
        messages += [f"  - {finding.message}" for finding in findings]
        for finding in findings:
            if finding.row is not None:
                records.append(_table_record(finding.row, "Pre-check", finding.best, "usage limits only"))
    else:
        messages.append("Pre-check: every limit can be met on its own, so some limits conflict with each other.")

    # --- Elastic model ------------------------------------------------------
    if elastic is None or not elastic.is_optimal:
        status = "not run" if elastic is None else elastic.status
        messages.append(f"Elastic model: no answer ({status}). The usage limits themselves leave no possible "
                        f"blend; see the pre-check.")
        return Diagnosis(messages, pd.DataFrame(records, columns=DIAGNOSIS_COLUMNS), findings, conflict)

    members = conflict.spec_rows + [f"{name} {side} use" for name, side, _ in conflict.usage_limits]
    if members:
        messages.append(f"Elastic model: these limits cannot all be met together: {', '.join(members)}.")
    for row_name in conflict.spec_rows:
        row = elastic.rows[row_name]
        others = [_condition(elastic.rows[other]) for other in conflict.spec_rows if other != row_name]
        others += [_usage_condition(*usage) for usage in conflict.usage_limits]
        holding = (f"while {' and '.join(others)} and all other limits are held" if others
                   else "while all other limits are held")
        best = achievable.get(row_name)
        if best is None:
            messages.append(f"  - {row_name} ({_condition(row)}): even without this limit, the other "
                            f"limits cannot all be met.")
        else:
            word = "highest" if row.bound == MIN else "lowest"
            messages.append(f"  - {row_name} ({_condition(row)}): the {word} {row.spec.expression} "
                            f"achievable {holding} is {_value(row, best)}.")
        reached = achieved_value(row, elastic.fractions) if elastic.violations.get(row_name, 0) > _TOL else None
        records.append(_table_record(row, "Elastic model", best, "all other limits", reached))
    for name, side, value in conflict.usage_limits:
        records.append({"Limit": f"{name} {side} use", "Limits": name, "Found by": "Elastic model",
                        "Internal target": value, "While holding": "usage limit"})

    changes = []
    for row_name, violation in elastic.violations.items():
        if violation > _TOL:
            row = elastic.rows[row_name]
            verb = "lower" if row.bound == MIN else "raise"
            changes.append(f"{verb} {row_name} from {_value(row, row.limit)} to "
                           f"{_value(row, achieved_value(row, elastic.fractions))}")
    if changes:
        messages.append("Smallest total relative change that makes a blend possible (one way; there may be "
                        "others): " + "; ".join(changes) + ".")
    return Diagnosis(messages, pd.DataFrame(records, columns=DIAGNOSIS_COLUMNS), findings, conflict)


def _table_record(row: LinearRow, found_by: str, best: float | None, holding: str,
                  elastic_target: float | None = None) -> dict[str, object]:
    """One Diagnosis table line for a spec limit."""
    if best is None:
        change = None
    else:
        change = row.limit - best if row.bound == MIN else best - row.limit
    return {
        "Limit": row.name,
        "Limits": row.spec.expression,
        "Found by": found_by,
        "Customer limit": row.customer_limit,
        "Internal target": row.limit,
        "Best achievable": best,
        "Change needed": change,
        "While holding": holding,
        "Elastic model target": elastic_target,
    }


def format_diagnosis(diagnosis: Diagnosis) -> str:
    """The diagnosis as console text."""
    return "\n".join(diagnosis.messages)

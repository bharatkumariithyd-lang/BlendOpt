"""What-if scenarios and price ranging.

* :func:`apply_changes` makes a changed copy of the data (prices, spec
  limits, usage limits). The original data is never modified.
* :func:`what_if` solves the base case and the scenario and compares them
  side by side: changes, cost, recipe, spec check and binding limits.
* :func:`run_scenarios` does the same for every named scenario of the
  optional WhatIf sheet, diagnosing infeasible ones automatically, and
  :func:`scenario_sheets` lays the results out as report sheets.
* :func:`price_ranging` finds, for each raw material, the price range over
  which the recipe stays the same.

Solving goes through ``pipeline`` and the tables through
``solve_and_report``; this module never touches PuLP.

Change formats (names as in the workbook):

    price_changes = {"Lime": 9500}                 Rs/t, in the material's own price basis
    spec_changes  = {"Al2O3": {"Min": 32}}         keys Min, Max, Margin; None removes a limit
    usage_changes = {"Fluorspar": {"Max": 0}}      keys Min, Max (wt % of dry blend); product
                                                   limits, like the UsageLimits sheet
"""

import math
from dataclasses import dataclass, replace

import pandas as pd

from blend_optimizer import config
from blend_optimizer import solve_and_report as sr
from blend_optimizer.data import BlendData, Scenario, UsageLimit
from blend_optimizer.diagnose import Diagnosis, format_diagnosis
from blend_optimizer.inputs import spec_limit_problems, suggest_close_match
from blend_optimizer.pipeline import INFEASIBLE_STATUSES, diagnose_infeasible, solve_blend
from blend_optimizer.solve_and_report import SolveResult

SHEET_SCENARIOS = "Scenarios"

_SPEC_FIELDS = {config.COL_MIN: "min_value", config.COL_MAX: "max_value", config.COL_MARGIN: "margin"}
_USAGE_FIELDS = (config.COL_MIN, config.COL_MAX)

# Two recipes count as "the same" if no raw material's share differs by more
# than this, in wt % of the dry blend.
# Why 0.01 wt %: in a 1,000 kg batch that is 0.1 kg, the default scale
# resolution. A change smaller than the scale can weigh is not a real recipe
# change: the operator would load exactly the same amounts. (Solver round-off
# is many orders of magnitude smaller, so it never counts as a change.)
RECIPE_TOLERANCE = 0.01  # wt %


# ===========================================================================
# Applying changes
# ===========================================================================

def _check_name(name: str, known: list[str], what: str) -> None:
    """Raise a helpful ValueError if ``name`` is not one of ``known``."""
    if name in known:
        return
    suggestion = suggest_close_match(name, known)
    hint = f" Did you mean '{suggestion}'?" if suggestion else f" Known: {', '.join(known)}."
    raise ValueError(f"Unknown {what} '{name}'.{hint}")


def _check_fields(name: str, fields: dict[str, object], allowed: tuple[str, ...] | list[str]) -> None:
    unknown = [field for field in fields if field not in allowed]
    if unknown:
        raise ValueError(f"'{name}': unknown field(s) {', '.join(unknown)}; use {', '.join(allowed)}.")


def apply_changes(
    data: BlendData,
    price_changes: dict[str, float] | None = None,
    spec_changes: dict[str, dict[str, float | None]] | None = None,
    usage_changes: dict[str, dict[str, float | None]] | None = None,
    usage_source: str | None = None,
) -> BlendData:
    """Return a copy of ``data`` with the changes applied (``data`` itself is unchanged).

    Args:
        data: The base case.
        price_changes: Material -> new price, Rs/t in that material's price basis.
        spec_changes: Spec name -> {"Min"/"Max"/"Margin": new value}. None
            removes a Min or Max (a Margin of None means 0).
        usage_changes: Material -> {"Min"/"Max": new product limit, wt %}.
            These replace the product (UsageLimits) limit; the plant-wide
            limit still applies, the tighter one wins. None resets a side
            (Min to 0, Max to 100).
        usage_source: Label for where the changed usage limits came from,
            shown in diagnosis messages (e.g. "WhatIf (No fluorspar)").
            None keeps "UsageLimits".

    Raises:
        ValueError: For unknown names or fields, or limits that make no sense.
    """
    materials = dict(data.raw_materials)
    for name, price in (price_changes or {}).items():
        _check_name(name, data.material_names, "raw material")
        if not isinstance(price, (int, float)) or isinstance(price, bool) or not price >= 0:
            raise ValueError(f"Price for '{name}' must be a number of 0 or more, got {price!r}.")
        materials[name] = replace(materials[name], cost_input=float(price))

    specs = list(data.specs)
    position = {spec.name: index for index, spec in enumerate(specs)}
    for name, fields in (spec_changes or {}).items():
        _check_name(name, list(position), "spec")
        _check_fields(name, fields, list(_SPEC_FIELDS))
        changes = {_SPEC_FIELDS[field]: value for field, value in fields.items()}
        if changes.get("margin", 0.0) is None:
            changes["margin"] = 0.0
        spec = replace(specs[position[name]], **changes)
        problems = spec_limit_problems(spec)
        if problems:
            raise ValueError(f"Spec '{name}': {problems[0]}")
        specs[position[name]] = spec

    product_limits = dict(data.product_limits)
    limit_sources = dict(data.product_limit_sources)
    for name, fields in (usage_changes or {}).items():
        _check_name(name, data.material_names, "raw material")
        _check_fields(name, fields, _USAGE_FIELDS)
        current = product_limits.get(name, UsageLimit())
        new_min = fields.get(config.COL_MIN, current.min_pct)
        new_max = fields.get(config.COL_MAX, current.max_pct)
        limit = UsageLimit(0.0 if new_min is None else float(new_min), 100.0 if new_max is None else float(new_max))
        if not 0 <= limit.min_pct <= limit.max_pct <= 100:
            raise ValueError(f"Usage limit for '{name}': need 0 <= Min <= Max <= 100, got "
                             f"Min {limit.min_pct:g}, Max {limit.max_pct:g}.")
        effective = materials[name].plant_limit.tighter(limit)
        if effective.min_pct > effective.max_pct:
            raise ValueError(f"Usage limit for '{name}' conflicts with the plant-wide limit: effective "
                             f"Min {effective.min_pct:g} % > effective Max {effective.max_pct:g} %.")
        product_limits[name] = limit
        if usage_source is not None:
            limit_sources[name] = usage_source

    return replace(data, raw_materials=materials, specs=specs, product_limits=product_limits,
                   product_limit_sources=limit_sources)


# ===========================================================================
# Side-by-side comparison
# ===========================================================================

def what_if(
    data: BlendData,
    price_changes: dict[str, float] | None = None,
    spec_changes: dict[str, dict[str, float | None]] | None = None,
    usage_changes: dict[str, dict[str, float | None]] | None = None,
) -> dict[str, pd.DataFrame]:
    """Re-solve with changes and compare the scenario with the base case.

    Args:
        data: The base case.
        price_changes, spec_changes, usage_changes: See :func:`apply_changes`.

    Returns:
        Tables, sheet name -> table (full precision; display with
        :func:`format_what_if` or ``solve_and_report.write_excel``):
        "Changes", "Cost", "Recipe", "Spec check", "Binding limits".
        If either case is infeasible its columns are blank and its status
        says so (a scenario can be the fix for an infeasible base case).
    """
    scenario_data = apply_changes(data, price_changes, spec_changes, usage_changes)
    return compare(data, solve_blend(data), scenario_data, solve_blend(scenario_data))


def compare(data: BlendData, base: SolveResult, scenario_data: BlendData,
            scenario: SolveResult) -> dict[str, pd.DataFrame]:
    """The side-by-side tables for an already solved base case and scenario."""
    return {
        "Changes": _changes_table(data, scenario_data),
        "Cost": _cost_comparison(data, base, scenario_data, scenario),
        "Recipe": _recipe_comparison(data, base, scenario_data, scenario),
        "Spec check": _spec_comparison(data, base, scenario_data, scenario),
        "Binding limits": _binding_comparison(data, base, scenario_data, scenario),
    }


def _changes_table(base: BlendData, scenario: BlendData) -> pd.DataFrame:
    """What differs between the two data sets, one line per changed value.

    The value columns are called "Base value"/"Scenario value" so the report
    treats them as typed limits: a P Max change from 0.03 to 0.035 is shown
    as 0.0300 -> 0.0350, not rounded to 2 decimals.
    """
    records = []
    for name, material in base.raw_materials.items():
        new = scenario.raw_materials[name]
        if new.cost_input != material.cost_input:
            records.append({"Change": "Price", "Name": name, "Unit": f"Rs/t ({material.price_basis})",
                            "Base value": material.cost_input, "Scenario value": new.cost_input})
    for old, new in zip(base.specs, scenario.specs):
        unit = "ratio" if old.is_ratio else f"wt % ({base.settings.spec_basis} basis)"
        for field, attribute in _SPEC_FIELDS.items():
            if getattr(old, attribute) != getattr(new, attribute):
                records.append({"Change": f"Spec {field}", "Name": old.name, "Unit": unit,
                                "Base value": getattr(old, attribute), "Scenario value": getattr(new, attribute)})
    for name in base.material_names:
        old, new = base.effective_limit(name), scenario.effective_limit(name)
        for label, before, after in (("Min", old.min_pct, new.min_pct), ("Max", old.max_pct, new.max_pct)):
            if before != after:
                records.append({"Change": f"Usage {label} (effective)", "Name": name, "Unit": sr.USAGE_UNIT,
                                "Base value": before, "Scenario value": after})
    return pd.DataFrame(records, columns=["Change", "Name", "Unit", "Base value", "Scenario value"])


def _cost_comparison(data: BlendData, base: SolveResult, scenario_data: BlendData,
                     scenario: SolveResult) -> pd.DataFrame:
    def costs(case_data: BlendData, result: SolveResult) -> dict[str, float]:
        if not result.is_optimal:
            return {}
        return sr.cost_table(case_data, result).set_index("Item")["Value"].to_dict()

    base_costs, scenario_costs = costs(data, base), costs(scenario_data, scenario)
    records = [{"Item": "Solver status", "Unit": "", "Base": base.status, "Scenario": scenario.status,
                "Change": None}]
    for item, unit in (("Cost (Rs per dry tonne of blend)", "Rs/dry t"),
                       ("Cost (Rs per as-received tonne of blend)", "Rs/as-received t"),
                       ("Cost per batch (Rs)", "Rs")):
        before, after = base_costs.get(item), scenario_costs.get(item)
        records.append({"Item": item, "Unit": unit, "Base": before, "Scenario": after,
                        "Change": None if before is None or after is None else after - before})
    return pd.DataFrame(records)


def _recipe_comparison(data: BlendData, base: SolveResult, scenario_data: BlendData,
                       scenario: SolveResult) -> pd.DataFrame:
    def column(case_data: BlendData, result: SolveResult, name: str) -> pd.Series:
        if not result.is_optimal:
            return pd.Series([math.nan] * (len(case_data.material_names) + 1))
        return sr.recipe_table(case_data, result)[name]

    wt, kg = "Dry blend (wt %)", "As-received mass (kg/batch)"
    table = pd.DataFrame({"Raw material": data.material_names + ["TOTAL"]})
    table["Base (wt %)"] = column(data, base, wt)
    table["Scenario (wt %)"] = column(scenario_data, scenario, wt)
    table["Change (wt %)"] = table["Scenario (wt %)"] - table["Base (wt %)"]
    table["Base (kg/batch, as received)"] = column(data, base, kg)
    table["Scenario (kg/batch, as received)"] = column(scenario_data, scenario, kg)
    return table


def _spec_comparison(data: BlendData, base: SolveResult, scenario_data: BlendData,
                     scenario: SolveResult) -> pd.DataFrame:
    def checks(case_data: BlendData, result: SolveResult) -> pd.DataFrame | None:
        return sr.spec_check_table(case_data, result).set_index("Spec") if result.is_optimal else None

    base_check, scenario_check = checks(data, base), checks(scenario_data, scenario)
    records = []
    for spec in scenario_data.specs:  # the scenario's limits are the ones shown
        records.append({
            "Spec": spec.name,
            "Limits": spec.expression,
            "Customer Min": spec.min_value,
            "Customer Max": spec.max_value,
            "Margin": spec.margin,
            "Base achieved": None if base_check is None else base_check.loc[spec.name, "Achieved"],
            "Scenario achieved": None if scenario_check is None else scenario_check.loc[spec.name, "Achieved"],
            "Base status": base.status if base_check is None else base_check.loc[spec.name, "Status"],
            "Scenario status": scenario.status if scenario_check is None else scenario_check.loc[spec.name, "Status"],
        })
    return pd.DataFrame(records)


def _binding_comparison(data: BlendData, base: SolveResult, scenario_data: BlendData,
                        scenario: SolveResult) -> pd.DataFrame:
    """Every limit binding in either case, with its saving per step in each (blank = not binding)."""
    columns = ["Constraint", "Limit unit", "Step", "Base saving per step (Rs/dry t)",
               "Scenario saving per step (Rs/dry t)"]

    def shadow(case_data: BlendData, result: SolveResult) -> pd.DataFrame:
        if not result.is_optimal:
            return pd.DataFrame(columns=sr.SHADOW_PRICE_COLUMNS).set_index("Constraint")
        return sr.shadow_price_table(case_data, result).set_index("Constraint")

    base_table, scenario_table = shadow(data, base), shadow(scenario_data, scenario)
    names = list(base_table.index) + [name for name in scenario_table.index if name not in base_table.index]
    records = []
    for name in names:
        source = base_table if name in base_table.index else scenario_table
        records.append({
            "Constraint": name,
            "Limit unit": source.loc[name, "Limit unit"],
            "Step": source.loc[name, "Step"],
            "Base saving per step (Rs/dry t)":
                base_table.loc[name, "Saving per step (Rs/dry t)"] if name in base_table.index else None,
            "Scenario saving per step (Rs/dry t)":
                scenario_table.loc[name, "Saving per step (Rs/dry t)"] if name in scenario_table.index else None,
        })
    return pd.DataFrame(records, columns=columns)


# Console columns for each comparison table (the Excel sheets show all columns)
_CONSOLE_COLUMNS = {
    "Changes": ["Change", "Name", "Unit", "Base value", "Scenario value"],
    "Cost": ["Item", "Base", "Scenario", "Change"],
    "Recipe": ["Raw material", "Base (wt %)", "Scenario (wt %)", "Change (wt %)"],
    "Spec check": ["Spec", "Base achieved", "Scenario achieved", "Base status", "Scenario status"],
    "Binding limits": ["Constraint", "Step", "Limit unit", "Base saving per step (Rs/dry t)",
                       "Scenario saving per step (Rs/dry t)"],
}


def format_what_if(tables: dict[str, pd.DataFrame], scale_resolution_kg: float = config.SCALE_RESOLUTION_KG,
                   title: str = "=== What-if: base case vs scenario ===") -> str:
    """The what-if comparison as console text."""
    lines = [title]
    for sheet, columns in _CONSOLE_COLUMNS.items():
        lines += ["", f"{sheet}:"]
        table = tables[sheet]
        lines.append("  none" if table.empty else sr.indented_table(table, columns, scale_resolution_kg))
    return "\n".join(lines)


def print_what_if(tables: dict[str, pd.DataFrame], scale_resolution_kg: float = config.SCALE_RESOLUTION_KG) -> None:
    """Print the what-if comparison to the console."""
    print(format_what_if(tables, scale_resolution_kg))


# ===========================================================================
# Named scenarios from the WhatIf sheet
# ===========================================================================

@dataclass
class ScenarioOutcome:
    """One WhatIf scenario, solved and compared with the base case.

    Attributes:
        scenario: The scenario as read from the WhatIf sheet.
        data: The data with the scenario's changes applied.
        result: The scenario's solver result.
        tables: The comparison tables (see :func:`what_if`).
        diagnosis: Explanation if the scenario is infeasible, else None.
    """

    scenario: Scenario
    data: BlendData
    result: SolveResult
    tables: dict[str, pd.DataFrame]
    diagnosis: Diagnosis | None = None


def scenario_data(data: BlendData, scenario: Scenario) -> BlendData:
    """The base data with one scenario's changes applied.

    Its usage changes are labelled "WhatIf (<scenario name>)", so a diagnosis
    names the scenario as the source of a ban or cap.
    """
    return apply_changes(data, scenario.price_changes, scenario.spec_changes, scenario.usage_changes,
                         usage_source=f"{config.SHEET_WHAT_IF} ({scenario.name})")


def check_scenarios(data: BlendData) -> list[str]:
    """Check that every scenario's changes make sense together, before anything is solved.

    The WhatIf rows were checked one by one when the workbook was read; this
    catches problems that only show once a scenario's changes are applied,
    such as a new Min above the Max.

    Returns:
        Error messages (empty if all scenarios are fine).
    """
    errors = []
    for scenario in data.scenarios:
        try:
            scenario_data(data, scenario)
        except ValueError as error:
            errors.append(f"WhatIf scenario '{scenario.name}': {error}")
    return errors


def run_scenarios(data: BlendData) -> list[ScenarioOutcome]:
    """Solve every WhatIf scenario and compare it with the base case.

    An infeasible scenario is diagnosed automatically, with the same
    pre-checks and elastic model as an infeasible base case.

    Raises:
        ValueError: If a scenario's changes are invalid (run
            :func:`check_scenarios` first to get all such errors at once).
    """
    base = solve_blend(data)
    outcomes = []
    for scenario in data.scenarios:
        changed = scenario_data(data, scenario)
        result = solve_blend(changed)
        diagnosis = diagnose_infeasible(changed) if result.status in INFEASIBLE_STATUSES else None
        outcomes.append(ScenarioOutcome(scenario, changed, result, compare(data, base, changed, result), diagnosis))
    return outcomes


_FORBIDDEN_IN_SHEET_NAMES = str.maketrans({character: " " for character in "[]:*?/\\"})


def scenario_sheet_name(scenario_name: str, used: set[str]) -> str:
    """Excel sheet name for a scenario: 'Scenario <name>', at most 31 characters, unique."""
    base = f"Scenario {scenario_name}".translate(_FORBIDDEN_IN_SHEET_NAMES).strip()[:31]
    name, counter = base, 2
    while name.lower() in used:
        suffix = f" ({counter})"
        name = base[:31 - len(suffix)] + suffix
        counter += 1
    used.add(name.lower())
    return name


def _changes_text(outcome: ScenarioOutcome, scale_resolution_kg: float) -> str:
    """The scenario's changes in one line, e.g. 'Price Fluorspar 28,000 -> 35,000'."""
    shown = sr.display_table(outcome.tables["Changes"], scale_resolution_kg)
    return "; ".join(f"{row['Change']} {row['Name']} {row['Base value'] or 'none'} -> {row['Scenario value'] or 'none'}"
                     for _, row in shown.iterrows())


def scenario_sheets(data: BlendData, outcomes: list[ScenarioOutcome],
                    reserved: set[str] | None = None) -> dict[str, pd.DataFrame | list[sr.Section]]:
    """Report sheets for the scenarios: a "Scenarios" overview plus one sheet per scenario.

    Each scenario sheet holds the comparison tables one under another, and
    a Diagnosis section if the scenario is infeasible.

    Args:
        data: The base data.
        outcomes: From :func:`run_scenarios`.
        reserved: Sheet names already used in the workbook (lower case).
    """
    if not outcomes:
        return {}
    resolution = data.settings.scale_resolution_kg
    used = set(reserved or set()) | {SHEET_SCENARIOS.lower()}
    sheets: dict[str, pd.DataFrame | list[sr.Section]] = {}
    overview = []
    for outcome in outcomes:
        name = outcome.scenario.name
        sheet = scenario_sheet_name(name, used)
        cost = outcome.tables["Cost"].set_index("Item")
        overview.append({
            "Scenario": name,
            "Sheet": sheet,
            "Changes": _changes_text(outcome, resolution),
            "Status": outcome.result.status,
            "Cost (Rs/dry t)": cost.loc["Cost (Rs per dry tonne of blend)", "Scenario"],
            "Change vs base (Rs/dry t)": cost.loc["Cost (Rs per dry tonne of blend)", "Change"],
        })
        sections = [sr.Section(f"{name}: {title.lower()}", table) for title, table in outcome.tables.items()]
        if outcome.diagnosis is not None:
            sections.append(sr.Section(f"{name}: diagnosis", outcome.diagnosis.table, outcome.diagnosis.messages))
        sheets[sheet] = sections
    return {SHEET_SCENARIOS: pd.DataFrame(overview), **sheets}


def format_scenario(outcome: ScenarioOutcome, scale_resolution_kg: float = config.SCALE_RESOLUTION_KG) -> str:
    """One scenario's comparison (and diagnosis, if infeasible) as console text."""
    text = format_what_if(outcome.tables, scale_resolution_kg,
                          title=f"=== Scenario: {outcome.scenario.name} (vs base case) ===")
    if outcome.diagnosis is not None:
        text += "\n\n" + format_diagnosis(outcome.diagnosis)
    return text


# ===========================================================================
# Price ranging
# ===========================================================================
#
# Method: re-solve the LP at trial prices (bisection). No solver-specific
# ranging report is used, so this works with any solver.
#
# Why a single interval: in an LP, a given recipe x stays optimal while one
# price p moves inside an interval around today's price, and outside that
# interval a different recipe is cheaper. The reason is that the cost of any
# recipe is a straight line in p. If x is cheaper than every other recipe y
# at two prices p1 and p2, then (being straight lines) it is also cheaper at
# every price in between. So the prices at which x stays optimal form one
# unbroken interval, and bisection can find its ends.
#
# For each raw material:
#   Lower end: solve at price 0. If the recipe is still the same there is no
#     lower limit. Otherwise the end lies between 0 (different) and today's
#     price (same): halve that gap, keeping one end "same" and the other
#     "different", until it is narrower than ``tolerance_rs``.
#   Upper end: the same, between today's price and search_factor x today's
#     price. If even that price leaves the recipe unchanged, no upper limit
#     was found in the range tested.
# Each halving step is one solve, so each end costs about log2(range / tolerance)
# solves (about 15-20). The ends reported are the last prices found to leave
# the recipe unchanged, so they are never outside the true range.
#
# Prices are changed and reported in each material's own price basis
# (as entered), because that is how they are bought and quoted.

PRICE_RANGING_COLUMNS = [
    "Raw material", "Price basis", "Current price (Rs/t)", "In recipe (wt %)",
    "Unchanged from (Rs/t)", "Unchanged up to (Rs/t)", "Note",
]


def _same_recipe(a: dict[str, float], b: dict[str, float]) -> bool:
    """True if no share differs by more than RECIPE_TOLERANCE wt % (fractions are 0..1)."""
    return all(100.0 * abs(a[name] - b[name]) <= RECIPE_TOLERANCE for name in a)


def _recipe_at_price(data: BlendData, name: str, price: float) -> dict[str, float] | None:
    """The optimal recipe with one material's price changed (None if not optimal)."""
    result = solve_blend(apply_changes(data, price_changes={name: price}))
    return result.fractions if result.is_optimal else None


def _bisect(data: BlendData, name: str, base: dict[str, float], same_price: float, other_price: float,
            tolerance_rs: float) -> float:
    """Narrow [same_price, other_price] until narrower than tolerance; return the "same" end."""
    while abs(other_price - same_price) > tolerance_rs:
        middle = (same_price + other_price) / 2.0
        recipe = _recipe_at_price(data, name, middle)
        if recipe is not None and _same_recipe(recipe, base):
            same_price = middle
        else:
            other_price = middle
    return same_price


def price_ranging(data: BlendData, tolerance_rs: float = 1.0, search_factor: float = 10.0) -> pd.DataFrame:
    """For each raw material, the price range over which the recipe stays the same.

    See the method notes above this function.

    Args:
        data: The base case (must be feasible).
        tolerance_rs: Accuracy of the range ends, Rs/t.
        search_factor: Upper prices are tested up to this multiple of the
            current price (or of the highest price in the list, for a
            material that is currently free).

    Returns:
        One row per raw material, prices in its own price basis. A blank
        range end means no limit was found; the Note says what was tested.
    """
    base_result = solve_blend(data)
    if not base_result.is_optimal:
        raise ValueError(f"The base case has no solution (status '{base_result.status}'); "
                         f"price ranging needs a feasible recipe.")
    base = base_result.fractions
    highest_price = max(material.cost_input for material in data.raw_materials.values())

    records = []
    for name, material in data.raw_materials.items():
        price = material.cost_input
        cap = search_factor * (price if price > 0 else max(highest_price, 1.0))
        notes = []

        # Lower end
        at_zero = _recipe_at_price(data, name, 0.0)
        if price == 0 or (at_zero is not None and _same_recipe(at_zero, base)):
            lower = None
        else:
            lower = _bisect(data, name, base, price, 0.0, tolerance_rs)

        # Upper end
        at_cap = _recipe_at_price(data, name, cap)
        if at_cap is not None and _same_recipe(at_cap, base):
            upper = None
        else:
            upper = _bisect(data, name, base, price, cap, tolerance_rs)

        in_recipe = 100.0 * base[name]
        if lower is None and upper is None:
            notes.append(f"Recipe unchanged at every price tested (Rs 0 to {sr.format_number(cap, sr.RS)})")
        else:
            if lower is None:
                notes.append("No lower limit: unchanged down to Rs 0")
            if upper is None:
                notes.append(f"No upper limit found up to Rs {sr.format_number(cap, sr.RS)} "
                             f"({search_factor:g}x the current price)")
        if base[name] <= sr._ZERO and lower is not None:
            notes.append(f"Not used now; enters the recipe below Rs {sr.format_number(lower, sr.RS)}")

        records.append({
            "Raw material": name,
            "Price basis": material.price_basis,
            "Current price (Rs/t)": price,
            "In recipe (wt %)": in_recipe,
            "Unchanged from (Rs/t)": lower,
            "Unchanged up to (Rs/t)": upper,
            "Note": "; ".join(notes),
        })
    return pd.DataFrame(records, columns=PRICE_RANGING_COLUMNS)


SHEET_PRICE_RANGES = "Price ranges"


def price_ranging_report(data: BlendData, tolerance_rs: float = 1.0,
                         search_factor: float = 10.0) -> tuple[pd.DataFrame, list[str]]:
    """The "Price ranges" report sheet for the base case: the table and its notes.

    If the base case is infeasible there is no recipe to keep unchanged, so
    the table is empty and the notes say why.

    Args:
        data: The base case.
        tolerance_rs, search_factor: See :func:`price_ranging`.

    Returns:
        (table, notes for under the table).
    """
    if not solve_blend(data).is_optimal:
        return (pd.DataFrame(columns=PRICE_RANGING_COLUMNS),
                ["Not available: price ranges need a feasible base case (see the Diagnosis sheet)."])
    notes = [
        "For each raw material on its own (all other prices as they are): the price band within which "
        "today's recipe stays the same, in the material's own price basis (Rs/t as entered).",
        f"Found by re-solving at trial prices (bisection), accurate to Rs {tolerance_rs:g}/t. A recipe counts "
        f"as unchanged if no share moves by more than {RECIPE_TOLERANCE:g} wt %.",
        f"A blank end means no limit was found in the range tested (Rs 0 up to {search_factor:g}x the current "
        f"price); the Note says which.",
        "For a raw material not in the recipe, 'Unchanged from' is the price below which it would enter.",
    ]
    return price_ranging(data, tolerance_rs, search_factor), notes

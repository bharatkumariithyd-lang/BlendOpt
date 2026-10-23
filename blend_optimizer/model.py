"""Build the PuLP problem from validated data.  (PuLP lives here and in solve_and_report.)

The LP, with x_i = dry mass fraction of raw material i:

    minimise    sum(dry_cost_i * x_i)                 Rs per dry tonne of blend
    subject to  sum(x_i) = 1                          "Sum_of_fractions"
                one row per spec bound                "<spec>_min" / "<spec>_max"
                min_use_i/100 <= x_i <= max_use_i/100 (variable bounds)

The spec rows come ready-made from ``linear_rows``; this module only copies
them into PuLP. It knows nothing about Excel.

Elastic version (``elastic=True``, used to diagnose infeasible problems):
every spec row may be missed by a violation v_k >= 0,

    >= rows:  sum(a_i x_i) + v_k >= rhs        <= rows:  sum(a_i x_i) - v_k <= rhs

and the objective becomes the total relative violation, sum(v_k / scale_k),
with scale_k from ``linear_rows.violation_scale``. Usage limits and
sum = 1 stay hard. The answer is always feasible (unless the usage limits
themselves are impossible) and shows which limits have to move.

Built in small steps (``_add_...`` functions) so that integer features can
be added later as extra steps without rewriting the rest, e.g.
    * whole-bag additions:  x_i = bag_size * n_i with integer n_i
    * minimum-if-used:      binary y_i, x_i >= min_i * y_i, x_i <= max_i * y_i
    * max number of materials:  sum(y_i) <= K
"""

import re
from dataclasses import dataclass, field

from pulp import LpConstraint, LpMinimize, LpProblem, LpVariable, lpSum

from blend_optimizer.data import BlendData
from blend_optimizer.linear_rows import GREATER_EQUAL, LinearRow, build_linear_rows, violation_scale

SUM_ROW_NAME = "Sum_of_fractions"


@dataclass
class BlendModel:
    """A built PuLP problem plus the maps needed to read its results.

    PuLP may change names (spaces become "_"), so we never look things up
    by PuLP name. The dicts below map *our* names to the PuLP objects.

    Keep this object alive while reading results: in PuLP 4 the variables
    only work while their LpProblem exists.

    Attributes:
        problem: The PuLP problem.
        data: The data it was built from.
        fractions: Raw material name -> its variable x_i.
        constraints: Row name -> PuLP constraint. Includes SUM_ROW_NAME.
        rows: Row name -> the LinearRow it was built from (spec rows only).
        violations: Row name -> violation variable v_k (elastic model only).
        elastic: True for the elastic (diagnosis) version.
    """

    problem: LpProblem
    data: BlendData
    fractions: dict[str, LpVariable]
    constraints: dict[str, LpConstraint]
    rows: dict[str, LinearRow]
    violations: dict[str, LpVariable] = field(default_factory=dict)
    elastic: bool = False


def build_model(data: BlendData, elastic: bool = False) -> BlendModel:
    """Build the least-cost blend LP, or its elastic version.

    Args:
        data: Validated input data (from ``inputs.load_inputs``).
        elastic: Build the elastic model (see the module notes) instead:
            spec rows may be violated, and the total relative violation is
            minimised instead of the cost.

    Returns:
        The model, ready for ``solve_and_report.solve``.
    """
    used_names: set[str] = set()
    problem = LpProblem(safe_name(data.settings.product_name, used_names, prefix="blend"), LpMinimize)

    fractions = _add_fraction_variables(problem, data, used_names)
    constraints = {SUM_ROW_NAME: _add_sum_to_one(problem, fractions, used_names)}
    rows = {row.name: row for row in build_linear_rows(data)}
    spec_constraints, violations = _add_spec_rows(problem, fractions, rows, used_names, elastic)
    constraints.update(spec_constraints)
    # Future integer/binary steps would be added here, e.g.
    #   _add_minimum_if_used(problem, fractions, data, used_names)
    if elastic:
        _set_violation_objective(problem, violations, rows, data)
    else:
        _set_cost_objective(problem, fractions, data)

    return BlendModel(problem=problem, data=data, fractions=fractions, constraints=constraints, rows=rows,
                      violations=violations, elastic=elastic)


# ---------------------------------------------------------------------------
# Building steps
# ---------------------------------------------------------------------------

def _add_fraction_variables(
    problem: LpProblem, data: BlendData, used_names: set[str]
) -> dict[str, LpVariable]:
    """One continuous variable x_i per raw material, bounded by its usage limits.

    Usage limits are in wt % of the dry blend; the variables are fractions,
    hence the division by 100. The effective limit is the tighter of the
    plant-wide and product limits.
    """
    fractions: dict[str, LpVariable] = {}
    for name in data.material_names:
        limit = data.effective_limit(name)
        fractions[name] = problem.add_variable(
            safe_name(name, used_names, prefix="x"),
            lowBound=limit.min_pct / 100.0,
            upBound=limit.max_pct / 100.0,
        )
    return fractions


def _add_sum_to_one(
    problem: LpProblem, fractions: dict[str, LpVariable], used_names: set[str]
) -> LpConstraint:
    """The fractions of the dry blend add up to 1."""
    used_names.add(SUM_ROW_NAME)
    return problem.addConstraint(lpSum(fractions.values()) == 1, SUM_ROW_NAME)


def _add_spec_rows(
    problem: LpProblem,
    fractions: dict[str, LpVariable],
    rows: dict[str, LinearRow],
    used_names: set[str],
    elastic: bool = False,
) -> tuple[dict[str, LpConstraint], dict[str, LpVariable]]:
    """Copy each LinearRow into the problem as a named constraint.

    In the elastic model each row also gets a violation variable v >= 0 that
    lets it be missed: added to a >= row, subtracted from a <= row.

    Returns:
        (row name -> constraint, row name -> violation variable).
    """
    constraints: dict[str, LpConstraint] = {}
    violations: dict[str, LpVariable] = {}
    for row_name, row in rows.items():
        expression = lpSum(coefficient * fractions[material] for material, coefficient in row.coefficients.items())
        if elastic:
            violation = problem.add_variable(safe_name(f"viol_{row_name}", used_names, prefix="v"), lowBound=0)
            violations[row_name] = violation
            expression = expression + violation if row.sense == GREATER_EQUAL else expression - violation
        if row.sense == GREATER_EQUAL:
            constraint = expression >= row.rhs
        else:
            constraint = expression <= row.rhs
        constraints[row_name] = problem.addConstraint(constraint, safe_name(row_name, used_names, prefix="c"))
    return constraints, violations


def _set_cost_objective(problem: LpProblem, fractions: dict[str, LpVariable], data: BlendData) -> None:
    """Minimise cost in Rs per dry tonne of blend: sum(dry_cost_i * x_i)."""
    problem.setObjective(lpSum(data.raw_materials[name].dry_cost * x for name, x in fractions.items()))


def _set_violation_objective(
    problem: LpProblem, violations: dict[str, LpVariable], rows: dict[str, LinearRow], data: BlendData
) -> None:
    """Elastic model: minimise the total relative violation, sum(v_k / scale_k)."""
    problem.setObjective(lpSum(
        (1.0 / violation_scale(rows[name], data)) * violation for name, violation in violations.items()
    ))


# ---------------------------------------------------------------------------
# Names
# ---------------------------------------------------------------------------

def safe_name(text: str, used_names: set[str], prefix: str) -> str:
    """A name that is safe in PuLP and in an .lp file, and not used before.

    Anything other than letters, digits and "_" becomes "_" (so
    "Calcined alumina" -> "Calcined_alumina"). ``prefix`` is added in front
    for variables (``x_``), and for any name that would otherwise start with
    a digit or be empty. If two names end up the same, "_2", "_3", ... is
    appended. The name is recorded in ``used_names``.

    Args:
        text: The readable name (raw material, spec row, product).
        used_names: Names already taken in this problem; updated.
        prefix: "x" for variables, "c" for constraints, "blend" for the problem.
    """
    cleaned = re.sub(r"[^A-Za-z0-9_]+", "_", text).strip("_")
    if prefix == "x" or not cleaned or cleaned[0].isdigit():
        cleaned = f"{prefix}_{cleaned}" if cleaned else prefix
    candidate, counter = cleaned, 2
    while candidate.lower() in {used.lower() for used in used_names}:
        candidate = f"{cleaned}_{counter}"
        counter += 1
    used_names.add(candidate)
    return candidate

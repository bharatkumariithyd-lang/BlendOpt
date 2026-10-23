"""Turn each spec limit into one linear row:  sum(a_i * x_i)  (>= or <=)  rhs.

Here x_i is the dry mass fraction of raw material i in the blend (sum x_i = 1).
This module is pure Python, so the formulation can be read and tested
without PuLP; ``model.py`` just copies these rows into the LP.

Every spec bound has the same shape: a blend value written as a fraction

    blend value = sum(N_i * x_i) / sum(g_i * x_i)

compared with a limit L, where

    N_i  numerator: the summed numerator components of material i (wt %, dry)
    g_i  denominator weight, which depends on the kind of row:

           composition, dry basis          g_i = 1
           composition, ignited basis      g_i = 1 - LOI_i/100
           composition, as-received basis  g_i = 1 / (1 - moisture_i/100)
           ratio (any basis)               g_i = D_i, the summed denominator
                                                 components of material i

For a composition row, g_i is the mass of material i "on the spec basis"
per dry tonne: one dry tonne is (1 - LOI/100) t once ignited, and
1 / (1 - m/100) t as received. Summing g_i * x_i gives the blend mass on
that basis per dry tonne of blend, so dividing by it converts the dry
component mass into wt % on the spec basis.

A ratio needs no basis: numerator and denominator would both be divided by
the same blend mass, which cancels.

Linearising: the denominator is positive, so multiplying through by it keeps
the direction of the inequality:

    sum(N_i x_i) / sum(g_i x_i) >= L    <=>    sum((N_i - L * g_i) * x_i) >= 0

which is linear in x. (Same for <=.) See :func:`linearise`.

On a dry basis sum(g_i * x_i) = sum(x_i) = 1, so the row is written in the
simpler form  sum(N_i * x_i) >= L  instead. Both are the same constraint,
but this one makes the .lp file read like the hand calculation:
"CaO_min: 92 x_Lime + 1 x_Calcined_alumina + 2 x_Bauxite >= 50".
"""

from dataclasses import dataclass

from blend_optimizer import config
from blend_optimizer.data import BlendData, RawMaterial, SpecRow

MIN = "min"
MAX = "max"
GREATER_EQUAL = ">="
LESS_EQUAL = "<="


@dataclass(frozen=True)
class LinearRow:
    """One bound of one spec row, as a linear constraint.

    Attributes:
        name: Constraint name, ``<spec name>_min`` or ``<spec name>_max``.
        spec: The spec row it comes from.
        bound: ``"min"`` or ``"max"``.
        sense: ``">="`` or ``"<="``.
        coefficients: Raw material name -> a_i (the LP coefficient).
        rhs: Right-hand side of the LP row.
        limit: Internal target L (customer limit with the margin applied).
        customer_limit: The customer's limit as written in Specs.
        numerator: Raw material name -> N_i.
        denominator: Raw material name -> g_i.
        limit_on_rhs: True for the simple form sum(N_i x_i) >= L (dry
            composition rows); False for the linearised form
            sum((N_i - L g_i) x_i) >= 0 (all other rows).
    """

    name: str
    spec: SpecRow
    bound: str
    sense: str
    coefficients: dict[str, float]
    rhs: float
    limit: float
    customer_limit: float
    numerator: dict[str, float]
    denominator: dict[str, float]
    limit_on_rhs: bool


def component_sum(material: RawMaterial, components: tuple[str, ...]) -> float:
    """Sum of the given components in one raw material, wt % dry basis."""
    return sum(material.composition[component] for component in components)


def basis_factor(material: RawMaterial, spec_basis: str) -> float:
    """g_i for a composition row: mass on the spec basis per dry tonne of material i.

    * dry:          1 dry tonne is 1 t                      -> 1
    * ignited:      it loses LOI on ignition                -> 1 - LOI/100
    * as_received:  it carries free moisture as delivered   -> 1 / (1 - moisture/100)

    Args:
        material: The raw material.
        spec_basis: One of ``config.SPEC_BASES``.
    """
    if spec_basis == config.BASIS_DRY:
        return 1.0
    if spec_basis == config.BASIS_IGNITED:
        return 1.0 - material.loi_pct / 100.0
    if spec_basis == config.BASIS_AS_RECEIVED:
        return 1.0 / (1.0 - material.moisture_pct / 100.0)
    raise ValueError(f"Unknown spec basis {spec_basis!r}; expected one of {config.SPEC_BASES}")


def linearise(numerator: dict[str, float], denominator: dict[str, float], limit: float) -> dict[str, float]:
    """Coefficients a_i = N_i - L * g_i of the row  sum(a_i * x_i)  (>= or <=)  0.

    This is  sum(N_i x_i) / sum(g_i x_i)  compared with L, multiplied through
    by the (positive) denominator. Example, basicity CaO/SiO2 >= 8:
    a_i = CaO_i - 8 * SiO2_i.

    Args:
        numerator: Material -> N_i.
        denominator: Material -> g_i (must be > 0 for the blend as a whole).
        limit: The limit L.
    """
    return {name: numerator[name] - limit * denominator[name] for name in numerator}


def build_linear_rows(data: BlendData) -> list[LinearRow]:
    """One LinearRow per spec bound (a spec with Min and Max gives two rows).

    * Composition row, dry basis:   sum(N_i x_i) >= L   (rhs = L)
    * Composition row, other basis: sum((N_i - L g_i) x_i) >= 0
    * Ratio row (any basis):        sum((N_i - R D_i) x_i) >= 0
    (and the same with <= for a Max.) L and R include the margin.

    Args:
        data: Validated input data.

    Returns:
        The rows in Specs order, Min before Max for each spec.
    """
    materials = data.raw_materials
    rows: list[LinearRow] = []
    for spec in data.specs:
        numerator = {name: component_sum(material, spec.numerator) for name, material in materials.items()}
        if spec.is_ratio:
            denominator = {name: component_sum(material, spec.denominator) for name, material in materials.items()}
        else:
            denominator = {name: basis_factor(material, data.settings.spec_basis)
                           for name, material in materials.items()}
        # Only a dry-basis composition row has g_i = 1, where sum(g_i x_i) = 1
        simple_form = not spec.is_ratio and data.settings.spec_basis == config.BASIS_DRY

        for bound, limit, customer_limit, sense in (
            (MIN, spec.min_target, spec.min_value, GREATER_EQUAL),
            (MAX, spec.max_target, spec.max_value, LESS_EQUAL),
        ):
            if limit is None:
                continue
            if simple_form:
                coefficients, rhs = dict(numerator), limit
            else:
                coefficients, rhs = linearise(numerator, denominator, limit), 0.0
            rows.append(LinearRow(
                name=f"{spec.name}_{bound}",
                spec=spec,
                bound=bound,
                sense=sense,
                coefficients=coefficients,
                rhs=rhs,
                limit=limit,
                customer_limit=customer_limit,
                numerator=numerator,
                denominator=denominator,
                limit_on_rhs=simple_form,
            ))
    return rows


def violation_scale(row: LinearRow, data: BlendData) -> float:
    """How big a "whole limit" is in this row's LP units (weights the elastic model).

    The elastic model lets each row be missed by a violation v >= 0 and
    minimises the total *relative* violation, sum(v / scale). For that to be
    relative, v must be divided by the size of the limit, in v's own unit:

    * Simple form, sum(N_i x_i) >= L: v is in the limit's unit (wt %), so
      scale = |L|.
    * Linearised form, sum((N_i - L g_i) x_i) >= 0: the left-hand side equals
      (blend value - L) * sum(g_i x_i), so v is in the limit's unit times
      sum(g_i x_i). That sum depends on the recipe, so a typical value is
      used: the average g_i of the raw materials that may be used.
      scale = |L| * typical g.  (For ignited/as-received rows g is near 1;
      for a ratio such as CaO/SiO2 it is a typical SiO2 content.)

    |L| is floored at ``config.ELASTIC_MIN_LIMIT`` so that a limit of 0 does
    not divide by zero.
    """
    size = max(abs(row.limit), config.ELASTIC_MIN_LIMIT)
    if row.limit_on_rhs:
        return size
    allowed = [name for name in data.material_names if data.effective_limit(name).max_pct > 0]
    weights = [row.denominator[name] for name in allowed if row.denominator[name] > 0]
    typical = sum(weights) / len(weights) if weights else 1.0
    return size * typical


def row_activity(row: LinearRow, fractions: dict[str, float]) -> float:
    """Left-hand side of the LP row, sum(a_i * x_i), for a given recipe."""
    return sum(coefficient * fractions[name] for name, coefficient in row.coefficients.items())


def blend_value(numerator: dict[str, float], denominator: dict[str, float],
                fractions: dict[str, float]) -> float:
    """Blend value sum(N_i x_i) / sum(g_i x_i) for a given recipe (fractions sum to 1).

    A ratio whose denominator is zero in the blend (e.g. CaO/SiO2 with no
    SiO2 at all) is returned as infinity, or NaN if the numerator is zero too.
    """
    top = sum(numerator[name] * x for name, x in fractions.items())
    bottom = sum(denominator[name] * x for name, x in fractions.items())
    if bottom == 0:
        return float("inf") if top > 0 else float("nan")
    return top / bottom


def achieved_value(row: LinearRow, fractions: dict[str, float]) -> float:
    """The value of the limited quantity in the blend, in the same unit as the limit."""
    return blend_value(row.numerator, row.denominator, fractions)

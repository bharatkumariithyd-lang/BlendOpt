"""Stage 3 tests: spec bases, ratio rows, margins and UsageLimits. Test 2 must pass."""

from dataclasses import replace

import numpy as np
import pandas as pd
import pytest

from blend_optimizer import config, examples
from blend_optimizer import solve_and_report as sr
from blend_optimizer.data import RawMaterial, UsageLimit
from blend_optimizer.inputs import load_inputs
from blend_optimizer.linear_rows import (
    GREATER_EQUAL,
    achieved_value,
    basis_factor,
    build_linear_rows,
    linearise,
    row_activity,
)
from blend_optimizer.model import SUM_ROW_NAME, build_model


def _solve(write_example, example, file_name="product.xlsx"):
    data = load_inputs(write_example(example, file_name))
    return data, sr.solve(build_model(data))


def _binding(result):
    return {name for name, slack in result.row_slack.items()
            if slack < config.BINDING_TOLERANCE and name != SUM_ROW_NAME}


def _with_spec(example, name, **cells):
    """Copy of ``example`` with cells of one Specs row changed, e.g. Min=12."""
    specs = example.specs.copy()
    for column, value in cells.items():
        specs.loc[specs[config.COL_NAME] == name, column] = value
    return replace(example, specs=specs)


# ===========================================================================
# Independent checks (written out here, NOT using linear_rows)
# ===========================================================================

def independent_value(data, fractions, spec):
    """Blend value of a spec, computed straight from the definitions."""
    rm = data.raw_materials

    def value(name, terms):  # sum(factor x wt %) for one material
        return sum(term.factor * rm[name].composition[term.component] for term in terms)

    top = sum(value(n, spec.numerator) * x for n, x in fractions.items())
    basis = data.settings.spec_basis
    if spec.is_ratio:        # mass of denominator components (basis cancels)
        bottom = sum(value(n, spec.denominator) * x for n, x in fractions.items())
    elif basis == "dry":     # dry blend mass
        bottom = sum(fractions.values())
    elif basis == "ignited":  # blend mass after ignition
        bottom = sum(x * (1 - rm[n].loi_pct / 100) for n, x in fractions.items())
    else:                    # blend mass as received (with moisture)
        bottom = sum(x / (1 - rm[n].moisture_pct / 100) for n, x in fractions.items())
    return top / bottom


def assert_specs_met(data, result, tol=1e-7):
    """Every spec target is met, and the report shows the independently computed value."""
    check = sr.spec_check_table(data, result).set_index("Spec")
    for spec in data.specs:
        value = independent_value(data, result.fractions, spec)
        assert check.loc[spec.name, "Achieved"] == pytest.approx(value, rel=1e-9)
        if spec.min_target is not None:
            assert value >= spec.min_target - tol, spec.name
        if spec.max_target is not None:
            assert value <= spec.max_target + tol, spec.name


def assert_optimal_by_duality(data, result, tol=1e-5):
    """Prove the recipe is optimal without trusting the solver's word for it.

    For a minimisation LP, a feasible x is optimal if there are duals pi
    (one per row) such that, with reduced costs dj_i = c_i - sum_k pi_k a_ki,
      * >= rows have pi >= 0, <= rows have pi <= 0, and pi = 0 wherever
        the row has slack (complementary slackness);
      * dj >= 0 where x_i sits at its lower bound, dj <= 0 at its upper
        bound, dj = 0 in between.
    We check these conditions on the duals returned by the solver.
    """
    fractions, pi = result.fractions, result.row_dual
    for name, material in data.raw_materials.items():
        dj = material.dry_cost - pi[SUM_ROW_NAME] - sum(
            pi[row_name] * row.coefficients[name] for row_name, row in result.rows.items())
        assert dj == pytest.approx(result.reduced_costs[name], abs=tol * 100)
        limit = data.effective_limit(name)
        x = 100 * fractions[name]
        at_lower, at_upper = x <= limit.min_pct + 1e-7, x >= limit.max_pct - 1e-7
        if at_lower and not at_upper:
            assert dj >= -tol, name
        elif at_upper and not at_lower:
            assert dj <= tol, name
        elif not at_lower and not at_upper:
            assert abs(dj) <= tol * 100, name
    for row_name, row in result.rows.items():
        if row.sense == GREATER_EQUAL:
            assert pi[row_name] >= -tol, row_name
        else:
            assert pi[row_name] <= tol, row_name
        assert abs(pi[row_name] * result.row_slack[row_name]) <= tol * 100, row_name


# ===========================================================================
# Test 2: the synthetic slag, all variants
# ===========================================================================

EXPECTED_DRY_RECIPE = {"Lime": 52.09, "Lime fines": 0.0, "Calcined dolomite": 2.60,
                       "Calcined bauxite": 17.13, "Raw bauxite": 25.83, "Fluorspar": 2.35}


def test_synthetic_slag_dry(write_example):
    data, result = _solve(write_example, examples.synthetic_slag_example())
    assert result.is_optimal
    assert result.cost_per_dry_t == pytest.approx(13004, abs=1)
    for name, expected in EXPECTED_DRY_RECIPE.items():
        assert 100 * result.fractions[name] == pytest.approx(expected, abs=0.05), name
    assert _binding(result) == {"CaO_min", "Al2O3_min", "CaF2_min", "P_max"}
    assert_specs_met(data, result)
    assert_optimal_by_duality(data, result)


def test_synthetic_slag_ignited(write_example):
    data, result = _solve(write_example, examples.synthetic_slag_example(spec_basis="ignited"))
    assert result.cost_per_dry_t == pytest.approx(13143, abs=1)
    assert "MgO_max" in _binding(result)
    check = sr.spec_check_table(data, result).set_index("Spec")
    assert check.loc["MgO", "Achieved"] == pytest.approx(8.0, abs=0.005)
    assert check.loc["MgO", "Status"] == "Binding (Max)"
    assert check.loc["MgO", "Basis"] == "ignited"
    assert_specs_met(data, result)
    assert_optimal_by_duality(data, result)


def test_synthetic_slag_margin(write_example):
    data, result = _solve(write_example, examples.synthetic_slag_example(margin=0.5))
    assert result.cost_per_dry_t == pytest.approx(13330, abs=1)
    check = sr.spec_check_table(data, result).set_index("Spec")
    assert (check.loc["CaO", "Customer Min"], check.loc["CaO", "Target Min"]) == (50, 50.5)
    assert (check.loc["Al2O3", "Customer Max"], check.loc["Al2O3", "Target Max"]) == (38, 37.5)
    assert check.loc["CaO", "Achieved"] == pytest.approx(50.5)
    assert_specs_met(data, result)
    assert_optimal_by_duality(data, result)


def test_synthetic_slag_as_received(write_example):
    """No expected value was given for this basis, so prove the answer instead."""
    data, result = _solve(write_example, examples.synthetic_slag_example(spec_basis="as_received"))
    assert result.is_optimal
    assert_specs_met(data, result)
    assert_optimal_by_duality(data, result)

    # On an as-received basis each component is its dry value divided by the
    # as-received tonnes per dry tonne of blend
    wet_per_dry = sr.as_received_tonnes_per_dry_tonne(data, result.fractions)
    composition = sr.composition_table(data, result).set_index("Component")["Blend (wt %)"]
    dry_cao = sum(data.raw_materials[n].composition["CaO"] * x for n, x in result.fractions.items())
    assert composition["CaO"] == pytest.approx(dry_cao / wet_per_dry)


def test_fluorspar_ban_is_infeasible(write_example):
    data, result = _solve(write_example, examples.synthetic_slag_example(ban_fluorspar=True))
    assert data.effective_limit("Fluorspar") == UsageLimit(0, 0)
    assert result.status == "Infeasible"


# ===========================================================================
# Ratio rows and margins
# ===========================================================================

def test_ratio_row_binds_with_margin(write_example):
    # Base dry basicity is 11.40 (slack). Customer Min 11.5 + margin 0.5 -> target 12.
    example = _with_spec(examples.synthetic_slag_example(), "Basicity", Min=11.5, Margin=0.5)
    data, result = _solve(write_example, example)
    check = sr.spec_check_table(data, result).set_index("Spec")
    assert check.loc["Basicity", "Basis"] == "ratio (any basis)"
    assert check.loc["Basicity", "Target Min"] == 12
    assert check.loc["Basicity", "Achieved"] == pytest.approx(12)
    assert check.loc["Basicity", "Status"] == "Binding (Min)"
    assert result.cost_per_dry_t > 13004
    assert_specs_met(data, result)
    assert_optimal_by_duality(data, result)


def test_ratio_max_row(write_example):
    # Base basicity is 11.40; with the other limits held the lowest achievable is 11.24
    example = _with_spec(examples.synthetic_slag_example(), "Basicity", Max=11.3)
    data, result = _solve(write_example, example)
    check = sr.spec_check_table(data, result).set_index("Spec")
    assert check.loc["Basicity", "Achieved"] == pytest.approx(11.3)
    assert check.loc["Basicity", "Status"] == "Binding (Max)"
    assert_specs_met(data, result)
    assert_optimal_by_duality(data, result)

    too_low = _with_spec(examples.synthetic_slag_example(), "Basicity", Max=11)
    assert _solve(write_example, too_low, "too_low.xlsx")[1].status == "Infeasible"


def test_ratio_does_not_depend_on_basis(write_example):
    rows = {}
    for basis in config.SPEC_BASES:
        data = load_inputs(write_example(examples.synthetic_slag_example(spec_basis=basis), f"{basis}.xlsx"))
        rows[basis] = next(row for row in build_linear_rows(data) if row.name == "Basicity_min")
    assert rows["dry"].coefficients == rows["ignited"].coefficients == rows["as_received"].coefficients
    assert rows["dry"].coefficients["Lime"] == pytest.approx(92 - 8 * 1.5)

    # In the ignited result, basicity is the ratio of the ignited CaO and SiO2 values
    data, result = _solve(write_example, examples.synthetic_slag_example(spec_basis="ignited"), "ign2.xlsx")
    composition = sr.composition_table(data, result).set_index("Component")["Blend (wt %)"]
    check = sr.spec_check_table(data, result).set_index("Spec")
    assert check.loc["Basicity", "Achieved"] == pytest.approx(composition["CaO"] / composition["SiO2"])


def test_margin_on_dry_row_costs_the_shadow_price(write_example):
    base_data, base = _solve(write_example, examples.hand_worked_example(), "base.xlsx")
    example = _with_spec(examples.hand_worked_example(), "CaO", Margin=1.0)
    data, result = _solve(write_example, example, "margin.xlsx")
    check = sr.spec_check_table(data, result).set_index("Spec")
    assert check.loc["CaO", "Target Min"] == 51
    assert check.loc["CaO", "Achieved"] == pytest.approx(51)
    # The LP is linear while the same limits bind: 1 % more CaO costs one shadow price
    assert result.cost_per_dry_t - base.cost_per_dry_t == pytest.approx(450.479, abs=0.01)


# ===========================================================================
# UsageLimits in the model
# ===========================================================================

def test_product_max_tighter_than_plant(write_example):
    example = replace(examples.synthetic_slag_example(),
                      usage_limits=examples.usage_limit_table([("Raw bauxite", None, 20)]))
    data, result = _solve(write_example, example)
    assert data.effective_limit("Raw bauxite") == UsageLimit(0, 20)
    assert 100 * result.fractions["Raw bauxite"] == pytest.approx(20)
    assert result.cost_per_dry_t > 13004
    assert_specs_met(data, result)
    assert_optimal_by_duality(data, result)


def test_product_min_forces_a_material_in(write_example):
    example = replace(examples.synthetic_slag_example(),
                      usage_limits=examples.usage_limit_table([("Lime fines", 5, None)]))
    data, result = _solve(write_example, example)
    assert data.effective_limit("Lime fines") == UsageLimit(5, 20)  # plant max 20 still applies
    assert 100 * result.fractions["Lime fines"] == pytest.approx(5)
    assert result.cost_per_dry_t > 13004
    assert_optimal_by_duality(data, result)


# ===========================================================================
# The basis conversion and linearisation themselves
# ===========================================================================

def test_basis_factor():
    material = RawMaterial("M", 1000, "dry", moisture_pct=6.0, loi_pct=24.0,
                           plant_limit=UsageLimit(), composition={"A": 50.0})
    assert basis_factor(material, "dry") == 1.0
    assert basis_factor(material, "ignited") == pytest.approx(0.76)
    assert basis_factor(material, "as_received") == pytest.approx(1 / 0.94)
    with pytest.raises(ValueError):
        basis_factor(material, "wet")


def test_linearise():
    assert linearise({"a": 50.0, "b": 10.0}, {"a": 1.0, "b": 0.8}, 20.0) == {"a": 30.0, "b": -6.0}


@pytest.mark.parametrize("basis", config.SPEC_BASES)
def test_linear_row_sign_matches_the_limit(write_example, basis):
    """For any recipe, the linear row is satisfied exactly when the real limit is.

    activity - rhs = (achieved - L) * sum(g_i x_i), and sum(g_i x_i) > 0,
    so both have the same sign.
    """
    data = load_inputs(write_example(examples.synthetic_slag_example(spec_basis=basis)))
    rng = np.random.default_rng(seed=1)
    names = data.material_names
    for _ in range(25):
        fractions = dict(zip(names, rng.dirichlet(np.ones(len(names)))))
        for row in build_linear_rows(data):
            achieved = achieved_value(row, fractions)
            assert achieved == pytest.approx(independent_value(data, fractions, row.spec))
            weight = sum(row.denominator[n] * x for n, x in fractions.items())
            lhs_minus_rhs = row_activity(row, fractions) - row.rhs
            expected = (achieved - row.limit) * (1.0 if row.limit_on_rhs else weight)
            assert lhs_minus_rhs == pytest.approx(expected, abs=1e-9)

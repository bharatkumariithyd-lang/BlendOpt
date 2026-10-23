"""Stage 5 tests: what-if scenarios and price ranging."""

import math

import numpy as np
import openpyxl
import pandas as pd
import pytest

from blend_optimizer import examples
from blend_optimizer import solve_and_report as sr
from blend_optimizer.inputs import load_inputs
from blend_optimizer.pipeline import solve_blend
from blend_optimizer.what_if import apply_changes, format_what_if, price_ranging, what_if


@pytest.fixture
def hand_worked(write_example):
    return load_inputs(write_example(examples.hand_worked_example(), "test1.xlsx"))


@pytest.fixture
def slag(write_example):
    return load_inputs(write_example(examples.synthetic_slag_example(), "test2.xlsx"))


# ===========================================================================
# apply_changes
# ===========================================================================

def test_apply_changes_leaves_the_base_case_alone(slag):
    scenario = apply_changes(slag, price_changes={"Lime": 9500}, spec_changes={"CaO": {"Min": 51}},
                             usage_changes={"Fluorspar": {"Max": 3}})
    # Base case untouched
    assert slag.raw_materials["Lime"].cost_input == 9000
    assert slag.specs[0].min_value == 50
    assert slag.product_limits == {}
    # Scenario changed; the dry cost follows the new as-received price automatically
    assert scenario.raw_materials["Lime"].dry_cost == pytest.approx(9500 / 0.995)
    assert scenario.specs[0].min_value == 51 and scenario.specs[0].max_value == 56
    assert scenario.effective_limit("Fluorspar").max_pct == 3


@pytest.mark.parametrize("changes, message", [
    ({"price_changes": {"Lyme": 9500}}, "Unknown raw material 'Lyme'. Did you mean 'Lime'?"),
    ({"price_changes": {"Lime": -1}}, "must be a number of 0 or more"),
    ({"spec_changes": {"AL2O3": {"Min": 32}}}, "Unknown spec 'AL2O3'. Did you mean 'Al2O3'?"),
    ({"spec_changes": {"CaO": {"Minimum": 51}}}, "unknown field(s) Minimum"),
    ({"spec_changes": {"CaO": {"Min": 57}}}, "Min (57) is greater than Max (56)"),
    ({"spec_changes": {"SiO2": {"Max": None}}}, "give a Min, a Max, or both"),
    ({"spec_changes": {"CaO": {"Margin": 4}}}, "after the margin the internal range is empty"),
    ({"usage_changes": {"Lime": {"Min": 60, "Max": 50}}}, "need 0 <= Min <= Max <= 100"),
    ({"usage_changes": {"Lime": {"Max": 120}}}, "need 0 <= Min <= Max <= 100"),
    ({"usage_changes": {"Lime fines": {"Min": 30}}}, "conflicts with the plant-wide limit"),
    ({"usage_changes": {"Lime": {"Maximum": 50}}}, "unknown field(s) Maximum"),
])
def test_apply_changes_rejects_bad_changes(slag, changes, message):
    with pytest.raises(ValueError) as info:
        apply_changes(slag, **changes)
    assert message in str(info.value)


# ===========================================================================
# what_if
# ===========================================================================

def test_no_changes_gives_identical_cases(slag):
    tables = what_if(slag)
    assert tables["Changes"].empty
    cost = tables["Cost"].set_index("Item")
    assert cost.loc["Cost (Rs per dry tonne of blend)", "Change"] == pytest.approx(0, abs=1e-9)
    recipe = tables["Recipe"].set_index("Raw material")
    assert recipe["Change (wt %)"].abs().max() == pytest.approx(0, abs=1e-9)


def test_price_change_with_recipe_unchanged(hand_worked):
    """Lime's recipe share is fixed by the limits, so a price rise just costs x_lime per Rs."""
    base_lime = solve_blend(hand_worked).fractions["Lime"]
    tables = what_if(hand_worked, price_changes={"Lime": 9500})
    cost = tables["Cost"].set_index("Item")
    assert cost.loc["Cost (Rs per dry tonne of blend)", "Change"] == pytest.approx(500 * base_lime)
    assert tables["Recipe"]["Change (wt %)"].abs().max() == pytest.approx(0, abs=1e-9)
    change = tables["Changes"].iloc[0]
    assert (change["Change"], change["Name"], change["Base value"], change["Scenario value"]) == \
           ("Price", "Lime", 9000, 9500)


def test_spec_change_matches_a_workbook_with_that_spec(hand_worked, write_example):
    tables = what_if(hand_worked, spec_changes={"Al2O3": {"Min": 32}})
    # The same case built the normal way, from a workbook
    direct = solve_blend(load_inputs(write_example(examples.hand_worked_example(al2o3_min=32), "al32.xlsx")))
    cost = tables["Cost"].set_index("Item")
    assert cost.loc["Cost (Rs per dry tonne of blend)", "Scenario"] == pytest.approx(direct.cost_per_dry_t)
    spec = tables["Spec check"].set_index("Spec")
    assert spec.loc["Al2O3", "Scenario achieved"] == pytest.approx(32)
    assert spec.loc["Al2O3", "Scenario status"] == "Binding (Min)"
    assert spec.loc["Al2O3", "Customer Min"] == 32


def test_margin_change_reproduces_the_margin_variant(slag):
    tables = what_if(slag, spec_changes={"CaO": {"Margin": 0.5}, "Al2O3": {"Margin": 0.5}})
    cost = tables["Cost"].set_index("Item")
    assert cost.loc["Cost (Rs per dry tonne of blend)", "Base"] == pytest.approx(13004, abs=1)
    assert cost.loc["Cost (Rs per dry tonne of blend)", "Scenario"] == pytest.approx(13330, abs=1)


def test_removing_a_slack_limit_changes_nothing(slag):
    tables = what_if(slag, spec_changes={"CaO": {"Max": None}})
    assert tables["Cost"].set_index("Item").loc["Cost (Rs per dry tonne of blend)", "Change"] == \
           pytest.approx(0, abs=1e-6)
    assert pd.isna(tables["Changes"].iloc[0]["Scenario value"])


def test_binding_limits_side_by_side(slag):
    tables = what_if(slag, spec_changes={"P": {"Max": 0.035}}, usage_changes={"Raw bauxite": {"Max": 20}})
    binding = tables["Binding limits"].set_index("Constraint")
    assert binding.loc["P_max", "Base saving per step (Rs/dry t)"] == pytest.approx(7350, abs=5)
    assert pd.isna(binding.loc["P_max", "Scenario saving per step (Rs/dry t)"])   # no longer binding
    assert pd.isna(binding.loc["Raw bauxite max use", "Base saving per step (Rs/dry t)"])
    assert binding.loc["Raw bauxite max use", "Scenario saving per step (Rs/dry t)"] > 0


def test_infeasible_scenario_is_shown_not_raised(slag):
    tables = what_if(slag, usage_changes={"Fluorspar": {"Max": 0}})
    cost = tables["Cost"].set_index("Item")
    assert cost.loc["Solver status", "Scenario"] == "Infeasible"
    assert pd.isna(cost.loc["Cost (Rs per dry tonne of blend)", "Scenario"])
    recipe = tables["Recipe"]
    assert recipe["Scenario (wt %)"].isna().all()
    assert (tables["Spec check"]["Scenario status"] == "Infeasible").all()
    shown = sr.display_table(recipe)
    assert (shown["Scenario (wt %)"] == "").all()  # blanks, not NaN


def test_what_if_display_and_excel(slag, tmp_path):
    tables = what_if(slag, price_changes={"Fluorspar": 35000}, spec_changes={"P": {"Max": 0.035}})
    text = format_what_if(tables)
    assert "Spec Max" in text and "0.0300" in text and "0.0350" in text  # typed decimals + 1
    assert "nan" not in text.lower()

    path = sr.write_excel(tables, tmp_path / "what_if.xlsx")
    workbook = openpyxl.load_workbook(path)
    assert workbook.sheetnames == ["Changes", "Cost", "Recipe", "Spec check", "Binding limits"]
    recipe = workbook["Recipe"]
    headers = [cell.value for cell in recipe[1]]
    column = headers.index("Scenario (kg/batch, as received)") + 1
    rows = [recipe.cell(row, column).value for row in range(2, 8)]
    assert round(sum(rows), 6) == recipe.cell(8, column).value == 1000  # weigh sheet adds up


# ===========================================================================
# price_ranging
# ===========================================================================

def _hand_worked_breakpoints():
    """Exact Test 1 price limits, worked out from the two recipes involved.

    Today's recipe V1 has CaO_min and Al2O3_min binding. When alumina gets
    cheap enough, or bauxite dear enough, the optimum moves to V3: no bauxite,
    Al2O3 = 30 % from lime + alumina (x_A = 29/84, x_L = 55/84). The price
    limit is where V1 and V3 cost the same.
    """
    a = np.array([[1, 1, 1], [92, 1, 2], [1, 85, 60]])
    lime1, alumina1, bauxite1 = np.linalg.solve(a, [1, 50, 30])
    lime3, alumina3 = 55 / 84, 29 / 84
    alumina_lower = (9000 * (lime1 - lime3) + 12000 * bauxite1) / (alumina3 - alumina1)
    bauxite_upper = (9000 * (lime3 - lime1) + 30000 * (alumina3 - alumina1)) / bauxite1
    return alumina_lower, bauxite_upper


def test_price_ranging_hand_worked(hand_worked):
    table = price_ranging(hand_worked).set_index("Raw material")
    alumina_lower, bauxite_upper = _hand_worked_breakpoints()

    # Alumina: unchanged at about Rs 13,300 and above
    assert table.loc["Calcined alumina", "Unchanged from (Rs/t)"] == pytest.approx(alumina_lower, abs=1.0)
    assert pd.isna(table.loc["Calcined alumina", "Unchanged up to (Rs/t)"])
    assert round(alumina_lower, -2) == 13300
    # Bauxite: unchanged up to about Rs 23,700 (exactly 23,750)
    assert table.loc["Bauxite", "Unchanged up to (Rs/t)"] == pytest.approx(bauxite_upper, abs=1.0)
    assert pd.isna(table.loc["Bauxite", "Unchanged from (Rs/t)"])
    assert bauxite_upper == pytest.approx(23750, abs=1)
    # Lime: unchanged at any price tested
    assert pd.isna(table.loc["Lime", "Unchanged from (Rs/t)"])
    assert pd.isna(table.loc["Lime", "Unchanged up to (Rs/t)"])
    assert table.loc["Lime", "Note"].startswith("Recipe unchanged at every price tested")


def test_price_ranging_ends_are_where_the_recipe_changes(hand_worked):
    table = price_ranging(hand_worked).set_index("Raw material")
    base = solve_blend(hand_worked).fractions

    def changed(name, price):
        recipe = solve_blend(apply_changes(hand_worked, price_changes={name: price})).fractions
        return max(abs(recipe[n] - base[n]) for n in base) > 1e-6

    lower = table.loc["Calcined alumina", "Unchanged from (Rs/t)"]
    assert not changed("Calcined alumina", lower) and changed("Calcined alumina", lower - 2)
    upper = table.loc["Bauxite", "Unchanged up to (Rs/t)"]
    assert not changed("Bauxite", upper) and changed("Bauxite", upper + 2)


def test_price_ranging_tolerance(hand_worked):
    alumina_lower, _ = _hand_worked_breakpoints()
    table = price_ranging(hand_worked, tolerance_rs=0.01).set_index("Raw material")
    assert table.loc["Calcined alumina", "Unchanged from (Rs/t)"] == pytest.approx(alumina_lower, abs=0.01)


def test_price_ranging_agrees_with_reduced_costs(slag):
    """For an unused material the lower end is its entry price (as-received basis)."""
    ranges = price_ranging(slag).set_index("Raw material")
    reduced = sr.reduced_cost_table(slag, solve_blend(slag)).set_index("Raw material")
    entry = reduced.loc["Lime fines", "Enters below (Rs/t as entered)"]
    assert ranges.loc["Lime fines", "Unchanged from (Rs/t)"] == pytest.approx(entry, abs=1.0)
    assert pd.isna(ranges.loc["Lime fines", "Unchanged up to (Rs/t)"])
    assert "enters the recipe below" in ranges.loc["Lime fines", "Note"]
    assert ranges.loc["Lime fines", "Price basis"] == "as_received"


def test_price_ranging_brackets_the_current_price(slag):
    for _, row in price_ranging(slag).iterrows():
        lower, upper, price = row["Unchanged from (Rs/t)"], row["Unchanged up to (Rs/t)"], row["Current price (Rs/t)"]
        assert pd.isna(lower) or lower <= price
        assert pd.isna(upper) or upper >= price


def test_price_ranging_needs_a_feasible_base(write_example):
    data = load_inputs(write_example(examples.hand_worked_example(al2o3_min=45)))
    with pytest.raises(ValueError, match="needs a feasible recipe"):
        price_ranging(data)


def test_scenario_can_fix_an_infeasible_base(write_example):
    """Stage 7: a what-if may be the fix for an infeasible base case, so it is compared, not refused."""
    data = load_inputs(write_example(examples.hand_worked_example(al2o3_min=45)))
    tables = what_if(data, spec_changes={"Al2O3": {"Min": 39}})
    cost = tables["Cost"].set_index("Item")
    assert (cost.loc["Solver status", "Base"], cost.loc["Solver status", "Scenario"]) == ("Infeasible", "Optimal")
    assert pd.isna(cost.loc["Cost (Rs per dry tonne of blend)", "Base"])
    assert pd.isna(cost.loc["Cost (Rs per dry tonne of blend)", "Change"])
    spec = tables["Spec check"].set_index("Spec")
    assert spec.loc["Al2O3", "Base status"] == "Infeasible"
    assert spec.loc["Al2O3", "Scenario achieved"] == pytest.approx(39)


def test_price_ranging_display(hand_worked):
    shown = sr.display_table(price_ranging(hand_worked)).set_index("Raw material")
    assert shown.loc["Calcined alumina", "Unchanged from (Rs/t)"] == "13,271"
    assert shown.loc["Calcined alumina", "Unchanged up to (Rs/t)"] == ""
    assert not math.isnan(price_ranging(hand_worked).iloc[0]["Current price (Rs/t)"])

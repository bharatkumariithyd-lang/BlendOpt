"""Stage 2 tests: the LP on a dry basis. Test 1 must be reproduced exactly."""

import gc
from dataclasses import replace
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

from blend_optimizer import config, examples
from blend_optimizer import solve_and_report as sr
from blend_optimizer.inputs import load_inputs
from blend_optimizer.linear_rows import build_linear_rows
from blend_optimizer.model import SUM_ROW_NAME, build_model, safe_name

PACKAGE = Path(__file__).resolve().parents[1] / "blend_optimizer"


def _solve(write_example, example, lp_path=None):
    data = load_inputs(write_example(example))
    return data, sr.solve(build_model(data), lp_path=lp_path)


def _by_name(table: pd.DataFrame, key_column: str) -> pd.DataFrame:
    return table.set_index(key_column)


# ===========================================================================
# Test 1: the hand-worked example
# ===========================================================================

def test_hand_worked_recipe_is_exact(write_example):
    data, result = _solve(write_example, examples.hand_worked_example())
    assert result.status == "Optimal" and result.is_optimal

    # The two binding limits plus sum = 1 fix the recipe: solve those 3 equations directly
    #           Lime  Alumina  Bauxite
    a = np.array([[1,    1,       1],    # sum of fractions = 1
                  [92,   1,       2],    # CaO   = 50
                  [1,    85,      60]])  # Al2O3 = 30
    exact = dict(zip(["Lime", "Calcined alumina", "Bauxite"], np.linalg.solve(a, [1, 50, 30])))
    for name, x in exact.items():
        assert result.fractions[name] == pytest.approx(x, abs=1e-9)
    exact_cost = 9000 * exact["Lime"] + 30000 * exact["Calcined alumina"] + 12000 * exact["Bauxite"]
    assert result.cost_per_dry_t == pytest.approx(exact_cost, abs=1e-6)

    # ... and the numbers from the hand calculation
    recipe = _by_name(sr.recipe_table(data, result), "Raw material")["Dry blend (wt %)"]
    assert round(recipe["Lime"], 2) == 53.40
    assert round(recipe["Calcined alumina"], 2) == 6.02
    assert round(recipe["Bauxite"], 2) == 40.58
    assert round(result.cost_per_dry_t) == 11482


def test_hand_worked_binding_and_slack(write_example):
    data, result = _solve(write_example, examples.hand_worked_example())
    check = _by_name(sr.spec_check_table(data, result), "Spec")

    assert check.loc["CaO", "Status"] == "Binding (Min)"
    assert check.loc["Al2O3", "Status"] == "Binding (Min)"
    assert check.loc["SiO2", "Status"] == "Slack"
    assert round(check.loc["SiO2", "Achieved"], 2) == 4.68
    assert check.loc["SiO2", "Slack to target"] == pytest.approx(6 - check.loc["SiO2", "Achieved"])
    assert check.loc["CaO", "Achieved"] == pytest.approx(50)
    assert result.row_slack["SiO2_max"] == pytest.approx(1.3245, abs=1e-4)


def test_hand_worked_composition_and_cost(write_example):
    data, result = _solve(write_example, examples.hand_worked_example())
    composition = _by_name(sr.composition_table(data, result), "Component")["Blend (wt %)"]
    assert composition["CaO"] == pytest.approx(50)
    assert composition["Al2O3"] == pytest.approx(30)
    assert composition["Sum of components"] == pytest.approx(100)  # "Others" closes the balance
    assert composition["LOI"] == 0 and composition["Moisture"] == 0

    cost = _by_name(sr.cost_table(data, result), "Item")["Value"]
    assert cost["Cost (Rs per dry tonne of blend)"] == pytest.approx(result.cost_per_dry_t)
    # No moisture: dry and as-received are the same
    assert cost["Cost (Rs per as-received tonne of blend)"] == pytest.approx(result.cost_per_dry_t)
    assert cost["Batch dry mass (kg)"] == pytest.approx(1000)
    assert cost["Cost per batch (Rs)"] == pytest.approx(result.cost_per_dry_t)


def test_activity_and_slack_are_consistent(write_example):
    _, result = _solve(write_example, examples.hand_worked_example())
    for name, row in result.rows.items():
        activity = result.row_activity[name]
        expected_slack = activity - row.rhs if row.sense == ">=" else row.rhs - activity
        assert result.row_slack[name] == pytest.approx(expected_slack, abs=1e-9)
    assert result.row_activity[SUM_ROW_NAME] == pytest.approx(1)


def test_lp_file_is_written_and_readable(write_example, tmp_path):
    _, result = _solve(write_example, examples.hand_worked_example(), lp_path=tmp_path / "out" / "test1.lp")
    text = Path(result.lp_file).read_text()
    assert "CaO_min: 2 x_Bauxite + x_Calcined_alumina + 92 x_Lime >= 50" in text
    assert "SiO2_max: 8 x_Bauxite + 6 x_Calcined_alumina + 2 x_Lime <= 6" in text
    assert "Sum_of_fractions: x_Bauxite + x_Calcined_alumina + x_Lime = 1" in text


def test_infeasible_is_reported_not_raised(write_example, tmp_path):
    data, result = _solve(write_example, examples.hand_worked_example(al2o3_min=45), lp_path=tmp_path / "inf.lp")
    assert result.status == "Infeasible"
    assert not result.is_optimal
    assert result.cost_per_dry_t is None and result.fractions == {}
    assert Path(result.lp_file).is_file()  # written before solving, for checking by hand
    with pytest.raises(ValueError, match="solver status is 'Infeasible'"):
        sr.recipe_table(data, result)


# ===========================================================================
# Model details
# ===========================================================================

def test_plant_max_use_is_a_variable_bound(write_example):
    table = examples.hand_worked_example().raw_materials.copy()
    table.loc[table[config.COL_NAME] == "Bauxite", config.COL_MAX_USE] = 30
    data, result = _solve(write_example, replace(examples.hand_worked_example(), raw_materials=table))

    assert result.fractions["Bauxite"] == pytest.approx(0.30)
    # With bauxite capped, Al2O3_min and the cap fix the recipe
    assert result.fractions["Lime"] == pytest.approx(0.565476, abs=1e-6)
    assert result.cost_per_dry_t == pytest.approx(12725.0)


def test_plant_min_use_is_a_variable_bound(write_example):
    table = examples.hand_worked_example().raw_materials.copy()
    table.loc[table[config.COL_NAME] == "Calcined alumina", config.COL_MIN_USE] = 10
    _, result = _solve(write_example, replace(examples.hand_worked_example(), raw_materials=table))
    assert result.fractions["Calcined alumina"] == pytest.approx(0.10)
    assert result.row_slack["Al2O3_min"] == pytest.approx(0, abs=1e-9)


def test_batch_is_as_received(write_example):
    # Lime with 10 % moisture (price on dry basis, so the recipe does not change)
    table = examples.hand_worked_example().raw_materials.copy()
    table.loc[table[config.COL_NAME] == "Lime", config.COL_MOISTURE] = 10
    data, result = _solve(write_example, replace(examples.hand_worked_example(), raw_materials=table))
    recipe = _by_name(sr.recipe_table(data, result), "Raw material")

    assert recipe.loc["TOTAL", "As-received mass (kg/batch)"] == pytest.approx(1000)
    lime = recipe.loc["Lime"]
    assert lime["As-received mass (kg/batch)"] == pytest.approx(lime["Dry mass (kg/batch)"] / 0.9)
    tonnes_wet_per_dry = result.fractions["Lime"] / 0.9 + result.fractions["Calcined alumina"] + result.fractions["Bauxite"]
    assert recipe.loc["TOTAL", "Dry mass (kg/batch)"] == pytest.approx(1000 / tonnes_wet_per_dry)

    cost = _by_name(sr.cost_table(data, result), "Item")["Value"]
    assert cost["Cost (Rs per as-received tonne of blend)"] == pytest.approx(result.cost_per_dry_t / tonnes_wet_per_dry)
    moisture = _by_name(sr.composition_table(data, result), "Component").loc["Moisture", "Blend (wt %)"]
    assert moisture == pytest.approx(100 * (tonnes_wet_per_dry - 1) / tonnes_wet_per_dry)


def test_result_survives_the_model(write_example):
    """SolveResult holds plain values, so it still works after the PuLP problem is gone."""
    data = load_inputs(write_example(examples.hand_worked_example()))
    model = build_model(data)
    result = sr.solve(model)
    del model
    gc.collect()
    assert isinstance(result.fractions["Lime"], float)
    assert sr.recipe_table(data, result).shape[0] == 4


def test_linear_rows_dry_basis(write_example):
    example = examples.hand_worked_example()
    specs = example.specs.copy()
    specs.loc[specs[config.COL_NAME] == "CaO", config.COL_MARGIN] = 1.0
    extra = examples.spec_table([("CaO+SiO2", "CaO+SiO2", None, None, 60, None)])
    data = load_inputs(write_example(replace(example, specs=pd.concat([specs, extra], ignore_index=True))))

    rows = {row.name: row for row in build_linear_rows(data)}
    assert list(rows) == ["CaO_min", "Al2O3_min", "SiO2_max", "CaO+SiO2_max"]
    cao = rows["CaO_min"]
    assert (cao.sense, cao.rhs, cao.limit, cao.customer_limit) == (">=", 51.0, 51.0, 50)  # margin applied
    assert cao.coefficients == {"Lime": 92, "Calcined alumina": 1, "Bauxite": 2}
    assert rows["SiO2_max"].sense == "<="
    assert rows["CaO+SiO2_max"].coefficients == {"Lime": 94, "Calcined alumina": 7, "Bauxite": 10}


def test_safe_names():
    used: set[str] = set()
    assert safe_name("Lime fines-2/x", used, prefix="x") == "x_Lime_fines_2_x"
    assert safe_name("A b", used, prefix="x") == "x_A_b"
    assert safe_name("A_b", used, prefix="x") == "x_A_b_2"  # would clash after cleaning
    assert safe_name("CaO+SiO2_max", used, prefix="c") == "CaO_SiO2_max"
    assert safe_name("2CaO_min", used, prefix="c") == "c_2CaO_min"  # must not start with a digit
    assert safe_name("%%%", used, prefix="c") == "c"


def test_pulp_only_in_model_and_solve_and_report():
    """The 'PuLP wall': no other module may import pulp."""
    allowed = {"model.py", "solve_and_report.py"}
    offenders = [
        path.name for path in PACKAGE.glob("*.py")
        if path.name not in allowed
        and any(line.lstrip().startswith(("import pulp", "from pulp")) for line in path.read_text().splitlines())
    ]
    assert offenders == []

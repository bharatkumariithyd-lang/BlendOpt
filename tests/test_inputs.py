"""Stage 1 tests: reading and validating the input workbooks."""

from dataclasses import replace

import openpyxl
import pandas as pd
import pytest

from blend_optimizer import config, examples
from blend_optimizer.data import (
    InputError,
    Settings,
    SpecRow,
    UsageLimit,
    price_from_dry_basis,
    price_to_dry_basis,
)
from blend_optimizer.data import Term
from blend_optimizer.inputs import (
    analysis_to_dry_basis,
    analysis_total_warning,
    load_inputs,
    parse_expression,
    suggest_close_match,
)


# ===========================================================================
# Helpers: copies of an example with one thing changed
# ===========================================================================

def _edit(table: pd.DataFrame, name: str, column: str, value: object) -> pd.DataFrame:
    """Copy of ``table`` with one cell changed; the row is found by its Name."""
    table = table.copy()
    mask = table[config.COL_NAME] == name
    assert mask.any(), f"no row named {name!r}"
    table.loc[mask, column] = value
    return table


def raw(example, name, column, value):
    """Change one RawMaterials cell."""
    return replace(example, raw_materials=_edit(example.raw_materials, name, column, value))


def spec(example, name, column, value):
    """Change one Specs cell."""
    return replace(example, specs=_edit(example.specs, name, column, value))


def setting(example, key, value):
    """Change one Settings value."""
    return replace(example, settings={**example.settings, key: value})


def rename_raw_column(example, old, new):
    """Rename one RawMaterials header."""
    return replace(example, raw_materials=example.raw_materials.rename(columns={old: new}))


def usage(example, rows):
    """Replace the UsageLimits sheet."""
    return replace(example, usage_limits=examples.usage_limit_table(rows))


def errors_of(path, **kwargs) -> InputError:
    """Load ``path`` expecting an InputError, and return it."""
    with pytest.raises(InputError) as info:
        load_inputs(path, **kwargs)
    return info.value


def slag():
    """Test 2 (dry basis), the base for most error cases."""
    return examples.synthetic_slag_example()


# ===========================================================================
# Valid workbooks load correctly
# ===========================================================================

def test_hand_worked_example_loads_cleanly(write_example):
    data = load_inputs(write_example(examples.hand_worked_example()))

    assert data.material_names == ["Lime", "Calcined alumina", "Bauxite"]
    assert data.components == ["CaO", "Al2O3", "SiO2", "Others"]
    assert data.raw_materials["Bauxite"].composition == {"CaO": 2, "Al2O3": 60, "SiO2": 8, "Others": 30}
    assert data.raw_materials["Calcined alumina"].dry_cost == 30000
    assert [s.name for s in data.specs] == ["CaO", "Al2O3", "SiO2"]
    assert data.specs[0] == SpecRow("CaO", ("CaO",), (), 50, None, 0.0)
    assert data.specs[2].max_value == 6 and data.specs[2].min_value is None
    assert data.settings.spec_basis == "dry"
    assert data.settings.batch_size_kg == 1000
    assert data.product_limits == {}
    assert data.effective_limit("Lime") == UsageLimit(0, 100)
    assert data.warnings == []


def test_synthetic_slag_prices_converted_to_dry_basis(write_example):
    data = load_inputs(write_example(slag()))
    rm = data.raw_materials

    # as_received: dry cost = cost / (1 - moisture/100)
    assert rm["Lime"].dry_cost == pytest.approx(9000 / 0.995)
    assert rm["Raw bauxite"].dry_cost == pytest.approx(7000 / 0.94)
    # dry: unchanged, even though it has moisture
    assert rm["Calcined bauxite"].dry_cost == 32000
    assert rm["Calcined bauxite"].cost_input == 32000

    assert rm["Lime fines"].plant_limit == UsageLimit(0, 20)
    basicity = data.specs[-1]
    assert basicity.is_ratio and basicity.expression == "CaO / SiO2"
    assert basicity.min_value == 8
    assert data.warnings == []


def test_spec_basis_and_margin_variants(write_example):
    ignited = load_inputs(write_example(examples.synthetic_slag_example(spec_basis="ignited")))
    assert ignited.settings.spec_basis == "ignited"

    margin = load_inputs(write_example(examples.synthetic_slag_example(margin=0.5), "margin.xlsx"))
    cao = margin.specs[0]
    assert (cao.min_target, cao.max_target) == (50.5, 55.5)
    sio2 = margin.specs[2]
    assert sio2.margin == 0 and sio2.max_target == 6


def test_fluorspar_ban_sets_product_limit(write_example):
    data = load_inputs(write_example(examples.synthetic_slag_example(ban_fluorspar=True)))
    assert data.product_limits == {"Fluorspar": UsageLimit(0, 0)}
    assert data.effective_limit("Fluorspar") == UsageLimit(0, 0)
    # The tighter of plant (max 20) and product (none) for an unlisted material
    assert data.effective_limit("Lime fines") == UsageLimit(0, 20)


def test_split_layout_matches_single_workbook(write_example):
    single = load_inputs(write_example(slag(), "single.xlsx"))
    master, product = slag().split()
    master_path = write_example(master, "master.xlsx")
    product_path = write_example(product, "product_only.xlsx")

    split = load_inputs(product_path, raw_materials_file=master_path)

    assert split.raw_materials == single.raw_materials
    assert split.specs == single.specs
    assert split.settings == single.settings
    assert split.sources[config.SHEET_RAW_MATERIALS].endswith("master.xlsx")
    assert split.sources[config.SHEET_SPECS].endswith("product_only.xlsx")
    assert split.warnings == []


def test_master_file_wins_over_product_raw_materials(write_example):
    master_path = write_example(slag().split()[0], "master.xlsx")
    # This product file still has its own (different) RawMaterials sheet
    product_path = write_example(raw(slag(), "Lime", config.COL_COST, 1), "product.xlsx")

    data = load_inputs(product_path, raw_materials_file=master_path)

    assert data.raw_materials["Lime"].cost_input == 9000
    assert any("is ignored because a master raw-material file" in w for w in data.warnings)


def test_component_sum_in_numerator(write_example):
    extra = examples.spec_table([("CaO+MgO", "CaO + MgO", None, None, 60, None)])
    example = replace(slag(), specs=pd.concat([slag().specs, extra], ignore_index=True))
    data = load_inputs(write_example(example))
    assert data.specs[-1].numerator == (Term(1.0, "CaO"), Term(1.0, "MgO"))  # plain sum: factor 1 each


def test_lenient_header_sheet_and_value_matching(write_example, tmp_path):
    example = rename_raw_column(slag(), config.COL_LOI, "loi %")
    example = raw(example, "Lime", config.COL_PRICE_BASIS, "As Received")
    example = replace(example, settings={"product name": "Lenient", "Spec basis": "Ignited", "Batch_size_kg": 500})
    path = write_example(example)
    workbook = openpyxl.load_workbook(path)
    workbook[config.SHEET_RAW_MATERIALS].title = "Raw Materials"
    workbook.save(path)

    data = load_inputs(path)

    assert data.raw_materials["Raw bauxite"].loi_pct == 24
    assert data.raw_materials["Lime"].price_basis == "as_received"
    assert data.settings == Settings(product_name="Lenient", spec_basis="ignited", batch_size_kg=500)
    assert len(data.specs) == 9


# ===========================================================================
# The deliberately broken workbook: every error reported in one go
# ===========================================================================

def test_broken_workbook_reports_every_error_at_once(write_example):
    error = errors_of(write_example(examples.broken_example()))
    text = "\n".join(error.errors)

    expected_fragments = [
        "'Lime': duplicate name",
        "'Calcined dolomite', Cost_Rs_per_t: '8,500 Rs' is not a number",
        "'Calcined bauxite', Price_basis: 'wet' is not allowed",
        "'Raw bauxite', Moisture_%: -6 is negative",
        "Min + margin = 54 is greater than Max - margin = 52",
        "component 'AL2O3' is not a column in RawMaterials. Did you mean 'Al2O3'?",
        "'MgO': give a Min, a Max, or both",
        "'calcined' is not allowed; use one of dry, as_received, ignited",
        "'Flourspar': not a raw material in RawMaterials. Did you mean 'Fluorspar'?",
    ]
    for fragment in expected_fragments:
        assert fragment in text, f"missing error: {fragment}"
    assert len(error.errors) == len(expected_fragments)
    assert str(error).startswith("9 input error(s) found:")

    # The bad Lime analysis is a warning, not an error, and is still reported
    assert len(error.warnings) == 1
    assert "'Lime': components + LOI = 87.34 %" in error.warnings[0]


# ===========================================================================
# One error at a time
# ===========================================================================

ERROR_CASES = [
    # --- RawMaterials -------------------------------------------------------
    pytest.param(lambda ex: rename_raw_column(ex, config.COL_COST, "Cost"),
                 ["required column 'Cost_Rs_per_t' is missing", "Did you mean 'Cost'?"], id="missing-cost-column"),
    pytest.param(lambda ex: rename_raw_column(ex, config.COL_MOISTURE, "Moistre_%"),
                 ["'Moistre_%' looks like a misspelling of 'Moisture_%'", "read as a chemical component"],
                 id="misspelt-fixed-column"),
    pytest.param(lambda ex: raw(ex, "Lime", config.COL_NAME, None), ["Name is blank"], id="blank-name"),
    pytest.param(lambda ex: raw(ex, "Lime fines", config.COL_NAME, "LIME"), ["duplicate name"], id="duplicate-name"),
    pytest.param(lambda ex: raw(ex, "Lime", config.COL_COST, "abc"), ["'abc' is not a number"], id="cost-text"),
    pytest.param(lambda ex: raw(ex, "Lime", config.COL_COST, -5), ["-5 is negative"], id="cost-negative"),
    pytest.param(lambda ex: raw(ex, "Lime", config.COL_COST, None), ["Cost_Rs_per_t: value is blank"], id="cost-blank"),
    pytest.param(lambda ex: raw(ex, "Lime", "SiO2", -1), ["'Lime', SiO2: -1 is negative"], id="component-negative"),
    pytest.param(lambda ex: raw(ex, "Lime", "CaO", 150), ["150 is above the maximum allowed (100)"],
                 id="component-above-100"),
    pytest.param(lambda ex: raw(ex, "Lime", "CaO", "<0.01"), ["'<0.01' is not a number"], id="component-text"),
    pytest.param(lambda ex: raw(ex, "Lime", config.COL_MOISTURE, 100), ["Moisture_%: must be below 100"],
                 id="moisture-100"),
    pytest.param(lambda ex: raw(ex, "Lime", config.COL_LOI, 100), ["LOI_%: must be below 100"], id="loi-100"),
    pytest.param(lambda ex: raw(ex, "Lime fines", config.COL_MIN_USE, 30),
                 ["Min_use_% (30) is greater than Max_use_% (20)"], id="plant-min-above-max"),
    pytest.param(lambda ex: raw(ex, "Lime", config.COL_PRICE_BASIS, None),
                 ["Price_basis: blank, but Moisture_% is 0.5"], id="price-basis-blank-with-moisture"),
    pytest.param(lambda ex: raw(ex, "Lime", config.COL_PRICE_BASIS, "wet"), ["'wet' is not allowed"],
                 id="price-basis-unknown"),
    # --- Specs --------------------------------------------------------------
    pytest.param(lambda ex: spec(ex, "Al2O3", config.COL_NUMERATOR, "AL2O3"), ["Did you mean 'Al2O3'?"],
                 id="component-typo"),
    pytest.param(lambda ex: spec(ex, "Al2O3", config.COL_NUMERATOR, "Xyz"),
                 ["component 'Xyz' is not a column", "Components available: CaO, SiO2"], id="component-unknown"),
    pytest.param(lambda ex: spec(ex, "Basicity", config.COL_DENOMINATOR, "Si02"), ["Did you mean 'SiO2'?"],
                 id="denominator-typo"),
    pytest.param(lambda ex: spec(ex, "CaO", config.COL_NUMERATOR, "CaO++MgO"), ["empty component name"],
                 id="bad-component-sum"),
    pytest.param(lambda ex: spec(ex, "CaO", config.COL_NUMERATOR, "CaO+CaO"), ["lists CaO more than once"],
                 id="repeated-component"),
    pytest.param(lambda ex: spec(ex, "CaO", config.COL_NUMERATOR, None), ["Numerator: value is blank"],
                 id="numerator-blank"),
    pytest.param(lambda ex: spec(ex, "SiO2", config.COL_MAX, None), ["give a Min, a Max, or both"],
                 id="spec-without-limits"),
    pytest.param(lambda ex: spec(ex, "CaO", config.COL_MIN, 57), ["Min (57) is greater than Max (56)"],
                 id="spec-min-above-max"),
    pytest.param(lambda ex: spec(ex, "CaO", config.COL_MARGIN, 4),
                 ["Min + margin = 54 is greater than Max - margin = 52"], id="margin-empties-range"),
    pytest.param(lambda ex: spec(ex, "P", config.COL_MARGIN, 0.05), ["Max - margin = -0.02 is negative"],
                 id="margin-makes-max-negative"),
    pytest.param(lambda ex: spec(ex, "CaO", config.COL_MARGIN, -0.5), ["-0.5 is negative"], id="margin-negative"),
    pytest.param(lambda ex: spec(spec(ex, "CaF2", config.COL_MAX, None), "CaF2", config.COL_MIN, 101),
                 ["above 100 % and can never be met"], id="composition-min-above-100"),
    pytest.param(lambda ex: spec(ex, "Al2O3", config.COL_NAME, "cao"), ["duplicate name"], id="duplicate-spec"),
    pytest.param(lambda ex: replace(ex, specs=None), ["required sheet 'Specs' not found"], id="no-specs-sheet"),
    # --- Settings -----------------------------------------------------------
    pytest.param(lambda ex: replace(ex, settings=None), ["required sheet 'Settings' not found"],
                 id="no-settings-sheet"),
    pytest.param(lambda ex: setting(ex, config.KEY_SPEC_BASIS, "calcined"), ["'calcined' is not allowed"],
                 id="spec-basis-unknown"),
    pytest.param(lambda ex: setting(ex, config.KEY_BATCH_SIZE, 0), ["must be greater than 0"], id="batch-zero"),
    pytest.param(lambda ex: setting(ex, config.KEY_BATCH_SIZE, "1 t"), ["'1 t' is not a number"], id="batch-text"),
    pytest.param(lambda ex: setting(ex, config.KEY_SCALE_RESOLUTION, 0), ["Scale_resolution_kg'", "greater than 0"],
                 id="scale-zero"),
    pytest.param(lambda ex: setting(ex, config.KEY_SCALE_RESOLUTION, "fine"), ["'fine' is not a number"],
                 id="scale-text"),
    # --- UsageLimits --------------------------------------------------------
    pytest.param(lambda ex: usage(ex, [("Flourspar", None, 0)]), ["Did you mean 'Fluorspar'?"],
                 id="usage-unknown-material"),
    pytest.param(lambda ex: usage(ex, [("Lime", 10, 5)]), ["Min_use_% (10) is greater than Max_use_% (5)"],
                 id="usage-min-above-max"),
    pytest.param(lambda ex: usage(ex, [("Lime fines", 30, None)]),
                 ["effective limits conflict", "Effective Min = 30 %", "effective Max = 20 %"],
                 id="usage-conflicts-with-plant-limit"),
    pytest.param(lambda ex: usage(ex, [("Lime", None, 50), ("Lime", None, 40)]), ["listed more than once"],
                 id="usage-duplicate"),
    pytest.param(lambda ex: usage(ex, [("Lime", None, 120)]), ["120 is above the maximum allowed (100)"],
                 id="usage-above-100"),
]


@pytest.mark.parametrize("mutate, fragments", ERROR_CASES)
def test_single_error_is_reported_clearly(write_example, mutate, fragments):
    error = errors_of(write_example(mutate(slag())))
    assert len(error.errors) == 1, error.errors
    for fragment in fragments:
        assert fragment in error.errors[0]


def test_duplicate_component_column(write_example):
    table = slag().raw_materials.copy()
    table.insert(len(table.columns), "CaO", 1.0, allow_duplicates=True)
    error = errors_of(write_example(replace(slag(), raw_materials=table)))
    assert error.errors == ["RawMaterials (product.xlsx): column 'CaO' appears more than once (column Q)."]


def test_duplicate_setting(write_example):
    path = write_example(slag())
    workbook = openpyxl.load_workbook(path)
    workbook[config.SHEET_SETTINGS].append([config.KEY_SPEC_BASIS, "ignited"])
    workbook.save(path)
    error = errors_of(path)
    assert len(error.errors) == 1 and "setting 'Spec_basis' appears more than once" in error.errors[0]


def test_no_component_columns(write_example):
    fixed_only = slag().raw_materials[list(config.RAW_MATERIAL_FIXED_COLUMNS)]
    error = errors_of(write_example(replace(slag(), raw_materials=fixed_only)))
    assert "no chemical component columns found" in error.errors[0]


def test_missing_raw_materials_everywhere(write_example):
    product_only = slag().split()[1]
    error = errors_of(write_example(product_only))
    assert len(error.errors) == 1
    assert "No 'RawMaterials' sheet in product.xlsx, and no master raw-material file" in error.errors[0]


def test_master_file_without_raw_materials_sheet(write_example):
    product_only = slag().split()[1]
    product_path = write_example(product_only, "product.xlsx")
    not_a_master = write_example(product_only, "not_master.xlsx")
    error = errors_of(product_path, raw_materials_file=not_a_master)
    assert "not_master.xlsx: required sheet 'RawMaterials' not found" in error.errors[0]


def test_missing_file(tmp_path):
    error = errors_of(tmp_path / "nope.xlsx")
    assert error.errors[0].startswith("File not found:")


# ===========================================================================
# Warnings (the run continues)
# ===========================================================================

def test_unknown_setting_and_unused_column_warn(write_example):
    example = setting(slag(), "Colour", "white")
    specs = example.specs.copy()
    specs["Comment"] = "note"
    data = load_inputs(write_example(replace(example, specs=specs)))
    assert any("unknown setting 'Colour' is ignored" in w for w in data.warnings)
    assert any("column 'Comment' is not used and is ignored" in w for w in data.warnings)


def test_scale_resolution_setting(write_example):
    assert load_inputs(write_example(slag(), "default.xlsx")).settings.scale_resolution_kg == \
           config.SCALE_RESOLUTION_KG  # key missing
    blank = load_inputs(write_example(setting(slag(), config.KEY_SCALE_RESOLUTION, None), "blank.xlsx"))
    assert blank.settings.scale_resolution_kg == config.SCALE_RESOLUTION_KG  # key present, value blank
    assert blank.warnings == []
    given = load_inputs(write_example(setting(slag(), config.KEY_SCALE_RESOLUTION, 0.5), "given.xlsx"))
    assert given.settings.scale_resolution_kg == 0.5


def test_blank_product_name_uses_file_name(write_example):
    data = load_inputs(write_example(setting(slag(), config.KEY_PRODUCT_NAME, None), "my_slag.xlsx"))
    assert data.settings.product_name == "my_slag"
    assert any("Product_name not given; using 'my_slag'" in w for w in data.warnings)


def test_as_received_price_without_moisture_warns(write_example):
    data = load_inputs(write_example(raw(slag(), "Lime", config.COL_MOISTURE, None)))
    assert data.raw_materials["Lime"].dry_cost == 9000
    assert any("'Lime': price is 'as_received' but Moisture_% is blank" in w for w in data.warnings)


def test_blank_price_basis_without_moisture_is_dry(write_example):
    # With no moisture, as_received and dry give the same price, so blank is fine
    example = raw(examples.hand_worked_example(), "Lime", config.COL_PRICE_BASIS, None)
    data = load_inputs(write_example(example))
    assert data.raw_materials["Lime"].price_basis == "dry"
    assert data.warnings == []


# ===========================================================================
# Lab analysis convention
# ===========================================================================

def _loi_free_version(example):
    """The same raw materials, but with analyses reported on an LOI-free basis."""
    table = example.raw_materials.copy()
    for component in examples.SYNTHETIC_SLAG_COMPONENTS:
        table[component] = [
            value / (1 - loi / 100) for value, loi in zip(table[component], table[config.COL_LOI])
        ]
    return replace(example, raw_materials=table)


def test_loi_free_convention_converts_to_dry_basis(write_example):
    dry_data = load_inputs(write_example(slag(), "dry.xlsx"))
    path = write_example(_loi_free_version(slag()), "loi_free.xlsx")

    data = load_inputs(path, lab_convention=config.LAB_LOI_FREE)

    assert data.lab_convention == config.LAB_LOI_FREE
    assert data.warnings == []  # LOI-free numbers total ~100 on their own
    for name, material in data.raw_materials.items():
        for component, value in material.composition.items():
            assert value == pytest.approx(dry_data.raw_materials[name].composition[component])


def test_total_warning_suggests_loi_free(write_example):
    # LOI-free numbers read with the default convention: components + LOI is far above 103
    data = load_inputs(write_example(_loi_free_version(slag())))
    bauxite = [w for w in data.warnings if "'Raw bauxite'" in w]
    assert len(bauxite) == 1
    assert "components + LOI = 123.49 %" in bauxite[0]
    assert "Components alone = 99.49 %, which looks like the 'loi_free' convention" in bauxite[0]


def test_total_warning_suggests_dry_incl_loi(write_example):
    # Dry-with-LOI numbers read as loi_free: components alone are too low
    data = load_inputs(write_example(slag()), lab_convention=config.LAB_LOI_FREE)
    bauxite = [w for w in data.warnings if "'Raw bauxite'" in w]
    assert len(bauxite) == 1
    assert "components alone = 75.61 %" in bauxite[0]
    assert "looks like the 'dry_incl_loi' convention" in bauxite[0]


def test_analysis_total_warning_cases():
    assert analysis_total_warning("A", 96.0, 3.0, config.LAB_DRY_INCL_LOI) is None
    assert analysis_total_warning("A", 99.0, 0.0, config.LAB_LOI_FREE) is None
    neither = analysis_total_warning("A", 80.0, 5.0, config.LAB_DRY_INCL_LOI)
    assert "Neither convention fits" in neither


def test_analysis_to_dry_basis():
    reported = {"CaO": 50.0, "SiO2": 50.0}
    assert analysis_to_dry_basis(reported, 20.0, config.LAB_DRY_INCL_LOI) == reported
    assert analysis_to_dry_basis(reported, 20.0, config.LAB_LOI_FREE) == {"CaO": 40.0, "SiO2": 40.0}


def test_unknown_lab_convention_rejected(write_example):
    with pytest.raises(ValueError):
        load_inputs(write_example(slag()), lab_convention="wet")


# ===========================================================================
# Small helpers
# ===========================================================================

def test_parse_expression():
    assert parse_expression("CaO") == [Term(1.0, "CaO")]
    assert parse_expression(" CaO + MgO ") == [Term(1.0, "CaO"), Term(1.0, "MgO")]
    for bad in ("", "CaO+", "+CaO", "CaO++MgO", "CaO+CaO"):
        with pytest.raises(ValueError):
            parse_expression(bad)


def test_suggest_close_match():
    known = ["CaO", "Al2O3", "SiO2", "Fe2O3"]
    assert suggest_close_match("AL2O3", known) == "Al2O3"
    assert suggest_close_match("Si02", known) == "SiO2"
    assert suggest_close_match("Unobtainium", known) is None


def test_price_conversions_round_trip():
    assert price_to_dry_basis(9000, "as_received", 10) == pytest.approx(10000)
    assert price_to_dry_basis(9000, "dry", 10) == 9000
    assert price_from_dry_basis(10000, "as_received", 10) == pytest.approx(9000)
    with pytest.raises(ValueError):
        price_to_dry_basis(9000, "wet", 10)


def test_usage_limit_tighter_and_spec_targets():
    assert UsageLimit(5, 80).tighter(UsageLimit(0, 50)) == UsageLimit(5, 50)
    row = SpecRow("CaO", ("CaO",), (), 50, 56, 0.5)
    assert (row.min_target, row.max_target, row.is_ratio) == (50.5, 55.5, False)
    assert SpecRow("B", ("CaO", "MgO"), ("SiO2",), 2, None).expression == "(CaO+MgO) / SiO2"
    assert SpecRow("C", ("CaO", "MgO"), (), None, 60).expression == "CaO+MgO"

"""Example datasets, used by ``make_templates.py`` and by the tests.

* Test 1: the hand-worked three-material example (expected results known exactly).
* Test 2: a synthetic slag with typical (not real) raw material data.
* A deliberately broken workbook, to show and test the validation messages.

The tables are written row by row in the same column order as the Excel
sheets, so they can be checked against the original data by eye.
"""

from dataclasses import dataclass, replace
from pathlib import Path
from typing import Sequence

import pandas as pd

from blend_optimizer import config
from blend_optimizer.templates import write_workbook


@dataclass(frozen=True, eq=False)
class ExampleWorkbook:
    """The sheets of one example workbook (None = sheet left out)."""

    raw_materials: pd.DataFrame | None = None
    specs: pd.DataFrame | None = None
    settings: dict[str, object] | None = None
    usage_limits: pd.DataFrame | None = None
    what_if: pd.DataFrame | None = None

    def write(self, path: str | Path) -> Path:
        """Write this example to an .xlsx file and return the path."""
        return write_workbook(path, self.raw_materials, self.specs, self.settings, self.usage_limits,
                              what_if=self.what_if)

    def split(self) -> tuple["ExampleWorkbook", "ExampleWorkbook"]:
        """Split into (master raw-material workbook, product-only workbook)."""
        master = ExampleWorkbook(raw_materials=self.raw_materials)
        product = replace(self, raw_materials=None)
        return master, product


# ---------------------------------------------------------------------------
# Table builders
# ---------------------------------------------------------------------------

def raw_material_table(components: Sequence[str], rows: Sequence[Sequence[object]]) -> pd.DataFrame:
    """RawMaterials table. Each row: Name, Cost, Price_basis, Moisture, LOI, Min_use, Max_use, *analysis."""
    columns = list(config.RAW_MATERIAL_FIXED_COLUMNS) + list(components)
    # dtype=object so tests can put text (e.g. "abc") into number columns
    return pd.DataFrame([list(row) for row in rows], columns=columns, dtype=object)


def spec_table(rows: Sequence[Sequence[object]]) -> pd.DataFrame:
    """Specs table. Each row: Name, Numerator, Denominator, Min, Max, Margin."""
    return pd.DataFrame([list(row) for row in rows], columns=list(config.SPEC_COLUMNS), dtype=object)


def usage_limit_table(rows: Sequence[Sequence[object]]) -> pd.DataFrame:
    """UsageLimits table. Each row: Name, Min_use_%, Max_use_%."""
    return pd.DataFrame([list(row) for row in rows], columns=list(config.USAGE_LIMIT_COLUMNS), dtype=object)


def what_if_table(rows: Sequence[Sequence[object]]) -> pd.DataFrame:
    """WhatIf table. Each row: Scenario, Change type, Item, New value."""
    return pd.DataFrame([list(row) for row in rows], columns=list(config.WHAT_IF_COLUMNS), dtype=object)


def settings_dict(product_name: str, spec_basis: str = config.DEFAULT_SPEC_BASIS,
                  batch_size_kg: float = config.DEFAULT_BATCH_SIZE_KG) -> dict[str, object]:
    """Settings sheet contents."""
    return {
        config.KEY_PRODUCT_NAME: product_name,
        config.KEY_SPEC_BASIS: spec_basis,
        config.KEY_BATCH_SIZE: batch_size_kg,
    }


# ---------------------------------------------------------------------------
# Test 1: hand-worked example
# ---------------------------------------------------------------------------
# Dry basis, no moisture, no LOI, prices on dry basis, no usage limits.
# "Others" makes each analysis total 100 and is not constrained.

HAND_WORKED_COMPONENTS = ("CaO", "Al2O3", "SiO2", "Others")

_HAND_WORKED_RAW_MATERIALS = [
    # Name,              Cost,  Basis, Moist, LOI, Min,  Max,   CaO, Al2O3, SiO2, Others
    ("Lime",              9000, "dry",  0,     0,   None, None,  92,   1,    2,    5),
    ("Calcined alumina", 30000, "dry",  0,     0,   None, None,   1,  85,    6,    8),
    ("Bauxite",          12000, "dry",  0,     0,   None, None,   2,  60,    8,   30),
]


def hand_worked_example(al2o3_min: float = 30.0) -> ExampleWorkbook:
    """Test 1. With ``al2o3_min=45`` it becomes the infeasibility test."""
    specs = spec_table([
        # Name,   Numerator, Denominator, Min,       Max,  Margin
        ("CaO",   "CaO",     None,        50,        None, None),
        ("Al2O3", "Al2O3",   None,        al2o3_min, None, None),
        ("SiO2",  "SiO2",    None,        None,      6,    None),
    ])
    return ExampleWorkbook(
        raw_materials=raw_material_table(HAND_WORKED_COMPONENTS, _HAND_WORKED_RAW_MATERIALS),
        specs=specs,
        settings=settings_dict("Test 1 hand-worked example"),
    )


# ---------------------------------------------------------------------------
# Test 2: synthetic slag (typical, not real, data)
# ---------------------------------------------------------------------------

SYNTHETIC_SLAG_COMPONENTS = ("CaO", "SiO2", "Al2O3", "MgO", "CaF2", "Fe2O3", "TiO2", "S", "P")

_SYNTHETIC_SLAG_RAW_MATERIALS = [
    # Name,               Cost,  Price_basis,   Moist, LOI,  Min,  Max,  CaO,  SiO2, Al2O3, MgO,  CaF2, Fe2O3, TiO2, S,    P
    ("Lime",               9000, "as_received", 0.5,   3.5,  None, None, 92.0, 1.5,  0.8,   1.2,  0,    0.3,   0,    0.03, 0.01),
    ("Lime fines",         6500, "as_received", 1.0,   6.0,  None, 20,   88.0, 2.5,  1.2,   1.5,  0,    0.6,   0,    0.05, 0.02),
    ("Calcined dolomite",  8500, "as_received", 0.5,   2.5,  None, None, 57.0, 1.5,  0.8,   37.0, 0,    0.6,   0,    0.02, 0.01),
    ("Calcined bauxite",  32000, "dry",         0.5,   1.0,  None, None, 1.0,  6.0,  85.0,  0.5,  0,    2.5,   3.5,  0.05, 0.05),
    ("Raw bauxite",        7000, "as_received", 6.0,   24.0, None, None, 1.5,  9.0,  58.0,  0.5,  0,    4.0,   2.5,  0.05, 0.06),
    ("Fluorspar",         28000, "as_received", 1.0,   1.5,  None, None, 1.5,  9.0,  1.0,   0.5,  85.0, 1.0,   0,    0.10, 0.02),
]


def _synthetic_slag_specs(margin: float = 0.0) -> pd.DataFrame:
    """Test 2 specs; ``margin`` is applied to CaO and Al2O3 only."""
    margin_cell = margin or None  # leave the cell blank when there is no margin
    return spec_table([
        # Name,      Numerator, Denominator, Min,  Max,  Margin
        ("CaO",      "CaO",     None,        50,   56,   margin_cell),
        ("Al2O3",    "Al2O3",   None,        30,   38,   margin_cell),
        ("SiO2",     "SiO2",    None,        None, 6,    None),
        ("MgO",      "MgO",     None,        None, 8,    None),
        ("CaF2",     "CaF2",    None,        2,    5,    None),
        ("Fe2O3",    "Fe2O3",   None,        None, 2,    None),
        ("S",        "S",       None,        None, 0.08, None),
        ("P",        "P",       None,        None, 0.03, None),
        ("Basicity", "CaO",     "SiO2",      8,    None, None),
    ])


def synthetic_slag_example(
    spec_basis: str = config.BASIS_DRY, margin: float = 0.0, ban_fluorspar: bool = False,
    with_what_if: bool = False,
) -> ExampleWorkbook:
    """Test 2 and its variants.

    Args:
        spec_basis: "dry", "ignited" or "as_received".
        margin: Safety margin on CaO and Al2O3 (the margin variant uses 0.5).
        ban_fluorspar: Add a UsageLimits sheet with Fluorspar Max 0
            (a customer ban; makes CaF2 Min 2 impossible).
        with_what_if: Add a WhatIf sheet with the example scenarios.
    """
    usage_limits = usage_limit_table([("Fluorspar", None, 0)]) if ban_fluorspar else None
    return ExampleWorkbook(
        raw_materials=raw_material_table(SYNTHETIC_SLAG_COMPONENTS, _SYNTHETIC_SLAG_RAW_MATERIALS),
        specs=_synthetic_slag_specs(margin),
        settings=settings_dict(f"Synthetic slag (Test 2, {spec_basis} basis)", spec_basis),
        usage_limits=usage_limits,
        what_if=synthetic_slag_what_if() if with_what_if else None,
    )


def synthetic_slag_what_if() -> pd.DataFrame:
    """Example WhatIf sheet for Test 2: four scenarios, one of them infeasible."""
    return what_if_table([
        # Scenario,              Change type,   Item,          New value
        ("Fluorspar dearer",     "Price",       "Fluorspar",   35000),
        ("Looser P",             "Spec Max",    "P",           0.035),
        ("Capped bauxite",       "Usage Max",   "Raw bauxite", 20),
        ("Capped bauxite",       "Spec Margin", "CaO",         0.5),
        ("No fluorspar",         "Usage Max",   "Fluorspar",   0),     # CaF2 Min 2 becomes impossible
    ])


# ---------------------------------------------------------------------------
# Deliberately broken workbook
# ---------------------------------------------------------------------------

def broken_example() -> ExampleWorkbook:
    """Test 2 with one mistake of each common kind, to test and show the messages.

    Errors (9): duplicate raw material name, cost typed as text, negative
    moisture, unknown price basis, component typo in Specs ("AL2O3"), margin
    that empties the CaO range, a spec with neither Min nor Max, unknown
    spec basis, and a usage limit for a misspelt material ("Flourspar").

    Warning (1): Lime's analysis no longer adds up (CaO typed as 80).
    """
    rows = [list(row) for row in _SYNTHETIC_SLAG_RAW_MATERIALS]
    rows[0][7] = 80.0          # Lime CaO 92 -> 80: total about 87 % (warning only)
    rows[1][0] = "Lime"        # "Lime fines" renamed to a duplicate name
    rows[2][1] = "8,500 Rs"    # Calcined dolomite cost is text
    rows[4][3] = -6.0          # Raw bauxite negative moisture
    rows[3][2] = "wet"         # Calcined bauxite unknown price basis

    specs = spec_table([
        # Name,      Numerator, Denominator, Min,  Max,  Margin
        ("CaO",      "CaO",     None,        50,   56,   4),     # 54 > 52 after margin
        ("Al2O3",    "AL2O3",   None,        30,   38,   None),  # typo
        ("SiO2",     "SiO2",    None,        None, 6,    None),
        ("MgO",      "MgO",     None,        None, None, None),  # no limit at all
        ("Basicity", "CaO",     "SiO2",      8,    None, None),
    ])
    return ExampleWorkbook(
        raw_materials=raw_material_table(SYNTHETIC_SLAG_COMPONENTS, rows),
        specs=specs,
        settings=settings_dict("Broken example", spec_basis="calcined"),
        usage_limits=usage_limit_table([("Flourspar", None, 0)]),
    )

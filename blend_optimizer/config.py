"""Central constants: conventions, defaults, tolerances, and the sheet/column names.

If your workbooks, lab reports or tolerances change, this is the file to edit.
Nothing here imports pandas or PuLP.
"""

# ---------------------------------------------------------------------------
# Lab analysis convention
# ---------------------------------------------------------------------------
# How the lab reports the chemical analysis of a raw material:
#
#   "dry_incl_loi"  wt% on a dry basis, LOI is one of the reported items,
#                   so components + LOI should total about 100.
#                   The numbers are used as they are.
#
#   "loi_free"      wt% on an ignited (LOI-free) basis, so the components
#                   alone total about 100. The loader converts them to a dry
#                   basis: 1 t dry material leaves (1 - LOI/100) t after
#                   ignition, so  c_dry = c_reported * (1 - LOI/100).
LAB_DRY_INCL_LOI = "dry_incl_loi"
LAB_LOI_FREE = "loi_free"
LAB_CONVENTIONS = (LAB_DRY_INCL_LOI, LAB_LOI_FREE)
LAB_ANALYSIS_CONVENTION = LAB_DRY_INCL_LOI

# A raw material whose analysis total (per the convention above) falls outside
# this range gets a warning. It is a warning, not an error: the run continues.
ANALYSIS_TOTAL_WARN_RANGE = (97.0, 103.0)

# ---------------------------------------------------------------------------
# Bases
# ---------------------------------------------------------------------------
# Price basis of Cost_Rs_per_t in RawMaterials
PRICE_AS_RECEIVED = "as_received"
PRICE_DRY = "dry"
PRICE_BASES = (PRICE_AS_RECEIVED, PRICE_DRY)

# Basis on which the customer's composition limits are expressed
BASIS_DRY = "dry"
BASIS_AS_RECEIVED = "as_received"
BASIS_IGNITED = "ignited"
SPEC_BASES = (BASIS_DRY, BASIS_AS_RECEIVED, BASIS_IGNITED)

# ---------------------------------------------------------------------------
# Defaults and tolerances
# ---------------------------------------------------------------------------
DEFAULT_SPEC_BASIS = BASIS_DRY
# Batch_size_kg is the total AS-RECEIVED mass weighed into the mixer
DEFAULT_BATCH_SIZE_KG = 1000.0
# A constraint whose slack is below this is reported as "Binding"
BINDING_TOLERANCE = 1e-6
# Elastic model (diagnosis of infeasible problems): violations are divided by
# the size of their limit, but never by less than this (wt % or ratio units),
# so a limit of 0 does not divide by zero.
ELASTIC_MIN_LIMIT = 0.01

# ---------------------------------------------------------------------------
# Report display
# ---------------------------------------------------------------------------
# All calculations use full precision. Numbers are rounded only when shown
# (console and Excel); the tables returned by the code are not rounded.
DISPLAY_DECIMALS_PCT = 2    # wt %, ratios, limits (minimum; spec rows may show more, see below)
DISPLAY_DECIMALS_KG = 1     # batch masses
DISPLAY_DECIMALS_RS = 0     # rupees: whole numbers
# Spec rows: limits, targets and achieved values are shown with one more
# decimal than the limits as typed (Min, Max, Margin), and at least
# DISPLAY_DECIMALS_PCT. E.g. P Max 0.03 -> 0.030. Capped at this:
DISPLAY_DECIMALS_PCT_MAX = 6
# Weighing-scale resolution: kg are shown rounded to a multiple of this,
# e.g. 0.5 shows 12.26 kg as 12.5 kg. A product can override it with the
# optional Settings key Scale_resolution_kg. On the weigh sheet the rounding
# difference is put on the largest material, so the rows add up to the total.
SCALE_RESOLUTION_KG = 0.1

# ---------------------------------------------------------------------------
# Workbook layout: sheet names
# ---------------------------------------------------------------------------
SHEET_RAW_MATERIALS = "RawMaterials"
SHEET_SPECS = "Specs"
SHEET_SETTINGS = "Settings"
SHEET_USAGE_LIMITS = "UsageLimits"
SHEET_WHAT_IF = "WhatIf"
SHEET_README = "README"

# ---------------------------------------------------------------------------
# Workbook layout: column names
# ---------------------------------------------------------------------------
COL_NAME = "Name"

# RawMaterials. Every column that is NOT one of these fixed columns is read as
# a chemical component (wt%, dry basis).
COL_COST = "Cost_Rs_per_t"
COL_PRICE_BASIS = "Price_basis"
COL_MOISTURE = "Moisture_%"
COL_LOI = "LOI_%"
COL_MIN_USE = "Min_use_%"
COL_MAX_USE = "Max_use_%"
RAW_MATERIAL_FIXED_COLUMNS = (
    COL_NAME,
    COL_COST,
    COL_PRICE_BASIS,
    COL_MOISTURE,
    COL_LOI,
    COL_MIN_USE,
    COL_MAX_USE,
)
RAW_MATERIAL_REQUIRED_COLUMNS = (COL_NAME, COL_COST)

# Specs
COL_NUMERATOR = "Numerator"
COL_DENOMINATOR = "Denominator"
COL_MIN = "Min"
COL_MAX = "Max"
COL_MARGIN = "Margin"
SPEC_COLUMNS = (COL_NAME, COL_NUMERATOR, COL_DENOMINATOR, COL_MIN, COL_MAX, COL_MARGIN)
SPEC_REQUIRED_COLUMNS = (COL_NAME, COL_NUMERATOR, COL_MIN, COL_MAX)
# Numerator/Denominator expressions: terms joined with "+"; a term is a
# component, or a factor and a component joined with "*", e.g. "Al2O3+1.8895*Al"
COMPONENT_SEPARATOR = "+"
FACTOR_SEPARATOR = "*"

# UsageLimits (optional sheet): same usage columns as RawMaterials
USAGE_LIMIT_COLUMNS = (COL_NAME, COL_MIN_USE, COL_MAX_USE)

# Settings: two columns, one row per key
COL_KEY = "Key"
COL_VALUE = "Value"
SETTINGS_COLUMNS = (COL_KEY, COL_VALUE)
KEY_PRODUCT_NAME = "Product_name"
KEY_SPEC_BASIS = "Spec_basis"
KEY_BATCH_SIZE = "Batch_size_kg"
KEY_SCALE_RESOLUTION = "Scale_resolution_kg"  # optional; blank = SCALE_RESOLUTION_KG
SETTINGS_KEYS = (KEY_PRODUCT_NAME, KEY_SPEC_BASIS, KEY_BATCH_SIZE, KEY_SCALE_RESOLUTION)

# WhatIf (optional sheet): one row per change; rows with the same Scenario
# name form one scenario, compared with the base case.
COL_SCENARIO = "Scenario"
COL_CHANGE_TYPE = "Change type"
COL_ITEM = "Item"
COL_NEW_VALUE = "New value"
WHAT_IF_COLUMNS = (COL_SCENARIO, COL_CHANGE_TYPE, COL_ITEM, COL_NEW_VALUE)
CHANGE_PRICE = "Price"              # Item: raw material; value: Rs/t in its price basis
CHANGE_SPEC_MIN = "Spec Min"        # Item: spec name; value: new Min (blank removes it)
CHANGE_SPEC_MAX = "Spec Max"        # Item: spec name; value: new Max (blank removes it)
CHANGE_SPEC_MARGIN = "Spec Margin"  # Item: spec name; value: new margin (blank = 0)
CHANGE_USAGE_MIN = "Usage Min"      # Item: raw material; value: product Min_use_% (blank = 0)
CHANGE_USAGE_MAX = "Usage Max"      # Item: raw material; value: product Max_use_% (blank = 100)
CHANGE_TYPES = (CHANGE_PRICE, CHANGE_SPEC_MIN, CHANGE_SPEC_MAX, CHANGE_SPEC_MARGIN,
                CHANGE_USAGE_MIN, CHANGE_USAGE_MAX)

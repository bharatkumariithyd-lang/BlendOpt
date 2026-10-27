"""Read and validate the input workbook(s) into a :class:`BlendData` object.

Two layouts are supported by the same function, :func:`load_inputs`:

1. One workbook per product with the sheets RawMaterials, Specs, Settings
   and (optional) UsageLimits.
2. A master raw-material workbook (RawMaterials only) plus a product workbook
   (Specs, Settings, optional UsageLimits).

Validation has two levels:

* **errors** stop the run. They are all collected first and raised together
  in one :class:`InputError`, so every problem can be fixed in one go.
* **warnings** let the run continue. They end up in ``BlendData.warnings``.

Header matching is forgiving about case, spaces, underscores and "%"
("loi %" is read as "LOI_%"). A header that is *nearly* a fixed column name
(a typo such as "Moistre_%") is an error: in RawMaterials it would otherwise
be read as a chemical component. Component names themselves must match
exactly, because "AL2O3" and "Al2O3" could be different things to the
optimizer; close matches are suggested in the error message.

Numerator/Denominator expressions (e.g. "Al2O3+1.8895*Al") are read by
:func:`parse_expression` only.

This module knows nothing about PuLP.
"""

import difflib
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl.utils import get_column_letter

from blend_optimizer import config
from blend_optimizer.rounding import format_fixed
from blend_optimizer.data import (
    BlendData,
    InputError,
    RawMaterial,
    Scenario,
    Settings,
    SpecRow,
    Term,
    UsageLimit,
)


# ===========================================================================
# Public helpers (also used by tests and other modules)
# ===========================================================================

# A factor: digits with an optional decimal point, e.g. 2, 1.8895, .5 (no sign, no exponent)
_NUMBER = re.compile(r"^(\d+(\.\d*)?|\.\d+)$")
# A term that starts with a number glued to a name, e.g. "1.8895Al" -> ("1.8895", "Al")
_NUMBER_THEN_NAME = re.compile(r"^(\d+(?:\.\d*)?|\.\d+)\s*(\S.*)$")


def parse_expression(text: str, known_components: Iterable[str] | None = None) -> list[Term]:
    """Read a Numerator/Denominator expression into (factor, component) terms.

    This is the ONE place where an expression is read; everything else uses
    the terms it returns. The syntax:

    * terms joined with "+", e.g. ``CaO+MgO``;
    * a term is a component (factor 1) or ``number*Component``, e.g.
      ``Al2O3+1.8895*Al``; spaces around "+" and "*" are allowed;
    * factors are plain positive numbers typed by the user. The tool knows
      no chemistry.

    Not allowed (each gives a clear message): an empty term ("++"), a factor
    without "*" ("1.8895Al"), a missing factor or component ("*Al",
    "1.8895*"), the factor after the component ("Al*1.8895"), a factor that
    is not a number, a factor of 0, subtraction ("-" between terms, which
    could make a ratio's denominator zero or negative), and the same
    component twice.

    A hyphen inside a component name ("Fe-total") is fine.

    (A future shorthand such as "Al as Al2O3" would be turned into
    "1.8895*Al" here, at the start, before anything else sees the text.)

    Args:
        text: The cell text.
        known_components: The component columns, if known. A term without
            "*" that exactly matches one of them is that component, even if it
            starts with a digit (a column "3CaO"). They also let "CaO-MgO" be
            recognised as an attempted subtraction rather than an unknown name.

    Returns:
        The terms, in the order written.

    Raises:
        ValueError: If the expression is not valid; the message names the
            expression and the term at fault.
    """
    text = str(text).strip()
    if not text:
        raise ValueError("the expression is empty")
    known = set(known_components) if known_components is not None else None
    subtraction = (f"'{text}': subtraction ('-') is not supported; write every term with a positive "
                   f"factor joined by '{config.COMPONENT_SEPARATOR}'")

    terms: list[Term] = []
    for raw_term in text.split(config.COMPONENT_SEPARATOR):
        term = raw_term.strip()
        if not term:
            raise ValueError(f"'{text}' has an empty component name: check for "
                             f"'{config.COMPONENT_SEPARATOR}{config.COMPONENT_SEPARATOR}', or a "
                             f"'{config.COMPONENT_SEPARATOR}' at the start or end")
        if known is not None and term in known:
            terms.append(Term(1.0, term, term))   # an exact column name always wins
            continue
        # "-" as an operator: at the start, with a space next to it, or joining known names
        if term.startswith("-") or re.search(r"\s-|-\s", term):
            raise ValueError(subtraction)
        if known is not None and "-" in term and all(part.strip() in known for part in term.split("-")):
            raise ValueError(subtraction)

        if config.FACTOR_SEPARATOR not in term:
            glued = _NUMBER_THEN_NAME.match(term)
            if glued:
                number, name = glued.groups()
                raise ValueError(f"in '{text}', the term '{term}' starts with a number; put "
                                 f"'{config.FACTOR_SEPARATOR}' between the factor and the component "
                                 f"(did you mean {number}{config.FACTOR_SEPARATOR}{name}?)")
            terms.append(Term(1.0, term, term))
            continue

        parts = [part.strip() for part in term.split(config.FACTOR_SEPARATOR)]
        if len(parts) != 2:
            raise ValueError(f"in '{text}', the term '{term}' has more than one '{config.FACTOR_SEPARATOR}'; "
                             f"write one factor per component, e.g. 1.8895{config.FACTOR_SEPARATOR}Al")
        factor_text, component = parts
        if not factor_text:
            raise ValueError(f"in '{text}', the term '{term}' has no factor before "
                             f"'{config.FACTOR_SEPARATOR}', e.g. 1.8895{config.FACTOR_SEPARATOR}Al")
        if not component:
            raise ValueError(f"in '{text}', the term '{term}' has no component after '{config.FACTOR_SEPARATOR}'")
        if "-" in factor_text:
            raise ValueError(subtraction)
        if not _NUMBER.match(factor_text):
            if _NUMBER.match(component):
                raise ValueError(f"in '{text}', the term '{term}' has the factor after the component; "
                                 f"write it first: {component}{config.FACTOR_SEPARATOR}{factor_text}")
            hint = " (use a decimal point, e.g. 1.8895)" if "," in factor_text else ""
            raise ValueError(f"in '{text}', the factor '{factor_text}' is not a number{hint}")
        factor = float(factor_text)
        if factor == 0:
            raise ValueError(f"in '{text}', the factor of {component} is 0; it must be greater than 0")
        terms.append(Term(factor, component, f"{factor_text}{config.FACTOR_SEPARATOR}{component}"))

    names = [term.component for term in terms]
    duplicates = sorted({name for name in names if names.count(name) > 1})
    if duplicates:
        raise ValueError(f"'{text}' lists {', '.join(duplicates)} more than once; "
                         f"write each component once, with the combined factor")
    return terms


def suggest_close_match(name: str, known: Iterable[str]) -> str | None:
    """Return the known name that ``name`` is most likely a typo of, or None.

    A case-insensitive exact match wins (``"AL2O3"`` -> ``"Al2O3"``);
    otherwise the closest spelling above a similarity cut-off is returned.
    """
    known = list(known)
    for candidate in known:
        if candidate.lower() == name.lower():
            return candidate
    matches = difflib.get_close_matches(name, known, n=1, cutoff=0.75)
    return matches[0] if matches else None


def analysis_to_dry_basis(
    composition: dict[str, float], loi_pct: float, convention: str
) -> dict[str, float]:
    """Convert a reported analysis to wt % on a dry basis.

    * ``"dry_incl_loi"``: already dry basis, returned unchanged.
    * ``"loi_free"``: reported per tonne of *ignited* material. One dry tonne
      leaves (1 - LOI/100) t after ignition, so each component per dry tonne
      is ``c_reported * (1 - LOI/100)``.

    Args:
        composition: Component -> wt % as reported by the lab.
        loi_pct: Loss on ignition, wt % dry basis.
        convention: One of ``config.LAB_CONVENTIONS``.

    Returns:
        Component -> wt %, dry basis.
    """
    if convention == config.LAB_DRY_INCL_LOI:
        return dict(composition)
    if convention == config.LAB_LOI_FREE:
        factor = 1.0 - loi_pct / 100.0
        return {component: value * factor for component, value in composition.items()}
    raise ValueError(f"Unknown lab convention {convention!r}; expected one of {config.LAB_CONVENTIONS}")


def analysis_total_warning(
    material_name: str, components_total: float, loi_pct: float, convention: str
) -> str | None:
    """Check that a reported analysis adds up, and say which convention it resembles.

    Under ``"dry_incl_loi"`` the total checked is components + LOI; under
    ``"loi_free"`` it is the components alone. If the total is outside
    ``config.ANALYSIS_TOTAL_WARN_RANGE``, the other convention's total is
    tried too, so the warning can say "these numbers look like loi_free".

    Args:
        material_name: Used in the message.
        components_total: Sum of the reported component values.
        loi_pct: Reported LOI.
        convention: The convention in use.

    Returns:
        A warning message, or None if the total is in range.
    """
    low, high = config.ANALYSIS_TOTAL_WARN_RANGE
    with_loi = components_total + loi_pct
    if convention == config.LAB_DRY_INCL_LOI:
        total, total_label = with_loi, "components + LOI"
        other_total, other_label, other_convention = components_total, "components alone", config.LAB_LOI_FREE
    else:
        total, total_label = components_total, "components alone"
        other_total, other_label, other_convention = with_loi, "components + LOI", config.LAB_DRY_INCL_LOI

    if low <= total <= high:
        return None
    message = (
        f"RawMaterials '{material_name}': {total_label} = {format_fixed(total, 2)} %, outside the "
        f"expected {low:g}-{high:g} % for the '{convention}' lab convention."
    )
    if loi_pct > 0 and low <= other_total <= high:
        message += (
            f" {other_label.capitalize()} = {format_fixed(other_total, 2)} %, which looks like the "
            f"'{other_convention}' convention. Check the lab report and "
            f"LAB_ANALYSIS_CONVENTION in config.py."
        )
    else:
        message += (
            f" Neither convention fits (components alone = {format_fixed(components_total, 2)} %, "
            f"components + LOI = {format_fixed(with_loi, 2)} %): check for a missing component "
            f"column, a typing error or a missing LOI."
        )
    return message


# ===========================================================================
# Main entry point
# ===========================================================================

def load_inputs(
    product_file: str | Path,
    raw_materials_file: str | Path | None = None,
    lab_convention: str = config.LAB_ANALYSIS_CONVENTION,
) -> BlendData:
    """Read and validate the input workbook(s).

    Args:
        product_file: Workbook with Specs, Settings and optional UsageLimits.
            In the one-workbook layout it also holds RawMaterials.
        raw_materials_file: Optional master workbook with the RawMaterials
            sheet. If given, RawMaterials is taken from here (a RawMaterials
            sheet in the product file is then ignored, with a warning).
        lab_convention: How component analyses are reported; see
            ``config.LAB_ANALYSIS_CONVENTION``.

    Returns:
        The validated data. Non-fatal findings are in ``.warnings``.

    Raises:
        InputError: If anything is wrong. ``.errors`` lists every problem
            found and ``.warnings`` the warnings found so far.
    """
    if lab_convention not in config.LAB_CONVENTIONS:
        raise ValueError(f"lab_convention must be one of {config.LAB_CONVENTIONS}, got {lab_convention!r}")

    issues = _Issues()
    product_path = Path(product_file)
    product_sheets = _read_workbook(product_path)

    # --- Where does RawMaterials come from? ---------------------------------
    if raw_materials_file is not None:
        rm_path = Path(raw_materials_file)
        rm_sheets = _read_workbook(rm_path)
        rm_raw = _find_sheet(rm_sheets, config.SHEET_RAW_MATERIALS, rm_path, issues, required=True)
        if _find_sheet(product_sheets, config.SHEET_RAW_MATERIALS, product_path, issues, required=False) is not None:
            issues.warn(
                f"{product_path.name} also has a {config.SHEET_RAW_MATERIALS} sheet. It is ignored "
                f"because a master raw-material file ({rm_path.name}) was given."
            )
    else:
        rm_path = product_path
        rm_raw = _find_sheet(product_sheets, config.SHEET_RAW_MATERIALS, product_path, issues, required=False)
        if rm_raw is None:
            issues.error(
                f"No '{config.SHEET_RAW_MATERIALS}' sheet in {product_path.name}, and no master "
                f"raw-material file was given. Add the sheet, or give the master raw-material "
                f"file with --raw-materials."
            )

    specs_raw = _find_sheet(product_sheets, config.SHEET_SPECS, product_path, issues, required=True)
    settings_raw = _find_sheet(product_sheets, config.SHEET_SETTINGS, product_path, issues, required=True)
    usage_raw = _find_sheet(product_sheets, config.SHEET_USAGE_LIMITS, product_path, issues, required=False)

    # --- Parse each sheet. Each parser records its own errors/warnings. ----
    materials: dict[str, RawMaterial] = {}
    components: list[str] | None = None  # None = unknown, so component checks are skipped
    listed_names: list[str] = []
    if rm_raw is not None:
        table = _make_table(rm_raw, config.SHEET_RAW_MATERIALS, rm_path, config.RAW_MATERIAL_FIXED_COLUMNS)
        materials, components, listed_names = _parse_raw_materials(table, issues, lab_convention)

    specs: list[SpecRow] = []
    spec_names: list[str] = []  # every name in Specs, valid or not (for the WhatIf check)
    if specs_raw is not None:
        table = _make_table(specs_raw, config.SHEET_SPECS, product_path, config.SPEC_COLUMNS)
        specs = _parse_specs(table, components, issues, listed_names)
        spec_names = [_text(row.get(config.COL_NAME)) for row in table.rows if _text(row.get(config.COL_NAME))]

    settings = Settings(product_name=product_path.stem)
    if settings_raw is not None:
        table = _make_table(settings_raw, config.SHEET_SETTINGS, product_path, config.SETTINGS_COLUMNS)
        settings = _parse_settings(table, product_path.stem, issues)

    product_limits: dict[str, UsageLimit] = {}
    if usage_raw is not None:
        table = _make_table(usage_raw, config.SHEET_USAGE_LIMITS, product_path, config.USAGE_LIMIT_COLUMNS)
        product_limits = _parse_usage_limits(table, materials, listed_names, issues)

    _warn_not_analysed(materials, specs, product_limits, issues)

    what_if_raw =_find_sheet(product_sheets, config.SHEET_WHAT_IF, product_path, issues, required=False)
    scenarios: list[Scenario] = []
    if what_if_raw is not None:
        table = _make_table(what_if_raw, config.SHEET_WHAT_IF, product_path, config.WHAT_IF_COLUMNS)
        scenarios = _parse_what_if(table, listed_names, spec_names, issues)

    if issues.errors:
        raise InputError(issues.errors, issues.warnings)

    sources = {config.SHEET_RAW_MATERIALS: str(rm_path), config.SHEET_SPECS: str(product_path),
               config.SHEET_SETTINGS: str(product_path)}
    for sheet, raw in ((config.SHEET_USAGE_LIMITS, usage_raw), (config.SHEET_WHAT_IF, what_if_raw)):
        if raw is not None:
            sources[sheet] = str(product_path)

    return BlendData(
        raw_materials=materials,
        components=components or [],
        specs=specs,
        settings=settings,
        product_limits=product_limits,
        lab_convention=lab_convention,
        warnings=issues.warnings,
        sources=sources,
        scenarios=scenarios,
    )


# ===========================================================================
# Reading sheets into simple tables
# ===========================================================================

class _Issues:
    """Collects errors and warnings while the workbook is checked."""

    def __init__(self) -> None:
        self.errors: list[str] = []
        self.warnings: list[str] = []

    def error(self, message: str) -> None:
        self.errors.append(message)

    def warn(self, message: str) -> None:
        self.warnings.append(message)


@dataclass
class _Row:
    """One non-blank data row of a sheet."""

    excel_row: int               # row number as shown in Excel (header is row 1)
    values: dict[str, object]    # column name -> raw cell value

    def get(self, column: str) -> object:
        """Cell value, or None if the column does not exist."""
        return self.values.get(column)


@dataclass
class _Table:
    """A sheet after header processing: column names plus data rows."""

    label: str                   # e.g. "RawMaterials (master.xlsx)", used in messages
    columns: list[str]
    rows: list[_Row]
    header_problems: list[str]   # blank or duplicate headers found while reading


def _read_workbook(path: Path) -> dict[str, pd.DataFrame]:
    """Read every sheet as raw cells, without letting pandas interpret the header.

    ``header=None`` keeps the header row as ordinary cells, so we can detect
    duplicate or blank headers ourselves (pandas would silently rename a
    second "CaO" to "CaO.1"). ``dtype=object`` keeps text as text, so a
    typed "abc" in a number column can be reported instead of crashing.
    """
    if not path.is_file():
        raise InputError([f"File not found: {path}"])
    try:
        return pd.read_excel(path, sheet_name=None, header=None, dtype=object, engine="openpyxl")
    except Exception as exc:  # corrupt file, wrong format, file open elsewhere, ...
        raise InputError([f"Could not read {path.name} as an Excel workbook: {exc}"]) from exc


def _normalise(text: str) -> str:
    """Lower-case and drop spaces, underscores, hyphens and '%' for lenient matching."""
    return re.sub(r"[\s_%\-]", "", str(text)).lower()


def _find_sheet(
    sheets: dict[str, pd.DataFrame], wanted: str, path: Path, issues: _Issues, required: bool
) -> pd.DataFrame | None:
    """Return the sheet called ``wanted`` (leniently matched), or None.

    If it is required and missing, an error is recorded with a suggestion.
    """
    if wanted in sheets:
        return sheets[wanted]
    for name, frame in sheets.items():
        if _normalise(name) == _normalise(wanted):
            return frame
    if required:
        suggestion = suggest_close_match(wanted, sheets)
        hint = f" Did you mean '{suggestion}'?" if suggestion else ""
        issues.error(
            f"{path.name}: required sheet '{wanted}' not found.{hint} "
            f"Sheets present: {', '.join(sheets) or 'none'}."
        )
    return None


def _make_table(raw: pd.DataFrame, sheet: str, path: Path, known_columns: Iterable[str]) -> _Table:
    """Turn raw cells into a :class:`_Table`.

    Row 1 is the header. Headers that match a known column leniently are
    renamed to the exact known name. Completely blank rows are dropped, and
    blank-header columns are dropped if they hold no data.
    """
    label = f"{sheet} ({path.name})"
    known = list(known_columns)
    if raw.empty:
        return _Table(label, [], [], [])

    problems: list[str] = []
    columns: list[str | None] = []
    for position, cell in enumerate(raw.iloc[0].tolist()):
        letter = get_column_letter(position + 1)
        if _is_blank(cell):
            has_data = not raw.iloc[1:, position].map(_is_blank).all()
            if has_data:
                problems.append(f"{label}: column {letter} has values but no header in row 1.")
            columns.append(None)
            continue
        header = str(cell).strip()
        for known_name in known:
            if _normalise(header) == _normalise(known_name):
                header = known_name
                break
        if header in columns:
            problems.append(f"{label}: column '{header}' appears more than once (column {letter}).")
            columns.append(None)  # keep the first one only
            continue
        columns.append(header)

    rows: list[_Row] = []
    for row_position in range(1, len(raw)):  # position 0 is the header
        values = {
            column: raw.iat[row_position, column_position]
            for column_position, column in enumerate(columns)
            if column is not None
        }
        if all(_is_blank(value) for value in values.values()):
            continue
        rows.append(_Row(excel_row=row_position + 1, values=values))

    return _Table(label, [c for c in columns if c is not None], rows, problems)


def _require_columns(table: _Table, required: Iterable[str], issues: _Issues) -> bool:
    """Record header problems and missing required columns. Returns True if all present."""
    for problem in table.header_problems:
        issues.error(problem)
    missing = [column for column in required if column not in table.columns]
    for column in missing:
        suggestion = _suggest_header(column, table.columns)
        hint = f" Did you mean '{suggestion}'?" if suggestion else ""
        issues.error(f"{table.label}: required column '{column}' is missing.{hint}")
    return not missing


def _suggest_header(wanted: str, headers: list[str]) -> str | None:
    """Guess which existing header was meant to be ``wanted``.

    Catches shortened headers ("Cost" for "Cost_Rs_per_t") as well as typos.
    Very short headers (e.g. a component "S") are not used for the
    shortened-name check, to avoid silly suggestions.
    """
    target = _normalise(wanted)
    for header in headers:
        short = _normalise(header)
        if len(short) >= 3 and (short in target or target in short):
            return header
    by_normalised = {_normalise(header): header for header in headers}
    close = difflib.get_close_matches(target, list(by_normalised), n=1, cutoff=0.6)
    return by_normalised[close[0]] if close else None


def _check_unknown_columns(
    table: _Table, known: Iterable[str], issues: _Issues, unknown_is_component: bool = False
) -> list[str]:
    """Handle columns that are not in ``known``.

    A column that looks like a misspelt known column is an error. Other
    unknown columns are either components (RawMaterials) or ignored with a
    warning (other sheets).

    Returns:
        The unknown columns that are not misspellings.
    """
    known = list(known)
    normalised_known = {_normalise(name): name for name in known}
    others: list[str] = []
    for column in table.columns:
        if column in known:
            continue
        close = difflib.get_close_matches(_normalise(column), list(normalised_known), n=1, cutoff=0.8)
        if close:
            consequence = "it would be read as a chemical component" if unknown_is_component else "it would be ignored"
            issues.error(
                f"{table.label}: column '{column}' looks like a misspelling of "
                f"'{normalised_known[close[0]]}'. Rename it; otherwise {consequence}."
            )
        elif unknown_is_component:
            others.append(column)
        else:
            issues.warn(f"{table.label}: column '{column}' is not used and is ignored.")
            others.append(column)
    return others


# ===========================================================================
# Cell parsing
# ===========================================================================

def _is_blank(value: object) -> bool:
    """True for empty cells: None, NaN, or text that is only spaces."""
    if isinstance(value, str):
        return value.strip() == ""
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _text(value: object) -> str:
    """Cell as stripped text ('' if blank)."""
    return "" if _is_blank(value) else str(value).strip()


def _parse_number(
    value: object,
    where: str,
    issues: _Issues,
    *,
    required: bool = False,
    default: float | None = None,
    minimum: float | None = 0.0,
    maximum: float | None = None,
) -> float | None:
    """Read a numeric cell.

    Blank cells give ``default`` (or an error if ``required``). Text that is a
    number ("12.5") is accepted. Anything else ("abc", "<0.01", TRUE, a date)
    is an error. Values outside [minimum, maximum] are errors.

    Returns:
        The number, or None if blank without default or if invalid (in which
        case an error has been recorded).
    """
    if _is_blank(value):
        if required:
            issues.error(f"{where}: value is blank.")
        return default

    number: float | None = None
    if isinstance(value, bool):  # bool is a subclass of int, so check it first
        number = None
    elif isinstance(value, (int, float)):
        number = float(value)
    elif isinstance(value, str):
        try:
            number = float(value.strip())
        except ValueError:
            number = None
    if number is None or not math.isfinite(number):
        issues.error(f"{where}: '{value}' is not a number. Use a number, or leave the cell blank.")
        return None

    if minimum is not None and number < minimum:
        if minimum == 0:
            issues.error(f"{where}: {number:g} is negative; it must be 0 or more.")
        else:
            issues.error(f"{where}: {number:g} is below the minimum allowed ({minimum:g}).")
        return None
    if maximum is not None and number > maximum:
        issues.error(f"{where}: {number:g} is above the maximum allowed ({maximum:g}).")
        return None
    return number


def _parse_choice(value: object, choices: tuple[str, ...], where: str, issues: _Issues) -> str | None:
    """Read a text cell that must be one of ``choices``.

    Matching ignores case, and spaces/hyphens count as underscores, so
    "As received" is read as "as_received". Blank returns None.
    """
    if _is_blank(value):
        return None
    text = re.sub(r"[\s\-]+", "_", str(value).strip().lower())
    if text in choices:
        return text
    issues.error(f"{where}: '{value}' is not allowed; use one of {', '.join(choices)}.")
    return None


def _parse_expression_cell(
    value: object, where: str, known: list[str] | None, issues: _Issues, required: bool,
    raw_material_names: Iterable[str] = (),
) -> tuple[Term, ...] | None:
    """Read a Numerator/Denominator cell (via :func:`parse_expression`) and check its components exist.

    A name that is not a component gets a message saying what it is instead:
    a property column (LOI_%, Moisture_%, the cost or usage columns), a raw
    material (which belongs in UsageLimits), or a likely typo.

    Args:
        known: Component columns of RawMaterials, or None to skip the check
            (when RawMaterials itself could not be read).
        raw_material_names: Every name listed in RawMaterials.

    Returns:
        The terms (``()`` for an optional blank cell), or None if invalid.
    """
    if _is_blank(value):
        if required:
            issues.error(f"{where}: value is blank; name at least one component.")
            return None
        return ()
    try:
        terms = parse_expression(str(value), known)
    except ValueError as exc:
        issues.error(f"{where}: {exc}.")
        return None

    if known is None:
        return tuple(terms)
    property_columns = {_normalise(column) for column in config.RAW_MATERIAL_FIXED_COLUMNS}
    materials_lower = {material.lower() for material in raw_material_names}
    ok = True
    for name in (term.component for term in terms):
        if name in known:
            continue
        ok = False
        suggestion = suggest_close_match(name, known)
        if _normalise(name) in property_columns:  # LOI_%, Moisture_%, Cost_Rs_per_t, ... (leniently matched)
            issues.error(f"{where}: '{name}' is a property column, not a component, so it can't be used in a spec.")
        elif name.lower() in materials_lower:
            issues.error(f"{where}: '{name}' is a raw material, not a component. To limit how much of it is used, "
                         f"put it in {config.SHEET_USAGE_LIMITS}.")
        elif suggestion:
            issues.error(
                f"{where}: component '{name}' is not a column in RawMaterials. "
                f"Did you mean '{suggestion}'? (Component names must match exactly.)"
            )
        else:
            issues.error(
                f"{where}: component '{name}' is not a column in RawMaterials. "
                f"Components available: {', '.join(known)}."
            )
    return tuple(terms) if ok else None


# ===========================================================================
# Sheet parsers
# ===========================================================================

def _parse_raw_materials(
    table: _Table, issues: _Issues, convention: str
) -> tuple[dict[str, RawMaterial], list[str] | None, list[str]]:
    """Parse the RawMaterials sheet.

    Returns:
        (materials, components, listed_names).
        ``materials`` holds only the valid rows. ``components`` is None if
        the sheet is too broken to know them (missing required columns).
        ``listed_names`` holds every name in the sheet, valid or not, so a
        UsageLimits row for a material with a typo in its cost is not also
        reported as "not a raw material".
    """
    if not _require_columns(table, config.RAW_MATERIAL_REQUIRED_COLUMNS, issues):
        return {}, None, []
    components = _check_unknown_columns(
        table, config.RAW_MATERIAL_FIXED_COLUMNS, issues, unknown_is_component=True
    )
    if not components:
        issues.error(f"{table.label}: no chemical component columns found (e.g. one column per oxide).")

    materials: dict[str, RawMaterial] = {}
    listed_names: list[str] = []
    seen: dict[str, int] = {}  # lower-case name -> Excel row where first seen
    for row in table.rows:
        name = _text(row.get(config.COL_NAME))
        if not name:
            issues.error(f"{table.label} row {row.excel_row}: {config.COL_NAME} is blank.")
            continue
        where = f"{table.label} row {row.excel_row} '{name}'"
        if name.lower() in seen:
            issues.error(
                f"{where}: duplicate name (already used in row {seen[name.lower()]}; "
                f"names must be unique, ignoring case)."
            )
            continue
        seen[name.lower()] = row.excel_row
        listed_names.append(name)
        errors_before = len(issues.errors)

        cost = _parse_number(row.get(config.COL_COST), f"{where}, {config.COL_COST}", issues, required=True)
        moisture = _parse_number(row.get(config.COL_MOISTURE), f"{where}, {config.COL_MOISTURE}", issues,
                                 default=0.0, maximum=100.0)
        loi = _parse_number(row.get(config.COL_LOI), f"{where}, {config.COL_LOI}", issues,
                            default=0.0, maximum=100.0)
        for value, column in ((moisture, config.COL_MOISTURE), (loi, config.COL_LOI)):
            if value is not None and value >= 100.0:
                issues.error(f"{where}, {column}: must be below 100.")

        price_basis = _parse_choice(row.get(config.COL_PRICE_BASIS), config.PRICE_BASES,
                                    f"{where}, {config.COL_PRICE_BASIS}", issues)
        if price_basis is None and _is_blank(row.get(config.COL_PRICE_BASIS)):
            if moisture:  # blank basis only matters if there is moisture
                issues.error(
                    f"{where}, {config.COL_PRICE_BASIS}: blank, but {config.COL_MOISTURE} is "
                    f"{moisture:g}. Say whether {config.COL_COST} is "
                    f"'{config.PRICE_AS_RECEIVED}' or '{config.PRICE_DRY}'."
                )
            else:
                price_basis = config.PRICE_DRY  # no moisture: both bases give the same price
        if price_basis == config.PRICE_AS_RECEIVED and _is_blank(row.get(config.COL_MOISTURE)):
            issues.warn(
                f"{where}: price is '{config.PRICE_AS_RECEIVED}' but {config.COL_MOISTURE} is "
                f"blank, so moisture is taken as 0 and the dry price equals the quoted price."
            )

        min_use = _parse_number(row.get(config.COL_MIN_USE), f"{where}, {config.COL_MIN_USE}", issues,
                                default=0.0, maximum=100.0)
        max_use = _parse_number(row.get(config.COL_MAX_USE), f"{where}, {config.COL_MAX_USE}", issues,
                                default=100.0, maximum=100.0)
        if min_use is not None and max_use is not None and min_use > max_use:
            issues.error(
                f"{where}: {config.COL_MIN_USE} ({min_use:g}) is greater than "
                f"{config.COL_MAX_USE} ({max_use:g})."
            )

        reported = {
            component: _parse_number(row.get(component), f"{where}, {component}", issues,
                                     default=0.0, maximum=100.0)
            for component in components
        }

        if len(issues.errors) > errors_before:
            continue  # this row has errors; don't build a half-valid material

        warning = analysis_total_warning(name, sum(reported.values()), loi, convention)
        if warning:
            issues.warn(warning)

        materials[name] = RawMaterial(
            name=name,
            cost_input=cost,
            price_basis=price_basis,
            moisture_pct=moisture,
            loi_pct=loi,
            plant_limit=UsageLimit(min_use, max_use),
            composition=analysis_to_dry_basis(reported, loi, convention),
            # Blank = "not analysed"; 0 = "analysed, none present". Both count as 0.
            not_analysed=frozenset(component for component in components if _is_blank(row.get(component))),
        )

    if not table.rows:
        issues.error(f"{table.label}: no raw materials listed.")
    return materials, components, listed_names


def _warn_not_analysed(
    materials: dict[str, RawMaterial], specs: list[SpecRow], product_limits: dict[str, UsageLimit], issues: _Issues
) -> None:
    """Warn about blank analysis cells that could change the answer.

    A blank cell means "not analysed" and a 0 means "analysed, none present";
    both count as 0. A blank matters when a spec uses that component and the
    material may be used (effective Max_use_% above 0): the recipe may then
    rely on a value nobody measured. One warning per material and component.
    """
    used: list[str] = []  # components used by any spec, in the order first used
    for spec in specs:
        for term in spec.numerator + spec.denominator:
            if term.component not in used:
                used.append(term.component)
    for name, material in materials.items():
        allowed = material.plant_limit.tighter(product_limits.get(name, UsageLimit())).max_pct > 0
        if not allowed:
            continue
        for component in used:
            if component in material.not_analysed:
                issues.warn(f"RawMaterials '{name}': {component} is blank (not analysed) and is counted as 0, "
                            f"but a spec uses it. Type 0 if it is genuinely absent.")


def _parse_specs(table: _Table, components: list[str] | None, issues: _Issues,
                 raw_material_names: list[str] | None = None) -> list[SpecRow]:
    """Parse the Specs sheet (one row per limit).

    Args:
        raw_material_names: Every name in RawMaterials, so a raw material
            typed in Specs gets a message pointing to UsageLimits.
    """
    raw_material_names = raw_material_names or []
    if not _require_columns(table, config.SPEC_REQUIRED_COLUMNS, issues):
        return []
    _check_unknown_columns(table, config.SPEC_COLUMNS, issues)

    specs: list[SpecRow] = []
    seen: dict[str, int] = {}
    for row in table.rows:
        name = _text(row.get(config.COL_NAME))
        if not name:
            issues.error(f"{table.label} row {row.excel_row}: {config.COL_NAME} is blank.")
            continue
        where = f"{table.label} row {row.excel_row} '{name}'"
        if name.lower() in seen:
            issues.error(
                f"{where}: duplicate name (already used in row {seen[name.lower()]}; "
                f"spec names become constraint names, so they must be unique)."
            )
            continue
        seen[name.lower()] = row.excel_row
        errors_before = len(issues.errors)

        numerator = _parse_expression_cell(row.get(config.COL_NUMERATOR), f"{where}, {config.COL_NUMERATOR}",
                                           components, issues, required=True,
                                           raw_material_names=raw_material_names)
        denominator = _parse_expression_cell(row.get(config.COL_DENOMINATOR),
                                             f"{where}, {config.COL_DENOMINATOR}", components, issues,
                                             required=False, raw_material_names=raw_material_names)
        min_value = _parse_number(row.get(config.COL_MIN), f"{where}, {config.COL_MIN}", issues)
        max_value = _parse_number(row.get(config.COL_MAX), f"{where}, {config.COL_MAX}", issues)
        margin = _parse_number(row.get(config.COL_MARGIN), f"{where}, {config.COL_MARGIN}", issues, default=0.0)

        if len(issues.errors) > errors_before:
            continue

        spec = SpecRow(name, numerator, denominator, min_value, max_value, margin)
        problems = spec_limit_problems(spec)
        for problem in problems:
            issues.error(f"{where}: {problem}")
        if not problems:
            specs.append(spec)

    if not table.rows:
        issues.warn(f"{table.label}: no spec rows; the cheapest raw material mix will be chosen without limits.")
    return specs


def spec_limit_problems(spec: SpecRow) -> list[str]:
    """Check that a spec's limits make sense, before and after the margin.

    Used for the Specs sheet and for what-if spec changes, so both follow
    the same rules.

    Returns:
        Problem descriptions; empty if the limits are fine.
    """
    if spec.min_value is None and spec.max_value is None:
        return [f"give a {config.COL_MIN}, a {config.COL_MAX}, or both."]
    for label, value in ((config.COL_MIN, spec.min_value), (config.COL_MAX, spec.max_value),
                         (config.COL_MARGIN, spec.margin)):
        if value is not None and value < 0:
            return [f"{label} {value:g} is negative; it must be 0 or more."]
    if spec.min_value is not None and spec.max_value is not None and spec.min_value > spec.max_value:
        return [f"{config.COL_MIN} ({spec.min_value:g}) is greater than {config.COL_MAX} ({spec.max_value:g})."]

    problems = []
    if spec.min_target is not None and spec.max_target is not None and spec.min_target > spec.max_target:
        problems.append(
            f"after the margin the internal range is empty: Min + margin = {spec.min_target:g} "
            f"is greater than Max - margin = {spec.max_target:g}. Reduce the margin."
        )
    if spec.max_target is not None and spec.max_target < 0:
        problems.append(f"Max - margin = {spec.max_target:g} is negative. Reduce the margin.")
    # A plain composition can never exceed 100 %. With a factor above 1 it can
    # (Al2O3+1.8895*Al counts metallic Al as more mass of Al2O3), so this quick
    # check only applies to plain expressions; impossible limits with factors
    # are found by the diagnosis pre-check instead.
    plain = all(term.factor <= 1 for term in spec.numerator)
    if not spec.is_ratio and plain and spec.min_target is not None and spec.min_target > 100:
        problems.append(f"Min + margin = {spec.min_target:g} % is above 100 % and can never be met.")
    return problems


def _parse_settings(table: _Table, default_name: str, issues: _Issues) -> Settings:
    """Parse the Settings sheet (columns Key, Value)."""
    if not _require_columns(table, config.SETTINGS_COLUMNS, issues):
        return Settings(product_name=default_name)
    _check_unknown_columns(table, config.SETTINGS_COLUMNS, issues)

    found: dict[str, tuple[object, str]] = {}  # key -> (value, where)
    for row in table.rows:
        raw_key = _text(row.get(config.COL_KEY))
        where = f"{table.label} row {row.excel_row}"
        if not raw_key:
            issues.error(f"{where}: the row has values but {config.COL_KEY} is blank.")
            continue
        key = next((k for k in config.SETTINGS_KEYS if _normalise(k) == _normalise(raw_key)), None)
        if key is None:
            suggestion = suggest_close_match(raw_key, config.SETTINGS_KEYS)
            hint = f" Did you mean '{suggestion}'?" if suggestion else ""
            issues.warn(f"{where}: unknown setting '{raw_key}' is ignored.{hint}")
            continue
        if key in found:
            issues.error(f"{where}: setting '{key}' appears more than once.")
            continue
        found[key] = (row.get(config.COL_VALUE), f"{where} '{key}'")

    product_name = default_name
    if config.KEY_PRODUCT_NAME in found and _text(found[config.KEY_PRODUCT_NAME][0]):
        product_name = _text(found[config.KEY_PRODUCT_NAME][0])
    else:
        issues.warn(f"{table.label}: {config.KEY_PRODUCT_NAME} not given; using '{default_name}'.")

    spec_basis = config.DEFAULT_SPEC_BASIS
    if config.KEY_SPEC_BASIS in found:
        value, where = found[config.KEY_SPEC_BASIS]
        spec_basis = _parse_choice(value, config.SPEC_BASES, where, issues) or config.DEFAULT_SPEC_BASIS

    batch_size = _positive_setting(found, config.KEY_BATCH_SIZE, config.DEFAULT_BATCH_SIZE_KG, issues)
    # Optional: blank or missing means the config.py value
    resolution = _positive_setting(found, config.KEY_SCALE_RESOLUTION, config.SCALE_RESOLUTION_KG, issues)

    return Settings(product_name=product_name, spec_basis=spec_basis, batch_size_kg=batch_size,
                    scale_resolution_kg=resolution)


def _positive_setting(found: dict[str, tuple[object, str]], key: str, default: float, issues: _Issues) -> float:
    """A numeric setting that must be > 0; blank or missing gives ``default``."""
    if key not in found:
        return default
    value, where = found[key]
    number = _parse_number(value, where, issues, default=default)
    if number is not None and number <= 0:
        issues.error(f"{where}: must be greater than 0.")
        return default
    return default if number is None else number


def _parse_usage_limits(
    table: _Table, materials: dict[str, RawMaterial], listed_names: list[str], issues: _Issues
) -> dict[str, UsageLimit]:
    """Parse the optional UsageLimits sheet and check it against the plant-wide limits.

    Args:
        materials: Valid raw materials (for the plant-wide limit check).
        listed_names: Every name in RawMaterials (for the "does it exist" check).
    """
    if not _require_columns(table, (config.COL_NAME,), issues):
        return {}
    _check_unknown_columns(table, config.USAGE_LIMIT_COLUMNS, issues)
    if config.COL_MIN_USE not in table.columns and config.COL_MAX_USE not in table.columns:
        issues.error(f"{table.label}: needs a {config.COL_MIN_USE} and/or {config.COL_MAX_USE} column.")
        return {}

    limits: dict[str, UsageLimit] = {}
    for row in table.rows:
        name = _text(row.get(config.COL_NAME))
        if not name:
            issues.error(f"{table.label} row {row.excel_row}: {config.COL_NAME} is blank.")
            continue
        where = f"{table.label} row {row.excel_row} '{name}'"
        if listed_names and name not in listed_names:
            suggestion = suggest_close_match(name, listed_names)
            hint = f" Did you mean '{suggestion}'?" if suggestion else ""
            issues.error(f"{where}: not a raw material in RawMaterials.{hint}")
            continue
        if name in limits:
            issues.error(f"{where}: listed more than once.")
            continue

        min_use = _parse_number(row.get(config.COL_MIN_USE), f"{where}, {config.COL_MIN_USE}", issues,
                                default=0.0, maximum=100.0)
        max_use = _parse_number(row.get(config.COL_MAX_USE), f"{where}, {config.COL_MAX_USE}", issues,
                                default=100.0, maximum=100.0)
        if min_use is None or max_use is None:
            continue
        if min_use > max_use:
            issues.error(f"{where}: {config.COL_MIN_USE} ({min_use:g}) is greater than {config.COL_MAX_USE} ({max_use:g}).")
            continue

        limit = UsageLimit(min_use, max_use)
        if name in materials:
            plant = materials[name].plant_limit
            effective = plant.tighter(limit)
            if effective.min_pct > effective.max_pct:
                issues.error(
                    f"{where}: the effective limits conflict. Effective Min = {effective.min_pct:g} % "
                    f"(plant {plant.min_pct:g}, product {limit.min_pct:g}) is greater than effective "
                    f"Max = {effective.max_pct:g} % (plant {plant.max_pct:g}, product {limit.max_pct:g})."
                )
                continue
        limits[name] = limit
    return limits


# What each WhatIf change type changes: (kind of item, field name in what_if.apply_changes)
_WHAT_IF_TARGETS = {
    config.CHANGE_PRICE: ("raw material", None),
    config.CHANGE_SPEC_MIN: ("spec", config.COL_MIN),
    config.CHANGE_SPEC_MAX: ("spec", config.COL_MAX),
    config.CHANGE_SPEC_MARGIN: ("spec", config.COL_MARGIN),
    config.CHANGE_USAGE_MIN: ("raw material", config.COL_MIN),
    config.CHANGE_USAGE_MAX: ("raw material", config.COL_MAX),
}


def _parse_what_if(
    table: _Table, material_names: list[str], spec_names: list[str], issues: _Issues
) -> list[Scenario]:
    """Parse the optional WhatIf sheet into scenarios (rows grouped by Scenario name).

    Each row is checked on its own here: known change type, existing item,
    valid number. Whether a scenario's changes make sense together (e.g. a
    new Min above the Max) is checked when the scenarios are run
    (``what_if.check_scenarios``), because that needs the changes applied.
    """
    if not _require_columns(table, config.WHAT_IF_COLUMNS, issues):
        return []
    _check_unknown_columns(table, config.WHAT_IF_COLUMNS, issues)

    grouped: dict[str, dict[str, dict]] = {}   # scenario -> {"price": {}, "spec": {}, "usage": {}}
    seen: dict[tuple[str, str, str], int] = {}
    for row in table.rows:
        scenario = _text(row.get(config.COL_SCENARIO))
        raw_type = _text(row.get(config.COL_CHANGE_TYPE))
        item = _text(row.get(config.COL_ITEM))
        where = f"{table.label} row {row.excel_row}" + (f" '{scenario}'" if scenario else "")
        if not scenario:
            issues.error(f"{where}: {config.COL_SCENARIO} is blank; give the scenario a name.")
            continue
        change_type = next((known for known in config.CHANGE_TYPES if _normalise(known) == _normalise(raw_type)), None)
        if change_type is None:
            suggestion = suggest_close_match(raw_type, config.CHANGE_TYPES) if raw_type else None
            hint = f" Did you mean '{suggestion}'?" if suggestion else ""
            issues.error(f"{where}: {config.COL_CHANGE_TYPE} '{raw_type}' is not known.{hint} "
                         f"Use one of: {', '.join(config.CHANGE_TYPES)}.")
            continue
        kind, field_name = _WHAT_IF_TARGETS[change_type]
        known = material_names if kind == "raw material" else spec_names
        if not item:
            issues.error(f"{where}: {config.COL_ITEM} is blank; give the {kind} to change.")
            continue
        if known and item not in known:
            suggestion = suggest_close_match(item, known)
            hint = f" Did you mean '{suggestion}'?" if suggestion else ""
            issues.error(f"{where}: '{item}' is not a {kind} in {'RawMaterials' if kind == 'raw material' else 'Specs'}.{hint}")
            continue
        key = (scenario, change_type, item)
        if key in seen:
            issues.error(f"{where}: '{change_type}' for '{item}' is already given in row {seen[key]} of this scenario.")
            continue
        seen[key] = row.excel_row

        value_where = f"{where}, {config.COL_NEW_VALUE}"
        maximum = 100.0 if change_type in (config.CHANGE_USAGE_MIN, config.CHANGE_USAGE_MAX) else None
        errors_before = len(issues.errors)
        value = _parse_number(row.get(config.COL_NEW_VALUE), value_where, issues,
                              required=change_type == config.CHANGE_PRICE, maximum=maximum)
        if len(issues.errors) > errors_before:
            continue

        changes = grouped.setdefault(scenario, {"price": {}, "spec": {}, "usage": {}})
        if change_type == config.CHANGE_PRICE:
            changes["price"][item] = value
        elif kind == "spec":
            changes["spec"].setdefault(item, {})[field_name] = value
        else:
            changes["usage"].setdefault(item, {})[field_name] = value

    return [Scenario(name, changes["price"], changes["spec"], changes["usage"]) for name, changes in grouped.items()]

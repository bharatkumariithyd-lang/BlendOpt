"""Find the least-cost blend for one product, plus any what-if scenarios.

Run from the project folder, for example:

    .venv\\Scripts\\python.exe run.py templates\\Test1_hand_worked.xlsx
    .venv\\Scripts\\python.exe run.py data\\my_product.xlsx --raw-materials data\\raw_materials.xlsx
    .venv\\Scripts\\python.exe run.py data\\my_product.xlsx --output output\\my_product_june.xlsx

What it does:
    1. Reads and checks the workbook(s). Errors stop the run (all are listed);
       warnings are printed and the run continues.
    2. Checks the WhatIf scenarios, if the product workbook has a WhatIf sheet.
    3. Solves the base case; if it is infeasible, explains why.
    4. Works out the price range of each raw material for the base case.
    5. Solves each scenario and compares it with the base case; an
       infeasible scenario is explained in the same way.
    6. Writes one report workbook (default: output\\<product>_result.xlsx)
       with the .lp file next to it, and prints a summary.
"""

import argparse
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from blend_optimizer import solve_and_report as sr
from blend_optimizer.data import InputError
from blend_optimizer.inputs import load_inputs
from blend_optimizer.pipeline import RunOutcome, optimise
from blend_optimizer.what_if import (
    SHEET_PRICE_RANGES,
    ScenarioOutcome,
    check_scenarios,
    format_scenario,
    price_ranging_report,
    run_scenarios,
    scenario_sheets,
)

DEFAULT_OUTPUT_FOLDER = Path(__file__).resolve().parent / "output"


@dataclass
class FullRun:
    """The base case, its price ranges and every scenario of one run."""

    base: RunOutcome
    price_ranges: pd.DataFrame
    scenarios: list[ScenarioOutcome] = field(default_factory=list)


def format_price_ranges(table: pd.DataFrame, notes: list[str], scale_resolution_kg: float) -> str:
    """The Price ranges sheet as console text."""
    lines = ["Price ranges (base case): today's recipe stays the same while one price moves within this band"]
    if table.empty:
        lines.append(f"  {notes[0]}")
    else:
        lines.append(sr.indented_table(table, ["Raw material", "Price basis", "Current price (Rs/t)",
                                               "Unchanged from (Rs/t)", "Unchanged up to (Rs/t)"],
                                       scale_resolution_kg))
    return "\n".join(lines)


def default_output_path(product_file: str | Path) -> Path:
    """output/<product file name>_result.xlsx"""
    return DEFAULT_OUTPUT_FOLDER / f"{Path(product_file).stem}_result.xlsx"


def run(
    product_file: str | Path,
    raw_materials_file: str | Path | None = None,
    output_path: str | Path | None = None,
    show: bool = True,
) -> FullRun:
    """Run the base case (with its price ranges) and all WhatIf scenarios; write one report workbook.

    Args:
        product_file: Product workbook (or the single workbook with everything).
        raw_materials_file: Optional master raw-material workbook.
        output_path: Report workbook; default output/<product>_result.xlsx.
        show: Print the summary to the console.

    Raises:
        InputError: If the workbook(s) or the scenarios contain errors.
    """
    data = load_inputs(product_file, raw_materials_file)
    scenario_errors = check_scenarios(data)
    if scenario_errors:
        raise InputError(scenario_errors, data.warnings)

    price_ranges, price_notes = price_ranging_report(data)
    scenarios = run_scenarios(data)
    output_path = Path(output_path) if output_path is not None else default_output_path(product_file)
    # The base report's sheet names, so scenario sheets never clash with them
    reserved = {"summary", "recipe", "composition", "spec check", "cost", "shadow prices", "reduced costs",
                "diagnosis", SHEET_PRICE_RANGES.lower()}
    extra_sheets = {SHEET_PRICE_RANGES: price_ranges, **scenario_sheets(data, scenarios, reserved)}
    base = optimise(data, output_path, show=show, extra_sheets=extra_sheets,
                    extra_notes={SHEET_PRICE_RANGES: price_notes})
    if show:
        resolution = data.settings.scale_resolution_kg
        print()
        print(format_price_ranges(price_ranges, price_notes, resolution))
        for outcome in scenarios:
            print()
            print(format_scenario(outcome, resolution))
    return FullRun(base, price_ranges, scenarios)


def main(argv: list[str] | None = None) -> int:
    """Command-line entry point. Returns the exit code (0 = report written)."""
    parser = argparse.ArgumentParser(
        description="Least-cost blend optimizer: find the cheapest raw material mix that meets the specs.")
    parser.add_argument("product_file", help="product workbook (.xlsx)")
    parser.add_argument("--raw-materials", dest="raw_materials_file", default=None,
                        help="master raw-material workbook, if RawMaterials is not in the product workbook")
    parser.add_argument("--output", dest="output_path", default=None,
                        help="report workbook to write (default: output\\<product>_result.xlsx)")
    parser.add_argument("--quiet", action="store_true", help="do not print the summary")
    arguments = parser.parse_args(argv)

    try:
        full_run = run(arguments.product_file, arguments.raw_materials_file, arguments.output_path,
                       show=not arguments.quiet)
    except InputError as error:
        print(error)
        if error.warnings:
            print(f"{len(error.warnings)} warning(s):")
            for warning in error.warnings:
                print(f"  - {warning}")
        print("Nothing was solved. Correct the workbook and run again.")
        return 1
    except PermissionError as error:
        print(f"Could not write the report: {error}")
        print("If the report is open in Excel, close it and run again.")
        return 1

    print()
    print(f"Report written to: {full_run.base.excel_path}")
    print(f"LP file:           {full_run.base.lp_path}")
    if full_run.scenarios:
        print(f"Scenarios:         {len(full_run.scenarios)} (see the Scenarios sheet)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""Create the blank templates and the example workbooks in the templates/ folder.

Run from the project folder:

    .venv\\Scripts\\python.exe make_templates.py

Blank templates (fill these in for real products):
    Template_single_product.xlsx         layout 1: everything in one workbook
    Template_master_raw_materials.xlsx   layout 2: master raw-material list
    Template_product.xlsx                layout 2: per-product specs and settings

Examples (the test datasets from blend_optimizer/examples.py):
    Test1_*.xlsx, Test2_*.xlsx, Example_master_raw_materials.xlsx,
    Test2_product_only.xlsx, Broken_example.xlsx
"""

from pathlib import Path

from blend_optimizer import config, examples
from blend_optimizer.templates import write_blank_templates

DEFAULT_FOLDER = Path(__file__).resolve().parent / "templates"


def main(folder: Path = DEFAULT_FOLDER) -> list[Path]:
    """Write all templates and example workbooks into ``folder``; return their paths."""
    # Pre-fill the blank RawMaterials sheet with the usual slag components
    written = write_blank_templates(folder, components=examples.SYNTHETIC_SLAG_COMPONENTS)

    single_workbooks = {
        "Test1_hand_worked.xlsx": examples.hand_worked_example(),
        "Test1_infeasible_Al2O3_min_45.xlsx": examples.hand_worked_example(al2o3_min=45),
        "Test2_synthetic_slag_dry.xlsx": examples.synthetic_slag_example(),
        "Test2_synthetic_slag_ignited.xlsx": examples.synthetic_slag_example(spec_basis=config.BASIS_IGNITED),
        "Test2_synthetic_slag_margin_0.5.xlsx": examples.synthetic_slag_example(margin=0.5),
        "Test2_synthetic_slag_fluorspar_ban.xlsx": examples.synthetic_slag_example(ban_fluorspar=True),
        "Test2_synthetic_slag_with_what_if.xlsx": examples.synthetic_slag_example(with_what_if=True),
        "Broken_example.xlsx": examples.broken_example(),
    }
    for file_name, example in single_workbooks.items():
        written.append(example.write(folder / file_name))

    # Layout 2 demonstration: Test 2 split into master + product workbooks
    master, product = examples.synthetic_slag_example().split()
    written.append(master.write(folder / "Example_master_raw_materials.xlsx"))
    written.append(product.write(folder / "Test2_product_only.xlsx"))
    return written


if __name__ == "__main__":
    for path in main():
        print(f"wrote {path}")

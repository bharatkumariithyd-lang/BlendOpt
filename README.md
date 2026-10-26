# Least-cost blend optimizer

This tool finds the **cheapest mix of raw materials** that meets a customer's composition specification:
for example a synthetic slag, a desulphurising compound or a deoxidiser. You describe the raw materials
(price, moisture, LOI, analysis) and the product specification in an Excel workbook. The tool works out the
recipe, shows the blend analysis, and tells you which limits are driving the cost and what it would save
to relax them. If the specification cannot be met, it explains why.

It also lets you try **what-if scenarios** (a price change, a tighter limit, a banned material) and compare
each one with today's recipe.

Contents:
1. [Before you start](#1-before-you-start)
2. [Filling in the workbook](#2-filling-in-the-workbook)
3. [Running the optimizer](#3-running-the-optimizer)
4. [Reading the report](#4-reading-the-report)
5. [Good to know](#5-good-to-know)
6. [For whoever maintains the tool](#6-for-whoever-maintains-the-tool)

---

## 1. Before you start

The tool lives in one folder (here: `D:\VS Python\Cost Optimizer`). The parts you will use:

| Folder / file | What it is |
|---|---|
| `templates\` | Blank workbooks to copy, and worked examples |
| `data\` | Put your real product workbooks here (this folder is kept private, it is not shared with the code) |
| `output\` | The reports appear here |
| `run.py` | The program you run |
| `.venv\` | The Python installation the tool uses. Leave it alone. |

**Opening a command window in the folder.** Either open the folder in VS Code and choose
*Terminal → New Terminal*, or open the folder in File Explorer, click in the address bar, type
`powershell` and press Enter.

**Check that it works** by running the first worked example:

```
.venv\Scripts\python.exe run.py templates\Test1_hand_worked.xlsx
```

You should see a recipe of Lime 53.40 %, Calcined alumina 6.02 % and Bauxite 40.58 %, costing
Rs 11,482 per dry tonne, and the line `Report written to: ...\output\Test1_hand_worked_result.xlsx`.

---

## 2. Filling in the workbook

### Two ways to organise your workbooks

* **One workbook per product.** Copy `templates\Template_single_product.xlsx`. It holds everything:
  raw materials, specs, settings, and the optional usage limits and what-if scenarios.
* **A master raw-material list plus one workbook per product.** Keep all raw materials (prices, analyses)
  in one master workbook (copy `templates\Template_master_raw_materials.xlsx`) and each product's specs in
  its own workbook (copy `templates\Template_product.xlsx`). When a price changes you update it once.

Copy the template into `data\` and rename it, e.g. `data\Slag_SS45.xlsx`. Every template has a README
sheet with a short version of the instructions below.

### General rules

* Keep the column headers in **row 1** and do not rename them. Blank rows are ignored.
* Type **plain numbers**: no units, no `<0.01`, no `nd`. A blank analysis cell counts as 0.
* Names (raw materials, components, specs) are matched **exactly**: `Al2O3` and `AL2O3` are different.
  If you make a typo, the error message suggests the name you probably meant.

### Sheet `RawMaterials`: one row per raw material

| Column | What to enter |
|---|---|
| `Name` | A unique name, e.g. `Calcined bauxite` |
| `Cost_Rs_per_t` | The price in Rs per tonne, as quoted |
| `Price_basis` | `as_received` (price per tonne as delivered, with its moisture) or `dry` (price per dry tonne). The tool converts everything to Rs per **dry** tonne: Rs 9,000/t as received with 0.5 % moisture is Rs 9,045 per dry tonne. May be left blank only if moisture is 0. |
| `Moisture_%` | Free moisture, wt % as received. Blank = 0. |
| `LOI_%` | Loss on ignition, wt % on a dry basis. Blank = 0. |
| `Min_use_%`, `Max_use_%` | Plant-wide limits on how much of this material may be in any blend, wt % of the dry blend. Blank = 0 and 100. E.g. `Max_use_% = 20` for a fines grade that segregates. |
| every other column | One column per chemical component (`CaO`, `SiO2`, `Al2O3`, `MgO`, `CaF2`, `Fe2O3`, `S`, `P` ...), wt %. The column header is the component name. Add, remove or reorder these columns freely. |

**Which analysis basis?** The analyses are taken to be on a **dry basis with LOI reported separately**,
so for each material the components plus LOI should add up to about 100 %. If a material totals less
than 97 % or more than 103 %, you get a warning (the run still continues). The warning also says if the
numbers look like an LOI-free (ignited-basis) analysis instead. If your lab reports that way, ask whoever
maintains the tool to switch it over (section 6).

### Sheet `Specs`: one row per limit

| Column | What to enter |
|---|---|
| `Name` | A unique label. In the report the limits are called `<Name>_min` and `<Name>_max`. |
| `Numerator` | What is limited: a component, or several joined with `+`, each optionally with a conversion factor (see below), e.g. `CaO+MgO` or `Al2O3+1.8895*Al` |
| `Denominator` | Leave blank for a composition limit. Fill in for a **ratio**, e.g. Numerator `CaO`, Denominator `SiO2` for basicity. Same syntax as Numerator. |
| `Min`, `Max` | The customer's limits. Leave one blank if there is no limit on that side; at least one is needed. |
| `Margin` | Optional safety margin, in the same unit as the limit. The optimizer aims inside the customer limits: Min + margin and Max − margin. Blank = 0. |

Examples:

| Name | Numerator | Denominator | Min | Max | Margin | Meaning |
|---|---|---|---|---|---|---|
| CaO | CaO | | 50 | 56 | 0.5 | CaO between 50 and 56 %; the optimizer aims for 50.5–55.5 % |
| SiO2 | SiO2 | | | 6 | | SiO2 at most 6 % |
| CaO+MgO | CaO+MgO | | 55 | | | CaO plus MgO at least 55 % |
| Basicity | CaO | SiO2 | 8 | | | CaO / SiO2 at least 8 |
| P | P | | | 0.03 | | P at most 0.03 % |
| Total Al2O3 | Al2O3+1.8895*Al | | 75 | | | Al2O3, with metallic Al counted as Al2O3, at least 75 % |

**Counting one component as another (conversion factors).** Some customers count a component in
another form. Metallic Al, for example, oxidises to Al2O3 in the slag, so it is counted in Al2O3 by
stoichiometry. Write the factor in front of the component with `*`:

```
Al2O3+1.8895*Al
```

This means "Al2O3 % + 1.8895 × Al %" for every raw material. 1.8895 is the mass of Al2O3 formed per unit
mass of Al: 101.96 / (2 × 26.98), the molar mass of Al2O3 divided by its two atoms of Al. Spaces around `+`
and `*` are fine; every term needs a positive factor (or none, meaning 1), and `-` is not allowed. The
factor works everywhere a plain component does: on every spec basis, in ratios (also in the Denominator),
with margins, in WhatIf changes, and in the diagnosis. The report shows the expression as you typed it.
The tool knows no chemistry: the factor is simply the number you type.

Common factors (molar mass of what it is counted as, per atom of the shared element, divided by the molar
mass of what was analysed):

| Analysed as | Counted as | Factor | Derivation |
|---|---|---|---|
| Al | Al2O3 | 1.8895 | 101.96 / (2 × 26.98) |
| Mg | MgO | 1.6583 | 40.304 / 24.305 |
| Ca | CaO | 1.3992 | 56.077 / 40.078 |
| Si | SiO2 | 2.1393 | 60.084 / 28.086 |
| CaC2 | CaO | 0.8748 | 56.077 / 64.100 |
| FeO | Fe2O3 | 1.1113 | 159.687 / (2 × 71.844) |
| CaF2 | F | 0.4867 | (2 × 18.998) / 78.075 |

So `CaO+1.3992*Ca+0.8748*CaC2` would count metallic Ca and carbide as CaO. If you used a helper column in
RawMaterials for this before (e.g. a column `Al_as_Al2O3` = 1.8895 × Al), it is no longer needed:
put the factor in the Specs sheet instead and delete the column.

### Sheet `Settings`: two columns, `Key` and `Value`

| Key | Value |
|---|---|
| `Product_name` | Shown on the report |
| `Spec_basis` | The basis the customer's composition limits are written on: `dry` (default), `ignited` (after loss on ignition, e.g. as the slag is in the ladle), or `as_received` (including the blend's moisture). Ratios such as basicity are the same on every basis. |
| `Batch_size_kg` | The total mass weighed into the mixer per batch, **as received** (default 1000) |
| `Scale_resolution_kg` | Optional. The resolution of your weighing scale, e.g. `0.5`; the weigh sheet is rounded to it. Blank = 0.1 kg. |

### Sheet `UsageLimits` (optional): limits for this product or customer only

Use this when a customer or product needs tighter limits than the plant-wide ones in RawMaterials.

| Column | What to enter |
|---|---|
| `Name` | A raw material name from RawMaterials |
| `Min_use_%`, `Max_use_%` | wt % of the dry blend. Where both a plant-wide and a product limit exist, the **tighter** one applies. `Max_use_% = 0` bans the material, e.g. a customer who does not accept fluorspar. |

### Sheet `WhatIf` (optional): scenarios to compare with the base case

Each row is one change. Rows with the same **Scenario** name together make one scenario. Every scenario
is compared with the base case (the workbook as it stands), not with the other scenarios.

| Column | What to enter |
|---|---|
| `Scenario` | Any name, e.g. `Fluorspar dearer` |
| `Change type` | One of the types below (there is a drop-down list) |
| `Item` | The raw material or spec Name being changed |
| `New value` | The new value |

| Change type | Item | New value |
|---|---|---|
| `Price` | raw material | New price, Rs/t, on that material's own price basis |
| `Spec Min` / `Spec Max` | spec Name | New limit. Blank removes that limit. |
| `Spec Margin` | spec Name | New margin. Blank = 0. |
| `Usage Min` / `Usage Max` | raw material | New product limit, wt % of dry blend (as in UsageLimits). Blank resets it. The plant-wide limit still applies. |

Example (this is `templates\Test2_synthetic_slag_with_what_if.xlsx`):

| Scenario | Change type | Item | New value |
|---|---|---|---|
| Fluorspar dearer | Price | Fluorspar | 35000 |
| Looser P | Spec Max | P | 0.035 |
| Capped bauxite | Usage Max | Raw bauxite | 20 |
| Capped bauxite | Spec Margin | CaO | 0.5 |
| No fluorspar | Usage Max | Fluorspar | 0 |

"Capped bauxite" makes two changes at once. "No fluorspar" cannot be met (nothing else contains CaF2),
and the report explains that, just as it would for the base case.

---

## 3. Running the optimizer

Open a command window in the tool's folder (section 1) and type one of:

```
.venv\Scripts\python.exe run.py data\Slag_SS45.xlsx
```
for a product workbook that contains its own raw materials, or

```
.venv\Scripts\python.exe run.py data\Slag_SS45.xlsx --raw-materials data\Raw_materials.xlsx
```
when the raw materials are in a master workbook. Add `--output output\Slag_SS45_June.xlsx` to choose the
report's name (the default is `output\<product>_result.xlsx`), and `--quiet` to skip the summary on screen.

**What happens:**

1. The workbook is checked. **Errors stop the run.** Every error is listed at once with the sheet, row and
   column, e.g.
   `Specs (Slag_SS45.xlsx) row 3 'Al2O3', Numerator: component 'AL2O3' is not a column in RawMaterials. Did you mean 'Al2O3'?`
   Correct them all and run again. **Warnings** (e.g. an analysis that does not add up to about 100 %)
   are shown, but the run continues. Check them anyway.
2. The cheapest blend is found and a summary is printed: cost, recipe, spec check, the limits that drive
   the cost, and the raw materials not used.
3. The price range of each raw material is worked out (see *Price ranges* below).
4. Each WhatIf scenario is solved and compared with the base case.
5. The report workbook is written to `output\`, with a `.lp` file of the same name next to it.

**Close the report in Excel before running again.** Otherwise the tool cannot overwrite it and says
`If the report is open in Excel, close it and run again.`

---

## 4. Reading the report

All figures are calculated at full precision and only rounded for display: wt % to 2 decimals (spec
limits get one more decimal than you typed, e.g. P 0.030), kg to the scale resolution, and rupees to whole
numbers. Empty cells mean "not applicable". Each sheet has short notes in italics below its table.

### Summary
The product, the **solver status** (`Optimal` = a recipe was found; `Infeasible` = no recipe meets all
limits, see Diagnosis), the spec basis, the batch size, which files were read, and any warnings.

### Recipe: the weigh sheet
One row per raw material:
* **Dry blend (wt %)** is the recipe on a dry basis.
* **As-received mass (kg/batch)** is what the operator weighs. This column adds up to `Batch_size_kg`.
* **Dry mass (kg/batch)** is the same amounts without their moisture.
* The masses are rounded to the scale resolution, and **each column adds up exactly to its TOTAL**. To
  achieve that, any rounding difference is put on the largest material; the note below the table says
  which row was adjusted and by how much.
* The price columns show the price as entered, its dry-basis equivalent, and each material's share of the
  blend cost.

### Composition
The blend's analysis for every component, on the spec basis, plus the blend's LOI (dry basis) and moisture
(as received). "Sum of components" is a check: on a dry basis it plus LOI should be about 100 %.

### Spec check
One row per spec: the customer's limits, the margin, the internal **targets** the optimizer aimed for,
the **achieved** value and its **status**:
* **Binding**: this limit is what holds the recipe where it is. The blend sits exactly on the target.
  Binding limits are the ones that cost money (see Shadow prices).
* **Slack**: the limit has room to spare. *Slack to target* says how much.

### Cost
The blend cost in Rs per dry tonne, Rs per as-received tonne, and per batch.

### Shadow prices: what each binding limit costs
For each binding limit: **how much you would save, in Rs per dry tonne of blend, by relaxing it**
(lowering a Min, raising a Max). The same figure is what tightening it would cost.
* *Saving per step* uses a step sized to the limit: per 1 % for main components, per 0.01 % for a limit
  such as P 0.03 %. Example: `P_max  raising the Max  0.01  ... 7,351` means allowing 0.04 % P instead of
  0.03 % would save about Rs 7,351 per tonne for the first part of that change.
* **Valid only near the current limit.** As the limit moves further, the recipe changes and the saving per
  step changes too. Use the WhatIf sheet to get the exact effect of a bigger change. (The "Looser P" example
  scenario shows this: relaxing P by 0.005 % saves Rs 623 per tonne, not half of 7,351.)
* Rows marked *Usage limit* are plant or product usage limits that are holding the recipe, e.g. a raw
  material held at its maximum.

Use this sheet when a customer asks for a tighter specification (what will it cost?) or when you want to
know which limit is worth negotiating.

### Reduced costs: raw materials not used
For each raw material the optimizer did not use: **how much cheaper it would have to become before it is
worth using**, on a dry basis and in its own price basis, and the price at which it would enter the
recipe. A note says if it is banned, or would not be used even if free.

### Price ranges: how far each price can move before the recipe changes
For each raw material, taken one at a time with all other prices as they are: the price band (Rs/t, in
that material's own price basis) within which **today's recipe stays exactly the same**. Inside the band
only the cost changes; outside it the optimizer would choose a different recipe.
* A blank end means no limit was found in the range tested (Rs 0 up to 10 times today's price); the Note
  says which. A material whose share is fixed by the limits can show "unchanged at every price tested".
* For a raw material not in the recipe, *Unchanged from* is the price below which it would come in. It
  agrees with the Reduced costs sheet.
* The ends are found by re-solving at trial prices, to within Rs 1/t. A recipe counts as unchanged if no
  share moves by more than 0.01 wt % (0.1 kg in a 1,000 kg batch).

Use it to judge whether a supplier's new price matters (inside the band: same recipe, new cost) and how
much negotiating room there is before a substitute becomes worthwhile. This sheet is for the base case
only; for a specific price change use a WhatIf scenario. If the base case cannot be met, the sheet just
says so.

### Diagnosis (only when no recipe meets the limits)
Explains in plain language why the specification cannot be met, in two steps:
1. **Pre-check.** Each limit is looked at on its own, with only the usage limits applied. This catches
   limits that no mix could ever meet, e.g. `No allowed raw material contains CaF2, so CaF2_min (CaF2 >=
   2.00 %) cannot be met`, and names the raw materials that would help but are banned or capped.
2. **Conflicting limits.** When every limit can be met on its own but not all together, the report names
   the limits that clash and, for each of them, **the best value achievable while all the other limits are
   held**. Example: `the highest Al2O3 achievable while CaO >= 50.00 % and all other limits are held is
   39.77 %`. It then suggests the smallest relative change that makes a blend possible. That is one way out;
   moving one of the other named limits also works.

What to do about it: relax one of the named limits (change Specs, or use a WhatIf scenario to try it), or
allow or add a raw material that supplies what is missing. If a problem has several separate conflicts, the
report shows one at a time: fix it and run again.

### Scenarios and the scenario sheets (only with a WhatIf sheet)
The **Scenarios** sheet lists every scenario with its changes, its status, its cost and the change against
the base case. Each scenario then has its own sheet (`Scenario <name>`) with, one under another:
* **changes**: what was changed, from what to what;
* **cost**: base case and scenario side by side;
* **recipe**: wt % and as-received kg, base and scenario, and the change;
* **spec check**: achieved value and status in both;
* **binding limits**: the shadow prices in both. A blank means "not binding in that case";
* **diagnosis**: only if the scenario cannot be met, in the same form as above.

### The `.lp` file
A plain-text copy of the optimization problem, for checking by hand. Each line is one limit, named like
`CaO_min`, written as the amounts of each raw material times its analysis, e.g.
`CaO_min: 2 x_Bauxite + x_Calcined_alumina + 92 x_Lime >= 50`. Here `x_Lime` is the share of lime in the
dry blend, as a fraction, and the analysis values are wt %.

---

## 5. Good to know

* The optimizer puts the blend exactly on the binding targets. Real raw materials vary, so use **margins**
  on the limits that matter, rather than relying on luck.
* Recipes are in continuous amounts. Whole bags, "at least 5 % if used at all", or "at most four raw
  materials" are not handled in this version.
* The cost is for raw materials only: no freight, processing or losses.

---

## 6. For whoever maintains the tool

* **Conventions and display:** `blend_optimizer\config.py` holds them all, with comments:
  * the lab analysis convention (`LAB_ANALYSIS_CONVENTION`: `dry_incl_loi` or `loi_free`);
  * the 97–103 % analysis check;
  * the display decimals;
  * the default scale resolution;
  * the binding tolerance;
  * the elastic-model settings.
* **Templates:** regenerate with `.venv\Scripts\python.exe make_templates.py`. The example datasets are in
  `blend_optimizer\examples.py`.
* **Tests:** `.venv\Scripts\python.exe -m pytest`. They include the hand-worked Test 1 and the synthetic
  slag Test 2 with their expected results.
* **Price ranging** is in every report; from Python it is
  `blend_optimizer.what_if.price_ranging(load_inputs(...))`, with options for the accuracy and search range.
* **Code layout** (in `blend_optimizer\`):

| Module | What it does |
|---|---|
| `inputs.py` | Reads and checks the workbooks |
| `linear_rows.py` | Turns each spec limit into one linear constraint (basis conversion, ratios) |
| `model.py`, `solve_and_report.py` | The only modules that use PuLP/HiGHS: build and solve the problem, build the report |
| `pipeline.py` | Runs a case and diagnoses it when it cannot be met |
| `diagnose.py` | The pre-checks and the plain-language diagnosis |
| `what_if.py` | What-if scenarios and price ranging |

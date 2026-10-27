# Least-cost blend optimizer: User Guide

## About this guide

This guide is for metallurgists and process engineers who prepare blended products (synthetic slags, desulphurising and deoxidising compounds) and want the cheapest recipe that meets a customer's specification. You do not need to know programming. You will type a few commands, and the guide gives every one of them, exactly.

The short reference is `README.md` in the tool's folder. This guide is the complete manual: it explains each step, each sheet, each column and each message.

**Conventions used in this guide**

| You see | It means |
|---|---|
| ₹ (Rs) | Indian rupees. The tool itself writes "Rs" in its reports and on screen, so headings and messages are quoted with "Rs". |
| wt % | Percent by mass. All analyses and limits are in wt %. |
| t, dry t | Metric tonne; a dry tonne is a tonne with its free moisture removed. |
| `text like this` | Something you type, or something the tool prints, exactly as shown. |
| Blocks with a grey background | Commands to type in Windows PowerShell, or text the tool prints. |

Every command in this guide has been run, and every screen or report excerpt is copied from a real run on the example workbooks that come with the tool (`Test1_hand_worked.xlsx`, `Test2_synthetic_slag_dry.xlsx` and `Test2_synthetic_slag_with_what_if.xlsx`) or on small copies of them. The examples use invented data, not real company data.

## Contents

- [1. What the tool does and when to use it](#1-what-the-tool-does-and-when-to-use-it)
  - [1.1 The problem it solves](#11-the-problem-it-solves)
  - [1.2 What it needs and what it gives back](#12-what-it-needs-and-what-it-gives-back)
  - [1.3 The whole idea on one page: Test 1](#13-the-whole-idea-on-one-page-test-1)
- [2. How it works, without mathematics](#2-how-it-works-without-mathematics)
  - [2.1 The mixing rule](#21-the-mixing-rule)
  - [2.2 Binding and slack limits](#22-binding-and-slack-limits)
  - [2.3 Why the cheapest recipe uses only a few raw materials](#23-why-the-cheapest-recipe-uses-only-a-few-raw-materials)
  - [2.4 Two raw materials: why a specification can be impossible](#24-two-raw-materials-why-a-specification-can-be-impossible)
- [3. One-time setup](#3-one-time-setup)
  - [3.1 What you need](#31-what-you-need)
  - [3.2 Install Python](#32-install-python)
  - [3.3 Open PowerShell in the tool's folder](#33-open-powershell-in-the-tools-folder)
  - [3.4 Create the virtual environment](#34-create-the-virtual-environment)
  - [3.5 Activate the virtual environment](#35-activate-the-virtual-environment)
  - [3.6 Install the packages](#36-install-the-packages)
  - [3.7 Check the installation](#37-check-the-installation)
  - [3.8 Choose the interpreter in VS Code (optional)](#38-choose-the-interpreter-in-vs-code-optional)
  - [3.9 The folder layout](#39-the-folder-layout)
  - [3.10 If you move or rename the folder](#310-if-you-move-or-rename-the-folder)
- [4. Choosing a workbook layout](#4-choosing-a-workbook-layout)
- [5. Filling in each sheet](#5-filling-in-each-sheet)
  - [5.1 Rules for every sheet](#51-rules-for-every-sheet)
  - [5.2 RawMaterials](#52-rawmaterials)
  - [5.3 Specs](#53-specs)
  - [5.4 Settings](#54-settings)
  - [5.5 UsageLimits (optional)](#55-usagelimits-optional)
  - [5.6 WhatIf (optional)](#56-whatif-optional)
  - [5.7 Checklist before your first real run](#57-checklist-before-your-first-real-run)
- [6. Running the tool](#6-running-the-tool)
  - [6.1 The command](#61-the-command)
  - [6.2 Options](#62-options)
  - [6.3 Where the report goes](#63-where-the-report-goes)
  - [6.4 Time-savers](#64-time-savers)
  - [6.5 Reading the screen](#65-reading-the-screen)
- [7. Reading the report](#7-reading-the-report)
  - [7.1 Summary](#71-summary)
  - [7.2 Recipe](#72-recipe)
  - [7.3 Composition](#73-composition)
  - [7.4 Spec check](#74-spec-check)
  - [7.5 Cost](#75-cost)
  - [7.6 Shadow prices](#76-shadow-prices)
  - [7.7 Reduced costs](#77-reduced-costs)
  - [7.8 Diagnosis](#78-diagnosis)
  - [7.9 Price ranges](#79-price-ranges)
  - [7.10 Scenarios](#710-scenarios)
  - [7.11 Scenario sheets](#711-scenario-sheets)
  - [7.12 The .lp file](#712-the-lp-file)
- [8. When there is no recipe (infeasible)](#8-when-there-is-no-recipe-infeasible)
  - [8.1 How the tool explains it](#81-how-the-tool-explains-it)
  - [8.2 Example 1: a banned raw material](#82-example-1-a-banned-raw-material)
  - [8.3 Example 2: two limits that conflict](#83-example-2-two-limits-that-conflict)
  - [8.4 What to do next](#84-what-to-do-next)
- [9. Everyday tasks](#9-everyday-tasks)
- [10. Before using a recipe in production](#10-before-using-a-recipe-in-production)
- [11. Troubleshooting](#11-troubleshooting)
- [12. Limitations](#12-limitations)
- [13. Glossary](#13-glossary)
- [Appendix A. Message reference](#appendix-a-message-reference)
- [Appendix B. For the maintainer](#appendix-b-for-the-maintainer)

---

## 1. What the tool does and when to use it

### 1.1 The problem it solves

A customer gives you a specification: for example CaO between 50 and 56 wt %, Al2O3 at least 30 wt %, SiO2 at most 6 wt %. You have several raw materials, each with its own analysis and price. Usually **many** recipes meet the specification. Some cost far more than others.

The tool does three things:

1. It finds the **cheapest** recipe that meets every limit.
2. It **proves** that no cheaper recipe exists. This is not a search that might miss something: the method (linear programming, explained in [chapter 2](#2-how-it-works-without-mathematics)) gives the exact optimum.
3. It **explains** the answer: which limits drive the cost, what it would save to relax each one, how much cheaper an unused raw material would have to be before it is worth using, and, when no recipe exists, why.

Use it when you quote a new product, when a raw material price changes, when a customer asks for a tighter or looser limit, when a supplier offers a new material, and to check whether today's production recipe is the cheapest one.

### 1.2 What it needs and what it gives back

| You provide (in an Excel workbook) | The tool gives back (in an Excel report and on screen) |
|---|---|
| Each raw material's price (₹ per tonne), moisture (wt %), LOI (wt %) and chemical analysis (wt %, dry basis) | The cheapest recipe, in wt % of the dry blend and in kg per batch |
| The customer's limits, e.g. CaO 50–56 wt % | The blend analysis and a check of every limit |
| Optional: limits on how much of a raw material may be used | The cost in ₹ per dry tonne, per as-received tonne and per batch |
| Optional: what-if scenarios to compare | What each binding limit costs, which raw materials are not used and why, the price range of each raw material, and a side-by-side comparison of every scenario |
| | If no recipe exists: a plain-language diagnosis |

### 1.3 The whole idea on one page: Test 1

The tool comes with a small worked example, `templates\Test1_hand_worked.xlsx`. It has three raw materials. Prices are per dry tonne and there is no moisture or LOI, so it is easy to check by hand.

**Raw materials** (sheet RawMaterials, analyses in wt %):

| Name | Cost_Rs_per_t | Price_basis | CaO | Al2O3 | SiO2 | Others |
|---|---|---|---|---|---|---|
| Lime | 9000 | dry | 92 | 1 | 2 | 5 |
| Calcined alumina | 30000 | dry | 1 | 85 | 6 | 8 |
| Bauxite | 12000 | dry | 2 | 60 | 8 | 30 |

**Customer limits** (sheet Specs): CaO at least 50 wt %, Al2O3 at least 30 wt %, SiO2 at most 6 wt %.

**What the tool prints** (from a real run of `python run.py templates\Test1_hand_worked.xlsx`):

```
Cost: Rs 11,482 per dry t of blend | Rs 11,482 per as-received t | Rs 11,482 per batch

Recipe (batch 1,000.0 kg as received, 1,000.0 kg dry):
      Raw material Dry blend (wt %) Dry mass (kg/batch) As-received mass (kg/batch)
              Lime            53.40               534.0                       534.0
  Calcined alumina             6.02                60.2                        60.2
           Bauxite            40.58               405.8                       405.8
             TOTAL           100.00             1,000.0                     1,000.0

Spec check:
   Spec Limits Basis Target Min Target Max Achieved        Status
    CaO    CaO   dry      50.00               50.00 Binding (Min)
  Al2O3  Al2O3   dry      30.00               30.00 Binding (Min)
   SiO2   SiO2   dry                  6.00     4.68         Slack

Shadow prices of binding limits (Rs per dry t saved by relaxing; valid only near the current limit):
  Constraint         Relax by Step       Limit unit Saving per step (Rs/dry t) Saving per 1 unit (Rs/dry t)
     CaO_min lowering the Min    1 wt % (dry basis)                        450                          450
   Al2O3_min lowering the Min    1 wt % (dry basis)                        738                          738
```

How to read it:

- The cheapest recipe is 53.40 wt % Lime, 6.02 wt % Calcined alumina and 40.58 wt % Bauxite. It costs ₹ 11,482 per dry tonne. No other mix of these three materials that meets the limits is cheaper.
- CaO and Al2O3 sit exactly on their minimums (**Binding**): these two limits hold the recipe where it is. SiO2 is 4.68 wt %, below its 6 wt % maximum (**Slack**): it does not affect the cost.
- If the customer accepted 49 wt % CaO instead of 50 wt %, the blend would cost about ₹ 450 per dry tonne less. One wt % less Al2O3 would save about ₹ 738 per dry tonne.

The rest of this guide explains each of these ideas and every other part of the report.

---

## 2. How it works, without mathematics

### 2.1 The mixing rule

When you mix raw materials, the blend's analysis is the **weighted average** of the materials' analyses, weighted by how much of each you use. The same is true for cost.

Test 1 example, with the shares from the recipe above:

- Blend CaO = 0.5340 × 92 + 0.0602 × 1 + 0.4058 × 2 = 50.00 wt %
- Blend cost = 0.5340 × ₹ 9,000 + 0.0602 × ₹ 30,000 + 0.4058 × ₹ 12,000 = ₹ 4,806 + ₹ 1,807 + ₹ 4,869 = ₹ 11,482 per dry tonne

(These three cost figures appear in the report's Recipe sheet, in the column "Cost share (Rs/dry t of blend)".)

Because every limit and the cost are weighted averages, doubling a share doubles its effect, and the effects simply add up. Mathematicians call such a problem **linear**. For linear problems there is an exact method, called **linear programming**, that finds the best answer and proves that nothing better exists. The program that carries out the method is called a **solver**; this tool uses a solver called HiGHS.

Ratios such as basicity (CaO / SiO2) are not weighted averages, but a ratio limit such as "CaO / SiO2 at least 8" can be rewritten as "CaO − 8 × SiO2 at least 0", which is. The tool does this for you.

### 2.2 Binding and slack limits

After the tool has found the cheapest recipe, every limit is in one of two states:

- **Binding**: the blend sits exactly on the limit. The limit is pushing against the recipe; if you relaxed it, the recipe would change and get cheaper. In Test 1, CaO (exactly 50.00 wt %, minimum 50) and Al2O3 (exactly 30.00 wt %, minimum 30) are binding.
- **Slack**: the blend is inside the limit with room to spare. Relaxing it would change nothing. In Test 1, SiO2 is 4.68 wt % against a maximum of 6 wt %: slack by 1.32 wt %.

Only binding limits cost money. The report tells you, for each binding limit, how much (see [7.6 Shadow prices](#76-shadow-prices)).

### 2.3 Why the cheapest recipe uses only a few raw materials

The cheapest recipe usually uses **no more raw materials than the number of binding limits plus one**. (A usage limit that holds a raw material at its maximum or minimum counts as a binding limit here.) Each binding limit "uses up" one degree of freedom; the "+ 1" is because the shares must add up to 100 %.

| Example | Binding limits | Raw materials used |
|---|---|---|
| Test 1 | 2 (CaO, Al2O3) | 3 of 3 |
| Test 2 (synthetic slag, dry basis) | 4 (CaO, Al2O3, CaF2, P) | 5 of 6: Lime fines is at 0 % |
| Test 1 with the Al2O3 row deleted from Specs | 0 | 1 of 3: 100 % Lime |

So a raw material at **0 % is normal, not an error**. It means the material brings nothing that the others cannot bring more cheaply. The report's Reduced costs sheet tells you how much cheaper it would have to be before it is worth using: in Test 2, Lime fines would have to fall from ₹ 6,500 to below ₹ 1,872 per tonne (as received).

The last row of the table also shows a danger: if a spec row is missing, the tool does not complain. It simply optimises without that limit (here: 100 % Lime). This is why you should read the Spec check sheet after every run ([7.4](#74-spec-check)).

### 2.4 Two raw materials: why a specification can be impossible

With only two raw materials there is only one decision: the share of the first one. This makes it easy to see why a specification can be impossible. The example uses two invented materials, P and Q (they also appear in the tool's own tests):

| Name | Al2O3 (wt %) | Al (wt %) | Price (₹ per dry t) |
|---|---|---|---|
| P | 60 | 10 | 10,000 |
| Q | 72 | 19 | 15,000 |

Each limit allows only a range of P's share:

| Limit | Blend value when the share of P is p | Allowed share of P |
|---|---|---|
| Al2O3 at least 65 wt % | 72 − 12 × p | at most 58.33 % |
| Al at most 12 wt % | 19 − 9 × p | at least 77.78 % |

```
share of P:   0 %           58.33 %       77.78 %          100 %
Al2O3 >= 65:  |================|
Al <= 12:                                    |================|
```

The two ranges do not overlap, so **no recipe meets both limits**. The problem is called **infeasible**. This is what the tool reports for this case (real run):

```
Status: Infeasible | spec basis: dry | lab convention: dry_incl_loi
No recipe: the solver status is 'Infeasible'.

No blend meets every limit (the problem is infeasible).
Pre-check: every limit can be met on its own, so some limits conflict with each other.
Elastic model: these limits cannot all be met together: Al2O3_min, Al_max.
  - Al2O3_min (Al2O3 >= 65.00 %): the highest Al2O3 achievable while Al <= 12.00 % and all other limits are held is 62.67 %.
  - Al_max (Al <= 12.00 %): the lowest Al achievable while Al2O3 >= 65.00 % and all other limits are held is 13.75 %.
```

The numbers match the picture: at 77.78 % P (the least P that keeps Al at 12 wt %), Al2O3 is only 62.67 wt %; at 58.33 % P (the most P that still gives 65 wt % Al2O3), Al is 13.75 wt %.

If the customer accepted Al up to 15 wt %, the second range would become "at least 44.44 %". The ranges then overlap between 44.44 % and 58.33 %, and the tool picks the cheapest share inside the overlap: 58.33 % P (P is the cheaper material), at ₹ 12,083 per dry tonne (real run).

With more raw materials there are more decisions and you cannot draw the picture any more, but the idea is the same: each limit allows a region of recipes, and the specification is infeasible when the regions have no recipe in common. [Chapter 8](#8-when-there-is-no-recipe-infeasible) shows how the tool explains such cases.

---

## 3. One-time setup

You do this once per computer. If someone has already set up the tool on your computer, skip to [3.7 Check the installation](#37-check-the-installation).

### 3.1 What you need

- A Windows 10 or 11 computer. (The tool also runs on macOS and Linux; the differences are noted where they matter.)
- **Python 3.12 or newer.** The packages the tool uses need at least 3.12. The tool was built and tested with Python 3.13.15, so 3.13 is the safest choice.
- An internet connection for the one-time installation of the packages (about a minute). After that the tool works completely offline.
- The tool's folder, for example `D:\VS Python\Cost Optimizer`.

### 3.2 Install Python

This step is done in a web browser and the Python installer, not with a command.

1. Go to python.org, open *Downloads → Windows*, and download the **Windows installer (64-bit)** for Python 3.13.
2. Run the installer and keep the default options. The installer also installs the Python launcher `py`.
3. Check the installation in PowerShell (see [3.3](#33-open-powershell-in-the-tools-folder) for how to open it):

```
py -3.13 --version
```

Expected output (the last digits can differ):

```
Python 3.13.15
```

### 3.3 Open PowerShell in the tool's folder

All commands in this guide are typed in **Windows PowerShell**, in the tool's folder.

1. Open the tool's folder in File Explorer.
2. Click in the address bar at the top, type `powershell` and press Enter.

A PowerShell window opens in that folder. Alternatively, open PowerShell from the Start menu and change to the folder with `cd`. Put the path in quotes because it contains a space:

```
cd "D:\VS Python\Cost Optimizer"
```

In VS Code, *Terminal → New Terminal* opens PowerShell in the open folder.

### 3.4 Create the virtual environment

A **virtual environment** is a private copy of Python, inside the tool's folder, with exactly the package versions the tool needs. It keeps the tool independent of anything else installed on your computer. It lives in a folder called `.venv`.

In PowerShell, in the tool's folder, type:

```
py -3.13 -m venv .venv
```

This takes a few seconds and prints nothing. A new folder `.venv` appears.

### 3.5 Activate the virtual environment

**Activating** tells this PowerShell window to use the Python inside `.venv`. You must do this **every time you open a new PowerShell window**.

```
.\.venv\Scripts\Activate.ps1
```

The prompt now starts with `(.venv)`:

```
(.venv) PS D:\VS Python\Cost Optimizer>
```

**If PowerShell refuses.** On many company computers, PowerShell blocks scripts. You then see a red message like this (real output; the path is shortened):

```
...\.venv\Scripts\Activate.ps1 cannot be loaded because running scripts
is disabled on this system. For more information, see about_Execution_Policies at
https:/go.microsoft.com/fwlink/?LinkID=135170.
```

Use one of these commands, then activate again:

```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```

```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

The first one allows scripts **in this window only**; you repeat it in each new window. The second allows them **permanently for your Windows account**. If PowerShell asks you to confirm, type `Y` and press Enter. If neither is allowed on your computer, ask IT, or use Command Prompt instead (below).

**Command Prompt alternative.** In Command Prompt (`cmd`), activate with:

```
.venv\Scripts\activate.bat
```

**Without activating.** You can also skip activation and type the full path to the Python in `.venv` instead of `python`, for example `.venv\Scripts\python.exe run.py templates\Test1_hand_worked.xlsx`. This guide uses the shorter `python` form and assumes the environment is active.

**macOS and Linux.** Create the environment with `python3 -m venv .venv` and activate it with `source .venv/bin/activate`. Use `/` instead of `\` in all paths.

### 3.6 Install the packages

With the environment active, type:

```
python -m pip install -r requirements.txt
```

This downloads and installs the exact package versions listed in `requirements.txt`, including PuLP (the modelling package) and HiGHS (the solver, package name `highspy`). Older versions of PuLP came with a solver called CBC; PuLP 4 does not, and the tool does not need it.

On the test computer this took 36 seconds and ended with:

```
Successfully installed PuLP-4.2.0 Pygments-2.21.0 annotated-types-0.8.0 colorama-0.4.6 cpsat-logutils-1.0.0 et_xmlfile-2.0.0 highspy-1.15.1 iniconfig-2.3.1 numpy-2.5.4 openpyxl-3.1.5 orloge-1.0.0 packaging-26.3 pandas-3.1.0 pluggy-1.7.0 pydantic-2.14.1 pydantic_core-2.50.1 pytest-9.1.1 python-dateutil-3.0.0 six-1.17.0 typing-inspection-0.4.4 typing_extensions-4.16.0 tzdata-2026.5
```

PowerShell may also show a red block containing `NativeCommandError` and a line like `[notice] A new release of pip is available`. This is only a notice, not an error. You can ignore it.

### 3.7 Check the installation

1. Check that the solver is found:

```
python -c "import pulp; print('PuLP', pulp.__version__, '| solvers:', pulp.listSolvers(onlyAvailable=True))"
```

Expected output:

```
PuLP 4.2.0 | solvers: ['HiGHS']
```

2. Run the first example:

```
python run.py templates\Test1_hand_worked.xlsx
```

The output must start like this:

```
=== Least-cost blend: Test 1 hand-worked example ===
Status: Optimal | spec basis: dry | lab convention: dry_incl_loi
Cost: Rs 11,482 per dry t of blend | Rs 11,482 per as-received t | Rs 11,482 per batch
```

and end with:

```
Report written to: D:\VS Python\Cost Optimizer\output\Test1_hand_worked_result.xlsx
LP file:           D:\VS Python\Cost Optimizer\output\Test1_hand_worked_result.lp
```

If you see this, the tool is ready.

### 3.8 Choose the interpreter in VS Code (optional)

This is a manual step in the VS Code window. It is only needed if you run the tool from VS Code.

1. Open the tool's folder in VS Code (*File → Open Folder*).
2. Press Ctrl+Shift+P, type `Python: Select Interpreter` and press Enter.
3. Choose the entry that points to `.venv` (it shows `.\.venv\Scripts\python.exe`).

New terminals in VS Code then activate the environment automatically.

### 3.9 The folder layout

| Folder or file | What it holds | In git? |
|---|---|---|
| `templates\` | Blank templates to copy, and the example workbooks used in this guide | Yes |
| `data\` | **Your** product and raw-material workbooks. Create it once (command below). | **No** |
| `output\` | The reports and `.lp` files the tool writes. Created automatically. | **No** |
| `run.py` | The program you run | Yes |
| `README.md`, `docs\` | The short reference and this guide | Yes |
| `blend_optimizer\`, `tests\`, `make_templates.py` | The program's parts and its tests. Do not edit. | Yes |
| `requirements.txt` | The exact package versions | Yes |
| `.venv\` | The virtual environment. Do not edit. | No |

**Keep your real workbooks in `data\`.** Recipes and prices are confidential. The tool's folder is kept in git (a version-control system), and its settings exclude `data\` and `output\`, so files there are never uploaded. Workbooks saved anywhere else in the folder, for example in `templates\`, could be uploaded by mistake.

Create the `data` folder once:

```
New-Item -ItemType Directory data
```

### 3.10 If you move or rename the folder

The virtual environment remembers the folder it was created in. If you move or rename the tool's folder, delete `.venv` and create it again:

```
Remove-Item -Recurse -Force .venv
```

Then repeat [3.4](#34-create-the-virtual-environment), [3.5](#35-activate-the-virtual-environment) and [3.6](#36-install-the-packages).

---

## 4. Choosing a workbook layout

The tool reads Excel workbooks. You can organise them in two ways:

| | Layout A: one workbook per product | Layout B: a master raw-material workbook plus one workbook per product |
|---|---|---|
| Raw materials | In each product workbook | Once, in the master workbook |
| Specs, Settings, UsageLimits, WhatIf | In the product workbook | In each product workbook |
| Template to copy | `templates\Template_single_product.xlsx` | `templates\Template_master_raw_materials.xlsx` (once) and `templates\Template_product.xlsx` (per product) |
| A price changes | Update every product workbook that uses it | Update the master once |
| Best for | Starting out; one or two products; a one-off quote | Many products from the same raw materials |

**Recommendation: start with layout A.** Everything for one product is in one file, which is easier to check. Move to layout B when you run several products from the same raw materials.

**To start a layout A workbook:**

1. Copy the template into `data\` under the product's name:

```
Copy-Item templates\Template_single_product.xlsx data\Slag_SS45.xlsx
```

2. Open `data\Slag_SS45.xlsx` in Excel and fill in the sheets ([chapter 5](#5-filling-in-each-sheet)).

**To start layout B:** copy `templates\Template_master_raw_materials.xlsx` to, for example, `data\Raw_materials.xlsx` (once), and `templates\Template_product.xlsx` to a file per product. You run a product with the extra option `--raw-materials` ([6.2](#62-options)).

If a product workbook in layout B still contains its own RawMaterials sheet, the tool uses the master and warns you. Real run, with a layout A workbook given a master file:

```
Input warnings (1):
  - Test2_synthetic_slag_dry.xlsx also has a RawMaterials sheet. It is ignored because a master raw-material file (Example_master_raw_materials.xlsx) was given.
```

The example workbooks show both layouts: `templates\Test2_synthetic_slag_dry.xlsx` is layout A; `templates\Example_master_raw_materials.xlsx` with `templates\Test2_product_only.xlsx` is the same product in layout B.

---

## 5. Filling in each sheet

A product workbook can have these sheets. Every template also has a **README** sheet with a short version of these instructions.

| Sheet | Required? | One row per |
|---|---|---|
| RawMaterials | Yes (in the product workbook, or in the master workbook for layout B) | raw material |
| Specs | Yes | customer limit |
| Settings | Yes | setting |
| UsageLimits | No | raw material with a product- or customer-specific limit |
| WhatIf | No | change in a what-if scenario |

### 5.1 Rules for every sheet

1. **Keep the headers in row 1** and do not rename the fixed column headers (such as `Cost_Rs_per_t`). The tool is forgiving about capital letters, spaces, underscores and the % sign in these fixed headers: a header typed `moisture %` is read as `Moisture_%`. A header that is *almost* right, such as `Moistre_%`, is an error, because it would otherwise be read as a chemical component.
2. **Type plain numbers.** No units, no `<0.01`, no `nd`. Text that is a number, such as `12.5`, is accepted. In an analysis column, a **blank** cell means "not analysed" and a **0** means "analysed, none present". Both count as 0 in the calculation; see [5.2](#52-rawmaterials) for the warning a blank can give.
3. **Names must match exactly** between sheets: a component name in Specs must be written exactly as the column header in RawMaterials (`Al2O3`, not `AL2O3`). If you make a typo, the error message suggests the name you probably meant.
4. Blank rows are ignored.
5. Every problem is reported at once when you run the tool, so you can correct them all before running again.

### 5.2 RawMaterials

One row per raw material. The first seven columns are fixed; **every other column is a chemical component**.

| Column | What to enter | Unit | If left blank |
|---|---|---|---|
| `Name` | A unique name. Two names that differ only in capital letters count as the same name, which is an error. | – | Error |
| `Cost_Rs_per_t` | The **landed cost**: price plus freight and handling to your plant, per tonne | ₹ per tonne, on the basis in `Price_basis` | Error |
| `Price_basis` | `as_received` (the price is per tonne as delivered, with its moisture) or `dry` (per dry tonne) | – | Treated as `dry` if `Moisture_%` is 0; an error otherwise |
| `Moisture_%` | Free moisture of the material as delivered | wt %, as received | 0 (with a warning if the price is `as_received`) |
| `LOI_%` | Loss on ignition | wt %, dry basis | 0 |
| `Min_use_%` | Plant-wide minimum share of this material in **any** blend | wt % of the dry blend | 0 |
| `Max_use_%` | Plant-wide maximum share, for real plant reasons (e.g. a fines grade that segregates) | wt % of the dry blend | 100 |
| any other column | The analysis of one component, e.g. `CaO`. The header is the component name. Add, remove or reorder these columns freely. Type 0 for a component that was analysed and is not present. | wt %, dry basis | Counted as 0, but taken as "not analysed": a warning if a spec uses the component (see below) |

**Worked example** (Test 2, first two rows):

| Name | Cost_Rs_per_t | Price_basis | Moisture_% | LOI_% | Min_use_% | Max_use_% | CaO | SiO2 | Al2O3 | MgO | CaF2 | Fe2O3 | TiO2 | S | P |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Lime | 9000 | as_received | 0.5 | 3.5 | | | 92 | 1.5 | 0.8 | 1.2 | 0 | 0.3 | 0 | 0.03 | 0.01 |
| Lime fines | 6500 | as_received | 1 | 6 | | 20 | 88 | 2.5 | 1.2 | 1.5 | 0 | 0.6 | 0 | 0.05 | 0.02 |

- Lime costs ₹ 9,000 per tonne **as delivered**, with 0.5 wt % moisture. The tool converts every price to ₹ per **dry** tonne: 9,000 ÷ (1 − 0.5/100) = ₹ 9,045 per dry tonne. (You will see 9,045 in the report's Recipe sheet, column "Dry cost (Rs/dry t)".)
- Lime fines may never exceed 20 wt % of the dry blend in any product (a plant-wide limit). Its Min_use_% is blank, so its minimum is 0.

**Good practice**

- Enter the **landed cost**, not the supplier's list price.
- Enter the **average of recent lots**, not one sample. One unusual sample gives one unusual recipe.
- List **every material that could be used**, not only today's. The tool will tell you if a material is not worth using, and at what price it would be.
- Use plant-wide `Min_use_%` and `Max_use_%` only for real plant reasons (handling, segregation, a contract minimum). Customer- or product-specific limits belong in the UsageLimits sheet ([5.5](#55-usagelimits-optional)).

**The analysis basis and the 97–103 % check**

The tool assumes that analyses are on a **dry basis with LOI reported separately**, so that for each material the components plus LOI add up to about 100 wt %. For Lime above: 92 + 1.5 + 0.8 + 1.2 + 0 + 0.3 + 0 + 0.03 + 0.01 = 95.84 wt %, plus LOI 3.5 = 99.34 wt %.

If a material adds up to less than 97 or more than 103 wt %, the tool prints a **warning**, not an error: the run continues. Example from the example workbook `templates\Broken_example.xlsx`, where Lime's CaO was typed as 80 instead of 92:

```
  - RawMaterials 'Lime': components + LOI = 87.34 %, outside the expected 97-103 % for the 'dry_incl_loi' lab convention. Neither convention fits (components alone = 83.84 %, components + LOI = 87.34 %): check for a missing component column, a typing error or a missing LOI.
```

Some laboratories report analyses on an **LOI-free** (ignited) basis, so the components alone add up to about 100 %. The tool can read such analyses too, but this is a setting for the whole tool, called `LAB_ANALYSIS_CONVENTION`, which the maintainer changes ([Appendix B](#appendix-b-for-the-maintainer)). The two conventions are `dry_incl_loi` (the default) and `loi_free`. The warning tells you when your numbers look like the other convention.

**Blank or 0: why it matters.** The tool only knows the values you type. A blank cell is counted as 0, so a value that was never measured can make the product fail a limit that the report says is met. That is why the tool treats a blank as "not analysed" and a 0 as "analysed, none present", and **warns** about a blank in a component that a spec uses, for every raw material the blend may use. Example (a copy of Test 2 where the P cell of Calcined bauxite was left blank instead of 0.05; real run):

```
Input warnings (1):
  - RawMaterials 'Calcined bauxite': P is blank (not analysed) and is counted as 0, but a spec uses it. Type 0 if it is genuinely absent.
```

- The material still adds up to 99.55 wt %, so the 97–103 % check does not catch it. This warning does.
- The run continues with P taken as 0. The tool now thinks Calcined bauxite carries no phosphorus, and chooses a different recipe that costs ₹ 12,381 instead of ₹ 13,004 per dry tonne.
- The report shows P = 0.024 wt % (limit 0.030 wt %). With the real analysis, the same recipe has **0.032 wt % P**: it would fail the customer's limit.

So fill in every value that was analysed, and **type 0** where a component is genuinely absent. There is no warning for a blank in a component that no spec uses (it cannot change the recipe), or for a material that may not be used at all (Max_use_% 0 in RawMaterials or UsageLimits).

### 5.3 Specs

One row per customer limit.

| Column | What to enter | If left blank |
|---|---|---|
| `Name` | A unique label for the limit. In the report the limits are called `<Name>_min` and `<Name>_max`. | Error |
| `Numerator` | What is limited: a component, or several joined with `+`, each optionally with a conversion factor (see below) | Error |
| `Denominator` | Leave blank for a composition limit. Fill it in for a **ratio** (same syntax as Numerator). | Composition limit |
| `Min` | The customer's minimum | No minimum |
| `Max` | The customer's maximum | No maximum (but at least one of Min and Max is needed) |
| `Margin` | Your safety margin, in the same unit as the limit | 0 |

**The kinds of limit, with examples**

| Kind | Name | Numerator | Denominator | Min | Max | Margin | Meaning |
|---|---|---|---|---|---|---|---|
| Minimum | Al2O3 | Al2O3 | | 30 | | | at least 30 wt % Al2O3 (Test 1) |
| Maximum | P | P | | | 0.03 | | at most 0.03 wt % P (Test 2) |
| Range | CaO | CaO | | 50 | 56 | | between 50 and 56 wt % CaO (Test 2) |
| Sum | CaO+MgO | CaO+MgO | | 55 | | | CaO plus MgO at least 55 wt % |
| Ratio | Basicity | CaO | SiO2 | 8 | | | CaO / SiO2 at least 8 (Test 2) |
| Conversion factor | Total Al2O3 | Al2O3+1.8895*Al | | 75 | | | Al2O3 with metallic Al counted as Al2O3, at least 75 wt % |
| With a margin | CaO | CaO | | 50 | 56 | 0.5 | the tool aims for 50.5–55.5 wt % |

**Conversion factors: counting one component as another**

Some customers count a component in another form. Metallic Al, for example, oxidises to Al2O3 in the slag, so it is counted in Al2O3 by stoichiometry. Write the factor in front of the component, joined with `*`:

```
Al2O3+1.8895*Al
```

This means "Al2O3 wt % + 1.8895 × Al wt %" for every raw material. 1.8895 is the mass of Al2O3 formed from one unit mass of Al: the molar mass of Al2O3 (101.96) divided by its two atoms of Al (2 × 26.98).

Rules:

1. Join terms with `+`. A term is a component (factor 1) or `number*Component`. Spaces around `+` and `*` are allowed.
2. The factor is a plain positive number with a decimal point (`1.8895`, not `1,8895`), written **before** the component.
3. `-` (subtraction) is not allowed. A hyphen inside a component name, such as `Fe-total`, is fine.
4. Each component may appear only once in an expression.
5. The tool knows no chemistry: the factor is simply the number you type.

Common factors, each worked out as the molar mass of what it is counted as, per atom of the shared element, divided by the molar mass of what was analysed:

| Analysed as | Counted as | Factor | Derivation |
|---|---|---|---|
| Al | Al2O3 | 1.8895 | 101.96 / (2 × 26.98) |
| Mg | MgO | 1.6583 | 40.304 / 24.305 |
| Ca | CaO | 1.3992 | 56.077 / 40.078 |
| Si | SiO2 | 2.1393 | 60.084 / 28.086 |
| CaC2 | CaO | 0.8748 | 56.077 / 64.100 |
| FeO | Fe2O3 | 1.1113 | 159.687 / (2 × 71.844) |
| CaF2 | F | 0.4867 | (2 × 18.998) / 78.075 |

For a metal counted as its oxide the rule reduces to: oxide molar mass ÷ (number of metal atoms in the oxide × atomic mass of the metal).

Example (real run on two invented materials P with 60 wt % Al2O3 and 10 wt % Al, and Q with 72 wt % Al2O3 and 19 wt % Al; spec "Total Al2O3", Numerator `Al2O3+1.8895*Al`, Min 80):

```
Spec check:
         Spec          Limits Basis Target Min Target Max Achieved        Status
  Total Al2O3 Al2O3+1.8895*Al   dry      80.00               80.00 Binding (Min)
```

The report always shows the expression as you typed it (spaces removed), so the reader sees that a conversion was applied. A conversion factor works everywhere a plain component does: on every spec basis, in ratios (also in the Denominator), with margins, in WhatIf changes and in the diagnosis.

If you used a helper column in RawMaterials for this before (e.g. a column `Al_as_Al2O3` = 1.8895 × Al), it is no longer needed: put the factor in Specs and delete the column. The tool's tests check that both ways give exactly the same recipe and cost.

**Units**

- Composition limits are in **wt %** on the spec basis ([5.4](#54-settings)). Enter a trace limit of 300 ppm as **0.03** (1 wt % = 10,000 ppm), not 300.
- Ratio limits have no unit: basicity 8 means CaO / SiO2 = 8.

**Margins only ever tighten a limit.** The tool aims for Min + margin and Max − margin, the **internal target**. The customer's limit is unchanged and shown next to the target in the report. A negative margin is an error. A margin larger than half the range is also an error, for example (from `Broken_example.xlsx`, CaO 50–56 with margin 4):

```
  - Specs (Broken_example.xlsx) row 2 'CaO': after the margin the internal range is empty: Min + margin = 54 is greater than Max - margin = 52. Reduce the margin.
```

### 5.4 Settings

Two columns, `Key` and `Value`, one row per setting.

| Key | Value | If missing or blank |
|---|---|---|
| `Product_name` | The product's name, shown on the report | The file name is used, with a warning |
| `Spec_basis` | The basis of the customer's composition limits: `dry`, `as_received` or `ignited` | `dry` |
| `Batch_size_kg` | The total mass weighed into the mixer per batch, **as received**, in kg | 1000 |
| `Scale_resolution_kg` | Optional. The resolution of your weighing scale, in kg (e.g. 0.5). The weigh sheet is rounded to it. | 0.1 |

Test 2 example:

| Key | Value |
|---|---|
| Product_name | Synthetic slag (Test 2, dry basis) |
| Spec_basis | dry |
| Batch_size_kg | 1000 |

**The spec basis**

The same blend has different analyses depending on how the sample is treated:

- `dry`: the analysis of the dried blend (free moisture removed, LOI still in it). This is how the raw material analyses are entered.
- `as_received`: the analysis of the blend as delivered, including its free moisture. Each component is a little lower than on a dry basis.
- `ignited`: the analysis after loss on ignition, as the material is when it has melted into the slag. Each component is higher than on a dry basis, because the volatile part (LOI) is gone.

Confirm the basis with the customer: it is usually stated in their specification or on their analysis certificate ("on dry basis", "as received", "on ignited basis" or "loss-free"). If it is not stated, ask. Ratios such as basicity are the same on every basis.

**Why it matters.** The basis changes the recipe and the cost. Test 2 run with `Spec_basis` = `dry` and then `ignited` (two real runs, all other data identical):

| | Dry basis | Ignited basis |
|---|---|---|
| Cost (₹ per dry t of blend) | 13,004 | 13,143 |
| Lime (wt %) | 52.09 | 40.75 |
| Calcined dolomite (wt %) | 2.60 | 18.20 |
| CaO achieved (wt %) | 50.00, binding | 52.11, slack |
| MgO achieved (wt %), Max 8 | 1.81, slack | 8.00, **binding** |

On an ignited basis every component counts for more, because the LOI is gone, and by a different amount for each raw material (Raw bauxite loses 24 wt % on ignition, Calcined bauxite only 1 wt %). The cheapest way to meet the limits changes: it uses far more Calcined dolomite, until the MgO maximum binds. Using the wrong basis gives the wrong recipe.

**Batch size and scale resolution.** `Batch_size_kg` is what the operator weighs: the as-received kg of all materials together. The report's weigh sheet ([7.2](#72-recipe)) adds up exactly to this number, rounded to `Scale_resolution_kg`.

An unknown key is ignored with a warning, for example (real run, a copy of Test 2 with an extra key `Colour`):

```
Input warnings (1):
  - Settings (Test2_unknown_setting.xlsx) row 5: unknown setting 'Colour' is ignored.
```

### 5.5 UsageLimits (optional)

Use this sheet for limits that apply to **this product or customer only**, for example a customer who does not accept fluorspar.

| Column | What to enter | If left blank |
|---|---|---|
| `Name` | A raw material name from RawMaterials | Error |
| `Min_use_%` | Minimum share for this product | 0 |
| `Max_use_%` | Maximum share for this product. **0 bans the material.** | 100 |

Both in wt % of the dry blend.

Example (`templates\Test2_synthetic_slag_fluorspar_ban.xlsx`):

| Name | Min_use_% | Max_use_% |
|---|---|---|
| Fluorspar | | 0 |

**The tighter limit applies.** When a material has a plant-wide limit in RawMaterials and a product limit here, the tool uses the tighter of the two on each side. Example: Lime fines has a plant-wide `Max_use_%` of 20. A product limit of 50 changes nothing (20 is tighter); a product limit of 10 lowers it to 10. If the two sides can no longer both be met (e.g. a product minimum of 30 against the plant maximum of 20), the tool reports an error.

Do not put a raw material name in the Specs sheet to ban it: Specs only accepts chemical components (see [11. Troubleshooting](#11-troubleshooting)).

### 5.6 WhatIf (optional)

A **scenario** is a set of changes to the workbook that the tool solves separately and compares with the unchanged workbook (the **base case**). You do not need to copy the workbook to try something: you list the changes here.

| Column | What to enter |
|---|---|
| `Scenario` | Any name, e.g. `Fluorspar dearer`. **Rows with the same name form one scenario.** |
| `Change type` | One of the six types below (the template has a drop-down list) |
| `Item` | The raw material or spec Name that the change applies to |
| `New value` | The new value |

| Change type | Item | New value | Blank New value means |
|---|---|---|---|
| `Price` | a raw material | New price, ₹ per tonne, on that material's own price basis | (not allowed) |
| `Spec Min` | a spec Name | New minimum | Remove the minimum |
| `Spec Max` | a spec Name | New maximum | Remove the maximum |
| `Spec Margin` | a spec Name | New margin | Margin 0 |
| `Usage Min` | a raw material | New product minimum, wt % of dry blend | Back to 0 |
| `Usage Max` | a raw material | New product maximum, wt % of dry blend | Back to 100 |

Usage changes replace the material's UsageLimits value for that scenario; the plant-wide limit in RawMaterials still applies, and the tighter one wins.

Example (`templates\Test2_synthetic_slag_with_what_if.xlsx`):

| Scenario | Change type | Item | New value |
|---|---|---|---|
| Fluorspar dearer | Price | Fluorspar | 35000 |
| Looser P | Spec Max | P | 0.035 |
| Capped bauxite | Usage Max | Raw bauxite | 20 |
| Capped bauxite | Spec Margin | CaO | 0.5 |
| No fluorspar | Usage Max | Fluorspar | 0 |

This gives four scenarios. "Capped bauxite" makes two changes at once. Each scenario is compared with the base case, not with the other scenarios. The results are in the report's Scenarios sheet and one sheet per scenario ([7.10](#710-scenarios), [7.11](#711-scenario-sheets)).

**The "current recipe" scenario: what is today's recipe costing?**

To measure what the optimum would save, cost **today's production recipe** with the same prices and limits. You do this with a scenario that fixes each raw material at its current share. Two things need care:

- WhatIf usage values are in **wt % of the dry blend**, while the shop floor works in **as-received kg**. You must convert first.
- The fixed shares must add up to **exactly 100 %**. Rounded shares often add up to 99.99 % or 100.01 %, and then no recipe exists.

**Worked example.** Today's recipe for the Test 2 product is, per 1,000 kg batch as weighed: Lime 515 kg, Calcined dolomite 40 kg, Calcined bauxite 205 kg, Raw bauxite 215 kg, Fluorspar 25 kg (invented figures).

1. For each material, work out the dry kg: **dry kg = wet kg × (1 − moisture / 100)**, with the moisture from RawMaterials.
2. Add up the dry kg.
3. Each material's share is **dry kg ÷ total dry kg × 100**.

| Raw material | As-received kg | Moisture_% | Dry kg | Dry wt % | Rounded |
|---|---|---|---|---|---|
| Lime | 515 | 0.5 | 512.425 | 52.1260 | 52.13 |
| Lime fines | 0 | 1.0 | 0 | 0 | 0 |
| Calcined dolomite | 40 | 0.5 | 39.800 | 4.0486 | 4.05 |
| Calcined bauxite | 205 | 0.5 | 203.975 | 20.7492 | 20.75 |
| Raw bauxite | 215 | 6.0 | 202.100 | 20.5585 | 20.56 |
| Fluorspar | 25 | 1.0 | 24.750 | 2.5177 | 2.52 |
| **Total** | **1,000** | | **983.050** | **100.0000** | **100.01** |

The rounded shares add up to 100.01 %.

4. **Fix every material except one** with `Usage Min` = `Usage Max` = its share, and **leave the largest one free** (here Lime). The free material takes whatever is left, so rounding can never break the 100 % rule. Set `Usage Max` = 0 for materials **not** in today's recipe, so that the free share cannot go to them.

| Scenario | Change type | Item | New value |
|---|---|---|---|
| Current recipe | Usage Min | Calcined dolomite | 4.05 |
| Current recipe | Usage Max | Calcined dolomite | 4.05 |
| Current recipe | Usage Min | Calcined bauxite | 20.75 |
| Current recipe | Usage Max | Calcined bauxite | 20.75 |
| Current recipe | Usage Min | Raw bauxite | 20.56 |
| Current recipe | Usage Max | Raw bauxite | 20.56 |
| Current recipe | Usage Min | Fluorspar | 2.52 |
| Current recipe | Usage Max | Fluorspar | 2.52 |
| Current recipe | Usage Max | Lime fines | 0 |

Result (real run, on a copy of Test 2 with these rows):

```
=== Scenario: Current recipe (vs base case) ===
...
Cost:
                                      Item    Base Scenario Change
                             Solver status Optimal  Optimal       
          Cost (Rs per dry tonne of blend)  13,004   13,944    940
  Cost (Rs per as-received tonne of blend)  12,745   13,708    963
                       Cost per batch (Rs)  12,745   13,708    963
```

Lime came out at 52.12 wt % (the remainder). Today's recipe meets the specification but costs **₹ 940 per dry tonne more** than the optimum (₹ 963 per as-received tonne and per 1,000 kg batch).

If instead you fix **all five** materials at their rounded shares, the scenario has no recipe, and the diagnosis says why (real run):

```
  - The minimum usage limits add up to 100.01 % of the blend, more than 100 % (Lime min 52.13 %, Calcined dolomite min 4.05 %, Calcined bauxite min 20.75 %, Raw bauxite min 20.56 %, Fluorspar min 2.52 %).
```

If today's recipe does not meet one of the limits, the "current recipe" scenario is also infeasible, and its diagnosis names the limit it fails. That is useful to know too.

### 5.7 Checklist before your first real run

- [ ] The workbook is in `data\`, not in `templates\` or elsewhere.
- [ ] Every raw material that could be used is listed, with its **landed** cost and the right `Price_basis`.
- [ ] Moisture and LOI are filled in for every material (blank means 0).
- [ ] Analyses are averages of recent lots, on a dry basis, and every component cell that should have a value has one (especially trace elements: P, S). A component that was analysed and is absent has a **0**, not a blank; no "not analysed" warning.
- [ ] Each material adds up to about 100 wt % (components + LOI); no total warning.
- [ ] Every limit on the customer's specification has a row in Specs, with the value in the right column (Min or Max), in wt % (not ppm).
- [ ] `Spec_basis` matches the customer's specification.
- [ ] `Batch_size_kg` is the as-received batch weight; `Scale_resolution_kg` matches your scale.
- [ ] Product- or customer-specific limits and bans are in UsageLimits, not in Specs.
- [ ] After the run: read the Spec check sheet and confirm that **every** customer limit is listed ([7.4](#74-spec-check)).

---

## 6. Running the tool

### 6.1 The command

1. **Close the workbook in Excel**, and close the previous report of the same product if it is open.
2. Open PowerShell in the tool's folder ([3.3](#33-open-powershell-in-the-tools-folder)).
3. Activate the environment ([3.5](#35-activate-the-virtual-environment)):

```
.\.venv\Scripts\Activate.ps1
```

4. Run the tool on your workbook. For a layout A workbook:

```
python run.py data\Slag_SS45.xlsx
```

For a layout B product, add the master raw-material file:

```
python run.py data\Slag_SS45.xlsx --raw-materials data\Raw_materials.xlsx
```

(In this guide these two commands were verified with the example workbooks in the same positions: `python run.py templates\Test2_synthetic_slag_dry.xlsx`, and `python run.py templates\Test2_product_only.xlsx --raw-materials templates\Example_master_raw_materials.xlsx`.)

5. Read the screen ([6.5](#65-reading-the-screen)), then open the report.

If a path contains a space, put it in quotes: `python run.py "data\SS45 slag.xlsx"`.

### 6.2 Options

`python run.py --help` lists them (real output):

```
usage: run.py [-h] [--raw-materials RAW_MATERIALS_FILE] [--output OUTPUT_PATH]
              [--quiet]
              product_file

Least-cost blend optimizer: find the cheapest raw material mix that meets the
specs.

positional arguments:
  product_file          product workbook (.xlsx)

options:
  -h, --help            show this help message and exit
  --raw-materials RAW_MATERIALS_FILE
                        master raw-material workbook, if RawMaterials is not
                        in the product workbook
  --output OUTPUT_PATH  report workbook to write (default:
                        output\<product>_result.xlsx)
  --quiet               do not print the summary
```

| Option | What it does | Example |
|---|---|---|
| *(product file)* | The workbook to run. Always first. | `data\Slag_SS45.xlsx` |
| `--raw-materials` | The master raw-material workbook (layout B) | `--raw-materials data\Raw_materials.xlsx` |
| `--output` | The name and place of the report | `--output output\Slag_SS45_June.xlsx` |
| `--quiet` | Print only where the report was written, not the summary | `--quiet` |
| `--help` | Show the list above | `--help` |

### 6.3 Where the report goes

By default the report is `output\<name of the product file>_result.xlsx`, and the `.lp` file ([7.12](#712-the-lp-file)) is next to it with the same name and the extension `.lp`. For `templates\Test1_hand_worked.xlsx` these are `output\Test1_hand_worked_result.xlsx` and `output\Test1_hand_worked_result.lp`. Running the same product again **overwrites** them.

To keep a history, put the date in the name with `--output`. PowerShell can insert today's date for you:

```
python run.py templates\Test2_synthetic_slag_dry.xlsx --output "output\Test2_synthetic_slag_dry_$(Get-Date -Format yyyy-MM-dd).xlsx" --quiet
```

Real output:

```
Report written to: output\Test2_synthetic_slag_dry_2026-10-26.xlsx
LP file:           output\Test2_synthetic_slag_dry_2026-10-26.lp
```

### 6.4 Time-savers

- **Tab completes names.** Type the start of a file or folder name, e.g. `python run.py data\Sl`, and press Tab. PowerShell completes it, so you avoid typing errors. Press Tab again to cycle through matches.
- **↑ repeats a command.** Press the up-arrow key to bring back the previous command, change it if needed, and press Enter.
- **Close the files in Excel first.** If the report from an earlier run is open in Excel, the tool cannot overwrite it (real output):

```
Could not write the report: [Errno 13] Permission denied: 'D:\\VS Python\\Cost Optimizer\\output\\Test1_hand_worked_result.xlsx'
If the report is open in Excel, close it and run again.
```

- **Run from the tool's folder.** The commands in this guide assume the PowerShell window is in the tool's folder. If you type `run.py` wrongly, or are in another folder, Python says (real output, with `rn.py` typed by mistake):

```
python: can't open file 'D:\\VS Python\\Cost Optimizer\\rn.py': [Errno 2] No such file or directory
```

### 6.5 Reading the screen

A run ends in one of three ways. **The details are always printed above the last lines**, so scroll up if the window is small.

**1. A recipe was found.** The summary starts with the product, the status `Optimal` and the cost, and ends with where the report was written. (Test 1, see [1.3](#13-the-whole-idea-on-one-page-test-1).)

**2. Errors: nothing was solved.** If the workbook has errors, the tool lists every error and warning and stops. Real output for `templates\Broken_example.xlsx`, an example workbook with nine deliberate mistakes:

```
9 input error(s) found:
  - RawMaterials (Broken_example.xlsx) row 3 'Lime': duplicate name (already used in row 2; names must be unique, ignoring case).
  - RawMaterials (Broken_example.xlsx) row 4 'Calcined dolomite', Cost_Rs_per_t: '8,500 Rs' is not a number. Use a number, or leave the cell blank.
  - RawMaterials (Broken_example.xlsx) row 5 'Calcined bauxite', Price_basis: 'wet' is not allowed; use one of as_received, dry.
  - RawMaterials (Broken_example.xlsx) row 6 'Raw bauxite', Moisture_%: -6 is negative; it must be 0 or more.
  - Specs (Broken_example.xlsx) row 2 'CaO': after the margin the internal range is empty: Min + margin = 54 is greater than Max - margin = 52. Reduce the margin.
  - Specs (Broken_example.xlsx) row 3 'Al2O3', Numerator: component 'AL2O3' is not a column in RawMaterials. Did you mean 'Al2O3'? (Component names must match exactly.)
  - Specs (Broken_example.xlsx) row 5 'MgO': give a Min, a Max, or both.
  - Settings (Broken_example.xlsx) row 3 'Spec_basis': 'calcined' is not allowed; use one of dry, as_received, ignited.
  - UsageLimits (Broken_example.xlsx) row 2 'Flourspar': not a raw material in RawMaterials. Did you mean 'Fluorspar'?
1 warning(s):
  - RawMaterials 'Lime': components + LOI = 87.34 %, outside the expected 97-103 % for the 'dry_incl_loi' lab convention. Neither convention fits (components alone = 83.84 %, components + LOI = 87.34 %): check for a missing component column, a typing error or a missing LOI.
Nothing was solved. Correct the workbook and run again.
```

Each message names the sheet, the file, the Excel row number and the item. Correct them all in Excel, save, and run again. No report is written.

**3. Warnings: the run continues.** Warnings appear near the top of the summary, under `Input warnings`, and in the report's Summary sheet. Read them: a warning is often a sign of a typing error. Example in [5.4](#54-settings).

**4. No recipe exists.** The status is `Infeasible`, followed by an explanation. The report is still written. See [chapter 8](#8-when-there-is-no-recipe-infeasible).

**Exit codes.** When the tool finishes it also returns a number to Windows, the **exit code**. You only need it if you run the tool from a batch file or a script that should react to the result:

| Exit code | Meaning |
|---|---|
| 0 | A recipe was found for the base case. |
| 1 | Input errors: nothing was solved (case 2 above). Also when the report could not be written, e.g. because it is open in Excel. |
| 2 | No recipe for the base case: it is infeasible (case 4 above). The report, with the diagnosis, is still written. |

A WhatIf scenario that has no recipe does not change the exit code; it is reported in its own sheet. To see the exit code of the last run, type `$LASTEXITCODE` in PowerShell (in a Command Prompt batch file, `%errorlevel%`). Real example, after running `templates\Test1_infeasible_Al2O3_min_45.xlsx`:

```
PS D:\VS Python\Cost Optimizer> $LASTEXITCODE
2
```

---

## 7. Reading the report

The report is an Excel workbook with these sheets, in this order:

| Sheet | When it appears |
|---|---|
| Summary | Always |
| Recipe, Composition, Spec check, Cost, Shadow prices, Reduced costs | When a recipe was found |
| Diagnosis | Only when no recipe exists |
| Price ranges | Always (empty, with a note, when no recipe exists) |
| Scenarios, and one sheet per scenario | Only when the product workbook has a WhatIf sheet |

General points:

- Numbers are calculated at full precision and only **rounded for display**: wt % to 2 decimals (spec limits get one more decimal than you typed, e.g. P 0.030), kg to the scale resolution, ₹ to whole rupees. An Excel cell holds the full value; widen the column or change the number format to see more.
- An exact half is always **rounded up**, the same way on screen and in Excel: 78.895 shows as 78.90, and ₹ 12,744.5 as ₹ 12,745.
- An **empty cell** means "not applicable" (e.g. no minimum).
- Short notes in *italics* under each table explain it.
- All excerpts below are from real runs of `templates\Test2_synthetic_slag_dry.xlsx` unless stated otherwise.

### 7.1 Summary

What was run, and how it went.

| Item | Value |
|---|---|
| Product | Synthetic slag (Test 2, dry basis) |
| Solver status | Optimal |
| Spec basis | dry |
| Lab analysis convention | dry_incl_loi |
| Batch size | 1,000.0 kg as received |
| Scale resolution | 0.1 kg |
| Raw materials read from | templates\Test2_synthetic_slag_dry.xlsx |
| Specs read from | templates\Test2_synthetic_slag_dry.xlsx |
| Input warnings | 0 |

| Item | Meaning |
|---|---|
| Product | `Product_name` from Settings |
| Solver status | `Optimal`: the cheapest recipe was found. `Infeasible`: no recipe meets every limit (see the Diagnosis sheet). |
| Spec basis | `Spec_basis` from Settings |
| Lab analysis convention | How analyses were read ([5.2](#52-rawmaterials)) |
| Batch size, Scale resolution | From Settings |
| Raw materials read from, Specs read from | The files used (useful for layout B) |
| Input warnings | The number of warnings; each one follows as "Warning 1", "Warning 2" … |

**Use it** to confirm you ran the right file with the right basis, and to read any warnings.

### 7.2 Recipe

The recipe and the weigh sheet.

| Raw material | Dry blend (wt %) | Dry mass (kg/batch) | As-received mass (kg/batch) | Moisture (%) | Price as entered (Rs/t) | Price basis | Dry cost (Rs/dry t) | Cost share (Rs/dry t of blend) |
|---|---|---|---|---|---|---|---|---|
| Lime | 52.09 | 510.5 | 513.1 | 0.50 | 9,000 | as_received | 9,045 | 4,712 |
| Lime fines | 0.00 | 0.0 | 0.0 | 1.00 | 6,500 | as_received | 6,566 | 0 |
| Calcined dolomite | 2.60 | 25.5 | 25.6 | 0.50 | 8,500 | as_received | 8,543 | 222 |
| Calcined bauxite | 17.13 | 167.9 | 168.7 | 0.50 | 32,000 | dry | 32,000 | 5,481 |
| Raw bauxite | 25.83 | 253.1 | 269.3 | 6.00 | 7,000 | as_received | 7,447 | 1,923 |
| Fluorspar | 2.35 | 23.1 | 23.3 | 1.00 | 28,000 | as_received | 28,283 | 665 |
| TOTAL | 100.00 | 980.1 | 1,000.0 | | | | | 13,004 |

*Batch = 1,000.0 kg AS RECEIVED (what is weighed into the mixer).*
*kg are rounded to the scale resolution (0.1 kg); each kg column adds up exactly to its TOTAL.*

| Column | Meaning |
|---|---|
| Raw material | Every raw material, including those not used (0.00) |
| Dry blend (wt %) | **The recipe**: the share of each material in the dry blend |
| Dry mass (kg/batch) | The same, as dry kg in one batch |
| As-received mass (kg/batch) | **What the operator weighs**: kg as delivered, including moisture. Adds up to `Batch_size_kg`. |
| Moisture (%) | From RawMaterials |
| Price as entered (Rs/t), Price basis | From RawMaterials |
| Dry cost (Rs/dry t) | The price converted to a dry tonne |
| Cost share (Rs/dry t of blend) | Share × dry cost: each material's contribution to the blend cost. The TOTAL is the blend cost. |

**Dry and as-received kg.** The batch size is **as received**: 1,000.0 kg is weighed in total. That batch contains only 980.1 kg of dry material, because the rest is moisture. Raw bauxite, with 6 wt % moisture, needs 269.3 kg as received to supply 253.1 kg dry (253.1 ÷ 0.94 = 269.3).

**Rounding to the scale.** Each kg is rounded to `Scale_resolution_kg`. Rounded figures do not always add up to the batch size, so the tool puts the rounding difference on the **largest** material and says so. With a 5 kg scale (real run on a copy of Test 2 with `Scale_resolution_kg` = 5):

```
       Raw material Dry blend (wt %) Dry mass (kg/batch) As-received mass (kg/batch)
               Lime            52.09               505.0                       510.0
         Lime fines             0.00                 0.0                         0.0
  Calcined dolomite             2.60                25.0                        25.0
   Calcined bauxite            17.13               170.0                       170.0
        Raw bauxite            25.83               255.0                       270.0
          Fluorspar             2.35                25.0                        25.0
              TOTAL           100.00               980.0                     1,000.0
  kg are rounded to the scale resolution (5 kg); each kg column adds up exactly to its TOTAL.
  Dry mass (kg/batch): rounding difference of -5.0 kg put on Lime (the largest).
  As-received mass (kg/batch): rounding difference of -5.0 kg put on Lime (the largest).
```

**Use it**: the As-received column is the weigh sheet for the shop floor; the Cost share column shows where the money goes.

### 7.3 Composition

The blend's analysis.

| Component | Basis | Blend (wt %) |
|---|---|---|
| CaO | dry | 50.00 |
| SiO2 | dry | 4.38 |
| Al2O3 | dry | 30.00 |
| MgO | dry | 1.81 |
| CaF2 | dry | 2.00 |
| Fe2O3 | dry | 1.66 |
| TiO2 | dry | 1.25 |
| S | dry | 0.040 |
| P | dry | 0.030 |
| Sum of components | dry | 91.17 |
| LOI | dry | 8.29 |
| Moisture | as_received | 1.99 |

| Column or row | Meaning |
|---|---|
| Component | Every component column of RawMaterials, limited or not (TiO2 has no limit here) |
| Basis | The spec basis for the components; LOI is always on a dry basis, moisture always as received |
| Blend (wt %) | The blend's content. A component limited by its own spec row is shown with one more decimal than the limit you typed (P 0.030). |
| Sum of components | A check: on a dry basis, Sum of components + LOI should be about 100 wt % (here 91.17 + 8.29 = 99.46) |
| LOI | The blend's loss on ignition, wt % dry basis |
| Moisture | The blend's free moisture, wt % as received |

**Use it** to compare with the lab analysis of a trial batch ([chapter 10](#10-before-using-a-recipe-in-production)). Make sure the lab reports on the same basis.

### 7.4 Spec check

Every spec row, with the customer limit, the target and the value achieved.

| Spec | Limits | Basis | Customer Min | Customer Max | Margin | Target Min | Target Max | Achieved | Status | Slack to target |
|---|---|---|---|---|---|---|---|---|---|---|
| CaO | CaO | dry | 50.00 | 56.00 | 0.00 | 50.00 | 56.00 | 50.00 | Binding (Min) | 0.00 |
| Al2O3 | Al2O3 | dry | 30.00 | 38.00 | 0.00 | 30.00 | 38.00 | 30.00 | Binding (Min) | 0.00 |
| SiO2 | SiO2 | dry | | 6.00 | 0.00 | | 6.00 | 4.38 | Slack | 1.62 |
| MgO | MgO | dry | | 8.00 | 0.00 | | 8.00 | 1.81 | Slack | 6.19 |
| CaF2 | CaF2 | dry | 2.00 | 5.00 | 0.00 | 2.00 | 5.00 | 2.00 | Binding (Min) | 0.00 |
| Fe2O3 | Fe2O3 | dry | | 2.00 | 0.00 | | 2.00 | 1.66 | Slack | 0.34 |
| S | S | dry | | 0.080 | 0.000 | | 0.080 | 0.040 | Slack | 0.040 |
| P | P | dry | | 0.030 | 0.000 | | 0.030 | 0.030 | Binding (Max) | 0.000 |
| Basicity | CaO / SiO2 | ratio (any basis) | 8.00 | | 0.00 | 8.00 | | 11.40 | Slack | 3.40 |

| Column | Meaning |
|---|---|
| Spec | The spec row's Name |
| Limits | What is limited, as typed: a component, a sum, a ratio (`CaO / SiO2`) or an expression with factors |
| Basis | The spec basis, or "ratio (any basis)" |
| Customer Min, Customer Max | The customer's limits, as typed in Specs |
| Margin | Your safety margin |
| Target Min, Target Max | What the tool aimed for: Min + margin and Max − margin |
| Achieved | The blend's value |
| Status | `Binding (Min)`: sitting on the minimum target; `Binding (Max)`: on the maximum target; `Slack`: room to spare. (If Min equals Max, `Binding (Min and Max)`.) |
| Slack to target | How far the achieved value is from the nearest target, in the spec's own unit |

**Read this sheet after every run**, and confirm that **every limit on the customer's specification is listed**. A missing spec row gives **no error**: the tool simply optimises without it. Real example: Test 1 with the Al2O3 row deleted from Specs. The tool chooses 100 % Lime, because nothing asks for Al2O3:

```
Cost: Rs 9,000 per dry t of blend | Rs 9,000 per as-received t | Rs 9,000 per batch

Recipe (batch 1,000.0 kg as received, 1,000.0 kg dry):
      Raw material Dry blend (wt %) Dry mass (kg/batch) As-received mass (kg/batch)
              Lime           100.00             1,000.0                     1,000.0
  Calcined alumina             0.00                 0.0                         0.0
           Bauxite             0.00                 0.0                         0.0
             TOTAL           100.00             1,000.0                     1,000.0
...
Spec check:
  Spec Limits Basis Target Min Target Max Achieved Status
   CaO    CaO   dry      50.00               92.00  Slack
  SiO2   SiO2   dry                  6.00     2.00  Slack
```

Also check that each value is in the right column: a maximum typed in the Min column changes the problem completely (see [11. Troubleshooting](#11-troubleshooting)).

### 7.5 Cost

| Item | Value | Unit |
|---|---|---|
| Cost (Rs per dry tonne of blend) | 13,004 | Rs/dry t |
| Cost (Rs per as-received tonne of blend) | 12,745 | Rs/as-received t |
| Batch size (kg, as received) | 1,000.0 | kg |
| Batch dry mass (kg) | 980.1 | kg |
| Cost per batch (Rs) | 12,745 | Rs |

- **Per dry tonne** is what the tool minimises; it is the figure to compare between recipes and scenarios.
- **Per as-received tonne** is lower, because an as-received tonne contains some moisture. It is the cost of a tonne as weighed and shipped.
- **Per batch** is the cost of one batch of `Batch_size_kg` (here 1,000 kg as received).

### 7.6 Shadow prices

The **shadow price** of a binding limit is how much the blend would get cheaper, in ₹ per dry tonne of blend, if the customer relaxed that limit a little: lowering a minimum or raising a maximum. Tightening it costs the same amount per step.

| Constraint | Type | Limits | Customer limit | Internal target | Limit unit | Relax by | Saving per 1 unit (Rs/dry t) | Step | Saving per step (Rs/dry t) |
|---|---|---|---|---|---|---|---|---|---|
| CaO_min | Composition | CaO | 50.00 | 50.00 | wt % (dry basis) | lowering the Min | 14 | 1 | 14 |
| Al2O3_min | Composition | Al2O3 | 30.00 | 30.00 | wt % (dry basis) | lowering the Min | 637 | 1 | 637 |
| CaF2_min | Composition | CaF2 | 2.00 | 2.00 | wt % (dry basis) | lowering the Min | 327 | 1 | 327 |
| P_max | Composition | P | 0.030 | 0.030 | wt % (dry basis) | raising the Max | 735,134 | 0.01 | 7,351 |

| Column | Meaning |
|---|---|
| Constraint | The limit: `<spec Name>_min` or `_max`, or `<raw material> max use` / `min use` for a usage limit |
| Type | `Composition`, `Ratio` or `Usage limit` |
| Limits | What is limited |
| Customer limit, Internal target | The limit as typed, and with the margin applied |
| Limit unit | The unit of one step |
| Relax by | `lowering the Min` or `raising the Max` |
| Saving per 1 unit (Rs/dry t) | Saving for relaxing by 1 wt % (or 1 ratio unit, or 1 wt % of a usage limit) |
| Step | A step sized to the limit: 1 for limits of 1 or more; below 1, one unit of the limit's first significant decimal (0.03 → 0.01, 0.5 → 0.1, 0.005 → 0.001) |
| Saving per step (Rs/dry t) | Saving for relaxing by one step. **Use this column**: for P, 1 wt % is meaningless, 0.01 wt % is a real step. |

Only binding limits appear. A slack limit has a shadow price of 0.

**Valid only near the current limit.** The saving holds while the same limits stay binding. Further away the recipe changes, and so does the saving. Two real checks on Test 2, done as WhatIf scenarios:

| Change | Predicted from the shadow price | Real saving (Scenarios sheet) |
|---|---|---|
| Al2O3 Min 30 → 29 | 1 step × ₹ 637 = ₹ 637 | ₹ 637 per dry t |
| P Max 0.030 → 0.031 | 0.1 step × ₹ 7,351 = ₹ 735 | ₹ 550 per dry t |

The Al2O3 prediction holds exactly; for P the recipe changes before even 0.001 wt % is reached. So use shadow prices to see **which limits are expensive**, and use a WhatIf scenario to get the exact figure for a specific change.

**In customer discussions.** Shadow prices show which limits are worth discussing. Here P is by far the most expensive limit per step, Al2O3 next, and CaO costs almost nothing. To put a figure on a change, multiply the exact saving by the annual quantity: if you sell 10,000 dry tonnes a year, accepting Al2O3 29 instead of 30 wt % is worth 10,000 × ₹ 637 = ₹ 63,70,000 (₹ 63.7 lakh) a year.

The notes under the table also warn that for ratio limits and limits on an ignited or as-received basis the saving drifts slightly as the limit moves, even before the recipe changes.

### 7.7 Reduced costs

For each raw material **not used** in the recipe: how much cheaper it would have to be before it is worth using.

| Raw material | Price as entered (Rs/t) | Price basis | Reduced cost (Rs/dry t) | Price drop needed (Rs/t as entered) | Enters below (Rs/t as entered) | Note |
|---|---|---|---|---|---|---|
| Lime fines | 6,500 | as_received | 4,674 | 4,628 | 1,872 | |

| Column | Meaning |
|---|---|
| Raw material | An unused raw material (with no minimum use) |
| Price as entered (Rs/t), Price basis | From RawMaterials |
| Reduced cost (Rs/dry t) | The price drop needed, per **dry** tonne |
| Price drop needed (Rs/t as entered) | The same drop on the material's own price basis (here per tonne as received, with 1 wt % moisture) |
| Enters below (Rs/t as entered) | The price below which the material would come into the recipe |
| Note | Explains special cases: a material banned by a usage limit ("Banned by its usage limit (Max 0 %)"), one that would be used if the ban were lifted, one that is "Break-even: another recipe of the same cost uses it", or one that "Would not be used even if free" |

In Test 2, Lime fines would enter the recipe only below ₹ 1,872 per tonne, far below its ₹ 6,500. The Price ranges sheet gives the same price as 1,873: that search works to ₹ 1, so the two can differ by ₹ 1.

When every raw material is used, the sheet is empty with the note *None: every raw material is used.*

**Use it** to answer a supplier: "at what price would your material be worth using?"

### 7.8 Diagnosis

This sheet appears **only when no recipe exists**. It lists the limits that have to move.

Real example (`templates\Test2_synthetic_slag_fluorspar_ban.xlsx`):

| Limit | Limits | Found by | Customer limit | Internal target | Best achievable | Change needed | While holding | Elastic model target |
|---|---|---|---|---|---|---|---|---|
| CaF2_min | CaF2 | Pre-check | 2.00 | 2.00 | 0.00 | 2.00 | usage limits only | |
| CaF2_min | CaF2 | Elastic model | 2.00 | 2.00 | 0.00 | 2.00 | all other limits | 0.00 |
| Fluorspar max use | Fluorspar | Elastic model | | 0.00 | | | usage limit | |

| Column | Meaning |
|---|---|
| Limit | The limit involved |
| Limits | What it limits |
| Found by | `Pre-check` (the limit cannot be met on its own) or `Elastic model` (it conflicts with other limits) |
| Customer limit, Internal target | As in the Spec check |
| Best achievable | The best value possible: with only the usage limits (pre-check), or with all the other limits held (elastic model) |
| Change needed | How far the limit would have to move to the best achievable value |
| While holding | What was kept in place when working out the best achievable value |
| Elastic model target | Where the elastic model's suggested change would put this limit (blank if it would not move it) |

The plain-language explanation is written under the table, and also printed on screen. [Chapter 8](#8-when-there-is-no-recipe-infeasible) explains how to read it.

### 7.9 Price ranges

For each raw material, taken on its own with all other prices unchanged: the price band within which **today's recipe stays exactly the same**. Inside the band only the cost changes; outside it, the tool would choose a different recipe.

| Raw material | Price basis | Current price (Rs/t) | In recipe (wt %) | Unchanged from (Rs/t) | Unchanged up to (Rs/t) | Note |
|---|---|---|---|---|---|---|
| Lime | as_received | 9,000 | 52.09 | 8,500 | 13,199 | |
| Lime fines | as_received | 6,500 | 0.00 | 1,873 | | No upper limit found up to Rs 65,000 (10x the current price); Not used now; enters the recipe below Rs 1,873 |
| Calcined dolomite | as_received | 8,500 | 2.60 | | 9,000 | No lower limit: unchanged down to Rs 0 |
| Calcined bauxite | dry | 32,000 | 17.13 | 15,726 | | No upper limit found up to Rs 320,000 (10x the current price) |
| Raw bauxite | as_received | 7,000 | 25.83 | | 17,223 | No lower limit: unchanged down to Rs 0 |
| Fluorspar | as_received | 28,000 | 2.35 | 517 | | No upper limit found up to Rs 280,000 (10x the current price) |

| Column | Meaning |
|---|---|
| Price basis, Current price (Rs/t) | As in RawMaterials; the band is on the same basis |
| In recipe (wt %) | Today's share |
| Unchanged from (Rs/t) | The lowest price with the same recipe. Blank: no lower limit (the recipe is the same down to ₹ 0). |
| Unchanged up to (Rs/t) | The highest price with the same recipe. Blank: none found up to 10 times today's price. |
| Note | Explains blanks and unused materials |

The bands are found by solving again at trial prices, accurate to ₹ 1 per tonne. A recipe counts as unchanged if no share moves by more than 0.01 wt % (0.1 kg in a 1,000 kg batch). In Test 1, Lime has the note *Recipe unchanged at every price tested (Rs 0 to 90,000)*: its share is fixed by the limits, whatever it costs.

On screen the same table is printed without the Note column, and the notes follow under it. Test 1 (real run):

```
Price ranges (base case): today's recipe stays the same while one price moves within this band
      Raw material Price basis Current price (Rs/t) Unchanged from (Rs/t) Unchanged up to (Rs/t)
              Lime         dry                9,000                                             
  Calcined alumina         dry               30,000                13,271                       
           Bauxite         dry               12,000                                       23,750
  Notes:
    Lime: Recipe unchanged at every price tested (Rs 0 to 90,000)
    Calcined alumina: No upper limit found up to Rs 300,000 (10x the current price)
    Bauxite: No lower limit: unchanged down to Rs 0
```

**Use it** when a supplier changes a price. Example: if Raw bauxite rises from ₹ 7,000 to ₹ 9,000 per tonne, that is inside its band (up to ₹ 17,223), so the recipe stays the same and only the cost rises. Calcined dolomite above ₹ 9,000 would change the recipe; run the tool again (or a WhatIf scenario) to see how.

When no recipe exists, the sheet is empty with the note *Not available: price ranges need a feasible base case (see the Diagnosis sheet).*

### 7.10 Scenarios

Appears when the product workbook has a WhatIf sheet: one row per scenario. Real run of `templates\Test2_synthetic_slag_with_what_if.xlsx`:

| Scenario | Sheet | Changes | Status | Cost (Rs/dry t) | Change vs base (Rs/dry t) |
|---|---|---|---|---|---|
| Fluorspar dearer | Scenario Fluorspar dearer | Price Fluorspar 28,000 -> 35,000 | Optimal | 13,170 | 166 |
| Looser P | Scenario Looser P | Spec Max P 0.0300 -> 0.0350 | Optimal | 12,381 | -623 |
| Capped bauxite | Scenario Capped bauxite | Spec Margin CaO 0.00 -> 0.50; Usage Max (effective) Raw bauxite 100.00 -> 20.00 | Optimal | 13,645 | 641 |
| No fluorspar | Scenario No fluorspar | Usage Max (effective) Fluorspar 100.00 -> 0.00 | Infeasible | | |

| Column | Meaning |
|---|---|
| Scenario | The name from the WhatIf sheet |
| Sheet | The sheet with the details |
| Changes | What the scenario changed, from → to. "(effective)" means the usage limit actually applied, after combining plant and product limits. |
| Status | `Optimal` or `Infeasible` |
| Cost (Rs/dry t) | The scenario's cost |
| Change vs base (Rs/dry t) | Scenario minus base: **negative = cheaper** than the base case |

Reading this example: a dearer fluorspar adds ₹ 166 per dry tonne; accepting 0.035 wt % P saves ₹ 623; capping raw bauxite at 20 wt % with a 0.5 wt % CaO margin adds ₹ 641; banning fluorspar makes the product impossible.

### 7.11 Scenario sheets

One sheet per scenario, named `Scenario <name>`, with these tables one under another: changes, cost, recipe, spec check and binding limits, plus a diagnosis when the scenario is infeasible. Real excerpt from `Scenario Looser P`:

**Looser P: changes**

| Change | Name | Unit | Base value | Scenario value |
|---|---|---|---|---|
| Spec Max | P | wt % (dry basis) | 0.0300 | 0.0350 |

**Looser P: cost**

| Item | Unit | Base | Scenario | Change |
|---|---|---|---|---|
| Solver status | | Optimal | Optimal | |
| Cost (Rs per dry tonne of blend) | Rs/dry t | 13,004 | 12,381 | -623 |
| Cost (Rs per as-received tonne of blend) | Rs/as-received t | 12,745 | 12,118 | -627 |
| Cost per batch (Rs) | Rs | 12,745 | 12,118 | -627 |

**Looser P: recipe**

| Raw material | Base (wt %) | Scenario (wt %) | Change (wt %) | Base (kg/batch, as received) | Scenario (kg/batch, as received) |
|---|---|---|---|---|---|
| Lime | 52.09 | 34.57 | -17.52 | 513.1 | 340.0 |
| Lime fines | 0.00 | 20.00 | 20.00 | 0.0 | 197.7 |
| Calcined dolomite | 2.60 | 0.00 | -2.60 | 25.6 | 0.0 |
| Calcined bauxite | 17.13 | 16.57 | -0.56 | 168.7 | 163.0 |
| Raw bauxite | 25.83 | 26.51 | 0.69 | 269.3 | 276.0 |
| Fluorspar | 2.35 | 2.35 | 0.00 | 23.3 | 23.3 |
| TOTAL | 100.00 | 100.00 | 0.00 | 1,000.0 | 1,000.0 |

**Looser P: spec check** (shortened)

| Spec | Limits | Customer Min | Customer Max | Margin | Base achieved | Scenario achieved | Base status | Scenario status |
|---|---|---|---|---|---|---|---|---|
| CaO | CaO | 50.00 | 56.00 | 0.00 | 50.00 | 50.00 | Binding (Min) | Binding (Min) |
| P | P | | 0.0350 | 0.0000 | 0.0300 | 0.0321 | Binding (Max) | Slack |

**Looser P: binding limits**

| Constraint | Limit unit | Step | Base saving per step (Rs/dry t) | Scenario saving per step (Rs/dry t) |
|---|---|---|---|---|
| CaO_min | wt % (dry basis) | 1 | 14 | 599 |
| Al2O3_min | wt % (dry basis) | 1 | 637 | 920 |
| CaF2_min | wt % (dry basis) | 1 | 327 | 862 |
| P_max | wt % (dry basis) | 0.01 | 7,351 | |
| Lime fines max use | wt % of dry blend | 1 | | 4 |

How to read it:

- The spec check shows the **scenario's** limits (here P Max 0.035) with both achieved values. With the looser limit, P is no longer binding: the blend settles at 0.0321 wt %.
- In the binding limits table a blank means "not binding in that case". With the looser P, Lime fines becomes worth using up to its plant maximum of 20 wt %, so its usage limit now binds.
- The recipe table gives the as-received kg for both cases: the scenario's weigh sheet.

### 7.12 The .lp file

Next to each report the tool writes a plain-text file with the extension `.lp`, for example `output\Test2_synthetic_slag_dry_result.lp`. It is the optimisation problem exactly as the solver received it. You do not need it for normal use, but it lets you check a limit by hand. Open it with Notepad.

In it, `x_Lime` stands for the share of Lime in the dry blend as a fraction (0.5209 = 52.09 wt %). Each limit is one line, named after the spec row. A composition limit on a dry basis reads like the hand calculation, for example from Test 1:

```
CaO_min: 2 x_Bauxite + x_Calcined_alumina + 92 x_Lime >= 50
```

(a coefficient of 1 is not written). A ratio limit is rewritten so that it is a weighted average. Test 2's basicity row, CaO / SiO2 at least 8, becomes "CaO − 8 × SiO2 at least 0":

```
Basicity_min: - 47 x_Calcined_bauxite + 45 x_Calcined_dolomite
 - 70.5 x_Fluorspar + 80 x_Lime + 68 x_Lime_fines - 70.5 x_Raw_bauxite >= 0
```

Check one coefficient by hand: for each raw material it is CaO − 8 × SiO2 from RawMaterials.

| Raw material | CaO | SiO2 | CaO − 8 × SiO2 |
|---|---|---|---|
| Lime | 92 | 1.5 | 92 − 12 = 80 |
| Lime fines | 88 | 2.5 | 88 − 20 = 68 |
| Calcined dolomite | 57 | 1.5 | 57 − 12 = 45 |
| Calcined bauxite | 1 | 6 | 1 − 48 = −47 |
| Raw bauxite | 1.5 | 9 | 1.5 − 72 = −70.5 |
| Fluorspar | 1.5 | 9 | 1.5 − 72 = −70.5 |

Limits on an ignited or as-received basis are rewritten in the same way, with 0 on the right-hand side. The line `Sum_of_fractions` says that the shares add up to 1 (100 %). When no recipe exists, the tool also writes `<report name>_elastic.lp`, the problem used for the diagnosis.

---

## 8. When there is no recipe (infeasible)

**Infeasible** means that no mix of the allowed raw materials meets every limit at the same time. The status is `Infeasible`, the report is still written, and the tool explains why, on screen and in the Diagnosis sheet.

### 8.1 How the tool explains it

The explanation has three steps.

1. **Pre-check.** The tool looks at each limit **on its own**, with only the usage limits applied, and asks: can any mix of the allowed raw materials reach it at all? This catches limits that can never be met, such as a component that no allowed material contains, and usage limits that leave no blend at all (minimums adding up to more than 100 %).
2. **Elastic model.** If every limit can be met on its own, some limits **conflict** with each other. The tool then lets each limit "stretch" and finds the smallest total stretch that makes a recipe possible, measured relative to the size of each limit. The limits it has to stretch, or that push against the stretched ones, are the candidates for the conflicting set.
3. **Only the limits really needed.** The tool then tests each candidate by taking it away. If there is still no recipe without it, it is not part of the conflict and is left out of the list. So every limit that is named matters: relaxing any one of them far enough removes this conflict. For each limit named, the tool works out the **best value achievable while all the other limits are held**.

It finally suggests **one** smallest change that makes a blend possible. There may be other ways: moving any of the conflicting limits will do.

### 8.2 Example 1: a banned raw material

`templates\Test2_synthetic_slag_fluorspar_ban.xlsx` is Test 2 with a UsageLimits row that bans Fluorspar (Max 0). Real output:

```
Status: Infeasible | spec basis: dry | lab convention: dry_incl_loi
No recipe: the solver status is 'Infeasible'.

No blend meets every limit (the problem is infeasible).
Pre-check (each limit on its own, with only the usage limits applied):
  - No allowed raw material contains CaF2, so CaF2_min (CaF2 >= 2.00 %) cannot be met. Raw materials that would help are restricted: Fluorspar (CaF2 85.00 %) is not allowed (Max 0 %, UsageLimits).
Elastic model: these limits cannot all be met together: CaF2_min, Fluorspar max use.
  - CaF2_min (CaF2 >= 2.00 %): the highest CaF2 achievable while Fluorspar <= 0 % (usage limit) and all other limits are held is 0.00 %.
Smallest total relative change that makes a blend possible (one way; there may be others): lower CaF2_min from 2.00 % to 0.00 %.
```

How to read it:

- The pre-check already finds the cause: only Fluorspar contains CaF2, and it is banned (by the UsageLimits sheet). The CaF2 minimum of 2 wt % can never be met.
- The elastic model confirms it: CaF2_min and the fluorspar ban cannot both hold.
- The way out: either the customer drops the CaF2 minimum, or the ban is lifted, or you add a raw material that contains CaF2 and is allowed.

### 8.3 Example 2: two limits that conflict

`templates\Test1_infeasible_Al2O3_min_45.xlsx` is Test 1 with the Al2O3 minimum raised from 30 to 45 wt %. Real output:

```
No blend meets every limit (the problem is infeasible).
Pre-check: every limit can be met on its own, so some limits conflict with each other.
Elastic model: these limits cannot all be met together: CaO_min, Al2O3_min.
  - CaO_min (CaO >= 50.00 %): the highest CaO achievable while Al2O3 >= 45.00 % and all other limits are held is 44.33 %.
  - Al2O3_min (Al2O3 >= 45.00 %): the highest Al2O3 achievable while CaO >= 50.00 % and all other limits are held is 39.77 %.
Smallest total relative change that makes a blend possible (one way; there may be others): lower CaO_min from 50.00 % to 44.33 %.
```

How to read it:

- Each limit can be met on its own (Calcined alumina alone has 85 wt % Al2O3; Lime alone has 92 wt % CaO), so the pre-check finds nothing.
- But not both together. With CaO at least 50 wt %, the most Al2O3 you can get is **39.77 wt %**; with Al2O3 at least 45 wt %, the most CaO you can get is 44.33 wt %. These two lines are the same conflict seen from each side.
- The suggestion is to lower CaO_min to 44.33 %, because that is the smaller change relative to the limit (11.3 % of 50, against 11.6 % of 45 for Al2O3). Lowering Al2O3_min to 39.77 % works just as well.

### 8.4 What to do next

1. **Check your inputs first.** Most infeasible results are typing errors, not real conflicts.
   - Is a row missing, or in the wrong sheet? A raw material name belongs in UsageLimits, not Specs.
   - Is a value in the wrong column? (Example below.)
   - Is an analysis wrong? A misplaced decimal point or a value in the wrong column of RawMaterials can make a limit impossible.

   Real example of a value in the wrong column: Test 1 with the SiO2 limit of 6 typed under **Min** instead of Max. The tool now looks for at least 6 wt % SiO2 and reports a conflict:

```
Elastic model: these limits cannot all be met together: CaO_min, SiO2_min.
  - CaO_min (CaO >= 50.00 %): the highest CaO achievable while SiO2 >= 6.00 % and all other limits are held is 32.00 %.
  - SiO2_min (SiO2 >= 6.00 %): the highest SiO2 achievable while CaO >= 50.00 % and all other limits are held is 4.68 %.
```

   The name `SiO2_min` gives the mistake away: the customer gave a maximum. Al2O3_min is not in the list: CaO_min and SiO2_min conflict even without it. (The 32.00 % and 4.68 % are worked out with Al2O3_min held as well, as "all other limits are held" says.)

2. **Then decide how to resolve a real conflict.** There are three ways:
   - **Relax a limit with the customer.** The diagnosis tells you how far: in Example 2, to 39.77 wt % Al2O3 or to 44.33 wt % CaO.
   - **Add a raw material** that breaks the conflict: one that brings what is missing without what is in excess (in Example 2, a material richer in both CaO and Al2O3, such as a calcium aluminate).
   - **Check the analyses** of the materials involved, and update them if they were out of date.
3. **Test the options with WhatIf** before talking to the customer. Example 2 with two scenarios (real run):

| Scenario | Change type | Item | New value |
|---|---|---|---|
| CaO 44 | Spec Min | CaO | 44 |
| Al2O3 39.5 | Spec Min | Al2O3 | 39.5 |

| Scenario | Sheet | Changes | Status | Cost (Rs/dry t) | Change vs base (Rs/dry t) |
|---|---|---|---|---|---|
| CaO 44 | Scenario CaO 44 | Spec Min CaO 50.00 -> 44.00 | Optimal | 19,850 | |
| Al2O3 39.5 | Scenario Al2O3 39.5 | Spec Min Al2O3 45.00 -> 39.50 | Optimal | 18,494 | |

   Both are possible, and accepting 39.5 wt % Al2O3 gives the cheaper blend. "Change vs base" is blank because the base case has no cost.
4. **One conflict at a time.** The tool shows one set of conflicting limits. A badly specified product can contain several separate conflicts: resolve the one shown, run again, and the next one (if any) appears. A line ending in "even without this limit, the other limits cannot all be met" is the sign of a second conflict: without the limit named, other limits still clash. Example (real run): Test 1 with the Al2O3 minimum at 45 wt % **and** the SiO2 maximum at 1 wt %:

```
Elastic model: these limits cannot all be met together: SiO2_max.
  - SiO2_max (SiO2 <= 1.00 %): even without this limit, the other limits cannot all be met.
```

   SiO2 ≤ 1 wt % is impossible on its own (the pre-check says so too), and without it CaO_min and Al2O3_min still conflict, as in Example 2.

---

## 9. Everyday tasks

Short recipes for common jobs. Each one assumes the environment is active ([3.5](#35-activate-the-virtual-environment)) and that you run the tool as in [6.1](#61-the-command).

### 9.1 A raw material price changes

1. Look at the **Price ranges** sheet of the last report. If the new price is inside the material's band, the recipe stays the same and only the cost changes.
2. To see the new cost before changing anything, add a WhatIf row, for example:

| Scenario | Change type | Item | New value |
|---|---|---|---|
| New fluorspar price | Price | Fluorspar | 35000 |

   (In the example workbook this scenario is called "Fluorspar dearer": it adds ₹ 166 per dry tonne, and the recipe does not change.)
3. When the price is confirmed, change `Cost_Rs_per_t` in RawMaterials (in the master workbook for layout B), save, and run again.

### 9.2 A supplier offers a new material or a new price

- **A new price for a material you already list:** as in 9.1.
- **A new material:**
  1. Add a row to RawMaterials: name, landed cost, price basis, moisture, LOI, and a value in **every** component column.
  2. Optional: add a WhatIf scenario with `Usage Max` = 0 for the new material, called for example "Without new material". It shows what the new material saves.
  3. Run the tool.
  4. Look at the Recipe sheet. If the new material is used, the Price ranges sheet shows how far its price can rise before the recipe changes. If it is not used, the Reduced costs sheet shows the price below which it would be worth buying ("Enters below (Rs/t as entered)"). That is your answer to the supplier.

### 9.3 A customer asks for a tighter or looser limit

1. Look at the **Spec check** sheet. If the limit is slack, a small change costs nothing.
2. If it is binding, the **Shadow prices** sheet gives the cost per step, for a quick first answer.
3. For the exact figure, add a WhatIf row with `Spec Min` or `Spec Max`. In the example, "Looser P" (P Max 0.030 → 0.035) saves ₹ 623 per dry tonne.
4. If the customer agrees, change the row in Specs and run again.
5. A tighter limit can make the product impossible; the tool then explains why ([chapter 8](#8-when-there-is-no-recipe-infeasible)).

### 9.4 A customer bans a raw material

1. To see the effect first, add a WhatIf row with `Usage Max` = 0 for that material. In the example, "No fluorspar" is infeasible because nothing else contains CaF2.
2. To apply the ban for this customer's product, add a row to the product workbook's **UsageLimits** sheet with `Max_use_%` = 0. Other products are not affected.
3. If the product becomes infeasible, [chapter 8](#8-when-there-is-no-recipe-infeasible) shows which limits conflict and what to discuss with the customer.

### 9.5 Comparing the optimum with the current production recipe

Use the "current recipe" scenario described in [5.6](#56-whatif-optional): convert the shop-floor kg to dry wt %, fix every material except the largest, and set `Usage Max` = 0 for the materials not used today. The "Change vs base" column in the Scenarios sheet is the cost of today's recipe above the optimum, in ₹ per dry tonne (₹ 940 in the example). Multiply by the annual dry tonnes for the yearly figure.

### 9.6 Running several products against one master raw-material file

Use layout B ([chapter 4](#4-choosing-a-workbook-layout)). The tool runs one product per command, so run the command once for each product:

```
python run.py data\Slag_SS45.xlsx --raw-materials data\Raw_materials.xlsx
```

To run every product workbook in a folder in one go, put the product workbooks in `data\products\` and type this single line:

```
foreach ($p in Get-ChildItem data\products\*.xlsx) { python run.py $p.FullName --raw-materials data\Raw_materials.xlsx --output "output\$($p.BaseName)_$(Get-Date -Format yyyy-MM-dd).xlsx" --quiet }
```

Real output with two product workbooks, `Slag_dry.xlsx` and `Slag_ignited.xlsx`:

```
Report written to: output\Slag_dry_2026-10-26.xlsx
LP file:           output\Slag_dry_2026-10-26.lp

Report written to: output\Slag_ignited_2026-10-26.xlsx
LP file:           output\Slag_ignited_2026-10-26.lp
```

Each product gets its own dated report. Open each one and check its Spec check sheet.

### 9.7 Keeping a history of reports

With the default name, each run overwrites the previous report of the same product. To keep a history, put the date in the report name with `--output` ([6.3](#63-where-the-report-goes)):

```
python run.py templates\Test2_synthetic_slag_dry.xlsx --output "output\Test2_synthetic_slag_dry_$(Get-Date -Format yyyy-MM-dd).xlsx" --quiet
```

Keep a copy of the input workbook with the same date as well, so you can see later which prices and analyses gave which recipe.

---

## 10. Before using a recipe in production

The tool finds the cheapest recipe that meets the limits you gave it. It does not know your plant. Before you change production:

1. **Cost the current recipe first** ([5.6](#56-whatif-optional)), so you know the real saving.
2. **Check the optimum with an experienced eye.** Is any material at a share you would not normally use? Is one material close to 100 %? Is a customer limit missing from the Spec check sheet? Would the recipe cause handling, segregation, dust or mixing problems that the tool cannot know about?
3. **Add the missing practical limits** as usage limits: in UsageLimits for this product, or as `Min_use_%` / `Max_use_%` in RawMaterials for every product. Run again. The Shadow prices sheet then shows what each practical limit costs (rows of type "Usage limit").
4. **Use margins** on the limits that matter, to cover the lot-to-lot variation of the raw materials.
5. **Make a trial batch** with the As-received kg from the Recipe sheet.
6. **Analyse the trial batch** and compare it with the Composition sheet, on the **same basis** (dry, ignited or as received). If the analyses differ by more than the usual lab and sampling scatter, the analyses in RawMaterials are probably out of date: update them and run again.
7. Only then release the recipe.

---

## 11. Troubleshooting

The messages in this table are copied from real runs. Every other message the tool can produce is listed in [Appendix A](#appendix-a-message-reference).

| Symptom | Likely cause | What to do |
|---|---|---|
| `…\.venv\Scripts\Activate.ps1 cannot be loaded because running scripts is disabled on this system.` | PowerShell blocks scripts on this computer | Run `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process` and activate again, or use Command Prompt ([3.5](#35-activate-the-virtual-environment)) |
| `ModuleNotFoundError: No module named 'pandas'` (or another package) | The environment is not active, so the system Python runs, without the packages | Activate ([3.5](#35-activate-the-virtual-environment)). If it is active, install the packages ([3.6](#36-install-the-packages)). |
| `python: can't open file 'D:\\VS Python\\Cost Optimizer\\rn.py': [Errno 2] No such file or directory` | A typo in `run.py`, or PowerShell is in another folder | Use Tab to complete names; `cd` to the tool's folder ([3.3](#33-open-powershell-in-the-tools-folder)) |
| `File not found: templates\Test1_hand_workd.xlsx` | A typo in the workbook name or folder | Use Tab to complete the name |
| `Nothing was solved. Correct the workbook and run again.` | The workbook has errors | Read the messages printed **above** this line; each names the sheet, row and item ([6.5](#65-reading-the-screen)) |
| `component 'AL2O3' is not a column in RawMaterials. Did you mean 'Al2O3'? (Component names must match exactly.)` | The name in Specs is spelled differently from the RawMaterials header | Make the spelling identical, including capital letters |
| `Specs (Test1_formula_in_Name.xlsx) row 5 'CaO / SiO2', Numerator: value is blank; name at least one component.` | The formula was written in **Name** instead of Numerator and Denominator | Name is only a label: write `CaO` in Numerator and `SiO2` in Denominator |
| `Numerator: 'Fluorspar' is a raw material, not a component. To limit how much of it is used, put it in UsageLimits.` | A **raw material** name was used in Specs (e.g. to ban it) | Limits on raw materials belong in UsageLimits ([5.5](#55-usagelimits-optional)); Specs only takes chemical components |
| `Numerator: 'LOI_%' is a property column, not a component, so it can't be used in a spec.` | A **property** column (`LOI_%`, `Moisture_%`, the cost or usage columns) was used in a spec | These columns describe the raw materials; they cannot be limited in Specs ([12. Limitations](#12-limitations)) |
| `Could not write the report: [Errno 13] Permission denied: …` and `If the report is open in Excel, close it and run again.` | The report from an earlier run is open in Excel | Close it in Excel and run again |
| `No 'RawMaterials' sheet in Test2_product_only.xlsx, and no master raw-material file was given. Add the sheet, or give the master raw-material file with --raw-materials.` | A layout B product was run without its master file | Add `--raw-materials data\Raw_materials.xlsx` ([6.2](#62-options)) |
| A warning `components + LOI = … %, outside the expected 97-103 %` | A typing error, a missing component column, a missing LOI, or analyses on the other convention | Check the material's analysis ([5.2](#52-rawmaterials)) |
| A warning `RawMaterials 'Calcined bauxite': P is blank (not analysed) and is counted as 0, but a spec uses it. Type 0 if it is genuinely absent.` | An analysis cell left **blank** for a component that a spec limits. The recipe would rely on a value nobody measured (here the report shows 0.024 wt % P; with the real analysis the same recipe has 0.032 wt %). | Fill in the analysed value, or type 0 if the component is genuinely absent ([5.2](#52-rawmaterials)) |
| An unexpected recipe, e.g. 100 % of one material | A **missing spec row** (Test 1 without its Al2O3 row gives 100 % Lime) | Compare the Spec check sheet with the customer's specification, limit by limit ([7.4](#74-spec-check)) |
| An unexpected recipe or `Infeasible`, and a limit appears as `…_min` where the customer gave a maximum (e.g. `SiO2_min`) | A value typed in the **Min column instead of Max** | Move the value to the right column |
| An unexpected recipe with no warning | A **wrong value** in an analysis cell, e.g. a 0 typed where the component is present, or a misplaced decimal point | Check the analyses of the materials the recipe uses, especially trace elements ([5.2](#52-rawmaterials)) |
| Status `Infeasible` (exit code 2) | No mix meets every limit | Read the diagnosis ([chapter 8](#8-when-there-is-no-recipe-infeasible)) |
| A batch file or script sees exit code 1 or 2 | 1: input errors (or the report could not be written); 2: no recipe for the base case | Run the same command by hand and read the screen ([6.5](#65-reading-the-screen)) |
| `The minimum usage limits add up to 100.01 % of the blend, more than 100 %` in a "current recipe" scenario | All materials were fixed at rounded shares | Fix every material except one ([5.6](#56-whatif-optional)) |
| A raw material at 0.00 % | It is not worth using at its price | Normal ([2.3](#23-why-the-cheapest-recipe-uses-only-a-few-raw-materials)); see Reduced costs for the price at which it would be used |
| Reduced costs says "Enters below 1,872" and Price ranges says 1,873 | The price search works to ₹ 1 | Normal |
| A red block with `NativeCommandError` and `[notice] A new release of pip is available` during installation | A pip notice, shown in red by PowerShell | Ignore it |
| An error that starts with `WhatIf scenario '…':` | The changes of one scenario do not fit together (e.g. a new Min above the Max) | Correct that scenario's rows ([Appendix A](#appendix-a-message-reference), A.7) |

---

## 12. Limitations

What this version of the tool does not do:

- **Continuous amounts only.** Shares can take any value. There are no whole bags, no "either 0 or at least 5 %" (a `Min_use_%` always applies), and no limit on the number of raw materials used.
- **Chemistry only.** Limits are on components, sums of components, ratios and conversion factors. Melting point, viscosity, particle size, density, dust and flowability are not modelled.
- **LOI and moisture cannot be limited** in Specs; they are properties of the raw materials, not components.
- **Usage limits are per raw material.** There is no limit on a group of materials together (e.g. "Raw bauxite plus Lime fines at most 30 %").
- **Raw material cost only.** The cost is the landed cost of the raw materials. Processing, energy, losses and yield are not included.
- **One conflict set at a time** in the diagnosis ([8.4](#84-what-to-do-next)).
- **Conversion factors are typed by you.** The tool knows no chemistry and does not check that a factor is right. Subtraction (`-`) is not allowed in expressions.
- **One lab convention for the whole tool**, set by the maintainer, not per material or per workbook.
- **Shadow prices are valid only near the current limit**; use WhatIf for the exact effect of a change.
- **Price ranges** vary one price at a time with all others fixed, are searched between ₹ 0 and 10 times today's price, and are accurate to ₹ 1.
- **Scenarios** are each compared with the base case, not with each other, and cannot add a new raw material (add it in RawMaterials).
- **One product per run** (a PowerShell loop can run several, [9.6](#96-running-several-products-against-one-master-raw-material-file)); report names carry no date unless you give one.
- **Analyses are taken as exact.** The tool does not model the variation of raw materials; use margins.

---

## 13. Glossary

| Term | Meaning |
|---|---|
| As received | The state of a material as delivered, including its free moisture. Prices and the batch weight are usually as received. |
| Base case | The workbook as it stands, without any WhatIf changes. Every scenario is compared with it. |
| Binding | A limit the blend sits exactly on. It holds the recipe where it is; relaxing it would make the blend cheaper. |
| Conversion factor | A number that counts one component as another, e.g. 1.8895 to count metallic Al as Al2O3 (`Al2O3+1.8895*Al`). |
| Dry basis | The analysis or mass of a material with its free moisture removed (LOI still included). Raw material analyses are entered on a dry basis. |
| Elastic model | The second step of the diagnosis: every limit may stretch, and the tool finds the smallest total stretch that makes a recipe possible. It shows which limits conflict. |
| Feasible / infeasible | Feasible: at least one recipe meets every limit. Infeasible: none does. |
| Ignited basis | The analysis after loss on ignition, i.e. without the volatile part. Components are higher than on a dry basis. |
| Landed cost | The price of a raw material delivered to your plant: purchase price plus freight and handling. |
| Linear programming | The mathematical method the tool uses. It finds the exact cheapest recipe when the blend's analysis and cost are weighted averages of the raw materials'. |
| LOI | Loss on ignition: the mass lost when the dried material is heated strongly (water of crystallisation, CO2 from carbonates, etc.), in wt % of the dry material. |
| .lp file | A plain-text copy of the optimisation problem, written next to each report, for checking by hand. |
| Margin | A safety distance you set inside a customer limit. The tool aims for Min + margin and Max − margin. A margin only ever tightens a limit. |
| Price range | The band within which a raw material's price can move without changing the recipe. |
| Reduced cost | For an unused raw material: how much cheaper it would have to be before it is worth using. |
| Scenario | A named set of changes in the WhatIf sheet, solved and compared with the base case. |
| Shadow price | For a binding limit: how much the blend cost falls (₹ per dry tonne) per step that the limit is relaxed. Valid only near the current limit. |
| Slack | A limit the blend meets with room to spare. Relaxing it changes nothing. |
| Solver | The program that carries out the linear programming. This tool uses HiGHS. |
| Spec basis | The basis of the customer's composition limits: dry, as received or ignited. Set in Settings. |
| Target | The limit the tool actually aims for: the customer limit with your margin applied. |
| Usage limit | A minimum or maximum share of one raw material in the dry blend: plant-wide (RawMaterials) or for one product (UsageLimits). |
| Virtual environment | The private copy of Python and its packages in the `.venv` folder, used only by this tool. |
| wt % | Percent by mass. |

---

## Appendix A. Message reference

These are the messages the tool can print, taken from the program. Parts in `<angle brackets>` are filled in with the actual file, row, name or value. Errors stop the run ("Nothing was solved"); warnings let it continue.

Most messages start with where the problem is, for example `RawMaterials (Slag_SS45.xlsx) row 4 'Lime', Cost_Rs_per_t:`, meaning sheet, file, Excel row, the row's name, and the column.

### A.1 Starting and finishing

| Message | Type | Meaning |
|---|---|---|
| `python: can't open file '<path>\<name>.py': [Errno 2] No such file or directory` | Python error | The script name is wrong, or you are in another folder |
| `ModuleNotFoundError: No module named '<package>'` | Python error | The environment is not active or not installed |
| `File not found: <path>` | Error | The workbook does not exist at that path |
| `Could not read <file> as an Excel workbook: <reason>` | Error | The file is damaged, not an .xlsx file, or open elsewhere |
| `<file>: required sheet '<sheet>' not found. Did you mean '<sheet>'? Sheets present: <list>.` | Error | A required sheet (RawMaterials, Specs or Settings) is missing or misnamed |
| `No 'RawMaterials' sheet in <file>, and no master raw-material file was given. Add the sheet, or give the master raw-material file with --raw-materials.` | Error | Add `--raw-materials <master file>` |
| `<file> also has a RawMaterials sheet. It is ignored because a master raw-material file (<master>) was given.` | Warning | Layout B: the master file is used |
| `<n> input error(s) found:` … `Nothing was solved. Correct the workbook and run again.` | Summary | The list of errors and warnings, then the run stops |
| `Could not write the report: <reason>` / `If the report is open in Excel, close it and run again.` | Error | The report file is locked |
| `Report written to: <path>` / `LP file: <path>` / `Scenarios: <n> (see the Scenarios sheet)` | Information | Where the results are |

Every run also ends with an exit code ([6.5](#65-reading-the-screen)): 0 = recipe found, 1 = input errors (or the report could not be written), 2 = no recipe for the base case.

### A.2 Headers and columns (any sheet)

| Message | Type | Meaning |
|---|---|---|
| `<sheet>: required column '<column>' is missing. Did you mean '<header>'?` | Error | A required column header is missing or misspelled |
| `<sheet>: column '<header>' looks like a misspelling of '<column>'. Rename it; otherwise it would be read as a chemical component.` (in other sheets: `… otherwise it would be ignored.`) | Error | Correct the header |
| `<sheet>: column '<header>' is not used and is ignored.` | Warning | An extra column in a sheet other than RawMaterials |
| `<sheet>: column <letter> has values but no header in row 1.` | Error | Add the header, or clear the column |
| `<sheet>: column '<header>' appears more than once (column <letter>).` | Error | Two columns with the same header |

### A.3 RawMaterials

| Message | Type | Meaning |
|---|---|---|
| `<sheet> row <n>: Name is blank.` | Error | Every row needs a name |
| `… duplicate name (already used in row <n>; names must be unique, ignoring case).` | Error | Two raw materials with the same name |
| `…, <column>: value is blank.` | Error | A required value (e.g. the cost) is missing |
| `…, <column>: '<text>' is not a number. Use a number, or leave the cell blank.` | Error | Text in a number column, e.g. `8,500 Rs` |
| `…, <column>: <value> is negative; it must be 0 or more.` | Error | |
| `…, <column>: <value> is above the maximum allowed (100).` | Error | A percentage above 100 |
| `…, <column>: must be below 100.` | Error | Moisture or LOI of 100 % or more |
| `…, Price_basis: '<text>' is not allowed; use one of as_received, dry.` | Error | |
| `…, Price_basis: blank, but Moisture_% is <value>. Say whether Cost_Rs_per_t is 'as_received' or 'dry'.` | Error | With moisture, the basis matters |
| `…: price is 'as_received' but Moisture_% is blank, so moisture is taken as 0 and the dry price equals the quoted price.` | Warning | Fill in the moisture |
| `…: Min_use_% (<a>) is greater than Max_use_% (<b>).` | Error | |
| `<sheet>: no chemical component columns found (e.g. one column per oxide).` | Error | |
| `<sheet>: no raw materials listed.` | Error | |
| `RawMaterials '<name>': components + LOI = <x> %, outside the expected 97-103 % for the 'dry_incl_loi' lab convention. Components alone = <y> %, which looks like the 'loi_free' convention. Check the lab report and LAB_ANALYSIS_CONVENTION in config.py.` | Warning | The analysis looks LOI-free ([5.2](#52-rawmaterials)) |
| `RawMaterials '<name>': … Neither convention fits (components alone = <y> %, components + LOI = <x> %): check for a missing component column, a typing error or a missing LOI.` | Warning | The analysis does not add up |
| `RawMaterials '<name>': <component> is blank (not analysed) and is counted as 0, but a spec uses it. Type 0 if it is genuinely absent.` | Warning | A blank analysis cell in a component that a spec uses, for a material the blend may use (effective Max above 0). Fill in the value, or type 0 ([5.2](#52-rawmaterials)) |

### A.4 Specs

| Message | Type | Meaning |
|---|---|---|
| `<sheet> row <n>: Name is blank.` | Error | |
| `… duplicate name (already used in row <n>; spec names become constraint names, so they must be unique).` | Error | |
| `…, Numerator: value is blank; name at least one component.` | Error | Often: the formula was written in Name |
| `…, <column>: component '<name>' is not a column in RawMaterials. Did you mean '<name>'? (Component names must match exactly.)` | Error | A spelling difference |
| `…, <column>: '<name>' is a property column, not a component, so it can't be used in a spec.` | Error | `LOI_%`, `Moisture_%`, `Cost_Rs_per_t`, `Min_use_%`, `Max_use_%` (or the same with other capitals, spaces or without `_%`) used in Numerator or Denominator |
| `…, <column>: '<name>' is a raw material, not a component. To limit how much of it is used, put it in UsageLimits.` | Error | A raw material name (in any capitals) used in Numerator or Denominator ([5.5](#55-usagelimits-optional)) |
| `…, <column>: component '<name>' is not a column in RawMaterials. Components available: <list>.` | Error | None of the above, and no close spelling: a typo, or a component column missing in RawMaterials |
| `…: give a Min, a Max, or both.` | Error | |
| `…: Min (<a>) is greater than Max (<b>).` | Error | |
| `…: after the margin the internal range is empty: Min + margin = <a> is greater than Max - margin = <b>. Reduce the margin.` | Error | |
| `…: Max - margin = <x> is negative. Reduce the margin.` | Error | |
| `…: Min + margin = <x> % is above 100 % and can never be met.` | Error | Only for plain components (not with a factor above 1) |
| `<sheet>: no spec rows; the cheapest raw material mix will be chosen without limits.` | Warning | |

Problems in a Numerator or Denominator expression (each preceded by the sheet, row and column):

| Message | Meaning |
|---|---|
| `the expression is empty` | |
| `'<expr>' has an empty component name: check for '++', or a '+' at the start or end` | |
| `'<expr>': subtraction ('-') is not supported; write every term with a positive factor joined by '+'` | |
| `in '<expr>', the term '<term>' starts with a number; put '*' between the factor and the component (did you mean <n>*<name>?)` | e.g. `1.8895Al` |
| `in '<expr>', the term '<term>' has more than one '*'; write one factor per component, e.g. 1.8895*Al` | |
| `in '<expr>', the term '<term>' has no factor before '*', e.g. 1.8895*Al` | |
| `in '<expr>', the term '<term>' has no component after '*'` | |
| `in '<expr>', the term '<term>' has the factor after the component; write it first: <n>*<name>` | e.g. `Al*1.8895` |
| `in '<expr>', the factor '<text>' is not a number` (with a comma: `(use a decimal point, e.g. 1.8895)`) | |
| `in '<expr>', the factor of <name> is 0; it must be greater than 0` | |
| `'<expr>' lists <name> more than once; write each component once, with the combined factor` | |

### A.5 Settings

| Message | Type | Meaning |
|---|---|---|
| `<sheet> row <n>: the row has values but Key is blank.` | Error | |
| `<sheet> row <n>: unknown setting '<key>' is ignored. Did you mean '<key>'?` | Warning | |
| `<sheet> row <n>: setting '<key>' appears more than once.` | Error | |
| `<sheet>: Product_name not given; using '<file name>'.` | Warning | |
| `… 'Spec_basis': '<text>' is not allowed; use one of dry, as_received, ignited.` | Error | |
| `… '<key>': must be greater than 0.` | Error | Batch_size_kg or Scale_resolution_kg |

### A.6 UsageLimits

| Message | Type | Meaning |
|---|---|---|
| `<sheet>: needs a Min_use_% and/or Max_use_% column.` | Error | |
| `<sheet> row <n>: Name is blank.` | Error | |
| `… '<name>': not a raw material in RawMaterials. Did you mean '<name>'?` | Error | |
| `… '<name>': listed more than once.` | Error | |
| `… Min_use_% (<a>) is greater than Max_use_% (<b>).` | Error | |
| `… the effective limits conflict. Effective Min = <a> % (plant <p>, product <q>) is greater than effective Max = <b> % (plant <p>, product <q>).` | Error | Plant and product limits cannot both be met |

### A.7 WhatIf

Checked when the workbook is read:

| Message | Type |
|---|---|
| `… Scenario is blank; give the scenario a name.` | Error |
| `… Change type '<text>' is not known. Did you mean '<type>'? Use one of: Price, Spec Min, Spec Max, Spec Margin, Usage Min, Usage Max.` | Error |
| `… Item is blank; give the <raw material / spec> to change.` | Error |
| `… '<item>' is not a raw material in RawMaterials.` / `… is not a spec in Specs.` (with a suggestion) | Error |
| `… '<type>' for '<item>' is already given in row <n> of this scenario.` | Error |
| `…, New value: value is blank.` | Error (Price needs a value) |

Checked when the run starts, for each scenario as a whole (`WhatIf scenario '<name>': …`):

| Message | Meaning |
|---|---|
| `WhatIf scenario '<name>': Spec '<spec>': <a spec problem from A.4>` | e.g. a new Min above the Max |
| `WhatIf scenario '<name>': Usage limit for '<material>': need 0 <= Min <= Max <= 100, got Min <a>, Max <b>.` | |
| `WhatIf scenario '<name>': Usage limit for '<material>' conflicts with the plant-wide limit: effective Min <a> % > effective Max <b> %.` | |

### A.8 Diagnosis (no recipe)

| Message | Meaning |
|---|---|
| `No recipe: the solver status is 'Infeasible'.` | No mix meets every limit |
| `Pre-check: every limit can be met on its own, so some limits conflict with each other.` | Look at the elastic model's list |
| `The minimum usage limits add up to <x> % of the blend, more than 100 % (<list>).` | Usage minimums leave no room |
| `The maximum usage limits add up to only <x> %, so the allowed raw materials cannot make up a whole blend (<list>).` | Usage maximums are too tight |
| `No allowed raw material contains <component>, so <limit> (<condition>) cannot be met.` | |
| `<limit> (<condition>) cannot be met by any mix of the allowed raw materials, even ignoring all other spec limits: the highest (or lowest) achievable <what> is <value>.` | |
| `Raw materials that would help are restricted: <material> (<content>) is not allowed (Max 0 %, <source>); …` / `… is capped at <x> % (<source>)` | `<source>` is `UsageLimits`, `plant-wide limit in RawMaterials`, or `WhatIf (<scenario>)` |
| `Elastic model: these limits cannot all be met together: <list>.` | The conflicting set, kept to the limits really needed: each one was tested by taking it away ([8.1](#81-how-the-tool-explains-it)) |
| `<limit> (<condition>): the highest (or lowest) <what> achievable while <other limits> and all other limits are held is <value>.` | The best possible value for this limit |
| `<limit> (<condition>): even without this limit, the other limits cannot all be met.` | There is a second, separate conflict among other limits; resolve the one shown and run again ([8.4](#84-what-to-do-next)) |
| `Smallest total relative change that makes a blend possible (one way; there may be others): lower (or raise) <limit> from <a> to <b>.` | One way out |
| `Elastic model: no answer (<status>). The usage limits themselves leave no possible blend; see the pre-check.` | Correct the usage limits first |

### A.9 Price ranges

Under the table on screen (`Notes:`, one line per raw material) and in the Note column of the Price ranges sheet ([7.9](#79-price-ranges)):

| Message | Meaning |
|---|---|
| `Recipe unchanged at every price tested (Rs 0 to <price>)` | The material's share is fixed by the limits, whatever it costs (in the range tested) |
| `No lower limit: unchanged down to Rs 0` | The recipe stays the same however cheap the material gets |
| `No upper limit found up to Rs <price> (<n>x the current price)` | The recipe stays the same up to the highest price tested |
| `Not used now; enters the recipe below Rs <price>` | The material is at 0 %; below this price it would be used |
| `Not available: price ranges need a feasible base case (see the Diagnosis sheet).` | No recipe exists, so there is no price range |

---

## Appendix B. For the maintainer

This appendix is for whoever looks after the tool. It contains the only programming detail in this guide.

**Project structure**

| Path | Responsibility |
|---|---|
| `run.py` | Command-line entry point: runs the base case, the price ranges and the WhatIf scenarios, writes one report |
| `blend_optimizer\config.py` | All constants: sheet and column names, lab convention, display decimals, tolerances |
| `blend_optimizer\data.py` | Data classes passed between modules (raw materials, spec rows, terms, scenarios) |
| `blend_optimizer\inputs.py` | Reads and validates the workbooks; `parse_expression` is the one reader of Numerator/Denominator expressions |
| `blend_optimizer\linear_rows.py` | Turns each spec limit into one linear row (basis conversion, ratios, factors) |
| `blend_optimizer\model.py` | Builds the PuLP problem (normal and elastic) |
| `blend_optimizer\solve_and_report.py` | Solves it (`make_solver` is the only place the solver is named), reads the results, builds the report tables, writes Excel and the console summary |
| `blend_optimizer\pipeline.py` | Solve → diagnose if infeasible (including the reduction of the conflict set to the limits really needed) → report |
| `blend_optimizer\rounding.py` | The one rounding rule for display (half up), used on screen and in Excel |
| `blend_optimizer\diagnose.py` | Pre-checks and the plain-language diagnosis (no solver) |
| `blend_optimizer\what_if.py` | What-if scenarios, the WhatIf sheet runs, price ranging |
| `blend_optimizer\templates.py`, `examples.py`, `make_templates.py` | Template writer, example datasets, template generator |
| `tests\` | The test suite |
| `docs\User_Guide.md`, `docs\build_html.py` | This guide and its HTML converter |

**PuLP is used only in `model.py` and `solve_and_report.py`.** A test fails if any other module imports it.

**Settings in `config.py`** that you may need to change:

| Constant | Default | What it controls |
|---|---|---|
| `LAB_ANALYSIS_CONVENTION` | `dry_incl_loi` | How analyses are read; `loi_free` converts LOI-free analyses with c_dry = c × (1 − LOI/100) |
| `ANALYSIS_TOTAL_WARN_RANGE` | (97.0, 103.0) | The analysis total check |
| `DISPLAY_DECIMALS_PCT`, `DISPLAY_DECIMALS_KG`, `DISPLAY_DECIMALS_RS` | 2, 1, 0 | Display rounding |
| `SCALE_RESOLUTION_KG` | 0.1 | Default scale resolution (Settings can override it) |
| `BINDING_TOLERANCE` | 1e-6 | Slack below this counts as binding |
| `ELASTIC_MIN_LIMIT` | 0.01 | The smallest limit size used to weight violations in the elastic model |

**Running the tests** (with the environment active):

```
python -m pytest
```

At the time of writing: `273 passed`.

**Regenerating the templates** (after changing `examples.py` or `templates.py`):

```
python make_templates.py
```

This rewrites the 13 workbooks in `templates\`.

**Rebuilding the HTML version of this guide** (after editing `docs\User_Guide.md`):

```
python docs\build_html.py
```

**PuLP 4.x notes.** PuLP 4 changed its API, and most examples online are for 2.x/3.x:

- Variables are created with `prob.add_variable(name, lowBound=…, upBound=…)`; `LpVariable(...)` raises `TypeError`.
- `prob.constraints()` returns a list; `prob.get_constraint_by_name(name)` finds one constraint.
- `prob.solve(solver)` returns an `LpSolveStats` object: use `.status` and `.has_solution`.
- CBC is not bundled and `LpSolverDefault` is `None`, so the solver is always passed explicitly (`HiGHS(msg=False)` in `make_solver`).
- After a HiGHS solve, `constraint.pi` holds the dual (d cost / d right-hand side) and `variable.dj` the reduced cost; both are populated for LPs.
- Variables only work while their `LpProblem` exists, so results are copied into a plain `SolveResult` immediately after solving.
- PuLP changes names (spaces become `_`), so the model keeps its own name maps instead of looking things up by PuLP name.

**Price ranging from Python**: `blend_optimizer.what_if.price_ranging(load_inputs(...), tolerance_rs=1.0, search_factor=10.0)`.

**Exit codes** are set in `run.py` (`EXIT_RECIPE_FOUND`, `EXIT_INPUT_ERRORS`, `EXIT_NO_RECIPE`).

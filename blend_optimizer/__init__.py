"""Least-cost blend optimizer.

Finds the cheapest mix of raw materials that meets a composition specification.
The modules are used in this order:

    inputs            read and validate the Excel workbook(s)        -> BlendData
    linear_rows       turn each spec limit into a linear row          (pure Python)
    model             build the PuLP problem (normal or elastic)      (PuLP)
    solve_and_report  solve it and build the report tables            (PuLP)
    diagnose          pre-checks and plain-language diagnosis         (pure Python)
    pipeline          solve -> diagnose if infeasible -> report
    what_if           what-if scenarios (WhatIf sheet), price ranging

``run.py`` in the project folder is the command-line entry point.
Only ``model`` and ``solve_and_report`` import PuLP.
"""

__version__ = "0.1.0"

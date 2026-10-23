"""Shared pytest fixtures.

Every test writes its own workbooks into pytest's temporary folder
(``tmp_path``), built from the datasets in ``blend_optimizer/examples.py``.
So the tests never depend on files in templates/ and never overwrite them.
"""

from pathlib import Path
from typing import Callable

import pytest

from blend_optimizer.examples import ExampleWorkbook


@pytest.fixture
def write_example(tmp_path: Path) -> Callable[..., Path]:
    """Return a function that writes an ExampleWorkbook into tmp_path and returns the file path."""

    def _write(example: ExampleWorkbook, file_name: str = "product.xlsx") -> Path:
        return example.write(tmp_path / file_name)

    return _write

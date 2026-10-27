"""The one rounding rule for everything the tool shows: round half up.

Python's built-in ``round()`` and ``f"{x:.2f}"`` round a tie to the even
digit, and work on the binary value, so 78.895 (stored as 78.894999...)
shows as 78.89 and 12,744.5 as 12,744. Excel shows 78.90 and 12,745. To keep
the screen, the Excel report and the messages identical, every displayed
number goes through :func:`round_half_up`.

Calculations are never rounded; this is for display only. (The weigh sheet
stores its kg already rounded to the scale, with this same rule.)
"""

from decimal import ROUND_HALF_UP, Decimal


def round_half_up(value: float, step: float = 1.0) -> float:
    """Round ``value`` to the nearest multiple of ``step``; a tie goes away from zero.

    The work is done in decimal on the numbers as written (their shortest
    form, e.g. 78.895), so a value that looks like a tie is treated as one.

    Examples: (78.895, 0.01) -> 78.90; (12744.5, 1) -> 12745;
    (12.25, 0.5) -> 12.5; (0.15, 0.1) -> 0.2; (-2.5, 1) -> -3.

    Args:
        value: The number to round.
        step: The rounding step: 0.01 for 2 decimals, 1 for whole numbers,
            or a scale resolution such as 0.5 kg.
    """
    exact_step = Decimal(repr(float(step)))
    steps = (Decimal(repr(float(value))) / exact_step).quantize(Decimal(1), rounding=ROUND_HALF_UP)
    return float(steps * exact_step)


def format_fixed(value: float, decimals: int, thousands: bool = False) -> str:
    """Text with a fixed number of decimals, rounded half up (e.g. 12,745 or 78.90).

    Args:
        value: The number.
        decimals: Decimals to show.
        thousands: Group thousands with commas.
    """
    rounded = round_half_up(value, 10.0 ** -decimals)
    if rounded == 0:
        rounded = 0.0  # never show "-0.00"
    separator = "," if thousands else ""
    return f"{rounded:{separator}.{decimals}f}"

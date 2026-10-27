"""Plain data classes passed between the modules.

These hold already-validated input data. They know nothing about Excel
(that is ``inputs``) and nothing about PuLP (that is ``model``).

Units used everywhere:
    compositions   wt %, dry basis
    usage limits   wt % of the dry blend
    costs          Rs per tonne (dry_cost: Rs per DRY tonne)
"""

from dataclasses import dataclass, field

from blend_optimizer import config


# ---------------------------------------------------------------------------
# Price conversions
# ---------------------------------------------------------------------------
# These live next to RawMaterial (rather than in inputs) so that
# RawMaterial.dry_cost is always computed from the current price. When
# what_if changes a price, the dry cost can never be left stale.

def price_to_dry_basis(cost: float, price_basis: str, moisture_pct: float) -> float:
    """Convert a quoted price to Rs per dry tonne.

    One tonne as received contains (1 - moisture/100) tonnes of dry material,
    so paying ``cost`` for it means paying ``cost / (1 - moisture/100)`` per
    dry tonne. A price already on a dry basis is returned unchanged.

    Args:
        cost: Price in Rs per tonne, on ``price_basis``.
        price_basis: ``"as_received"`` or ``"dry"``.
        moisture_pct: Free moisture, wt % as received (0 <= moisture < 100).

    Returns:
        Price in Rs per dry tonne.
    """
    if price_basis == config.PRICE_DRY:
        return cost
    if price_basis == config.PRICE_AS_RECEIVED:
        return cost / (1.0 - moisture_pct / 100.0)
    raise ValueError(f"Unknown price basis {price_basis!r}; expected one of {config.PRICE_BASES}")


def price_from_dry_basis(dry_cost: float, price_basis: str, moisture_pct: float) -> float:
    """Inverse of :func:`price_to_dry_basis`: Rs per dry tonne -> price on ``price_basis``.

    Used to report results (e.g. price ranges) in the same basis the user typed.
    """
    if price_basis == config.PRICE_DRY:
        return dry_cost
    if price_basis == config.PRICE_AS_RECEIVED:
        return dry_cost * (1.0 - moisture_pct / 100.0)
    raise ValueError(f"Unknown price basis {price_basis!r}; expected one of {config.PRICE_BASES}")


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class UsageLimit:
    """Allowed share of one raw material in the dry blend, in wt %."""

    min_pct: float = 0.0
    max_pct: float = 100.0

    def tighter(self, other: "UsageLimit") -> "UsageLimit":
        """Combine two limits, keeping the stricter bound on each side."""
        return UsageLimit(
            min_pct=max(self.min_pct, other.min_pct),
            max_pct=min(self.max_pct, other.max_pct),
        )


@dataclass(frozen=True)
class RawMaterial:
    """One raw material, as read from the RawMaterials sheet.

    Attributes:
        name: Unique identifier (the Name column).
        cost_input: Price exactly as typed, Rs per tonne on ``price_basis``.
        price_basis: ``"as_received"`` or ``"dry"``.
        moisture_pct: Free moisture, wt % as received.
        loi_pct: Loss on ignition, wt % dry basis.
        plant_limit: Plant-wide usage limit (Min_use_% / Max_use_% columns).
        composition: Component -> wt %, dry basis (already converted if the
            lab reports on an LOI-free basis). Contains every component
            column of the sheet; blanks are 0.
        not_analysed: Components whose cell was left blank ("not analysed"),
            as opposed to typed as 0 ("analysed, none present"). Both count
            as 0 in the calculation; a blank in a component that a spec uses
            gives a warning.
    """

    name: str
    cost_input: float
    price_basis: str
    moisture_pct: float
    loi_pct: float
    plant_limit: UsageLimit
    composition: dict[str, float]
    not_analysed: frozenset[str] = frozenset()

    @property
    def dry_cost(self) -> float:
        """Price in Rs per dry tonne (what the optimizer minimises)."""
        return price_to_dry_basis(self.cost_input, self.price_basis, self.moisture_pct)


@dataclass(frozen=True)
class Term:
    """One term of a Numerator or Denominator expression: factor x component.

    "Al2O3+1.8895*Al" has two terms: (1, Al2O3) and (1.8895, Al). The
    factor is a number typed by the user; the tool knows no chemistry.

    Attributes:
        factor: Multiplies the component's wt % (1 for a plain component).
        component: A component column of RawMaterials.
        written: The term as typed (spaces removed), shown in reports.
            Not part of equality: "1*Al" and "Al" are the same term.
    """

    factor: float
    component: str
    written: str = field(default="", compare=False)

    @property
    def label(self) -> str:
        """The term as shown in reports: as typed, or built from factor and component."""
        if self.written:
            return self.written
        if self.factor == 1:
            return self.component
        return f"{self.factor:g}{config.FACTOR_SEPARATOR}{self.component}"


@dataclass(frozen=True)
class SpecRow:
    """One row of the Specs sheet.

    A row with an empty ``denominator`` is a composition limit on the
    numerator expression (a sum of factor x component terms). A row with a
    denominator is a ratio, e.g. numerator (CaO) and denominator (SiO2) for
    basicity.

    Attributes:
        name: Unique label; constraints are named ``<name>_min`` / ``<name>_max``.
        numerator: Terms of the numerator expression.
        denominator: Terms of the denominator expression; ``()`` if none.
        min_value: Customer minimum, or None for no minimum.
        max_value: Customer maximum, or None for no maximum.
        margin: Safety margin, same unit as the limits, applied inward.
    """

    name: str
    numerator: tuple[Term, ...]
    denominator: tuple[Term, ...] = ()
    min_value: float | None = None
    max_value: float | None = None
    margin: float = 0.0

    def __post_init__(self) -> None:
        # Plain component names are accepted too and mean factor 1, which is
        # handy in code and tests: SpecRow("CaO", ("CaO",), ...). They are
        # turned into Terms here, so everything else only ever sees Terms.
        # (The dataclass is frozen, so the converted value is set this way, once.)
        for side in ("numerator", "denominator"):
            terms = tuple(item if isinstance(item, Term) else Term(1.0, str(item), str(item))
                          for item in getattr(self, side))
            object.__setattr__(self, side, terms)

    @property
    def is_ratio(self) -> bool:
        """True if this row limits a ratio rather than a composition."""
        return len(self.denominator) > 0

    @property
    def min_target(self) -> float | None:
        """Internal minimum used by the optimizer: customer Min + margin."""
        return None if self.min_value is None else self.min_value + self.margin

    @property
    def max_target(self) -> float | None:
        """Internal maximum used by the optimizer: customer Max - margin."""
        return None if self.max_value is None else self.max_value - self.margin

    @property
    def expression(self) -> str:
        """What is limited, as typed: e.g. ``"Al2O3+1.8895*Al"`` or ``"(CaO+MgO) / SiO2"``."""
        def joined(terms: tuple[Term, ...], brackets: bool) -> str:
            text = config.COMPONENT_SEPARATOR.join(term.label for term in terms)
            return f"({text})" if brackets and len(terms) > 1 else text

        if not self.is_ratio:
            return joined(self.numerator, brackets=False)
        return f"{joined(self.numerator, True)} / {joined(self.denominator, True)}"


@dataclass(frozen=True)
class Settings:
    """Values from the Settings sheet (defaults are in ``config``)."""

    product_name: str
    spec_basis: str = config.DEFAULT_SPEC_BASIS
    # Total AS-RECEIVED mass weighed into the mixer, kg
    batch_size_kg: float = config.DEFAULT_BATCH_SIZE_KG
    # Weighing-scale resolution used to round kg in the report
    scale_resolution_kg: float = config.SCALE_RESOLUTION_KG


@dataclass(frozen=True)
class Scenario:
    """One named what-if scenario from the WhatIf sheet.

    The change dictionaries have the format of ``what_if.apply_changes``.

    Attributes:
        name: Scenario name, as written in the WhatIf sheet.
        price_changes: Raw material -> new price (Rs/t, its own price basis).
        spec_changes: Spec name -> {"Min"/"Max"/"Margin": new value or None}.
        usage_changes: Raw material -> {"Min"/"Max": new product limit or None}.
    """

    name: str
    price_changes: dict[str, float] = field(default_factory=dict)
    spec_changes: dict[str, dict[str, float | None]] = field(default_factory=dict)
    usage_changes: dict[str, dict[str, float | None]] = field(default_factory=dict)


@dataclass
class BlendData:
    """Everything the optimizer needs, validated and in consistent units.

    Attributes:
        raw_materials: Name -> RawMaterial, in sheet order.
        components: All component column names, in sheet order.
        specs: Spec rows, in sheet order.
        settings: Product settings.
        product_limits: Name -> product/customer usage limit (UsageLimits sheet).
            Materials not listed have no product limit.
        lab_convention: The lab analysis convention the data was read with.
        warnings: Messages that did not stop the run.
        sources: Sheet name -> file it was read from.
        scenarios: What-if scenarios from the optional WhatIf sheet, in the
            order they first appear.
        product_limit_sources: Raw material -> where its product limit came
            from, when not the UsageLimits sheet (e.g. "WhatIf (No fluorspar)"
            for a scenario's usage change). Used in diagnosis messages.
    """

    raw_materials: dict[str, RawMaterial]
    components: list[str]
    specs: list[SpecRow]
    settings: Settings
    product_limits: dict[str, UsageLimit] = field(default_factory=dict)
    lab_convention: str = config.LAB_ANALYSIS_CONVENTION
    warnings: list[str] = field(default_factory=list)
    sources: dict[str, str] = field(default_factory=dict)
    scenarios: list[Scenario] = field(default_factory=list)
    product_limit_sources: dict[str, str] = field(default_factory=dict)

    @property
    def material_names(self) -> list[str]:
        """Raw material names in sheet order."""
        return list(self.raw_materials)

    def effective_limit(self, name: str) -> UsageLimit:
        """Usage limit actually applied: the tighter of plant-wide and product limits."""
        plant = self.raw_materials[name].plant_limit
        product = self.product_limits.get(name, UsageLimit())
        return plant.tighter(product)


class InputError(Exception):
    """The input workbook(s) contain errors.

    All errors are collected before this is raised, so one run reports every
    problem. Warnings found up to that point are attached as well.
    """

    def __init__(self, errors: list[str], warnings: list[str] | None = None) -> None:
        self.errors = list(errors)
        self.warnings = list(warnings or [])
        lines = [f"{len(self.errors)} input error(s) found:"]
        lines += [f"  - {message}" for message in self.errors]
        super().__init__("\n".join(lines))

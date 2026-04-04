"""Configuration dataclasses for the Transmission NPV Model.

All monetary values are in millions of dollars (USD).
All rates are expressed as decimals (e.g., 0.05 = 5%).
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from typing import Optional


@dataclass
class ProjectConfig:
    """Top-level project parameters."""
    name: str = "HV Transmission Network"
    start_year: int = 2026
    construction_years: int = 3
    operating_years: int = 35
    currency: str = "USD"

    @property
    def total_years(self) -> int:
        return self.construction_years + self.operating_years

    @property
    def cod_year(self) -> int:
        """Commercial Operation Date year."""
        return self.start_year + self.construction_years

    @property
    def end_year(self) -> int:
        return self.start_year + self.total_years - 1


@dataclass
class CapexConfig:
    """Capital expenditure parameters."""
    # Transmission lines
    transmission_lines_km: float = 250.0
    cost_per_km: float = 2.5  # $M per km

    # Substations
    substation_count: int = 4
    cost_per_substation: float = 45.0  # $M each

    # Transformers
    transformer_count: int = 8
    cost_per_transformer: float = 12.0  # $M each

    # Other costs
    land_easements: float = 35.0  # $M
    construction_overhead_pct: float = 0.12
    contingency_pct: float = 0.10

    # Construction phasing: fraction of total capex spent each construction year
    construction_phasing: list = field(
        default_factory=lambda: [0.20, 0.45, 0.35]
    )

    # Post-COD sustaining capex as % of opening RAB
    sustaining_capex_pct: float = 0.015
    capex_escalation_rate: float = 0.025

    def compute_total(self) -> dict:
        """Compute total capex breakdown in $M."""
        lines = self.transmission_lines_km * self.cost_per_km
        substations = self.substation_count * self.cost_per_substation
        transformers = self.transformer_count * self.cost_per_transformer
        land = self.land_easements
        direct = lines + substations + transformers + land
        overhead = direct * self.construction_overhead_pct
        subtotal = direct + overhead
        contingency = subtotal * self.contingency_pct
        total = subtotal + contingency
        return {
            "transmission_lines": lines,
            "substations": substations,
            "transformers": transformers,
            "land_easements": land,
            "overhead": overhead,
            "contingency": contingency,
            "total": total,
        }


@dataclass
class OpexConfig:
    """Operating expenditure parameters."""
    maintenance_pct_of_rab: float = 0.015
    staffing_base: float = 8.0  # $M/year at COD
    staffing_growth_rate: float = 0.02
    grid_losses_pct: float = 0.025  # fraction of energy throughput
    energy_price_for_losses: float = 50.0  # $/MWh
    insurance_pct_of_rab: float = 0.003
    regulatory_compliance: float = 1.5  # $M/year
    opex_escalation_rate: float = 0.025  # CPI


@dataclass
class RevenueConfig:
    """Revenue parameters (used in merchant mode or for reference)."""
    capacity_charge_per_mw_year: float = 55_000.0  # $/MW/year
    contracted_capacity_mw: float = 1500.0
    capacity_growth_rate: float = 0.015
    energy_charge_per_mwh: float = 4.50
    annual_energy_gwh: float = 8000.0
    energy_growth_rate: float = 0.02
    connection_fee_annual: float = 5.0  # $M/year
    revenue_escalation_rate: float = 0.025


@dataclass
class RegulatoryConfig:
    """RAB regulatory framework parameters."""
    allowed_return_on_equity: float = 0.10
    allowed_return_on_debt: float = 0.05
    regulatory_gearing: float = 0.60  # 60% debt
    regulatory_depreciation_years: int = 40
    depreciation_method: str = "straight_line"
    regulatory_period_years: int = 5
    efficiency_factor: float = 0.01  # annual X-factor on controllable opex
    inflation_rate: float = 0.025

    @property
    def vanilla_wacc(self) -> float:
        """Pre-tax vanilla WACC used by regulator for allowed return."""
        return (
            self.regulatory_gearing * self.allowed_return_on_debt
            + (1 - self.regulatory_gearing) * self.allowed_return_on_equity
        )


@dataclass
class TaxConfig:
    """Tax parameters."""
    corporate_tax_rate: float = 0.30
    tax_depreciation_years: int = 25  # accelerated
    tax_depreciation_method: str = "straight_line"
    loss_carry_forward: bool = True


@dataclass
class FundingConfig:
    """Debt/equity structure and cost of capital."""
    gearing_ratio: float = 0.60  # 60% debt
    cost_of_debt: float = 0.05
    cost_of_equity: float = 0.12
    risk_free_rate: float = 0.035
    equity_risk_premium: float = 0.06
    asset_beta: float = 0.45
    debt_tenor_years: int = 15

    def compute_wacc(self, tax_rate: float) -> float:
        """Post-tax nominal WACC."""
        return (
            self.gearing_ratio * self.cost_of_debt * (1 - tax_rate)
            + (1 - self.gearing_ratio) * self.cost_of_equity
        )


@dataclass
class ModelConfig:
    """Top-level model configuration composing all sub-configs."""
    project: ProjectConfig = field(default_factory=ProjectConfig)
    capex: CapexConfig = field(default_factory=CapexConfig)
    opex: OpexConfig = field(default_factory=OpexConfig)
    revenue: RevenueConfig = field(default_factory=RevenueConfig)
    regulatory: RegulatoryConfig = field(default_factory=RegulatoryConfig)
    tax: TaxConfig = field(default_factory=TaxConfig)
    funding: FundingConfig = field(default_factory=FundingConfig)

    def to_dict(self) -> dict:
        return asdict(self)

    def to_json(self, path: Optional[str] = None) -> str:
        data = self.to_dict()
        json_str = json.dumps(data, indent=2)
        if path:
            with open(path, "w") as f:
                f.write(json_str)
        return json_str

    @classmethod
    def from_json(cls, path: str) -> ModelConfig:
        with open(path) as f:
            data = json.load(f)
        return cls(
            project=ProjectConfig(**data.get("project", {})),
            capex=CapexConfig(**data.get("capex", {})),
            opex=OpexConfig(**data.get("opex", {})),
            revenue=RevenueConfig(**data.get("revenue", {})),
            regulatory=RegulatoryConfig(**data.get("regulatory", {})),
            tax=TaxConfig(**data.get("tax", {})),
            funding=FundingConfig(**data.get("funding", {})),
        )

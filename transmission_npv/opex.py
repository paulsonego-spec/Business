"""Operating expenditure projection."""

from __future__ import annotations

import pandas as pd

from transmission_npv.config import OpexConfig, ProjectConfig, RevenueConfig


def compute_opex_projection(
    opex_config: OpexConfig,
    project_config: ProjectConfig,
    revenue_config: RevenueConfig,
    rab_schedule: pd.DataFrame,
) -> pd.DataFrame:
    """Project operating expenditure over the asset life.

    OPEX begins at Commercial Operation Date. Components:
    - Maintenance: % of opening RAB
    - Staffing: base amount growing annually
    - Grid losses: loss_pct * throughput * electricity price
    - Insurance: % of opening RAB
    - Regulatory compliance: fixed annual cost

    All controllable costs escalate by CPI.

    Returns DataFrame indexed by year with columns:
        maintenance, staffing, grid_losses, insurance,
        regulatory_compliance, total_opex
    """
    years = rab_schedule.index.tolist()
    cod_year = project_config.cod_year

    data = {col: [0.0] * len(years) for col in [
        "maintenance", "staffing", "grid_losses",
        "insurance", "regulatory_compliance", "total_opex",
    ]}

    for i, year in enumerate(years):
        if year < cod_year:
            continue

        years_operating = year - cod_year
        escalation = (1 + opex_config.opex_escalation_rate) ** years_operating
        opening_rab = rab_schedule.loc[year, "opening_rab"]

        # Maintenance as % of RAB
        data["maintenance"][i] = opening_rab * opex_config.maintenance_pct_of_rab

        # Staffing with real growth + CPI
        data["staffing"][i] = (
            opex_config.staffing_base
            * (1 + opex_config.staffing_growth_rate) ** years_operating
            * (1 + opex_config.opex_escalation_rate) ** years_operating
        )

        # Grid losses: throughput grows with energy growth rate
        throughput_gwh = revenue_config.annual_energy_gwh * (
            (1 + revenue_config.energy_growth_rate) ** years_operating
        )
        throughput_mwh = throughput_gwh * 1000
        loss_cost = (
            opex_config.grid_losses_pct
            * throughput_mwh
            * opex_config.energy_price_for_losses
            / 1e6  # convert to $M
        )
        data["grid_losses"][i] = loss_cost

        # Insurance as % of RAB
        data["insurance"][i] = opening_rab * opex_config.insurance_pct_of_rab

        # Regulatory compliance (escalated)
        data["regulatory_compliance"][i] = (
            opex_config.regulatory_compliance * escalation
        )

    df = pd.DataFrame(data, index=years)
    df.index.name = "year"
    df["total_opex"] = (
        df["maintenance"] + df["staffing"] + df["grid_losses"]
        + df["insurance"] + df["regulatory_compliance"]
    )
    return df

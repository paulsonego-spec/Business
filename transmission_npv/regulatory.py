"""Regulated Asset Base (RAB) framework calculations.

Handles RAB roll-forward, regulatory depreciation, and allowed revenue
under a regulated utility model.
"""

from __future__ import annotations

import pandas as pd

from transmission_npv.config import RegulatoryConfig, ProjectConfig


def compute_rab_rollforward(
    regulatory_config: RegulatoryConfig,
    project_config: ProjectConfig,
    capex_schedule: pd.DataFrame,
) -> pd.DataFrame:
    """Compute RAB roll-forward over the project life.

    RAB logic:
        Opening RAB + capex - regulatory depreciation = Closing RAB

    Depreciation is straight-line over regulatory_depreciation_years,
    applied to each capex tranche separately.

    Returns DataFrame indexed by year with columns:
        opening_rab, capex_additions, regulatory_depreciation, closing_rab
    """
    years = capex_schedule.index.tolist()
    n_years = len(years)

    opening_rab = [0.0] * n_years
    capex_additions = [0.0] * n_years
    reg_depreciation = [0.0] * n_years
    closing_rab = [0.0] * n_years

    dep_life = regulatory_config.regulatory_depreciation_years

    # Track depreciation from each year's capex tranche
    tranche_annual_dep = {}
    for i, year in enumerate(years):
        capex_this_year = capex_schedule.loc[year, "total_capex"]
        capex_additions[i] = capex_this_year

        if capex_this_year > 0 and dep_life > 0:
            tranche_annual_dep[year] = capex_this_year / dep_life

        # Opening RAB
        if i == 0:
            opening_rab[i] = 0.0
        else:
            opening_rab[i] = closing_rab[i - 1]

        # Total depreciation this year from all active tranches
        total_dep = 0.0
        for tranche_year, annual_dep in tranche_annual_dep.items():
            tranche_age = year - tranche_year
            if 0 <= tranche_age < dep_life:
                total_dep += annual_dep
        reg_depreciation[i] = total_dep

        # Closing RAB (floor at zero)
        closing_rab[i] = max(
            0.0, opening_rab[i] + capex_additions[i] - reg_depreciation[i]
        )

    df = pd.DataFrame({
        "opening_rab": opening_rab,
        "capex_additions": capex_additions,
        "regulatory_depreciation": reg_depreciation,
        "closing_rab": closing_rab,
    }, index=years)
    df.index.name = "year"

    return df


def compute_allowed_revenue(
    regulatory_config: RegulatoryConfig,
    project_config: ProjectConfig,
    rab_schedule: pd.DataFrame,
    opex_projection: pd.DataFrame,
) -> pd.DataFrame:
    """Compute regulated allowed revenue.

    Allowed revenue = return on RAB + regulatory depreciation + opex allowance

    Return on RAB uses the average of opening and closing RAB multiplied
    by the vanilla WACC set by the regulator.

    Returns DataFrame indexed by year with columns:
        average_rab, return_on_rab, depreciation_allowance,
        opex_allowance, total_allowed_revenue
    """
    years = rab_schedule.index.tolist()
    cod_year = project_config.cod_year
    wacc = regulatory_config.vanilla_wacc

    avg_rab = (rab_schedule["opening_rab"] + rab_schedule["closing_rab"]) / 2
    return_on_rab = avg_rab * wacc
    dep_allowance = rab_schedule["regulatory_depreciation"]

    # Opex allowance: pass through actual opex (pre efficiency factor)
    opex_allowance = opex_projection["total_opex"].reindex(years, fill_value=0.0)

    # No revenue during construction
    for year in years:
        if year < cod_year:
            return_on_rab.loc[year] = 0.0
            dep_allowance.loc[year] = 0.0
            opex_allowance.loc[year] = 0.0

    total = return_on_rab + dep_allowance + opex_allowance

    df = pd.DataFrame({
        "average_rab": avg_rab,
        "return_on_rab": return_on_rab,
        "depreciation_allowance": dep_allowance,
        "opex_allowance": opex_allowance,
        "total_allowed_revenue": total,
    }, index=years)
    df.index.name = "year"

    return df

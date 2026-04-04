"""Debt/equity structure, WACC, and interest expense calculations."""

from __future__ import annotations

import pandas as pd

from transmission_npv.config import FundingConfig, TaxConfig, ProjectConfig


def compute_wacc(funding_config: FundingConfig, tax_config: TaxConfig) -> dict:
    """Compute weighted average cost of capital.

    Returns dict with pre_tax_wacc, post_tax_wacc, cost_of_equity, cost_of_debt.
    """
    g = funding_config.gearing_ratio
    kd = funding_config.cost_of_debt
    ke = funding_config.cost_of_equity
    t = tax_config.corporate_tax_rate

    post_tax_wacc = g * kd * (1 - t) + (1 - g) * ke
    pre_tax_wacc = g * kd + (1 - g) * ke

    return {
        "cost_of_equity": ke,
        "cost_of_debt": kd,
        "gearing": g,
        "pre_tax_wacc": pre_tax_wacc,
        "post_tax_wacc": post_tax_wacc,
    }


def compute_debt_schedule(
    funding_config: FundingConfig,
    project_config: ProjectConfig,
    capex_schedule: pd.DataFrame,
) -> pd.DataFrame:
    """Compute debt drawdown, balance, and interest expense.

    Debt = gearing_ratio * cumulative capex during construction.
    Post-construction, debt amortises over the debt tenor.
    Interest = opening debt balance * cost of debt.

    Returns DataFrame indexed by year with columns:
        opening_debt, drawdown, repayment, closing_debt, interest_expense
    """
    years = capex_schedule.index.tolist()
    n = len(years)
    cod_year = project_config.cod_year
    gearing = funding_config.gearing_ratio
    rate = funding_config.cost_of_debt
    tenor = funding_config.debt_tenor_years

    opening_debt = [0.0] * n
    drawdown = [0.0] * n
    repayment = [0.0] * n
    closing_debt = [0.0] * n
    interest = [0.0] * n

    peak_debt = 0.0

    for i, year in enumerate(years):
        # Opening balance
        if i > 0:
            opening_debt[i] = closing_debt[i - 1]

        capex = capex_schedule.loc[year, "total_capex"]

        if year < cod_year:
            # Construction: draw debt proportional to capex
            drawdown[i] = capex * gearing
        else:
            drawdown[i] = 0.0
            # Amortise debt over tenor starting from COD
            if peak_debt > 0 and tenor > 0:
                annual_repayment = peak_debt / tenor
                repayment[i] = min(annual_repayment, opening_debt[i])
            else:
                repayment[i] = 0.0

        closing_debt[i] = opening_debt[i] + drawdown[i] - repayment[i]
        interest[i] = opening_debt[i] * rate

        # Track peak debt at COD
        if year == cod_year - 1:
            peak_debt = closing_debt[i]

    # If peak_debt wasn't set (single year construction), use max
    if peak_debt == 0.0:
        peak_debt = max(closing_debt)

    # Recalculate repayments with correct peak_debt
    for i, year in enumerate(years):
        if i > 0:
            opening_debt[i] = closing_debt[i - 1]
        capex = capex_schedule.loc[year, "total_capex"]

        if year < cod_year:
            drawdown[i] = capex * gearing
            repayment[i] = 0.0
        else:
            drawdown[i] = 0.0
            if peak_debt > 0 and tenor > 0:
                annual_repayment = peak_debt / tenor
                repayment[i] = min(annual_repayment, max(0.0, opening_debt[i]))
            else:
                repayment[i] = 0.0

        closing_debt[i] = max(0.0, opening_debt[i] + drawdown[i] - repayment[i])
        interest[i] = opening_debt[i] * rate

    df = pd.DataFrame({
        "opening_debt": opening_debt,
        "drawdown": drawdown,
        "repayment": repayment,
        "closing_debt": closing_debt,
        "interest_expense": interest,
    }, index=years)
    df.index.name = "year"
    return df

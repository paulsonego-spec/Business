"""Tax calculations including tax depreciation and tax shields."""

from __future__ import annotations

import pandas as pd

from transmission_npv.config import TaxConfig, ProjectConfig


def compute_tax_depreciation(
    tax_config: TaxConfig,
    project_config: ProjectConfig,
    capex_schedule: pd.DataFrame,
) -> pd.Series:
    """Compute tax depreciation schedule.

    Uses accelerated straight-line depreciation over tax_depreciation_years,
    applied tranche-by-tranche to each year's capex.

    Returns Series indexed by year with total tax depreciation.
    """
    years = capex_schedule.index.tolist()
    dep_life = tax_config.tax_depreciation_years
    tax_dep = pd.Series(0.0, index=years, name="tax_depreciation")

    for tranche_year in years:
        capex = capex_schedule.loc[tranche_year, "total_capex"]
        if capex <= 0:
            continue
        annual_dep = capex / dep_life
        for year in years:
            age = year - tranche_year
            if 0 <= age < dep_life:
                tax_dep.loc[year] += annual_dep

    tax_dep.index.name = "year"
    return tax_dep


def compute_tax(
    tax_config: TaxConfig,
    project_config: ProjectConfig,
    revenue: pd.DataFrame,
    opex: pd.DataFrame,
    interest_expense: pd.Series,
    tax_depreciation: pd.Series,
) -> pd.DataFrame:
    """Compute corporate tax.

    Taxable income = revenue - opex - interest expense - tax depreciation
    Tax payable = max(0, taxable_income) * tax_rate (with loss carry-forward)

    Returns DataFrame indexed by year with columns:
        taxable_income, tax_losses_carried, tax_payable
    """
    years = revenue.index.tolist()
    cod_year = project_config.cod_year

    total_revenue = revenue["total_revenue"]
    total_opex = opex["total_opex"].reindex(years, fill_value=0.0)
    interest = interest_expense.reindex(years, fill_value=0.0)
    tax_dep = tax_depreciation.reindex(years, fill_value=0.0)

    taxable_income = total_revenue - total_opex - interest - tax_dep

    tax_payable = [0.0] * len(years)
    carried_losses = [0.0] * len(years)
    accumulated_loss = 0.0

    for i, year in enumerate(years):
        income = taxable_income.iloc[i]

        if tax_config.loss_carry_forward:
            income_after_losses = income - accumulated_loss
            if income_after_losses > 0:
                accumulated_loss = 0.0
                tax_payable[i] = income_after_losses * tax_config.corporate_tax_rate
            else:
                accumulated_loss = -income_after_losses
                tax_payable[i] = 0.0
        else:
            tax_payable[i] = max(0.0, income) * tax_config.corporate_tax_rate

        carried_losses[i] = accumulated_loss

    df = pd.DataFrame({
        "taxable_income": taxable_income.values,
        "tax_losses_carried": carried_losses,
        "tax_payable": tax_payable,
    }, index=years)
    df.index.name = "year"
    return df

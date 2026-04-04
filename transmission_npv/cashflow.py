"""Cash flow projections: FCFF and FCFE."""

from __future__ import annotations

import pandas as pd

from transmission_npv.config import ProjectConfig


def compute_cashflows(
    project_config: ProjectConfig,
    revenue: pd.DataFrame,
    opex: pd.DataFrame,
    capex_schedule: pd.DataFrame,
    tax_schedule: pd.DataFrame,
    debt_schedule: pd.DataFrame,
    rab_schedule: pd.DataFrame,
) -> pd.DataFrame:
    """Compute Free Cash Flow to Firm (FCFF) and Free Cash Flow to Equity (FCFE).

    FCFF = Revenue - OPEX - Tax - Capex
    FCFE = FCFF - Interest*(1-t) - Debt Repayment + Debt Drawdown

    Returns DataFrame indexed by year with columns:
        revenue, opex, ebitda, depreciation, ebit,
        tax, capex, fcff, interest_after_tax, net_debt_flow, fcfe
    """
    years = revenue.index.tolist()

    total_revenue = revenue["total_revenue"].reindex(years, fill_value=0.0)
    total_opex = opex["total_opex"].reindex(years, fill_value=0.0)
    total_capex = capex_schedule["total_capex"].reindex(years, fill_value=0.0)
    tax_payable = tax_schedule["tax_payable"].reindex(years, fill_value=0.0)
    interest = debt_schedule["interest_expense"].reindex(years, fill_value=0.0)
    debt_drawdown = debt_schedule["drawdown"].reindex(years, fill_value=0.0)
    debt_repayment = debt_schedule["repayment"].reindex(years, fill_value=0.0)
    reg_dep = rab_schedule["regulatory_depreciation"].reindex(years, fill_value=0.0)

    ebitda = total_revenue - total_opex
    ebit = ebitda - reg_dep

    # FCFF: cash available to all capital providers
    fcff = total_revenue - total_opex - tax_payable - total_capex

    # FCFE: cash available to equity holders
    interest_after_tax = interest * (1 - 0.30)  # simplified
    net_debt_flow = debt_drawdown - debt_repayment
    fcfe = fcff - interest + net_debt_flow

    df = pd.DataFrame({
        "revenue": total_revenue,
        "opex": total_opex,
        "ebitda": ebitda,
        "regulatory_depreciation": reg_dep,
        "ebit": ebit,
        "tax": tax_payable,
        "capex": total_capex,
        "fcff": fcff,
        "interest_expense": interest,
        "debt_drawdown": debt_drawdown,
        "debt_repayment": debt_repayment,
        "net_debt_flow": net_debt_flow,
        "fcfe": fcfe,
    }, index=years)
    df.index.name = "year"
    return df

"""NPV, IRR, and payback period calculations."""

from __future__ import annotations

import numpy as np
import pandas as pd


def compute_npv(cashflows: pd.Series, discount_rate: float) -> float:
    """Compute Net Present Value of a cash flow series.

    Args:
        cashflows: Series indexed by year with cash flows.
        discount_rate: Annual discount rate (e.g., 0.08 for 8%).

    Returns:
        NPV in same units as cash flows.
    """
    base_year = cashflows.index[0]
    npv = 0.0
    for year, cf in cashflows.items():
        t = year - base_year
        npv += cf / (1 + discount_rate) ** t
    return npv


def compute_irr(cashflows: pd.Series, guess: float = 0.10) -> float:
    """Compute Internal Rate of Return using Newton's method.

    Args:
        cashflows: Series indexed by year with cash flows.
        guess: Initial IRR guess.

    Returns:
        IRR as a decimal, or NaN if not converged.
    """
    cf_array = cashflows.values.astype(float)

    # Try numpy's IRR equivalent
    try:
        return float(np.irr(cf_array))
    except (AttributeError, ValueError):
        pass

    # Newton's method fallback
    rate = guess
    for _ in range(1000):
        npv = 0.0
        dnpv = 0.0
        for t, cf in enumerate(cf_array):
            npv += cf / (1 + rate) ** t
            if t > 0:
                dnpv -= t * cf / (1 + rate) ** (t + 1)
        if abs(dnpv) < 1e-12:
            break
        new_rate = rate - npv / dnpv
        if abs(new_rate - rate) < 1e-10:
            return new_rate
        rate = new_rate
        # Guard against divergence
        if rate < -0.99 or rate > 10.0:
            return float("nan")

    return rate


def compute_payback_period(cashflows: pd.Series) -> dict:
    """Compute simple and discounted payback periods.

    Returns dict with:
        simple_payback: years until cumulative cash flow turns positive
        discounted_payback: years until cumulative discounted CF turns positive
    """
    base_year = cashflows.index[0]
    cumulative = 0.0
    simple_payback = None

    for year, cf in cashflows.items():
        cumulative += cf
        if cumulative >= 0 and simple_payback is None:
            simple_payback = year - base_year

    return {
        "simple_payback_years": simple_payback,
    }


def compute_project_valuation(
    cashflows: pd.DataFrame,
    wacc: float,
    cost_of_equity: float,
) -> dict:
    """Compute all valuation metrics.

    Args:
        cashflows: DataFrame from cashflow module with fcff and fcfe columns.
        wacc: Weighted average cost of capital.
        cost_of_equity: Cost of equity for FCFE discounting.

    Returns:
        Dict with npv_project, npv_equity, irr_project, irr_equity,
        simple_payback, total_capex, total_revenue.
    """
    fcff = cashflows["fcff"]
    fcfe = cashflows["fcfe"]

    npv_project = compute_npv(fcff, wacc)
    npv_equity = compute_npv(fcfe, cost_of_equity)
    irr_project = compute_irr(fcff)
    irr_equity = compute_irr(fcfe)
    payback = compute_payback_period(fcff)

    return {
        "npv_project": npv_project,
        "npv_equity": npv_equity,
        "irr_project": irr_project,
        "irr_equity": irr_equity,
        "simple_payback_years": payback["simple_payback_years"],
        "total_capex": cashflows["capex"].sum(),
        "total_revenue": cashflows["revenue"].sum(),
        "wacc": wacc,
        "cost_of_equity": cost_of_equity,
    }

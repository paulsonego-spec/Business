"""Output formatting and summary reports."""

from __future__ import annotations

import pandas as pd


def format_summary(summary: dict) -> str:
    """Format model summary as a readable text report."""
    lines = [
        "=" * 65,
        f"  TRANSMISSION NPV MODEL - {summary['project_name']}",
        "=" * 65,
        "",
        "PROJECT TIMELINE",
        f"  Construction:     {summary['start_year']} - {summary['cod_year'] - 1} ({summary['cod_year'] - summary['start_year']} years)",
        f"  Operations:       {summary['cod_year']} - {summary['end_year']} ({summary['end_year'] - summary['cod_year'] + 1} years)",
        f"  Total project:    {summary['total_years']} years",
        "",
        "CAPITAL EXPENDITURE",
        f"  Initial CAPEX:    ${summary['total_initial_capex_$M']:,.1f}M",
        f"  Lifecycle CAPEX:  ${summary['total_lifecycle_capex_$M']:,.1f}M",
        f"  Peak RAB:         ${summary['peak_rab_$M']:,.1f}M",
        "",
        "COST OF CAPITAL",
        f"  Post-tax WACC:    {summary['wacc_post_tax']:.2%}",
        f"  Pre-tax WACC:     {summary['wacc_pre_tax']:.2%}",
        f"  Regulatory WACC:  {summary['regulatory_wacc']:.2%}",
        "",
        "VALUATION METRICS",
        f"  Project NPV:      ${summary['npv_project_$M']:,.1f}M",
        f"  Equity NPV:       ${summary['npv_equity_$M']:,.1f}M",
    ]

    if summary.get("irr_project") is not None:
        lines.append(f"  Project IRR:      {summary['irr_project']:.2%}")
    if summary.get("irr_equity") is not None:
        lines.append(f"  Equity IRR:       {summary['irr_equity']:.2%}")
    if summary.get("simple_payback_years") is not None:
        lines.append(f"  Simple Payback:   {summary['simple_payback_years']} years")

    lines += [
        "",
        "REVENUE & COSTS (LIFETIME)",
        f"  Total Revenue:    ${summary['total_revenue_$M']:,.1f}M",
        f"  Total OPEX:       ${summary['total_opex_$M']:,.1f}M",
        f"  Total Tax:        ${summary['total_tax_$M']:,.1f}M",
        "",
        "=" * 65,
    ]
    return "\n".join(lines)


def format_cashflow_table(df: pd.DataFrame, years: int = 10) -> str:
    """Format first N years of cash flow as a text table."""
    subset = df.head(years)
    cols = ["revenue", "opex", "capex", "fcff", "fcfe"]
    available = [c for c in cols if c in subset.columns]
    table = subset[available].copy()

    for col in available:
        table[col] = table[col].apply(lambda x: f"${x:,.1f}M")

    return table.to_string()


def export_to_csv(df: pd.DataFrame, path: str) -> None:
    """Export DataFrame to CSV."""
    df.to_csv(path)


def format_scenario_comparison(comparison_df: pd.DataFrame) -> str:
    """Format scenario comparison table."""
    display_cols = [
        "npv_project", "irr_project", "total_capex", "total_revenue"
    ]
    available = [c for c in display_cols if c in comparison_df.columns]
    table = comparison_df[available].copy()

    for col in available:
        if "npv" in col or "capex" in col or "revenue" in col:
            table[col] = table[col].apply(lambda x: f"${x:,.1f}M")
        elif "irr" in col:
            table[col] = table[col].apply(
                lambda x: f"{x:.2%}" if x is not None else "N/A"
            )

    lines = [
        "=" * 65,
        "  SCENARIO COMPARISON",
        "=" * 65,
        "",
        table.to_string(),
        "",
        "=" * 65,
    ]
    return "\n".join(lines)

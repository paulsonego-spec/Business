#!/usr/bin/env python3
"""Example: Run the Transmission NPV Model with scenario comparison."""

import sys
import os

# Allow running from the examples/ directory
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from transmission_npv.config import ModelConfig
from transmission_npv.model import TransmissionNPVModel
from transmission_npv.sensitivity import run_tornado_analysis, run_scenario_comparison
from transmission_npv.reports import (
    format_summary,
    format_cashflow_table,
    format_scenario_comparison,
)


def main():
    # --- Base Case ---
    print("\n" + "=" * 65)
    print("  RUNNING BASE CASE")
    print("=" * 65 + "\n")

    model = TransmissionNPVModel()
    model.run()

    print(format_summary(model.summary()))

    print("\n\nCASH FLOW PROJECTION (First 15 Years)")
    print("-" * 65)
    df = model.to_dataframe()
    print(format_cashflow_table(df, years=15))

    # --- Scenario Comparison ---
    print("\n\n")
    scenarios = {
        "Base Case": ModelConfig.from_json("scenarios/base_case.json"),
        "Upside": ModelConfig.from_json("scenarios/upside.json"),
        "Downside": ModelConfig.from_json("scenarios/downside.json"),
    }

    def run_fn(config):
        return TransmissionNPVModel(config).run_and_summarize()

    comparison = run_scenario_comparison(
        {name: cfg for name, cfg in scenarios.items()},
        lambda cfg: TransmissionNPVModel(cfg).run().valuation,
    )
    print(format_scenario_comparison(comparison))

    # --- Tornado Analysis ---
    print("\n\nTORNADO SENSITIVITY ANALYSIS (+/- 20%)")
    print("-" * 65)

    base_config = ModelConfig()
    tornado = run_tornado_analysis(
        base_config,
        lambda cfg: TransmissionNPVModel(cfg).run().valuation,
        parameters=[
            "capex.cost_per_km",
            "capex.cost_per_substation",
            "funding.cost_of_equity",
            "funding.cost_of_debt",
            "regulatory.allowed_return_on_equity",
            "opex.staffing_base",
        ],
        variation_pct=0.20,
    )
    print(tornado.to_string(index=False))

    # --- Export ---
    output_path = "output_cashflows.csv"
    df.to_csv(output_path)
    print(f"\n\nFull cash flow projection exported to: {output_path}")


if __name__ == "__main__":
    main()

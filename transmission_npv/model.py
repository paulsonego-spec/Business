"""Main orchestrator class tying all modules together."""

from __future__ import annotations

import pandas as pd

from transmission_npv.config import ModelConfig
from transmission_npv.capex import compute_capex_schedule
from transmission_npv.regulatory import compute_rab_rollforward, compute_allowed_revenue
from transmission_npv.opex import compute_opex_projection
from transmission_npv.revenue import compute_revenue
from transmission_npv.tax import compute_tax_depreciation, compute_tax
from transmission_npv.funding import compute_wacc, compute_debt_schedule
from transmission_npv.cashflow import compute_cashflows
from transmission_npv.valuation import compute_project_valuation


class TransmissionNPVModel:
    """Regulated HV transmission network NPV model.

    Chains all financial modules in sequence:
    1. CAPEX schedule
    2. RAB roll-forward
    3. OPEX projection
    4. Allowed revenue (RAB-based)
    5. Revenue
    6. Debt schedule & interest
    7. Tax depreciation & tax
    8. Cash flows (FCFF, FCFE)
    9. Valuation (NPV, IRR, payback)
    """

    def __init__(self, config: ModelConfig | None = None):
        self.config = config or ModelConfig()
        self._results = {}
        self._has_run = False

    def run(self) -> TransmissionNPVModel:
        """Execute the full model. Returns self for chaining."""
        cfg = self.config

        # 1. CAPEX schedule
        self.capex_schedule = compute_capex_schedule(cfg.capex, cfg.project)

        # 2. RAB roll-forward
        self.rab_schedule = compute_rab_rollforward(
            cfg.regulatory, cfg.project, self.capex_schedule
        )

        # 3. OPEX projection (needs RAB for maintenance/insurance calc)
        self.opex_projection = compute_opex_projection(
            cfg.opex, cfg.project, cfg.revenue, self.rab_schedule
        )

        # 4. Allowed revenue
        self.allowed_revenue = compute_allowed_revenue(
            cfg.regulatory, cfg.project, self.rab_schedule, self.opex_projection
        )

        # 5. Revenue
        self.revenue = compute_revenue(cfg.project, self.allowed_revenue)

        # 6. Debt schedule
        self.debt_schedule = compute_debt_schedule(
            cfg.funding, cfg.project, self.capex_schedule
        )

        # 7. Tax
        self.tax_depreciation = compute_tax_depreciation(
            cfg.tax, cfg.project, self.capex_schedule
        )
        self.tax_schedule = compute_tax(
            cfg.tax, cfg.project, self.revenue, self.opex_projection,
            self.debt_schedule["interest_expense"], self.tax_depreciation
        )

        # 8. Cash flows
        self.cashflows = compute_cashflows(
            cfg.project, self.revenue, self.opex_projection,
            self.capex_schedule, self.tax_schedule, self.debt_schedule,
            self.rab_schedule
        )

        # 9. WACC & Valuation
        self.wacc_details = compute_wacc(cfg.funding, cfg.tax)
        self.valuation = compute_project_valuation(
            self.cashflows,
            self.wacc_details["post_tax_wacc"],
            self.wacc_details["cost_of_equity"],
        )

        self._has_run = True
        return self

    def summary(self) -> dict:
        """Return key valuation metrics as a dict."""
        if not self._has_run:
            self.run()

        capex_breakdown = self.config.capex.compute_total()
        return {
            "project_name": self.config.project.name,
            "start_year": self.config.project.start_year,
            "cod_year": self.config.project.cod_year,
            "end_year": self.config.project.end_year,
            "total_years": self.config.project.total_years,
            "total_initial_capex_$M": round(capex_breakdown["total"], 1),
            "total_lifecycle_capex_$M": round(self.cashflows["capex"].sum(), 1),
            "peak_rab_$M": round(self.rab_schedule["closing_rab"].max(), 1),
            "wacc_post_tax": round(self.wacc_details["post_tax_wacc"], 4),
            "wacc_pre_tax": round(self.wacc_details["pre_tax_wacc"], 4),
            "regulatory_wacc": round(self.config.regulatory.vanilla_wacc, 4),
            "npv_project_$M": round(self.valuation["npv_project"], 1),
            "npv_equity_$M": round(self.valuation["npv_equity"], 1),
            "irr_project": round(self.valuation["irr_project"], 4)
            if self.valuation["irr_project"] is not None
            else None,
            "irr_equity": round(self.valuation["irr_equity"], 4)
            if self.valuation["irr_equity"] is not None
            else None,
            "simple_payback_years": self.valuation["simple_payback_years"],
            "total_revenue_$M": round(self.cashflows["revenue"].sum(), 1),
            "total_opex_$M": round(self.cashflows["opex"].sum(), 1),
            "total_tax_$M": round(self.cashflows["tax"].sum(), 1),
        }

    def to_dataframe(self) -> pd.DataFrame:
        """Return consolidated year-by-year projection."""
        if not self._has_run:
            self.run()

        years = self.cashflows.index

        df = pd.DataFrame(index=years)
        df.index.name = "year"

        # CAPEX
        df["capex"] = self.capex_schedule["total_capex"]

        # RAB
        df["opening_rab"] = self.rab_schedule["opening_rab"]
        df["closing_rab"] = self.rab_schedule["closing_rab"]
        df["regulatory_depreciation"] = self.rab_schedule["regulatory_depreciation"]

        # OPEX
        df["opex"] = self.opex_projection["total_opex"]

        # Revenue
        df["revenue"] = self.revenue["total_revenue"]

        # Debt
        df["debt_balance"] = self.debt_schedule["closing_debt"]
        df["interest_expense"] = self.debt_schedule["interest_expense"]

        # Tax
        df["tax_payable"] = self.tax_schedule["tax_payable"]

        # Cash flows
        df["fcff"] = self.cashflows["fcff"]
        df["fcfe"] = self.cashflows["fcfe"]

        return df

    def run_and_summarize(self) -> dict:
        """Convenience: run model and return summary metrics."""
        return self.run().summary()

"""Tests for regulatory module."""

import pytest
from transmission_npv.config import (
    ModelConfig, CapexConfig, ProjectConfig, RegulatoryConfig, OpexConfig, RevenueConfig,
)
from transmission_npv.capex import compute_capex_schedule
from transmission_npv.regulatory import compute_rab_rollforward, compute_allowed_revenue
from transmission_npv.opex import compute_opex_projection


def _build_rab():
    """Helper to build RAB schedule with defaults."""
    cfg = ModelConfig()
    capex = compute_capex_schedule(cfg.capex, cfg.project)
    rab = compute_rab_rollforward(cfg.regulatory, cfg.project, capex)
    return cfg, capex, rab


def test_rab_balance():
    """Closing RAB = Opening RAB + Capex - Depreciation."""
    cfg, capex, rab = _build_rab()

    for year in rab.index:
        expected_closing = (
            rab.loc[year, "opening_rab"]
            + rab.loc[year, "capex_additions"]
            - rab.loc[year, "regulatory_depreciation"]
        )
        assert abs(rab.loc[year, "closing_rab"] - max(0, expected_closing)) < 0.01


def test_rab_starts_at_zero():
    """RAB should start at zero (greenfield project)."""
    _, _, rab = _build_rab()
    assert rab.iloc[0]["opening_rab"] == 0.0


def test_rab_grows_during_construction():
    """RAB should increase during construction years."""
    cfg, _, rab = _build_rab()
    cod_year = cfg.project.cod_year

    for year in range(cfg.project.start_year, cod_year):
        assert rab.loc[year, "closing_rab"] > rab.loc[year, "opening_rab"]


def test_allowed_revenue_zero_during_construction():
    """No revenue during construction."""
    cfg, capex, rab = _build_rab()
    opex = compute_opex_projection(cfg.opex, cfg.project, cfg.revenue, rab)
    allowed = compute_allowed_revenue(cfg.regulatory, cfg.project, rab, opex)

    for year in range(cfg.project.start_year, cfg.project.cod_year):
        assert allowed.loc[year, "total_allowed_revenue"] == 0.0


def test_allowed_revenue_positive_post_cod():
    """Revenue should be positive after COD."""
    cfg, capex, rab = _build_rab()
    opex = compute_opex_projection(cfg.opex, cfg.project, cfg.revenue, rab)
    allowed = compute_allowed_revenue(cfg.regulatory, cfg.project, rab, opex)

    assert allowed.loc[cfg.project.cod_year, "total_allowed_revenue"] > 0

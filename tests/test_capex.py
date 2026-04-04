"""Tests for capital expenditure module."""

import pytest
from transmission_npv.config import CapexConfig, ProjectConfig
from transmission_npv.capex import compute_capex_schedule


def test_capex_total_matches_breakdown():
    """Total initial capex from schedule should approximate the breakdown total."""
    capex_cfg = CapexConfig()
    project_cfg = ProjectConfig()
    breakdown = capex_cfg.compute_total()

    schedule = compute_capex_schedule(capex_cfg, project_cfg)

    # Construction years only (before sustaining capex)
    construction = schedule.loc[
        project_cfg.start_year : project_cfg.cod_year - 1
    ]
    # Sum excluding sustaining capex
    initial_capex = (
        construction["transmission_lines"].sum()
        + construction["substations"].sum()
        + construction["transformers"].sum()
        + construction["land_easements"].sum()
        + construction["overhead"].sum()
        + construction["contingency"].sum()
    )
    # Should be close to breakdown total (within escalation adjustment)
    assert initial_capex > breakdown["total"] * 0.95
    assert initial_capex < breakdown["total"] * 1.10


def test_construction_phasing_coverage():
    """All construction years should have capex."""
    capex_cfg = CapexConfig()
    project_cfg = ProjectConfig()
    schedule = compute_capex_schedule(capex_cfg, project_cfg)

    for year in range(project_cfg.start_year, project_cfg.cod_year):
        assert schedule.loc[year, "total_capex"] > 0


def test_sustaining_capex_post_cod():
    """Sustaining capex should start from COD."""
    capex_cfg = CapexConfig()
    project_cfg = ProjectConfig()
    schedule = compute_capex_schedule(capex_cfg, project_cfg)

    for year in range(project_cfg.start_year, project_cfg.cod_year):
        assert schedule.loc[year, "sustaining_capex"] == 0.0

    # Post-COD should have sustaining capex
    assert schedule.loc[project_cfg.cod_year, "sustaining_capex"] > 0


def test_schedule_spans_full_project():
    """Schedule should cover all project years."""
    capex_cfg = CapexConfig()
    project_cfg = ProjectConfig()
    schedule = compute_capex_schedule(capex_cfg, project_cfg)

    assert len(schedule) == project_cfg.total_years
    assert schedule.index[0] == project_cfg.start_year
    assert schedule.index[-1] == project_cfg.end_year

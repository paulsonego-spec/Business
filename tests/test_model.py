"""Integration tests for the full model."""

import pytest
from transmission_npv.config import ModelConfig
from transmission_npv.model import TransmissionNPVModel


def test_model_runs_without_error():
    """Model should run to completion with default config."""
    model = TransmissionNPVModel()
    model.run()
    assert model._has_run


def test_model_summary_keys():
    """Summary should contain all expected keys."""
    model = TransmissionNPVModel()
    summary = model.run().summary()

    expected_keys = [
        "project_name", "start_year", "cod_year", "end_year",
        "total_years", "total_initial_capex_$M", "total_lifecycle_capex_$M",
        "peak_rab_$M", "wacc_post_tax", "npv_project_$M", "npv_equity_$M",
        "irr_project", "total_revenue_$M", "total_opex_$M", "total_tax_$M",
    ]
    for key in expected_keys:
        assert key in summary, f"Missing key: {key}"


def test_model_dataframe_shape():
    """Consolidated DataFrame should have correct dimensions."""
    config = ModelConfig()
    model = TransmissionNPVModel(config).run()
    df = model.to_dataframe()

    assert len(df) == config.project.total_years
    assert "fcff" in df.columns
    assert "fcfe" in df.columns
    assert "revenue" in df.columns
    assert "capex" in df.columns


def test_model_npv_is_reasonable():
    """Project NPV should be in a reasonable range for default params."""
    model = TransmissionNPVModel().run()
    summary = model.summary()

    # With regulated returns, NPV should exist
    assert summary["npv_project_$M"] is not None
    # Total capex should be roughly $900M+
    assert summary["total_initial_capex_$M"] > 800
    assert summary["total_initial_capex_$M"] < 1200


def test_model_revenue_zero_during_construction():
    """No revenue during construction years."""
    config = ModelConfig()
    model = TransmissionNPVModel(config).run()
    df = model.to_dataframe()

    for year in range(config.project.start_year, config.project.cod_year):
        assert df.loc[year, "revenue"] == 0.0


def test_model_from_json(tmp_path):
    """Model should work when loaded from JSON config."""
    config = ModelConfig()
    json_path = str(tmp_path / "test_config.json")
    config.to_json(json_path)

    loaded = ModelConfig.from_json(json_path)
    model = TransmissionNPVModel(loaded).run()
    summary = model.summary()
    assert summary["npv_project_$M"] is not None


def test_run_and_summarize():
    """Convenience method should work."""
    summary = TransmissionNPVModel().run_and_summarize()
    assert "npv_project_$M" in summary

"""Sensitivity analysis and scenario comparison."""

from __future__ import annotations

import copy
from dataclasses import asdict
from typing import Callable

import pandas as pd

from transmission_npv.config import ModelConfig


def run_sensitivity(
    base_config: ModelConfig,
    run_model_fn: Callable[[ModelConfig], dict],
    parameter_path: str,
    variations: list[float],
) -> pd.DataFrame:
    """Run one-way sensitivity analysis on a single parameter.

    Args:
        base_config: Base case ModelConfig.
        run_model_fn: Function that takes ModelConfig and returns dict of metrics.
        parameter_path: Dot-separated path to parameter, e.g. "capex.cost_per_km".
        variations: List of multipliers, e.g. [0.8, 0.9, 1.0, 1.1, 1.2].

    Returns:
        DataFrame with columns: variation, parameter_value, npv_project, irr_project, etc.
    """
    results = []
    parts = parameter_path.split(".")
    sub_config_name = parts[0]
    param_name = parts[1]

    base_sub = getattr(base_config, sub_config_name)
    base_value = getattr(base_sub, param_name)

    for mult in variations:
        config = copy.deepcopy(base_config)
        sub = getattr(config, sub_config_name)
        new_value = base_value * mult
        setattr(sub, param_name, new_value)

        metrics = run_model_fn(config)
        metrics["variation"] = mult
        metrics["parameter_value"] = new_value
        metrics["parameter"] = parameter_path
        results.append(metrics)

    return pd.DataFrame(results)


def run_tornado_analysis(
    base_config: ModelConfig,
    run_model_fn: Callable[[ModelConfig], dict],
    parameters: list[str],
    variation_pct: float = 0.20,
) -> pd.DataFrame:
    """Run tornado analysis varying each parameter +/- variation_pct.

    Returns DataFrame with columns:
        parameter, low_npv, base_npv, high_npv, low_value, high_value
    """
    base_metrics = run_model_fn(base_config)
    base_npv = base_metrics["npv_project"]

    results = []
    for param_path in parameters:
        low_results = run_sensitivity(
            base_config, run_model_fn, param_path, [1.0 - variation_pct]
        )
        high_results = run_sensitivity(
            base_config, run_model_fn, param_path, [1.0 + variation_pct]
        )

        results.append({
            "parameter": param_path,
            "low_npv": low_results.iloc[0]["npv_project"],
            "base_npv": base_npv,
            "high_npv": high_results.iloc[0]["npv_project"],
            "swing": abs(
                high_results.iloc[0]["npv_project"]
                - low_results.iloc[0]["npv_project"]
            ),
        })

    df = pd.DataFrame(results).sort_values("swing", ascending=True)
    return df


def run_scenario_comparison(
    configs: dict[str, ModelConfig],
    run_model_fn: Callable[[ModelConfig], dict],
) -> pd.DataFrame:
    """Run multiple scenarios and compare results.

    Args:
        configs: Dict mapping scenario name to ModelConfig.
        run_model_fn: Function that takes ModelConfig and returns dict of metrics.

    Returns:
        DataFrame indexed by scenario name with metric columns.
    """
    results = {}
    for name, config in configs.items():
        results[name] = run_model_fn(config)
    return pd.DataFrame(results).T

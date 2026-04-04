"""Capital expenditure scheduling and projection."""

from __future__ import annotations

import numpy as np
import pandas as pd

from transmission_npv.config import CapexConfig, ProjectConfig


def compute_capex_schedule(
    capex_config: CapexConfig, project_config: ProjectConfig
) -> pd.DataFrame:
    """Build year-by-year capex schedule.

    Returns DataFrame indexed by year with columns:
        transmission_lines, substations, transformers, land_easements,
        overhead, contingency, sustaining_capex, total_capex
    """
    years = list(range(
        project_config.start_year,
        project_config.end_year + 1,
    ))
    df = pd.DataFrame(index=years, columns=[
        "transmission_lines", "substations", "transformers",
        "land_easements", "overhead", "contingency",
        "sustaining_capex", "total_capex",
    ], dtype=float)
    df[:] = 0.0
    df.index.name = "year"

    breakdown = capex_config.compute_total()
    phasing = capex_config.construction_phasing

    # Distribute initial capex across construction years
    for i, fraction in enumerate(phasing):
        year = project_config.start_year + i
        if year > project_config.end_year:
            break
        escalation = (1 + capex_config.capex_escalation_rate) ** i
        for category in ["transmission_lines", "substations", "transformers",
                         "land_easements", "overhead", "contingency"]:
            df.loc[year, category] = breakdown[category] * fraction * escalation

    # Sustaining capex post-COD (as % of cumulative initial capex)
    total_initial = breakdown["total"]
    cod_year = project_config.cod_year
    for year in years:
        if year >= cod_year:
            years_from_cod = year - cod_year
            escalation = (1 + capex_config.capex_escalation_rate) ** (
                project_config.construction_years + years_from_cod
            )
            df.loc[year, "sustaining_capex"] = (
                total_initial * capex_config.sustaining_capex_pct * escalation
            )

    df["total_capex"] = df.drop(columns="total_capex").sum(axis=1)

    return df

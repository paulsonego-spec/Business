"""Revenue projection for regulated transmission network."""

from __future__ import annotations

import pandas as pd

from transmission_npv.config import ProjectConfig


def compute_revenue(
    project_config: ProjectConfig,
    allowed_revenue: pd.DataFrame,
) -> pd.DataFrame:
    """Compute revenue projection.

    Under the RAB regulated model, revenue equals the allowed revenue
    determined by the regulator (return on RAB + depreciation + opex).

    Returns DataFrame indexed by year with columns:
        return_on_rab, depreciation_allowance, opex_allowance, total_revenue
    """
    years = allowed_revenue.index.tolist()
    cod_year = project_config.cod_year

    df = pd.DataFrame({
        "return_on_rab": allowed_revenue["return_on_rab"],
        "depreciation_allowance": allowed_revenue["depreciation_allowance"],
        "opex_allowance": allowed_revenue["opex_allowance"],
        "total_revenue": allowed_revenue["total_allowed_revenue"],
    }, index=years)
    df.index.name = "year"

    # Zero out pre-COD years
    df.loc[df.index < cod_year] = 0.0

    return df

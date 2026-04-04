# Transmission Network NPV Model

Net Present Value financial model for a regulated high-voltage electricity transmission network business using the Regulated Asset Base (RAB) framework.

## Overview

This model projects the financial performance of an HV transmission network investment over a 38-year lifecycle (3-year construction + 35-year operation). Under the RAB framework, the regulator sets allowed revenue based on:

- **Return on RAB**: WACC applied to the Regulated Asset Base
- **Regulatory depreciation**: Straight-line depreciation of assets
- **OPEX allowance**: Pass-through of operating costs

## Installation

```bash
pip install -r requirements.txt
```

## Quick Start

```python
from transmission_npv import ModelConfig, TransmissionNPVModel

# Run with default parameters
model = TransmissionNPVModel()
model.run()
print(model.summary())

# Load a scenario
config = ModelConfig.from_json("scenarios/base_case.json")
model = TransmissionNPVModel(config)
model.run()
```

## Running the Example

```bash
python examples/run_model.py
```

## Interactive Analysis

```bash
jupyter notebook notebooks/analysis.ipynb
```

## Project Structure

```
transmission_npv/
├── config.py        # Parameter dataclasses (CapexConfig, OpexConfig, etc.)
├── capex.py         # Capital expenditure scheduling
├── opex.py          # Operating expenditure projection
├── revenue.py       # Regulated revenue calculation
├── regulatory.py    # RAB roll-forward and allowed returns
├── tax.py           # Corporate tax and tax depreciation
├── funding.py       # Debt/equity structure and WACC
├── cashflow.py      # FCFF and FCFE projections
├── valuation.py     # NPV, IRR, payback period
├── sensitivity.py   # Sensitivity and scenario analysis
├── model.py         # Main orchestrator (TransmissionNPVModel)
└── reports.py       # Output formatting
```

## Key Parameters (Base Case)

| Parameter | Value |
|---|---|
| Transmission lines | 250 km @ $2.5M/km |
| Substations | 4 @ $45M each |
| Total initial CAPEX | ~$940M |
| Regulatory depreciation | 40 years straight-line |
| Gearing | 60% debt / 40% equity |
| Cost of debt | 5.0% |
| Cost of equity | 12.0% |
| Corporate tax rate | 30% |
| Operating life | 35 years |

## Scenarios

- **Base Case** (`scenarios/base_case.json`): Default parameters
- **Upside** (`scenarios/upside.json`): Lower CAPEX (-10%), favorable regulatory returns, lower WACC
- **Downside** (`scenarios/downside.json`): CAPEX overrun (+15%), 4-year construction, higher cost of capital

## Model Outputs

- Project NPV (FCFF discounted at WACC)
- Equity NPV (FCFE discounted at cost of equity)
- Project and equity IRR
- Simple payback period
- Year-by-year cash flow projections
- RAB roll-forward schedule
- Sensitivity tornado charts
- Scenario comparison tables

## Testing

```bash
python -m pytest tests/ -v
```

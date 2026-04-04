"""Tests for valuation module."""

import pytest
import pandas as pd
from transmission_npv.valuation import compute_npv, compute_irr, compute_payback_period


def test_npv_known_values():
    """Test NPV against hand-calculated values."""
    # $100 investment, $50/year for 3 years at 10%
    cf = pd.Series([-100, 50, 50, 50], index=[2026, 2027, 2028, 2029])
    npv = compute_npv(cf, 0.10)
    # NPV = -100 + 50/1.1 + 50/1.21 + 50/1.331 = -100 + 45.45 + 41.32 + 37.57 = 24.34
    assert abs(npv - 24.34) < 0.5


def test_npv_zero_rate():
    """NPV at 0% discount rate = sum of cash flows."""
    cf = pd.Series([-100, 50, 50, 50], index=[2026, 2027, 2028, 2029])
    npv = compute_npv(cf, 0.0)
    assert abs(npv - 50.0) < 0.01


def test_irr_known_value():
    """Test IRR for a simple case."""
    # $100 investment returning $110 in 1 year -> IRR = 10%
    cf = pd.Series([-100, 110], index=[2026, 2027])
    irr = compute_irr(cf)
    assert abs(irr - 0.10) < 0.01


def test_irr_break_even():
    """IRR where NPV = 0 at the IRR."""
    cf = pd.Series([-1000, 400, 400, 400], index=[2026, 2027, 2028, 2029])
    irr = compute_irr(cf)
    # Verify NPV ~ 0 at the IRR
    npv_at_irr = compute_npv(cf, irr)
    assert abs(npv_at_irr) < 0.1


def test_payback_period():
    """Test simple payback calculation."""
    cf = pd.Series([-100, 30, 30, 30, 30], index=[2026, 2027, 2028, 2029, 2030])
    result = compute_payback_period(cf)
    # Cumulative: -100, -70, -40, -10, +20 -> payback at year 4 (2030)
    assert result["simple_payback_years"] == 4

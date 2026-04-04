"""Transmission Network NPV Financial Model.

A comprehensive Net Present Value model for regulated high-voltage
electricity transmission network businesses using the RAB framework.
"""

from transmission_npv.config import ModelConfig
from transmission_npv.model import TransmissionNPVModel

__version__ = "1.0.0"
__all__ = ["ModelConfig", "TransmissionNPVModel"]

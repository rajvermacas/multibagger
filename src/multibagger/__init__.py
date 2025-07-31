"""
Multibagger - Financial Investment Analysis Tool

A comprehensive tool for analyzing Excel workbooks containing financial data
and generating structured investment analysis reports.
"""

from .stock_analyzer import analyze_stock_workbook

__version__ = "0.1.0"
__author__ = "Multibagger Team"
__email__ = "team@multibagger.com"

__all__ = ["analyze_stock_workbook"]
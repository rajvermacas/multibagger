"""
Basic test suite for the stock analyzer system.
"""

import pytest
import os
import tempfile
import json
from unittest.mock import Mock, patch, MagicMock
import pandas as pd

# Import the modules to test
from src.multibagger.stock_analyzer import StockAnalyzer, analyze_stock_workbook
from src.multibagger.data_extractor import ExcelDataExtractor
from src.multibagger.financial_calculator import FinancialCalculator
from src.multibagger.utils import validate_excel_file, save_json_report


class TestStockAnalyzer:
    """Test cases for StockAnalyzer class."""

    def test_stock_analyzer_initialization(self):
        """Test StockAnalyzer initialization."""
        analyzer = StockAnalyzer("test.xlsx", "INFO")
        assert analyzer.excel_path == "test.xlsx"
        assert analyzer.extracted_data == {}
        assert analyzer.calculated_metrics == {}
        assert analyzer.final_analysis == {}

    @patch('src.multibagger.stock_analyzer.validate_excel_file')
    def test_validate_input_success(self, mock_validate):
        """Test successful input validation."""
        mock_validate.return_value = True
        analyzer = StockAnalyzer("test.xlsx")
        assert analyzer.validate_input() is True

    @patch('src.multibagger.stock_analyzer.validate_excel_file')
    def test_validate_input_failure(self, mock_validate):
        """Test failed input validation."""
        mock_validate.return_value = False
        analyzer = StockAnalyzer("test.xlsx")
        assert analyzer.validate_input() is False

    @patch('src.multibagger.stock_analyzer.ExcelDataExtractor')
    def test_extract_financial_data_success(self, mock_extractor_class):
        """Test successful financial data extraction."""
        # Mock the extractor
        mock_extractor = Mock()
        mock_extractor.extract_all_data.return_value = {
            'company_info': {'company_name': 'Test Company'},
            'profit_loss': {'years': [2020, 2021, 2022]},
            'balance_sheet': {},
            'cash_flow': {}
        }
        mock_extractor_class.return_value = mock_extractor
        
        analyzer = StockAnalyzer("test.xlsx")
        result = analyzer.extract_financial_data()
        
        assert result is True
        assert analyzer.extracted_data['company_info']['company_name'] == 'Test Company'

    @patch('src.multibagger.stock_analyzer.ExcelDataExtractor')
    def test_extract_financial_data_failure(self, mock_extractor_class):
        """Test failed financial data extraction."""
        mock_extractor = Mock()
        mock_extractor.extract_all_data.return_value = {}
        mock_extractor_class.return_value = mock_extractor
        
        analyzer = StockAnalyzer("test.xlsx")
        result = analyzer.extract_financial_data()
        
        assert result is False

    @patch('src.multibagger.stock_analyzer.FinancialCalculator')
    def test_calculate_financial_metrics_success(self, mock_calculator_class):
        """Test successful financial metrics calculation."""
        mock_calculator = Mock()
        mock_calculator.calculate_all_metrics.return_value = {
            'growth_metrics': {'revenue_cagr_5yr': 15.0},
            'investment_score': {'total_score': 75, 'recommendation': 'BUY'}
        }
        mock_calculator_class.return_value = mock_calculator
        
        analyzer = StockAnalyzer("test.xlsx")
        analyzer.extracted_data = {'company_info': {'company_name': 'Test'}}
        
        result = analyzer.calculate_financial_metrics()
        assert result is True
        assert analyzer.calculated_metrics['investment_score']['recommendation'] == 'BUY'


class TestExcelDataExtractor:
    """Test cases for ExcelDataExtractor class."""

    def test_extractor_initialization(self):
        """Test ExcelDataExtractor initialization."""
        extractor = ExcelDataExtractor("test.xlsx")
        assert extractor.excel_path == "test.xlsx"
        assert extractor.workbook is None
        assert extractor.sheets == {}

    def test_find_year_row(self):
        """Test finding year row in sheet data."""
        # Create mock DataFrame with years
        data = pd.DataFrame({
            'A': ['Label1', 'Years', 'Revenue'],
            'B': ['Value1', 2020, 1000],
            'C': ['Value2', 2021, 1200],
            'D': ['Value3', 2022, 1400]
        })
        
        extractor = ExcelDataExtractor("test.xlsx")
        year_row = extractor.find_year_row(data)
        assert year_row == 1  # Row with years

    def test_extract_years(self):
        """Test extracting years from a row."""
        year_row = pd.Series(['Years', 2020, 2021, 2022, 2023])
        
        extractor = ExcelDataExtractor("test.xlsx")
        years = extractor.extract_years(year_row)
        assert years == [2020, 2021, 2022, 2023]

    def test_find_row_by_label(self):
        """Test finding row by label."""
        data = pd.DataFrame({
            'A': ['Company Name', 'Revenue', 'Profit'],
            'B': ['Test Corp', 1000, 200],
            'C': ['', 1200, 250]
        })
        
        extractor = ExcelDataExtractor("test.xlsx")
        row_idx = extractor.find_row_by_label(data, 'Revenue')
        assert row_idx == 1


class TestFinancialCalculator:
    """Test cases for FinancialCalculator class."""

    def test_calculator_initialization(self):
        """Test FinancialCalculator initialization."""
        test_data = {'profit_loss': {'years': [2020, 2021, 2022]}}
        calculator = FinancialCalculator(test_data)
        assert calculator.data == test_data

    def test_safe_divide(self):
        """Test safe division function."""
        calculator = FinancialCalculator({})
        assert calculator.safe_divide(10, 2) == 5.0
        assert calculator.safe_divide(10, 0) == 0.0
        assert calculator.safe_divide(0, 5) == 0.0

    def test_calculate_cagr(self):
        """Test CAGR calculation."""
        calculator = FinancialCalculator({})
        values = {2020: 1000, 2021: 1200, 2022: 1440}
        cagr = calculator.calculate_cagr(values)
        assert abs(cagr - 20.0) < 1.0  # Approximately 20% CAGR

    def test_calculate_cagr_insufficient_data(self):
        """Test CAGR calculation with insufficient data."""
        calculator = FinancialCalculator({})
        values = {2020: 1000}  # Only one data point
        cagr = calculator.calculate_cagr(values)
        assert cagr == 0.0

    def test_calculate_growth_metrics(self):
        """Test growth metrics calculation."""
        test_data = {
            'profit_loss': {
                'years': [2020, 2021, 2022],
                'sales': {2020: 1000, 2021: 1200, 2022: 1440},
                'net_profit': {2020: 100, 2021: 130, 2022: 180},
                'eps': {2020: 10, 2021: 13, 2022: 18}
            },
            'quarterly': {
                'revenue': [250, 300, 350, 400]
            }
        }
        
        calculator = FinancialCalculator(test_data)
        growth_metrics = calculator.calculate_growth_metrics()
        
        assert 'revenue_cagr_3yr' in growth_metrics
        assert 'profit_cagr_3yr' in growth_metrics
        assert 'eps_growth_rate' in growth_metrics
        assert growth_metrics['revenue_cagr_3yr'] > 0


class TestUtilityFunctions:
    """Test cases for utility functions."""

    def test_validate_excel_file_nonexistent(self):
        """Test validation of non-existent file."""
        result = validate_excel_file("nonexistent.xlsx")
        assert result is False

    def test_validate_excel_file_wrong_extension(self):
        """Test validation of file with wrong extension."""
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            tmp.write(b"test content")
            tmp_path = tmp.name
        
        try:
            result = validate_excel_file(tmp_path)
            assert result is False
        finally:
            os.unlink(tmp_path)

    @patch('src.multibagger.utils.ensure_directory_exists')
    @patch('builtins.open')
    @patch('json.dump')
    def test_save_json_report(self, mock_json_dump, mock_open, mock_ensure_dir):
        """Test JSON report saving."""
        mock_ensure_dir.return_value = True
        
        test_data = {'company_info': {'name': 'Test Company'}}
        result = save_json_report(test_data, 'Test Company')
        
        assert result is not None
        assert 'Test_Company_analysis_' in result


class TestIntegration:
    """Integration tests."""

    @patch('src.multibagger.stock_analyzer.validate_excel_file')
    @patch('src.multibagger.stock_analyzer.ExcelDataExtractor')
    @patch('src.multibagger.stock_analyzer.save_json_report')
    def test_analyze_stock_workbook_integration(self, mock_save, mock_extractor_class, mock_validate):
        """Test the main analyze_stock_workbook function."""
        # Setup mocks
        mock_validate.return_value = True
        mock_save.return_value = "/path/to/result.json"
        
        # Mock extractor
        mock_extractor = Mock()
        mock_extractor.extract_all_data.return_value = {
            'company_info': {'company_name': 'Test Company', 'current_price': 100},
            'profit_loss': {
                'years': [2020, 2021, 2022],
                'sales': {2020: 1000, 2021: 1200, 2022: 1440},
                'net_profit': {2020: 100, 2021: 130, 2022: 180},
                'operating_profit': {2020: 150, 2021: 180, 2022: 220},
                'eps': {2020: 10, 2021: 13, 2022: 18}
            },
            'balance_sheet': {
                'years': [2020, 2021, 2022],
                'total_equity': {2020: 500, 2021: 600, 2022: 720},
                'total_debt': {2020: 200, 2021: 220, 2022: 240},
                'current_assets': {2020: 300, 2021: 350, 2022: 400},
                'current_liabilities': {2020: 100, 2021: 120, 2022: 140},
                'total_assets': {2020: 800, 2021: 950, 2022: 1100}
            },
            'cash_flow': {
                'years': [2020, 2021, 2022],
                'operating_cash_flow': {2020: 120, 2021: 150, 2022: 200},
                'capex': {2020: 50, 2021: 60, 2022: 70},
                'free_cash_flow': {2020: 70, 2021: 90, 2022: 130}
            }
        }
        mock_extractor_class.return_value = mock_extractor
        
        # Run the analysis
        result = analyze_stock_workbook("test.xlsx")
        
        # Verify result
        assert result == "/path/to/result.json"
        mock_validate.assert_called_once()
        mock_extractor.extract_all_data.assert_called_once()
        mock_save.assert_called_once()


if __name__ == "__main__":
    pytest.main([__file__])
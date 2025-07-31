"""
Stock Analyzer - Main Analysis Module

This is the main module that orchestrates the complete stock analysis process.
It extracts data from Excel workbooks, calculates financial metrics, and saves
the results to JSON files for AI agent consumption.
"""

import os
import logging
from datetime import datetime
from typing import Dict, Any, Optional
from pathlib import Path

from .data_extractor import ExcelDataExtractor
from .financial_calculator import FinancialCalculator
from .utils import (
    validate_excel_file, 
    save_json_report, 
    setup_logging,
    safe_get_nested_value
)
from .config import get_recommendation_from_score

logger = logging.getLogger(__name__)


class StockAnalyzer:
    """
    Main stock analyzer class that coordinates the entire analysis process.
    """
    
    def __init__(self, excel_path: str, log_level: str = "INFO"):
        """
        Initialize the stock analyzer.
        
        Args:
            excel_path (str): Path to the Excel workbook
            log_level (str): Logging level
        """
        self.excel_path = excel_path
        self.extracted_data = {}
        self.calculated_metrics = {}
        self.final_analysis = {}
        
        # Setup logging
        setup_logging(log_level)
        logger.info(f"StockAnalyzer initialized for: {excel_path}")
    
    def validate_input(self) -> bool:
        """
        Validate the input Excel file.
        
        Returns:
            bool: True if valid, False otherwise
        """
        logger.info("Validating input Excel file...")
        is_valid = validate_excel_file(self.excel_path)
        
        if is_valid:
            logger.info("Input validation successful")
        else:
            logger.error("Input validation failed")
        
        return is_valid
    
    def extract_financial_data(self) -> bool:
        """
        Extract financial data from the Excel workbook.
        
        Returns:
            bool: True if successful, False otherwise
        """
        logger.info("Starting financial data extraction...")
        
        try:
            extractor = ExcelDataExtractor(self.excel_path)
            self.extracted_data = extractor.extract_all_data()
            
            if not self.extracted_data:
                logger.error("No data extracted from Excel file")
                return False
            
            # Log extraction summary
            company_name = safe_get_nested_value(
                self.extracted_data, 
                ['company_info', 'company_name'], 
                'Unknown Company'
            )
            logger.info(f"Data extraction completed for: {company_name}")
            
            # Log data availability
            sheets_found = []
            for sheet_type in ['profit_loss', 'balance_sheet', 'cash_flow', 'quarterly']:
                if sheet_type in self.extracted_data and self.extracted_data[sheet_type]:
                    sheets_found.append(sheet_type)
            
            logger.info(f"Data found in sheets: {', '.join(sheets_found)}")
            return True
            
        except Exception as e:
            logger.error(f"Data extraction failed: {str(e)}")
            return False
    
    def calculate_financial_metrics(self) -> bool:
        """
        Calculate comprehensive financial metrics and ratios.
        
        Returns:
            bool: True if successful, False otherwise
        """
        logger.info("Starting financial metrics calculation...")
        
        try:
            calculator = FinancialCalculator(self.extracted_data)
            self.calculated_metrics = calculator.calculate_all_metrics()
            
            if not self.calculated_metrics:
                logger.error("No metrics calculated")
                return False
            
            # Log calculation summary
            total_score = safe_get_nested_value(
                self.calculated_metrics,
                ['investment_score', 'total_score'],
                0
            )
            recommendation = safe_get_nested_value(
                self.calculated_metrics,
                ['investment_score', 'recommendation'],
                'UNKNOWN'
            )
            
            logger.info(f"Metrics calculation completed. Score: {total_score}, Recommendation: {recommendation}")
            return True
            
        except Exception as e:
            logger.error(f"Metrics calculation failed: {str(e)}")
            return False
    
    def compile_final_analysis(self) -> bool:
        """
        Compile the final analysis combining all data and metrics.
        
        Returns:
            bool: True if successful, False otherwise
        """
        logger.info("Compiling final analysis...")
        
        try:
            # Get company information
            company_info = self.extracted_data.get('company_info', {})
            company_name = company_info.get('company_name', 'Unknown Company')
            
            # Prepare historical data in structured format
            historical_data = self._prepare_historical_data()
            
            # Prepare quarterly data
            quarterly_data = self._prepare_quarterly_data()
            
            # Get investment score and recommendation
            investment_score = self.calculated_metrics.get('investment_score', {})
            score = investment_score.get('total_score', 0)
            recommendation_data = get_recommendation_from_score(score)
            
            # Compile final analysis
            self.final_analysis = {
                'company_info': {
                    'name': company_name,
                    'current_price': company_info.get('current_price', 0.0),
                    'market_cap': company_info.get('market_cap', 0.0),
                    'face_value': company_info.get('face_value', 0.0),
                    'outstanding_shares': company_info.get('outstanding_shares', 0.0),
                    'analysis_date': datetime.now().strftime('%Y-%m-%d')
                },
                
                'historical_data': historical_data,
                
                'quarterly_data': quarterly_data,
                
                'calculated_metrics': {
                    'growth_metrics': self.calculated_metrics.get('growth_metrics', {}),
                    'profitability_ratios': self.calculated_metrics.get('profitability_ratios', {}),
                    'efficiency_ratios': self.calculated_metrics.get('efficiency_ratios', {}),
                    'leverage_ratios': self.calculated_metrics.get('leverage_ratios', {}),
                    'liquidity_ratios': self.calculated_metrics.get('liquidity_ratios', {}),
                    'valuation_ratios': self.calculated_metrics.get('valuation_ratios', {}),
                    'cash_flow_ratios': self.calculated_metrics.get('cash_flow_ratios', {})
                },
                
                'investment_score': {
                    'total_score': score,
                    'category_scores': investment_score.get('category_scores', {}),
                    'recommendation': recommendation_data.get('recommendation', 'UNKNOWN'),
                    'confidence_level': recommendation_data.get('confidence', 'LOW')
                },
                
                'investment_thesis': self.calculated_metrics.get('investment_thesis', {}),
                
                'analysis_metadata': {
                    'excel_file_path': self.excel_path,
                    'analysis_timestamp': datetime.now().isoformat(),
                    'data_quality': self._assess_data_quality(),
                    'missing_data_points': self._identify_missing_data(),
                    'calculation_errors': [],
                    'recommendations_basis': 'Fundamental analysis based on multi-year financial trends',
                    'version': '1.0.0'
                }
            }
            
            logger.info(f"Final analysis compiled for {company_name}")
            return True
            
        except Exception as e:
            logger.error(f"Final analysis compilation failed: {str(e)}")
            return False
    
    def _prepare_historical_data(self) -> Dict[str, Any]:
        """
        Prepare historical financial data in structured format.
        
        Returns:
            Dict[str, Any]: Structured historical data
        """
        pl_data = self.extracted_data.get('profit_loss', {})
        bs_data = self.extracted_data.get('balance_sheet', {})
        cf_data = self.extracted_data.get('cash_flow', {})
        
        # Get years from any available dataset
        years = pl_data.get('years', bs_data.get('years', cf_data.get('years', [])))
        
        historical = {
            'years': years,
            'revenue': [pl_data.get('sales', {}).get(year, 0) for year in years],
            'operating_profit': [pl_data.get('operating_profit', {}).get(year, 0) for year in years],
            'net_profit': [pl_data.get('net_profit', {}).get(year, 0) for year in years],
            'eps': [pl_data.get('eps', {}).get(year, 0) for year in years],
            'dividend': [pl_data.get('dividend', {}).get(year, 0) for year in years],
            'total_equity': [bs_data.get('total_equity', {}).get(year, 0) for year in years],
            'total_debt': [bs_data.get('total_debt', {}).get(year, 0) for year in years],
            'current_assets': [bs_data.get('current_assets', {}).get(year, 0) for year in years],
            'current_liabilities': [bs_data.get('current_liabilities', {}).get(year, 0) for year in years],
            'fixed_assets': [bs_data.get('fixed_assets', {}).get(year, 0) for year in years],
            'operating_cash_flow': [cf_data.get('operating_cash_flow', {}).get(year, 0) for year in years],
            'capex': [cf_data.get('capex', {}).get(year, 0) for year in years],
            'free_cash_flow': [cf_data.get('free_cash_flow', {}).get(year, 0) for year in years]
        }
        
        return historical
    
    def _prepare_quarterly_data(self) -> Dict[str, Any]:
        """
        Prepare quarterly data in structured format.
        
        Returns:
            Dict[str, Any]: Structured quarterly data
        """
        quarterly_raw = self.extracted_data.get('quarterly', {})
        
        return {
            'quarters': quarterly_raw.get('quarters', []),
            'revenue': quarterly_raw.get('revenue', []),
            'net_profit': quarterly_raw.get('net_profit', [])
        }
    
    def _assess_data_quality(self) -> str:
        """
        Assess the quality of extracted data.
        
        Returns:
            str: Data quality assessment (HIGH, MEDIUM, LOW)
        """
        try:
            # Check data completeness
            pl_data = self.extracted_data.get('profit_loss', {})
            bs_data = self.extracted_data.get('balance_sheet', {})
            cf_data = self.extracted_data.get('cash_flow', {})
            
            data_sources = [pl_data, bs_data, cf_data]
            available_sources = sum(1 for source in data_sources if source and source.get('years'))
            
            # Check years of data
            years = pl_data.get('years', [])
            years_count = len(years)
            
            # Check completeness of key metrics
            key_metrics = ['sales', 'net_profit', 'operating_profit']
            complete_metrics = 0
            for metric in key_metrics:
                if metric in pl_data and pl_data[metric]:
                    complete_metrics += 1
            
            # Scoring
            if available_sources >= 3 and years_count >= 5 and complete_metrics >= 2:
                return "HIGH"
            elif available_sources >= 2 and years_count >= 3 and complete_metrics >= 1:
                return "MEDIUM"
            else:
                return "LOW"
                
        except Exception:
            return "LOW"
    
    def _identify_missing_data(self) -> list:
        """
        Identify missing or incomplete data points.
        
        Returns:
            list: List of missing data points
        """
        missing = []
        
        try:
            # Check for missing sheets
            required_sheets = ['profit_loss', 'balance_sheet', 'cash_flow']
            for sheet in required_sheets:
                if sheet not in self.extracted_data or not self.extracted_data[sheet]:
                    missing.append(f"Missing {sheet} data")
            
            # Check for missing key metrics
            pl_data = self.extracted_data.get('profit_loss', {})
            key_metrics = ['sales', 'net_profit', 'operating_profit', 'eps']
            
            for metric in key_metrics:
                if metric not in pl_data or not pl_data[metric]:
                    missing.append(f"Missing {metric} data")
            
            # Check for insufficient years
            years = pl_data.get('years', [])
            if len(years) < 3:
                missing.append(f"Insufficient historical data: only {len(years)} years")
                
        except Exception as e:
            missing.append(f"Error assessing data completeness: {str(e)}")
        
        return missing
    
    def save_analysis(self) -> Optional[str]:
        """
        Save the final analysis to JSON file.
        
        Returns:
            Optional[str]: Path to saved file, None if failed
        """
        logger.info("Saving analysis to JSON file...")
        
        try:
            company_name = safe_get_nested_value(
                self.final_analysis,
                ['company_info', 'name'],
                'Unknown_Company'
            )
            
            file_path = save_json_report(self.final_analysis, company_name)
            
            if file_path:
                logger.info(f"Analysis saved successfully to: {file_path}")
                return file_path
            else:
                logger.error("Failed to save analysis")
                return None
                
        except Exception as e:
            logger.error(f"Error saving analysis: {str(e)}")
            return None
    
    def run_complete_analysis(self) -> Optional[str]:
        """
        Run the complete stock analysis process.
        
        Returns:
            Optional[str]: Path to saved JSON file, None if failed
        """
        logger.info("Starting complete stock analysis...")
        
        # Step 1: Validate input
        if not self.validate_input():
            logger.error("Analysis aborted due to input validation failure")
            return None
        
        # Step 2: Extract financial data
        if not self.extract_financial_data():
            logger.error("Analysis aborted due to data extraction failure")
            return None
        
        # Step 3: Calculate financial metrics
        if not self.calculate_financial_metrics():
            logger.error("Analysis aborted due to metrics calculation failure")
            return None
        
        # Step 4: Compile final analysis
        if not self.compile_final_analysis():
            logger.error("Analysis aborted due to final compilation failure")
            return None
        
        # Step 5: Save analysis
        result_path = self.save_analysis()
        
        if result_path:
            company_name = safe_get_nested_value(
                self.final_analysis,
                ['company_info', 'name'],
                'Unknown Company'
            )
            score = safe_get_nested_value(
                self.final_analysis,
                ['investment_score', 'total_score'],
                0
            )
            recommendation = safe_get_nested_value(
                self.final_analysis,
                ['investment_score', 'recommendation'],
                'UNKNOWN'
            )
            
            logger.info(f"Complete analysis finished for {company_name}")
            logger.info(f"Investment Score: {score}/100, Recommendation: {recommendation}")
            logger.info(f"Results saved to: {result_path}")
        else:
            logger.error("Analysis completed but failed to save results")
        
        return result_path


def analyze_stock_workbook(excel_path: str, log_level: str = "INFO") -> Optional[str]:
    """
    Main function to analyze a stock Excel workbook and return JSON file path.
    
    This is the primary entry point for the stock analysis system.
    
    Args:
        excel_path (str): Path to the Excel workbook containing financial data
        log_level (str): Logging level (DEBUG, INFO, WARNING, ERROR)
        
    Returns:
        Optional[str]: Path to the generated JSON analysis file, None if failed
        
    Example:
        >>> json_path = analyze_stock_workbook("data/company.xlsx")
        >>> if json_path:
        ...     print(f"Analysis saved to: {json_path}")
        ... else:
        ...     print("Analysis failed")
    """
    try:
        # Create analyzer instance
        analyzer = StockAnalyzer(excel_path, log_level)
        
        # Run complete analysis
        result_path = analyzer.run_complete_analysis()
        
        return result_path
        
    except Exception as e:
        logger.error(f"Critical error in analyze_stock_workbook: {str(e)}")
        return None


# Additional utility functions for external use
def get_analysis_summary(json_path: str) -> Optional[Dict[str, Any]]:
    """
    Get a summary of analysis results from a saved JSON file.
    
    Args:
        json_path (str): Path to the JSON analysis file
        
    Returns:
        Optional[Dict[str, Any]]: Summary of key metrics, None if failed
    """
    try:
        from .utils import load_json_report
        
        data = load_json_report(json_path)
        if not data:
            return None
        
        return {
            'company_name': safe_get_nested_value(data, ['company_info', 'name'], 'Unknown'),
            'investment_score': safe_get_nested_value(data, ['investment_score', 'total_score'], 0),
            'recommendation': safe_get_nested_value(data, ['investment_score', 'recommendation'], 'UNKNOWN'),
            'confidence': safe_get_nested_value(data, ['investment_score', 'confidence_level'], 'LOW'),
            'revenue_cagr_5yr': safe_get_nested_value(data, ['calculated_metrics', 'growth_metrics', 'revenue_cagr_5yr'], 0),
            'net_profit_margin': safe_get_nested_value(data, ['calculated_metrics', 'profitability_ratios', 'net_profit_margin'], 0),
            'roe': safe_get_nested_value(data, ['calculated_metrics', 'profitability_ratios', 'roe'], 0),
            'debt_to_equity': safe_get_nested_value(data, ['calculated_metrics', 'leverage_ratios', 'debt_to_equity'], 0),
            'pe_ratio': safe_get_nested_value(data, ['calculated_metrics', 'valuation_ratios', 'pe_ratio'], 0),
            'analysis_date': safe_get_nested_value(data, ['company_info', 'analysis_date'], 'Unknown')
        }
        
    except Exception as e:
        logger.error(f"Error getting analysis summary: {str(e)}")
        return None


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) != 2:
        print("Usage: python stock_analyzer.py <excel_file_path>")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    result = analyze_stock_workbook(excel_file, "INFO")
    
    if result:
        print(f"Analysis completed successfully. Results saved to: {result}")
        
        # Print summary
        summary = get_analysis_summary(result)
        if summary:
            print("\n--- Analysis Summary ---")
            print(f"Company: {summary['company_name']}")
            print(f"Investment Score: {summary['investment_score']}/100")
            print(f"Recommendation: {summary['recommendation']} ({summary['confidence']} confidence)")
            print(f"Revenue CAGR (5yr): {summary['revenue_cagr_5yr']:.2f}%")
            print(f"Net Profit Margin: {summary['net_profit_margin']:.2f}%")
            print(f"ROE: {summary['roe']:.2f}%")
            print(f"P/E Ratio: {summary['pe_ratio']:.2f}")
    else:
        print("Analysis failed. Please check the logs for details.")
        sys.exit(1)
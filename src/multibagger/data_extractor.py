"""
Data Extractor Module for Excel Workbook Analysis

This module handles reading and extracting financial data from Excel workbooks
containing 5 standard sheets: Profit & Loss, Quarters, Balance Sheet, Cash Flow, and Data Sheet.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any
import logging
import re
import os
from datetime import datetime

logger = logging.getLogger(__name__)


class ExcelDataExtractor:
    """
    Extracts financial data from Excel workbooks with standardized sheet structure.
    """
    
    def __init__(self, excel_path: str):
        """
        Initialize the data extractor with an Excel file path.
        
        Args:
            excel_path (str): Path to the Excel workbook
        """
        self.excel_path = excel_path
        self.workbook = None
        self.sheets = {}
        self.extracted_data = {}
        
    def load_workbook(self) -> bool:
        """
        Load the Excel workbook and read all sheets.
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Check if file exists
            if not os.path.exists(self.excel_path):
                logger.error(f"Excel file not found: {self.excel_path}")
                return False
            
            # Read all sheets from the Excel file
            self.workbook = pd.read_excel(self.excel_path, sheet_name=None, engine='openpyxl')
            self.sheets = self.workbook
            
            if not self.sheets:
                logger.error("No sheets found in the workbook")
                return False
                
            logger.info(f"Successfully loaded workbook with sheets: {list(self.sheets.keys())}")
            return True
        except FileNotFoundError:
            logger.error(f"Excel file not found: {self.excel_path}")
            return False
        except PermissionError:
            logger.error(f"Permission denied accessing file: {self.excel_path}")
            return False
        except Exception as e:
            logger.error(f"Failed to load workbook {self.excel_path}: {str(e)}")
            return False
    
    def find_year_row(self, sheet_data: pd.DataFrame) -> Optional[int]:
        """
        Find the row containing year headers in the sheet.
        
        Args:
            sheet_data (pd.DataFrame): The sheet data
            
        Returns:
            Optional[int]: Row index containing years, None if not found
        """
        for idx, row in sheet_data.iterrows():
            # Look for cells containing 4-digit years (datetime objects or strings)
            year_count = 0
            for cell in row:
                if pd.notna(cell):
                    # Check if it's a datetime object (including datetime.datetime and pd.Timestamp)
                    if hasattr(cell, 'year'):
                        year = cell.year
                        if 1990 <= year <= 2030:  # Reasonable year range
                            year_count += 1
                    # Check string/numeric for year patterns
                    elif isinstance(cell, (int, float, str)):
                        cell_str = str(cell)
                        if re.search(r'\b(19|20)\d{2}\b', cell_str):
                            year_count += 1
            
            if year_count >= 3:  # At least 3 years found
                return idx
        return None
    
    def extract_years(self, year_row: pd.Series) -> List[int]:
        """
        Extract year values from a row.
        
        Args:
            year_row (pd.Series): Row containing years
            
        Returns:
            List[int]: List of extracted years
        """
        years = []
        for cell in year_row:
            if pd.notna(cell):
                # Handle datetime objects (including datetime.datetime and pd.Timestamp)
                if hasattr(cell, 'year'):
                    year = cell.year
                    if 1990 <= year <= 2030:  # Reasonable year range
                        years.append(year)
                # Handle string/numeric representations
                else:
                    cell_str = str(cell)
                    year_match = re.search(r'\b(19|20)(\d{2})\b', cell_str)
                    if year_match:
                        year = int(year_match.group(0))
                        if 1990 <= year <= 2030:  # Reasonable year range
                            years.append(year)
        return sorted(years)
    
    def find_row_by_label(self, sheet_data: pd.DataFrame, label: str) -> Optional[int]:
        """
        Find row index by searching for a label in the first column.
        
        Args:
            sheet_data (pd.DataFrame): The sheet data
            label (str): Label to search for
            
        Returns:
            Optional[int]: Row index if found, None otherwise
        """
        for idx, row in sheet_data.iterrows():
            first_cell = row.iloc[0] if len(row) > 0 else None
            if pd.notna(first_cell):
                cell_str = str(first_cell).lower().strip()
                if label.lower() in cell_str:
                    return idx
        return None
    
    def extract_numeric_values(self, row: pd.Series, years: List[int]) -> Dict[int, float]:
        """
        Extract numeric values from a row corresponding to years.
        
        Args:
            row (pd.Series): Data row
            years (List[int]): List of years
            
        Returns:
            Dict[int, float]: Dictionary mapping years to values
        """
        values = {}
        # Skip first column (label) and extract values
        for i, year in enumerate(years):
            try:
                if i + 1 < len(row):
                    value = row.iloc[i + 1]
                    if pd.notna(value) and isinstance(value, (int, float)):
                        values[year] = float(value)
                    else:
                        values[year] = 0.0
                else:
                    values[year] = 0.0
            except (ValueError, IndexError):
                values[year] = 0.0
        return values
    
    def extract_from_profit_loss_sheet(self) -> Dict[str, Any]:
        """
        Extract data from Profit & Loss sheet.
        
        Returns:
            Dict[str, Any]: Extracted P&L data
        """
        sheet_name = self._find_sheet_by_name(['profit', 'p&l', 'pl', 'income'])
        if not sheet_name:
            logger.warning("Profit & Loss sheet not found")
            return {}
        
        sheet_data = self.sheets[sheet_name]
        year_row_idx = self.find_year_row(sheet_data)
        
        if year_row_idx is None:
            logger.warning("Year row not found in Profit & Loss sheet")
            return {}
        
        years = self.extract_years(sheet_data.iloc[year_row_idx])
        
        # Define metrics to extract
        metrics = {
            'sales': ['sales', 'revenue', 'total revenue', 'net sales'],
            'operating_profit': ['operating profit', 'ebit', 'operating income'],
            'net_profit': ['net profit', 'net income', 'profit after tax', 'pat'],
            'eps': ['eps', 'earnings per share'],
            'dividend': ['dividend', 'dividend per share', 'dps']
        }
        
        extracted = {'years': years}
        
        for metric_key, search_terms in metrics.items():
            found = False
            for term in search_terms:
                row_idx = self.find_row_by_label(sheet_data, term)
                if row_idx is not None:
                    values = self.extract_numeric_values(sheet_data.iloc[row_idx], years)
                    extracted[metric_key] = values
                    found = True
                    break
            
            if not found:
                extracted[metric_key] = {year: 0.0 for year in years}
                logger.warning(f"Metric '{metric_key}' not found in P&L sheet")
        
        return extracted
    
    def extract_from_balance_sheet(self) -> Dict[str, Any]:
        """
        Extract data from Balance Sheet.
        
        Returns:
            Dict[str, Any]: Extracted balance sheet data
        """
        sheet_name = self._find_sheet_by_name(['balance', 'bs', 'balance sheet'])
        if not sheet_name:
            logger.warning("Balance Sheet not found")
            return {}
        
        sheet_data = self.sheets[sheet_name]
        year_row_idx = self.find_year_row(sheet_data)
        
        if year_row_idx is None:
            logger.warning("Year row not found in Balance Sheet")
            return {}
        
        years = self.extract_years(sheet_data.iloc[year_row_idx])
        
        # Define balance sheet metrics
        metrics = {
            'total_equity': ['total equity', 'shareholders equity', 'equity'],
            'total_debt': ['total debt', 'total borrowings', 'debt'],
            'current_assets': ['current assets'],
            'current_liabilities': ['current liabilities'],
            'fixed_assets': ['fixed assets', 'property plant equipment', 'ppe'],
            'total_assets': ['total assets']
        }
        
        extracted = {'years': years}
        
        for metric_key, search_terms in metrics.items():
            found = False
            for term in search_terms:
                row_idx = self.find_row_by_label(sheet_data, term)
                if row_idx is not None:
                    values = self.extract_numeric_values(sheet_data.iloc[row_idx], years)
                    extracted[metric_key] = values
                    found = True
                    break
            
            if not found:
                extracted[metric_key] = {year: 0.0 for year in years}
                logger.warning(f"Metric '{metric_key}' not found in Balance Sheet")
        
        return extracted
    
    def extract_from_cash_flow_sheet(self) -> Dict[str, Any]:
        """
        Extract data from Cash Flow sheet.
        
        Returns:
            Dict[str, Any]: Extracted cash flow data
        """
        sheet_name = self._find_sheet_by_name(['cash flow', 'cf', 'cashflow'])
        if not sheet_name:
            logger.warning("Cash Flow sheet not found")
            return {}
        
        sheet_data = self.sheets[sheet_name]
        year_row_idx = self.find_year_row(sheet_data)
        
        if year_row_idx is None:
            logger.warning("Year row not found in Cash Flow sheet")
            return {}
        
        years = self.extract_years(sheet_data.iloc[year_row_idx])
        
        # Define cash flow metrics
        metrics = {
            'operating_cash_flow': ['operating cash flow', 'cash from operations', 'ocf'],
            'capex': ['capex', 'capital expenditure', 'investments'],
            'financing_cash_flow': ['financing cash flow', 'cash from financing']
        }
        
        extracted = {'years': years}
        
        for metric_key, search_terms in metrics.items():
            found = False
            for term in search_terms:
                row_idx = self.find_row_by_label(sheet_data, term)
                if row_idx is not None:
                    values = self.extract_numeric_values(sheet_data.iloc[row_idx], years)
                    extracted[metric_key] = values
                    found = True
                    break
            
            if not found:
                extracted[metric_key] = {year: 0.0 for year in years}
                logger.warning(f"Metric '{metric_key}' not found in Cash Flow sheet")
        
        # Calculate free cash flow
        if 'operating_cash_flow' in extracted and 'capex' in extracted:
            extracted['free_cash_flow'] = {}
            for year in years:
                ocf = extracted['operating_cash_flow'].get(year, 0)
                capex = abs(extracted['capex'].get(year, 0))  # Capex is usually negative
                extracted['free_cash_flow'][year] = ocf - capex
        
        return extracted
    
    def extract_from_quarters_sheet(self) -> Dict[str, Any]:
        """
        Extract quarterly data for trend analysis.
        
        Returns:
            Dict[str, Any]: Extracted quarterly data
        """
        sheet_name = self._find_sheet_by_name(['quarters', 'quarterly', 'q'])
        if not sheet_name:
            logger.warning("Quarters sheet not found")
            return {}
        
        sheet_data = self.sheets[sheet_name]
        
        # Look for quarter headers (Q1, Q2, Q3, Q4, or specific dates)
        quarter_row_idx = None
        quarters = []
        
        for idx, row in sheet_data.iterrows():
            quarter_count = 0
            row_quarters = []
            for cell in row:
                if pd.notna(cell):
                    # Handle datetime objects for quarters (including datetime.datetime and pd.Timestamp)
                    if hasattr(cell, 'strftime'):
                        quarter_count += 1
                        row_quarters.append(cell.strftime('%Y-%m-%d'))
                    else:
                        cell_str = str(cell).upper()
                        if re.search(r'Q[1-4]', cell_str) or 'MAR' in cell_str or 'JUN' in cell_str or 'SEP' in cell_str or 'DEC' in cell_str:
                            quarter_count += 1
                            row_quarters.append(cell_str)
            
            if quarter_count >= 4:  # At least 4 quarters
                quarter_row_idx = idx
                quarters = row_quarters
                break
        
        if quarter_row_idx is None:
            logger.warning("Quarter headers not found")
            return {}
        
        # Extract quarterly revenue and profit
        metrics = {
            'revenue': ['sales', 'revenue', 'total revenue'],
            'net_profit': ['net profit', 'net income', 'profit after tax']
        }
        
        extracted = {'quarters': quarters}
        
        for metric_key, search_terms in metrics.items():
            found = False
            for term in search_terms:
                row_idx = self.find_row_by_label(sheet_data, term)
                if row_idx is not None:
                    values = []
                    for i in range(len(quarters)):
                        try:
                            if i + 1 < len(sheet_data.iloc[row_idx]):
                                value = sheet_data.iloc[row_idx].iloc[i + 1]
                                if pd.notna(value) and isinstance(value, (int, float)):
                                    values.append(float(value))
                                else:
                                    values.append(0.0)
                            else:
                                values.append(0.0)
                        except (ValueError, IndexError):
                            values.append(0.0)
                    
                    extracted[metric_key] = values
                    found = True
                    break
            
            if not found:
                extracted[metric_key] = [0.0] * len(quarters)
                logger.warning(f"Quarterly metric '{metric_key}' not found")
        
        return extracted
    
    def extract_from_data_sheet(self) -> Dict[str, Any]:
        """
        Extract company metadata and consolidated data.
        
        Returns:
            Dict[str, Any]: Extracted company data
        """
        sheet_name = self._find_sheet_by_name(['data', 'company', 'info', 'summary'])
        if not sheet_name:
            logger.warning("Data sheet not found")
            return {}
        
        sheet_data = self.sheets[sheet_name]
        extracted = {}
        
        # First, try to find company name in column headers across all sheets
        company_name = self._extract_company_name_from_headers()
        if company_name:
            extracted['company_name'] = company_name
        
        # Look for company information in data sheet
        company_fields = {
            'current_price': ['price', 'current price', 'share price', 'cmp'],
            'market_cap': ['market cap', 'market capitalization', 'mcap'],
            'face_value': ['face value', 'fv', 'par value'],
            'outstanding_shares': ['shares', 'outstanding shares', 'shares outstanding', 'number of shares']
        }
        
        for field_key, search_terms in company_fields.items():
            found = False
            for term in search_terms:
                row_idx = self.find_row_by_label(sheet_data, term)
                if row_idx is not None:
                    # Look for value in adjacent cells
                    row = sheet_data.iloc[row_idx]
                    for i in range(1, min(len(row), 5)):  # Check next few cells
                        value = row.iloc[i]
                        if pd.notna(value):
                            try:
                                extracted[field_key] = float(value)
                                found = True
                                break
                            except (ValueError, TypeError):
                                continue
                
                if found:
                    break
            
            if not found:
                extracted[field_key] = 0.0
                logger.warning(f"Company field '{field_key}' not found in Data sheet")
        
        # Fallback for company name if not found in headers
        if 'company_name' not in extracted:
            extracted['company_name'] = "Unknown Company"
            logger.warning("Company name not found in any sheet headers")
        
        return extracted
    
    def _find_sheet_by_name(self, keywords: List[str]) -> Optional[str]:
        """
        Find sheet by matching keywords in sheet names.
        
        Args:
            keywords (List[str]): Keywords to search for
            
        Returns:
            Optional[str]: Sheet name if found, None otherwise
        """
        for sheet_name in self.sheets.keys():
            sheet_name_lower = sheet_name.lower()
            for keyword in keywords:
                if keyword.lower() in sheet_name_lower:
                    return sheet_name
        return None
    
    def _extract_company_name_from_headers(self) -> Optional[str]:
        """
        Extract company name from column headers across all sheets.
        
        Returns:
            Optional[str]: Company name if found, None otherwise
        """
        try:
            # Check all sheets for company name in column headers
            for sheet_name, sheet_data in self.sheets.items():
                # Skip customization sheets
                if 'custom' in sheet_name.lower():
                    continue
                    
                for col in sheet_data.columns:
                    col_str = str(col).strip()
                    # Skip generic column names
                    if col_str.startswith('Unnamed:') or col_str in ['SCREENER.IN', 'Narration']:
                        continue
                    
                    # Look for company name patterns
                    if (len(col_str) > 5 and 
                        any(keyword in col_str.upper() for keyword in ['LTD', 'LIMITED', 'INC', 'CORP', 'COMPANY', 'INFORMATICS'])):
                        logger.info(f"Found company name in {sheet_name} headers: {col_str}")
                        return col_str
                        
            return None
            
        except Exception as e:
            logger.error(f"Error extracting company name from headers: {str(e)}")
            return None
    
    def extract_all_data(self) -> Dict[str, Any]:
        """
        Extract data from all sheets in the workbook.
        
        Returns:
            Dict[str, Any]: Complete extracted data
        """
        if not self.load_workbook():
            return {}
        
        self.extracted_data = {
            'company_info': self.extract_from_data_sheet(),
            'profit_loss': self.extract_from_profit_loss_sheet(),
            'balance_sheet': self.extract_from_balance_sheet(),
            'cash_flow': self.extract_from_cash_flow_sheet(),
            'quarterly': self.extract_from_quarters_sheet(),
            'extraction_timestamp': datetime.now().isoformat()
        }
        
        logger.info("Data extraction completed successfully")
        return self.extracted_data


def extract_financial_data(excel_path: str) -> Dict[str, Any]:
    """
    Convenience function to extract all financial data from an Excel workbook.
    
    Args:
        excel_path (str): Path to the Excel workbook
        
    Returns:
        Dict[str, Any]: Extracted financial data
    """
    extractor = ExcelDataExtractor(excel_path)
    return extractor.extract_all_data()
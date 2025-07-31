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
        Find row index by searching for a label in the first column with fuzzy matching.
        
        Args:
            sheet_data (pd.DataFrame): The sheet data
            label (str): Label to search for
            
        Returns:
            Optional[int]: Row index if found, None otherwise
        """
        label_lower = label.lower().strip()
        
        for idx, row in sheet_data.iterrows():
            first_cell = row.iloc[0] if len(row) > 0 else None
            if pd.notna(first_cell):
                cell_str = str(first_cell).lower().strip()
                
                # Exact match
                if label_lower == cell_str:
                    return idx
                
                # Substring match
                if label_lower in cell_str or cell_str in label_lower:
                    return idx
                    
                # Fuzzy matching - remove common words and check key terms
                cell_words = set(cell_str.replace('-', ' ').split())
                label_words = set(label_lower.replace('-', ' ').split())
                
                # Remove common filler words
                filler_words = {'from', 'to', 'the', 'and', 'or', 'of', 'in', 'for', 'with', 'total'}
                cell_words = cell_words - filler_words
                label_words = label_words - filler_words
                
                # Check if key words match
                if label_words and cell_words and len(label_words & cell_words) >= min(len(label_words), 2):
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
        
        # Define balance sheet metrics with comprehensive search terms
        metrics = {
            'total_equity': ['total equity', 'shareholders equity', 'equity', 'equity share capital'],
            'total_debt': ['total debt', 'total borrowings', 'debt', 'borrowings'],
            'current_assets': ['current assets', 'current asset', 'working capital', 'debtors', 'inventory'],
            'current_liabilities': ['current liabilities', 'current liability', 'other liabilities'],
            'fixed_assets': ['fixed assets', 'property plant equipment', 'ppe', 'net block', 'gross block', 'capital work in progress'],
            'total_assets': ['total assets']  # Remove generic 'total' to avoid conflicts
        }
        
        extracted = {'years': years}
        used_rows = set()  # Track which rows have been used to prevent duplicates
        
        for metric_key, search_terms in metrics.items():
            found = False
            matched_term = None
            matched_row = None
            
            for term in search_terms:
                row_idx = self.find_row_by_label(sheet_data, term)
                if row_idx is not None and row_idx not in used_rows:
                    values = self.extract_numeric_values(sheet_data.iloc[row_idx], years)
                    extracted[metric_key] = values
                    used_rows.add(row_idx)  # Mark this row as used
                    found = True
                    matched_term = term
                    matched_row = str(sheet_data.iloc[row_idx].iloc[0])
                    logger.info(f"Balance Sheet: '{metric_key}' mapped to '{matched_row}' using search term '{matched_term}'")
                    break
                elif row_idx is not None and row_idx in used_rows:
                    logger.debug(f"Balance Sheet: Row {row_idx} ('{str(sheet_data.iloc[row_idx].iloc[0])}') already used, skipping for '{metric_key}'")
            
            if not found:
                extracted[metric_key] = {year: 0.0 for year in years}
                logger.warning(f"Balance Sheet: '{metric_key}' not found. Searched for: {search_terms}")
        
        # Calculate total_assets if not found directly
        if 'total_assets' in extracted and all(v == 0.0 for v in extracted['total_assets'].values()):
            logger.info("Balance Sheet: Attempting to calculate total_assets from components")
            calculated_total_assets = {}
            
            for year in years:
                # Try to calculate as current_assets + fixed_assets
                current_assets = extracted.get('current_assets', {}).get(year, 0)
                fixed_assets = extracted.get('fixed_assets', {}).get(year, 0)
                
                # Also check if we can use total_debt + total_equity (balance equation)
                total_debt = extracted.get('total_debt', {}).get(year, 0)
                total_equity = extracted.get('total_equity', {}).get(year, 0)
                
                # Use the larger of the two methods (more reliable)
                method1 = current_assets + fixed_assets  # Assets = Current + Fixed
                method2 = total_debt + total_equity      # Assets = Liabilities + Equity
                
                if method2 > 0 and method2 > method1:
                    calculated_total_assets[year] = method2
                    logger.debug(f"Balance Sheet: Year {year} - using debt+equity method: {method2}")
                elif method1 > 0:
                    calculated_total_assets[year] = method1
                    logger.debug(f"Balance Sheet: Year {year} - using current+fixed assets method: {method1}")
                else:
                    calculated_total_assets[year] = 0.0
            
            # Update total_assets if we got meaningful values
            if any(v > 0 for v in calculated_total_assets.values()):
                extracted['total_assets'] = calculated_total_assets
                non_zero_count = sum(1 for v in calculated_total_assets.values() if v > 0)
                logger.info(f"Balance Sheet: Successfully calculated total_assets for {non_zero_count}/{len(years)} years")
        
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
        
        # Define cash flow metrics with comprehensive search terms
        metrics = {
            'operating_cash_flow': ['operating cash flow', 'cash from operations', 'ocf', 'cash from operating activity', 'operating activity'],
            'capex': ['capex', 'capital expenditure', 'investments', 'cash from investing activity', 'investing activity', 'capital expenditures'],
            'financing_cash_flow': ['financing cash flow', 'cash from financing', 'cash from financing activity', 'financing activity']
        }
        
        extracted = {'years': years}
        used_rows = set()  # Track which rows have been used to prevent duplicates
        
        for metric_key, search_terms in metrics.items():
            found = False
            matched_term = None
            matched_row = None
            
            for term in search_terms:
                row_idx = self.find_row_by_label(sheet_data, term)
                if row_idx is not None and row_idx not in used_rows:
                    values = self.extract_numeric_values(sheet_data.iloc[row_idx], years)
                    extracted[metric_key] = values
                    used_rows.add(row_idx)  # Mark this row as used
                    found = True
                    matched_term = term
                    matched_row = str(sheet_data.iloc[row_idx].iloc[0])
                    logger.info(f"Cash Flow: '{metric_key}' mapped to '{matched_row}' using search term '{matched_term}'")
                    break
                elif row_idx is not None and row_idx in used_rows:
                    logger.debug(f"Cash Flow: Row {row_idx} ('{str(sheet_data.iloc[row_idx].iloc[0])}') already used, skipping for '{metric_key}'")
            
            if not found:
                extracted[metric_key] = {year: 0.0 for year in years}
                logger.warning(f"Cash Flow: '{metric_key}' not found. Searched for: {search_terms}")
        
        # Calculate free cash flow
        if 'operating_cash_flow' in extracted and 'capex' in extracted:
            extracted['free_cash_flow'] = {}
            for year in years:
                ocf = extracted['operating_cash_flow'].get(year, 0)
                capex = abs(extracted['capex'].get(year, 0))  # Capex is usually negative
                extracted['free_cash_flow'][year] = ocf - capex
        
        # Validate cash flow data quality - check for identical values indicating mapping errors
        if 'operating_cash_flow' in extracted and 'capex' in extracted:
            ocf_values = list(extracted['operating_cash_flow'].values())
            capex_values = list(extracted['capex'].values())
            
            # Check if OCF and CAPEX have identical values (indicating wrong mapping)
            if ocf_values == capex_values:
                logger.error("Cash Flow validation: Operating Cash Flow and CAPEX have identical values - likely mapping error!")
                logger.error(f"OCF values: {ocf_values}")
                logger.error(f"CAPEX values: {capex_values}")
            else:
                non_zero_ocf = sum(1 for v in ocf_values if v != 0.0)
                non_zero_capex = sum(1 for v in capex_values if v != 0.0)
                logger.info(f"Cash Flow validation: OCF has {non_zero_ocf}/{len(ocf_values)} non-zero values, CAPEX has {non_zero_capex}/{len(capex_values)} non-zero values")
        
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
                        # Check if it's a valid date (not just time like 00:00:00)
                        if hasattr(cell, 'year') and cell.year > 1900:  # Valid year
                            quarter_count += 1
                            row_quarters.append(cell.strftime('%Y-%m-%d'))
                        else:
                            # Handle time-only cells like "00:00:00" by creating placeholder dates
                            quarter_count += 1
                            quarter_num = len(row_quarters) % 4  # 0,1,2,3 for Q1,Q2,Q3,Q4
                            placeholder_year = 2020 + (len(row_quarters) // 4)  # Increment year every 4 quarters
                            
                            # Use proper quarter-end dates: Q1=Mar-31, Q2=Jun-30, Q3=Sep-30, Q4=Dec-31
                            if quarter_num == 0:  # Q1
                                quarter_date = f"{placeholder_year:04d}-03-31"
                            elif quarter_num == 1:  # Q2  
                                quarter_date = f"{placeholder_year:04d}-06-30"
                            elif quarter_num == 2:  # Q3
                                quarter_date = f"{placeholder_year:04d}-09-30"
                            else:  # Q4
                                quarter_date = f"{placeholder_year:04d}-12-31"
                            
                            row_quarters.append(quarter_date)
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
            logger.warning("Quarter headers not found in sheet")
            logger.info("Quarterly analysis will be skipped - only annual data will be used")
            return {}
        
        logger.info(f"Found quarterly headers in row {quarter_row_idx}: {quarters[:5]}{'...' if len(quarters) > 5 else ''}")
        
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
        
        # Validate quarterly data quality
        total_quarters = len(quarters)
        if total_quarters > 0:
            revenue_values = extracted.get('revenue', [])
            profit_values = extracted.get('net_profit', [])
            
            # Check if all values are zero
            all_revenue_zero = all(v == 0.0 for v in revenue_values)
            all_profit_zero = all(v == 0.0 for v in profit_values)
            
            if all_revenue_zero and all_profit_zero:
                logger.warning(f"Quarterly data: Quarters sheet found but contains no actual financial data (all {total_quarters} quarters are zero)")
                logger.warning("This indicates the Excel file may not have meaningful quarterly breakdowns - annual data will be used for analysis")
            elif all_revenue_zero:
                logger.warning(f"Quarterly data: All {total_quarters} quarters have zero revenue - may indicate incomplete data")
            elif all_profit_zero:
                logger.warning(f"Quarterly data: All {total_quarters} quarters have zero profit - may indicate losses or incomplete data")
            else:
                non_zero_revenue = sum(1 for v in revenue_values if v != 0.0)
                non_zero_profit = sum(1 for v in profit_values if v != 0.0)
                logger.info(f"Quarterly data: Found valid data - {non_zero_revenue}/{total_quarters} quarters with revenue, {non_zero_profit}/{total_quarters} with profit")
                
            # Log quarter date range for verification
            if quarters:
                logger.info(f"Quarterly data: Date range from {quarters[0]} to {quarters[-1]}")
        
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
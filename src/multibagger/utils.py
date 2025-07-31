"""
Utility Functions Module

This module provides utility functions for file handling, JSON operations,
date/time utilities, and other helper functions for the stock analysis system.
"""

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional
import logging

logger = logging.getLogger(__name__)


def ensure_directory_exists(directory_path: str) -> bool:
    """
    Ensure that a directory exists, creating it if necessary.
    
    Args:
        directory_path (str): Path to the directory
        
    Returns:
        bool: True if directory exists or was created successfully
    """
    try:
        Path(directory_path).mkdir(parents=True, exist_ok=True)
        return True
    except Exception as e:
        logger.error(f"Failed to create directory {directory_path}: {str(e)}")
        return False


def generate_report_filename(company_name: str, file_type: str = "json") -> str:
    """
    Generate a standardized filename for reports.
    
    Args:
        company_name (str): Name of the company
        file_type (str): File extension (default: json)
        
    Returns:
        str: Generated filename
    """
    # Clean company name for filename
    clean_name = "".join(c for c in company_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
    clean_name = clean_name.replace(' ', '_')
    
    # Generate timestamp
    timestamp = datetime.now().strftime("%H%M%S")
    
    return f"{clean_name}_analysis_{timestamp}.{file_type}"


def get_reports_directory() -> str:
    """
    Get the reports directory path for today's date.
    
    Returns:
        str: Path to today's reports directory
    """
    today = datetime.now().strftime("%Y-%m-%d")
    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    reports_dir = os.path.join(base_dir, "resources", "reports", today)
    return reports_dir


def save_json_report(data: Dict[str, Any], company_name: str) -> Optional[str]:
    """
    Save analysis data to a JSON file in the reports directory.
    
    Args:
        data (Dict[str, Any]): Analysis data to save
        company_name (str): Company name for filename generation
        
    Returns:
        Optional[str]: Full path to saved file, None if failed
    """
    try:
        # Ensure reports directory exists
        reports_dir = get_reports_directory()
        if not ensure_directory_exists(reports_dir):
            return None
        
        # Generate filename
        filename = generate_report_filename(company_name, "json")
        file_path = os.path.join(reports_dir, filename)
        
        # Add metadata to the data
        data_with_metadata = {
            **data,
            'analysis_metadata': {
                **data.get('analysis_metadata', {}),
                'file_saved_at': datetime.now().isoformat(),
                'file_path': file_path,
                'file_name': filename
            }
        }
        
        # Save to JSON file
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data_with_metadata, f, indent=2, ensure_ascii=False, default=str)
        
        logger.info(f"Analysis data saved to: {file_path}")
        return file_path
        
    except Exception as e:
        logger.error(f"Failed to save JSON report: {str(e)}")
        return None


def load_json_report(file_path: str) -> Optional[Dict[str, Any]]:
    """
    Load analysis data from a JSON file.
    
    Args:
        file_path (str): Path to the JSON file
        
    Returns:
        Optional[Dict[str, Any]]: Loaded data, None if failed
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        logger.info(f"Analysis data loaded from: {file_path}")
        return data
        
    except Exception as e:
        logger.error(f"Failed to load JSON report from {file_path}: {str(e)}")
        return None


def validate_excel_file(file_path: str) -> bool:
    """
    Validate that an Excel file exists and is accessible.
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        bool: True if file is valid, False otherwise
    """
    if not os.path.exists(file_path):
        logger.error(f"Excel file does not exist: {file_path}")
        return False
    
    if not os.path.isfile(file_path):
        logger.error(f"Path is not a file: {file_path}")
        return False
    
    # Check file extension
    valid_extensions = ['.xlsx', '.xls']
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext not in valid_extensions:
        logger.error(f"Invalid file extension {file_ext}. Must be one of {valid_extensions}")
        return False
    
    # Check file size (basic sanity check)
    file_size = os.path.getsize(file_path)
    if file_size == 0:
        logger.error(f"Excel file is empty: {file_path}")
        return False
    
    if file_size > 100 * 1024 * 1024:  # 100MB limit
        logger.warning(f"Excel file is very large ({file_size / (1024*1024):.1f}MB): {file_path}")
    
    return True


def format_currency(amount: float, currency: str = "₹") -> str:
    """
    Format a number as currency with appropriate scaling.
    
    Args:
        amount (float): Amount to format
        currency (str): Currency symbol
        
    Returns:
        str: Formatted currency string
    """
    if amount == 0:
        return f"{currency}0"
    
    abs_amount = abs(amount)
    sign = "-" if amount < 0 else ""
    
    if abs_amount >= 10000000:  # 1 crore
        return f"{sign}{currency}{abs_amount/10000000:.2f}Cr"
    elif abs_amount >= 100000:  # 1 lakh
        return f"{sign}{currency}{abs_amount/100000:.2f}L"
    elif abs_amount >= 1000:  # 1 thousand
        return f"{sign}{currency}{abs_amount/1000:.2f}K"
    else:
        return f"{sign}{currency}{abs_amount:.2f}"


def format_percentage(value: float, decimal_places: int = 2) -> str:
    """
    Format a number as a percentage.
    
    Args:
        value (float): Value to format
        decimal_places (int): Number of decimal places
        
    Returns:
        str: Formatted percentage string
    """
    if value == 0:
        return "0.00%"
    return f"{value:.{decimal_places}f}%"


def safe_get_nested_value(data: Dict[str, Any], keys: list, default: Any = None) -> Any:
    """
    Safely get a nested value from a dictionary.
    
    Args:
        data (Dict[str, Any]): Dictionary to search
        keys (list): List of keys for nested access
        default (Any): Default value if key path not found
        
    Returns:
        Any: Found value or default
    """
    current = data
    for key in keys:
        if isinstance(current, dict) and key in current:
            current = current[key]
        else:
            return default
    return current


def setup_logging(log_level: str = "INFO") -> None:
    """
    Set up logging configuration for the application.
    
    Args:
        log_level (str): Logging level (DEBUG, INFO, WARNING, ERROR)
    """
    # Create logs directory
    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    logs_dir = os.path.join(base_dir, "logs")
    ensure_directory_exists(logs_dir)
    
    # Configure logging
    log_file = os.path.join(logs_dir, f"multibagger_{datetime.now().strftime('%Y%m%d')}.log")
    
    logging.basicConfig(
        level=getattr(logging, log_level.upper()),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    logger.info(f"Logging configured. Log file: {log_file}")


def clean_numeric_value(value: Any) -> float:
    """
    Clean and convert a value to float, handling various formats.
    
    Args:
        value (Any): Value to clean and convert
        
    Returns:
        float: Cleaned numeric value
    """
    if value is None:
        return 0.0
    
    if isinstance(value, (int, float)):
        return float(value)
    
    if isinstance(value, str):
        # Remove common formatting characters
        cleaned = value.replace(',', '').replace('₹', '').replace('%', '').strip()
        
        # Handle negative values in parentheses
        if cleaned.startswith('(') and cleaned.endswith(')'):
            cleaned = '-' + cleaned[1:-1]
        
        try:
            return float(cleaned)
        except ValueError:
            return 0.0
    
    return 0.0


def get_financial_year_from_date(date_str: str) -> Optional[int]:
    """
    Extract financial year from a date string.
    
    Args:
        date_str (str): Date string in various formats
        
    Returns:
        Optional[int]: Financial year, None if parsing fails
    """
    try:
        # Common date formats
        date_formats = ['%Y-%m-%d', '%d-%m-%Y', '%m/%d/%Y', '%d/%m/%Y', '%Y']
        
        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(date_str, fmt)
                # Financial year in India: April to March
                if date_obj.month >= 4:
                    return date_obj.year + 1
                else:
                    return date_obj.year
            except ValueError:
                continue
        
        return None
        
    except Exception:
        return None


def calculate_time_weighted_average(values: Dict[int, float], weight_recent: bool = True) -> float:
    """
    Calculate time-weighted average of values.
    
    Args:
        values (Dict[int, float]): Dictionary of year -> value
        weight_recent (bool): Whether to give more weight to recent values
        
    Returns:
        float: Time-weighted average
    """
    if not values:
        return 0.0
    
    sorted_years = sorted(values.keys())
    total_weight = 0
    weighted_sum = 0
    
    for i, year in enumerate(sorted_years):
        if weight_recent:
            weight = i + 1  # More weight to recent years
        else:
            weight = 1  # Equal weight
        
        weighted_sum += values[year] * weight
        total_weight += weight
    
    return weighted_sum / total_weight if total_weight > 0 else 0.0


def generate_summary_statistics(values: list) -> Dict[str, float]:
    """
    Generate summary statistics for a list of values.
    
    Args:
        values (list): List of numeric values
        
    Returns:
        Dict[str, float]: Summary statistics
    """
    if not values:
        return {'mean': 0, 'median': 0, 'std': 0, 'min': 0, 'max': 0}
    
    import statistics
    
    clean_values = [v for v in values if v is not None and not (isinstance(v, float) and (v != v))]  # Remove None and NaN
    
    if not clean_values:
        return {'mean': 0, 'median': 0, 'std': 0, 'min': 0, 'max': 0}
    
    try:
        return {
            'mean': round(statistics.mean(clean_values), 2),
            'median': round(statistics.median(clean_values), 2),
            'std': round(statistics.stdev(clean_values) if len(clean_values) > 1 else 0, 2),
            'min': round(min(clean_values), 2),
            'max': round(max(clean_values), 2)
        }
    except Exception:
        return {'mean': 0, 'median': 0, 'std': 0, 'min': 0, 'max': 0}
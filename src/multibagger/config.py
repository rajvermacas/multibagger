"""
Configuration Module

This module contains configuration settings, scoring thresholds, and constants
used throughout the stock analysis system.
"""

from typing import Dict, Any

# Investment Scoring Configuration (100 points total)
SCORING_FRAMEWORK = {
    'growth_quality': {
        'weight': 20,
        'thresholds': {
            'excellent': {'min_cagr': 15, 'points': 20},
            'good': {'min_cagr': 10, 'points': 15},
            'fair': {'min_cagr': 5, 'points': 10},
            'poor': {'min_cagr': 0, 'points': 0}
        }
    },
    'profitability': {
        'weight': 20,
        'thresholds': {
            'excellent': {'min_npm': 15, 'improving': True, 'points': 20},
            'good': {'min_npm': 10, 'stable': True, 'points': 15},
            'fair': {'min_npm': 5, 'points': 10},
            'poor': {'min_npm': 0, 'declining': True, 'points': 0}
        }
    },
    'financial_health': {
        'weight': 20,
        'thresholds': {
            'excellent': {'max_de': 0.5, 'min_liquidity': 2.0, 'points': 20},
            'good': {'max_de': 1.0, 'min_liquidity': 1.5, 'points': 15},
            'fair': {'max_de': 2.0, 'min_liquidity': 1.0, 'points': 10},
            'poor': {'max_de': float('inf'), 'min_liquidity': 0, 'points': 0}
        }
    },
    'cash_flow_quality': {
        'weight': 20,
        'thresholds': {
            'excellent': {'min_ocf_np': 1.0, 'growing_fcf': True, 'points': 20},
            'good': {'min_ocf_np': 0.8, 'points': 15},
            'fair': {'min_ocf_np': 0.6, 'points': 10},
            'poor': {'min_ocf_np': 0, 'points': 0}
        }
    },
    'valuation': {
        'weight': 20,
        'thresholds': {
            'excellent': {'max_pe': 15, 'min_growth': 10, 'points': 20},
            'good': {'max_pe': 25, 'points': 15},
            'fair': {'max_pe': 35, 'points': 10},
            'poor': {'max_pe': float('inf'), 'points': 5}
        }
    }
}

# Investment Recommendation Thresholds
RECOMMENDATION_THRESHOLDS = {
    'strong_buy': {'min_score': 70, 'confidence': 'HIGH'},
    'buy': {'min_score': 50, 'confidence': 'MEDIUM'},
    'hold': {'min_score': 30, 'confidence': 'MEDIUM'},
    'avoid': {'min_score': 0, 'confidence': 'HIGH'}
}

# Financial Ratio Benchmarks
RATIO_BENCHMARKS = {
    'profitability': {
        'operating_margin': {
            'excellent': 20.0,
            'good': 15.0,
            'fair': 10.0,
            'poor': 5.0
        },
        'net_profit_margin': {
            'excellent': 15.0,
            'good': 10.0,
            'fair': 5.0,
            'poor': 2.0
        },
        'roe': {
            'excellent': 20.0,
            'good': 15.0,
            'fair': 10.0,
            'poor': 5.0
        },
        'roce': {
            'excellent': 20.0,
            'good': 15.0,
            'fair': 10.0,
            'poor': 5.0
        }
    },
    'leverage': {
        'debt_to_equity': {
            'excellent': 0.5,
            'good': 1.0,
            'fair': 2.0,
            'poor': float('inf')
        },
        'interest_coverage': {
            'excellent': 10.0,
            'good': 5.0,
            'fair': 2.5,
            'poor': 1.0
        }
    },
    'liquidity': {
        'current_ratio': {
            'excellent': 2.0,
            'good': 1.5,
            'fair': 1.0,
            'poor': 0.5
        },
        'quick_ratio': {
            'excellent': 1.5,
            'good': 1.0,
            'fair': 0.8,
            'poor': 0.5
        }
    },
    'efficiency': {
        'asset_turnover': {
            'excellent': 1.5,
            'good': 1.0,
            'fair': 0.5,
            'poor': 0.2
        },
        'receivables_days': {
            'excellent': 30,
            'good': 45,
            'fair': 60,
            'poor': 90
        }
    },
    'valuation': {
        'pe_ratio': {
            'excellent': 15.0,
            'good': 25.0,
            'fair': 35.0,
            'poor': 50.0
        },
        'pb_ratio': {
            'excellent': 1.0,
            'good': 2.0,
            'fair': 3.0,
            'poor': 5.0
        },
        'peg_ratio': {
            'excellent': 0.5,
            'good': 1.0,
            'fair': 1.5,
            'poor': 2.0
        }
    }
}

# Growth Rate Classifications
GROWTH_CLASSIFICATIONS = {
    'revenue_cagr': {
        'high_growth': 20.0,
        'moderate_growth': 10.0,
        'slow_growth': 5.0,
        'no_growth': 0.0
    },
    'profit_cagr': {
        'high_growth': 25.0,
        'moderate_growth': 15.0,
        'slow_growth': 8.0,
        'no_growth': 0.0
    }
}

# Cash Flow Quality Indicators
CASH_FLOW_INDICATORS = {
    'ocf_to_net_profit': {
        'excellent': 1.2,
        'good': 1.0,
        'fair': 0.8,
        'poor': 0.6
    },
    'fcf_margin': {
        'excellent': 10.0,
        'good': 5.0,
        'fair': 2.0,
        'poor': 0.0
    }
}

# Risk Assessment Thresholds
RISK_THRESHOLDS = {
    'financial_risk': {
        'high_debt': 2.0,  # D/E ratio
        'low_liquidity': 1.0,  # Current ratio
        'poor_interest_coverage': 2.5,
        'negative_fcf_years': 2
    },
    'business_risk': {
        'declining_revenue_years': 2,
        'margin_compression': 5.0,  # Percentage points
        'market_cap_threshold': 500  # Crores
    },
    'valuation_risk': {
        'overvalued_pe': 40.0,
        'high_pb': 5.0,
        'expensive_peg': 2.0
    }
}

# Data Quality Thresholds
DATA_QUALITY_CONFIG = {
    'minimum_years': 3,
    'preferred_years': 5,
    'complete_data_threshold': 0.8,  # 80% of data points should be available
    'outlier_detection': {
        'revenue_growth_limit': 100.0,  # % per year
        'margin_limit': 50.0,  # % margin
        'ratio_upper_limit': 100.0
    }
}

# Excel Sheet Identification Patterns
SHEET_PATTERNS = {
    'profit_loss': ['profit', 'p&l', 'pl', 'income', 'statement'],
    'balance_sheet': ['balance', 'bs', 'balance sheet', 'position'],
    'cash_flow': ['cash flow', 'cf', 'cashflow', 'cash'],
    'quarterly': ['quarters', 'quarterly', 'q1', 'q2', 'q3', 'q4'],
    'data_sheet': ['data', 'company', 'info', 'summary', 'overview']
}

# Metric Search Patterns
METRIC_PATTERNS = {
    'revenue': ['sales', 'revenue', 'total revenue', 'net sales', 'turnover'],
    'operating_profit': ['operating profit', 'ebit', 'operating income', 'ebitda'],
    'net_profit': ['net profit', 'net income', 'profit after tax', 'pat', 'profit'],
    'eps': ['eps', 'earnings per share', 'earning per share'],
    'dividend': ['dividend', 'dividend per share', 'dps'],
    'total_equity': ['total equity', 'shareholders equity', 'equity', 'net worth'],
    'total_debt': ['total debt', 'total borrowings', 'debt', 'borrowings'],
    'current_assets': ['current assets', 'ca'],
    'current_liabilities': ['current liabilities', 'cl'],
    'fixed_assets': ['fixed assets', 'property plant equipment', 'ppe', 'fa'],
    'total_assets': ['total assets', 'assets'],
    'operating_cash_flow': ['operating cash flow', 'cash from operations', 'ocf'],
    'capex': ['capex', 'capital expenditure', 'investments', 'capital investments']
}

# Company Information Patterns
COMPANY_INFO_PATTERNS = {
    'company_name': ['company', 'name', 'company name', 'scrip'],
    'current_price': ['price', 'current price', 'share price', 'cmp', 'market price'],
    'market_cap': ['market cap', 'market capitalization', 'mcap', 'market value'],
    'face_value': ['face value', 'fv', 'par value', 'nominal value'],
    'outstanding_shares': ['shares', 'outstanding shares', 'shares outstanding', 'equity shares']
}

# Output Configuration
OUTPUT_CONFIG = {
    'json_indent': 2,
    'currency_symbol': '₹',
    'number_format': {
        'decimal_places': 2,
        'percentage_places': 2,
        'ratio_places': 2
    },
    'date_format': '%Y-%m-%d',
    'timestamp_format': '%Y-%m-%d %H:%M:%S'
}

# Logging Configuration
LOGGING_CONFIG = {
    'level': 'INFO',
    'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    'file_max_size': 10 * 1024 * 1024,  # 10MB
    'backup_count': 5
}

# File Size Limits
FILE_LIMITS = {
    'max_excel_size': 100 * 1024 * 1024,  # 100MB
    'max_json_size': 50 * 1024 * 1024,    # 50MB
    'min_excel_size': 1024                 # 1KB
}

# Analysis Configuration
ANALYSIS_CONFIG = {
    'time_weights': {
        'recent_weight_multiplier': 1.5,
        'years_for_recent': 3
    },
    'trend_analysis': {
        'minimum_points': 3,
        'volatility_threshold': 0.3,
        'growth_consistency_threshold': 0.2
    },
    'peer_comparison': {
        'sector_multiplier_range': (0.8, 1.2),
        'industry_benchmark_weight': 0.3
    }
}


def get_ratio_benchmark(category: str, ratio: str, level: str) -> float:
    """
    Get benchmark value for a specific ratio.
    
    Args:
        category (str): Category (profitability, leverage, etc.)
        ratio (str): Ratio name
        level (str): Benchmark level (excellent, good, fair, poor)
        
    Returns:
        float: Benchmark value
    """
    try:
        return RATIO_BENCHMARKS[category][ratio][level]
    except KeyError:
        return 0.0


def get_scoring_threshold(category: str, level: str) -> Dict[str, Any]:
    """
    Get scoring threshold for a category and level.
    
    Args:
        category (str): Scoring category
        level (str): Threshold level
        
    Returns:
        Dict[str, Any]: Threshold configuration
    """
    try:
        return SCORING_FRAMEWORK[category]['thresholds'][level]
    except KeyError:
        return {}


def get_recommendation_from_score(score: int) -> Dict[str, str]:
    """
    Get investment recommendation based on score.
    
    Args:
        score (int): Investment score
        
    Returns:
        Dict[str, str]: Recommendation and confidence level
    """
    for rec_type, config in RECOMMENDATION_THRESHOLDS.items():
        if score >= config['min_score']:
            return {
                'recommendation': rec_type.upper().replace('_', ' '),
                'confidence': config['confidence']
            }
    
    return {'recommendation': 'AVOID', 'confidence': 'HIGH'}


def validate_configuration() -> bool:
    """
    Validate that all configuration values are properly set.
    
    Returns:
        bool: True if configuration is valid
    """
    required_configs = [
        SCORING_FRAMEWORK,
        RECOMMENDATION_THRESHOLDS,
        RATIO_BENCHMARKS,
        SHEET_PATTERNS,
        METRIC_PATTERNS
    ]
    
    for config in required_configs:
        if not config:
            return False
    
    # Validate scoring framework totals to 100
    total_weight = sum(cat['weight'] for cat in SCORING_FRAMEWORK.values())
    if total_weight != 100:
        return False
    
    return True


# Initialize configuration validation
if not validate_configuration():
    raise ValueError("Configuration validation failed. Please check config values.")
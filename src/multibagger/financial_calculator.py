"""
Financial Calculator Module

This module implements comprehensive financial ratio calculations and metrics
for investment analysis based on extracted financial data.
"""

import numpy as np
from typing import Dict, List, Optional, Tuple, Any
import logging
import math

logger = logging.getLogger(__name__)


class FinancialCalculator:
    """
    Calculates financial ratios and metrics for investment analysis.
    """
    
    def __init__(self, financial_data: Dict[str, Any]):
        """
        Initialize calculator with extracted financial data.
        
        Args:
            financial_data (Dict[str, Any]): Extracted financial data from Excel
        """
        self.data = financial_data
        self.calculated_metrics = {}
        
    def safe_divide(self, numerator: float, denominator: float) -> float:
        """
        Safely divide two numbers, returning 0 if denominator is 0.
        
        Args:
            numerator (float): Numerator
            denominator (float): Denominator
            
        Returns:
            float: Division result or 0 if denominator is 0
        """
        if denominator == 0 or math.isnan(denominator) or math.isnan(numerator):
            return 0.0
        return numerator / denominator
    
    def calculate_cagr(self, values: Dict[int, float], years: Optional[int] = None) -> float:
        """
        Calculate Compound Annual Growth Rate (CAGR).
        
        Args:
            values (Dict[int, float]): Dictionary of year -> value
            years (Optional[int]): Number of years to consider (from end)
            
        Returns:
            float: CAGR percentage
        """
        if not values or len(values) < 2:
            return 0.0
        
        sorted_years = sorted(values.keys())
        
        if years:
            # Take last 'years' data points
            sorted_years = sorted_years[-years:]
        
        if len(sorted_years) < 2:
            return 0.0
        
        start_value = values[sorted_years[0]]
        end_value = values[sorted_years[-1]]
        num_years = sorted_years[-1] - sorted_years[0]
        
        if start_value <= 0 or end_value <= 0 or num_years <= 0:
            return 0.0
        
        try:
            cagr = (pow(end_value / start_value, 1 / num_years) - 1) * 100
            return round(cagr, 2)
        except (ValueError, ZeroDivisionError, OverflowError):
            return 0.0
    
    def calculate_growth_metrics(self) -> Dict[str, Any]:
        """
        Calculate various growth metrics.
        
        Returns:
            Dict[str, Any]: Growth metrics
        """
        metrics = {}
        
        # Extract P&L data
        pl_data = self.data.get('profit_loss', {})
        sales = pl_data.get('sales', {})
        net_profit = pl_data.get('net_profit', {})
        
        # Revenue CAGR for different periods
        metrics['revenue_cagr_3yr'] = self.calculate_cagr(sales, 3)
        metrics['revenue_cagr_5yr'] = self.calculate_cagr(sales, 5)
        metrics['revenue_cagr_10yr'] = self.calculate_cagr(sales, 10)
        
        # Profit CAGR
        metrics['profit_cagr_3yr'] = self.calculate_cagr(net_profit, 3)
        metrics['profit_cagr_5yr'] = self.calculate_cagr(net_profit, 5)
        metrics['profit_cagr_10yr'] = self.calculate_cagr(net_profit, 10)
        
        # EPS growth
        eps = pl_data.get('eps', {})
        metrics['eps_growth_rate'] = self.calculate_cagr(eps, 5)
        
        # Quarterly growth momentum
        quarterly_data = self.data.get('quarterly', {})
        if quarterly_data and 'revenue' in quarterly_data:
            q_revenue = quarterly_data['revenue']
            if len(q_revenue) >= 4:
                # Calculate QoQ growth for last 4 quarters
                qoq_growth = []
                for i in range(1, min(len(q_revenue), 5)):
                    if q_revenue[i-1] > 0:
                        growth = ((q_revenue[i] - q_revenue[i-1]) / q_revenue[i-1]) * 100
                        qoq_growth.append(growth)
                
                metrics['quarterly_growth_momentum'] = round(np.mean(qoq_growth), 2) if qoq_growth else 0.0
            else:
                metrics['quarterly_growth_momentum'] = 0.0
        else:
            metrics['quarterly_growth_momentum'] = 0.0
        
        return metrics
    
    def calculate_profitability_ratios(self) -> Dict[str, Any]:
        """
        Calculate profitability ratios.
        
        Returns:
            Dict[str, Any]: Profitability ratios
        """
        metrics = {}
        
        # Get latest year data
        pl_data = self.data.get('profit_loss', {})
        bs_data = self.data.get('balance_sheet', {})
        
        years = pl_data.get('years', [])
        if not years:
            return {key: 0.0 for key in ['operating_margin', 'net_profit_margin', 'roe', 'roce', 'roa', 'gross_margin']}
        
        latest_year = max(years)
        
        # Get values for latest year
        sales = pl_data.get('sales', {}).get(latest_year, 0)
        operating_profit = pl_data.get('operating_profit', {}).get(latest_year, 0)
        net_profit = pl_data.get('net_profit', {}).get(latest_year, 0)
        total_equity = bs_data.get('total_equity', {}).get(latest_year, 0)
        total_assets = bs_data.get('total_assets', {}).get(latest_year, 0)
        
        # Operating Margin
        metrics['operating_margin'] = round(self.safe_divide(operating_profit, sales) * 100, 2)
        
        # Net Profit Margin
        metrics['net_profit_margin'] = round(self.safe_divide(net_profit, sales) * 100, 2)
        
        # Return on Equity (ROE)
        metrics['roe'] = round(self.safe_divide(net_profit, total_equity) * 100, 2)
        
        # Return on Capital Employed (ROCE)
        total_debt = bs_data.get('total_debt', {}).get(latest_year, 0)
        capital_employed = total_equity + total_debt
        metrics['roce'] = round(self.safe_divide(operating_profit, capital_employed) * 100, 2)
        
        # Return on Assets (ROA)
        metrics['roa'] = round(self.safe_divide(net_profit, total_assets) * 100, 2)
        
        # Gross Margin (assuming operating profit as proxy)
        metrics['gross_margin'] = round(self.safe_divide(operating_profit * 1.5, sales) * 100, 2)  # Approximation
        
        return metrics
    
    def calculate_efficiency_ratios(self) -> Dict[str, Any]:
        """
        Calculate efficiency ratios.
        
        Returns:
            Dict[str, Any]: Efficiency ratios
        """
        metrics = {}
        
        pl_data = self.data.get('profit_loss', {})
        bs_data = self.data.get('balance_sheet', {})
        
        years = pl_data.get('years', [])
        if not years:
            return {key: 0.0 for key in ['asset_turnover', 'working_capital_turnover', 'inventory_turnover', 'receivables_days']}
        
        latest_year = max(years)
        
        # Get values
        sales = pl_data.get('sales', {}).get(latest_year, 0)
        total_assets = bs_data.get('total_assets', {}).get(latest_year, 0)
        current_assets = bs_data.get('current_assets', {}).get(latest_year, 0)
        current_liabilities = bs_data.get('current_liabilities', {}).get(latest_year, 0)
        
        # Asset Turnover
        metrics['asset_turnover'] = round(self.safe_divide(sales, total_assets), 2)
        
        # Working Capital Turnover
        working_capital = current_assets - current_liabilities
        metrics['working_capital_turnover'] = round(self.safe_divide(sales, working_capital), 2)
        
        # Inventory Turnover (approximation)
        inventory_approx = current_assets * 0.3  # Assume 30% of current assets is inventory
        metrics['inventory_turnover'] = round(self.safe_divide(sales, inventory_approx), 2)
        
        # Receivables Days (approximation)
        receivables_approx = current_assets * 0.4  # Assume 40% of current assets is receivables
        metrics['receivables_days'] = round(self.safe_divide(receivables_approx * 365, sales), 0)
        
        return metrics
    
    def calculate_leverage_ratios(self) -> Dict[str, Any]:
        """
        Calculate leverage ratios.
        
        Returns:
            Dict[str, Any]: Leverage ratios
        """
        metrics = {}
        
        pl_data = self.data.get('profit_loss', {})
        bs_data = self.data.get('balance_sheet', {})
        
        years = pl_data.get('years', [])
        if not years:
            return {key: 0.0 for key in ['debt_to_equity', 'debt_to_ebitda', 'interest_coverage', 'financial_leverage']}
        
        latest_year = max(years)
        
        # Get values
        total_debt = bs_data.get('total_debt', {}).get(latest_year, 0)
        total_equity = bs_data.get('total_equity', {}).get(latest_year, 0)
        operating_profit = pl_data.get('operating_profit', {}).get(latest_year, 0)
        net_profit = pl_data.get('net_profit', {}).get(latest_year, 0)
        total_assets = bs_data.get('total_assets', {}).get(latest_year, 0)
        
        # Debt to Equity
        metrics['debt_to_equity'] = round(self.safe_divide(total_debt, total_equity), 2)
        
        # Debt to EBITDA (approximation: EBITDA = Operating Profit * 1.2)
        ebitda_approx = operating_profit * 1.2
        metrics['debt_to_ebitda'] = round(self.safe_divide(total_debt, ebitda_approx), 2)
        
        # Interest Coverage (approximation: Interest = Debt * 0.08)
        interest_approx = total_debt * 0.08
        metrics['interest_coverage'] = round(self.safe_divide(operating_profit, interest_approx), 2)
        
        # Financial Leverage
        metrics['financial_leverage'] = round(self.safe_divide(total_assets, total_equity), 2)
        
        return metrics
    
    def calculate_liquidity_ratios(self) -> Dict[str, Any]:
        """
        Calculate liquidity ratios.
        
        Returns:
            Dict[str, Any]: Liquidity ratios
        """
        metrics = {}
        
        bs_data = self.data.get('balance_sheet', {})
        years = bs_data.get('years', [])
        
        if not years:
            return {key: 0.0 for key in ['current_ratio', 'quick_ratio', 'cash_ratio']}
        
        latest_year = max(years)
        
        # Get values
        current_assets = bs_data.get('current_assets', {}).get(latest_year, 0)
        current_liabilities = bs_data.get('current_liabilities', {}).get(latest_year, 0)
        
        # Current Ratio
        metrics['current_ratio'] = round(self.safe_divide(current_assets, current_liabilities), 2)
        
        # Quick Ratio (approximation: exclude inventory)
        quick_assets = current_assets * 0.7  # Assume 70% of current assets are liquid
        metrics['quick_ratio'] = round(self.safe_divide(quick_assets, current_liabilities), 2)
        
        # Cash Ratio (approximation)
        cash_equivalents = current_assets * 0.3  # Assume 30% of current assets is cash
        metrics['cash_ratio'] = round(self.safe_divide(cash_equivalents, current_liabilities), 2)
        
        return metrics
    
    def calculate_valuation_ratios(self) -> Dict[str, Any]:
        """
        Calculate valuation ratios.
        
        Returns:
            Dict[str, Any]: Valuation ratios
        """
        metrics = {}
        
        company_info = self.data.get('company_info', {})
        pl_data = self.data.get('profit_loss', {})
        bs_data = self.data.get('balance_sheet', {})
        
        current_price = company_info.get('current_price', 0)
        outstanding_shares = company_info.get('outstanding_shares', 0)
        market_cap = company_info.get('market_cap', 0)
        
        # If market cap not available, calculate it
        if market_cap == 0 and current_price > 0 and outstanding_shares > 0:
            market_cap = current_price * outstanding_shares
        
        years = pl_data.get('years', [])
        if not years:
            return {key: 0.0 for key in ['pe_ratio', 'pb_ratio', 'ev_ebitda', 'peg_ratio']}
        
        latest_year = max(years)
        
        # Get values
        net_profit = pl_data.get('net_profit', {}).get(latest_year, 0)
        eps = pl_data.get('eps', {}).get(latest_year, 0)
        total_equity = bs_data.get('total_equity', {}).get(latest_year, 0)
        operating_profit = pl_data.get('operating_profit', {}).get(latest_year, 0)
        total_debt = bs_data.get('total_debt', {}).get(latest_year, 0)
        
        # P/E Ratio
        if eps > 0:
            metrics['pe_ratio'] = round(self.safe_divide(current_price, eps), 2)
        else:
            metrics['pe_ratio'] = round(self.safe_divide(market_cap, net_profit), 2)
        
        # P/B Ratio
        book_value_per_share = self.safe_divide(total_equity, outstanding_shares) if outstanding_shares > 0 else 0
        metrics['pb_ratio'] = round(self.safe_divide(current_price, book_value_per_share), 2)
        
        # EV/EBITDA
        enterprise_value = market_cap + total_debt
        ebitda_approx = operating_profit * 1.2  # Approximation
        metrics['ev_ebitda'] = round(self.safe_divide(enterprise_value, ebitda_approx), 2)
        
        # PEG Ratio
        profit_growth = self.calculate_cagr(pl_data.get('net_profit', {}), 3)
        if profit_growth > 0:
            metrics['peg_ratio'] = round(self.safe_divide(metrics['pe_ratio'], profit_growth), 2)
        else:
            metrics['peg_ratio'] = 0.0
        
        return metrics
    
    def calculate_cash_flow_ratios(self) -> Dict[str, Any]:
        """
        Calculate cash flow ratios.
        
        Returns:
            Dict[str, Any]: Cash flow ratios
        """
        metrics = {}
        
        cf_data = self.data.get('cash_flow', {})
        pl_data = self.data.get('profit_loss', {})
        
        years = cf_data.get('years', [])
        if not years:
            return {key: 0.0 for key in ['ocf_to_net_profit', 'fcf_to_revenue', 'cash_conversion_cycle']}
        
        latest_year = max(years)
        
        # Get values
        operating_cash_flow = cf_data.get('operating_cash_flow', {}).get(latest_year, 0)
        free_cash_flow = cf_data.get('free_cash_flow', {}).get(latest_year, 0)
        net_profit = pl_data.get('net_profit', {}).get(latest_year, 0)
        sales = pl_data.get('sales', {}).get(latest_year, 0)
        
        # OCF to Net Profit
        metrics['ocf_to_net_profit'] = round(self.safe_divide(operating_cash_flow, net_profit), 2)
        
        # FCF to Revenue
        metrics['fcf_to_revenue'] = round(self.safe_divide(free_cash_flow, sales) * 100, 2)
        
        # Cash Conversion Cycle (approximation)
        metrics['cash_conversion_cycle'] = 58  # Default approximation
        
        return metrics
    
    def calculate_investment_score(self) -> Dict[str, Any]:
        """
        Calculate overall investment score based on scoring framework.
        
        Returns:
            Dict[str, Any]: Investment scoring details
        """
        # Get all calculated metrics
        growth_metrics = self.calculate_growth_metrics()
        profitability_ratios = self.calculate_profitability_ratios()
        leverage_ratios = self.calculate_leverage_ratios()
        cash_flow_ratios = self.calculate_cash_flow_ratios()
        valuation_ratios = self.calculate_valuation_ratios()
        
        scoring = {}
        
        # Growth Quality (20 points)
        revenue_cagr = growth_metrics.get('revenue_cagr_5yr', 0)
        if revenue_cagr > 15:
            scoring['growth_quality'] = 20
        elif revenue_cagr > 10:
            scoring['growth_quality'] = 15
        elif revenue_cagr > 5:
            scoring['growth_quality'] = 10
        else:
            scoring['growth_quality'] = 0
        
        # Profitability (20 points)
        npm = profitability_ratios.get('net_profit_margin', 0)
        if npm > 15:
            scoring['profitability'] = 20
        elif npm > 10:
            scoring['profitability'] = 15
        elif npm > 5:
            scoring['profitability'] = 10
        else:
            scoring['profitability'] = 0
        
        # Financial Health (20 points)
        de_ratio = leverage_ratios.get('debt_to_equity', 0)
        current_ratio = self.calculate_liquidity_ratios().get('current_ratio', 0)
        
        if de_ratio < 0.5 and current_ratio > 2:
            scoring['financial_health'] = 20
        elif de_ratio < 1.0 and current_ratio > 1.5:
            scoring['financial_health'] = 15
        elif de_ratio < 2.0 and current_ratio > 1:
            scoring['financial_health'] = 10
        else:
            scoring['financial_health'] = 0
        
        # Cash Flow Quality (20 points)
        ocf_np_ratio = cash_flow_ratios.get('ocf_to_net_profit', 0)
        fcf_revenue = cash_flow_ratios.get('fcf_to_revenue', 0)
        
        if ocf_np_ratio > 1.0 and fcf_revenue > 5:
            scoring['cash_flow_quality'] = 20
        elif ocf_np_ratio > 0.8:
            scoring['cash_flow_quality'] = 15
        elif ocf_np_ratio > 0.6:
            scoring['cash_flow_quality'] = 10
        else:
            scoring['cash_flow_quality'] = 0
        
        # Valuation (20 points)
        pe_ratio = valuation_ratios.get('pe_ratio', 0)
        
        if pe_ratio < 15 and revenue_cagr > 10:
            scoring['valuation'] = 20
        elif pe_ratio < 25:
            scoring['valuation'] = 15
        elif pe_ratio < 35:
            scoring['valuation'] = 10
        else:
            scoring['valuation'] = 5
        
        # Total score
        total_score = sum(scoring.values())
        
        # Recommendation
        if total_score >= 70:
            recommendation = "STRONG BUY"
            confidence = "HIGH"
        elif total_score >= 50:
            recommendation = "BUY"
            confidence = "MEDIUM"
        elif total_score >= 30:
            recommendation = "HOLD"
            confidence = "MEDIUM"
        else:
            recommendation = "AVOID"
            confidence = "HIGH"
        
        return {
            'total_score': total_score,
            'category_scores': scoring,
            'recommendation': recommendation,
            'confidence_level': confidence
        }
    
    def generate_investment_thesis(self) -> Dict[str, Any]:
        """
        Generate investment thesis with bull/bear cases and risks.
        
        Returns:
            Dict[str, Any]: Investment thesis
        """
        growth_metrics = self.calculate_growth_metrics()
        profitability_ratios = self.calculate_profitability_ratios()
        leverage_ratios = self.calculate_leverage_ratios()
        cash_flow_ratios = self.calculate_cash_flow_ratios()
        
        bull_case = []
        bear_case = []
        key_risks = []
        
        # Bull case points
        if growth_metrics.get('revenue_cagr_5yr', 0) > 10:
            bull_case.append(f"Strong revenue growth with {growth_metrics['revenue_cagr_5yr']:.1f}% CAGR over 5 years")
        
        if profitability_ratios.get('net_profit_margin', 0) > 10:
            bull_case.append(f"Healthy profit margins at {profitability_ratios['net_profit_margin']:.1f}%")
        
        if cash_flow_ratios.get('ocf_to_net_profit', 0) > 0.8:
            bull_case.append(f"Strong cash generation with OCF/NP ratio of {cash_flow_ratios['ocf_to_net_profit']:.2f}")
        
        if leverage_ratios.get('debt_to_equity', 0) < 1.0:
            bull_case.append(f"Conservative leverage with D/E ratio of {leverage_ratios['debt_to_equity']:.2f}")
        
        # Bear case points
        if growth_metrics.get('revenue_cagr_5yr', 0) < 5:
            bear_case.append("Slow revenue growth indicating maturity or challenges")
        
        if profitability_ratios.get('net_profit_margin', 0) < 5:
            bear_case.append("Low profit margins indicating competitive pressure")
        
        if leverage_ratios.get('debt_to_equity', 0) > 2.0:
            bear_case.append("High leverage increases financial risk")
        
        # Key risks
        key_risks.extend([
            "Market competition and competitive positioning",
            "Economic downturn impact on business performance",
            "Regulatory changes affecting industry dynamics",
            "Management execution and capital allocation decisions"
        ])
        
        return {
            'bull_case_points': bull_case,
            'bear_case_points': bear_case,
            'key_risks': key_risks
        }
    
    def calculate_all_metrics(self) -> Dict[str, Any]:
        """
        Calculate all financial metrics and ratios.
        
        Returns:
            Dict[str, Any]: Complete set of calculated metrics
        """
        self.calculated_metrics = {
            'growth_metrics': self.calculate_growth_metrics(),
            'profitability_ratios': self.calculate_profitability_ratios(),
            'efficiency_ratios': self.calculate_efficiency_ratios(),
            'leverage_ratios': self.calculate_leverage_ratios(),
            'liquidity_ratios': self.calculate_liquidity_ratios(),
            'valuation_ratios': self.calculate_valuation_ratios(),
            'cash_flow_ratios': self.calculate_cash_flow_ratios(),
            'investment_score': self.calculate_investment_score(),
            'investment_thesis': self.generate_investment_thesis()
        }
        
        logger.info("All financial metrics calculated successfully")
        return self.calculated_metrics


def calculate_financial_metrics(financial_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Convenience function to calculate all financial metrics.
    
    Args:
        financial_data (Dict[str, Any]): Extracted financial data
        
    Returns:
        Dict[str, Any]: Calculated financial metrics
    """
    calculator = FinancialCalculator(financial_data)
    return calculator.calculate_all_metrics()
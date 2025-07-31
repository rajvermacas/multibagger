#!/usr/bin/env python3
"""
Comprehensive Batch Stock Analysis Script

This script processes all Excel files in the stocks directory and generates
both JSON analysis files and comprehensive markdown investment reports.
"""

import os
import sys
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from multibagger.stock_analyzer import analyze_stock_workbook, get_analysis_summary
from multibagger.utils import setup_logging


class BatchStockAnalyzer:
    """Handles batch processing of multiple stock Excel files."""
    
    def __init__(self, stocks_directory: str, reports_directory: str):
        """
        Initialize batch analyzer.
        
        Args:
            stocks_directory: Path to directory containing Excel files
            reports_directory: Path to directory for saving reports
        """
        self.stocks_directory = Path(stocks_directory)
        self.reports_directory = Path(reports_directory)
        self.analysis_log = []
        self.successful_analyses = []
        self.failed_analyses = []
        
        # Setup logging
        setup_logging("INFO")
        self.logger = logging.getLogger(__name__)
        
        # Ensure reports directory exists
        self.reports_directory.mkdir(parents=True, exist_ok=True)
        
        # Create analysis log file
        self.log_file = self.reports_directory / "analysis_log.txt"
        
    def get_excel_files(self) -> List[Path]:
        """Get all Excel files from the stocks directory."""
        excel_files = []
        for file_path in self.stocks_directory.glob("*.xlsx"):
            if not file_path.name.startswith("~"):  # Skip temp files
                excel_files.append(file_path)
        
        return sorted(excel_files)
    
    def run_analysis_on_file(self, excel_path: Path) -> Optional[str]:
        """
        Run analysis on a single Excel file.
        
        Args:
            excel_path: Path to Excel file
            
        Returns:
            Path to generated JSON file or None if failed
        """
        try:
            self.logger.info(f"Starting analysis for: {excel_path.name}")
            
            # Run the stock analysis
            json_path = analyze_stock_workbook(str(excel_path), "INFO")
            
            if json_path:
                self.logger.info(f"Analysis successful for {excel_path.name}")
                self.successful_analyses.append({
                    'excel_file': excel_path.name,
                    'json_file': json_path,
                    'timestamp': datetime.now().isoformat()
                })
                return json_path
            else:
                self.logger.error(f"Analysis failed for {excel_path.name}")
                self.failed_analyses.append({
                    'excel_file': excel_path.name,
                    'error': 'Analysis returned None',
                    'timestamp': datetime.now().isoformat()
                })
                return None
                
        except Exception as e:
            error_msg = f"Critical error analyzing {excel_path.name}: {str(e)}"
            self.logger.error(error_msg)
            self.failed_analyses.append({
                'excel_file': excel_path.name,
                'error': error_msg,
                'timestamp': datetime.now().isoformat()
            })
            return None
    
    def create_markdown_report(self, json_path: str) -> Optional[str]:
        """
        Create comprehensive markdown investment report from JSON data.
        
        Args:
            json_path: Path to JSON analysis file
            
        Returns:
            Path to markdown report or None if failed
        """
        try:
            # Load JSON data
            with open(json_path, 'r') as f:
                data = json.load(f)
            
            company_name = data.get('company_info', {}).get('name', 'Unknown_Company')
            timestamp = datetime.now().strftime('%H%M%S')
            
            # Clean company name for filename
            clean_name = company_name.upper().replace(' ', '_').replace('.', '').replace('-', '_')
            markdown_filename = f"{clean_name}_Investment_Report_{timestamp}.md"
            markdown_path = self.reports_directory / markdown_filename
            
            # Generate markdown content
            markdown_content = self._generate_markdown_content(data)
            
            # Save markdown file
            with open(markdown_path, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            
            self.logger.info(f"Markdown report created: {markdown_path}")
            return str(markdown_path)
            
        except Exception as e:
            self.logger.error(f"Error creating markdown report for {json_path}: {str(e)}")
            return None
    
    def _generate_markdown_content(self, data: Dict[str, Any]) -> str:
        """Generate comprehensive markdown report content."""
        
        # Extract key data
        company_info = data.get('company_info', {})
        metrics = data.get('calculated_metrics', {})
        investment_score = data.get('investment_score', {})
        historical = data.get('historical_data', {})
        thesis = data.get('investment_thesis', {})
        metadata = data.get('analysis_metadata', {})
        
        # Helper function to safely get values
        def safe_get(d, key, default=0):
            return d.get(key, default) if d else default
        
        def format_currency(value):
            """Format currency values in Crores."""
            if value == 0:
                return "₹0.00 Cr"
            return f"₹{value/100:.2f} Cr" if value >= 100 else f"₹{value:.2f} Cr"
        
        def format_percentage(value):
            """Format percentage values."""
            return f"{value:.2f}%" if value != 0 else "0.00%"
        
        def get_quality_rating(value, thresholds):
            """Get quality rating based on thresholds."""
            if value >= thresholds.get('excellent', 20):
                return "Excellent"
            elif value >= thresholds.get('good', 15):
                return "Good"
            elif value >= thresholds.get('fair', 10):
                return "Fair"
            else:
                return "Poor"
        
        # Start building markdown content
        markdown = f"""# {company_info.get('name', 'Unknown Company')} - Investment Analysis Report

## Executive Summary

- **Investment Recommendation**: {investment_score.get('recommendation', 'UNKNOWN')}
- **Investment Score**: {investment_score.get('total_score', 0)}/100
- **Confidence Level**: {investment_score.get('confidence_level', 'LOW')}
- **Analysis Date**: {company_info.get('analysis_date', 'Unknown')}
- **Investment Horizon**: 3-5 years

## Company Overview

- **Company**: {company_info.get('name', 'Unknown Company')}
- **Current Price**: ₹{company_info.get('current_price', 0):.2f}
- **Market Cap**: {format_currency(company_info.get('market_cap', 0))}
- **Face Value**: ₹{company_info.get('face_value', 0):.2f}
- **Outstanding Shares**: {company_info.get('outstanding_shares', 0):.2f} Cr

## Financial Performance Analysis

### 📈 Growth Metrics

"""
        
        # Growth metrics table
        growth = metrics.get('growth_metrics', {})
        markdown += """| Metric       | 3-Year | 5-Year | 10-Year | Assessment |
|--------------|--------|--------|---------|------------|
"""
        
        revenue_3yr = safe_get(growth, 'revenue_cagr_3yr')
        revenue_5yr = safe_get(growth, 'revenue_cagr_5yr')
        revenue_10yr = safe_get(growth, 'revenue_cagr_10yr')
        profit_3yr = safe_get(growth, 'profit_cagr_3yr')
        profit_5yr = safe_get(growth, 'profit_cagr_5yr')
        profit_10yr = safe_get(growth, 'profit_cagr_10yr')
        eps_3yr = safe_get(growth, 'eps_cagr_3yr')
        eps_5yr = safe_get(growth, 'eps_cagr_5yr')
        eps_10yr = safe_get(growth, 'eps_cagr_10yr')
        
        markdown += f"""| Revenue CAGR | {format_percentage(revenue_3yr)} | {format_percentage(revenue_5yr)} | {format_percentage(revenue_10yr)} | {get_quality_rating(revenue_5yr, {'excellent': 15, 'good': 10, 'fair': 5})} |
| Profit CAGR  | {format_percentage(profit_3yr)} | {format_percentage(profit_5yr)} | {format_percentage(profit_10yr)} | {get_quality_rating(profit_5yr, {'excellent': 20, 'good': 15, 'fair': 10})} |
| EPS Growth   | {format_percentage(eps_3yr)} | {format_percentage(eps_5yr)} | {format_percentage(eps_10yr)} | {get_quality_rating(eps_5yr, {'excellent': 15, 'good': 10, 'fair': 5})} |

**Growth Analysis**: """
        
        # Add growth analysis
        if revenue_5yr > 15:
            markdown += f"The company demonstrates strong revenue growth with a 5-year CAGR of {format_percentage(revenue_5yr)}. "
        elif revenue_5yr > 10:
            markdown += f"The company shows moderate revenue growth with a 5-year CAGR of {format_percentage(revenue_5yr)}. "
        else:
            markdown += f"Revenue growth has been modest with a 5-year CAGR of {format_percentage(revenue_5yr)}. "
        
        if profit_5yr > revenue_5yr and profit_5yr > 20:
            markdown += "Profit growth is outpacing revenue growth, indicating improving operational efficiency and margin expansion."
        elif profit_5yr > 15:
            markdown += "Profit growth is healthy, suggesting good operational management."
        else:
            markdown += "Profit growth needs attention as it may indicate margin pressure or operational challenges."
        
        # Profitability section
        profitability = metrics.get('profitability_ratios', {})
        markdown += f"""

### 💰 Profitability Ratios

| Metric            | Current | Trend   | Benchmark | Status           |
|-------------------|---------|---------|-----------|------------------|
"""
        
        op_margin = safe_get(profitability, 'operating_margin')
        net_margin = safe_get(profitability, 'net_profit_margin')
        roe = safe_get(profitability, 'roe')
        roce = safe_get(profitability, 'roce')
        
        def get_trend_arrow(current, historical_avg):
            if current > historical_avg * 1.1:
                return "↑"
            elif current < historical_avg * 0.9:
                return "↓"
            else:
                return "→"
        
        def get_status(value, benchmark):
            if value >= benchmark:
                return "Good"
            elif value >= benchmark * 0.7:
                return "Fair"
            else:
                return "Poor"
        
        markdown += f"""| Operating Margin  | {format_percentage(op_margin)} | → | >15% | {get_status(op_margin, 15)} |
| Net Profit Margin | {format_percentage(net_margin)} | → | >10% | {get_status(net_margin, 10)} |
| ROE               | {format_percentage(roe)} | → | >15% | {get_status(roe, 15)} |
| ROCE              | {format_percentage(roce)} | → | >15% | {get_status(roce, 15)} |

**Profitability Analysis**: """
        
        if op_margin > 15 and net_margin > 10:
            markdown += "The company demonstrates strong profitability with healthy operating and net margins. "
        elif op_margin > 10 or net_margin > 5:
            markdown += "Profitability metrics are moderate and within acceptable ranges. "
        else:
            markdown += "Profitability metrics are below industry benchmarks and need improvement. "
        
        if roe > 15 and roce > 15:
            markdown += "Returns to shareholders and capital employed are excellent, indicating efficient capital allocation."
        else:
            markdown += "Returns could be improved through better capital allocation and operational efficiency."
        
        # Balance sheet strength
        leverage = metrics.get('leverage_ratios', {})
        liquidity = metrics.get('liquidity_ratios', {})
        
        debt_equity = safe_get(leverage, 'debt_to_equity')
        current_ratio = safe_get(liquidity, 'current_ratio')
        interest_coverage = safe_get(leverage, 'interest_coverage_ratio')
        
        markdown += f"""

### 🏦 Balance Sheet Strength

| Metric            | Current | Benchmark | Assessment                 |
|-------------------|---------|-----------|----------------------------|
| Debt/Equity       | {debt_equity:.2f} | <1.0      | {get_quality_rating(1/max(debt_equity, 0.1) * 10, {'excellent': 10, 'good': 5, 'fair': 2})} |
| Current Ratio     | {current_ratio:.2f} | >2.0      | {get_quality_rating(current_ratio * 10, {'excellent': 25, 'good': 20, 'fair': 15})} |
| Interest Coverage | {interest_coverage:.2f} | >5.0      | {get_quality_rating(min(interest_coverage, 20), {'excellent': 10, 'good': 5, 'fair': 2})} |

**Financial Health**: """
        
        if debt_equity < 0.5:
            markdown += "The company maintains a conservative debt profile with low leverage. "
        elif debt_equity < 1.0:
            markdown += "Debt levels are manageable and within acceptable limits. "
        else:
            markdown += "High debt levels may pose financial risks and limit flexibility. "
        
        if current_ratio > 2.0:
            markdown += "Strong liquidity position ensures ability to meet short-term obligations."
        elif current_ratio > 1.5:
            markdown += "Adequate liquidity but could be strengthened."
        else:
            markdown += "Liquidity concerns as current ratio is below recommended levels."
        
        # Cash flow analysis
        cash_flow = metrics.get('cash_flow_ratios', {})
        
        ocf_net_ratio = safe_get(cash_flow, 'ocf_to_net_income')
        fcf_revenue = safe_get(cash_flow, 'fcf_to_revenue') * 100  # Convert to percentage
        
        markdown += f"""

### 💸 Cash Flow Analysis

| Metric          | Value    | Quality                    |
|-----------------|----------|----------------------------|
| OCF/Net Profit  | {ocf_net_ratio:.2f} | {get_quality_rating(ocf_net_ratio * 10, {'excellent': 12, 'good': 10, 'fair': 8})} |
| FCF/Revenue     | {format_percentage(fcf_revenue)} | {get_quality_rating(fcf_revenue, {'excellent': 10, 'good': 5, 'fair': 2})} |
| Cash Conversion | Data not available | Unable to calculate |

**Cash Flow Quality**: """
        
        if ocf_net_ratio > 1.2:
            markdown += "Excellent cash conversion with operating cash flow exceeding net profit. "
        elif ocf_net_ratio > 1.0:
            markdown += "Good cash conversion quality indicating healthy business operations. "
        else:
            markdown += "Cash conversion needs attention as operating cash flow is below net profit levels. "
        
        if fcf_revenue > 10:
            markdown += "Strong free cash flow generation provides flexibility for growth investments and dividends."
        elif fcf_revenue > 5:
            markdown += "Moderate free cash flow generation supports business operations."
        else:
            markdown += "Limited free cash flow generation may constrain growth opportunities."
        
        # Valuation metrics
        valuation = metrics.get('valuation_ratios', {})
        
        pe_ratio = safe_get(valuation, 'pe_ratio')
        pb_ratio = safe_get(valuation, 'pb_ratio')
        peg_ratio = safe_get(valuation, 'peg_ratio')
        
        markdown += f"""

### 📊 Valuation Metrics

| Metric    | Current | Historical Range | Assessment                  |
|-----------|---------|------------------|-----------------------------|
| P/E Ratio | {pe_ratio:.2f} | Data not available | {'Cheap' if pe_ratio < 15 else 'Fair' if pe_ratio < 25 else 'Expensive'} |
| P/B Ratio | {pb_ratio:.2f} | Data not available | {'Cheap' if pb_ratio < 2 else 'Fair' if pb_ratio < 4 else 'Expensive'} |
| PEG Ratio | {peg_ratio:.2f} | <1.0 Attractive  | {'Attractive' if peg_ratio < 1 else 'Fair' if peg_ratio < 2 else 'Expensive'} |

## Investment Scoring Breakdown

"""
        
        # Investment scoring table
        category_scores = investment_score.get('category_scores', {})
        
        def score_to_performance(score, max_score):
            percentage = (score / max_score) * 100 if max_score > 0 else 0
            if percentage >= 80:
                return "Excellent"
            elif percentage >= 60:
                return "Good"
            elif percentage >= 40:
                return "Fair"
            else:
                return "Poor"
        
        markdown += f"""| Category          | Score | Max | Performance                |
|-------------------|-------|-----|----------------------------|
| Growth Quality    | {safe_get(category_scores, 'growth_score'):.0f} | 20  | {score_to_performance(safe_get(category_scores, 'growth_score'), 20)} |
| Profitability     | {safe_get(category_scores, 'profitability_score'):.0f} | 20  | {score_to_performance(safe_get(category_scores, 'profitability_score'), 20)} |
| Financial Health  | {safe_get(category_scores, 'financial_health_score'):.0f} | 20  | {score_to_performance(safe_get(category_scores, 'financial_health_score'), 20)} |
| Cash Flow Quality | {safe_get(category_scores, 'cash_flow_score'):.0f} | 20  | {score_to_performance(safe_get(category_scores, 'cash_flow_score'), 20)} |
| Valuation         | {safe_get(category_scores, 'valuation_score'):.0f} | 20  | {score_to_performance(safe_get(category_scores, 'valuation_score'), 20)} |
| **Total Score**   | **{investment_score.get('total_score', 0):.0f}** | **100** | **{score_to_performance(investment_score.get('total_score', 0), 100)}** |

## Investment Thesis

### 🟢 Bull Case (Reasons to Invest)

"""
        
        # Add bull points
        bull_points = thesis.get('bull_points', []) if thesis else []
        if bull_points:
            for point in bull_points:
                markdown += f"- {point}\n"
        else:
            # Generate bull points based on metrics
            if revenue_5yr > 15:
                markdown += f"- Strong revenue growth with 5-year CAGR of {format_percentage(revenue_5yr)}\n"
            if roe > 15:
                markdown += f"- Excellent return on equity of {format_percentage(roe)} indicates efficient management\n"
            if debt_equity < 0.5:
                markdown += f"- Conservative debt profile with D/E ratio of {debt_equity:.2f}\n"
            if op_margin > 15:
                markdown += f"- Strong operating margins of {format_percentage(op_margin)} show operational efficiency\n"
        
        markdown += """
### 🔴 Bear Case (Key Concerns)

"""
        
        # Add bear points
        bear_points = thesis.get('bear_points', []) if thesis else []
        if bear_points:
            for point in bear_points:
                markdown += f"- {point}\n"
        else:
            # Generate bear points based on metrics
            if revenue_5yr < 5:
                markdown += f"- Slow revenue growth with 5-year CAGR of only {format_percentage(revenue_5yr)}\n"
            if net_margin < 5:
                markdown += f"- Low net profit margins of {format_percentage(net_margin)} indicate profitability challenges\n"
            if debt_equity > 1.0:
                markdown += f"- High debt levels with D/E ratio of {debt_equity:.2f} may limit flexibility\n"
            if pe_ratio > 30:
                markdown += f"- High P/E ratio of {pe_ratio:.2f} suggests expensive valuation\n"
        
        markdown += """
### ⚖️ Key Risks

"""
        
        # Add risk factors
        risk_factors = thesis.get('risk_factors', []) if thesis else []
        if risk_factors:
            for risk in risk_factors:
                markdown += f"- {risk}\n"
        else:
            # Generate generic risk factors
            markdown += "- Market volatility and economic downturns\n"
            markdown += "- Industry-specific regulatory changes\n"
            markdown += "- Competition from established and new players\n"
            markdown += "- Execution risks in growth initiatives\n"
        
        # Historical performance
        years = historical.get('years', [])
        revenues = historical.get('revenue', [])
        
        markdown += """
## Historical Performance

### Revenue Trend (Last 10 Years)

"""
        
        if years and revenues and len(years) > 1:
            markdown += "| Year | Revenue (₹ Cr) | YoY Growth |\n"
            markdown += "|------|----------------|------------|\n"
            
            for i, year in enumerate(years):
                revenue = revenues[i] if i < len(revenues) else 0
                if i > 0 and revenues[i-1] > 0:
                    growth = ((revenue - revenues[i-1]) / revenues[i-1]) * 100
                    markdown += f"| {year} | {revenue/100:.2f} | {growth:.1f}% |\n"
                else:
                    markdown += f"| {year} | {revenue/100:.2f} | - |\n"
        else:
            markdown += "Historical revenue data not available.\n"
        
        # Final recommendation
        total_score = investment_score.get('total_score', 0)
        recommendation = investment_score.get('recommendation', 'UNKNOWN')
        
        markdown += f"""
## Final Investment Recommendation

### 🎯 Recommendation: {recommendation}

**Rationale**: """
        
        if total_score >= 70:
            markdown += f"With an investment score of {total_score}/100, this stock shows strong fundamentals across multiple metrics. "
        elif total_score >= 50:
            markdown += f"With an investment score of {total_score}/100, this stock shows mixed fundamentals with both strengths and areas for improvement. "
        else:
            markdown += f"With an investment score of {total_score}/100, this stock shows weak fundamentals and significant risks. "
        
        # Position sizing recommendations
        markdown += f"""

**Position Sizing**:
"""
        
        if total_score >= 70:
            markdown += "- **Full Position**: Consider for 3-5% portfolio allocation\n"
            markdown += "- **Buy on dips**: Suitable for systematic investment\n"
        elif total_score >= 50:
            markdown += "- **Partial Position**: Consider for 1-2% portfolio allocation\n"
            markdown += "- **Wait for better entry**: Monitor for improvement in metrics\n"
        else:
            markdown += "- **Avoid**: Not recommended for investment at current levels\n"
            markdown += "- **Watch list only**: Monitor for significant improvements\n"
        
        markdown += f"""
**Entry Strategy**: """
        
        if pe_ratio > 0:
            if pe_ratio < 15:
                markdown += "Current valuation appears attractive for entry.\n"
            elif pe_ratio < 25:
                markdown += "Fair valuation - consider dollar-cost averaging.\n"
            else:
                markdown += "Expensive valuation - wait for price correction.\n"
        else:
            markdown += "Valuation assessment limited due to lack of earnings data.\n"
        
        markdown += f"""
**Key Monitoring Points**:
- Revenue growth trajectory and market share trends
- Margin expansion or contraction patterns  
- Debt reduction progress and financial leverage
- Cash flow generation and capital allocation
- Industry dynamics and competitive positioning

**Price Targets & Risk Management**:
- **Target Return**: Based on score and growth potential
- **Stop Loss**: Set at 20-25% below entry price
- **Time Horizon**: 3-5 years for full thesis to play out

## Data Quality & Limitations

**Data Quality Assessment**: {metadata.get('data_quality', 'Unknown')}

**Missing Data Points**: """
        
        missing_data = metadata.get('missing_data_points', [])
        if missing_data:
            for missing in missing_data:
                markdown += f"- {missing}\n"
        else:
            markdown += "- None identified\n"
        
        markdown += f"""
**Analysis Limitations**: 
- Analysis based on historical financial data only
- Market sentiment and qualitative factors not included
- Future projections are estimates based on past performance
- Recommend independent research before investment decisions

---

*Report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} using automated financial analysis system.*
*This report is for informational purposes only and should not be considered as investment advice.*
"""
        
        return markdown
    
    def run_batch_analysis(self) -> Dict[str, Any]:
        """
        Run complete batch analysis on all Excel files.
        
        Returns:
            Dictionary with analysis results and statistics
        """
        self.logger.info("Starting batch stock analysis...")
        
        # Get all Excel files
        excel_files = self.get_excel_files()
        total_files = len(excel_files)
        
        self.logger.info(f"Found {total_files} Excel files to analyze")
        
        if total_files == 0:
            self.logger.warning("No Excel files found in the directory")
            return {
                'total_files': 0,
                'successful_analyses': 0,
                'failed_analyses': 0,
                'success_rate': 0,
                'reports_generated': []
            }
        
        # Process each file
        markdown_reports = []
        
        for i, excel_file in enumerate(excel_files, 1):
            self.logger.info(f"Processing file {i}/{total_files}: {excel_file.name}")
            
            # Run stock analysis
            json_path = self.run_analysis_on_file(excel_file)
            
            if json_path:
                # Create markdown report
                markdown_path = self.create_markdown_report(json_path)
                if markdown_path:
                    markdown_reports.append(markdown_path)
            
            # Log progress
            self.log_analysis_progress(i, total_files)
        
        # Create summary statistics
        successful_count = len(self.successful_analyses)
        failed_count = len(self.failed_analyses)
        success_rate = (successful_count / total_files) * 100 if total_files > 0 else 0
        
        # Save analysis log
        self.save_analysis_log()
        
        results = {
            'total_files': total_files,
            'successful_analyses': successful_count,
            'failed_analyses': failed_count,
            'success_rate': success_rate,
            'reports_generated': len(markdown_reports),
            'successful_files': [item['excel_file'] for item in self.successful_analyses],
            'failed_files': [item['excel_file'] for item in self.failed_analyses],
            'markdown_reports': markdown_reports
        }
        
        self.logger.info(f"Batch analysis completed. Success rate: {success_rate:.1f}%")
        return results
    
    def log_analysis_progress(self, current: int, total: int):
        """Log analysis progress."""
        progress = (current / total) * 100
        self.logger.info(f"Progress: {current}/{total} files processed ({progress:.1f}%)")
    
    def save_analysis_log(self):
        """Save detailed analysis log to file."""
        try:
            log_content = f"""Stock Analysis Batch Log
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

SUMMARY:
- Total files processed: {len(self.successful_analyses) + len(self.failed_analyses)}
- Successful analyses: {len(self.successful_analyses)}
- Failed analyses: {len(self.failed_analyses)}
- Success rate: {(len(self.successful_analyses) / (len(self.successful_analyses) + len(self.failed_analyses)) * 100):.1f}%

SUCCESSFUL ANALYSES:
"""
            
            for success in self.successful_analyses:
                log_content += f"✓ {success['excel_file']} -> {success['json_file']} ({success['timestamp']})\n"
            
            log_content += "\nFAILED ANALYSES:\n"
            
            for failure in self.failed_analyses:
                log_content += f"✗ {failure['excel_file']} - {failure['error']} ({failure['timestamp']})\n"
            
            with open(self.log_file, 'w') as f:
                f.write(log_content)
            
            self.logger.info(f"Analysis log saved to: {self.log_file}")
            
        except Exception as e:
            self.logger.error(f"Error saving analysis log: {str(e)}")


def main():
    """Main function to run batch analysis."""
    # Set up directories
    stocks_dir = "/root/projects/Multibagger/resources/Stocks/2025-07-31"
    reports_dir = "/root/projects/Multibagger/resources/reports/2025-07-31"
    
    # Create analyzer
    analyzer = BatchStockAnalyzer(stocks_dir, reports_dir)
    
    # Run batch analysis
    results = analyzer.run_batch_analysis()
    
    # Print summary
    print("\n" + "="*60)
    print("BATCH STOCK ANALYSIS SUMMARY")
    print("="*60)
    print(f"Total files processed: {results['total_files']}")
    print(f"Successful analyses: {results['successful_analyses']}")
    print(f"Failed analyses: {results['failed_analyses']}")
    print(f"Success rate: {results['success_rate']:.1f}%")
    print(f"Markdown reports generated: {results['reports_generated']}")
    
    if results['failed_files']:
        print(f"\nFailed files:")
        for failed_file in results['failed_files']:
            print(f"  - {failed_file}")
    
    print(f"\nReports saved to: {reports_dir}")
    print("="*60)
    
    return results


if __name__ == "__main__":
    main()
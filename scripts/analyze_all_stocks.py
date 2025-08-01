#!/usr/bin/env python3
"""
Script to analyze all stock Excel files systematically
"""

import os
import sys
import time
import logging
import argparse
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from multibagger.stock_analyzer import analyze_stock_workbook

def setup_logging():
    """Setup logging for the analysis process"""
    log_dir = Path(__file__).parent.parent / "logs"
    log_dir.mkdir(exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_dir / "bulk_analysis.log"),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(description="Analyze all stock Excel files in a directory")
    parser.add_argument(
        "stocks_dir",
        help="Path to directory containing stock Excel files"
    )
    return parser.parse_args()

def main():
    """Main function to analyze all stock files"""
    args = parse_arguments()
    logger = setup_logging()
    
    # Define paths from command line argument
    stocks_dir = Path(args.stocks_dir)
    
    # Validate directory exists
    if not stocks_dir.exists():
        logger.error(f"Directory does not exist: {stocks_dir}")
        sys.exit(1)
    
    if not stocks_dir.is_dir():
        logger.error(f"Path is not a directory: {stocks_dir}")
        sys.exit(1)
    
    # List all Excel files
    excel_files = list(stocks_dir.glob("*.xlsx"))
    logger.info(f"Found {len(excel_files)} Excel files to analyze")
    
    # Track results
    successful = []
    failed = []
    
    for excel_file in sorted(excel_files):
        logger.info(f"\n{'='*50}")
        logger.info(f"Analyzing: {excel_file.name}")
        logger.info(f"{'='*50}")
        
        try:
            # Run analysis - returns file path or None
            result_path = analyze_stock_workbook(str(excel_file))
            
            if result_path:
                successful.append(excel_file.name)
                logger.info(f"✅ Successfully analyzed: {excel_file.name}")
                logger.info(f"📄 Report saved to: {result_path}")
            else:
                failed.append({
                    'file': excel_file.name,
                    'error': 'Analysis returned None - check logs for details'
                })
                logger.error(f"❌ Failed to analyze: {excel_file.name}")
                
        except Exception as e:
            failed.append({
                'file': excel_file.name,
                'error': str(e)
            })
            logger.error(f"❌ Exception analyzing {excel_file.name}: {str(e)}")
        
        # Small delay between analyses
        time.sleep(1)
    
    # Print summary
    logger.info(f"\n{'='*60}")
    logger.info("ANALYSIS SUMMARY")
    logger.info(f"{'='*60}")
    logger.info(f"Total files: {len(excel_files)}")
    logger.info(f"Successfully analyzed: {len(successful)}")
    logger.info(f"Failed: {len(failed)}")
    
    if successful:
        logger.info(f"\n✅ SUCCESSFUL ANALYSES ({len(successful)}):")
        for file in successful:
            logger.info(f"  - {file}")
    
    if failed:
        logger.info(f"\n❌ FAILED ANALYSES ({len(failed)}):")
        for item in failed:
            logger.info(f"  - {item['file']}: {item['error']}")
    
    return successful, failed

if __name__ == "__main__":
    successful, failed = main()
    
    # Exit with appropriate code
    sys.exit(0 if len(failed) == 0 else 1)
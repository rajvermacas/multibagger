# Multi-Stock Analysis Workflow

**Think harder** about this comprehensive stock analysis task to ensure accuracy and thoroughness.

## Objective
Analyze all stock-related files in the directory `$ARGUMENTS` using specialized sub-agents to generate comprehensive markdown reports, then create a consolidated ranking report.

## Workflow Steps

### Phase 1: File Discovery and Validation
1. **Scan the target directory** `$ARGUMENTS` for all relevant stock data files
2. **Validate file formats** and ensure they contain analyzable stock data
3. **Create a list** of valid files for processing
4. **Log any issues** with file accessibility or format problems

### Phase 2: Individual Stock Analysis
For each valid file identified:

1. **Delegate to the stock-analyst sub-agent**

2. **Monitor sub-agent progress** and log completion status
3. **Verify report generation** was successful for each file

### Phase 3: Quality Assurance
1. **Verify all reports** were generated successfully
2. **Check for missing data** or incomplete analyses
3. **Re-process any failed analyses** with detailed error logging

### Phase 4: Comparative Analysis and Ranking
1. **Extract key metrics** from all individual reports

2. **Create a consolidated ranking** sorting stocks in descending order from investment perspective to buy and hold for atleast 5 years

3. **Generate the master report** as `resources/reports/$(date +%Y-%m-%d)/CONSOLIDATED_STOCK_RANKING.md`
   - Date format: YYYY-MM-DD (e.g., `2025-07-31`)

## Error Handling
- **Log all errors** to `resources/reports/$(date +%Y-%m-%d)/analysis_log.txt`
  - Date format: YYYY-MM-DD (e.g., `2025-07-31`)
- **Continue processing** remaining files if individual analyses fail
- **Generate partial reports** with clear annotations about missing data

## Final Deliverables
All files saved in `resources/reports/$(date +%Y-%m-%d)/` directory (YYYY-MM-DD format):
1. Individual stock analysis reports (one per stock)
2. Consolidated ranking report
3. Execution log with timestamps and status updates
4. Summary statistics on the analysis batch

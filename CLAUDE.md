# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an Excel VBA project for downloading option pricing data from LSEG (London Stock Exchange Group) data feeds. The system automates the collection of option prices for various underlyings, strikes, and maturities, computes Greeks, and prepares data for database upload.

## Key Commands

### Excel VBA Development
- **Generate RICs**: Execute `GenerateAllRICs()` in RICconfiguration module to build option ticker codes
- **Main Download Process**: Execute `MainDownloadProcess()` in OptionDownload module to start data collection
- **Initialize Workbook**: Execute `InitializeWorkbook()` in OptionDownload module to create necessary sheets if missing
- **Refresh LSEG Data**: Use `Application.Run "WorkspaceRefreshWorksheet", True, 120000, sheetName` for LSEG data refresh

### Data Quality Checks
- Run quality reports through the QualityReport worksheet after downloads
- Check for missing prices, especially near ATM strikes
- Validate implied volatility bounds

## Architecture

### Core Modules
- **RICconfiguration.bas**: Generates LSEG RIC codes for options based on maturity and strike ranges
- **OptionDownload.bas**: Main download orchestration, batch processing, and workbook initialization
- **VanillaOptions.bas**: Black-Scholes Greeks calculations (GBlackScholesNGreeks, Black76, etc.)
- **Distributions.bas**: Additional option pricing calculations
- **IVOL.bas**: Implied volatility calculations

### Data Flow
1. Configuration parameters set in Config sheet (underlying, strike ranges, dates)
2. RIC codes generated for all strike/maturity combinations
3. LSEG formulas refresh to pull option prices (uses RHistory function)
4. Greeks calculated using Black-Scholes functions
5. Data staged and quality-checked
6. Final data prepared for CSV export

### Key Worksheets
- **Config**: Master configuration (underlying ticker, strike ranges, batch sizes)
- **RIC_List**: Generated option RIC codes with processing status tracking
- **DataCollection**: Active data collection workspace
- **Staging**: Temporary data storage during processing
- **QualityReport**: Data validation and missing price analysis
- **Future et co**: Underlying price and interest rate data

### LSEG Integration
- Uses RHistory formula: `=RHistory(RIC,".Timestamp;.Close","START:date END:date INTERVAL:1D")`
- Month codes for options: Calls (A-L), Puts (M-X)
- RIC format: RootRIC + Strike + MonthCode + Year (e.g., "1EW7000T25")

## Important Notes

- Process is intentionally semi-manual to maintain control and monitoring capability
- LSEG data refresh is slow - processes option by option
- Batch processing implemented to handle large strike/maturity combinations
- Greeks calculation uses custom VBA functions, not external libraries
- Data quality checks essential due to potential missing prices for illiquid strikes
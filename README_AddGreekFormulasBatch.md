# Add Greek Formulas Batch Process

## Overview

`RunAddGreekFormulasBatch` is a VBA macro that calculates Black-Scholes Greeks for large datasets (10k+ rows) by processing in batches to prevent Excel from becoming unresponsive.

## How to Run

1. Open the worksheet containing your option data
2. Make sure the worksheet is the **active sheet**
3. Run the macro:
   - Press `Alt + F8` to open the Macros dialog
   - Select `RunAddGreekFormulasBatch`
   - Click **Run**
4. Confirm when prompted
5. Watch progress in the status bar

## Required Column Structure

The worksheet must have the following columns (same as Staging sheet):

| Column | Name | Description |
|--------|------|-------------|
| A | Spot_Date | Date of the price observation |
| B | Premium | Option price (required - used to detect data rows) |
| D | Maturity | Option expiration date |
| E | Interest_Rate | Risk-free interest rate |
| F | Spot | Underlying spot price |
| G | Strike | Option strike price |
| H | Type | "C" for Call, "P" for Put |

## Output Columns (Calculated)

### Primary Greeks (Columns I-N)

| Column | Greek | Description |
|--------|-------|-------------|
| I | Implied Volatility | Calculated via bisection method |
| J | Delta | Price sensitivity to spot |
| K | Vega | Price sensitivity to volatility |
| L | Gamma | Delta sensitivity to spot |
| M | Theta | Time decay |
| N | Rho | Price sensitivity to interest rate |

### Second-Order Greeks (Columns U-AA)

| Column | Greek | Description |
|--------|-------|-------------|
| U | DDeltaDVol | Vanna - Delta sensitivity to vol |
| V | DDeltaDVolDVol | Second-order vanna |
| W | DDeltaDTime | Charm - Delta decay |
| X | DGammaDSpot | Speed - Gamma sensitivity to spot |
| Y | DGammaDVol | Zomma - Gamma sensitivity to vol |
| Z | DVegaDVol | Vomma - Vega sensitivity to vol |
| AA | DVegaDVolDVol | Ultima - Third-order vega |

## Processing Details

- **Batch Size**: 500 rows per batch (default)
- **Process per batch**:
  1. Add Greek formulas to the batch
  2. Calculate the worksheet
  3. Convert formulas to values (paste special values)
  4. Move to next batch
- **Progress**: Displayed in Excel status bar
- **Completion**: Message box shows total rows processed

## Error Handling

- Rows with empty or error values in Premium (column B) are skipped
- Greek calculations that fail return "NA"
- If an error occurs, calculation mode and screen updating are restored

## Calling from VBA

```vba
' Run on active sheet with default batch size (500)
RunAddGreekFormulasBatch

' Run on specific sheet with custom batch size
AddGreekFormulasBatch ThisWorkbook.Sheets("MyData"), 1000
```

## Performance

| Rows | Approximate Time |
|------|------------------|
| 1,000 | ~10 seconds |
| 10,000 | ~2 minutes |
| 50,000 | ~10 minutes |

*Times vary based on system performance and Excel version.*

# Hardcoded Constants and Parameters

This document lists all constants and parameters that are hardcoded in VBA modules (not configurable via Excel named ranges).

---

## RIC Generation Constants (RICconfiguration.bas)

| Constant | Value | Description |
|----------|-------|-------------|
| `UNDERLYING_MONTH_MODE` | `"Quarter End"` | How to determine underlying future month. Options: `"Same Month"` or `"Quarter End"` |
| `OPTION_STRIKE_DECIMALS` | `1` | Number of decimal places for strike in RIC. Options: `1` or `2` |
| `OPTION_YEAR_DIGITS` | `1` | Number of year digits in option RIC. Options: `1` or `2` |
| `g_ChainStepSize` | `7` | Column spacing between chain downloads in ChainDownload sheet |

---

## Sheet Name Constants

| Constant | Value | Module(s) |
|----------|-------|-----------|
| `SHEET_CONFIG` | `"Config"` | RICconfiguration, OptionDownload, SetupConfig |
| `SHEET_RIC_LIST` | `"RIC_List"` | RICconfiguration, OptionDownload |
| `SHEET_COLLECTION` | `"DataCollection"` | OptionDownload, SetupConfig |
| `SHEET_STAGING` | `"Staging"` | OptionDownload, SetupConfig |
| `SHEET_QUALITY` | `"QualityReport"` | OptionDownload, SetupConfig |
| `SHEET_FUTURE` | `"Future et co"` | OptionDownload, SetupConfig |
| `SHEET_PROGRESS` | `"Progress"` | SetupConfig |
| `SHEET_MAIN` | `"Main"` | SetupConfig |

---

## Named Range References (Excel)

These are names of Excel named ranges referenced in code:

| Constant/Reference | Named Range | Description |
|--------------------|-------------|-------------|
| `MONTH_CALL` | `"monthCall"` | Call month codes (A-L) |
| `MONTH_PUT` | `"monthPut"` | Put month codes (M-X) |
| `WEEKLY_CALL` | `"weeklyCall"` | Weekly call type code |
| `WEEKLY_PUT` | `"weeklyPut"` | Weekly put type code |
| `OPTION_FREQUENCY` | `"optionFrequency"` | Monthly or weekly options |
| `OPTION_MONTH_CODE_METHOD` | `"optionMonthCodeMethod"` | Same Month or Next Month |
| `RANGE_DOWNLOAD` | `"UnderlyingDownload"` | First column for underlying download |
| `RANGE_UNDERLYING_START_DATE` | `"UnderlyingStartDate"` | Start date for underlying data |
| `RANGE_UNDERLYING_END_DATE` | `"UnderlyingEndDate"` | End date for underlying data |
| `RANGE_RFR` | `"RFR"` | Risk-free rate |
| - | `"rootRIC"` | Root RIC for options |
| - | `"rootBB"` | Bloomberg root ticker |
| - | `"rootUnderlyingBB"` | Bloomberg root for underlying |
| - | `"rootUnderlyingRIC"` | LSEG RIC root for underlying |
| - | `"methodRICBB"` | Method for RIC to Bloomberg conversion |
| - | `"monthFutureBloomberg"` | Future month codes for Bloomberg |
| - | `"maturityDate"` | Maturity dates range |
| - | `"minStrikePut"` / `"maxStrikePut"` | Put strike range |
| - | `"minStrikeCall"` / `"maxStrikeCall"` | Call strike range |
| - | `"steps"` | Strike step size |

---

## Timing & Timeout Constants

| Location | Value | Description |
|----------|-------|-------------|
| `OptionDownload:101` | `60` seconds | Future refresh timeout |
| `OptionDownload:686` | `60` checks | Batch refresh timeout (60 checks x 3s = 3 min) |
| `RICconfiguration:1203` | `60` checks | Chain download timeout (60 checks x 3s = 3 min) |
| `RICconfiguration:1167` | `120000` ms | LSEG WorkspaceRefreshWorksheet timeout |
| Various | `TimeValue("00:00:02")` to `("00:00:05")` | OnTime polling intervals |

---

## Data Layout Constants

| Constant | Value | Location | Description |
|----------|-------|----------|-------------|
| `ROW_SPACING` | `300` | OptionDownload (multiple) | Rows between RHistory formulas in DataCollection |
| `MAX_UNDERLYING_ROWS` | `10000` | OptionDownload:66 | Maximum rows for underlying price data |
| Start column for chains | `7` (column G) | RICconfiguration:1285 | First column for option chain data |

---

## State Machine Constants (RICconfiguration.bas)

| Constant | Value | Description |
|----------|-------|-------------|
| `CHAIN_STATE_IDLE` | `0` | No chain download active |
| `CHAIN_STATE_DOWNLOADING_CHAINS` | `1` | Downloading chain-of-chains |
| `CHAIN_STATE_PROCESSING_CHAINS` | `2` | Processing chain list |
| `CHAIN_STATE_DOWNLOADING_OPTIONS` | `3` | Downloading option data |
| `CHAIN_STATE_PROCESSING_OPTIONS` | `4` | Processing option data |

---

## UI/Formatting Constants

### Conditional Formatting Colors (RGB)
| Color | RGB Value | Usage |
|-------|-----------|-------|
| Light Gray | `RGB(200, 200, 200)` | Header backgrounds |
| Light Green | `RGB(200, 255, 200)` | Success/Yes status |
| Light Red | `RGB(255, 200, 200)` | Error status |
| Light Yellow | `RGB(255, 255, 200)` | Processing/In-progress status |
| Light Blue | `RGB(200, 200, 255)` | Staging sheet headers |

### Number Formats
| Format | Usage |
|--------|-------|
| `"#,##0"` | Strike prices, integers |
| `"#,##0.00"` | Premiums, prices with 2 decimals |
| `"0.00%"` | Interest rates, percentages |
| `"0.0000"` | Greeks, high precision values |
| `"mm/dd/yyyy"` | Maturity dates |

---

## Mathematical Constants (Distributions.bas)

Used in normal distribution calculations (should not be modified):

| Constants | Values | Purpose |
|-----------|--------|---------|
| `a1` to `a5` | Various | CND approximation coefficients |
| `A()`, `B()`, `c()` | Arrays | NormSInv coefficients |
| `XX()`, `W()` | Arrays | Gauss-Legendre quadrature weights |

---

## Notes

1. **To make a constant configurable**: Add a named range in the Config sheet and update the VBA code to read from it instead of using the hardcoded value.

2. **Critical constants for RIC format**:
   - `OPTION_YEAR_DIGITS` - Must match LSEG's expected format for the specific option chain
   - `OPTION_STRIKE_DECIMALS` - Must match LSEG's expected strike format
   - `UNDERLYING_MONTH_MODE` - Determines how underlying future is selected

3. **Year 2000 assumption**: The code assumes years are 2000+ when converting 2-digit year codes (see line 141 in RICconfiguration.bas: `ulYear = 2000 + CInt(ricDict("UnderlyingYearCode"))`).

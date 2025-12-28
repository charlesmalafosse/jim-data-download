
Create a plan for these 2 new parameters
Option Strike Param: Adjust to have strike on 4 digits if finishing by 0. for x = 111, it is 1110 and not 11100
Option Year character: Use 1 character year and not 2. Have as an option
It is
OGBM1200L5^L25
And not 
OGBM12000L25^L25

● Done. Here's a summary of the changes:

  Added constants (lines 21-22):
  Public Const OPTION_STRIKE_DECIMALS As Integer = 1  ' 1 or 2 decimal places for strike
  Public Const OPTION_YEAR_DIGITS As Integer = 1      ' 1 or 2 year digits

  Modified FormatStrikeForRIC (lines 330-350):
  - Now checks OPTION_STRIKE_DECIMALS constant
  - With value 1: strike 120 → "1200"
  - With value 2: strike 120 → "12000" (original behavior)

  Modified year code calculations (lines 275, 277):
  - Now uses OPTION_YEAR_DIGITS instead of hardcoded 2
  - With value 1: year 2025 → "5"
  - With value 2: year 2025 → "25" (original behavior)

  Result with current settings (both = 1):
  - Strike 120, Year 2025 → OGBM1200L5 (correct format)

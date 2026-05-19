' ============================================
' RIC GENERATOR MODULE
' Generates complete list of option RICs based on
' maturity dates and strike ranges in Config sheet
' ============================================

Option Explicit

' ============================================
' GLOBAL SHEET NAME CONSTANTS
' ============================================
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_RIC_LIST As String = "RIC_List"
Public Const MONTH_CALL = "monthCall"
Public Const MONTH_PUT = "monthPut"
Public Const WEEKLY_CALL = "weeklyCall"
Public Const WEEKLY_PUT = "weeklyPut"
Public Const OPTION_FREQUENCY = "optionFrequency"
Public Const OPTION_MONTH_CODE_METHOD = "optionMonthCodeMethod"
Public Const UNDERLYING_MONTH_MODE_RANGE = "underlyingMonthMode"  ' Config named range
Public Const UNDERLYING_MONTH_MODE_DEFAULT = "Same Month"          ' Fallback when range not set
                                                                   ' Valid values: "Same Month", "Quarter End", "Next 2 Month"
' DEPRECATED - use strikeMultiplier named range in Config sheet instead
' Public Const OPTION_STRIKE_DECIMALS As Integer = 1  ' 1 or 2 decimal places for strike
Public Const OPTION_YEAR_DIGITS As Integer = 2      ' 1 or 2 year digits

' OnTime Chain State for DownloadFromChain
Public g_ChainState As Long
Public g_ChainSheet As Worksheet
Public g_ChainRootRIC As String
Public g_ChainIndex As Long
Public g_ChainTotalChains As Long
Public g_ChainStepSize As Long
Public g_ChainRefreshCheckCount As Long
Public g_ChainStopRequested As Boolean
Public g_RICListSheet As Worksheet
Public g_ChainBatchStart As Long
Public g_ChainCleanRICs() As String
Public g_ChainColumns() As Long

Public Const CHAIN_BATCH_SIZE As Long = 3
Public Const CHAIN_STATE_IDLE As Long = 0
Public Const CHAIN_STATE_DOWNLOADING_CHAINS As Long = 1
Public Const CHAIN_STATE_PROCESSING_CHAINS As Long = 2
Public Const CHAIN_STATE_DOWNLOADING_OPTIONS As Long = 3
Public Const CHAIN_STATE_PROCESSING_OPTIONS As Long = 4

' ============================================
' MAIN RIC GENERATION FUNCTION
' ============================================
Sub GenerateAllRICs()
    ' Generate all RICs using the flat-step strike range (Config!steps).
    If Not ValidateBloombergConfig() Then Exit Sub
    WriteRICListToSheet BuildCompleteRICList()
End Sub

Public Sub GenerateAllRICsMoneyness()
    ' Generate all RICs using moneyness-band variable-step strikes.
    ' Strikes come from Config!moneynessBands anchored to Config!spotMin /
    ' spotMax instead of the flat Config!steps range.
    If Not ValidateBloombergConfig() Then Exit Sub

    Dim strikes As Collection
    Set strikes = GetMoneynessStrikeRange()
    If strikes Is Nothing Then Exit Sub  ' invalid config - message already shown

    If strikes.count = 0 Then
        MsgBox "Moneyness generator produced no strikes - check spot/band config.", _
               vbExclamation, "RIC Generation"
        Exit Sub
    End If

    ' Puts and calls share the same strike grid (call/put of same strike adjacent)
    WriteRICListToSheet BuildCompleteRICList(strikes, strikes)
End Sub

Private Function ValidateBloombergConfig() As Boolean
    ' Validate the Bloomberg-conversion named ranges in Config. Shows a
    ' message and returns False on the first problem; True if all present.
    Dim methodRICBB As String
    Dim rootUnderlyingBB As String
    Dim rootUnderlyingRIC As String

    On Error Resume Next
    methodRICBB = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("methodRICBB").Value)
    rootUnderlyingBB = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootUnderlyingBB").Value)
    rootUnderlyingRIC = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootUnderlyingRIC").Value)
    On Error GoTo 0

    If UCase(methodRICBB) <> "FUTURE" Then
        MsgBox "Invalid or missing 'methodRICBB' in Config sheet!" & vbNewLine & _
               "Current value: '" & methodRICBB & "'" & vbNewLine & _
               "Supported methods: 'Future'", _
               vbCritical, "Configuration Error"
        ValidateBloombergConfig = False
        Exit Function
    End If

    If rootUnderlyingBB = "" Then
        MsgBox "Missing 'rootUnderlyingBB' named range in Config sheet!" & vbNewLine & _
               "Please set the Bloomberg root ticker for the underlying.", _
               vbCritical, "Configuration Error"
        ValidateBloombergConfig = False
        Exit Function
    End If

    If rootUnderlyingRIC = "" Then
        MsgBox "Missing 'rootUnderlyingRIC' named range in Config sheet!" & vbNewLine & _
               "Please set the LSEG RIC root for the underlying (e.g., 'FGBM').", _
               vbCritical, "Configuration Error"
        ValidateBloombergConfig = False
        Exit Function
    End If

    ValidateBloombergConfig = True
End Function

Private Sub WriteRICListToSheet(ricList As Collection)
    ' Write a built RIC list to the RIC_List sheet: headers, the cols A-I
    ' output loop (underlying RIC, Bloomberg ticker, expiry suffix), and
    ' formatting. Shared by GenerateAllRICs and GenerateAllRICsMoneyness.
    Dim outputSheet As Worksheet
    Dim ricDict As Object
    Dim i As Long
    Dim lastRow As Long
    Dim rootUnderlyingBB As String
    Dim rootUnderlyingRIC As String
    Dim underlyingRIC As String

    ' Bloomberg-conversion roots (already validated non-empty by caller)
    On Error Resume Next
    rootUnderlyingBB = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootUnderlyingBB").Value)
    rootUnderlyingRIC = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootUnderlyingRIC").Value)
    On Error GoTo 0

    ' Check if sheet exists, create if it doesn't
    On Error Resume Next
    Set outputSheet = ThisWorkbook.Sheets(SHEET_RIC_LIST)
    On Error GoTo 0

    If outputSheet Is Nothing Then
        Set outputSheet = ThisWorkbook.Sheets.Add
        outputSheet.Name = SHEET_RIC_LIST
    Else
        ' Clear existing content in columns A to O
        lastRow = outputSheet.Cells(outputSheet.Rows.count, "A").End(xlUp).Row
        If lastRow > 0 Then
            outputSheet.Range("A1:O" & lastRow).Clear
        End If
    End If

    ' Add headers
    With outputSheet
        .Range("A1").Value = "RIC"
        .Range("B1").Value = "Maturity"
        .Range("C1").Value = "Strike"
        .Range("D1").Value = "Type"
        .Range("E1").Value = "Month Code"
        .Range("F1").Value = "Year"
        .Range("G1").Value = "Underlying LSEG"
        .Range("H1").Value = "Bloom_Ticker"  ' Bloomberg ticker
        .Range("I1").Value = "Processed"     ' Processing status (shifted from H)
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = RGB(200, 200, 200)
    End With

    ' Output all RICs
    i = 2
    Dim ric As Variant
    For Each ric In ricList
        Set ricDict = ric  ' Each item in collection is a dictionary
        With outputSheet
            .Cells(i, 1).Value = ricDict("FullRIC")
            .Cells(i, 2).Value = ricDict("Maturity")
            .Cells(i, 3).Value = ricDict("Strike")
            .Cells(i, 4).Value = ricDict("OptionType")
            .Cells(i, 5).Value = ricDict("MonthCode")
            .Cells(i, 6).Value = ricDict("YearCode")
            ' Column G: Build underlying RIC from root + future month code + underlying year code
            underlyingRIC = rootUnderlyingRIC & ricDict("FutureMonthCode") & ricDict("UnderlyingYearCode")
            ' Add expiration suffix if underlying has expired (^decade format: ^2 for 2020s, ^3 for 2030s)
            Dim ulYear As Integer
            Dim underlyingExpiryDate As Date
            ulYear = 2000 + CInt(ricDict("UnderlyingYearCode"))
            underlyingExpiryDate = DateSerial(ulYear, ricDict("UnderlyingMonth") + 1, 0)  ' Last day of underlying month
            If underlyingExpiryDate < Date Then
                underlyingRIC = underlyingRIC & "^" & CStr((ulYear - 2000) \ 10)
            End If
            .Cells(i, 7).Value = underlyingRIC
            ' Column H: Bloomberg ticker from underlying RIC
            .Cells(i, 8).Value = RICToBloomberg(underlyingRIC, rootUnderlyingBB)
            ' Column I: Initialize Processed column as "No"
            .Cells(i, 9).Value = "No"
        End With
        i = i + 1
    Next

    ' Format
    outputSheet.Columns("A:I").AutoFit
    outputSheet.Range("B:B").NumberFormat = "mm/dd/yyyy"
    outputSheet.Range("C:C").NumberFormat = "#,##0"

    ' Add conditional formatting to Processed column (I) for visual feedback
    With outputSheet.Range("I2:I" & ricList.count + 1).FormatConditions
        .Delete
        ' Green for "Yes"
        .Add Type:=xlTextString, String:="Yes", TextOperator:=xlContains
        .Item(.count).Interior.Color = RGB(200, 255, 200)
        ' Red for "Error"
        .Add Type:=xlTextString, String:="Error", TextOperator:=xlContains
        .Item(.count).Interior.Color = RGB(255, 200, 200)
        ' Yellow for "Processing"
        .Add Type:=xlTextString, String:="Processing", TextOperator:=xlContains
        .Item(.count).Interior.Color = RGB(255, 255, 200)
    End With

    MsgBox "Generated " & ricList.count & " RICs!" & vbNewLine & _
           "Check '" & SHEET_RIC_LIST & "' sheet for details." & vbNewLine & _
           "Column I tracks processing status.", vbInformation
End Sub

' ============================================
' MONEYNESS-BAND STRIKE GENERATOR
' ============================================

Function GetMoneynessStrikeRange() As Collection
    ' Build a strike list with variable step per moneyness band.
    ' Reads Config named ranges:
    '   spotMin, spotMax  - expected spot low / high over the period
    '   moneynessBands    - 3-col range: lower moneyness, upper moneyness, step
    ' Lower-bound strike = spotMin*(1+lowerMoneyness);
    ' upper-bound strike = spotMax*(1+upperMoneyness).
    ' Returns Nothing (after a MsgBox) on invalid config.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)

    Dim spotMin As Double, spotMax As Double
    Dim bandsRng As Range
    On Error Resume Next
    spotMin = ws.Range("spotMin").Value
    spotMax = ws.Range("spotMax").Value
    Set bandsRng = ws.Range("moneynessBands")
    On Error GoTo 0

    If spotMin <= 0 Or spotMax <= 0 Or spotMax <= spotMin Then
        MsgBox "Invalid 'spotMin' / 'spotMax' in Config sheet!" & vbNewLine & _
               "Both must be positive and spotMax must be greater than spotMin.", _
               vbCritical, "Configuration Error"
        Set GetMoneynessStrikeRange = Nothing
        Exit Function
    End If

    If bandsRng Is Nothing Then
        MsgBox "Missing 'moneynessBands' named range in Config sheet!" & vbNewLine & _
               "Expected a 3-column range: lower moneyness, upper moneyness, step.", _
               vbCritical, "Configuration Error"
        Set GetMoneynessStrikeRange = Nothing
        Exit Function
    End If

    ' Collect strikes into a dictionary keyed by rounded value (dedup overlaps)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim lowerM As Variant, upperM As Variant, stepV As Variant
    Dim rawLo As Double, rawHi As Double, stepSize As Double
    Dim startK As Double, k As Double
    Dim bandCount As Long
    bandCount = 0

    For r = 1 To bandsRng.Rows.count
        lowerM = bandsRng.Cells(r, 1).Value
        upperM = bandsRng.Cells(r, 2).Value
        stepV = bandsRng.Cells(r, 3).Value

        ' Stop at the first blank row
        If IsEmpty(lowerM) And IsEmpty(upperM) And IsEmpty(stepV) Then Exit For

        ' Skip rows that aren't fully specified / valid
        If Not (IsNumeric(lowerM) And IsNumeric(upperM) And IsNumeric(stepV)) Then GoTo NextBand
        If CDbl(stepV) <= 0 Then GoTo NextBand
        If CDbl(upperM) <= CDbl(lowerM) Then GoTo NextBand

        stepSize = CDbl(stepV)
        rawLo = spotMin * (1 + CDbl(lowerM))
        rawHi = spotMax * (1 + CDbl(upperM))

        ' First grid point >= rawLo, snapped to the step grid
        startK = Int(rawLo / stepSize) * stepSize
        If startK < rawLo Then startK = startK + stepSize

        Dim n As Long
        n = 0
        Do
            k = Round(startK + n * stepSize, 6)
            If k > rawHi Then Exit Do
            If k > 0 Then dict(CStr(k)) = k
            n = n + 1
        Loop

        bandCount = bandCount + 1
NextBand:
    Next r

    If bandCount = 0 Then
        MsgBox "No valid rows found in 'moneynessBands'." & vbNewLine & _
               "Each row needs: lower moneyness, upper moneyness, step (>0).", _
               vbCritical, "Configuration Error"
        Set GetMoneynessStrikeRange = Nothing
        Exit Function
    End If

    ' Return as a Collection (BuildStrikeUnion sorts/dedups downstream anyway)
    Dim result As New Collection
    Dim key As Variant
    For Each key In dict.Keys
        result.Add dict(key)
    Next key

    Set GetMoneynessStrikeRange = result
End Function

' ============================================
' BUILD COMPLETE RIC LIST
' ============================================

Function BuildCompleteRICList(Optional ByVal putStrikesIn As Collection, _
                              Optional ByVal callStrikesIn As Collection) As Collection
    ' Generate RICs in (maturity asc, strike asc, PUT-then-CALL) order so a
    ' top-to-bottom scroll of RIC_List shows the earliest maturity first,
    ' the lowest strike within each maturity, and PUT/CALL of the same
    ' (maturity, strike) adjacent.
    ' putStrikesIn / callStrikesIn: optional pre-built strike collections
    ' (used by the moneyness generator). When omitted, the flat-step
    ' Config strike ranges are used via GetStrikeRange.
    Dim ricList As New Collection
    Dim maturities As Collection
    Dim putStrikes As Collection
    Dim callStrikes As Collection
    Dim ricInfo As Object  ' Dictionary

    Set maturities = GetMaturityDates()

    If putStrikesIn Is Nothing Then
        Set putStrikes = GetStrikeRange("PUT")
    Else
        Set putStrikes = putStrikesIn
    End If

    If callStrikesIn Is Nothing Then
        Set callStrikes = GetStrikeRange("CALL")
    Else
        Set callStrikes = callStrikesIn
    End If

    ' Sorted unique union of all strikes (for the outer loop)
    Dim allStrikes() As Double
    Dim strikeCount As Long
    strikeCount = BuildStrikeUnion(putStrikes, callStrikes, allStrikes)

    ' Sorted maturities (oldest first) for the inner loop
    Dim matDates() As Date
    Dim matCount As Long
    matCount = MaturitiesAsSortedArray(maturities, matDates)

    If strikeCount = 0 Or matCount = 0 Then
        Set BuildCompleteRICList = ricList
        Exit Function
    End If

    ' Membership dictionaries — O(1) check whether a strike exists on each side
    Dim putDict As Object, callDict As Object
    Set putDict = CreateObject("Scripting.Dictionary")
    Set callDict = CreateObject("Scripting.Dictionary")

    Dim s As Variant
    For Each s In putStrikes
        putDict(CStr(CDbl(s))) = True
    Next s
    For Each s In callStrikes
        callDict(CStr(CDbl(s))) = True
    Next s

    ' Iterate maturity asc -> strike asc -> PUT then CALL
    Dim i As Long, j As Long
    Dim strikeKey As String
    For j = 0 To matCount - 1
        For i = 0 To strikeCount - 1
            strikeKey = CStr(allStrikes(i))
            If putDict.Exists(strikeKey) Then
                Set ricInfo = CreateRICInfo(matDates(j), allStrikes(i), "PUT")
                ricList.Add ricInfo
            End If
            If callDict.Exists(strikeKey) Then
                Set ricInfo = CreateRICInfo(matDates(j), allStrikes(i), "CALL")
                ricList.Add ricInfo
            End If
        Next i
    Next j

    Set BuildCompleteRICList = ricList
End Function

Private Function BuildStrikeUnion(putStrikes As Collection, callStrikes As Collection, _
                                  ByRef result() As Double) As Long
    ' Union of put + call strikes, deduped, ascending. Returns count.
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim s As Variant
    For Each s In putStrikes
        dict(CDbl(s)) = True
    Next s
    For Each s In callStrikes
        dict(CDbl(s)) = True
    Next s

    Dim count As Long
    count = dict.count
    If count = 0 Then
        ReDim result(0 To 0)
        BuildStrikeUnion = 0
        Exit Function
    End If

    ReDim result(0 To count - 1)
    Dim k As Variant, idx As Long
    idx = 0
    For Each k In dict.Keys
        result(idx) = CDbl(k)
        idx = idx + 1
    Next k

    ' Insertion sort (small N)
    Dim i As Long, j As Long, tmp As Double
    For i = 1 To count - 1
        tmp = result(i)
        j = i - 1
        Do While j >= 0
            If result(j) <= tmp Then Exit Do
            result(j + 1) = result(j)
            j = j - 1
        Loop
        result(j + 1) = tmp
    Next i

    BuildStrikeUnion = count
End Function

Private Function MaturitiesAsSortedArray(maturities As Collection, _
                                         ByRef result() As Date) As Long
    ' Copy maturities to a 0-based array sorted ascending. Returns count.
    Dim count As Long
    count = maturities.count
    If count = 0 Then
        ReDim result(0 To 0)
        MaturitiesAsSortedArray = 0
        Exit Function
    End If

    ReDim result(0 To count - 1)
    Dim i As Long
    For i = 1 To count
        result(i - 1) = CDate(maturities(i))
    Next i

    ' Insertion sort
    Dim j As Long, tmpDate As Date
    For i = 1 To count - 1
        tmpDate = result(i)
        j = i - 1
        Do While j >= 0
            If result(j) <= tmpDate Then Exit Do
            result(j + 1) = result(j)
            j = j - 1
        Loop
        result(j + 1) = tmpDate
    Next i

    MaturitiesAsSortedArray = count
End Function

' ============================================
' CREATE INDIVIDUAL RIC (Returns Dictionary)
' ============================================

Function CreateRICInfo(maturityDate As Date, strike As Double, optionType As String) As Object
    Dim rootRIC As String
    Dim monthCode As String
    Dim monthCodeCallForExpiredRIC As String
    Dim FutureMonthCode As String
    Dim yearCode As String
    Dim strikeStr As String
    Dim ricDict As Object
    Dim ricMonth As Integer
    Dim optFrequency As String
    Dim monthCodeMethod As String

    ' Create dictionary to hold RIC information
    Set ricDict = CreateObject("Scripting.Dictionary")

    ' Get values
    rootRIC = ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootRIC").Value
    optFrequency = GetOptionFrequency()
    monthCodeMethod = GetOptionMonthCodeMethod()

    ' Track option year (may differ from maturity year if month rolls over Dec->Jan)
    Dim optionYear As Integer
    optionYear = Year(maturityDate)  ' Default to maturity year

    ' Get type/month code based on frequency and month code method
    If optFrequency = "weekly" Then
        ' Weekly: apply optionMonthCodeMethod same as monthly
        If monthCodeMethod = "Same Month" Then
            ricMonth = Month(maturityDate)
        Else
            ' Next Month: use following month for month code lookup
            ricMonth = Month(maturityDate) + 1
            If ricMonth > 12 Then
                ricMonth = 1
                optionYear = Year(maturityDate) + 1  ' Year rolls over with month
            End If
        End If
        monthCode = GetMonthCodeFromTable(ricMonth, optionType)
        monthCodeCallForExpiredRIC = GetMonthCodeFromTable(ricMonth, "CALL")
    Else
        ' Monthly: use optionMonthCodeMethod to determine month
        If monthCodeMethod = "Same Month" Then
            ' Same Month: use maturity month
            ricMonth = Month(maturityDate)
        ElseIf monthCodeMethod = "Next 2 Month" Then
            ' Next Month: use following month for month code lookup
            ricMonth = Month(maturityDate) + 2
            If ricMonth > 12 Then
                ricMonth = ricMonth - 12
                optionYear = Year(maturityDate) + 1  ' Year rolls over with month
            End If
        Else
            ' Next Month: use following month for month code lookup
            ricMonth = Month(maturityDate) + 1
            If ricMonth > 12 Then
                ricMonth = 1
                optionYear = Year(maturityDate) + 1  ' Year rolls over with month
            End If
        End If
        monthCode = GetMonthCodeFromTable(ricMonth, optionType)
        monthCodeCallForExpiredRIC = GetMonthCodeFromTable(ricMonth, "CALL")
    End If

    ' Get future month code (for underlying RIC/Bloomberg ticker)
    Dim underlyingMonth As Integer
    Dim underlyingYear As Integer
    Dim underlyingMode As String
    underlyingYear = optionYear ' Year(maturityDate)
    underlyingMode = GetUnderlyingMonthMode()

    If underlyingMode = "Quarter End" Then
        underlyingMonth = GetQuarterEndMonth(Month(maturityDate))
        ' If maturity is in Dec (Q1) and quarter end is March, it's next year
        If Month(maturityDate) = 12 And underlyingMonth = 3 Then
            underlyingYear = underlyingYear + 1
        End If
    Else
        underlyingMonth = ricMonth  ' Same Month: use option's month
        underlyingYear = optionYear
    End If
    FutureMonthCode = GetFutureMonthCode(underlyingMonth)

    yearCode = Right(CStr(optionYear), OPTION_YEAR_DIGITS)
    Dim underlyingYearCode As String
    underlyingYearCode = Right(CStr(underlyingYear), 2)  ' Futures always use 2-digit years

    ' Use appropriate strike formatter based on frequency
    If optFrequency = "weekly" Then
        strikeStr = FormatStrikeForWeeklyRIC(strike)
    Else
        strikeStr = FormatStrikeForRIC(strike)
    End If

    ' Populate dictionary
    ricDict.Add "FullRIC", BuildRICString(rootRIC, strikeStr, monthCode, yearCode, maturityDate, monthCodeCallForExpiredRIC, optFrequency, optionYear)
    ricDict.Add "Maturity", maturityDate
    ricDict.Add "Strike", strike
    ricDict.Add "OptionType", optionType
    ricDict.Add "MonthCode", monthCode
    ricDict.Add "FutureMonthCode", FutureMonthCode
    ricDict.Add "UnderlyingMonth", underlyingMonth  ' Month number for underlying (used for expiry check)
    ricDict.Add "UnderlyingYearCode", underlyingYearCode  ' Year code for underlying (may differ from option year if Dec->Mar)
    ricDict.Add "YearCode", yearCode

    Set CreateRICInfo = ricDict
End Function

' ============================================
' BUILD RIC STRING
' ============================================

Function BuildRICString(rootRIC As String, strikeStr As String, monthCode As String, yearCode As String, maturityDate As Date, monthCodeCallForExpiredRIC As String, optFrequency As String, optionYear As Integer) As String
    ' Builds the complete RIC string
    ' Monthly format: 1EW7000T25 (rootRIC, strike, monthCode for following month, year)
    ' Weekly format:  1E3W1005L25 (rootRIC minus W, occurrence, W, strike, monthCode for maturity month, year)
    Dim occurrence As String
    Dim rootWithoutW As String

    If optFrequency = "weekly" Then
        occurrence = CStr(GetDayOccurrenceInMonth(maturityDate))
        ' Weekly format: {root minus W}{occurrence}W{strike}{monthCode}{yearCode}
        rootWithoutW = Left(rootRIC, Len(rootRIC) - 1)
        BuildRICString = rootWithoutW & occurrence & "W" & strikeStr & monthCode & yearCode
    Else
        BuildRICString = rootRIC & strikeStr & monthCode & yearCode
    End If

    ' Add ^{monthCode}{yearCode} suffix if maturity date is before today (expired option)
    ' NOTE: Expired suffix ALWAYS uses 2-digit year, regardless of OPTION_YEAR_DIGITS
    If maturityDate < Date Then
        Dim expiredYearCode As String
        expiredYearCode = Right(CStr(optionYear), 2)  ' Always 2 digits for expired suffix
        BuildRICString = BuildRICString & "^" & monthCodeCallForExpiredRIC & expiredYearCode
    End If
End Function

' ============================================
' FORMAT STRIKE FOR RIC
' ============================================

Function GetStrikeMultiplier() As Double
    ' Get strike multiplier from Config sheet
    ' Default: 10 (1 implied decimal). Use 100 for 2 implied decimals (quarter strikes)
    On Error Resume Next
    GetStrikeMultiplier = ThisWorkbook.Sheets(SHEET_CONFIG).Range("strikeMultiplier").Value
    On Error GoTo 0
    If GetStrikeMultiplier = 0 Then GetStrikeMultiplier = 10
End Function

Function FormatStrikeForRIC(strike As Double) As String
    ' Format strike for RIC using configurable multiplier
    ' Examples with multiplier=100: 80.75 -> "8075", 100 -> "10000"
    Dim strikeInt As Long
    Dim multiplier As Double

    multiplier = GetStrikeMultiplier()

    ' Multiply and convert to integer (avoids floating point format rounding)
    strikeInt = CLng(strike * multiplier)

    FormatStrikeForRIC = CStr(strikeInt)
End Function

Function FormatStrikeForWeeklyRIC(strike As Double) As String
    ' Format strike for a weekly option RIC using the configurable
    ' Config!strikeMultiplier (same as monthly via FormatStrikeForRIC).
    ' Examples with multiplier=10:  100 -> "1000", 100.5 -> "1005"
    ' Examples with multiplier=100: 100 -> "10000", 100.5 -> "10050"
    Dim multiplier As Double
    multiplier = GetStrikeMultiplier()
    FormatStrikeForWeeklyRIC = CStr(CLng(strike * multiplier))
End Function

' ============================================
' GET DAY OCCURRENCE IN MONTH (for weekly options)
' ============================================

Function GetDayOccurrenceInMonth(d As Date) As Integer
    ' Returns which occurrence of the weekday this date is in its month
    ' e.g., 2025-12-15 (Monday) -> 3 (3rd Monday of December)
    Dim firstOfMonth As Date
    Dim dayOfWeek As Integer
    Dim firstOccurrence As Date
    Dim occurrence As Integer

    firstOfMonth = DateSerial(Year(d), Month(d), 1)
    dayOfWeek = Weekday(d)  ' 1=Sunday, 2=Monday, etc.

    ' Find first occurrence of this weekday in the month
    firstOccurrence = firstOfMonth + ((dayOfWeek - Weekday(firstOfMonth) + 7) Mod 7)

    ' Calculate which occurrence this date is
    occurrence = ((Day(d) - Day(firstOccurrence)) \ 7) + 1

    GetDayOccurrenceInMonth = occurrence
End Function

' ============================================
' GET OPTION FREQUENCY FROM CONFIG
' ============================================

Function GetOptionFrequency() As String
    ' Returns "monthly" or "weekly" from Config sheet
    On Error Resume Next
    GetOptionFrequency = LCase(Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range(OPTION_FREQUENCY).Value))
    On Error GoTo 0

    ' Default to monthly if not set or invalid
    If GetOptionFrequency <> "weekly" Then
        GetOptionFrequency = "monthly"
    End If
End Function

' ============================================
' GET OPTION MONTH CODE METHOD FROM CONFIG
' ============================================

Function GetOptionMonthCodeMethod() As String
    ' Returns "Same Month" or "Next Month" from Config sheet
    ' Determines whether option month code uses maturity month or following month
    On Error Resume Next
    GetOptionMonthCodeMethod = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range(OPTION_MONTH_CODE_METHOD).Value)
    On Error GoTo 0

    ' Default to "Next Month" if not set or invalid
'    If LCase(GetOptionMonthCodeMethod) <> "same month" Then
'        GetOptionMonthCodeMethod = "Next Month"
'    Else
'        GetOptionMonthCodeMethod = "Same Month"
'    End If
End Function

' ============================================
' GET UNDERLYING MONTH MODE FROM CONFIG
' ============================================

Function GetUnderlyingMonthMode() As String
    ' Returns the underlying-future month rule from Config:
    '   "Same Month"    - underlying matches the option's month (default)
    '   "Quarter End"   - underlying is the next quarterly future (Mar/Jun/Sep/Dec)
    '                     Use this for instruments like 10Y UST options whose
    '                     underlying is always the next end-of-quarter future.
    '   "Next 2 Month"  - underlying is two months ahead of the option month
    ' Falls back to UNDERLYING_MONTH_MODE_DEFAULT when the named range is missing.
    Dim mode As String
    On Error Resume Next
    mode = Trim(CStr(ThisWorkbook.Sheets(SHEET_CONFIG).Range(UNDERLYING_MONTH_MODE_RANGE).Value))
    On Error GoTo 0

    If mode = "" Then
        GetUnderlyingMonthMode = UNDERLYING_MONTH_MODE_DEFAULT
    Else
        GetUnderlyingMonthMode = mode
    End If
End Function

' ============================================
' GET QUARTER END MONTH
' ============================================

Function GetQuarterEndMonth(maturityMonth As Integer) As Integer
    ' Returns the quarter end month (3, 6, 9, or 12) for the given month
    ' Used when underlyingMonthMode (Config) = "Quarter End"
    ' Q1: Dec, Jan, Feb -> March | Q2: Mar, Apr, May -> June
    ' Q3: Jun, Jul, Aug -> Sep   | Q4: Sep, Oct, Nov -> December
    Select Case maturityMonth
        Case 12, 1, 2: GetQuarterEndMonth = 3   ' Q1 -> March
        Case 3, 4, 5: GetQuarterEndMonth = 6    ' Q2 -> June
        Case 6, 7, 8: GetQuarterEndMonth = 9    ' Q3 -> September
        Case 9, 10, 11: GetQuarterEndMonth = 12 ' Q4 -> December
    End Select
End Function

' ============================================
' BUILD BLOOMBERG TICKER FROM CONFIG
' ============================================

Function BuildBloombergTicker(optionType As String, strike As Double, maturityDate As Date) As String
    ' Build Bloomberg ticker from rootBB + type + strike + maturity
    ' Format: "ES1 Index C 6000 12/19/2025"
    Dim rootBB As String

    On Error Resume Next
    rootBB = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootBB").Value)
    On Error GoTo 0

    If rootBB = "" Then
        BuildBloombergTicker = ""
        Exit Function
    End If

    BuildBloombergTicker = rootBB & " " & Left(optionType, 1) & " " & strike & " " & Format(maturityDate, "mm/dd/yyyy")
End Function

Function RICToBloomberg(ByVal ric As String, ByVal BBGRoot As String) As String
    '-----------------------------------------------------------
    ' Converts a Reuters RIC futures ticker to Bloomberg format
    ' Inputs:
    '   RIC     - Reuters ticker (e.g., "FGBMZ5", "FGBMZ25", "FGBMZ25^2")
    '   BBGRoot - Bloomberg root (e.g., "OE")
    ' Output:
    '   Bloomberg ticker (e.g., "OEZ25 Comdty")
    '-----------------------------------------------------------

    Dim monthCode As String
    Dim YearPart As String
    Dim YearNum As Integer
    Dim i As Integer
    Dim ValidMonths As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim cleanRIC As String
    Dim CaratPos As Integer

    ' Build ValidMonths from monthFutureBloomberg named range
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error Resume Next
    Set rng = ws.Range("monthFutureBloomberg")
    On Error GoTo ErrorHandler

    If rng Is Nothing Then
        RICToBloomberg = "#ERROR: Missing 'monthFutureBloomberg' range in Config"
        Exit Function
    End If

    ValidMonths = ""
    For Each cell In rng
        If Trim(cell.Value) <> "" Then
            ValidMonths = ValidMonths & Trim(cell.Value)
        End If
    Next cell

    On Error GoTo ErrorHandler

    ' Trim input
    cleanRIC = Trim(UCase(ric))
    BBGRoot = Trim(UCase(BBGRoot))

    ' Yellow-key suffix: default " Comdty". If the root (rootUnderlyingBB)
    ' ends with "Comdty" or "Index", strip it off and use it as the suffix.
    Dim bbgSuffix As String
    bbgSuffix = " Comdty"
    Dim rootSp As Long
    rootSp = InStrRev(BBGRoot, " ")
    If rootSp > 0 Then
        Dim rootLastWord As String
        rootLastWord = Mid(BBGRoot, rootSp + 1)
        If rootLastWord = "COMDTY" Then
            bbgSuffix = " Comdty"
            BBGRoot = RTrim(Left(BBGRoot, rootSp - 1))
        ElseIf rootLastWord = "INDEX" Then
            bbgSuffix = " Index"
            BBGRoot = RTrim(Left(BBGRoot, rootSp - 1))
        End If
    End If

    ' Strip expiration suffix (e.g., ^2) if present
    CaratPos = InStr(cleanRIC, "^")
    If CaratPos > 0 Then
        cleanRIC = Left(cleanRIC, CaratPos - 1)
    End If

    ' Validate inputs
    If Len(cleanRIC) < 2 Or Len(BBGRoot) = 0 Then
        RICToBloomberg = "#ERROR: Invalid input"
        Exit Function
    End If

    ' Extract year (last 1-2 digits from end)
    If IsNumeric(Right(cleanRIC, 2)) Then
        YearPart = Right(cleanRIC, 2)
        monthCode = Mid(cleanRIC, Len(cleanRIC) - 2, 1)
    ElseIf IsNumeric(Right(cleanRIC, 1)) Then
        YearPart = Right(cleanRIC, 1)
        monthCode = Mid(cleanRIC, Len(cleanRIC) - 1, 1)
    Else
        RICToBloomberg = "#ERROR: Cannot extract year"
        Exit Function
    End If

    ' Validate month code
    If InStr(ValidMonths, monthCode) = 0 Then
        RICToBloomberg = "#ERROR: Invalid month code"
        Exit Function
    End If

    ' Convert year to 2-digit format
    YearNum = CInt(YearPart)
    If YearNum < 10 Then
        YearPart = "2" & YearNum
    ElseIf YearNum >= 10 And YearNum < 100 Then
        YearPart = CStr(YearNum)
    End If

    ' Build Bloomberg ticker
    RICToBloomberg = BBGRoot & monthCode & YearPart & bbgSuffix
    Exit Function

ErrorHandler:
    RICToBloomberg = "#ERROR: " & Err.Description

End Function

Function GetWeekNumberFromDate(maturityDate As Date) As Integer
    '-----------------------------------------------------------
    ' Calculates which week of the month the date falls in (1-5)
    ' Based on which Friday of the month it is
    ' Week 1: days 1-7, Week 2: days 8-14, Week 3: days 15-21,
    ' Week 4: days 22-28, Week 5: days 29-31
    '-----------------------------------------------------------
    Dim dayOfMonth As Integer

    dayOfMonth = Day(maturityDate)

    If dayOfMonth <= 7 Then
        GetWeekNumberFromDate = 1
    ElseIf dayOfMonth <= 14 Then
        GetWeekNumberFromDate = 2
    ElseIf dayOfMonth <= 21 Then
        GetWeekNumberFromDate = 3
    ElseIf dayOfMonth <= 28 Then
        GetWeekNumberFromDate = 4
    Else
        GetWeekNumberFromDate = 5
    End If
End Function

Function GetUnderlyingBBGSuffix() As String
    ' Yellow-key suffix for underlying Bloomberg tickers, from Config!
    ' rootUnderlyingBB. " Index" if rootUnderlyingBB ends with "Index",
    ' otherwise the default " Comdty".
    Dim rub As String, sp As Long
    On Error Resume Next
    rub = UCase(Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootUnderlyingBB").Value))
    On Error GoTo 0
    sp = InStrRev(rub, " ")
    If sp > 0 Then
        If Mid(rub, sp + 1) = "INDEX" Then
            GetUnderlyingBBGSuffix = " Index"
            Exit Function
        End If
    End If
    GetUnderlyingBBGSuffix = " Comdty"
End Function

Function BuildOptionBloombergTicker(underlyingBloomTicker As String, optType As String, strike As Double) As String
    '-----------------------------------------------------------
    ' Builds option Bloomberg ticker from underlying Bloomberg ticker
    ' Inputs:
    '   underlyingBloomTicker - Bloomberg ticker of underlying (e.g., "OEZ25 Comdty")
    '   optType               - "CALL" or "PUT"
    '   strike                - Strike price (e.g., 70)
    ' Output:
    '   Option Bloomberg ticker (e.g., "OEZ25 C70")
    '-----------------------------------------------------------
    Dim baseTickerPart As String
    Dim CallPut As String
    Dim strikeStr As String
    Dim spacePos As Integer

    ' Strip suffix (e.g., " Comdty", " Index") from underlying ticker
    spacePos = InStr(underlyingBloomTicker, " ")
    If spacePos > 0 Then
        baseTickerPart = Left(underlyingBloomTicker, spacePos - 1)
    Else
        baseTickerPart = underlyingBloomTicker
    End If

    ' Determine Call/Put indicator
    If UCase(Left(optType, 1)) = "C" Then
        CallPut = "C"
    Else
        CallPut = "P"
    End If

    ' Format strike (remove decimals if whole number)
    If strike = Int(strike) Then
        strikeStr = CStr(CLng(strike))
    Else
        strikeStr = Format(strike, "0.0##")
    End If

    ' Build option Bloomberg ticker: "OEZ25C 70 Comdty" / "...Index"
    ' Yellow-key suffix follows Config!rootUnderlyingBB.
    BuildOptionBloombergTicker = baseTickerPart & CallPut & " " & strikeStr & GetUnderlyingBBGSuffix()
End Function

Function BuildWeeklyOptionBloombergTicker(underlyingBloomTicker As String, optType As String, strike As Double, weekNum As Integer) As String
    '-----------------------------------------------------------
    ' Builds weekly option Bloomberg ticker from underlying Bloomberg ticker
    ' Inputs:
    '   underlyingBloomTicker - Bloomberg ticker of underlying (e.g., "OEZ25 Comdty")
    '   optType               - "CALL" or "PUT"
    '   strike                - Strike price (e.g., 100.5)
    '   weekNum               - Week number (1-5)
    ' Output:
    '   Weekly option Bloomberg ticker, " Comdty" suffix appended.
    '
    ' Template mode: if rootBB (Config) contains a "(...)" token, rootBB is
    ' treated as a template and these placeholders are substituted
    ' (case-insensitive); any literal characters are kept verbatim:
    '   (Week Code)   -> week number 1-5
    '   (Month Code)  -> month letter  (from the underlying Bloomberg ticker)
    '   (Year Code)   -> year digits   (from the underlying Bloomberg ticker)
    '   (Year 1digit) -> last digit of the year
    '   (Option Code) -> "C" / "P"
    '   (Strike)      -> formatted strike
    '   e.g. rootBB = "IMDW_(Month Code)(Year Code)(Option Code)(Week Code)(Strike)"
    ' Yellow-key suffix: " Comdty" is appended by default. If the template
    ' itself ends with "Comdty" or "Index", that wins and nothing is added.
    ' Legacy mode: if rootBB has no token, the old fixed layout is used.
    '-----------------------------------------------------------
    Dim basePart As String
    Dim rootBB As String
    Dim rootUnderlyingBB As String
    Dim monthYear As String
    Dim spacePos As Integer
    Dim CallPut As String
    Dim strikeStr As String

    ' Strip suffix from underlying ticker: "OEZ25 Comdty" -> "OEZ25"
    spacePos = InStr(underlyingBloomTicker, " ")
    If spacePos > 0 Then
        basePart = Left(underlyingBloomTicker, spacePos - 1)
    Else
        basePart = underlyingBloomTicker
    End If

    ' Get rootBB (option root) and rootUnderlyingBB (underlying root) from Config
    On Error Resume Next
    rootBB = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootBB").Value)
    rootUnderlyingBB = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootUnderlyingBB").Value)
    On Error GoTo 0

    ' Strip any " Comdty" / " Index" yellow-key suffix from rootUnderlyingBB
    ' so only the bare root is used for the month/year split below.
    Dim ruSp As Long, ruLast As String
    ruSp = InStrRev(rootUnderlyingBB, " ")
    If ruSp > 0 Then
        ruLast = UCase(Mid(rootUnderlyingBB, ruSp + 1))
        If ruLast = "COMDTY" Or ruLast = "INDEX" Then
            rootUnderlyingBB = RTrim(Left(rootUnderlyingBB, ruSp - 1))
        End If
    End If

    ' Split basePart into root and month+year using UNDERLYING root
    ' e.g., "OEZ25" with rootUnderlyingBB="OE" -> monthYear="Z25"
    If Len(rootUnderlyingBB) > 0 And Left(basePart, Len(rootUnderlyingBB)) = rootUnderlyingBB Then
        monthYear = Mid(basePart, Len(rootUnderlyingBB) + 1)
    Else
        ' Fallback: assume 2-character underlying root
        monthYear = Mid(basePart, 3)
    End If

    ' Use rootBB (option root) for output; fallback to underlying root if not set
    If Len(rootBB) = 0 Then rootBB = rootUnderlyingBB
    If Len(rootBB) = 0 Then rootBB = Left(basePart, 2)

    ' Determine Call/Put indicator
    If UCase(Left(optType, 1)) = "C" Then
        CallPut = "C"
    Else
        CallPut = "P"
    End If

    ' Format strike (remove decimals if whole number)
    If strike = Int(strike) Then
        strikeStr = CStr(CLng(strike))
    Else
        strikeStr = Format(strike, "0.0##")
    End If

    ' Split monthYear (e.g. "Z25") into month letter + year digits
    Dim monthCode As String, yearCode As String
    monthCode = Left(monthYear, 1)
    yearCode = Mid(monthYear, 2)

    If InStr(rootBB, "(") > 0 Then
        ' Template mode: substitute placeholder tokens, keep literals verbatim
        Dim tpl As String
        tpl = rootBB
        tpl = Replace(tpl, "(Week Code)", CStr(weekNum), 1, -1, vbTextCompare)
        tpl = Replace(tpl, "(Month Code)", monthCode, 1, -1, vbTextCompare)
        tpl = Replace(tpl, "(Year Code)", yearCode, 1, -1, vbTextCompare)
        tpl = Replace(tpl, "(Year 1digit)", Right(yearCode, 1), 1, -1, vbTextCompare)
        tpl = Replace(tpl, "(Option Code)", CallPut, 1, -1, vbTextCompare)
        tpl = Replace(tpl, "(Strike)", strikeStr, 1, -1, vbTextCompare)
        tpl = RTrim(tpl)

        ' If the template already ends with a yellow-key suffix (Comdty or
        ' Index), use it as-is; otherwise append the default " Comdty".
        Dim lastWord As String, sp As Long
        sp = InStrRev(tpl, " ")
        If sp > 0 Then lastWord = Mid(tpl, sp + 1) Else lastWord = tpl
        If UCase(lastWord) = "COMDTY" Or UCase(lastWord) = "INDEX" Then
            BuildWeeklyOptionBloombergTicker = tpl
        Else
            BuildWeeklyOptionBloombergTicker = tpl & " Comdty"
        End If
    Else
        ' Legacy mode: old fixed layout "OE2Z25C 100 Comdty" / "...Index"
        BuildWeeklyOptionBloombergTicker = rootBB & CStr(weekNum) & monthYear & _
                                           CallPut & " " & strikeStr & GetUnderlyingBBGSuffix()
    End If
End Function

Function RICWeeklyOptionToBloomberg(ric As String, weekNum As Integer, _
    Optional StrikeDivisor As Double = 100) As String
    '-----------------------------------------------------------
    ' Converts a Reuters RIC futures option to Bloomberg format (Weekly)
    ' Inputs:
    '   RIC           - Reuters option ticker (e.g., "1E3W1005L25")
    '   WeekNum       - Week number (1-5)
    '   StrikeDivisor - Divisor to convert strike (default 100)
    ' Uses:
    '   Named Range "rootBB" - Bloomberg weekly option root
    ' Output:
    '   Bloomberg ticker (e.g., "OE2Z25 C100.50")
    ' Note: Call/Put determined from option month code (A-L=Call, M-X=Put)
    '-----------------------------------------------------------

    Dim BBGRoot As String
    Dim OptionMonthCode As String
    Dim FutureMonthCode As String
    Dim YearPart As String
    Dim StrikePart As String
    Dim CallPut As String
    Dim ValidFutureMonths As String
    Dim CallMonthCodes As String
    Dim PutMonthCodes As String
    Dim i As Integer
    Dim MonthPos As Integer
    Dim YearEndPos As Integer
    Dim StrikeValue As Double
    Dim TempRIC As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim MonthIndex As Integer

    On Error GoTo ErrorHandler

    ' Get Bloomberg root from named range
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    BBGRoot = Trim(UCase(ws.Range("rootBB").Value))

    ' Build ValidFutureMonths from monthFutureBloomberg named range
    On Error Resume Next
    Set rng = ws.Range("monthFutureBloomberg")
    On Error GoTo ErrorHandler

    If rng Is Nothing Then
        RICWeeklyOptionToBloomberg = "#ERROR: Missing 'monthFutureBloomberg' range in Config"
        Exit Function
    End If

    ValidFutureMonths = ""
    For Each cell In rng
        If Trim(cell.Value) <> "" Then
            ValidFutureMonths = ValidFutureMonths & Trim(cell.Value)
        End If
    Next cell

    ' Build Call month codes from monthCall named range
    CallMonthCodes = ""
    On Error Resume Next
    Set rng = ws.Range(MONTH_CALL)
    On Error GoTo ErrorHandler
    If Not rng Is Nothing Then
        For Each cell In rng
            If Trim(cell.Value) <> "" Then
                CallMonthCodes = CallMonthCodes & Trim(cell.Value)
            End If
        Next cell
    End If

    ' Build Put month codes from monthPut named range
    PutMonthCodes = ""
    On Error Resume Next
    Set rng = ws.Range(MONTH_PUT)
    On Error GoTo ErrorHandler
    If Not rng Is Nothing Then
        For Each cell In rng
            If Trim(cell.Value) <> "" Then
                PutMonthCodes = PutMonthCodes & Trim(cell.Value)
            End If
        Next cell
    End If

    On Error GoTo ErrorHandler

    ' Trim and uppercase input
    TempRIC = Trim(UCase(ric))

    ' Strip expiration suffix if present
    Dim CaratPos As Integer
    CaratPos = InStr(TempRIC, "^")
    If CaratPos > 0 Then
        TempRIC = Left(TempRIC, CaratPos - 1)
    End If

    ' Validate inputs
    If Len(TempRIC) < 4 Then
        RICWeeklyOptionToBloomberg = "#ERROR: Invalid RIC"
        Exit Function
    End If

    If Len(BBGRoot) = 0 Then
        RICWeeklyOptionToBloomberg = "#ERROR: rootBB named range is empty"
        Exit Function
    End If

    If weekNum < 1 Or weekNum > 5 Then
        RICWeeklyOptionToBloomberg = "#ERROR: WeekNum must be 1-5"
        Exit Function
    End If

    ' Find option month code position (look for call or put month codes followed by year digit)
    MonthPos = 0
    For i = 1 To Len(TempRIC)
        Dim ch As String
        ch = Mid(TempRIC, i, 1)
        If InStr(CallMonthCodes, ch) > 0 Or InStr(PutMonthCodes, ch) > 0 Then
            If i < Len(TempRIC) Then
                If IsNumeric(Mid(TempRIC, i + 1, 1)) Then
                    MonthPos = i
                    Exit For
                End If
            End If
        End If
    Next i

    If MonthPos = 0 Then
        RICWeeklyOptionToBloomberg = "#ERROR: Cannot find option month code"
        Exit Function
    End If

    OptionMonthCode = Mid(TempRIC, MonthPos, 1)

    ' Determine Call/Put from the option month code
    If InStr(CallMonthCodes, OptionMonthCode) > 0 Then
        CallPut = "C"
        MonthIndex = InStr(CallMonthCodes, OptionMonthCode)
    ElseIf InStr(PutMonthCodes, OptionMonthCode) > 0 Then
        CallPut = "P"
        MonthIndex = InStr(PutMonthCodes, OptionMonthCode)
    Else
        RICWeeklyOptionToBloomberg = "#ERROR: Cannot determine Call/Put from month code"
        Exit Function
    End If

    ' Get the corresponding future month code
    If MonthIndex > 0 And MonthIndex <= Len(ValidFutureMonths) Then
        FutureMonthCode = Mid(ValidFutureMonths, MonthIndex, 1)
    Else
        RICWeeklyOptionToBloomberg = "#ERROR: Cannot map to future month code"
        Exit Function
    End If

    ' Extract year (1-2 digits after month code)
    If MonthPos + 2 <= Len(TempRIC) And IsNumeric(Mid(TempRIC, MonthPos + 1, 2)) Then
        YearPart = Mid(TempRIC, MonthPos + 1, 2)
        YearEndPos = MonthPos + 2
    ElseIf IsNumeric(Mid(TempRIC, MonthPos + 1, 1)) Then
        YearPart = Mid(TempRIC, MonthPos + 1, 1)
        YearEndPos = MonthPos + 1
    Else
        RICWeeklyOptionToBloomberg = "#ERROR: Cannot extract year"
        Exit Function
    End If

    ' Convert year to 2-digit format
    If Len(YearPart) = 1 Then
        YearPart = "2" & YearPart
    End If

    ' Extract strike (numeric portion before month code)
    StrikePart = ""
    For i = 1 To MonthPos - 1
        If IsNumeric(Mid(TempRIC, i, 1)) Or Mid(TempRIC, i, 1) = "." Then
            StrikePart = StrikePart & Mid(TempRIC, i, 1)
        End If
    Next i

    ' If no strike before month, check after year
    If Len(StrikePart) = 0 Then
        For i = YearEndPos + 1 To Len(TempRIC)
            If IsNumeric(Mid(TempRIC, i, 1)) Or Mid(TempRIC, i, 1) = "." Then
                StrikePart = StrikePart & Mid(TempRIC, i, 1)
            End If
        Next i
    End If

    If Len(StrikePart) = 0 Then
        RICWeeklyOptionToBloomberg = "#ERROR: Cannot extract strike"
        Exit Function
    End If

    ' Convert strike to decimal
    StrikeValue = CDbl(StrikePart) / StrikeDivisor

    ' Format strike
    Dim strikeStr As String
    If StrikeValue = Int(StrikeValue) Then
        strikeStr = CStr(Int(StrikeValue))
    Else
        strikeStr = Format(StrikeValue, "0.0##")
    End If

    ' Build Bloomberg ticker with week number
    RICWeeklyOptionToBloomberg = BBGRoot & weekNum & FutureMonthCode & YearPart & " " & CallPut & strikeStr
    Exit Function

ErrorHandler:
    RICWeeklyOptionToBloomberg = "#ERROR: " & Err.Description

End Function

' ============================================
' GET MATURITY DATES FROM CONFIG
' ============================================

Function GetMaturityDates() As Collection
    Dim maturities As New Collection
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    Set rng = ws.Range("maturityDate")   ' Named range for maturities
    
    For Each cell In rng.Cells
        If IsDate(cell.Value) Then
            maturities.Add CDate(cell.Value)
        ElseIf IsEmpty(cell.Value) Then
            Exit For   ' stop at first empty cell
        End If
    Next cell
    
    If maturities.count = 0 Then
        MsgBox "No maturity dates found in " & SHEET_CONFIG & " sheet!", vbExclamation
    End If
    
    Set GetMaturityDates = maturities
End Function

' ============================================
' GET STRIKE RANGE
' ============================================

Function GetStrikeRange(optionType As String) As Collection
    Dim strikes As New Collection
    Dim ws As Worksheet
    Dim minStrike As Double
    Dim maxStrike As Double
    Dim stepSize As Double
    Dim currentStrike As Double
    
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    stepSize = ws.Range("steps").Value
    
    If optionType = "PUT" Then
        minStrike = ws.Range("minStrikePut").Value
        maxStrike = ws.Range("maxStrikePut").Value
    Else
        minStrike = ws.Range("minStrikeCall").Value
        maxStrike = ws.Range("maxStrikeCall").Value
    End If
    
    If minStrike = 0 Or maxStrike = 0 Or stepSize = 0 Then
        MsgBox "Invalid strike range configuration!", vbExclamation
        Exit Function
    End If
    
    currentStrike = minStrike
    Do While currentStrike <= maxStrike
        strikes.Add currentStrike
        currentStrike = currentStrike + stepSize
    Loop
    
    Set GetStrikeRange = strikes
End Function

' ============================================
' GET MONTH CODE FROM TABLE
' ============================================

Function GetMonthCodeFromTable(monthNum As Integer, optionType As String) As String
    Dim ws As Worksheet
    Dim i As Integer
    Dim rng As Range
    Dim Offset As Integer
    
    
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    
    ' Use named ranges for CALL vs PUT
    If optionType = "CALL" Then
        Set rng = ws.Range(MONTH_CALL)
        Offset = 1
    Else
        Set rng = ws.Range(MONTH_PUT)
        Offset = 2
    End If
    
    ' Loop through the range to find matching month number
    For i = 1 To rng.Rows.count
        If ws.Cells(rng.Row + i - 1, rng.Column - Offset).Value = monthNum Then  ' Column E has month numbers
            GetMonthCodeFromTable = rng.Cells(i, 1).Value
            Exit Function
        End If
    Next i
    
    ' If not found, raise error
    Err.Raise vbObjectError + 513, "GetMonthCodeFromTable", _
              "No month code found for month " & monthNum & " and option type " & optionType
End Function

Function GetWeeklyTypeCodeFromRange(optionType As String) As String
    ' Get the type code for weekly options from named range
    ' Weekly options have no month notion - just CALL or PUT code
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)

    On Error Resume Next
    If optionType = "CALL" Then
        Set rng = ws.Range(WEEKLY_CALL)
    Else
        Set rng = ws.Range(WEEKLY_PUT)
    End If
    On Error GoTo 0

    If rng Is Nothing Then
        Err.Raise vbObjectError + 514, "GetWeeklyTypeCodeFromRange", _
                  "Named range not found for weekly " & optionType
    End If

    GetWeeklyTypeCodeFromRange = rng.Cells(1, 1).Value
End Function

Function GetFutureMonthCode(monthNum As Integer) As String
    '-----------------------------------------------------------
    ' Returns the future month code for a given month number (1-12)
    ' Reads from monthFutureBloomberg named range in Config sheet
    ' Standard future month codes: F,G,H,J,K,M,N,Q,U,V,X,Z
    '-----------------------------------------------------------
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Integer

    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)

    On Error Resume Next
    Set rng = ws.Range("monthFutureBloomberg")
    On Error GoTo 0

    If rng Is Nothing Then
        Err.Raise vbObjectError + 515, "GetFutureMonthCode", _
                  "Named range 'monthFutureBloomberg' not found in Config sheet"
    End If

    ' The range should have 12 cells, one for each month
    ' Month 1 (Jan) = first cell, Month 12 (Dec) = last cell
    If monthNum < 1 Or monthNum > 12 Then
        Err.Raise vbObjectError + 516, "GetFutureMonthCode", _
                  "Invalid month number: " & monthNum
    End If

    ' Get the month code from the appropriate position
    i = 0
    For Each cell In rng
        i = i + 1
        If i = monthNum Then
            GetFutureMonthCode = Trim(cell.Value)
            Exit Function
        End If
    Next cell

    ' If we get here, range didn't have enough cells
    Err.Raise vbObjectError + 517, "GetFutureMonthCode", _
              "monthFutureBloomberg range does not have " & monthNum & " cells"
End Function

' ============================================
' GET SPECIFIC RIC FOR OPTION
' ============================================

Function GetRICForOption(strike As Double, maturityDate As Date, optionType As String) As String
    ' Quick function to get single RIC string (used in main process)
    Dim ricDict As Object
    Set ricDict = CreateRICInfo(maturityDate, strike, optionType)
    GetRICForOption = ricDict("FullRIC")
End Function

' ============================================
' BUILD RIC LOOKUP DICTIONARY
' ============================================

Function BuildRICLookupDictionary() As Object
    ' Creates a dictionary for fast RIC lookups
    ' Key: "Strike_Maturity_Type" -> Value: Full RIC
    
    Dim lookupDict As Object
    Dim ricList As Collection
    Dim ricDict As Object
    Dim lookupKey As String
    Dim ric As Variant
    
    Set lookupDict = CreateObject("Scripting.Dictionary")
    Set ricList = BuildCompleteRICList()
    
    For Each ric In ricList
        Set ricDict = ric
        
        ' Create lookup key
        lookupKey = ricDict("Strike") & "_" & _
                   Format(ricDict("Maturity"), "yyyymmdd") & "_" & _
                   ricDict("OptionType")
        
        ' Add to lookup dictionary
        lookupDict.Add lookupKey, ricDict("FullRIC")
    Next
    
    Set BuildRICLookupDictionary = lookupDict
End Function


' ============================================
' DOWNLOAD FROM OPTION CHAIN
' ============================================

Sub DownloadFromChain()
    ' Downloads option chain from LSEG and populates RIC_List sheet
    ' Uses OnTime chain pattern for async LSEG refresh (non-blocking)
    Dim ws As Worksheet
    Dim chainRIC As String

    On Error GoTo ErrorHandler

    ' Initialize state
    g_ChainState = CHAIN_STATE_IDLE
    g_ChainStopRequested = False
    g_ChainRefreshCheckCount = 0
    g_ChainIndex = 0
    g_ChainStepSize = 7

    ' Get root RIC from config
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    g_ChainRootRIC = Trim(ws.Range("rootRIC").Value)

    If g_ChainRootRIC = "" Then
        MsgBox "Please specify root RIC in Config sheet!", vbExclamation
        Exit Sub
    End If

    ' Create chain RIC for option chain download
    chainRIC = "0#" & g_ChainRootRIC & "+"

    ' Create or clear chain download sheet
    Set g_ChainSheet = CreateChainDownloadSheet()

    ' Setup chain download formula
    g_ChainSheet.Range("A1").Value = "Chain RIC"
    g_ChainSheet.Range("B1").Value = "Chain Data"
    g_ChainSheet.Range("C1").Value = "Status"
    g_ChainSheet.Range("A1:E1").Font.Bold = True

    ' Add chain download formula
    g_ChainSheet.Range("A2").Value = chainRIC
    g_ChainSheet.Range("B2").Formula = "=@TR(""" & chainRIC & """,""CF_NAME"",""CH=Fd RH=IN"")"
    g_ChainSheet.Range("C2").Value = "Downloading chains..."

    Application.StatusBar = "Downloading option chain of chains for " & g_ChainRootRIC & "..."

    ' Set state to downloading chains
    g_ChainState = CHAIN_STATE_DOWNLOADING_CHAINS
    g_ChainRefreshCheckCount = 0

    ' Trigger LSEG refresh (non-blocking call)
    Application.Run "WorkspaceRefreshWorksheet", True, 120000, g_ChainSheet.Name

    ' Schedule check via OnTime (non-blocking - allows LSEG to populate data)
    Application.OnTime Now + TimeValue("00:00:05"), "DownloadFromChain_CheckChainReady"

    ' VBA execution ends here - OnTime chain runs asynchronously
    Exit Sub

ErrorHandler:
    g_ChainState = CHAIN_STATE_IDLE
    Application.StatusBar = False
    MsgBox "Error in DownloadFromChain: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number, vbCritical
    If Not g_ChainSheet Is Nothing Then
        g_ChainSheet.Range("C2").Value = "Error: " & Err.Description
    End If
End Sub

' ============================================
' PHASE 1: Check if chain-of-chains data is ready
' ============================================
Sub DownloadFromChain_CheckChainReady()
    Dim readyCount As Long
    Dim totalCount As Long

    ' Force recalculation before checking
    g_ChainSheet.Calculate

    ' Check stop flag
    If g_ChainStopRequested Then
        DownloadFromChain_Abort
        Exit Sub
    End If

    ' Check timeout (60 checks x 3 sec = ~3 min)
    g_ChainRefreshCheckCount = g_ChainRefreshCheckCount + 1
    If g_ChainRefreshCheckCount > 60 Then
        Application.StatusBar = "Chain download timeout - proceeding anyway..."
        DownloadFromChain_ProcessChains
        Exit Sub
    End If

    ' Check if chain data is ready (look for data in B3)
    If Not IsEmpty(g_ChainSheet.Range("B3").Value) And g_ChainSheet.Range("B3").Value <> "0" Then
        ' Data ready, proceed to process chains
        Application.StatusBar = "Chain data ready - processing..."
        DownloadFromChain_ProcessChains
    Else
        ' Still waiting, check for LSEG status messages
        Dim cellText As String
        cellText = CStr(g_ChainSheet.Range("B2").Text)

        If InStr(1, cellText, "Retrieving", vbTextCompare) > 0 Or _
           InStr(1, cellText, "Requesting", vbTextCompare) > 0 Or _
           InStr(1, cellText, "Loading", vbTextCompare) > 0 Then
            ' Still downloading
            Application.StatusBar = "Downloading chains... (check #" & g_ChainRefreshCheckCount & ")"
        Else
            Application.StatusBar = "Waiting for chain data... (check #" & g_ChainRefreshCheckCount & ")"
        End If

        ' Reschedule check
        Application.OnTime Now + TimeValue("00:00:03"), "DownloadFromChain_CheckChainReady"
    End If
End Sub

' ============================================
' PHASE 2: Process chains and setup option downloads
' ============================================
Sub DownloadFromChain_ProcessChains()
    Dim lastRow As Long
    Dim i As Long
    Dim chainRICCode As String
    Dim cleanChainRIC As String
    Dim optionColumn As Long
    Dim chainCount As Long

    On Error GoTo ErrorHandler

    ' Check if data was actually downloaded
    If IsEmpty(g_ChainSheet.Range("B3").Value) Or g_ChainSheet.Range("B3").Value = "0" Then
        g_ChainSheet.Range("C2").Value = "No data"
        MsgBox "No option chain data found for " & g_ChainRootRIC & ". Please check if the root RIC is correct.", vbExclamation
        g_ChainState = CHAIN_STATE_IDLE
        Application.StatusBar = False
        Exit Sub
    End If

    g_ChainState = CHAIN_STATE_PROCESSING_CHAINS
    g_ChainSheet.Range("C2").Value = "Processing chains..."
    Application.StatusBar = "Processing chain data..."

    ' Setup RIC_List sheet
    Set g_RICListSheet = SetupRICListSheetForChain()

    ' Find last row with data in chain sheet
    lastRow = g_ChainSheet.Cells(g_ChainSheet.Rows.count, "B").End(xlUp).Row

    ' Setup headers for Stage 2 processing
    g_ChainSheet.Range("D1").Value = "Chain Index"
    g_ChainSheet.Range("E1").Value = "Clean Chain RIC"
    g_ChainSheet.Range("F1").Value = "Option Column"

    ' Build arrays of all chain RICs and their column positions
    chainCount = 0
    ReDim g_ChainCleanRICs(0 To lastRow - 3)
    ReDim g_ChainColumns(0 To lastRow - 3)

    For i = 3 To lastRow
        chainRICCode = Trim(CStr(g_ChainSheet.Cells(i, 2).Value))

        If chainRICCode = "" Or chainRICCode = "0" Then GoTo NextChainRIC

        ' Extract clean chain RIC
        If Left(chainRICCode, 1) = "/" Then
            cleanChainRIC = Mid(chainRICCode, 2)
        Else
            cleanChainRIC = chainRICCode
        End If

        ' Calculate option column
        optionColumn = 7 + chainCount * g_ChainStepSize

        ' Store in arrays
        g_ChainCleanRICs(chainCount) = cleanChainRIC
        g_ChainColumns(chainCount) = optionColumn

        ' Add column header
        g_ChainSheet.Cells(1, optionColumn).Value = "Chain " & chainCount & " (" & cleanChainRIC & ")"

        chainCount = chainCount + 1

NextChainRIC:
    Next i

    ' Resize arrays to actual count
    If chainCount > 0 Then
        ReDim Preserve g_ChainCleanRICs(0 To chainCount - 1)
        ReDim Preserve g_ChainColumns(0 To chainCount - 1)
    End If

    g_ChainTotalChains = chainCount
    g_ChainBatchStart = 0

    If g_ChainTotalChains = 0 Then
        g_ChainSheet.Range("C2").Value = "No valid chains found"
        MsgBox "No valid chain RICs found.", vbExclamation
        g_ChainState = CHAIN_STATE_IDLE
        Application.StatusBar = False
        Exit Sub
    End If

    g_ChainSheet.Range("C2").Value = "Downloading " & g_ChainTotalChains & " option chains in batches of " & CHAIN_BATCH_SIZE & "..."

    ' Start batched download
    DownloadFromChain_ProcessNextBatch
    Exit Sub

ErrorHandler:
    g_ChainState = CHAIN_STATE_IDLE
    Application.StatusBar = False
    MsgBox "Error in DownloadFromChain_ProcessChains: " & Err.Description, vbCritical
End Sub

' ============================================
' PHASE 2b: Process next batch of chains
' ============================================
Sub DownloadFromChain_ProcessNextBatch()
    Dim batchEnd As Long
    Dim i As Long

    ' Check stop flag
    If g_ChainStopRequested Then
        DownloadFromChain_Abort
        Exit Sub
    End If

    batchEnd = g_ChainBatchStart + CHAIN_BATCH_SIZE - 1
    If batchEnd > g_ChainTotalChains - 1 Then batchEnd = g_ChainTotalChains - 1

    Application.StatusBar = "Downloading chains " & (g_ChainBatchStart + 1) & "-" & (batchEnd + 1) & " of " & g_ChainTotalChains & "..."
    g_ChainSheet.Range("C2").Value = "Downloading chains " & (g_ChainBatchStart + 1) & "-" & (batchEnd + 1) & " of " & g_ChainTotalChains

    ' Place TR formulas for this batch
    For i = g_ChainBatchStart To batchEnd
        DownloadOptionsFromSingleChain g_ChainSheet, g_ChainColumns(i), g_ChainCleanRICs(i)
    Next i

    ' Set state and trigger refresh
    g_ChainState = CHAIN_STATE_DOWNLOADING_OPTIONS
    g_ChainRefreshCheckCount = 0

    ' Trigger LSEG refresh (non-blocking)
    Application.Run "WorkspaceRefreshWorksheet", True, 120000, g_ChainSheet.Name

    ' Schedule check via OnTime
    Application.OnTime Now + TimeValue("00:00:05"), "DownloadFromChain_CheckBatchReady"
End Sub

' ============================================
' PHASE 2c: Check if current batch is ready
' ============================================
Sub DownloadFromChain_CheckBatchReady()
    Dim readyCount As Long
    Dim totalCount As Long

    ' Force recalculation before checking
    g_ChainSheet.Calculate

    ' Check stop flag
    If g_ChainStopRequested Then
        DownloadFromChain_Abort
        Exit Sub
    End If

    ' Check timeout (60 checks x 3 sec = ~3 min)
    g_ChainRefreshCheckCount = g_ChainRefreshCheckCount + 1
    If g_ChainRefreshCheckCount > 60 Then
        Application.StatusBar = "Batch timeout - proceeding to next batch..."
        GoTo AdvanceBatch
    End If

    ' Check if data is ready
    If IsLSEGDataReady(g_ChainSheet, readyCount, totalCount, 10) Then
        Application.StatusBar = "Batch ready (" & readyCount & "/" & totalCount & ") - advancing..."
        GoTo AdvanceBatch
    Else
        Application.StatusBar = "Downloading chains " & (g_ChainBatchStart + 1) & "-" & _
            Application.Min(g_ChainBatchStart + CHAIN_BATCH_SIZE, g_ChainTotalChains) & " of " & g_ChainTotalChains & _
            "... " & readyCount & "/" & totalCount & " ready (check #" & g_ChainRefreshCheckCount & ")"
        ' Reschedule check
        Application.OnTime Now + TimeValue("00:00:03"), "DownloadFromChain_CheckBatchReady"
    End If
    Exit Sub

AdvanceBatch:
    g_ChainBatchStart = g_ChainBatchStart + CHAIN_BATCH_SIZE

    If g_ChainBatchStart < g_ChainTotalChains Then
        ' More chains to process - schedule next batch
        Application.OnTime Now + TimeValue("00:00:02"), "DownloadFromChain_ProcessNextBatch"
    Else
        ' All batches done - proceed to completion
        Application.StatusBar = "All chain batches downloaded - finalizing..."
        Application.OnTime Now + TimeValue("00:00:02"), "DownloadFromChain_Complete"
    End If
End Sub

' ============================================
' PHASE 3: (CheckOptionsReady removed - replaced by batched DownloadFromChain_CheckBatchReady)
' ============================================

' ============================================
' PHASE 4: Complete - process option data to RIC_List
' ============================================
Sub DownloadFromChain_Complete()
    On Error GoTo ErrorHandler

    g_ChainState = CHAIN_STATE_PROCESSING_OPTIONS
    g_ChainSheet.Range("C2").Value = "Processing options..."
    Application.StatusBar = "Processing option data to RIC_List..."

    ' Process all downloaded option data
    ProcessAllOptionDataByColumns g_ChainSheet, g_RICListSheet, g_ChainTotalChains, g_ChainStepSize

    ' Reset state
    g_ChainState = CHAIN_STATE_IDLE
    g_ChainSheet.Range("C2").Value = "Complete"
    Application.StatusBar = False

    MsgBox "Option chain download complete! Check " & SHEET_RIC_LIST & " sheet for results.", vbInformation
    Exit Sub

ErrorHandler:
    g_ChainState = CHAIN_STATE_IDLE
    Application.StatusBar = False
    MsgBox "Error in DownloadFromChain_Complete: " & Err.Description, vbCritical
End Sub

' ============================================
' Stop handler for DownloadFromChain
' ============================================
Sub StopDownloadFromChain()
    g_ChainStopRequested = True
    Application.StatusBar = "Stop requested - will halt after current operation..."
    MsgBox "Download will stop after current phase completes.", vbInformation
End Sub

' ============================================
' Abort handler for DownloadFromChain
' ============================================
Sub DownloadFromChain_Abort()
    g_ChainState = CHAIN_STATE_IDLE
    g_ChainStopRequested = False
    Application.StatusBar = False
    MsgBox "Download stopped.", vbInformation
End Sub

' ============================================
' CREATE CHAIN DOWNLOAD SHEET
' ============================================

Function CreateChainDownloadSheet() As Worksheet
    Dim ws As Worksheet
    Dim sheetName As String

    sheetName = "ChainDownload"

    ' Check if sheet already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create new sheet if it doesn't exist
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    Else
        ' Clear existing sheet contents
        ws.Cells.Clear
    End If

    Set CreateChainDownloadSheet = ws
End Function

' ============================================
' PROCESS CHAIN DATA
' ============================================

Sub ProcessChainData(chainSheet As Worksheet)
    Dim ricListSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim chainRICCode As String
    Dim cleanChainRIC As String
    Dim chainIndex As Long
    Dim optionColumn As Long
    Dim stepSize As Long

    Application.StatusBar = "Processing chain data..."
    stepSize = 7

    ' Setup RIC_List sheet
    Set ricListSheet = SetupRICListSheetForChain()

    ' Find last row with data in chain sheet (from initial chain download)
    lastRow = chainSheet.Cells(chainSheet.Rows.count, "B").End(xlUp).Row

    ' Setup headers for Stage 2 processing
    chainSheet.Range("D1").Value = "Chain Index"
    chainSheet.Range("E1").Value = "Clean Chain RIC"
    chainSheet.Range("F1").Value = "Option Column"

    chainIndex = 0  ' Track chain processing order

    ' Stage 1: Process each chain RIC from the chain-of-chains download
    For i = 3 To lastRow  ' Start from row 3 (skip header and first returned)
        ' Get the chain RIC from the first download
        chainRICCode = Trim(CStr(chainSheet.Cells(i, 2).Value))

        ' Skip if no data
        If chainRICCode = "" Or chainRICCode = "0" Then GoTo NextChainRIC

        ' Extract clean chain RIC (remove leading "/" if present)
        If Left(chainRICCode, 1) = "/" Then
            cleanChainRIC = Mid(chainRICCode, 2)
        Else
            cleanChainRIC = chainRICCode
        End If

        ' Calculate option column (start from column G = 7)
        optionColumn = 7 + chainIndex * stepSize

        ' Store chain processing information
        'chainSheet.Cells(i, 4).Value = chainIndex  ' Chain Index
        'chainSheet.Cells(i, 5).Value = cleanChainRIC  ' Clean Chain RIC
        'chainSheet.Cells(i, 6).Value = optionColumn  ' Option Column

        ' Add column header for this chain's options
        chainSheet.Cells(1, optionColumn).Value = "Chain " & chainIndex & " (" & cleanChainRIC & ")"

        ' Stage 2: Download individual options from this chain RIC
        Application.StatusBar = "Downloading options from chain " & chainIndex & ": " & cleanChainRIC
        DownloadOptionsFromSingleChain chainSheet, optionColumn, cleanChainRIC

        chainIndex = chainIndex + 1

NextChainRIC:
    Next i
    
    ' Call Worksheet refresh
    RefreshLSEGWithTimeout chainSheet, 120

    ' Wait for all TR formulas to refresh
    Application.StatusBar = "Waiting for data refresh..."
    Application.Wait Now + TimeValue("00:00:05")

    ' Stage 3: Process all downloaded option data and copy to RIC_List
    Application.StatusBar = "Processing option data..."
    ProcessAllOptionDataByColumns chainSheet, ricListSheet, chainIndex, stepSize

    Application.StatusBar = False
End Sub

' ============================================
' DOWNLOAD OPTIONS FROM SINGLE CHAIN
' ============================================

Sub DownloadOptionsFromSingleChain(chainSheet As Worksheet, optionColumn As Long, chainRIC As String)
    ' Downloads individual option RICs from a single chain RIC
    ' Uses the chain RIC to get the list of option instruments
    ' Places formula in the specified column to avoid collisions

    ' Add the chain RIC formula to download individual options
    ' Place the formula in row 2 of the specified column
    chainSheet.Cells(2, optionColumn).Formula = _
        "=@TR(""" & chainRIC & """,""CF_NAME;STRIKE_PRC;EXPIR_DATE;PUTCALLIND;UNDERLYING"",""CH=Fd RH=IN"")"
End Sub

' ============================================
' SETUP RIC_LIST SHEET FOR CHAIN
' ============================================

Function SetupRICListSheetForChain() As Worksheet
    Dim ws As Worksheet

    ' Get or create RIC_List sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_RIC_LIST)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = SHEET_RIC_LIST
    Else
        ' Clear existing content
        ws.Cells.Clear
    End If

    ' Setup headers
    With ws
        .Range("A1").Value = "RIC"
        .Range("B1").Value = "Maturity"
        .Range("C1").Value = "Strike"
        .Range("D1").Value = "Type"
        .Range("E1").Value = "Month Code"
        .Range("F1").Value = "Year"
        .Range("G1").Value = "Underlying LSEG"
        .Range("H1").Value = "Bloom_Ticker"
        .Range("I1").Value = "Processed"
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = RGB(200, 200, 200)
    End With

    Set SetupRICListSheetForChain = ws
End Function

' ============================================
' PROCESS ALL OPTION DATA BY COLUMNS
' ============================================

Sub ProcessAllOptionDataByColumns(chainSheet As Worksheet, ricListSheet As Worksheet, totalChains As Long, stepSize As Long)
    ' New column-based approach to process option data from separate columns
    Dim col As Long
    Dim ricListRow As Long
    Dim totalOptions As Long
    Dim errorCount As Long
    Dim startColumn As Long
    Dim optionColumn As Long

    On Error GoTo ErrorHandler

    ricListRow = 2  ' Start from row 2 (after header)
    totalOptions = 0
    errorCount = 0
    startColumn = 7  ' Options start from column G

    chainSheet.Calculate

    If totalChains = 0 Then
        MsgBox "No chains found to process.", vbExclamation
        Exit Sub
    End If

    ' Process each chain's option column
    For col = 0 To totalChains - 1
        optionColumn = startColumn + col * stepSize

        Application.StatusBar = "Processing chain " & col & " options from column " & Chr(64 + optionColumn) & "..."

        ' Process this column's option data
        ProcessSingleColumnOptions chainSheet, ricListSheet, optionColumn, ricListRow, totalOptions, errorCount
    Next col

    ' Format the RIC_List sheet
    FormatRICListSheet ricListSheet, ricListRow

    ' Show completion message
    ShowCompletionMessage totalOptions, errorCount, totalChains
    Exit Sub

ErrorHandler:
    MsgBox "Error in ProcessAllOptionDataByColumns: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Processing stopped at chain " & col, vbCritical
End Sub

' ============================================
' PROCESS SINGLE COLUMN OPTIONS
' ============================================

Sub ProcessSingleColumnOptions(chainSheet As Worksheet, ricListSheet As Worksheet, _
                              optionColumn As Long, ByRef ricListRow As Long, _
                              ByRef totalOptions As Long, ByRef errorCount As Long)
    Dim lastRow As Long
    Dim i As Long
    Dim optionDataText As String
    Dim optionLines As Variant
    Dim j As Long
    Dim optionFields As Variant
    Dim ricCode As String
    Dim strike As Variant
    Dim expirDate As Variant
    Dim putCallInd As String
    Dim optionType As String
    Dim monthCode As String
    Dim yearCode As String
    Dim RICvalue As String
    Dim methodRICBB As String
    Dim rootUnderlyingBB As String
    Dim underlyingRIC As String

    ' Read Bloomberg conversion method from Config
    methodRICBB = ""
    rootUnderlyingBB = ""
    On Error Resume Next
    methodRICBB = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("methodRICBB").Value)
    rootUnderlyingBB = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootUnderlyingBB").Value)
    On Error GoTo 0

    ' Find last row with data in this column
    lastRow = chainSheet.Cells(chainSheet.Rows.count, optionColumn).End(xlUp).Row

    ' Process each row in this column starting from row 2
    For i = 4 To lastRow
        optionDataText = Trim(CStr(chainSheet.Cells(i, optionColumn).Value))

        ' Skip if no data
        If optionDataText = "" Or optionDataText = "0" Then GoTo NextOptionRow
        
        RICvalue = ""
        If Left(chainSheet.Cells(i, optionColumn).Value, 1) = "/" Then
            RICvalue = Mid(chainSheet.Cells(i, optionColumn).Value, 2)
        Else
            RICvalue = chainSheet.Cells(i, optionColumn).Value
        End If
        
        ' Populate RIC_List row
        With ricListSheet
            .Cells(ricListRow, 1).Value = RICvalue  ' RIC
            .Cells(ricListRow, 2).Value = chainSheet.Cells(i, optionColumn + 3).Value ' Maturity
            .Cells(ricListRow, 3).Value = chainSheet.Cells(i, optionColumn + 2).Value ' Strike
            .Cells(ricListRow, 4).Value = chainSheet.Cells(i, optionColumn + 4).Value ' Type
            .Cells(ricListRow, 5).Value = "n/a" ' Month Code
            .Cells(ricListRow, 6).Value = "n/a"  ' Year
            .Cells(ricListRow, 7).Value = chainSheet.Cells(i, optionColumn + 5).Value ' Underlying

            ' Build Bloomberg ticker based on methodRICBB
            underlyingRIC = CStr(chainSheet.Cells(i, optionColumn + 5).Value)
            If UCase(methodRICBB) = "FUTURE" Then
                If rootUnderlyingBB = "" Then
                    MsgBox "Missing 'rootUnderlyingBB' named range in Config sheet!" & vbNewLine & _
                           "Please set the Bloomberg root ticker for the underlying.", _
                           vbCritical, "Configuration Error"
                    Exit Sub
                End If
                ' Convert underlying RIC to Bloomberg format
                .Cells(ricListRow, 8).Value = RICToBloomberg(underlyingRIC, rootUnderlyingBB)
            Else
                MsgBox "Invalid or missing 'methodRICBB' in Config sheet!" & vbNewLine & _
                       "Current value: '" & methodRICBB & "'" & vbNewLine & _
                       "Supported methods: 'Future'", _
                       vbCritical, "Configuration Error"
                Exit Sub
            End If

            .Cells(ricListRow, 9).Value = "No"  ' Processed
        End With
        ricListRow = ricListRow + 1
        totalOptions = totalOptions + 1

NextOptionRow:
    Next i
End Sub


' ============================================
' FORMAT RIC LIST SHEET
' ============================================

Sub FormatRICListSheet(ricListSheet As Worksheet, ricListRow As Long)
    With ricListSheet
        .Columns("A:I").AutoFit
        .Range("B:B").NumberFormat = "mm/dd/yyyy"
        .Range("C:C").NumberFormat = "#,##0"

        ' Add conditional formatting to Processed column (I)
        If ricListRow > 2 Then
            With .Range("I2:I" & ricListRow - 1).FormatConditions
                .Delete
                .Add Type:=xlTextString, String:="Yes", TextOperator:=xlContains
                .Item(.count).Interior.Color = RGB(200, 255, 200)
                .Add Type:=xlTextString, String:="Error", TextOperator:=xlContains
                .Item(.count).Interior.Color = RGB(255, 200, 200)
            End With
        End If
    End With
End Sub

' ============================================
' SHOW COMPLETION MESSAGE
' ============================================

Sub ShowCompletionMessage(totalOptions As Long, errorCount As Long, totalChains As Long)
    Dim resultMsg As String
    resultMsg = "Processed " & totalOptions & " option RICs from " & totalChains & " option chains!" & vbNewLine & _
                "Data copied to " & SHEET_RIC_LIST & " sheet."

    If errorCount > 0 Then
        resultMsg = resultMsg & vbNewLine & vbNewLine & "Note: " & errorCount & " items had month code generation errors."
    End If

    MsgBox resultMsg, vbInformation
End Sub


' ============================================
' HELPER FUNCTIONS TO PARSE TR RESULTS
' ============================================
' Note: These functions are kept for compatibility but are no longer
' used in the main DownloadFromChain process

Function GetStrikeFromTRResult(chainSheet As Worksheet, rowNum As Long) As Variant
    ' Parse strike price from TR result - DEPRECATED
    ' New process uses direct parsing in ProcessAllOptionData
    Dim trResult As String
    Dim parts As Variant

    trResult = CStr(chainSheet.Cells(rowNum, 5).Value)  ' Updated to column E
    parts = Split(trResult, ";")

    If UBound(parts) >= 1 Then
        GetStrikeFromTRResult = Val(parts(1))
    Else
        GetStrikeFromTRResult = 0
    End If
End Function

Function GetExpiryFromTRResult(chainSheet As Worksheet, rowNum As Long) As Variant
    ' Parse expiry date from TR result - DEPRECATED
    ' New process uses direct parsing in ProcessAllOptionData
    Dim trResult As String
    Dim parts As Variant

    trResult = CStr(chainSheet.Cells(rowNum, 5).Value)  ' Updated to column E
    parts = Split(trResult, ";")

    If UBound(parts) >= 2 Then
        GetExpiryFromTRResult = parts(2)
    Else
        GetExpiryFromTRResult = ""
    End If
End Function

Function GetPutCallFromTRResult(chainSheet As Worksheet, rowNum As Long) As String
    ' Parse put/call indicator from TR result - DEPRECATED
    ' New process uses direct parsing in ProcessAllOptionData
    Dim trResult As String
    Dim parts As Variant

    trResult = CStr(chainSheet.Cells(rowNum, 5).Value)  ' Updated to column E
    parts = Split(trResult, ";")

    If UBound(parts) >= 3 Then
        GetPutCallFromTRResult = Trim(parts(3))
    Else
        GetPutCallFromTRResult = ""
    End If
End Function



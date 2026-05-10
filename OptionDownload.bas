' ============================================
' MODULE OPTIONDOWNLOAD
' ============================================

Option Explicit

' Configuration Variables
Public g_NamePrefix As String
Public g_RootRIC As String
Public g_StrikeStep As Integer
Public g_LotSize As Long
Public g_Currency As String
Public g_DateStart As Date
Public g_DateEnd As Date
Public g_PutStrikeMin As Double
Public g_PutStrikeMax As Double
Public g_CallStrikeMin As Double
Public g_CallStrikeMax As Double

' Progress Tracking
Public g_CurrentMaturity As Date
Public g_CurrentStrike As Double
Public g_CurrentType As String
Public g_BatchSize As Integer
Public g_BatchCounter As Long  ' Track batch number for auto-save

' OnTime Chain State Management
Public Enum BatchProcessState
    bpsIdle = 0
    bpsSetupFormulas = 1
    bpsRefreshing = 2
    bpsProcessingResults = 3
End Enum

Public g_BatchState As BatchProcessState
Public g_BatchStartRow As Long
Public g_BatchEndRow As Long
Public g_FormulaCount As Long
Public g_StopRequested As Boolean
Public g_NextScheduledProc As String
Public g_RefreshCheckCount As Long

' Future Refresh OnTime State
Public g_FutureSheet As Worksheet
Public g_FutureRefreshStartTime As Double
Public g_FutureRefreshTimeout As Long
Public g_FutureUnderlyingCount As Long
Public g_FutureFormulaStartCol As Long  ' First column where RHistory formula is placed
Public g_FutureFormulaStartRow As Long  ' First row where RHistory data starts

' Sheet Names
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_RIC_LIST As String = "RIC_List"  ' Now used for progress tracking
Public Const SHEET_COLLECTION As String = "DataCollection"
Public Const SHEET_QUALITY As String = "QualityReport"
Public Const SHEET_FUTURE As String = "Future et co"

' RANGE FUTURE DOWNLOAD
Public Const RANGE_DOWNLOAD As String = "UnderlyingDownload"  '1st column for 1st underlying. Expand right for more underlyings, +3 columns each
Public Const RANGE_UNDERLYING_START_DATE As String = "UnderlyingStartDate"
Public Const RANGE_UNDERLYING_END_DATE As String = "UnderlyingEndDate"
Public Const RANGE_RFR As String = "RFR"

' Data Limits
Public Const MAX_UNDERLYING_ROWS As Long = 10000  ' Maximum rows for underlying price data and VLOOKUP ranges
Public Const ROW_SPACING As Long = 1000  ' Spacing between formulas in DataCollection sheet
Public Const ENABLE_RETRY_ON_FAILURE As Boolean = False  ' Set to True to retry failed downloads with alternate RIC format

' In-memory batch buffer for direct CSV export
Private g_BatchRows() As Variant
Private g_BatchRowCount As Long
Private Const BATCH_INITIAL_CAPACITY As Long = 10000
Private Const BATCH_COL_COUNT As Long = 29

' Types
Type BatchInfo
    maturityDate As Date
    optionType As String
    strikeStart As Double
    strikeEnd As Double
    Status As String
    RecordsFound As Long
    Errors As Long
End Type

Type OptionData
    spotDate As Date
    premium As Double
    Ticker As String
    maturity As Date
    InterestRate As Double
    spot As Double
    strike As Double
    optionType As String
    impliedVol As Double
    IsValid As Boolean
    ErrorMsg As String
End Type

' ============================================
' Future Sheet Refresh with OnTime (non-blocking)
' ============================================

Sub RefreshFutureSheet()
    ' Entry point - starts the async refresh
    Set g_FutureSheet = ThisWorkbook.Worksheets(SHEET_FUTURE)
    g_FutureRefreshStartTime = Timer
    g_FutureRefreshTimeout = 60

    Application.StatusBar = "Starting LSEG refresh for " & SHEET_FUTURE & "..."

    ' Start the LSEG refresh
    On Error Resume Next
    Application.Run "WorkspaceRefreshWorksheet", True, g_FutureRefreshTimeout * 1000, g_FutureSheet.Name
    On Error GoTo 0

    ' Schedule first check after 3 seconds
    Application.OnTime Now + TimeValue("00:00:03"), "RefreshFutureSheet_CheckReady"
End Sub

Sub RefreshFutureSheet_CheckReady()
    ' Guard: ignore stale OnTime callbacks after hard stop
    If g_FutureSheet Is Nothing Then Exit Sub

    Dim readyCount As Long
    Dim totalCount As Long
    Dim elapsed As Double

    elapsed = Timer - g_FutureRefreshStartTime

    ' Calculate before check
    g_FutureSheet.Calculate
    DoEvents

    ' Check if data is ready
    If IsLSEGDataReady(g_FutureSheet, readyCount, totalCount, 10) Then
        ' Data is ready - complete
        Application.OnTime Now + TimeValue("00:00:02"), "RefreshFutureSheet_Complete"
        Exit Sub
    End If

    ' Check timeout
    If elapsed > g_FutureRefreshTimeout Then
        Application.StatusBar = "Refresh timeout for " & SHEET_FUTURE & " - completing anyway..."
        Application.OnTime Now + TimeValue("00:00:02"), "RefreshFutureSheet_Complete"
        Exit Sub
    End If

    ' Update status and schedule next check
    Application.StatusBar = "Refreshing " & SHEET_FUTURE & "... " & Format(elapsed, "0") & "s - " & _
                           readyCount & " of " & totalCount & " cells ready"

    Application.OnTime Now + TimeValue("00:00:02"), "RefreshFutureSheet_CheckReady"
End Sub

Sub RefreshFutureSheet_Complete()
    ' Guard: ignore stale OnTime callbacks after hard stop
    If g_FutureSheet Is Nothing Then Exit Sub

    g_FutureSheet.Calculate
    DoEvents

    Application.StatusBar = False
    MsgBox "Double check data in : " & SHEET_FUTURE, vbExclamation
End Sub

' ============================================
' Refresh Future Underlyings with OnTime (non-blocking)
' ============================================

Sub RefreshFutureUnderlyings()
    Dim wsRIC As Worksheet
    Dim wsFuture As Worksheet
    Dim wsConfig As Worksheet
    Dim uniqueUnderlyings As Collection
    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    Dim underlyingValue As String
    Dim startCol As Long
    Dim startRow As Long
    Dim currentCol As Long
    Dim underlying As Variant
    Dim formulaRow As Long
    Dim formulaCol As Long
    Dim headerRow As Long
    Dim headerCol As Long
    Dim underlyingsArray() As String
    Dim arraySize As Long
    Dim temp As String
    Dim clearEndCol As Long

    On Error GoTo ErrorHandler

    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)
    Set wsConfig = ThisWorkbook.Worksheets(SHEET_CONFIG)
    Set uniqueUnderlyings = New Collection

    ' Step 1: Extract unique underlyings from RIC_List column G
    Application.StatusBar = "Extracting unique underlyings from RIC_List..."
    lastRow = wsRIC.Cells(wsRIC.Rows.count, "A").End(xlUp).Row

    For i = 2 To lastRow
        underlyingValue = Trim(CStr(wsRIC.Cells(i, 7).Value))  ' Column G

        If underlyingValue <> "" And underlyingValue <> "0" Then
            ' Add to unique collection (will ignore duplicates)
            On Error Resume Next
            uniqueUnderlyings.Add underlyingValue, underlyingValue
            On Error GoTo ErrorHandler
        End If
    Next i

    If uniqueUnderlyings.count = 0 Then
        MsgBox "No underlyings found in RIC_List column G", vbExclamation
        Exit Sub
    End If

    ' Step 2: Convert collection to array and sort alphabetically
    Application.StatusBar = "Sorting underlyings alphabetically..."
    arraySize = uniqueUnderlyings.count
    ReDim underlyingsArray(1 To arraySize)

    i = 1
    For Each underlying In uniqueUnderlyings
        underlyingsArray(i) = CStr(underlying)
        i = i + 1
    Next underlying

    ' Bubble sort by reversed name (e.g. "CCK7" -> "7KCC") so year sorts first, then month code, then future code
    For i = 1 To arraySize - 1
        For j = i + 1 To arraySize
            If StrReverse(underlyingsArray(i)) > StrReverse(underlyingsArray(j)) Then
                temp = underlyingsArray(i)
                underlyingsArray(i) = underlyingsArray(j)
                underlyingsArray(j) = temp
            End If
        Next j
    Next i

    ' Step 2b: Record RICs to RicBloomberg named range (on Config sheet)
    Application.StatusBar = "Recording RICs to RicBloomberg range..."
    Dim ricBloombergRange As Range
    Dim ricDataStartRow As Long
    Dim ricDataEndRow As Long
    Dim ricRowNum As Long

    On Error Resume Next
    Set ricBloombergRange = wsConfig.Range("RicBloomberg")
    On Error GoTo ErrorHandler

    If Not ricBloombergRange Is Nothing Then
        ' Calculate data rows (skip header row 1)
        ricDataStartRow = ricBloombergRange.Row + 1  ' Skip header (N19 + 1 = N20)
        ricDataEndRow = ricBloombergRange.Row + ricBloombergRange.Rows.count - 1  ' N31

        ' Clear previous data in column 1 only (preserve header at row 1 of range)
        wsConfig.Range(wsConfig.Cells(ricDataStartRow, ricBloombergRange.Column), _
                      wsConfig.Cells(ricDataEndRow, ricBloombergRange.Column)).ClearContents

        ' Add sorted RICs to column 1 (column N)
        ricRowNum = ricDataStartRow
        For i = 1 To arraySize
            If ricRowNum > ricDataEndRow Then
                ' Exceeded available space - warn user
                MsgBox "Warning: More RICs (" & arraySize & ") than available space in RicBloomberg range (" & _
                       (ricDataEndRow - ricDataStartRow + 1) & " rows)." & vbNewLine & _
                       "Only the first " & (ricDataEndRow - ricDataStartRow + 1) & " RICs were recorded.", _
                       vbExclamation
                Exit For
            End If

            wsConfig.Cells(ricRowNum, ricBloombergRange.Column).Value = underlyingsArray(i)
            ricRowNum = ricRowNum + 1
        Next i
    Else
        MsgBox "Named range 'RicBloomberg' not found in " & SHEET_CONFIG & vbNewLine & _
               "Skipping RIC recording step.", vbExclamation
    End If

    ' Step 3: Get starting position from RANGE_DOWNLOAD
    On Error Resume Next
    startCol = wsFuture.Range(RANGE_DOWNLOAD).Column
    startRow = wsFuture.Range(RANGE_DOWNLOAD).Row
    On Error GoTo ErrorHandler

    If startCol = 0 Or startRow = 0 Then
        MsgBox "Named range '" & RANGE_DOWNLOAD & "' not found in " & SHEET_FUTURE, vbExclamation
        Exit Sub
    End If

    ' Step 3b: Trim/extend date and formula columns (A-AF) to match date range
    Application.StatusBar = "Adjusting date range in columns A-AF..."
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim numDays As Long
    Dim lastUsedRow As Long
    Dim targetLastRow As Long
    Dim dateFirstRow As Long
    Dim formulaTemplateRow As Long

    dateFirstRow = 11       ' Row 11 = first date (hardcoded value)
    formulaTemplateRow = 12 ' Row 12 = expandable formulas (e.g. =WORKDAY(A11,1))

    ' Workbook may be in manual calc mode — force a recalculation so any
    ' formula-driven UnderlyingStartDate/EndDate cells return current values.
    Application.Calculate

    On Error Resume Next
    dtStart = wsFuture.Range(RANGE_UNDERLYING_START_DATE).Value
    dtEnd = wsFuture.Range(RANGE_UNDERLYING_END_DATE).Value
    On Error GoTo ErrorHandler

    If dtStart > 0 And dtEnd > 0 And dtEnd > dtStart Then
        ' Cap end date at today (no data exists for future dates)
        Dim effectiveEnd As Date
        If dtEnd > Date Then
            effectiveEnd = Date
        Else
            effectiveEnd = dtEnd
        End If

        ' Use NETWORKDAYS for exact business-day count (column A uses WORKDAY)
        numDays = Application.WorksheetFunction.NetworkDays(dtStart, effectiveEnd) + 5

        ' Target: row 11 + numDays rows
        targetLastRow = dateFirstRow + numDays

        ' Set the start date in A11
        wsFuture.Cells(dateFirstRow, 1).Value = dtStart

        ' Find current last used row in column A
        lastUsedRow = wsFuture.Cells(wsFuture.Rows.count, 1).End(xlUp).Row

        ' Clear excess rows beyond target (columns A through AF = 1 through 32)
        If lastUsedRow > targetLastRow Then
            wsFuture.Range(wsFuture.Cells(targetLastRow + 1, 1), _
                          wsFuture.Cells(lastUsedRow, 32)).ClearContents
        End If

        ' Extend formulas from row 12 down to targetLastRow if needed
        If targetLastRow > formulaTemplateRow Then
            ' Fill column A with WORKDAY formula from row 12 down
            wsFuture.Cells(formulaTemplateRow, 1).Formula = "=WORKDAY(A" & (formulaTemplateRow - 1) & ",1)"
            If targetLastRow > formulaTemplateRow Then
                wsFuture.Range(wsFuture.Cells(formulaTemplateRow, 1), _
                              wsFuture.Cells(formulaTemplateRow, 1)).AutoFill _
                    Destination:=wsFuture.Range(wsFuture.Cells(formulaTemplateRow, 1), _
                                               wsFuture.Cells(targetLastRow, 1))
            End If

            ' Fill columns B through AF (2-32) with formulas from row 12 template
            Dim lastTemplateCol As Long
            lastTemplateCol = wsFuture.Cells(formulaTemplateRow, wsFuture.Columns.count).End(xlToLeft).Column
            If lastTemplateCol > 32 Then lastTemplateCol = 32
            If lastTemplateCol >= 2 And targetLastRow > formulaTemplateRow Then
                wsFuture.Range(wsFuture.Cells(formulaTemplateRow, 2), _
                              wsFuture.Cells(formulaTemplateRow, lastTemplateCol)).AutoFill _
                    Destination:=wsFuture.Range(wsFuture.Cells(formulaTemplateRow, 2), _
                                               wsFuture.Cells(targetLastRow, lastTemplateCol))
            End If
        End If

        Application.StatusBar = "Date range set: " & Format(dtStart, "yyyy-mm-dd") & " to " & Format(effectiveEnd, "yyyy-mm-dd") & _
                               " (" & numDays & " business days + margin, rows 11 to " & targetLastRow & ")"

    End If

    ' Step 4: Clear existing underlying data
    Application.StatusBar = "Clearing existing underlying data..."
    ' Find the last used column by scanning for empty blocks.
    ' Each underlying block spans 3 columns starting at (startCol - 1):
    '   col -1: RHistory formula + "Date" header
    '   col  0: underlying ticker name (in startRow) + "Last Price" header
    '   col +1: "Added: ..." metadata
    ' We scan at startRow which holds the ticker name (the named range column).
    clearEndCol = startCol
    Dim scanCol As Long
    scanCol = startCol
    Do While True
        If wsFuture.Cells(startRow, scanCol).Value <> "" Then
            clearEndCol = scanCol + 1  ' This block occupies (scanCol-1, scanCol, scanCol+1)
        End If
        scanCol = scanCol + 3
        ' Stop if we find 2 consecutive empty blocks
        If wsFuture.Cells(startRow, scanCol).Value = "" And _
           wsFuture.Cells(startRow, scanCol + 3).Value = "" Then
            Exit Do
        End If
    Loop

    ' Clear the data area - clear up to MAX_UNDERLYING_ROWS from startRow (not entire sheet).
    ' Start one column LEFT of startCol to include the leftmost formula/header column
    ' of every block (the named range UnderlyingDownload sits in the middle column).
    Dim clearEndRow As Long
    clearEndRow = startRow + MAX_UNDERLYING_ROWS

    wsFuture.Range(wsFuture.Cells(startRow, startCol - 1), wsFuture.Cells(clearEndRow, clearEndCol)).ClearContents

    ' Step 5: Add all underlyings in alphabetical order
    Application.StatusBar = "Adding " & arraySize & " underlyings in alphabetical order..."

    currentCol = startCol

    For i = 1 To arraySize
        ' Calculate formula position
        formulaRow = startRow + 2
        formulaCol = currentCol - 1
        headerRow = startRow + 1
        headerCol = currentCol - 1

        ' Add header
        wsFuture.Cells(headerRow, headerCol).Value = "Date"
        wsFuture.Cells(headerRow, headerCol + 1).Value = "Last Price"

        ' Add the underlying symbol in the RANGE_DOWNLOAD row/column (before formula so cell ref works)
        wsFuture.Cells(startRow, currentCol).Value = underlyingsArray(i)

        ' Add the RHistory formula referencing the header cell for the RIC and named ranges for dates
        Dim ricCellRef As String
        ricCellRef = wsFuture.Cells(startRow, currentCol).Address(False, False)
        wsFuture.Cells(formulaRow, formulaCol).Formula = _
            "=RHistory(" & ricCellRef & "," & _
            """.Timestamp;.Close""," & _
            """START:""&TEXT(" & RANGE_UNDERLYING_START_DATE & ",""YYYY-MM-DD"")&"" END:""&TEXT(" & RANGE_UNDERLYING_END_DATE & ",""YYYY-MM-DD"")&"" INTERVAL:1D"",,""Sort:ASC"")"

        ' Add metadata in the next column
        wsFuture.Cells(startRow, currentCol + 1).Value = "Added: " & Format(Now, "yyyy-mm-dd hh:mm")

        ' Move to next 3-column block
        currentCol = currentCol + 3
    Next i

    ' Step 5b: Clear stale data below row 11 in every rate named range so the
    ' LSEG spill has a clean canvas — prevents #SPILL! errors and stale tail rows.
    ClearRateRangeDataRows wsFuture

    ' Store state for async completion
    Set g_FutureSheet = wsFuture
    g_FutureRefreshStartTime = Timer
    g_FutureRefreshTimeout = 120
    g_FutureUnderlyingCount = arraySize
    g_FutureFormulaStartCol = startCol - 1  ' First RHistory formula column
    g_FutureFormulaStartRow = startRow + 2  ' First RHistory data row

    ' Step 6: Start LSEG refresh (simple pattern matching DownloadFromChain)
    Application.StatusBar = "Starting LSEG refresh for " & arraySize & " underlyings..."

    ' Trigger LSEG refresh
    Application.Run "WorkspaceRefreshWorksheet", True, g_FutureRefreshTimeout * 1000, g_FutureSheet.Name

    ' Schedule check via OnTime (allows LSEG to populate data)
    Application.OnTime Now + TimeValue("00:00:05"), "RefreshFutureUnderlyings_CheckReady"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error in RefreshFutureUnderlyings: " & Err.Description, vbExclamation
End Sub

Sub RefreshFutureUnderlyings_CheckReady()
    ' Guard: ignore stale OnTime callbacks after hard stop
    If g_FutureSheet Is Nothing Then Exit Sub

    Dim readyCount As Long
    Dim totalCount As Long
    Dim elapsed As Double
    Dim i As Long
    Dim checkCol As Long
    Dim cellText As String

    On Error GoTo ErrorHandler

    elapsed = Timer - g_FutureRefreshStartTime
    ' Handle midnight rollover
    If elapsed < 0 Then elapsed = elapsed + 86400

    ' Calculate before check
    g_FutureSheet.Calculate
    DoEvents

    ' Check if data is ready by examining actual RHistory formula columns
    ' RHistory formulas are at columns: g_FutureFormulaStartCol, +3, +6, etc.
    ' Data (close price) is in column + 1 relative to formula column
    readyCount = 0
    totalCount = g_FutureUnderlyingCount

    For i = 1 To g_FutureUnderlyingCount
        ' Calculate the price column for this underlying (formula col + 1)
        checkCol = g_FutureFormulaStartCol + ((i - 1) * 3) + 1

        ' Check the first data row for this underlying
        cellText = CStr(g_FutureSheet.Cells(g_FutureFormulaStartRow, checkCol).Text)

        ' Check if cell is ready (not showing LSEG loading messages)
        If InStr(1, cellText, "Retrieving", vbTextCompare) = 0 And _
           InStr(1, cellText, "Requesting", vbTextCompare) = 0 And _
           InStr(1, cellText, "Loading", vbTextCompare) = 0 And _
           cellText <> "" Then
            readyCount = readyCount + 1
        End If
    Next i

    ' Check if all underlyings are ready
    If readyCount = totalCount Then
        ' Data is ready - schedule completion with delay
        Application.OnTime Now + TimeValue("00:00:02"), "RefreshFutureUnderlyings_Complete"
        Exit Sub
    End If

    ' Check timeout
    If elapsed > g_FutureRefreshTimeout Then
        Application.StatusBar = "Refresh timeout - completing anyway..."
        Application.OnTime Now + TimeValue("00:00:02"), "RefreshFutureUnderlyings_Complete"
        Exit Sub
    End If

    ' Update status and schedule next check
    Application.StatusBar = "Refreshing " & g_FutureUnderlyingCount & " underlyings... " & _
                           Format(elapsed, "0") & "s - " & readyCount & " of " & totalCount & " ready"

    Application.OnTime Now + TimeValue("00:00:02"), "RefreshFutureUnderlyings_CheckReady"
    Exit Sub

ErrorHandler:
    Application.StatusBar = "Error during refresh check: " & Err.Description
    Application.OnTime Now + TimeValue("00:00:02"), "RefreshFutureUnderlyings_Complete"
End Sub

Sub RefreshFutureUnderlyings_Complete()
    ' Guard: ignore stale OnTime callbacks after hard stop
    If g_FutureSheet Is Nothing Then Exit Sub

    On Error Resume Next

    ' First calc: ensure LSEG formulas have settled (col A WORKDAYs, col B-AF
    ' templates, rate-range LSEG cols) before TrimRateRanges reads any row counts.
    g_FutureSheet.Calculate
    DoEvents

    ' Trim rate ranges to match the date range now that LSEG has populated data.
    ' Each named range points to the first LSEG formula cell (row 11).
    ' Columns 1-2 are LSEG-populated; column 3 is an expansion formula.
    TrimRateRanges g_FutureSheet

    ' Second calc: TrimRateRanges autofilled col-3 expansion formulas — those
    ' new cells are dirty but unevaluated under manual calc mode. Recalc so
    ' the sheet shows their computed values, not blank.
    g_FutureSheet.Calculate
    DoEvents

    Application.StatusBar = False

    MsgBox "Added " & g_FutureUnderlyingCount & " underlyings to " & SHEET_FUTURE & " in alphabetical order." & vbNewLine & _
           "Data refresh complete. Please verify the downloaded data." & vbNewLine & _
           "IMPORTANT: Add Bloomberg equivalent for underlying RICs in RicBloomberg range.", vbInformation
End Sub

Sub ClearRateRangeDataRows(wsFuture As Worksheet)
    ' Clear all rows below the template row of each rate named range so the
    ' LSEG spill has a clean canvas to fill. Preserves row 11 (the templates:
    ' LSEG formulas in cols 1-2, expansion formula in col 3).
    Dim rateRangeNames As Variant
    Dim rateRng As Range
    Dim rateCol As Long, rateRow As Long
    Dim col1Last As Long, col2Last As Long, col3Last As Long
    Dim maxLast As Long

    rateRangeNames = Array("RatesUSD", "RatesUSD6m", "RatesUSD1y", "RatesUSD2y", _
                           "RatesEUR", "RatesEUR3m", "RatesEUR6m", "RatesEUR1y", "RatesGBP")

    Application.StatusBar = "Clearing rate ranges below template row..."

    Dim rn As Variant
    For Each rn In rateRangeNames
        Set rateRng = Nothing
        On Error Resume Next
        Set rateRng = wsFuture.Range(CStr(rn))
        On Error GoTo 0
        If rateRng Is Nothing Then GoTo NextRange

        rateCol = rateRng.Column
        rateRow = rateRng.Row

        ' Find the deepest used row across the 3 columns of this range
        col1Last = wsFuture.Cells(wsFuture.Rows.count, rateCol).End(xlUp).Row
        col2Last = wsFuture.Cells(wsFuture.Rows.count, rateCol + 1).End(xlUp).Row
        col3Last = wsFuture.Cells(wsFuture.Rows.count, rateCol + 2).End(xlUp).Row
        maxLast = col1Last
        If col2Last > maxLast Then maxLast = col2Last
        If col3Last > maxLast Then maxLast = col3Last

        ' Clear from rateRow + 1 (one row below the template) down to the deepest used row
        If maxLast > rateRow Then
            wsFuture.Range(wsFuture.Cells(rateRow + 1, rateCol), _
                          wsFuture.Cells(maxLast, rateCol + 2)).ClearContents
        End If
NextRange:
    Next rn
End Sub

Sub TrimRateRanges(wsFuture As Worksheet)
    ' For each rate named range (3-column block: LSEG date, LSEG value, expansion formula),
    ' expand col 3 to cover the col-A date axis (targetLastRow) by default. Only trim
    ' back to LSEG's actual data extent when LSEG clearly returned fewer rows than the
    ' date range — never under-extend just because LSEG hasn't fully populated yet.
    '
    ' The "template" cell for col 3's expansion formula sits at the SAME row as the
    ' named range itself (row 11). Autofill is sourced from that row.
    Dim rateRangeNames As Variant
    Dim rateRng As Range
    Dim rateCol As Long
    Dim rateRow As Long
    Dim lsegCol1Last As Long
    Dim lsegCol2Last As Long
    Dim lsegLastRow As Long
    Dim col3CurrentLast As Long
    Dim col3EndRow As Long
    Dim targetLastRow As Long
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim effectiveEnd As Date

    ' Compute targetLastRow from the date axis named ranges (same as Step 3b)
    On Error Resume Next
    dtStart = wsFuture.Range(RANGE_UNDERLYING_START_DATE).Value
    dtEnd = wsFuture.Range(RANGE_UNDERLYING_END_DATE).Value
    On Error GoTo 0

    If dtStart = 0 Or dtEnd = 0 Or dtEnd <= dtStart Then Exit Sub

    If dtEnd > Date Then
        effectiveEnd = Date
    Else
        effectiveEnd = dtEnd
    End If

    targetLastRow = 11 + Application.WorksheetFunction.NetworkDays(dtStart, effectiveEnd) + 5

    rateRangeNames = Array("RatesUSD", "RatesUSD6m", "RatesUSD1y", "RatesUSD2y", _
                           "RatesEUR", "RatesEUR3m", "RatesEUR6m", "RatesEUR1y", "RatesGBP")

    Application.StatusBar = "Aligning rate-range expansion formulas..."

    Dim rn As Variant
    For Each rn In rateRangeNames
        Set rateRng = Nothing
        On Error Resume Next
        Set rateRng = wsFuture.Range(CStr(rn))
        On Error GoTo 0

        If rateRng Is Nothing Then GoTo NextRange

        rateCol = rateRng.Column
        rateRow = rateRng.Row    ' Template row (= row of the named range, e.g., row 11)

        ' Inspect LSEG's data extent across both data columns; spill may show
        ' on either, so take the larger.
        lsegCol1Last = wsFuture.Cells(wsFuture.Rows.count, rateCol).End(xlUp).Row
        lsegCol2Last = wsFuture.Cells(wsFuture.Rows.count, rateCol + 1).End(xlUp).Row
        If lsegCol1Last > lsegCol2Last Then
            lsegLastRow = lsegCol1Last
        Else
            lsegLastRow = lsegCol2Last
        End If

        ' Decide where col 3 should end:
        '   - Default: targetLastRow (col A axis) so the 3rd column always
        '     visually matches the date range.
        '   - Trim back only when LSEG clearly returned fewer rows: requires
        '     lsegLastRow > rateRow + 5 (a sanity guard so a not-yet-populated
        '     LSEG range doesn't truncate the formula).
        If lsegLastRow > rateRow + 5 And lsegLastRow < targetLastRow Then
            col3EndRow = lsegLastRow
        Else
            col3EndRow = targetLastRow
        End If

        ' Clear col 3 beyond col3EndRow (whatever a previous run left behind).
        col3CurrentLast = wsFuture.Cells(wsFuture.Rows.count, rateCol + 2).End(xlUp).Row
        If col3CurrentLast > col3EndRow Then
            wsFuture.Range(wsFuture.Cells(col3EndRow + 1, rateCol + 2), _
                          wsFuture.Cells(col3CurrentLast, rateCol + 2)).ClearContents
        End If

        ' Extend col 3's expansion formula from the template row (rateRow) down
        ' to col3EndRow. Autofill source is the named-range row itself.
        If col3EndRow > rateRow And _
           Not IsEmpty(wsFuture.Cells(rateRow, rateCol + 2).Value) Then
            wsFuture.Cells(rateRow, rateCol + 2).AutoFill _
                Destination:=wsFuture.Range(wsFuture.Cells(rateRow, rateCol + 2), _
                                           wsFuture.Cells(col3EndRow, rateCol + 2))
        End If
NextRange:
    Next rn
End Sub



' ============================================
' MODULE 2: Main Process Controller
' ============================================

Sub InitializeWorkbook()
    ' Create necessary sheets if they don't exist
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim i As Integer
    
    sheetNames = Array(SHEET_CONFIG, SHEET_RIC_LIST, SHEET_COLLECTION, SHEET_QUALITY)
    
    For i = 0 To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheetNames(i))
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = sheetNames(i)
        End If
        On Error GoTo 0
    Next i
    
    ' Setup headers
    SetupRICListSheet  ' Setup RIC_List with all needed columns
    SetupQualitySheet
End Sub

Sub MainDownloadProcess()
    Dim response As Integer
    Dim totalRICs As Long

    Application.StatusBar = "Initializing workbook..."
    ' Initialize
    InitializeWorkbook

    Application.StatusBar = "Loading configuration..."
    ' Load configuration
    If Not LoadConfiguration() Then
        MsgBox "Please complete configuration in Config sheet", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

    Application.StatusBar = "Checking RIC list..."
    ' Check if RIC_List has data
    If Not CheckRICListExists() Then
        MsgBox "Please run GenerateAllRICs first to create the RIC list!", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

    ' Check that all RICs have underlying ticker filled in
    Dim missingUnderlyingRow As Long
    If Not CheckAllUnderlyingsFilled(missingUnderlyingRow) Then
        MsgBox "Missing underlying ticker in RIC_List!" & vbNewLine & vbNewLine & _
               "Row " & missingUnderlyingRow & " (and possibly others) is missing the underlying ticker in column G." & vbNewLine & _
               "Please fill in all underlying tickers before running the download.", _
               vbExclamation, "Missing Underlying Ticker"
        Application.StatusBar = False
        Exit Sub
    End If

    ' Check that all underlyings exist in Future et co with actual data
    Application.StatusBar = "Checking underlyings have data in Future et co..."
    Dim missingUnderlyingData As String
    If Not CheckUnderlyingsHaveData(missingUnderlyingData) Then
        MsgBox "Missing underlying data in 'Future et co' sheet!" & vbNewLine & vbNewLine & _
               "The following underlyings are missing or have no data:" & vbNewLine & _
               missingUnderlyingData & vbNewLine & vbNewLine & _
               "Please download the underlying futures data first.", _
               vbExclamation, "Missing Underlying Data"
        Application.StatusBar = False
        Exit Sub
    End If

    ' Verify the date axis in column A of Future et co covers the option
    ' download range — RFR and spot VLOOKUPs depend on these dates.
    Application.StatusBar = "Checking date range in Future et co column A..."
    Dim dateRangeError As String
    If Not CheckUnderlyingDateRange(dateRangeError) Then
        MsgBox "Date range mismatch in 'Future et co' column A:" & vbNewLine & vbNewLine & _
               dateRangeError & vbNewLine & vbNewLine & _
               "Run RefreshFutureUnderlyings with the correct UnderlyingStartDate / UnderlyingEndDate first.", _
               vbExclamation, "Underlying Date Range Mismatch"
        Application.StatusBar = False
        Exit Sub
    End If

    totalRICs = CountUnprocessedRICs()

    ' Show summary
    response = MsgBox("Starting download for: " & g_RootRIC & vbNewLine & _
                     "Date Range: " & g_DateStart & " to " & g_DateEnd & vbNewLine & _
                     "RICs to process: " & totalRICs & vbNewLine & _
                     vbNewLine & "Continue?", vbYesNo + vbQuestion)

    If response = vbNo Then
        Application.StatusBar = False
        Exit Sub
    End If

    Application.StatusBar = "Checking underlying data availability..."
    ' Check Underlying data
    If Not CheckUnderlyings() Then
        MsgBox "Process stopped: Missing underlying futures data." & vbNewLine & _
               "Please download the missing underlyings first before processing options.", _
               vbExclamation, "Missing Underlyings"
        Application.StatusBar = False
        Exit Sub
    End If

    Application.StatusBar = "Starting batch processing for " & totalRICs & " RICs..."

    ' Start batch processing chain
    ' NOTE: This triggers async OnTime chain - quality report generated at end by ProcessBatch_Complete
    ProcessAllBatchesFromRICList

    ' VBA execution ends here - OnTime chain runs asynchronously
End Sub

' ============================================
' Batch Processing with OnTime Chain Architecture
' ============================================
' This uses Application.OnTime to break VBA execution between phases,
' allowing LSEG add-in to populate data asynchronously.
'
' Flow: ProcessAllBatches ? SetupFormulas ? [OnTime] ? CheckRefresh
'       ? ProcessResults ? TriggerNext ? [loop back to SetupFormulas]
'
' To stop processing: Run StopBatchProcessing() or press ESC during dialogs
' ============================================

Sub ProcessAllBatchesFromRICList()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim batchStart As Long
    Dim batchEnd As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Initialize
    g_BatchCounter = 0
    g_StopRequested = False
    g_BatchState = bpsSetupFormulas

    ' Reset any rows stuck in "Processing" from a prior interrupted run
    Dim resetRow As Long
    For resetRow = 2 To lastRow
        If ws.Cells(resetRow, 9).Value = "Processing" Then
            ws.Cells(resetRow, 9).Value = "No"
        End If
    Next resetRow

    ' Find first unprocessed batch
    batchStart = FindNextUnprocessedRIC(2)
    g_BatchCounter = (batchStart - 2) / g_BatchSize
    If batchStart = 0 Then
        MsgBox "No unprocessed RICs found!", vbInformation
        Exit Sub
    End If

    batchEnd = Application.Min(batchStart + g_BatchSize - 1, lastRow)

    ' Increment counter
    g_BatchCounter = g_BatchCounter + 1

    ' Save batch range and trigger chain
    g_BatchStartRow = batchStart
    g_BatchEndRow = batchEnd

    ' Mark as processing
    MarkBatchStatus batchStart, batchEnd, "Processing"

    ' Start the chain
    ProcessBatch_SetupFormulas
End Sub

' ============================================
' PHASE 1: Setup Formulas and Trigger Refresh
' ============================================
Sub ProcessBatch_SetupFormulas()
    ' Guard: ignore stale OnTime callbacks after hard stop
    If g_BatchState = bpsIdle Then Exit Sub

    Dim wsRIC As Worksheet
    Dim wsCollection As Worksheet
    Dim i As Long
    Dim ric As String
    Dim currentRow As Long

    ' Check stop flag
    If g_StopRequested Then
        ProcessBatch_Abort
        Exit Sub
    End If

    g_BatchState = bpsSetupFormulas
    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    Set wsCollection = ThisWorkbook.Worksheets(SHEET_COLLECTION)

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Clearing collection sheet..."
    ClearCollectionSheet

    ' Setup formulas
    currentRow = 2
    g_FormulaCount = 0

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Setting up formulas..."

    For i = g_BatchStartRow To g_BatchEndRow
        ric = wsRIC.Cells(i, 1).Value

        ' Skip if already processed
        If wsRIC.Cells(i, 9).Value = "Yes" Then GoTo NextRIC  ' Column I: Processed

        ' Update status
        If g_FormulaCount Mod 10 = 0 Then
            Application.StatusBar = "Batch #" & g_BatchCounter & ": Preparing RIC " & (g_FormulaCount + 1) & " - " & ric
        End If

        currentRow = 2 + (g_FormulaCount * ROW_SPACING)

        ' Setup formula
        wsCollection.Cells(currentRow, 1).Formula = BuildRHistoryFormula(ric, g_DateStart, g_DateEnd)

        ' Setup MINIMAL anchor metadata only (full metadata deferred to Phase 3 after LSEG refresh)
        ' This avoids setting up 1000 rows of metadata for RICs that may return no data
        SetupMinimalAnchorMetadata wsCollection, currentRow, _
                                   wsRIC.Cells(i, 3).Value, _
                                   wsRIC.Cells(i, 4).Value, _
                                   wsRIC.Cells(i, 2).Value, _
                                   i, wsRIC.Cells(i, 7).Value, ric

        g_FormulaCount = g_FormulaCount + 1

NextRIC:
    Next i

    ' Only proceed if there's data
    If g_FormulaCount > 0 Then
        Application.StatusBar = "Batch #" & g_BatchCounter & ": Refreshing LSEG data for " & g_FormulaCount & " RICs..."
        g_BatchState = bpsRefreshing
        g_RefreshCheckCount = 0

        ' Trigger LSEG refresh
        RefreshLSEGCollectionSheet

        ' Schedule check after 5 seconds
        g_NextScheduledProc = "ProcessBatch_CheckRefresh"
        Application.OnTime Now + TimeValue("00:00:05"), g_NextScheduledProc
    Else
        ' No formulas, skip to next batch
        ProcessBatch_TriggerNext
    End If
End Sub

' ============================================
' PHASE 2: Check if Refresh Complete
' ============================================
Sub ProcessBatch_CheckRefresh()
    ' Guard: ignore stale OnTime callbacks after hard stop
    If g_BatchState = bpsIdle Then Exit Sub

    Dim wsCollection As Worksheet
    Dim readyCount As Long
    Dim totalChecks As Long

    ' Check stop flag
    If g_StopRequested Then
        ProcessBatch_Abort
        Exit Sub
    End If

    ' Check timeout
    g_RefreshCheckCount = g_RefreshCheckCount + 1
    If g_RefreshCheckCount > 60 Then  ' 60 checks × 3 sec = 3 min timeout
        MsgBox "LSEG refresh timeout for batch #" & g_BatchCounter & " - proceeding anyway", vbExclamation
        ' Schedule via OnTime to let LSEG finish populating all columns
        Application.OnTime Now + TimeValue("00:00:02"), "ProcessBatch_ProcessResults"
        Exit Sub
    End If

    Set wsCollection = ThisWorkbook.Worksheets(SHEET_COLLECTION)
    wsCollection.Calculate

    ' Check if data ready (with progress tracking)
    If IsDataReady(wsCollection, readyCount, totalChecks) Then
        ' Data ready, schedule processing via OnTime to let LSEG finish populating all columns
        Application.StatusBar = "Batch #" & g_BatchCounter & ": All data ready (" & readyCount & "/" & totalChecks & " cells) - processing in 2 seconds..."
        Application.OnTime Now + TimeValue("00:00:02"), "ProcessBatch_ProcessResults"
    Else
        ' Still waiting, reschedule - show progress in status bar
        Application.StatusBar = "Batch #" & g_BatchCounter & ": Waiting for LSEG data... " & _
                               readyCount & " of " & totalChecks & " cells ready (check #" & g_RefreshCheckCount & ")"
        g_NextScheduledProc = "ProcessBatch_CheckRefresh"
        Application.OnTime Now + TimeValue("00:00:03"), g_NextScheduledProc
    End If
End Sub

' ============================================
' PHASE 3: Process Results
' ============================================
Sub ProcessBatch_ProcessResults()
    ' Guard: ignore stale OnTime callbacks after hard stop
    If g_BatchState = bpsIdle Then Exit Sub

    Dim wsCollection As Worksheet
    Dim i As Long
    Dim processRow As Long

    ' Check stop flag
    If g_StopRequested Then
        ProcessBatch_Abort
        Exit Sub
    End If

    g_BatchState = bpsProcessingResults
    Set wsCollection = ThisWorkbook.Worksheets(SHEET_COLLECTION)

    ' Reset the in-memory batch buffer for this batch's CSV
    ResetBatchBuffer

    ' Setup metadata and Greek formulas ONLY for rows with actual LSEG data
    ' This is the optimized "lazy initialization" approach - metadata was deferred from Phase 1
    Application.StatusBar = "Batch #" & g_BatchCounter & ": Setting up metadata for data rows..."
    For i = 0 To g_FormulaCount - 1
        processRow = 2 + (i * ROW_SPACING)
        SetupMetadataForDataRows wsCollection, processRow, ROW_SPACING
    Next i

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Adding Greek formulas to data rows..."
    For i = 0 To g_FormulaCount - 1
        processRow = 2 + (i * ROW_SPACING)
        AddGreekFormulasToDataRows wsCollection, processRow, ROW_SPACING
    Next i

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Calculating Greeks..."
    wsCollection.Calculate
    ' Wait for calculation to complete
    Dim calcTimeout As Long
    calcTimeout = 0
    Do While Application.CalculationState <> xlDone
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")  ' Wait 1 second
        calcTimeout = calcTimeout + 1
        
       If calcTimeout > 20 Then  ' 20 second timeout
           If AreGreeksComputed(wsCollection, g_FormulaCount, ROW_SPACING) Then
               Exit Do  ' Greeks are done, Excel just doesn't know it
           Else
               MsgBox "Calculation timeout - Greeks not fully computed, proceeding anyway", vbExclamation
               Exit Do
           End If
       End If
    Loop

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Validating and buffering data..."
    ValidateAndUpdateRICListWithSpacing wsCollection, g_FormulaCount

    ' Retry failed RICs with alternate format (toggle expired suffix)
    If ENABLE_RETRY_ON_FAILURE Then
        Application.StatusBar = "Batch #" & g_BatchCounter & ": Retrying failed RICs..."
        RetryFailedRICsInBatch wsCollection, g_BatchStartRow, g_BatchEndRow
    End If

    'Application.StatusBar = "Batch #" & g_BatchCounter & ": Final calculations..."
    'Application.Calculate

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Saving to CSV..."
    WriteBatchToCSV GetBatchNumberFromRow(g_BatchStartRow)

    ' Save workbook every 5 batches
    If g_BatchCounter Mod 5 = 0 Then
        Application.StatusBar = "Saving workbook (batch " & g_BatchCounter & ")..."
        ThisWorkbook.Save
    End If

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Complete!"
    ShowBatchSummaryFromRICList g_BatchStartRow, g_BatchEndRow

    ' Trigger next batch
    ProcessBatch_TriggerNext
End Sub

' ============================================
' PHASE 4: Trigger Next Batch or Complete
' ============================================
Sub ProcessBatch_TriggerNext()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim nextStart As Long
    Dim nextEnd As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Find next batch
    nextStart = FindNextUnprocessedRIC(g_BatchEndRow + 1)

    If nextStart > 0 And nextStart <= lastRow Then
        ' Found next batch
        nextEnd = Application.Min(nextStart + g_BatchSize - 1, lastRow)
        g_BatchCounter = g_BatchCounter + 1

        ' Save new batch range
        g_BatchStartRow = nextStart
        g_BatchEndRow = nextEnd

        MarkBatchStatus nextStart, nextEnd, "Processing"

        ' Schedule next batch
        g_NextScheduledProc = "ProcessBatch_SetupFormulas"
        Application.OnTime Now + TimeValue("00:00:02"), g_NextScheduledProc
    Else
        ' All done
        ProcessBatch_Complete
    End If
End Sub

' ============================================
' Completion Handler
' ============================================
Sub ProcessBatch_Complete()
    g_BatchState = bpsIdle
    Application.StatusBar = "All batches complete! Generating quality report..."

    GenerateQualityReport

    Application.StatusBar = False
    MsgBox "All batches processed! Check Quality Report for summary.", vbInformation
End Sub

' ============================================
' Helper Functions for OnTime Chain
' ============================================

' Check if LSEG data has loaded (with optional progress tracking)
Function IsDataReady(ws As Worksheet, Optional ByRef outReadyCount As Long, Optional ByRef outTotalChecks As Long) As Boolean
    Dim checkRow As Long
    Dim readyCount As Long
    Dim totalChecks As Long
    Dim cellValue As Variant
    Dim cellText As String
    Dim i As Long
    Dim samplePositions() As Long
    Dim maxSamples As Long

    totalChecks = 0
    readyCount = 0
    maxSamples = Application.Min(10, g_FormulaCount)  ' Increased from 5 to 10 samples

    ' Build array of positions to check (first, last, and evenly distributed middle positions)
    ReDim samplePositions(1 To maxSamples)

    If g_FormulaCount > 0 Then
        For i = 1 To maxSamples
            If i = 1 Then
                ' First formula
                samplePositions(i) = 2
            ElseIf i = maxSamples And g_FormulaCount > 1 Then
                ' Last formula
                samplePositions(i) = 2 + ((g_FormulaCount - 1) * ROW_SPACING)
            Else
                ' Evenly distributed middle positions
                samplePositions(i) = 2 + (((i - 1) * g_FormulaCount \ maxSamples) * ROW_SPACING)
            End If
        Next i

        ' Check sampled positions
        For i = 1 To maxSamples
            checkRow = samplePositions(i)
            totalChecks = totalChecks + 1

            cellValue = ws.Cells(checkRow, 2).Value
            cellText = CStr(ws.Cells(checkRow, 2).Text)

            ' Check if cell is ready (no longer shows LSEG status messages)
            If InStr(1, cellText, "Retrieving", vbTextCompare) = 0 And _
               InStr(1, cellText, "Requesting", vbTextCompare) = 0 And _
               InStr(1, cellText, "Loading", vbTextCompare) = 0 Then
                readyCount = readyCount + 1
            End If
        Next i
    End If

    ' Return progress info via optional ByRef parameters
    outReadyCount = readyCount
    outTotalChecks = totalChecks

    ' Consider ready if ALL checked cells are no longer refreshing
    IsDataReady = (totalChecks > 0 And readyCount = totalChecks)
End Function

' Stop batch processing (graceful - waits for current phase)
Sub StopBatchProcessing()
    g_StopRequested = True
    Application.StatusBar = "Stop requested - will halt after current operation..."
    MsgBox "Batch processing will stop after current phase completes.", vbInformation
End Sub

' Emergency stop - cancels ALL pending OnTime callbacks and resets state
' Use this after a hard stop (Escape + Stop in VBA editor) to prevent ghost execution
Sub EmergencyStop()
    On Error Resume Next  ' OnTime errors if no matching call is pending

    ' Cancel batch processing chain callbacks
    If g_NextScheduledProc <> "" Then
        Application.OnTime EarliestTime:=Now, Procedure:=g_NextScheduledProc, Schedule:=False
    End If
    Application.OnTime EarliestTime:=Now, Procedure:="ProcessBatch_SetupFormulas", Schedule:=False
    Application.OnTime EarliestTime:=Now, Procedure:="ProcessBatch_CheckRefresh", Schedule:=False
    Application.OnTime EarliestTime:=Now, Procedure:="ProcessBatch_ProcessResults", Schedule:=False

    ' Cancel future sheet refresh callbacks
    Application.OnTime EarliestTime:=Now, Procedure:="RefreshFutureSheet_CheckReady", Schedule:=False
    Application.OnTime EarliestTime:=Now, Procedure:="RefreshFutureSheet_Complete", Schedule:=False
    Application.OnTime EarliestTime:=Now, Procedure:="RefreshFutureUnderlyings_CheckReady", Schedule:=False
    Application.OnTime EarliestTime:=Now, Procedure:="RefreshFutureUnderlyings_Complete", Schedule:=False

    ' Cancel DownloadFromChain callbacks
    Application.OnTime EarliestTime:=Now, Procedure:="DownloadFromChain_CheckChainReady", Schedule:=False
    Application.OnTime EarliestTime:=Now, Procedure:="DownloadFromChain_ProcessNextBatch", Schedule:=False
    Application.OnTime EarliestTime:=Now, Procedure:="DownloadFromChain_CheckBatchReady", Schedule:=False
    Application.OnTime EarliestTime:=Now, Procedure:="DownloadFromChain_Complete", Schedule:=False
    g_ChainState = CHAIN_STATE_IDLE
    g_ChainStopRequested = False

    On Error GoTo 0

    ' Reset all state
    g_BatchState = bpsIdle
    g_StopRequested = False
    g_NextScheduledProc = ""
    Application.StatusBar = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "All pending operations cancelled and state reset.", vbInformation
End Sub

' Abort handler
Sub ProcessBatch_Abort()
    g_BatchState = bpsIdle
    g_StopRequested = False
    g_NextScheduledProc = ""
    Application.StatusBar = False
    MsgBox "Processing stopped at batch #" & g_BatchCounter & vbNewLine & _
           "Progress saved in RIC_List sheet.", vbInformation
End Sub

' ============================================
' Helper Functions for Formula Building
' ============================================

' Find the column number for a specific underlying in the Future sheet
Function FindUnderlyingColumn(underlyingTicker As String) As Long
    Dim wsFuture As Worksheet
    Dim startCol As Long
    Dim startRow As Long
    Dim currentCol As Long
    Dim foundUnderlying As String

    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)

    ' Get starting position from RANGE_DOWNLOAD
    On Error Resume Next
    startCol = wsFuture.Range(RANGE_DOWNLOAD).Column
    startRow = wsFuture.Range(RANGE_DOWNLOAD).Row
    On Error GoTo 0

    FindUnderlyingColumn = 0  ' Default: not found

    If startCol > 0 And startRow > 0 Then
        ' Scan columns in steps of 3 (each underlying uses 3 columns)
        currentCol = startCol
        Do While True
            foundUnderlying = Trim(CStr(wsFuture.Cells(startRow, currentCol).Value))

            If foundUnderlying = underlyingTicker Then
                FindUnderlyingColumn = currentCol
                Exit Function
            End If

            currentCol = currentCol + 3  ' Jump to next underlying block

            ' Exit if we find empty blocks (no more underlyings)
            If wsFuture.Cells(startRow, currentCol).Value = "" And _
               wsFuture.Cells(startRow, currentCol + 3).Value = "" Then
                Exit Do
            End If
        Loop
    End If
End Function

' Build VLOOKUP formula for underlying spot price
Function BuildSpotVLOOKUPFormula(rowNum As Long, underlyingTicker As String) As String
    Dim wsFuture As Worksheet
    Dim startRow As Long
    Dim underlyingCol As Long

    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)

    ' Get starting row from RANGE_DOWNLOAD
    On Error Resume Next
    startRow = wsFuture.Range(RANGE_DOWNLOAD).Row
    On Error GoTo 0

    ' Find the column for this underlying
    underlyingCol = FindUnderlyingColumn(underlyingTicker)

    If underlyingCol > 0 And startRow > 0 Then
        ' Build VLOOKUP formula: lookup date in column A, return price from underlying's column
        ' Range: Date column (underlyingCol - 1) and Price column (underlyingCol)
        BuildSpotVLOOKUPFormula = "=IFERROR(VLOOKUP(A" & rowNum & ",'" & SHEET_FUTURE & "'!" & _
            wsFuture.Cells(startRow + 2, underlyingCol - 1).Address(False, False) & ":" & _
            wsFuture.Cells(startRow + MAX_UNDERLYING_ROWS, underlyingCol).Address(False, False) & _
            ",2,TRUE),"""")"
    Else
        ' If column not found, return empty formula
        BuildSpotVLOOKUPFormula = ""
    End If
End Function

' Build VLOOKUP formula for underlying spot price in R1C1 notation (for range-based operations)
' Returns formula that can be applied to entire column range at once
Function BuildSpotVLOOKUPFormulaR1C1(underlyingTicker As String) As String
    Dim wsFuture As Worksheet
    Dim startRow As Long
    Dim underlyingCol As Long
    Dim lookupStartRow As Long
    Dim lookupEndRow As Long
    Dim dateCol As Long

    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)

    ' Get starting row from RANGE_DOWNLOAD
    On Error Resume Next
    startRow = wsFuture.Range(RANGE_DOWNLOAD).Row
    On Error GoTo 0

    ' Find the column for this underlying
    underlyingCol = FindUnderlyingColumn(underlyingTicker)

    If underlyingCol > 0 And startRow > 0 Then
        ' Calculate row/column references for R1C1 formula
        lookupStartRow = startRow + 2
        lookupEndRow = startRow + MAX_UNDERLYING_ROWS
        dateCol = underlyingCol - 1

        ' Build R1C1 VLOOKUP formula: lookup date from column 1, return price from underlying's column
        ' R1C1 format: VLOOKUP(RC1, 'Future et co'!R{startRow}C{dateCol}:R{endRow}C{underlyingCol}, 2, TRUE)
        BuildSpotVLOOKUPFormulaR1C1 = "=IFERROR(VLOOKUP(RC1,'" & SHEET_FUTURE & "'!R" & lookupStartRow & "C" & dateCol & ":R" & lookupEndRow & "C" & underlyingCol & ",2,TRUE),"""")"
    Else
        ' If column not found, return empty formula
        BuildSpotVLOOKUPFormulaR1C1 = ""
    End If
End Function

' ============================================
' OPTIMIZED: Setup minimal metadata at anchor row only (before LSEG refresh)
' Full metadata is deferred to SetupMetadataForDataRows after LSEG returns data
' ============================================
Sub SetupMinimalAnchorMetadata(ws As Worksheet, startRow As Long, _
                               strike As Double, optType As String, maturity As Date, _
                               ricRowRef As Long, underlyingTicker As String, optionRic As String)
    ' Store ONLY what's needed to process this block later
    ' These values are read by SetupMetadataForDataRows after LSEG refresh
    ws.Cells(startRow, 15).Value = ricRowRef         ' RIC_Row_Ref - needed for validation
    ws.Cells(startRow, 16).Value = optionRic         ' RIC - needed for tracking
    ws.Cells(startRow, 7).Value = strike             ' Strike - temp storage for deferred setup
    ws.Cells(startRow, 8).Value = Left(optType, 1)   ' Type - temp storage
    ws.Cells(startRow, 4).Value = maturity           ' Maturity - temp storage
    ws.Cells(startRow, 29).Value = underlyingTicker  ' Underlying - temp storage
End Sub

' ============================================
' LEGACY: Original function that sets up ALL rows before knowing if data exists
' Kept for reference/rollback - not used in optimized flow
' ============================================
Sub SetupRHistoryAndMetadata_LEGACY(ws As Worksheet, startRow As Long, maxRows As Long, _
                             strike As Double, optType As String, maturity As Date, ricRowRef As Long, underlyingTicker As String, optionRic As String)
    Dim i As Long
    Dim endRow As Long
    Dim wsFuture As Worksheet
    Dim underlyingCol As Long
    Dim rfrRange As Range
    Dim rfrRow As Long
    Dim rfrCol As Long
    Dim rfrLastRow As Long
    Dim spotFormula As String
    Dim optFreq As String
    Dim weekNum As Integer
    Dim optionBloomTicker As String
    Dim underlyingBloomTicker As String

    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)

    ' Find the underlying column using helper function
    underlyingCol = FindUnderlyingColumn(underlyingTicker)

    ' Get RFR range position and find last row with data
    On Error Resume Next
    Set rfrRange = wsFuture.Range(RANGE_RFR)
    On Error GoTo 0

    If Not rfrRange Is Nothing Then
        rfrRow = rfrRange.Row
        rfrCol = rfrRange.Column
        rfrLastRow = wsFuture.Cells(wsFuture.Rows.count, 1).End(xlUp).Row
    End If

    ' Get option frequency from Config
    optFreq = GetOptionFrequency()

    ' Get underlying Bloomberg ticker for building option ticker
    underlyingBloomTicker = GetBloombergTicker(underlyingTicker, ricRowRef)

    ' Build option Bloomberg ticker using simpler approach with already-available data
    If optFreq = "weekly" Then
        weekNum = GetWeekNumberFromDate(maturity)
        optionBloomTicker = BuildWeeklyOptionBloombergTicker(underlyingBloomTicker, optType, strike, weekNum)
    Else
        optionBloomTicker = BuildOptionBloombergTicker(underlyingBloomTicker, optType, strike)
    End If

    endRow = startRow + maxRows - 1

    ' Setup basic metadata and VLOOKUP formulas (NO Greek formulas yet)
    For i = startRow To endRow
        ' Store metadata - Column C: Option Bloomberg ticker
        ws.Cells(i, 3).Value = optionBloomTicker
        ws.Cells(i, 4).Value = maturity

        ' Column E: Interest_rate - VLOOKUP from RFR range with dynamic last row
        If Not rfrRange Is Nothing Then
            ws.Cells(i, 5).Formula = "=IFERROR(VLOOKUP(A" & i & ",'" & SHEET_FUTURE & "'!" & _
                wsFuture.Range(wsFuture.Cells(rfrRow, 1), wsFuture.Cells(rfrLastRow, rfrCol)).Address(False, False) & _
                "," & rfrCol & ",TRUE),""not found"")"
        Else
            ws.Cells(i, 5).Value = "not found"
        End If

        ' Column F: Spot - VLOOKUP from underlying data using helper function
        spotFormula = BuildSpotVLOOKUPFormula(i, underlyingTicker)
        If spotFormula <> "" Then
            ws.Cells(i, 6).Formula = spotFormula
        Else
            ' Fallback if underlying column not found
            ws.Cells(i, 6).Value = GetSpotPrice(underlyingTicker)
        End If

        ' Store additional metadata
        ws.Cells(i, 7).Value = strike
        ws.Cells(i, 8).Value = Left(optType, 1)
        ws.Cells(i, 15).Value = ricRowRef
        ws.Cells(i, 16).Value = optionRic
        ws.Cells(i, 17).Value = g_LotSize
        ws.Cells(i, 18).Value = g_NamePrefix & " " & Left(optType, 1) & " " & strike & " " & Format(maturity, "mmm-yyyy")
        ws.Cells(i, 19).Value = GetBloombergTicker(underlyingTicker, ricRowRef)
        ws.Cells(i, 20).Value = g_Currency
        ws.Cells(i, 21).Value = 0
        ws.Cells(i, 29).Value = underlyingTicker  ' RIC_Underlying (column AC)
    Next i
End Sub

' Add Greek formulas ONLY to rows with premium data (after LSEG refresh)
Sub AddGreekFormulasToDataRows(ws As Worksheet, startRow As Long, maxRows As Long)
    Dim i As Long
    Dim endRow As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim rowCount As Long
    Dim formulaRange As Range
    Dim originalCalcMode As XlCalculation

    ' Save original calculation mode
    originalCalcMode = Application.Calculation

    ' Disable screen updating and calculation for speed
    'Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    'Application.EnableEvents = False

    On Error GoTo CleanUp

    endRow = startRow + maxRows - 1
    firstDataRow = 0
    lastDataRow = 0

    ' Find first and last rows with premium data (skip errors like #N/A)
    For i = startRow To endRow
        If Not IsEmpty(ws.Cells(i, 1).Value) And Not IsError(ws.Cells(i, 1).Value) And _
           Not IsEmpty(ws.Cells(i, 2).Value) And Not IsError(ws.Cells(i, 2).Value) Then
            If firstDataRow = 0 Then firstDataRow = i
            lastDataRow = i
        ElseIf firstDataRow > 0 Then
            Exit For ' No more data
        End If
    Next i

    ' Exit if no data found
    If firstDataRow = 0 Or lastDataRow = 0 Then GoTo CleanUp

    rowCount = lastDataRow - firstDataRow + 1

    ' Use R1C1 notation for relative formulas - much faster!
    ' Add Greek formulas to the data range only

    ' Column I (9): Implied Volatility (handle #N/A errors and calc errors)
    ws.Range(ws.Cells(firstDataRow, 9), ws.Cells(lastDataRow, 9)).FormulaR1C1 = _
        "=IF(OR(RC[-7]="""",ISERROR(RC[-7])),"""",IFERROR(GBlackScholesImpVolBisection(LOWER(RC[-1]),RC[-3],RC[-2],(RC[-5]-RC[-8])/365,RC[-4],0,RC[-7]),""NA""))"

    ' Column J (10): Delta
    ws.Range(ws.Cells(firstDataRow, 10), ws.Cells(lastDataRow, 10)).FormulaR1C1 = _
        "=IF(OR(RC[-8]="""",ISERROR(RC[-8])),"""",IFERROR(GBlackScholesNGreeks(""d"",LOWER(RC[-2]),RC[-4],RC[-3],(RC[-6]-RC[-9])/365,RC[-5],0,RC[-1]),""NA""))"

    ' Column K (11): Vega
    ws.Range(ws.Cells(firstDataRow, 11), ws.Cells(lastDataRow, 11)).FormulaR1C1 = _
        "=IF(OR(RC[-9]="""",ISERROR(RC[-9])),"""",IFERROR(GBlackScholesNGreeks(""v"",LOWER(RC[-3]),RC[-5],RC[-4],(RC[-7]-RC[-10])/365,RC[-6],0,RC[-2]),""NA""))"

    ' Column L (12): Gamma
    ws.Range(ws.Cells(firstDataRow, 12), ws.Cells(lastDataRow, 12)).FormulaR1C1 = _
        "=IF(OR(RC[-10]="""",ISERROR(RC[-10])),"""",IFERROR(GBlackScholesNGreeks(""g"",LOWER(RC[-4]),RC[-6],RC[-5],(RC[-8]-RC[-11])/365,RC[-7],0,RC[-3]),""NA""))"

    ' Column M (13): Theta
    ws.Range(ws.Cells(firstDataRow, 13), ws.Cells(lastDataRow, 13)).FormulaR1C1 = _
        "=IF(OR(RC[-11]="""",ISERROR(RC[-11])),"""",IFERROR(GBlackScholesNGreeks(""t"",LOWER(RC[-5]),RC[-7],RC[-6],(RC[-9]-RC[-12])/365,RC[-8],0,RC[-4]),""NA""))"

    ' Column N (14): Rho
    ws.Range(ws.Cells(firstDataRow, 14), ws.Cells(lastDataRow, 14)).FormulaR1C1 = _
        "=IF(OR(RC[-12]="""",ISERROR(RC[-12])),"""",IFERROR(GBlackScholesNGreeks(""r"",LOWER(RC[-6]),RC[-8],RC[-7],(RC[-10]-RC[-13])/365,RC[-9],0,RC[-5]),""NA""))"

    ' Column V (22): DDELTA/DVOL
    ws.Range(ws.Cells(firstDataRow, 22), ws.Cells(lastDataRow, 22)).FormulaR1C1 = _
        "=IF(OR(RC[-20]="""",ISERROR(RC[-20])),"""",IFERROR(CGBlackScholes(""dddv"",LOWER(RC[-14]),RC[-16],RC[-15],(RC[-18]-RC[-21])/365,RC[-17],0,RC[-13],RC[-12]),""NA""))"

    ' Column W (23): DDELTA/DVOLDVOL
    ws.Range(ws.Cells(firstDataRow, 23), ws.Cells(lastDataRow, 23)).FormulaR1C1 = _
        "=IF(OR(RC[-21]="""",ISERROR(RC[-21])),"""",IFERROR(CGBlackScholes(""dvv"",LOWER(RC[-15]),RC[-17],RC[-16],(RC[-19]-RC[-22])/365,RC[-18],0,RC[-14],RC[-13]),""NA""))"

    ' Column X (24): Charm (DDELTA/DTIME)
    ws.Range(ws.Cells(firstDataRow, 24), ws.Cells(lastDataRow, 24)).FormulaR1C1 = _
        "=IF(OR(RC[-22]="""",ISERROR(RC[-22])),"""",IFERROR(CGBlackScholes(""dt"",LOWER(RC[-16]),RC[-18],RC[-17],(RC[-20]-RC[-23])/365,RC[-19],0,RC[-15],RC[-14]),""NA""))"

    ' Column Y (25): DGamma/DSpot
    ws.Range(ws.Cells(firstDataRow, 25), ws.Cells(lastDataRow, 25)).FormulaR1C1 = _
        "=IF(OR(RC[-23]="""",ISERROR(RC[-23])),"""",IFERROR(CGBlackScholes(""gps"",LOWER(RC[-17]),RC[-19],RC[-18],(RC[-21]-RC[-24])/365,RC[-20],0,RC[-16],RC[-15]),""NA""))"

    ' Column Z (26): Zomma (DGAMMA/DVOL)
    ws.Range(ws.Cells(firstDataRow, 26), ws.Cells(lastDataRow, 26)).FormulaR1C1 = _
        "=IF(OR(RC[-24]="""",ISERROR(RC[-24])),"""",IFERROR(CGBlackScholes(""gpv"",LOWER(RC[-18]),RC[-20],RC[-19],(RC[-22]-RC[-25])/365,RC[-21],0,RC[-17],RC[-16]),""NA""))"

    ' Column AA (27): Vomma (DVEGA/DVOL)
    ws.Range(ws.Cells(firstDataRow, 27), ws.Cells(lastDataRow, 27)).FormulaR1C1 = _
        "=IF(OR(RC[-25]="""",ISERROR(RC[-25])),"""",IFERROR(CGBlackScholes(""dvdv"",LOWER(RC[-19]),RC[-21],RC[-20],(RC[-23]-RC[-26])/365,RC[-22],0,RC[-18],RC[-17]),""NA""))"

    ' Column AB (28): Ultima (DVEGA/DVOLDVOL)
    ws.Range(ws.Cells(firstDataRow, 28), ws.Cells(lastDataRow, 28)).FormulaR1C1 = _
        "=IF(OR(RC[-26]="""",ISERROR(RC[-26])),"""",IFERROR(CGBlackScholes(""vvv"",LOWER(RC[-20]),RC[-22],RC[-21],(RC[-24]-RC[-27])/365,RC[-23],0,RC[-19],RC[-18]),""NA""))"

    ' Calculate the worksheet to populate formulas
    'ws.Calculate

CleanUp:
    ' Re-enable Excel features and restore original calculation mode
    'Application.ScreenUpdating = True
    Application.Calculation = originalCalcMode
    'Application.EnableEvents = True
End Sub

' ============================================
' Check if all Greek formulas have computed values
' Reads entire range into memory in one COM call, then checks in-memory
' Returns True if all data rows with premium input have evaluated Greeks
' ============================================
Private Function AreGreeksComputed(ws As Worksheet, formulaCount As Long, rowSpacing As Long) As Boolean
    Dim lastRow As Long
    Dim data As Variant
    Dim i As Long, r As Long
    Dim blockStartRow As Long
    Dim endRow As Long
    Dim cellVal As Variant
    Dim col As Long

    ' Greek columns to check (relative to array: col 2=1, col 9=8, col 10=9, etc.)
    ' Array columns: 1=B(premium), 8=I(IV), 9=J(Delta), 10=K(Vega), 11=L(Gamma), 12=M(Theta), 13=N(Rho)
    '               21=V(dDelta/dVol), 22=W, 23=X(Charm), 24=Y, 25=Z(Zomma), 26=AA(Vomma), 27=AB(Ultima)
    Dim greekCols As Variant
    greekCols = Array(8, 9, 10, 11, 12, 13, 21, 22, 23, 24, 25, 26, 27)

    If formulaCount = 0 Then
        AreGreeksComputed = True
        Exit Function
    End If

    ' Calculate the extent of data: from row 2 to last possible data row
    lastRow = 2 + (formulaCount - 1) * rowSpacing + rowSpacing - 1

    ' Read cols B through AB (2 through 28) into array in one shot
    data = ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 28)).Value

    ' Check each formula block
    For i = 0 To formulaCount - 1
        blockStartRow = 1 + (i * rowSpacing)  ' Array is 1-based, block starts at row offset
        endRow = blockStartRow + rowSpacing - 1
        If endRow > UBound(data, 1) Then endRow = UBound(data, 1)

        ' Scan data rows within this block
        For r = blockStartRow To endRow
            ' Column 1 in array = col B(2) = premium
            If Not IsEmpty(data(r, 1)) And Not IsError(data(r, 1)) Then
                ' This row has input data — check all Greek columns
                Dim g As Long
                For g = LBound(greekCols) To UBound(greekCols)
                    col = greekCols(g)
                    If col > UBound(data, 2) Then GoTo NextGreek

                    cellVal = data(r, col)

                    ' A computed cell is: numeric, empty string, or "NA"
                    ' An uncomputed cell would be: Error, or still showing formula
                    If IsError(cellVal) Then
                        AreGreeksComputed = False
                        Exit Function
                    End If

                    If Not IsNumeric(cellVal) And cellVal <> "" And cellVal <> "NA" Then
                        AreGreeksComputed = False
                        Exit Function
                    End If
NextGreek:
                Next g
            End If
        Next r
    Next i

    AreGreeksComputed = True
End Function

' ============================================
' OPTIMIZED: Setup metadata ONLY for rows with actual LSEG data (after refresh)
' Phase 2: Uses range-based operations instead of cell-by-cell for ~99% faster performance
' ============================================
Sub SetupMetadataForDataRows(ws As Worksheet, startRow As Long, maxRows As Long)
    Dim i As Long
    Dim endRow As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim wsFuture As Worksheet
    Dim underlyingCol As Long
    Dim rfrRange As Range
    Dim rfrRow As Long
    Dim rfrCol As Long
    Dim rfrLastRow As Long
    Dim spotFormulaR1C1 As String
    Dim optFreq As String
    Dim weekNum As Integer
    Dim optionBloomTicker As String
    Dim underlyingBloomTicker As String
    Dim originalCalcMode As XlCalculation
    Dim nameString As String
    Dim optTypeShort As String
    Dim rfrFormulaR1C1 As String

    ' Read anchor row values (stored by SetupMinimalAnchorMetadata)
    Dim strike As Double
    Dim optType As String
    Dim maturity As Date
    Dim ricRowRef As Long
    Dim underlyingTicker As String
    Dim optionRic As String

    ' Save original calculation mode and disable for speed
    originalCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    endRow = startRow + maxRows - 1
    firstDataRow = 0
    lastDataRow = 0

    ' Find first and last rows with actual LSEG data (same pattern as AddGreekFormulasToDataRows)
    For i = startRow To endRow
        If Not IsEmpty(ws.Cells(i, 1).Value) And Not IsError(ws.Cells(i, 1).Value) And _
           Not IsEmpty(ws.Cells(i, 2).Value) And Not IsError(ws.Cells(i, 2).Value) Then
            If firstDataRow = 0 Then firstDataRow = i
            lastDataRow = i
        ElseIf firstDataRow > 0 Then
            Exit For ' No more data
        End If
    Next i

    ' Exit if no data found - nothing to set up
    If firstDataRow = 0 Or lastDataRow = 0 Then GoTo CleanUp

    ' Read anchor row metadata (stored by SetupMinimalAnchorMetadata)
    strike = ws.Cells(startRow, 7).Value
    optType = ws.Cells(startRow, 8).Value
    maturity = ws.Cells(startRow, 4).Value
    ricRowRef = ws.Cells(startRow, 15).Value
    optionRic = ws.Cells(startRow, 16).Value
    underlyingTicker = ws.Cells(startRow, 29).Value

    ' Setup references for VLOOKUP formulas
    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)
    underlyingCol = FindUnderlyingColumn(underlyingTicker)

    ' Get RFR range position
    On Error Resume Next
    Set rfrRange = wsFuture.Range(RANGE_RFR)
    On Error GoTo CleanUp

    If Not rfrRange Is Nothing Then
        rfrRow = rfrRange.Row
        rfrCol = rfrRange.Column
        rfrLastRow = wsFuture.Cells(wsFuture.Rows.count, 1).End(xlUp).Row
    End If

    ' Get option frequency and build Bloomberg ticker
    optFreq = GetOptionFrequency()
    underlyingBloomTicker = GetBloombergTicker(underlyingTicker, ricRowRef)

    If optFreq = "weekly" Then
        weekNum = GetWeekNumberFromDate(maturity)
        optionBloomTicker = BuildWeeklyOptionBloombergTicker(underlyingBloomTicker, optType, strike, weekNum)
    Else
        optionBloomTicker = BuildOptionBloombergTicker(underlyingBloomTicker, optType, strike)
    End If

    ' Prepare common values
    optTypeShort = Left(optType, 1)
    nameString = g_NamePrefix & " " & optTypeShort & " " & strike & " " & Format(maturity, "mmm-yyyy")

    ' ============================================
    ' RANGE-BASED OPERATIONS: Fill entire columns at once
    ' ============================================

    ' Column C (3): Option Bloomberg ticker - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 3), ws.Cells(lastDataRow, 3)).Value = optionBloomTicker

    ' Column D (4): Maturity - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 4), ws.Cells(lastDataRow, 4)).Value = maturity

    ' Column E (5): Interest_rate - VLOOKUP using R1C1 notation for entire range
    If Not rfrRange Is Nothing Then
        ' R1C1 format: VLOOKUP(RC1, 'Future et co'!R{rfrRow}C1:R{rfrLastRow}C{rfrCol}, {rfrCol}, TRUE)
        rfrFormulaR1C1 = "=IFERROR(VLOOKUP(RC1,'" & SHEET_FUTURE & "'!R" & rfrRow & "C1:R" & rfrLastRow & "C" & rfrCol & "," & rfrCol & ",TRUE),""not found"")"
        ws.Range(ws.Cells(firstDataRow, 5), ws.Cells(lastDataRow, 5)).FormulaR1C1 = rfrFormulaR1C1
    Else
        ws.Range(ws.Cells(firstDataRow, 5), ws.Cells(lastDataRow, 5)).Value = "not found"
    End If

    ' Column F (6): Spot - VLOOKUP using R1C1 notation for entire range
    spotFormulaR1C1 = BuildSpotVLOOKUPFormulaR1C1(underlyingTicker)
    If spotFormulaR1C1 <> "" Then
        ws.Range(ws.Cells(firstDataRow, 6), ws.Cells(lastDataRow, 6)).FormulaR1C1 = spotFormulaR1C1
    Else
        ws.Range(ws.Cells(firstDataRow, 6), ws.Cells(lastDataRow, 6)).Value = GetSpotPrice(underlyingTicker)
    End If

    ' Column G (7): Strike - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 7), ws.Cells(lastDataRow, 7)).Value = strike

    ' Column H (8): Type - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 8), ws.Cells(lastDataRow, 8)).Value = optTypeShort

    ' Column O (15): RIC row reference - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 15), ws.Cells(lastDataRow, 15)).Value = ricRowRef

    ' Column P (16): Option RIC - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 16), ws.Cells(lastDataRow, 16)).Value = optionRic

    ' Column Q (17): Lot size - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 17), ws.Cells(lastDataRow, 17)).Value = g_LotSize

    ' Column R (18): Name - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 18), ws.Cells(lastDataRow, 18)).Value = nameString

    ' Column S (19): Underlying Bloomberg ticker - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 19), ws.Cells(lastDataRow, 19)).Value = underlyingBloomTicker

    ' Column T (20): Currency - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 20), ws.Cells(lastDataRow, 20)).Value = g_Currency

    ' Column U (21): Placeholder (0) - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 21), ws.Cells(lastDataRow, 21)).Value = 0

    ' Column AC (29): Underlying RIC - same value for all rows
    ws.Range(ws.Cells(firstDataRow, 29), ws.Cells(lastDataRow, 29)).Value = underlyingTicker

CleanUp:
    Application.Calculation = originalCalcMode
End Sub

' ============================================
' OPTIMIZED: Copy data rows to staging using array transfer
' Phase 2: Uses array-based operations instead of cell-by-cell for ~99% faster performance
' ============================================
' Build a remapped destination array from a DataCollection block and append
' it to the in-memory batch buffer (g_BatchRows) for the current batch's CSV.
Sub BuildAndAppendBatchRows(ws As Worksheet, startRow As Long, maxRows As Long)
    Dim firstDataRow As Long, lastDataRow As Long
    Dim srcData As Variant
    Dim destData() As Variant
    Dim rowCount As Long, i As Long, destIdx As Long
    Dim internalIdValue As String
    Dim endRow As Long
    Dim cellVal As Variant

    ' Get Internal_ID once
    On Error Resume Next
    internalIdValue = Trim(ThisWorkbook.Sheets(SHEET_CONFIG).Range("internalId").Value)
    On Error GoTo 0

    endRow = startRow + maxRows - 1
    firstDataRow = 0
    lastDataRow = 0

    ' Find first and last rows with valid premium data (not empty, not error, numeric)
    For i = startRow To endRow
        cellVal = ws.Cells(i, 2).Value
        If Not IsEmpty(ws.Cells(i, 1).Value) And Not IsError(ws.Cells(i, 1).Value) And _
           Not IsEmpty(cellVal) And Not IsError(cellVal) And IsNumeric(cellVal) Then
            If firstDataRow = 0 Then firstDataRow = i
            lastDataRow = i
        ElseIf firstDataRow > 0 And IsEmpty(ws.Cells(i, 1).Value) Then
            Exit For ' No more data in this section
        End If
    Next i

    If firstDataRow = 0 Or lastDataRow = 0 Then Exit Sub

    rowCount = lastDataRow - firstDataRow + 1

    ' Read source data into array (single operation - read columns 1-29)
    srcData = ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(lastDataRow, 29)).Value2

    ' Prepare destination array with column remapping
    ' First pass: count valid rows (skip any rows with errors in the middle)
    Dim validRowCount As Long
    validRowCount = 0
    For i = 1 To rowCount
        If Not IsEmpty(srcData(i, 1)) And Not IsError(srcData(i, 1)) And _
           Not IsEmpty(srcData(i, 2)) And Not IsError(srcData(i, 2)) And IsNumeric(srcData(i, 2)) Then
            validRowCount = validRowCount + 1
        End If
    Next i

    If validRowCount = 0 Then Exit Sub

    ReDim destData(1 To validRowCount, 1 To BATCH_COL_COUNT)

    ' Second pass: populate destination array with column remapping
    destIdx = 0
    For i = 1 To rowCount
        If Not IsEmpty(srcData(i, 1)) And Not IsError(srcData(i, 1)) And _
           Not IsEmpty(srcData(i, 2)) And Not IsError(srcData(i, 2)) And IsNumeric(srcData(i, 2)) Then
            destIdx = destIdx + 1
            ' Spot_Date / Maturity: .Value2 returns Doubles (date serials like 46234)
            ' for date-formatted cells. Cast to Date so FormatCSVField writes them
            ' as "yyyy-mm-dd hh:mm:ss" rather than "46234".
            If IsNumeric(srcData(i, 1)) Then
                destData(destIdx, 1) = CDate(srcData(i, 1))    ' Spot_Date (col 1)
            Else
                destData(destIdx, 1) = srcData(i, 1)
            End If
            destData(destIdx, 2) = srcData(i, 2)    ' Premium (col 2)
            destData(destIdx, 3) = srcData(i, 3)    ' Ticker (col 3)
            If IsNumeric(srcData(i, 4)) Then
                destData(destIdx, 4) = CDate(srcData(i, 4))    ' Maturity (col 4)
            Else
                destData(destIdx, 4) = srcData(i, 4)
            End If
            destData(destIdx, 5) = srcData(i, 5)    ' Interest_rate (col 5)
            destData(destIdx, 6) = srcData(i, 6)    ' Spot (col 6)
            destData(destIdx, 7) = srcData(i, 7)    ' Strike (col 7)
            destData(destIdx, 8) = srcData(i, 8)    ' Type (col 8)
            destData(destIdx, 9) = srcData(i, 9)    ' IV (col 9)
            destData(destIdx, 10) = srcData(i, 10)  ' Delta (col 10)
            destData(destIdx, 11) = srcData(i, 11)  ' Vega (col 11)
            destData(destIdx, 12) = srcData(i, 12)  ' Gamma (col 12)
            destData(destIdx, 13) = srcData(i, 13)  ' Theta (col 13)
            destData(destIdx, 14) = srcData(i, 14)  ' Rho (col 14)
            destData(destIdx, 15) = srcData(i, 17)  ' Lot_size (col 17 -> 15)
            destData(destIdx, 16) = srcData(i, 18)  ' Name (col 18 -> 16)
            destData(destIdx, 17) = srcData(i, 19)  ' Reference (col 19 -> 17)
            destData(destIdx, 18) = srcData(i, 20)  ' ccy_pair (col 20 -> 18)
            destData(destIdx, 19) = internalIdValue ' Internal_ID (from Config)
            destData(destIdx, 20) = srcData(i, 21)  ' Dividend (col 21 -> 20)
            destData(destIdx, 21) = srcData(i, 22)  ' DDELTA_DVOL (col 22 -> 21)
            destData(destIdx, 22) = srcData(i, 23)  ' DDELTA_DVOLDVOL (col 23 -> 22)
            destData(destIdx, 23) = srcData(i, 24)  ' DDELTA_DTIME (col 24 -> 23)
            destData(destIdx, 24) = srcData(i, 25)  ' DGAMMA_DSPOT (col 25 -> 24)
            destData(destIdx, 25) = srcData(i, 26)  ' DGAMMA_DVOL (col 26 -> 25)
            destData(destIdx, 26) = srcData(i, 27)  ' DVEGA_DVOL (col 27 -> 26)
            destData(destIdx, 27) = srcData(i, 28)  ' DVEGA_DVOLDVOL (col 28 -> 27)
            destData(destIdx, 28) = srcData(i, 16)  ' RIC (col 16 -> 28)
            destData(destIdx, 29) = srcData(i, 29)  ' RIC_Underlying (col 29)
        End If
    Next i

    AppendRowsToBatch destData, validRowCount
End Sub

' Reset the in-memory batch buffer at the start of every batch.
Sub ResetBatchBuffer()
    ReDim g_BatchRows(1 To BATCH_INITIAL_CAPACITY, 1 To BATCH_COL_COUNT)
    g_BatchRowCount = 0
End Sub

' Append a remapped block to the in-memory batch buffer, growing capacity
' (doubling) as needed via ReDim Preserve. srcData is 1-based, 29 cols.
Sub AppendRowsToBatch(srcData As Variant, srcRowCount As Long)
    If srcRowCount <= 0 Then Exit Sub

    ' Determine current capacity (UBound errors if array was never ReDim'd)
    Dim currentCapacity As Long
    currentCapacity = 0
    On Error Resume Next
    currentCapacity = UBound(g_BatchRows, 1)
    On Error GoTo 0

    If currentCapacity = 0 Then
        ReDim g_BatchRows(1 To BATCH_INITIAL_CAPACITY, 1 To BATCH_COL_COUNT)
        currentCapacity = BATCH_INITIAL_CAPACITY
    End If

    ' Grow if needed (double until it fits)
    Do While g_BatchRowCount + srcRowCount > currentCapacity
        currentCapacity = currentCapacity * 2
    Loop
    If currentCapacity > UBound(g_BatchRows, 1) Then
        ReDim Preserve g_BatchRows(1 To currentCapacity, 1 To BATCH_COL_COUNT)
    End If

    Dim i As Long, c As Long
    For i = 1 To srcRowCount
        For c = 1 To BATCH_COL_COUNT
            g_BatchRows(g_BatchRowCount + i, c) = srcData(i, c)
        Next c
    Next i
    g_BatchRowCount = g_BatchRowCount + srcRowCount
End Sub

' Headers for the batch CSV (29 columns, in SQL-friendly order).
Function GetBatchCSVHeaders() As Variant
    Dim h(1 To 1, 1 To BATCH_COL_COUNT) As Variant
    h(1, 1) = "Spot_Date"
    h(1, 2) = "Premium"
    h(1, 3) = "Ticker"
    h(1, 4) = "Maturity"
    h(1, 5) = "Interest_rate"
    h(1, 6) = "Spot"
    h(1, 7) = "Strike"
    h(1, 8) = "Type"
    h(1, 9) = "Implied_Volatility"
    h(1, 10) = "Delta"
    h(1, 11) = "Vega"
    h(1, 12) = "Gamma"
    h(1, 13) = "Theta"
    h(1, 14) = "Rho"
    h(1, 15) = "Lot_size"
    h(1, 16) = "Name"
    h(1, 17) = "Reference"
    h(1, 18) = "ccy_pair"
    h(1, 19) = "Internal_ID"
    h(1, 20) = "Dividend"
    h(1, 21) = "DDELTA/DVOL"
    h(1, 22) = "DDELTA/DVOLDVOL"
    h(1, 23) = "DDELTA/DTIME"
    h(1, 24) = "DGAMMA/DSPOT"
    h(1, 25) = "DGAMMA/DVOL"
    h(1, 26) = "DVEGA/DVOL"
    h(1, 27) = "DVEGA/DVOLDVOL"
    h(1, 28) = "RIC"
    h(1, 29) = "RIC_Underlying"
    GetBatchCSVHeaders = h
End Function

' Write the in-memory batch buffer to a timestamped CSV via WriteArrayToCSV.
Sub WriteBatchToCSV(Optional batchNumber As Long = 0)
    If g_BatchRowCount <= 0 Then Exit Sub

    Dim csvPath As String, fileName As String

    If batchNumber > 0 Then
        fileName = g_RootRIC & "_" & Format(Date, "yyyymmdd") & "_" & Format(Now, "HHmmss") & "_batch" & batchNumber & ".csv"
    Else
        fileName = g_RootRIC & "_" & Format(Date, "yyyymmdd") & "_" & Format(Now, "HHmmss") & ".csv"
    End If
    csvPath = ThisWorkbook.Path & "\" & fileName

    ' Slice g_BatchRows down to actual row count before writing
    Dim sliced() As Variant
    ReDim sliced(1 To g_BatchRowCount, 1 To BATCH_COL_COUNT)
    Dim r As Long, c As Long
    For r = 1 To g_BatchRowCount
        For c = 1 To BATCH_COL_COUNT
            sliced(r, c) = g_BatchRows(r, c)
        Next c
    Next r

    WriteArrayToCSV sliced, csvPath, GetBatchCSVHeaders()

    Application.StatusBar = "Auto-saved: " & fileName & " (" & g_BatchRowCount & " rows)"
End Sub

' Modified validation function to handle spacing
Sub ValidateAndUpdateRICListWithSpacing(wsCollection As Worksheet, formulaCount As Long)
    Dim wsRIC As Worksheet
    Dim i As Long
    Dim formulaRow As Long
    Dim ricRow As Long
    Dim premium As Variant
    Dim iv As Variant
    Dim delta As Variant
    Dim validationResult As String
    Dim lastPremium As Double
    Dim lastIV As Double
    Dim dataFound As Boolean

    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)

    ' Caller (ProcessBatch_ProcessResults) already runs Calculate and waits
    ' for xlDone before invoking us — no need to recalculate here.

    For i = 0 To formulaCount - 1
        formulaRow = 2 + (i * ROW_SPACING)
        
        ' Get the RIC row reference
        ricRow = wsCollection.Cells(formulaRow, 15).Value
        
        If ricRow > 0 Then
            dataFound = False
            lastPremium = 0
            lastIV = 0
            
            ' Check all rows in this formula's space for the latest data
            Dim checkRow As Long
            For checkRow = formulaRow To formulaRow + ROW_SPACING - 1
                If IsEmpty(wsCollection.Cells(checkRow, 1).Value) Then
                    Exit For  ' No more data in this section
                End If
                
                premium = wsCollection.Cells(checkRow, 2).Value
                ' Check for error first (VBA doesn't short-circuit And)
                If Not IsEmpty(premium) And Not IsError(premium) Then
                    If IsNumeric(premium) Then
                        If premium > 0 Then
                            dataFound = True
                            lastPremium = premium

                            ' Get IV if available
                            iv = wsCollection.Cells(checkRow, 9).Value
                            If Not IsError(iv) And IsNumeric(iv) Then
                                lastIV = iv
                            End If

                            ' Get Delta if available
                            delta = wsCollection.Cells(checkRow, 10).Value
                        End If
                    End If
                End If
            Next checkRow
            
            ' Update RIC_List with results
            If dataFound Then
                ' Successful download
                wsRIC.Cells(ricRow, 9).Value = "Yes"  ' Processed (column I)
                wsRIC.Cells(ricRow, 10).Value = Now     ' Process_Time (column J)
                wsRIC.Cells(ricRow, 11).Value = lastPremium ' Premium (column K)

                If lastIV > 0 Then
                    wsRIC.Cells(ricRow, 12).Value = lastIV  ' IV (column L)
                    validationResult = ValidateIV(lastIV, wsRIC.Cells(ricRow, 3).Value, _
                                                 GetSpotPrice(wsRIC.Cells(ricRow, 7).Value), wsRIC.Cells(ricRow, 2).Value)
                    wsRIC.Cells(ricRow, 14).Value = validationResult  ' Validation (column N)
                End If

                If Not IsError(delta) And IsNumeric(delta) Then
                    wsRIC.Cells(ricRow, 13).Value = delta  ' Delta (column M)
                End If

                ' Copy data to staging - export all data that was successfully downloaded
                ' (validation is for quality info, not a gate for export)
                ' Find the last row with data in this section
                Dim lastDataRow As Long
                lastDataRow = formulaRow
                Dim findRow As Long
                For findRow = formulaRow To formulaRow + ROW_SPACING - 1
                    Dim cellVal As Variant
                    cellVal = wsCollection.Cells(findRow, 2).Value
                    If Not IsEmpty(cellVal) And Not IsError(cellVal) Then
                        lastDataRow = findRow
                    Else
                        Exit For
                    End If
                Next findRow
                ' Copy the rows with data to staging
                BuildAndAppendBatchRows wsCollection, formulaRow, lastDataRow - formulaRow + 1
            Else
                ' Failed download
                wsRIC.Cells(ricRow, 9).Value = "Error"  ' Processed (column I)
                wsRIC.Cells(ricRow, 15).Value = "No data returned"  ' Error_Message (column O)
            End If
        End If
    Next i
End Sub


' ============================================
' MODULE 3: RIC List Management Functions
' ============================================

Sub SetupRICListSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_RIC_LIST
    End If
    
    ' Check if headers already exist (from GenerateAllRICs)
    If ws.Range("A1").Value <> "RIC" Then
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
        End With
    End If

    ' Add additional tracking columns if they don't exist
    With ws
        If .Range("J1").Value = "" Then .Range("J1").Value = "Process_Time"
        If .Range("K1").Value = "" Then .Range("K1").Value = "Premium"
        If .Range("L1").Value = "" Then .Range("L1").Value = "IV"
        If .Range("M1").Value = "" Then .Range("M1").Value = "Delta"
        If .Range("N1").Value = "" Then .Range("N1").Value = "Validation"
        If .Range("O1").Value = "" Then .Range("O1").Value = "Error_Message"

        ' Format headers
        .Range("A1:O1").Font.Bold = True
        .Range("A1:O1").Interior.Color = RGB(200, 200, 200)

        ' Add conditional formatting to Processed column (I)
        Dim lastRow As Long
        lastRow = .Cells(.Rows.count, "A").End(xlUp).Row
        If lastRow > 1 Then
            With .Range("I2:I" & lastRow).FormatConditions
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
        End If
    End With

    ' AutoFit columns
    ws.Columns("A:O").AutoFit
End Sub

Function CheckRICListExists() As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    On Error GoTo 0

    If ws Is Nothing Then
        CheckRICListExists = False
    Else
        ' Check if there's data beyond header
        CheckRICListExists = ws.Cells(ws.Rows.count, "A").End(xlUp).Row > 1
    End If
End Function

Function CheckAllUnderlyingsFilled(ByRef missingRow As Long) As Boolean
    ' Checks that all rows in RIC_List have an underlying ticker in column G
    ' Returns True if all filled, False if any are missing
    ' missingRow returns the first row with missing underlying (0 if all filled)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    missingRow = 0

    For i = 2 To lastRow
        If Trim(ws.Cells(i, 7).Value) = "" Then  ' Column G: Underlying LSEG
            missingRow = i
            CheckAllUnderlyingsFilled = False
            Exit Function
        End If
    Next i

    CheckAllUnderlyingsFilled = True
End Function

Function CheckUnderlyingsHaveData(ByRef missingList As String) As Boolean
    ' Checks that all unique underlyings from RIC_List exist in Future et co with actual data
    ' Returns True if all have data, False if any are missing
    ' missingList returns comma-separated list of underlyings without data
    Dim wsRIC As Worksheet
    Dim wsFuture As Worksheet
    Dim uniqueUnderlyings As Collection
    Dim i As Long
    Dim lastRow As Long
    Dim underlyingValue As String
    Dim startCol As Long
    Dim startRow As Long
    Dim underlyingCol As Long
    Dim dataCount As Long

    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)
    Set uniqueUnderlyings = New Collection

    missingList = ""

    ' Step 1: Extract unique underlyings from RIC_List column G
    lastRow = wsRIC.Cells(wsRIC.Rows.count, "A").End(xlUp).Row

    For i = 2 To lastRow
        underlyingValue = Trim(CStr(wsRIC.Cells(i, 7).Value))
        If underlyingValue <> "" Then
            On Error Resume Next
            uniqueUnderlyings.Add underlyingValue, underlyingValue
            On Error GoTo 0
        End If
    Next i

    If uniqueUnderlyings.count = 0 Then
        CheckUnderlyingsHaveData = True
        Exit Function
    End If

    ' Step 2: Get starting position from RANGE_DOWNLOAD
    On Error Resume Next
    startCol = wsFuture.Range(RANGE_DOWNLOAD).Column
    startRow = wsFuture.Range(RANGE_DOWNLOAD).Row
    On Error GoTo 0

    If startCol = 0 Or startRow = 0 Then
        missingList = "Cannot find RANGE_DOWNLOAD in Future et co"
        CheckUnderlyingsHaveData = False
        Exit Function
    End If

    ' Step 3: Check each unique underlying exists and has data
    Dim underlying As Variant
    Dim foundCol As Long
    Dim currentCol As Long
    Dim foundUnderlying As String

    For Each underlying In uniqueUnderlyings
        foundCol = 0

        ' Search for underlying in Future et co (every 3rd column)
        currentCol = startCol
        Do While True
            foundUnderlying = Trim(CStr(wsFuture.Cells(startRow, currentCol).Value))

            If foundUnderlying = CStr(underlying) Then
                foundCol = currentCol
                Exit Do
            End If

            currentCol = currentCol + 3

            ' Exit if empty blocks
            If wsFuture.Cells(startRow, currentCol).Value = "" And _
               wsFuture.Cells(startRow, currentCol + 3).Value = "" Then
                Exit Do
            End If
        Loop

        ' Check if found and has data
        If foundCol = 0 Then
            ' Underlying not found in Future et co
            If missingList <> "" Then missingList = missingList & ", "
            missingList = missingList & CStr(underlying) & " (not found)"
        Else
            ' Found - check if there's actual price data (check a few rows below header)
            dataCount = 0
            For i = startRow + 2 To startRow + 10
                If IsNumeric(wsFuture.Cells(i, foundCol).Value) And _
                   wsFuture.Cells(i, foundCol).Value <> 0 Then
                    dataCount = dataCount + 1
                End If
            Next i

            If dataCount = 0 Then
                If missingList <> "" Then missingList = missingList & ", "
                missingList = missingList & CStr(underlying) & " (no data)"
            End If
        End If
    Next underlying

    CheckUnderlyingsHaveData = (missingList = "")
End Function

' Verify that column A of Future et co covers [g_DateStart, min(g_DateEnd, today)].
' RFR and spot VLOOKUPs use this column as the date key, so any gap leaves
' option rows without a rate or spot price. Returns False with a message
' describing exactly what's wrong.
Function CheckUnderlyingDateRange(ByRef errorMsg As String) As Boolean
    Const DATE_FIRST_ROW As Long = 11

    Dim wsFuture As Worksheet
    Dim firstDate As Variant
    Dim lastDate As Variant
    Dim lastDateRow As Long
    Dim requiredEnd As Date

    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)

    ' Force calc — workbook may be in manual mode, and column A is filled
    ' with WORKDAY formulas whose cached values may be stale.
    Application.Calculate

    ' First date is hardcoded at row 11 of column A
    firstDate = wsFuture.Cells(DATE_FIRST_ROW, 1).Value

    ' Find the last populated row in column A
    lastDateRow = wsFuture.Cells(wsFuture.Rows.count, 1).End(xlUp).Row

    If lastDateRow < DATE_FIRST_ROW Or Not IsDate(firstDate) Then
        errorMsg = "No dates found in column A starting at row " & DATE_FIRST_ROW & "." & vbNewLine & _
                   "The date axis is empty - underlyings have never been refreshed."
        CheckUnderlyingDateRange = False
        Exit Function
    End If

    lastDate = wsFuture.Cells(lastDateRow, 1).Value

    If Not IsDate(lastDate) Then
        errorMsg = "Last value in column A row " & lastDateRow & " is not a valid date: '" & lastDate & "'."
        CheckUnderlyingDateRange = False
        Exit Function
    End If

    ' Required end of underlying data: the option end date, capped at today
    ' (no historical data exists for future dates).
    If g_DateEnd > Date Then
        requiredEnd = Date
    Else
        requiredEnd = g_DateEnd
    End If

    ' Underlying must START on or before the option start date
    If CDate(firstDate) > g_DateStart Then
        errorMsg = "First date in Future et co column A is " & Format(firstDate, "yyyy-mm-dd") & _
                   ", which is AFTER the option start date " & Format(g_DateStart, "yyyy-mm-dd") & "." & vbNewLine & _
                   "Move UnderlyingStartDate to " & Format(g_DateStart, "yyyy-mm-dd") & " (or earlier) and re-refresh."
        CheckUnderlyingDateRange = False
        Exit Function
    End If

    ' Underlying must END on or after the required end date
    If CDate(lastDate) < requiredEnd Then
        errorMsg = "Last date in Future et co column A is " & Format(lastDate, "yyyy-mm-dd") & _
                   ", which is BEFORE the required end date " & Format(requiredEnd, "yyyy-mm-dd") & "." & vbNewLine & _
                   "Move UnderlyingEndDate to " & Format(requiredEnd, "yyyy-mm-dd") & " (or later) and re-refresh."
        CheckUnderlyingDateRange = False
        Exit Function
    End If

    errorMsg = ""
    CheckUnderlyingDateRange = True
End Function

Function CountUnprocessedRICs() As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim count As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    count = 0
    For i = 2 To lastRow
        If ws.Cells(i, 9).Value = "No" Then  ' Column I: Processed - only count explicit "No"
            count = count + 1
        End If
    Next i

    CountUnprocessedRICs = count
End Function

Function FindNextUnprocessedRIC(startFrom As Long) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    For i = startFrom To lastRow
        If ws.Cells(i, 9).Value = "No" Then  ' Column I: Processed
            FindNextUnprocessedRIC = i
            Exit Function
        End If
    Next i
    
    FindNextUnprocessedRIC = 0  ' No unprocessed RICs found
End Function

' Calculate batch number based on row position in RIC_List
' This ensures batch numbers are consistent even when resuming from a specific row
Function GetBatchNumberFromRow(startRow As Long) As Long
    ' Row 2 is first data row (row 1 is header)
    ' Batch 1 = rows 2 to (2 + batchSize - 1)
    ' Batch 2 = rows (2 + batchSize) to (2 + 2*batchSize - 1)
    ' Formula: ((startRow - 2) \ batchSize) + 1
    If g_BatchSize > 0 Then
        GetBatchNumberFromRow = ((startRow - 2) \ g_BatchSize) + 1
    Else
        GetBatchNumberFromRow = 1
    End If
End Function

Sub MarkBatchStatus(startRow As Long, endRow As Long, Status As String)
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)

    For i = startRow To endRow
        If ws.Cells(i, 9).Value <> "Yes" Then  ' Don't overwrite successful downloads
            ws.Cells(i, 9).Value = Status  ' Column I: Processed
            If Status = "Processing" Then
                ws.Cells(i, 10).Value = Now  ' Column J: Process_Time
            End If
        End If
    Next i
End Sub

Sub ValidateAndUpdateRICList(wsCollection As Worksheet, startRow As Long, endRow As Long)
    Dim wsRIC As Worksheet
    Dim i As Long
    Dim ricRow As Long
    Dim premium As Variant
    Dim iv As Variant
    Dim delta As Variant
    Dim validationResult As String
    
    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    
    For i = startRow To endRow
        ' Get the RIC row reference
        ricRow = wsCollection.Cells(i, 15).Value
        
        If ricRow > 0 Then
            ' Get data from collection sheet
            premium = wsCollection.Cells(i, 2).Value
            iv = wsCollection.Cells(i, 9).Value
            delta = wsCollection.Cells(i, 10).Value

            ' Validate and update RIC_List (check error first - VBA doesn't short-circuit And)
            Dim premiumValid As Boolean
            premiumValid = False
            If Not IsEmpty(premium) And Not IsError(premium) Then
                If IsNumeric(premium) Then
                    If premium > 0 Then
                        premiumValid = True
                    End If
                End If
            End If

            If premiumValid Then
                ' Successful download
                wsRIC.Cells(ricRow, 9).Value = "Yes"  ' Processed (column I)
                wsRIC.Cells(ricRow, 10).Value = Now     ' Process_Time (column J)
                wsRIC.Cells(ricRow, 11).Value = premium ' Premium (column K)

                If Not IsError(iv) And IsNumeric(iv) Then
                    wsRIC.Cells(ricRow, 12).Value = iv  ' IV (column L)
                    validationResult = ValidateIV(CDbl(iv), wsRIC.Cells(ricRow, 3).Value, _
                                                 GetSpotPrice(wsRIC.Cells(ricRow, 7).Value), wsRIC.Cells(ricRow, 2).Value)
                    wsRIC.Cells(ricRow, 14).Value = validationResult  ' Validation (column N)
                End If

                If Not IsError(delta) And IsNumeric(delta) Then
                    wsRIC.Cells(ricRow, 13).Value = delta  ' Delta (column M)
                End If

                ' Copy to staging - export all data that was successfully downloaded
                ' (validation is for quality info, not a gate for export)
                BuildAndAppendBatchRows wsCollection, i, 1
            Else
                ' Failed download
                wsRIC.Cells(ricRow, 9).Value = "Error"  ' Processed (column I)
                wsRIC.Cells(ricRow, 15).Value = "No data returned"  ' Error_Message (column O)
            End If
        End If
    Next i
End Sub

Sub ShowBatchSummaryFromRICList(startRow As Long, endRow As Long)
    Dim ws As Worksheet
    Dim successCount As Long
    Dim errorCount As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    
    successCount = 0
    errorCount = 0
    
    For i = startRow To endRow
        If ws.Cells(i, 9).Value = "Yes" Then  ' Column I: Processed
            successCount = successCount + 1
        ElseIf ws.Cells(i, 9).Value = "Error" Then
            errorCount = errorCount + 1
        End If
    Next i
    
'    MsgBox "Batch Complete!" & vbNewLine & vbNewLine & _
'          "Rows processed: " & startRow & " to " & endRow & vbNewLine & _
'          "Successful: " & successCount & vbNewLine & _
'          "Errors: " & errorCount & vbNewLine & _
'          "Skipped: " & (endRow - startRow + 1 - successCount - errorCount), _
'          vbInformation, "Batch Summary"
End Sub

' ============================================
' Keep existing RIC Builder Functions
' ============================================

Function BuildOptionRIC(rootRIC As String, strike As Double, _
                       maturityDate As Date, optionType As String) As String
    Dim monthCode As String
    Dim yearCode As String
    Dim strikeStr As String
    Dim ricMonth As Integer
    Dim monthCodeMethod As String

    ' Determine month based on optionMonthCodeMethod
    monthCodeMethod = GetOptionMonthCodeMethod()
    If monthCodeMethod = "Same Month" Then
        ricMonth = Month(maturityDate)
    Else
        ' Next Month
        ricMonth = Month(maturityDate) + 1
        If ricMonth > 12 Then ricMonth = 1
    End If

    monthCode = GetMonthCode(ricMonth, optionType)
    yearCode = Right(Year(maturityDate), 1)
    strikeStr = Replace(CStr(strike), ".", "")

    BuildOptionRIC = rootRIC & strikeStr & monthCode & yearCode

    ' Add suffix for expired options
    If maturityDate < Date Then
        BuildOptionRIC = BuildOptionRIC & "^" & monthCode & yearCode
    End If
End Function

Function GetMonthCode(monthNum As Integer, optionType As String) As String
    Dim callCodes As Variant
    Dim putCodes As Variant
    
    callCodes = Array("", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L")
    putCodes = Array("", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X")
    
    If optionType = "CALL" Then
        GetMonthCode = callCodes(monthNum)
    Else
        GetMonthCode = putCodes(monthNum)
    End If
End Function

Function BuildRHistoryFormula(ric As String, startDate As Date, endDate As Date) As String
    Dim startNum As Long
    Dim endNum As Long
    
    startNum = CLng(startDate)
    endNum = CLng(endDate)
    
    BuildRHistoryFormula = "=RHistory(""" & ric & """," & _
                          """.Timestamp;.Close"",""START:" & Format(startNum, "yyyy-mm-dd") & _
                          " END:" & Format(endNum, "yyyy-mm-dd") & " INTERVAL:1D"")"
End Function

' ============================================
' Keep existing refresh and calculation functions
' ============================================

Sub RefreshLSEGCollectionSheet()
    Dim wsCollection As Worksheet
    Set wsCollection = ThisWorkbook.Worksheets(SHEET_COLLECTION)
    'LSEG Download
    RefreshLSEGWithTimeout wsCollection, 60
    
    'FAKE DATA @@@
    'CopyFakeDownloadToDataCollection

    Application.StatusBar = False
End Sub

Sub CopyFakeDownloadToDataCollection()
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim lastRow As Long
    
    ' Set references
    Set wsSrc = ThisWorkbook.Sheets("FAKE_DOWNLOAD")
    Set wsDst = ThisWorkbook.Sheets("DataCollection")
    
    ' Find last used row in source (col A)
    lastRow = wsSrc.Cells(wsSrc.Rows.count, "A").End(xlUp).Row
    
    ' Clear destination columns A & B first (optional)
    wsDst.Range("A:B").ClearContents
    
    ' Copy values from FAKE_DOWNLOAD to Data Collection
    wsSrc.Range("A1:B" & lastRow).Copy
    wsDst.Range("A1").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
End Sub


Function ValidateIV(impliedVol As Double, strike As Double, _
                   spot As Double, maturity As Date) As String
    Dim moneyness As Double
    Dim timeToExp As Double

    moneyness = strike / spot
    timeToExp = (maturity - Date) / 365

    ' Check for convergence failures or invalid values
    If impliedVol < 0 Or IsEmpty(impliedVol) Or Not IsNumeric(impliedVol) Then
        ValidateIV = "Missing"
    ElseIf impliedVol = 0 Then
        ValidateIV = "Convergence Failed"
    ElseIf impliedVol < 0.001 Then
        ValidateIV = "Too Low"
    ElseIf impliedVol > 2 Then
        ValidateIV = "Too High"
    ElseIf impliedVol > 1.5 Then
        ValidateIV = "High"
    ElseIf timeToExp < 0 Then
        ValidateIV = "Expired"
    Else
        ValidateIV = "OK"
    End If
End Function

' ============================================
' RIC Retry Functions - Handle expired RIC suffix toggling
' ============================================

Function HasExpiredRICSuffix(ric As String) As Boolean
    ' Check if RIC has ^XX## suffix pattern (e.g., ^T25, ^L26)
    ' Pattern: caret, letter A-X, 2 digits
    Dim suffixStart As Long
    suffixStart = InStr(ric, "^")

    If suffixStart > 0 Then
        Dim suffix As String
        suffix = Mid(ric, suffixStart + 1)
        ' Should be 3 chars: letter + 2 digits
        If Len(suffix) = 3 Then
            If suffix Like "[A-X][0-9][0-9]" Then
                HasExpiredRICSuffix = True
                Exit Function
            End If
        End If
    End If
    HasExpiredRICSuffix = False
End Function

Function RemoveExpiredRICSuffix(ric As String) As String
    ' Remove ^XX## suffix if present
    Dim suffixStart As Long
    suffixStart = InStr(ric, "^")

    If suffixStart > 0 And HasExpiredRICSuffix(ric) Then
        RemoveExpiredRICSuffix = Left(ric, suffixStart - 1)
    Else
        RemoveExpiredRICSuffix = ric
    End If
End Function

Function AddExpiredRICSuffix(ric As String, monthCodeCall As String, yearCode As String) As String
    ' Add ^{monthCodeCall}{yearCode} suffix to RIC
    ' monthCodeCall should be the CALL month code (A-L)
    If HasExpiredRICSuffix(ric) Then
        ' Already has suffix, return as-is
        AddExpiredRICSuffix = ric
    Else
        AddExpiredRICSuffix = ric & "^" & monthCodeCall & yearCode
    End If
End Function

Function GetAlternateRIC(ric As String, monthCodeCall As String, yearCode As String) As String
    ' Toggle the expired RIC suffix
    ' If RIC has suffix, remove it; if not, add it
    If HasExpiredRICSuffix(ric) Then
        GetAlternateRIC = RemoveExpiredRICSuffix(ric)
    Else
        GetAlternateRIC = AddExpiredRICSuffix(ric, monthCodeCall, yearCode)
    End If
End Function

Function GetMonthCodeCallFromRIC(ric As String, maturityMonth As Long) As String
    ' Get the CALL month code for adding expired suffix
    ' Call month codes: A=Jan, B=Feb, ... L=Dec
    Dim monthCodes As String
    monthCodes = "ABCDEFGHIJKL"
    GetMonthCodeCallFromRIC = Mid(monthCodes, maturityMonth, 1)
End Function

Function ExtractYearCodeFromRIC(ric As String) As String
    ' Extract year code (1-2 digits) from end of RIC before any ^ suffix
    ' Examples: "FLG6800O26" -> "26", "OGBL1160V5" -> "5", "FLG6800O26^C26" -> "26"
    Dim cleanRIC As String
    Dim i As Long
    Dim yearDigits As String

    ' Remove ^ suffix if present
    If InStr(ric, "^") > 0 Then
        cleanRIC = Left(ric, InStr(ric, "^") - 1)
    Else
        cleanRIC = ric
    End If

    ' Extract trailing digits (1-2)
    yearDigits = ""
    For i = Len(cleanRIC) To 1 Step -1
        If IsNumeric(Mid(cleanRIC, i, 1)) Then
            yearDigits = Mid(cleanRIC, i, 1) & yearDigits
            If Len(yearDigits) = 2 Then Exit For
        Else
            Exit For
        End If
    Next i

    ExtractYearCodeFromRIC = yearDigits
End Function

Sub RetryFailedRICsInBatch(wsCollection As Worksheet, batchStartRow As Long, batchEndRow As Long)
    ' Retry failed RICs by toggling the expired suffix
    ' Called after initial validation to give failed RICs a second chance
    Dim wsRIC As Worksheet
    Dim i As Long
    Dim failedCount As Long
    Dim retryCount As Long
    Dim successCount As Long
    Dim ric As String
    Dim alternateRIC As String
    Dim maturityDate As Date
    Dim maturityMonth As Long
    Dim yearCode As String
    Dim monthCodeCall As String
    Dim formulaRow As Long

    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)

    ' Count failed RICs in batch
    failedCount = 0
    For i = batchStartRow To batchEndRow
        If wsRIC.Cells(i, 9).Value = "Error" Then
            failedCount = failedCount + 1
        End If
    Next i

    If failedCount = 0 Then
        Exit Sub  ' No failures to retry
    End If

    Application.StatusBar = "Retrying " & failedCount & " failed RICs with alternate format..."

    ' Update formulas for failed RICs with alternate RIC
    retryCount = 0
    For i = batchStartRow To batchEndRow
        If wsRIC.Cells(i, 9).Value = "Error" Then
            ric = wsRIC.Cells(i, 1).Value
            maturityDate = wsRIC.Cells(i, 2).Value
            ' Determine month based on optionMonthCodeMethod
            If GetOptionMonthCodeMethod() = "Same Month" Then
                maturityMonth = Month(maturityDate)
            Else
                maturityMonth = Month(maturityDate) + 1
                If maturityMonth > 12 Then maturityMonth = 1
            End If
            yearCode = wsRIC.Cells(i, 6).Value
            ' If year code is n/a or empty, extract from RIC
            If yearCode = "n/a" Or yearCode = "" Then
                yearCode = ExtractYearCodeFromRIC(ric)
            End If
            ' Ensure 2-digit year for expired suffix (pad with leading 2 if single digit)
            If Len(yearCode) = 1 Then
                yearCode = "2" & yearCode  ' Assumes 2020s decade
            End If
            monthCodeCall = GetMonthCodeCallFromRIC(ric, maturityMonth)

            alternateRIC = GetAlternateRIC(ric, monthCodeCall, yearCode)

            ' Find the corresponding formula row in DataCollection
            ' Formula rows are spaced by ROW_SPACING, starting at row 2
            formulaRow = 2 + (retryCount * ROW_SPACING)

            ' Find actual formula row by scanning for this RIC's row reference
            Dim j As Long
            For j = 0 To g_FormulaCount - 1
                formulaRow = 2 + (j * ROW_SPACING)
                If wsCollection.Cells(formulaRow, 15).Value = i Then
                    ' Found the formula row for this RIC
                    ' Update the RHistory formula with alternate RIC
                    wsCollection.Cells(formulaRow, 1).Formula = BuildRHistoryFormula(alternateRIC, g_DateStart, g_DateEnd)
                    wsCollection.Cells(formulaRow, 16).Value = alternateRIC  ' Store alternate RIC

                    ' Mark as retry in progress
                    wsRIC.Cells(i, 15).Value = "Retry: " & alternateRIC  ' Error_Message column
                    Exit For
                End If
            Next j

            retryCount = retryCount + 1
        End If
    Next i

    If retryCount = 0 Then
        Exit Sub
    End If

    ' Refresh LSEG data for updated formulas
    Application.StatusBar = "Refreshing LSEG data for " & retryCount & " retry RICs..."
    RefreshLSEGWithTimeout wsCollection, 30

    ' Wait for data to load
    Application.Wait Now + TimeValue("00:00:03")
    wsCollection.Calculate
    Application.Wait Now + TimeValue("00:00:02")

    ' Re-validate retried RICs
    Application.StatusBar = "Validating retry results..."
    successCount = 0

    For i = batchStartRow To batchEndRow
        If Left(wsRIC.Cells(i, 15).Value, 6) = "Retry:" Then
            alternateRIC = Trim(Mid(wsRIC.Cells(i, 15).Value, 8))

            ' Find formula row for this RIC
            For j = 0 To g_FormulaCount - 1
                formulaRow = 2 + (j * ROW_SPACING)
                If wsCollection.Cells(formulaRow, 15).Value = i Then
                    ' Check if data returned
                    Dim premium As Variant
                    Dim dataFound As Boolean
                    Dim lastPremium As Double
                    Dim lastIV As Double
                    Dim iv As Variant
                    Dim delta As Variant
                    Dim checkRow As Long

                    dataFound = False
                    lastPremium = 0
                    lastIV = 0

                    For checkRow = formulaRow To formulaRow + ROW_SPACING - 1
                        If IsEmpty(wsCollection.Cells(checkRow, 1).Value) Then
                            Exit For
                        End If

                        premium = wsCollection.Cells(checkRow, 2).Value
                        If Not IsEmpty(premium) And Not IsError(premium) Then
                            If IsNumeric(premium) Then
                                If premium > 0 Then
                                    dataFound = True
                                    lastPremium = premium
                                    iv = wsCollection.Cells(checkRow, 9).Value
                                    If Not IsError(iv) And IsNumeric(iv) Then
                                        lastIV = iv
                                    End If
                                    delta = wsCollection.Cells(checkRow, 10).Value
                                End If
                            End If
                        End If
                    Next checkRow

                    If dataFound Then
                        ' Retry successful! Update RIC_List with alternate RIC
                        successCount = successCount + 1
                        wsRIC.Cells(i, 1).Value = alternateRIC  ' Update RIC (column A)
                        wsRIC.Cells(i, 9).Value = "Yes"  ' Processed (column I)
                        wsRIC.Cells(i, 10).Value = Now  ' Process_Time (column J)
                        wsRIC.Cells(i, 11).Value = lastPremium  ' Premium (column K)
                        wsRIC.Cells(i, 15).Value = "Retry succeeded"  ' Error_Message (column O)

                        If lastIV > 0 Then
                            wsRIC.Cells(i, 12).Value = lastIV  ' IV (column L)
                            wsRIC.Cells(i, 14).Value = ValidateIV(lastIV, wsRIC.Cells(i, 3).Value, _
                                                     GetSpotPrice(wsRIC.Cells(i, 7).Value), wsRIC.Cells(i, 2).Value)
                        End If

                        If Not IsError(delta) And IsNumeric(delta) Then
                            wsRIC.Cells(i, 13).Value = delta  ' Delta (column M)
                        End If

                        ' Copy retry data to staging
                        Dim lastDataRow As Long
                        lastDataRow = formulaRow
                        Dim findRow As Long
                        For findRow = formulaRow To formulaRow + ROW_SPACING - 1
                            Dim cellVal As Variant
                            cellVal = wsCollection.Cells(findRow, 2).Value
                            If Not IsEmpty(cellVal) And Not IsError(cellVal) Then
                                lastDataRow = findRow
                            Else
                                Exit For
                            End If
                        Next findRow
                        BuildAndAppendBatchRows wsCollection, formulaRow, lastDataRow - formulaRow + 1
                    Else
                        ' Retry also failed
                        wsRIC.Cells(i, 15).Value = "Retry failed: " & alternateRIC
                    End If

                    Exit For
                End If
            Next j
        End If
    Next i

    Application.StatusBar = "Retry complete: " & successCount & " of " & retryCount & " succeeded"
End Sub

' ============================================
' Bulk-add expired suffix over a maturity range
' ============================================

' Sheet name and cell layout used by the picker. Shared by the entry sub
' and the OK/Cancel button callbacks below.
Private Const EXPIRED_PICKER_SHEET As String = "_ExpiredSuffixPicker"
Private Const EXPIRED_PICKER_START_CELL As String = "B4"
Private Const EXPIRED_PICKER_END_CELL As String = "B5"
Private Const EXPIRED_PICKER_LIST_FIRST_ROW As Long = 20

Public Sub AddExpiredSuffixForMaturityRange()
    ' Build a picker sheet with two dropdowns and OK/Cancel buttons, then
    ' return. The buttons drive the rest of the flow (ExpiredSuffixPicker_OK
    ' / ExpiredSuffixPicker_Cancel) so the user can interact with the
    ' dropdowns freely — no modal MsgBox blocking the worksheet.
    Dim wsRIC As Worksheet
    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)

    Dim sortedDates() As Date
    Dim n As Long
    n = CollectUniqueMaturities(wsRIC, sortedDates)
    If n = 0 Then
        MsgBox "No maturities found in RIC_List column B.", vbExclamation
        Exit Sub
    End If

    BuildExpiredSuffixPickerSheet sortedDates, n
End Sub

Private Sub BuildExpiredSuffixPickerSheet(sortedDates() As Date, n As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(EXPIRED_PICKER_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = EXPIRED_PICKER_SHEET
    End If

    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    On Error Resume Next
    ws.Cells.Validation.Delete
    On Error GoTo 0

    ' Remove any leftover buttons / shapes from a previous run
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' Write the source list down column D (any length, no DV string-limit issue)
    Dim i As Long
    For i = 0 To n - 1
        ws.Cells(EXPIRED_PICKER_LIST_FIRST_ROW + i, "D").Value = _
            Format(sortedDates(i), "yyyy-mm-dd")
    Next i
    Dim listRange As String
    listRange = "=$D$" & EXPIRED_PICKER_LIST_FIRST_ROW & _
                ":$D$" & (EXPIRED_PICKER_LIST_FIRST_ROW + n - 1)

    ' Header / instructions
    ws.Range("A1").Value = "Add Expired Suffix - pick maturity range"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Pick start and end from the dropdowns, then click OK below."

    ws.Range("A4").Value = "Start maturity:"
    ws.Range("A5").Value = "End maturity:"
    ws.Range("A4:A5").Font.Bold = True

    ' Default to first / last so user can just hit OK for the full range
    ws.Range(EXPIRED_PICKER_START_CELL).Value = Format(sortedDates(0), "yyyy-mm-dd")
    ws.Range(EXPIRED_PICKER_END_CELL).Value = Format(sortedDates(n - 1), "yyyy-mm-dd")

    With ws.Range(EXPIRED_PICKER_START_CELL).Validation
        .Add Type:=xlValidateList, Formula1:=listRange
        .InCellDropdown = True
    End With
    With ws.Range(EXPIRED_PICKER_END_CELL).Validation
        .Add Type:=xlValidateList, Formula1:=listRange
        .InCellDropdown = True
    End With

    ws.Columns("A:B").AutoFit
    ws.Columns("D:D").Hidden = True

    ' OK / Cancel buttons (Forms-style — OnAction routes to public macros)
    Dim btnTop As Double, btnLeft As Double, btnWidth As Double, btnHeight As Double
    btnLeft = ws.Range("B7").Left
    btnTop = ws.Range("B7").Top
    btnWidth = 80
    btnHeight = 26

    Dim okBtn As Shape, cancelBtn As Shape
    Set okBtn = ws.Shapes.AddFormControl(xlButtonControl, btnLeft, btnTop, btnWidth, btnHeight)
    okBtn.Name = "btnExpiredOK"
    okBtn.TextFrame.Characters.Text = "OK"
    okBtn.OnAction = "ExpiredSuffixPicker_OK"

    Set cancelBtn = ws.Shapes.AddFormControl(xlButtonControl, _
                                             btnLeft + btnWidth + 10, btnTop, _
                                             btnWidth, btnHeight)
    cancelBtn.Name = "btnExpiredCancel"
    cancelBtn.TextFrame.Characters.Text = "Cancel"
    cancelBtn.OnAction = "ExpiredSuffixPicker_Cancel"

    ws.Activate
    ws.Range(EXPIRED_PICKER_START_CELL).Select
End Sub

Public Sub ExpiredSuffixPicker_OK()
    ' Triggered by the OK button on the picker sheet. Reads the selected
    ' dates, applies the expired suffix, hides the picker, and reports.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(EXPIRED_PICKER_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim startStr As String, endStr As String
    startStr = Trim(CStr(ws.Range(EXPIRED_PICKER_START_CELL).Value))
    endStr = Trim(CStr(ws.Range(EXPIRED_PICKER_END_CELL).Value))

    If startStr = "" Or endStr = "" Then
        MsgBox "Start or end maturity not selected.", vbExclamation
        Exit Sub
    End If
    If Not IsDate(startStr) Or Not IsDate(endStr) Then
        MsgBox "Could not parse selected maturities ('" & startStr & "', '" & endStr & "').", _
               vbExclamation
        Exit Sub
    End If

    Dim startDate As Date, endDate As Date
    startDate = CDate(startStr)
    endDate = CDate(endStr)
    If endDate < startDate Then
        Dim tmpDate As Date
        tmpDate = startDate
        startDate = endDate
        endDate = tmpDate
    End If

    Dim wsRIC As Worksheet
    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)

    Dim updated As Long
    updated = ApplyExpiredSuffixToRange(wsRIC, startDate, endDate)

    ' Return to RIC_List and hide the picker
    On Error Resume Next
    wsRIC.Activate
    ws.Visible = xlSheetHidden
    On Error GoTo 0

    MsgBox "Expired suffix added to " & updated & " RIC(s) " & _
           "between " & Format(startDate, "yyyy-mm-dd") & " and " & _
           Format(endDate, "yyyy-mm-dd") & ".", vbInformation
End Sub

Public Sub ExpiredSuffixPicker_Cancel()
    ' Triggered by the Cancel button on the picker sheet. Hides it and
    ' returns the user to RIC_List without applying anything.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(EXPIRED_PICKER_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_RIC_LIST).Activate
    ws.Visible = xlSheetHidden
    On Error GoTo 0
End Sub

Private Function CollectUniqueMaturities(wsRIC As Worksheet, ByRef sortedDates() As Date) As Long
    Dim lastRow As Long
    lastRow = wsRIC.Cells(wsRIC.Rows.count, 2).End(xlUp).Row
    If lastRow < 2 Then
        ReDim sortedDates(0 To 0)
        CollectUniqueMaturities = 0
        Exit Function
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, v As Variant, key As String
    For i = 2 To lastRow
        v = wsRIC.Cells(i, 2).Value
        If IsDate(v) Then
            key = Format(CDate(v), "yyyy-mm-dd")
            If Not dict.Exists(key) Then dict.Add key, CDate(v)
        End If
    Next i

    Dim count As Long
    count = dict.count
    If count = 0 Then
        ReDim sortedDates(0 To 0)
        CollectUniqueMaturities = 0
        Exit Function
    End If

    ReDim sortedDates(0 To count - 1)
    Dim k As Variant, idx As Long
    idx = 0
    For Each k In dict.Keys
        sortedDates(idx) = dict(k)
        idx = idx + 1
    Next k

    ' Insertion sort (count is small)
    Dim a As Long, b As Long, tmp As Date
    For a = 1 To count - 1
        tmp = sortedDates(a)
        b = a - 1
        Do While b >= 0
            If sortedDates(b) <= tmp Then Exit Do
            sortedDates(b + 1) = sortedDates(b)
            b = b - 1
        Loop
        sortedDates(b + 1) = tmp
    Next a

    CollectUniqueMaturities = count
End Function

Private Function ApplyExpiredSuffixToRange(wsRIC As Worksheet, _
                                           startDate As Date, _
                                           endDate As Date) As Long
    Dim lastRow As Long
    lastRow = wsRIC.Cells(wsRIC.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim updated As Long
    Dim i As Long
    Dim ric As String, newRIC As String
    Dim maturityDate As Date, maturityMonth As Long
    Dim yearCode As String, monthCodeCall As String
    Dim mat As Variant

    Application.ScreenUpdating = False
    For i = 2 To lastRow
        ric = CStr(wsRIC.Cells(i, 1).Value)
        mat = wsRIC.Cells(i, 2).Value
        If ric = "" Or Not IsDate(mat) Then GoTo NextRow

        maturityDate = CDate(mat)
        If maturityDate < startDate Or maturityDate > endDate Then GoTo NextRow
        If HasExpiredRICSuffix(ric) Then GoTo NextRow

        If GetOptionMonthCodeMethod() = "Same Month" Then
            maturityMonth = Month(maturityDate)
        Else
            maturityMonth = Month(maturityDate) + 1
            If maturityMonth > 12 Then maturityMonth = 1
        End If

        yearCode = CStr(wsRIC.Cells(i, 6).Value)
        If yearCode = "n/a" Or yearCode = "" Then yearCode = ExtractYearCodeFromRIC(ric)
        If Len(yearCode) = 1 Then yearCode = "2" & yearCode

        monthCodeCall = GetMonthCodeCallFromRIC(ric, maturityMonth)
        newRIC = AddExpiredRICSuffix(ric, monthCodeCall, yearCode)

        If newRIC <> ric Then
            wsRIC.Cells(i, 1).Value = newRIC
            updated = updated + 1
        End If
NextRow:
    Next i
    Application.ScreenUpdating = True

    ApplyExpiredSuffixToRange = updated
End Function

' ============================================
' Keep remaining helper functions
' ============================================

Sub GenerateQualityReport()
    Dim ws As Worksheet
    Dim wsRIC As Worksheet
    Dim row As Long
    Dim i As Long
    Dim lastRow As Long
    Dim totalProcessed As Long, totalSuccess As Long, totalErrors As Long
    Dim ivCounts As Object
    Dim maturityCounts As Object
    Dim status As String
    Dim validation As String
    Dim maturityVal As Variant
    Dim maturityKey As String
    Dim counts As Variant

    Set ws = ThisWorkbook.Worksheets(SHEET_QUALITY)
    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    Set ivCounts = CreateObject("Scripting.Dictionary")
    Set maturityCounts = CreateObject("Scripting.Dictionary")

    ws.Cells.Clear

    ' Pre-seed IV categories so they appear in a fixed, meaningful order
    Dim ivCategories As Variant
    ivCategories = Array("OK", "High", "Too High", "Too Low", "Missing", "Convergence Failed", "Expired")
    For i = 0 To UBound(ivCategories)
        ivCounts.Add CStr(ivCategories(i)), 0
    Next i

    lastRow = wsRIC.Cells(wsRIC.Rows.count, "A").End(xlUp).Row

    ' Single pass over RIC_List
    For i = 2 To lastRow
        status = CStr(wsRIC.Cells(i, 9).Value)         ' Column I: Processed
        validation = Trim(CStr(wsRIC.Cells(i, 14).Value))  ' Column N: Validation
        maturityVal = wsRIC.Cells(i, 2).Value          ' Column B: Maturity

        ' Tally Processed/Success/Error
        If status <> "No" And status <> "" Then
            totalProcessed = totalProcessed + 1
            If status = "Yes" Then
                totalSuccess = totalSuccess + 1
            ElseIf status = "Error" Then
                totalErrors = totalErrors + 1
            End If
        End If

        ' Tally IV validation
        If validation <> "" Then
            If ivCounts.Exists(validation) Then
                ivCounts(validation) = ivCounts(validation) + 1
            Else
                ivCounts.Add validation, 1
            End If
        End If

        ' Tally maturity coverage by year-month
        If IsDate(maturityVal) Then
            maturityKey = Format(maturityVal, "yyyy-mm")
            If Not maturityCounts.Exists(maturityKey) Then
                maturityCounts.Add maturityKey, Array(0, 0, 0)
            End If
            counts = maturityCounts(maturityKey)
            counts(0) = counts(0) + 1
            If status = "Yes" Then counts(1) = counts(1) + 1
            If status = "Error" Then counts(2) = counts(2) + 1
            maturityCounts(maturityKey) = counts
        End If
    Next i

    ' --- Header ---
    ws.Range("A1").Value = "Option Data Quality Report"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Range("A3").Value = "Root RIC: " & g_RootRIC

    ' --- Summary Statistics ---
    row = 5
    ws.Cells(row, 1).Value = "SUMMARY STATISTICS"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row + 1, 1).Value = "Total RICs:"
    ws.Cells(row + 1, 2).Value = lastRow - 1
    ws.Cells(row + 2, 1).Value = "Processed:"
    ws.Cells(row + 2, 2).Value = totalProcessed
    ws.Cells(row + 3, 1).Value = "Successful:"
    ws.Cells(row + 3, 2).Value = totalSuccess
    ws.Cells(row + 4, 1).Value = "Errors:"
    ws.Cells(row + 4, 2).Value = totalErrors
    ws.Cells(row + 5, 1).Value = "Success Rate:"
    If totalProcessed > 0 Then
        ws.Cells(row + 5, 2).Value = totalSuccess / totalProcessed
        ws.Cells(row + 5, 2).NumberFormat = "0.0%"
    End If

    ' --- IV Validation Breakdown ---
    row = row + 7
    ws.Cells(row, 1).Value = "IV VALIDATION BREAKDOWN"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).Value = "Category"
    ws.Cells(row, 2).Value = "Count"
    ws.Cells(row, 3).Value = "Pct"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Font.Bold = True

    Dim ivKey As Variant
    Dim ivTotal As Long
    For Each ivKey In ivCounts.Keys
        ivTotal = ivTotal + CLng(ivCounts(ivKey))
    Next ivKey

    For Each ivKey In ivCounts.Keys
        row = row + 1
        ws.Cells(row, 1).Value = CStr(ivKey)
        ws.Cells(row, 2).Value = ivCounts(ivKey)
        If ivTotal > 0 Then
            ws.Cells(row, 3).Value = CLng(ivCounts(ivKey)) / ivTotal
            ws.Cells(row, 3).NumberFormat = "0.0%"
        End If
    Next ivKey

    ' --- Maturity Coverage ---
    row = row + 2
    ws.Cells(row, 1).Value = "MATURITY COVERAGE (by year-month)"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).Value = "Year-Month"
    ws.Cells(row, 2).Value = "Total"
    ws.Cells(row, 3).Value = "Success"
    ws.Cells(row, 4).Value = "Errors"
    ws.Cells(row, 5).Value = "Success%"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Font.Bold = True

    If maturityCounts.count > 0 Then
        ' Copy keys to array and sort (yyyy-mm strings sort chronologically)
        Dim sortedKeys() As String
        Dim k As Long
        Dim m As Long, n As Long
        Dim tmp As String
        ReDim sortedKeys(1 To maturityCounts.count)
        k = 1
        Dim mKey As Variant
        For Each mKey In maturityCounts.Keys
            sortedKeys(k) = CStr(mKey)
            k = k + 1
        Next mKey
        For m = 1 To UBound(sortedKeys) - 1
            For n = m + 1 To UBound(sortedKeys)
                If sortedKeys(m) > sortedKeys(n) Then
                    tmp = sortedKeys(m)
                    sortedKeys(m) = sortedKeys(n)
                    sortedKeys(n) = tmp
                End If
            Next n
        Next m

        Dim countsArr As Variant
        For m = 1 To UBound(sortedKeys)
            row = row + 1
            countsArr = maturityCounts(sortedKeys(m))
            ws.Cells(row, 1).Value = sortedKeys(m)
            ws.Cells(row, 2).Value = countsArr(0)
            ws.Cells(row, 3).Value = countsArr(1)
            ws.Cells(row, 4).Value = countsArr(2)
            If CLng(countsArr(0)) > 0 Then
                ws.Cells(row, 5).Value = CLng(countsArr(1)) / CLng(countsArr(0))
                ws.Cells(row, 5).NumberFormat = "0.0%"
            End If
        Next m
    End If

    ws.Columns("A:E").AutoFit
End Sub

' Keep remaining helper functions unchanged...
Function LoadConfiguration() As Boolean
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    g_RootRIC = ws.Range("rootRIC").Value
    g_StrikeStep = ws.Range("steps").Value
    g_LotSize = ws.Range("lotSize").Value
    g_Currency = ws.Range("currency").Value
    g_DateStart = ws.Range("dateStart").Value
    g_DateEnd = ws.Range("dateEnd").Value
    g_BatchSize = ws.Range("batchSize").Value
    g_NamePrefix = ws.Range("namePrefix").Value
    
    g_PutStrikeMin = ws.Range("minStrikePut").Value
    g_PutStrikeMax = ws.Range("maxStrikePut").Value
    g_CallStrikeMin = ws.Range("minStrikeCall").Value
    g_CallStrikeMax = ws.Range("maxStrikeCall").Value

    LoadConfiguration = True
    Exit Function
    
ErrorHandler:
    LoadConfiguration = False
End Function

Function GetSpotPrice(underlyingTicker As String) As Double
    Dim wsFuture As Worksheet
    Dim startCol As Long
    Dim startRow As Long
    Dim currentCol As Long
    Dim foundUnderlying As String
    Dim priceCol As Long
    Dim lastRow As Long
    Dim spotPrice As Double

    On Error Resume Next

    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)

    ' Get starting position from RANGE_DOWNLOAD
    startCol = wsFuture.Range(RANGE_DOWNLOAD).Column
    startRow = wsFuture.Range(RANGE_DOWNLOAD).Row

    If startCol = 0 Or startRow = 0 Then
        ' Named range not found, return 0
        GetSpotPrice = 0
        On Error GoTo 0
        Exit Function
    End If

    ' Scan for the underlying ticker (every 3rd column, same pattern as RefreshFutureUnderlyings)
    currentCol = startCol

    Do While True
        foundUnderlying = Trim(CStr(wsFuture.Cells(startRow, currentCol).Value))

        If foundUnderlying = underlyingTicker Then
            ' Found the matching underlying
            ' The "Last Price" column is at currentCol + 1 (relative to the Date column at currentCol - 1)
            ' So the actual price data is at currentCol
            priceCol = currentCol

            ' Find the last non-empty row in the price column
            lastRow = wsFuture.Cells(wsFuture.Rows.count, priceCol).End(xlUp).Row

            ' Get the most recent spot price (skip header rows)
            If lastRow > startRow + 2 Then
                spotPrice = wsFuture.Cells(lastRow, priceCol).Value

                If spotPrice > 0 Then
                    GetSpotPrice = spotPrice
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        End If

        currentCol = currentCol + 3

        ' Break if we find 3 consecutive empty blocks
        If wsFuture.Cells(startRow, currentCol).Value = "" And _
           wsFuture.Cells(startRow, currentCol + 3).Value = "" And _
           wsFuture.Cells(startRow, currentCol + 6).Value = "" Then
            Exit Do
        End If
    Loop

    ' If we reach here, underlying not found or no valid price - return 0
    GetSpotPrice = 0

    On Error GoTo 0
End Function

Function GetBloombergTicker(underlyingRIC As String, Optional ricRowRef As Long = 0) As String
    ' Get Bloomberg ticker for an option
    ' Priority: 1) Bloom_Ticker from RIC_List (if ricRowRef provided)
    '           2) rootBB named range from Config sheet
    '           3) underlyingRIC as fallback
    Dim wsConfig As Worksheet
    Dim wsRICList As Worksheet
    Dim bloomTicker As String
    Dim rootBB As String

    ' First, try to get Bloom_Ticker from RIC_List if row reference provided
    If ricRowRef > 0 Then
        On Error Resume Next
        Set wsRICList = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
        On Error GoTo 0

        If Not wsRICList Is Nothing Then
            bloomTicker = Trim(CStr(wsRICList.Cells(ricRowRef, 8).Value))  ' Column H = Bloom_Ticker
            If bloomTicker <> "" Then
                GetBloombergTicker = bloomTicker
                Exit Function
            End If
        End If
    End If

    ' Fallback: use rootBB named range from Config sheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets(SHEET_CONFIG)
    rootBB = Trim(CStr(wsConfig.Range("rootBB").Value))
    On Error GoTo 0

    If rootBB <> "" Then
        GetBloombergTicker = rootBB
    Else
        ' Final fallback: return input RIC
        GetBloombergTicker = underlyingRIC
    End If
End Function

Function GetRiskFreeRate() As Double
    Dim wsFuture As Worksheet
    Dim rfrRange As Range
    Dim lastRow As Long
    Dim rfr As Double

    On Error Resume Next
    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)
    Set rfrRange = wsFuture.Range(RANGE_RFR)

    If Not rfrRange Is Nothing Then
        lastRow = rfrRange.Cells(rfrRange.Rows.count, 1).End(xlUp).Row - rfrRange.Row + 1
        If lastRow > 1 Then
            rfr = rfrRange.Cells(lastRow, 1).Value
            If rfr > 0 Then
                GetRiskFreeRate = rfr
                On Error GoTo 0
                Exit Function
            End If
        End If
    End If

    GetRiskFreeRate = 0.04
    On Error GoTo 0
End Function

Sub ClearCollectionSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_COLLECTION)
    ws.Cells.Clear

    ' Set all column headers including new Greeks
    ws.Range("A1").Value = "Spot_Date"
    ws.Range("B1").Value = "Premium"
    ws.Range("C1").Value = "Ticker"
    ws.Range("D1").Value = "Maturity"
    ws.Range("E1").Value = "Interest_Rate"
    ws.Range("F1").Value = "Spot"
    ws.Range("G1").Value = "Strike"
    ws.Range("H1").Value = "Type"
    ws.Range("I1").Value = "Implied_Volatility"
    ws.Range("J1").Value = "Delta"
    ws.Range("K1").Value = "Vega"
    ws.Range("L1").Value = "Gamma"
    ws.Range("M1").Value = "Theta"
    ws.Range("N1").Value = "Rho"
    ws.Range("O1").Value = "RIC_Row_Ref"
    ws.Range("P1").Value = "RIC"
    ws.Range("Q1").Value = "Lot_size"
    ws.Range("R1").Value = "Name"
    ws.Range("S1").Value = "Reference"
    ws.Range("T1").Value = "ccy_pair"
    ws.Range("U1").Value = "Dividend"
    ws.Range("V1").Value = "DDELTA/DVOL"
    ws.Range("W1").Value = "DDELTA/DVOLDVOL"
    ws.Range("X1").Value = "DDELTA/DTIME"
    ws.Range("Y1").Value = "DGAMMA/DSPOT"
    ws.Range("Z1").Value = "DGAMMA/DVOL"
    ws.Range("AA1").Value = "DVEGA/DVOL"
    ws.Range("AB1").Value = "DVEGA/DVOLDVOL"
    ws.Range("AC1").Value = "RIC_Underlying"

    ws.Range("A1:AC1").Font.Bold = True

    ' Format column A (Spot_Date) as YYYY-MM-DD hh:mm:ss for CSV export
    ws.Columns("A:A").NumberFormat = "yyyy-mm-dd hh:mm:ss"

    ' Format column D (Maturity) as YYYY-MM-DD hh:mm:ss for CSV export
    ws.Columns("D:D").NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Sub SetupQualitySheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_QUALITY)
    ws.Cells.Clear
End Sub

Function CheckUnderlyings() As Boolean
    Dim wsRIC As Worksheet
    Dim wsFuture As Worksheet
    Dim uniqueUnderlyings As Collection
    Dim existingUnderlyings As Collection
    Dim missingUnderlyings As Collection
    Dim underlyingInfo As String
    Dim i As Long
    Dim lastRow As Long
    Dim underlyingValue As String
    Dim startCol As Long
    Dim rowDownload As Long
    Dim currentCol As Long
    Dim foundUnderlying As String
    Dim reportMsg As String

    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)
    Set uniqueUnderlyings = New Collection
    Set existingUnderlyings = New Collection
    Set missingUnderlyings = New Collection

    ' Step 1: Extract unique underlyings from RIC_List column G
    Application.StatusBar = "Extracting unique underlyings from RIC_List..."
    lastRow = wsRIC.Cells(wsRIC.Rows.count, "A").End(xlUp).Row

    For i = 2 To lastRow
        underlyingValue = Trim(CStr(wsRIC.Cells(i, 7).Value))  ' Column G

        If underlyingValue <> "" And underlyingValue <> "0" Then
            ' Add to unique collection (will ignore duplicates)
            On Error Resume Next
            uniqueUnderlyings.Add underlyingValue, underlyingValue
            On Error GoTo 0
        End If
    Next i

    ' Step 2: Get starting column from named range
    On Error Resume Next
    startCol = wsFuture.Range(RANGE_DOWNLOAD).Column
    rowDownload = wsFuture.Range(RANGE_DOWNLOAD).Row
    On Error GoTo 0

    If startCol = 0 Then
        MsgBox "Named range '" & RANGE_DOWNLOAD & "' not found in " & SHEET_FUTURE, vbExclamation
        Exit Function
    End If

    ' Step 3: Scan SHEET_FUTURE for existing underlyings (every 3rd column)
    Application.StatusBar = "Scanning SHEET_FUTURE for existing underlyings..."
    currentCol = startCol

    Do While True
        foundUnderlying = Trim(CStr(wsFuture.Cells(rowDownload, currentCol).Value))   ' Metadata column

        If foundUnderlying <> "" Then
            existingUnderlyings.Add foundUnderlying & " (Col " & (64 + currentCol) & ")", foundUnderlying
        End If

        currentCol = currentCol + 3  ' Move to next 3-column block

        ' Break if we find 3 consecutive empty blocks
        If wsFuture.Cells(rowDownload, currentCol).Value = "" And _
           wsFuture.Cells(rowDownload, currentCol + 3).Value = "" And _
           wsFuture.Cells(rowDownload, currentCol + 6).Value = "" Then
            Exit Do
        End If
    Loop

    ' Step 4: Compare and identify missing underlyings
    Dim underlying As Variant
    Dim found As Boolean

    For Each underlying In uniqueUnderlyings
        found = False

        Dim existing As Variant
        For Each existing In existingUnderlyings
            If InStr(CStr(existing), CStr(underlying)) > 0 Then
                found = True
                Exit For
            End If
        Next existing

        If Not found Then
            missingUnderlyings.Add underlying
        End If
    Next underlying

    ' Step 5: Generate report
    reportMsg = "UNDERLYINGS CHECK REPORT" & vbNewLine & String(30, "=") & vbNewLine & vbNewLine

    If existingUnderlyings.count > 0 Then
        reportMsg = reportMsg & "EXISTING UNDERLYINGS:" & vbNewLine
        For Each existing In existingUnderlyings
            reportMsg = reportMsg & "  " & existing & vbNewLine
        Next existing
        reportMsg = reportMsg & vbNewLine
    End If

    If missingUnderlyings.count > 0 Then
        reportMsg = reportMsg & "MISSING UNDERLYINGS (need download):" & vbNewLine
        For Each underlying In missingUnderlyings
            reportMsg = reportMsg & "  " & underlying & vbNewLine
        Next underlying
        reportMsg = reportMsg & vbNewLine
        MsgBox reportMsg, vbInformation, "Underlyings Check Results"
    Else
        reportMsg = reportMsg & "All underlyings are available in SHEET_FUTURE!"
    End If

    Application.StatusBar = False
    'MsgBox reportMsg, vbInformation, "Underlyings Check Results"

    ' Return True if all underlyings are available, False if some are missing
    CheckUnderlyings = (missingUnderlyings.count = 0)
End Function

' ============================================
' CSV EXPORT HELPER FUNCTIONS (High Performance)
' ============================================

' Format a single field for CSV output with proper escaping
Private Function FormatCSVField(val As Variant) As String
    Dim s As String

    If IsError(val) Then
        FormatCSVField = ""
    ElseIf IsEmpty(val) Or IsNull(val) Then
        FormatCSVField = ""
    ElseIf IsDate(val) Then
        FormatCSVField = Format(val, "yyyy-mm-dd hh:mm:ss")
    Else
        s = CStr(val)
        ' Escape fields containing comma, quote, or newline
        If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
            s = Replace(s, """", """""")  ' Double up any quotes
            s = """" & s & """"           ' Wrap in quotes
        End If
        FormatCSVField = s
    End If
End Function

' Write a 2D array to CSV file using VBA file I/O (much faster than Excel SaveAs)
Private Sub WriteArrayToCSV(data As Variant, filePath As String, Optional headers As Variant)
    Dim fileNum As Integer
    Dim i As Long, j As Long
    Dim rowStr As String
    Dim colCount As Long
    Dim rowCount As Long
    Dim fieldVal As String

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    ' Write headers if provided
    If Not IsMissing(headers) Then
        If IsArray(headers) Then
            colCount = UBound(headers, 2)
            rowStr = ""
            For j = 1 To colCount
                If j > 1 Then rowStr = rowStr & ","
                rowStr = rowStr & FormatCSVField(headers(1, j))
            Next j
            Print #fileNum, rowStr
        End If
    End If

    ' Handle case where data might be a single value (1 row, 1 col)
    If Not IsArray(data) Then
        Print #fileNum, FormatCSVField(data)
        Close #fileNum
        Exit Sub
    End If

    ' Write data rows
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)

    For i = 1 To rowCount
        rowStr = ""
        For j = 1 To colCount
            If j > 1 Then rowStr = rowStr & ","
            rowStr = rowStr & FormatCSVField(data(i, j))
        Next j
        Print #fileNum, rowStr
    Next i

    Close #fileNum
End Sub

' ============================================
' ADD GREEK FORMULAS TO EXTERNAL SHEET (Batch Processing)
' ============================================

' Entry point: Run from Excel on the active sheet
Public Sub RunAddGreekFormulasBatch()
    ' Wrapper to run AddGreekFormulasBatch on the active sheet
    ' Can be assigned to a button or run from Macros menu

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Confirm with user
    If MsgBox("Add Greek formulas to all rows in '" & ws.Name & "'?" & vbNewLine & vbNewLine & _
              "This will process columns I-N (primary Greeks) and U-AA (second-order Greeks)." & vbNewLine & _
              "Existing values in these columns will be overwritten.", _
              vbYesNo + vbQuestion, "Add Greek Formulas") = vbNo Then
        Exit Sub
    End If

    AddGreekFormulasBatch ws, 500
End Sub

Public Sub AddGreekFormulasBatch(ws As Worksheet, Optional batchSize As Long = 500)
    ' Adds Black-Scholes Greek formulas to a sheet with DataCollection structure
    ' Processes in batches: add formulas -> calculate -> paste values
    ' This allows processing 10k+ rows without Excel becoming unresponsive
    '
    ' Parameters:
    '   ws - Worksheet with DataCollection column structure
    '   batchSize - Number of rows per batch (default 500)
    '
    ' Column structure expected:
    '   A: Spot_Date, B: Premium, D: Maturity, E: Interest_Rate
    '   F: Spot, G: Strike, H: Type (C/P)
    '   I-N: Primary Greeks (IV, Delta, Vega, Gamma, Theta, Rho)
    '   V-AB: Second-order Greeks

    Dim lastRow As Long
    Dim batchStart As Long
    Dim batchEnd As Long
    Dim totalRows As Long
    Dim processedRows As Long
    Dim calcMode As XlCalculation

    On Error GoTo ErrorHandler

    ' Find last row with Premium data (column B)
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found in column B (Premium).", vbExclamation
        Exit Sub
    End If

    totalRows = lastRow - 1  ' Exclude header
    processedRows = 0

    ' Store and set calculation mode
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    ' Process in batches
    batchStart = 2  ' Start after header

    Do While batchStart <= lastRow
        batchEnd = Application.Min(batchStart + batchSize - 1, lastRow)

        ' Update status
        Application.StatusBar = "Processing rows " & batchStart & " to " & batchEnd & _
                               " of " & lastRow & " (" & _
                               Format(processedRows / totalRows, "0%") & ")"

        ' Add formulas to Greek columns (I-N, V-AB)
        AddGreekFormulasToRange ws, batchStart, batchEnd

        ' Calculate the batch
        ws.Calculate

        ' Convert formulas to values
        ConvertGreekFormulasToValues ws, batchStart, batchEnd

        ' Update counters
        processedRows = processedRows + (batchEnd - batchStart + 1)
        batchStart = batchEnd + 1

        DoEvents  ' Keep Excel responsive
    Loop

    ' Cleanup
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Greek formulas processed successfully!" & vbNewLine & _
           "Rows processed: " & totalRows & vbNewLine & _
           "Worksheet: " & ws.Name, vbInformation
    Exit Sub

ErrorHandler:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error in AddGreekFormulasBatch: " & Err.Description, vbCritical
End Sub

Private Sub AddGreekFormulasToRange(ws As Worksheet, startRow As Long, endRow As Long)
    ' Add Greek formulas using FormulaR1C1 for efficiency
    ' Columns: I=IV, J=Delta, K=Vega, L=Gamma, M=Theta, N=Rho
    '          V-AB = Second-order Greeks
    '
    ' Formula references (R1C1 notation):
    '   RC1 = Spot_Date (A), RC2 = Premium (B), RC4 = Maturity (D)
    '   RC5 = Interest_Rate (E), RC6 = Spot (F), RC7 = Strike (G), RC8 = Type (H)
    '   RC9 = Implied_Volatility (I)

    Dim rng As Range

    ' Column I (9): Implied Volatility
    Set rng = ws.Range(ws.Cells(startRow, 9), ws.Cells(endRow, 9))
    rng.FormulaR1C1 = "=IF(OR(RC2="""",ISERROR(RC2)),""""," & _
        "IFERROR(GBlackScholesImpVolBisection(RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC2),""NA""))"

    ' Column J (10): Delta
    Set rng = ws.Range(ws.Cells(startRow, 10), ws.Cells(endRow, 10))
    rng.FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-1]=""NA""),""""," & _
        "IFERROR(GBlackScholesNGreeks(""d"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC[-1]),""NA""))"

    ' Column K (11): Vega
    Set rng = ws.Range(ws.Cells(startRow, 11), ws.Cells(endRow, 11))
    rng.FormulaR1C1 = "=IF(OR(RC[-2]="""",RC[-2]=""NA""),""""," & _
        "IFERROR(GBlackScholesNGreeks(""v"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC[-2]),""NA""))"

    ' Column L (12): Gamma
    Set rng = ws.Range(ws.Cells(startRow, 12), ws.Cells(endRow, 12))
    rng.FormulaR1C1 = "=IF(OR(RC[-3]="""",RC[-3]=""NA""),""""," & _
        "IFERROR(GBlackScholesNGreeks(""g"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC[-3]),""NA""))"

    ' Column M (13): Theta
    Set rng = ws.Range(ws.Cells(startRow, 13), ws.Cells(endRow, 13))
    rng.FormulaR1C1 = "=IF(OR(RC[-4]="""",RC[-4]=""NA""),""""," & _
        "IFERROR(GBlackScholesNGreeks(""t"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC[-4]),""NA""))"

    ' Column N (14): Rho
    Set rng = ws.Range(ws.Cells(startRow, 14), ws.Cells(endRow, 14))
    rng.FormulaR1C1 = "=IF(OR(RC[-5]="""",RC[-5]=""NA""),""""," & _
        "IFERROR(GBlackScholesNGreeks(""r"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC[-5]),""NA""))"

    ' Second-order Greeks (U-AB) use CGBlackScholes
    ' Column U (21): DDeltaDVol (Vanna)
    Set rng = ws.Range(ws.Cells(startRow, 21), ws.Cells(endRow, 21))
    rng.FormulaR1C1 = "=IF(OR(RC9="""",RC9=""NA""),""""," & _
        "IFERROR(CGBlackScholes(""dddv"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC9),""NA""))"

    ' Column V (22): DDeltaDVolDVol
    Set rng = ws.Range(ws.Cells(startRow, 22), ws.Cells(endRow, 22))
    rng.FormulaR1C1 = "=IF(OR(RC9="""",RC9=""NA""),""""," & _
        "IFERROR(CGBlackScholes(""dvv"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC9),""NA""))"

    ' Column W (23): DDeltaDTime (Charm)
    Set rng = ws.Range(ws.Cells(startRow, 23), ws.Cells(endRow, 23))
    rng.FormulaR1C1 = "=IF(OR(RC9="""",RC9=""NA""),""""," & _
        "IFERROR(CGBlackScholes(""dt"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC9),""NA""))"

    ' Column X (24): DGammaDSpot
    Set rng = ws.Range(ws.Cells(startRow, 24), ws.Cells(endRow, 24))
    rng.FormulaR1C1 = "=IF(OR(RC9="""",RC9=""NA""),""""," & _
        "IFERROR(CGBlackScholes(""gps"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC9),""NA""))"

    ' Column Y (25): DGammaDVol (Zomma)
    Set rng = ws.Range(ws.Cells(startRow, 25), ws.Cells(endRow, 25))
    rng.FormulaR1C1 = "=IF(OR(RC9="""",RC9=""NA""),""""," & _
        "IFERROR(CGBlackScholes(""gpv"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC9),""NA""))"

    ' Column Z (26): DVegaDVol (Vomma)
    Set rng = ws.Range(ws.Cells(startRow, 26), ws.Cells(endRow, 26))
    rng.FormulaR1C1 = "=IF(OR(RC9="""",RC9=""NA""),""""," & _
        "IFERROR(CGBlackScholes(""dvdv"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC9),""NA""))"

    ' Column AA (27): DVegaDVolDVol (Ultima)
    Set rng = ws.Range(ws.Cells(startRow, 27), ws.Cells(endRow, 27))
    rng.FormulaR1C1 = "=IF(OR(RC9="""",RC9=""NA""),""""," & _
        "IFERROR(CGBlackScholes(""vvv"",RC8,RC6,RC7,(RC4-RC1)/365,RC5,0,RC9),""NA""))"
End Sub

Private Sub ConvertGreekFormulasToValues(ws As Worksheet, startRow As Long, endRow As Long)
    ' Convert formula columns to values (removes formulas, keeps calculated results)
    ' This is faster than Copy/PasteSpecial and doesn't use clipboard
    Dim rng As Range

    ' Primary Greeks (I:N = columns 9-14)
    Set rng = ws.Range(ws.Cells(startRow, 9), ws.Cells(endRow, 14))
    rng.Value = rng.Value

    ' Second-order Greeks (U:AA = columns 21-27)
    Set rng = ws.Range(ws.Cells(startRow, 21), ws.Cells(endRow, 27))
    rng.Value = rng.Value
End Sub


' RunAddGreekFormulasBatch()
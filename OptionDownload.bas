Attribute VB_Name = "OptionDownload"
' ============================================
' MODULE 1: Global Configuration and Types
' ============================================

Option Explicit

' Configuration Variables
Public g_NamePrefix As String
Public g_RootRIC As String
Public g_SpotPrice As Double
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

' Sheet Names
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_RIC_LIST As String = "RIC_List"  ' Now used for progress tracking
Public Const SHEET_COLLECTION As String = "DataCollection"
Public Const SHEET_STAGING As String = "Staging"
Public Const SHEET_QUALITY As String = "QualityReport"
Public Const SHEET_FUTURE As String = "Future et co"

' RANGE FUTURE DOWNLOAD
Public Const RANGE_DOWNLOAD As String = "UnderlyingDownload"  '1st column for 1st underlying. Expand right for more underlyings, +3 columns each
Public Const RANGE_UNDERLYING_START_DATE As String = "UnderlyingStartDate"
Public Const RANGE_UNDERLYING_END_DATE As String = "UnderlyingEndDate"
Public Const RANGE_RFR As String = "RFR"

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
' Keep existing refresh and calculation functions
' ============================================

Sub RefreshFutureSheet()
    Dim wsFuture As Worksheet
    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)

    RefreshLSEGWithTimeout wsFuture, 60

    MsgBox "Double check data in : " & SHEET_FUTURE, vbExclamation

    Application.StatusBar = False
End Sub

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

    ' Simple bubble sort
    For i = 1 To arraySize - 1
        For j = i + 1 To arraySize
            If underlyingsArray(i) > underlyingsArray(j) Then
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

        ' Clear previous data in both columns (preserve header at row 1 of range)
        wsConfig.Range(wsConfig.Cells(ricDataStartRow, ricBloombergRange.Column), _
                      wsConfig.Cells(ricDataEndRow, ricBloombergRange.Column + 1)).ClearContents

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

    ' Step 4: Clear existing underlying data
    Application.StatusBar = "Clearing existing underlying data..."
    ' Calculate end column (each underlying uses 3 columns, clear up to 100 columns as reasonable limit)
    clearEndCol = startCol + 99

    ' Clear the data area - clear up to 10000 rows from startRow (not entire sheet)
    ' This prevents accidentally clearing formulas/data far below the expected range
    Dim clearEndRow As Long
    clearEndRow = startRow + 10000

    wsFuture.Range(wsFuture.Cells(startRow, startCol), wsFuture.Cells(clearEndRow, clearEndCol)).ClearContents

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

        ' Add the RHistory formula
        wsFuture.Cells(formulaRow, formulaCol).Formula = _
            "=RHistory(""" & underlyingsArray(i) & """," & _
            """.Timestamp;.Close""," & _
            """NBROWS:5000 INTERVAL:1D"",,""Sort:ASC"")"

        ' Add the underlying symbol in the RANGE_DOWNLOAD row/column
        wsFuture.Cells(startRow, currentCol).Value = underlyingsArray(i)

        ' Add metadata in the next column
        wsFuture.Cells(startRow, currentCol + 1).Value = "Added: " & Format(Now, "yyyy-mm-dd hh:mm")

        ' Move to next 3-column block
        currentCol = currentCol + 3
    Next i

    ' Step 6: Refresh LSEG workspace
    Application.StatusBar = "Refreshing LSEG data for " & arraySize & " underlyings..."
    RefreshLSEGWithTimeout wsFuture, 120

    MsgBox "Added " & arraySize & " underlyings to SHEET_FUTURE in alphabetical order." & vbNewLine & _
           "Data refresh complete. Please verify the downloaded data." & vbNewLine & _
           "IMPORTANT: Add Bloomberg equivalent for underlying RICs in RicBloomberg range.", vbInformation

    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error in RefreshFutureUnderlyings: " & Err.Description, vbExclamation
End Sub



' ============================================
' MODULE 2: Main Process Controller
' ============================================

Sub InitializeWorkbook()
    ' Create necessary sheets if they don't exist
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim i As Integer
    
    sheetNames = Array(SHEET_CONFIG, SHEET_RIC_LIST, SHEET_COLLECTION, _
                      SHEET_STAGING, SHEET_QUALITY)
    
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
    SetupStagingSheet
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
    g_BatchState = bpsIdle

    ' Find first unprocessed batch
    batchStart = FindNextUnprocessedRIC(2)
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
    Dim wsRIC As Worksheet
    Dim wsCollection As Worksheet
    Dim i As Long
    Dim ric As String
    Dim currentRow As Long
    Const ROW_SPACING As Long = 300

    ' Check stop flag
    If g_StopRequested Then
        ProcessBatch_Abort
        Exit Sub
    End If

    g_BatchState = bpsSetupFormulas
    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    Set wsCollection = ThisWorkbook.Worksheets(SHEET_COLLECTION)

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Clearing collection and staging sheets..."
    ClearCollectionSheet
    ClearStagingSheet

    ' Setup formulas
    currentRow = 2
    g_FormulaCount = 0

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Setting up formulas..."

    For i = g_BatchStartRow To g_BatchEndRow
        ric = wsRIC.Cells(i, 1).Value

        ' Skip if already processed
        If wsRIC.Cells(i, 8).Value = "Yes" Then GoTo NextRIC

        ' Update status
        If g_FormulaCount Mod 10 = 0 Then
            Application.StatusBar = "Batch #" & g_BatchCounter & ": Preparing RIC " & (g_FormulaCount + 1) & " - " & ric
        End If

        currentRow = 2 + (g_FormulaCount * ROW_SPACING)

        ' Setup formula
        wsCollection.Cells(currentRow, 1).Formula = BuildRHistoryFormula(ric, g_DateStart, g_DateEnd)

        ' Store metadata
        wsCollection.Cells(currentRow, 7).Value = wsRIC.Cells(i, 3).Value
        wsCollection.Cells(currentRow, 8).Value = Left(wsRIC.Cells(i, 4).Value, 1)
        wsCollection.Cells(currentRow, 4).Value = wsRIC.Cells(i, 2).Value
        wsCollection.Cells(currentRow, 15).Value = i
        wsCollection.Cells(currentRow, 16).Value = ric

        ' Setup metadata only (Greek formulas added after refresh)
        SetupRHistoryAndMetadata wsCollection, currentRow, ROW_SPACING, _
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
        ProcessBatch_ProcessResults
        Exit Sub
    End If

    Set wsCollection = ThisWorkbook.Worksheets(SHEET_COLLECTION)

    ' Check if data ready (with progress tracking)
    If IsDataReady(wsCollection, readyCount, totalChecks) Then
        ' Data ready, proceed
        Application.StatusBar = "Batch #" & g_BatchCounter & ": All data ready (" & readyCount & "/" & totalChecks & " cells) - processing..."
        ProcessBatch_ProcessResults
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
    Dim wsCollection As Worksheet
    Dim i As Long
    Dim processRow As Long
    Const ROW_SPACING As Long = 300

    ' Check stop flag
    If g_StopRequested Then
        ProcessBatch_Abort
        Exit Sub
    End If

    g_BatchState = bpsProcessingResults
    Set wsCollection = ThisWorkbook.Worksheets(SHEET_COLLECTION)

    ' Add Greek formulas to rows with data (after LSEG refresh)
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
        
'        If calcTimeout > 30 Then  ' 30 second timeout
'            MsgBox "Calculation timeout - proceeding anyway", vbExclamation
'            Exit Do
'        End If
    Loop

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Validating and copying data to staging..."
    ValidateAndUpdateRICListWithSpacing wsCollection, g_FormulaCount

    'Application.StatusBar = "Batch #" & g_BatchCounter & ": Final calculations..."
    'Application.Calculate

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Saving to CSV..."
    SaveStagingToCSV g_BatchCounter

    ' Save workbook every 3 batches
    If g_BatchCounter Mod 3 = 0 Then
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
    Const ROW_SPACING As Long = 300

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

' Stop batch processing
Sub StopBatchProcessing()
    g_StopRequested = True
    Application.StatusBar = "Stop requested - will halt after current operation..."
    MsgBox "Batch processing will stop after current phase completes.", vbInformation
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

' Helper function to build VLOOKUP formula for underlying spot price
Function BuildSpotVLOOKUPFormula(rowNum As Long, underlyingTicker As String) As String
    ' Build VLOOKUP formula to get spot price from SHEET_FUTURE based on date in column A
    ' Returns formula that looks up the underlying's price column
    BuildSpotVLOOKUPFormula = "=IFERROR(INDEX('" & SHEET_FUTURE & "'!GetUnderlyingPriceColumn(""" & underlyingTicker & """),MATCH(A" & rowNum & ",'" & SHEET_FUTURE & "'!GetUnderlyingDateColumn(""" & underlyingTicker & """),1)),"""")"
End Function

' New optimized approach: Setup minimal metadata before LSEG refresh
Sub SetupRHistoryAndMetadata(ws As Worksheet, startRow As Long, maxRows As Long, _
                             strike As Double, optType As String, maturity As Date, ricRowRef As Long, underlyingTicker As String, optionRic As String)
    Dim i As Long
    Dim endRow As Long
    Dim wsFuture As Worksheet
    Dim underlyingCol As Long
    Dim startCol As Long
    Dim startRowFuture As Long
    Dim currentCol As Long
    Dim foundUnderlying As String
    Dim rfrRange As Range
    Dim rfrRow As Long
    Dim rfrCol As Long
    Dim rfrLastRow As Long

    ' Find the underlying column in SHEET_FUTURE
    Set wsFuture = ThisWorkbook.Worksheets(SHEET_FUTURE)
    On Error Resume Next
    startCol = wsFuture.Range(RANGE_DOWNLOAD).Column
    startRowFuture = wsFuture.Range(RANGE_DOWNLOAD).Row
    On Error GoTo 0

    underlyingCol = 0
    If startCol > 0 Then
        currentCol = startCol
        While currentCol <= 100
            foundUnderlying = Trim(CStr(wsFuture.Cells(startRowFuture, currentCol).Value))
            If foundUnderlying = underlyingTicker Then
                underlyingCol = currentCol
            End If
            currentCol = currentCol + 3
            If wsFuture.Cells(startRowFuture, currentCol).Value = "" And _
               wsFuture.Cells(startRowFuture, currentCol + 3).Value = "" Then
            End If
        Wend
    End If

    ' Get RFR range position and find last row with data
    On Error Resume Next
    Set rfrRange = wsFuture.Range(RANGE_RFR)
    On Error GoTo 0

    If Not rfrRange Is Nothing Then
        rfrRow = rfrRange.Row
        rfrCol = rfrRange.Column
        rfrLastRow = wsFuture.Cells(wsFuture.Rows.count, 1).End(xlUp).Row
    End If

    endRow = startRow + maxRows - 1

    ' Setup basic metadata and VLOOKUP formulas (NO Greek formulas yet)
    For i = startRow To endRow
        ' Store metadata
        ws.Cells(i, 3).Value = GetBloombergTicker(underlyingTicker) & " " & Left(optType, 1) & " " & strike
        ws.Cells(i, 4).Value = maturity

        ' Column E: Interest_rate - VLOOKUP from RFR range with dynamic last row
        If Not rfrRange Is Nothing Then
            ws.Cells(i, 5).Formula = "=IFERROR(VLOOKUP(A" & i & ",'" & SHEET_FUTURE & "'!" & _
                wsFuture.Range(wsFuture.Cells(rfrRow, 1), wsFuture.Cells(rfrLastRow, rfrCol)).Address(False, False) & _
                "," & rfrCol & ",TRUE),""not found"")"
        Else
            ws.Cells(i, 5).Value = "not found"
        End If

        ' Column F: Spot - VLOOKUP from underlying data (matches 10000-row clearing limit)
        If underlyingCol > 0 Then
            ws.Cells(i, 6).Formula = "=IFERROR(VLOOKUP(A" & i & ",'" & SHEET_FUTURE & "'!" & _
                wsFuture.Cells(startRowFuture + 2, underlyingCol - 1).Address(False, False) & ":" & _
                wsFuture.Cells(startRowFuture + 10000, underlyingCol).Address(False, False) & ",2,TRUE),"""")"
        Else
            ws.Cells(i, 6).Value = GetSpotPrice(underlyingTicker)
        End If

        ' Store additional metadata
        ws.Cells(i, 7).Value = strike
        ws.Cells(i, 8).Value = Left(optType, 1)
        ws.Cells(i, 15).Value = ricRowRef
        ws.Cells(i, 16).Value = optionRic
        ws.Cells(i, 17).Value = g_LotSize
        ws.Cells(i, 18).Value = g_NamePrefix & " " & Left(optType, 1) & " " & strike & " " & Format(maturity, "mmm-yyyy")
        ws.Cells(i, 19).Value = GetBloombergTicker(underlyingTicker)
        ws.Cells(i, 20).Value = g_Currency
        ws.Cells(i, 21).Value = 0
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

    ' Find first and last rows with premium data
    For i = startRow To endRow
        If Not IsEmpty(ws.Cells(i, 1).Value) And Not IsEmpty(ws.Cells(i, 2).Value) Then
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

    ' Column I (9): Implied Volatility
    ws.Range(ws.Cells(firstDataRow, 9), ws.Cells(lastDataRow, 9)).FormulaR1C1 = _
        "=IF(RC[-7]="""","""",GBlackScholesImpVolBisection(LOWER(RC[-1]),RC[-3],RC[-2],(RC[-5]-RC[-8])/365,RC[-4],0,RC[-7]))"

    ' Column J (10): Delta
    ws.Range(ws.Cells(firstDataRow, 10), ws.Cells(lastDataRow, 10)).FormulaR1C1 = _
        "=IF(RC[-8]="""","""",GBlackScholesNGreeks(""d"",LOWER(RC[-2]),RC[-4],RC[-3],(RC[-6]-RC[-9])/365,RC[-5],0,RC[-1]))"

    ' Column K (11): Vega
    ws.Range(ws.Cells(firstDataRow, 11), ws.Cells(lastDataRow, 11)).FormulaR1C1 = _
        "=IF(RC[-9]="""","""",GBlackScholesNGreeks(""v"",LOWER(RC[-3]),RC[-5],RC[-4],(RC[-7]-RC[-10])/365,RC[-6],0,RC[-2]))"

    ' Column L (12): Gamma
    ws.Range(ws.Cells(firstDataRow, 12), ws.Cells(lastDataRow, 12)).FormulaR1C1 = _
        "=IF(RC[-10]="""","""",GBlackScholesNGreeks(""g"",LOWER(RC[-4]),RC[-6],RC[-5],(RC[-8]-RC[-11])/365,RC[-7],0,RC[-3]))"

    ' Column M (13): Theta
    ws.Range(ws.Cells(firstDataRow, 13), ws.Cells(lastDataRow, 13)).FormulaR1C1 = _
        "=IF(RC[-11]="""","""",GBlackScholesNGreeks(""t"",LOWER(RC[-5]),RC[-7],RC[-6],(RC[-9]-RC[-12])/365,RC[-8],0,RC[-4]))"

    ' Column N (14): Rho
    ws.Range(ws.Cells(firstDataRow, 14), ws.Cells(lastDataRow, 14)).FormulaR1C1 = _
        "=IF(RC[-12]="""","""",GBlackScholesNGreeks(""r"",LOWER(RC[-6]),RC[-8],RC[-7],(RC[-10]-RC[-13])/365,RC[-9],0,RC[-5]))"

    ' Column V (22): DDELTA/DVOL
    ws.Range(ws.Cells(firstDataRow, 22), ws.Cells(lastDataRow, 22)).FormulaR1C1 = _
        "=IF(RC[-20]="""","""",CGBlackScholes(""dddv"",LOWER(RC[-14]),RC[-16],RC[-15],(RC[-18]-RC[-21])/365,RC[-17],0,RC[-13],RC[-12]))"

    ' Column W (23): DDELTA/DVOLDVOL
    ws.Range(ws.Cells(firstDataRow, 23), ws.Cells(lastDataRow, 23)).FormulaR1C1 = _
        "=IF(RC[-21]="""","""",CGBlackScholes(""dvv"",LOWER(RC[-15]),RC[-17],RC[-16],(RC[-19]-RC[-22])/365,RC[-18],0,RC[-14],RC[-13]))"

    ' Column X (24): Charm (DDELTA/DTIME)
    ws.Range(ws.Cells(firstDataRow, 24), ws.Cells(lastDataRow, 24)).FormulaR1C1 = _
        "=IF(RC[-22]="""","""",CGBlackScholes(""dt"",LOWER(RC[-16]),RC[-18],RC[-17],(RC[-20]-RC[-23])/365,RC[-19],0,RC[-15],RC[-14]))"

    ' Column Y (25): DGamma/DSpot
    ws.Range(ws.Cells(firstDataRow, 25), ws.Cells(lastDataRow, 25)).FormulaR1C1 = _
        "=IF(RC[-23]="""","""",CGBlackScholes(""gps"",LOWER(RC[-17]),RC[-19],RC[-18],(RC[-21]-RC[-24])/365,RC[-20],0,RC[-16],RC[-15]))"

    ' Column Z (26): Zomma (DGAMMA/DVOL)
    ws.Range(ws.Cells(firstDataRow, 26), ws.Cells(lastDataRow, 26)).FormulaR1C1 = _
        "=IF(RC[-24]="""","""",CGBlackScholes(""gpv"",LOWER(RC[-18]),RC[-20],RC[-19],(RC[-22]-RC[-25])/365,RC[-21],0,RC[-17],RC[-16]))"

    ' Column AA (27): Vomma (DVEGA/DVOL)
    ws.Range(ws.Cells(firstDataRow, 27), ws.Cells(lastDataRow, 27)).FormulaR1C1 = _
        "=IF(RC[-25]="""","""",CGBlackScholes(""dvdv"",LOWER(RC[-19]),RC[-21],RC[-20],(RC[-23]-RC[-26])/365,RC[-22],0,RC[-18],RC[-17]))"

    ' Column AB (28): Ultima (DVEGA/DVOLDVOL)
    ws.Range(ws.Cells(firstDataRow, 28), ws.Cells(lastDataRow, 28)).FormulaR1C1 = _
        "=IF(RC[-26]="""","""",CGBlackScholes(""vvv"",LOWER(RC[-20]),RC[-22],RC[-21],(RC[-24]-RC[-27])/365,RC[-23],0,RC[-19],RC[-18]))"

    ' Calculate the worksheet to populate formulas
    'ws.Calculate

CleanUp:
    ' Re-enable Excel features and restore original calculation mode
    'Application.ScreenUpdating = True
    Application.Calculation = originalCalcMode
    'Application.EnableEvents = True
End Sub

' New function to copy only rows with LSEG data to staging
Sub CopyDataRowsToStaging(ws As Worksheet, startRow As Long, maxRows As Long)
    Dim i As Long
    Dim endRow As Long
    Dim wsDest As Worksheet
    Dim NextRow As Long

    Set wsDest = ThisWorkbook.Worksheets(SHEET_STAGING)

    endRow = startRow + maxRows - 1

    ' Copy only rows that have premium data
    For i = startRow To endRow
        If Not IsEmpty(ws.Cells(i, 2).Value) And IsNumeric(ws.Cells(i, 2).Value) Then
            ' This row has LSEG data, copy it to staging with proper column mapping
            NextRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).Row + 1

            ' Map columns to staging sheet (matching CSV export format)
            ' Use .Value2 to preserve full numeric precision without formatting
            wsDest.Cells(NextRow, 1).Value2 = ws.Cells(i, 1).Value2   ' Spot_Date
            wsDest.Cells(NextRow, 2).Value2 = ws.Cells(i, 2).Value2   ' Premium
            wsDest.Cells(NextRow, 3).Value2 = ws.Cells(i, 3).Value2   ' Ticker
            wsDest.Cells(NextRow, 4).Value2 = ws.Cells(i, 4).Value2   ' Maturity
            wsDest.Cells(NextRow, 5).Value2 = ws.Cells(i, 5).Value2   ' Interest_rate
            wsDest.Cells(NextRow, 6).Value2 = ws.Cells(i, 6).Value2   ' Spot
            wsDest.Cells(NextRow, 7).Value2 = ws.Cells(i, 7).Value2   ' Strike
            wsDest.Cells(NextRow, 8).Value2 = ws.Cells(i, 8).Value2   ' Type
            wsDest.Cells(NextRow, 9).Value2 = ws.Cells(i, 9).Value2   ' Implied_Volatility
            wsDest.Cells(NextRow, 10).Value2 = ws.Cells(i, 10).Value2 ' Delta
            wsDest.Cells(NextRow, 11).Value2 = ws.Cells(i, 11).Value2 ' Vega
            wsDest.Cells(NextRow, 12).Value2 = ws.Cells(i, 12).Value2 ' Gamma
            wsDest.Cells(NextRow, 13).Value2 = ws.Cells(i, 13).Value2 ' Theta
            wsDest.Cells(NextRow, 14).Value2 = ws.Cells(i, 14).Value2 ' Rho
            wsDest.Cells(NextRow, 15).Value2 = ws.Cells(i, 17).Value2 ' Lot_size (from col Q)
            wsDest.Cells(NextRow, 16).Value2 = ws.Cells(i, 18).Value2 ' Name
            wsDest.Cells(NextRow, 17).Value2 = ws.Cells(i, 19).Value2 ' Reference
            wsDest.Cells(NextRow, 18).Value2 = ws.Cells(i, 20).Value2 ' ccy_pair
            wsDest.Cells(NextRow, 19).Value2 = ws.Cells(i, 21).Value2 ' Dividend
            ' DDELTA_DSPOT removed - columns shifted
            wsDest.Cells(NextRow, 20).Value2 = ws.Cells(i, 22).Value2 ' DDELTA_DVOL (from col V)
            wsDest.Cells(NextRow, 21).Value2 = ws.Cells(i, 23).Value2 ' DDELTA_DVOLDVOL (from col W)
            wsDest.Cells(NextRow, 22).Value2 = ws.Cells(i, 24).Value2 ' DDELTA_DTIME (from col X)
            wsDest.Cells(NextRow, 23).Value2 = ws.Cells(i, 25).Value2 ' DGAMMA_DSPOT (from col Y)
            wsDest.Cells(NextRow, 24).Value2 = ws.Cells(i, 26).Value2 ' DGAMMA_DVOL (from col Z)
            wsDest.Cells(NextRow, 25).Value2 = ws.Cells(i, 27).Value2 ' DVEGA_DVOL (from col AA)
            wsDest.Cells(NextRow, 26).Value2 = ws.Cells(i, 28).Value2 ' DVEGA_DVOLDVOL (from col AB)
        Else
            ' No more data in this section
            Exit For
        End If
    Next i
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
    Const ROW_SPACING As Long = 300
    
    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    
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
                If Not IsEmpty(premium) And IsNumeric(premium) And premium > 0 Then
                    dataFound = True
                    lastPremium = premium
                    
                    ' Get IV if available
                    iv = wsCollection.Cells(checkRow, 9).Value
                    If IsNumeric(iv) Then
                        lastIV = iv
                    End If
                    
                    ' Get Delta if available
                    delta = wsCollection.Cells(checkRow, 10).Value
                End If
            Next checkRow
            
            ' Update RIC_List with results
            If dataFound Then
                ' Successful download
                wsRIC.Cells(ricRow, 8).Value = "Yes"  ' Processed
                wsRIC.Cells(ricRow, 9).Value = Now     ' Process_Time
                wsRIC.Cells(ricRow, 10).Value = lastPremium ' Premium
                
                If lastIV > 0 Then
                    wsRIC.Cells(ricRow, 11).Value = lastIV  ' IV
                    validationResult = ValidateIV(lastIV, wsRIC.Cells(ricRow, 3).Value, _
                                                 GetSpotPrice(wsRIC.Cells(ricRow, 7).Value), wsRIC.Cells(ricRow, 2).Value)
                    wsRIC.Cells(ricRow, 13).Value = validationResult  ' Validation
                End If
                
                If IsNumeric(delta) Then
                    wsRIC.Cells(ricRow, 12).Value = delta  ' Delta
                End If
                
                ' Copy last row to staging if valid
                If validationResult = "OK" Or validationResult = "High" Then
                    ' Find the last row with data in this section
                    Dim lastDataRow As Long
                    lastDataRow = formulaRow
                    Dim findRow As Long
                    For findRow = formulaRow To formulaRow + ROW_SPACING - 1
                        If Not IsEmpty(wsCollection.Cells(findRow, 2).Value) Then
                            lastDataRow = findRow
                        Else
                            Exit For
                        End If
                    Next findRow
                    ' Copy the rows with data to staging
                    CopyDataRowsToStaging wsCollection, formulaRow, lastDataRow - formulaRow + 1
                End If
            Else
                ' Failed download
                wsRIC.Cells(ricRow, 8).Value = "Error"
                wsRIC.Cells(ricRow, 14).Value = "No data returned"  ' Error_Message
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
            .Range("G1").Value = "Underlying"
            .Range("H1").Value = "Processed"
        End With
    End If
    
    ' Add additional tracking columns if they don't exist
    With ws
        If .Range("I1").Value = "" Then .Range("I1").Value = "Process_Time"
        If .Range("J1").Value = "" Then .Range("J1").Value = "Premium"
        If .Range("K1").Value = "" Then .Range("K1").Value = "IV"
        If .Range("L1").Value = "" Then .Range("L1").Value = "Delta"
        If .Range("M1").Value = "" Then .Range("M1").Value = "Validation"
        If .Range("N1").Value = "" Then .Range("N1").Value = "Error_Message"
        
        ' Format headers
        .Range("A1:N1").Font.Bold = True
        .Range("A1:N1").Interior.Color = RGB(200, 200, 200)
        
        ' Add conditional formatting to Processed column
        Dim lastRow As Long
        lastRow = .Cells(.Rows.count, "A").End(xlUp).Row
        If lastRow > 1 Then
            With .Range("H2:H" & lastRow).FormatConditions
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
    ws.Columns("A:N").AutoFit
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

Function CountUnprocessedRICs() As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim count As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    count = 0
    For i = 2 To lastRow
        If ws.Cells(i, 8).Value <> "Yes" Then  ' Column H: Processed
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
        If ws.Cells(i, 8).Value <> "Yes" Then  ' Column H: Processed
            FindNextUnprocessedRIC = i
            Exit Function
        End If
    Next i
    
    FindNextUnprocessedRIC = 0  ' No unprocessed RICs found
End Function

Sub MarkBatchStatus(startRow As Long, endRow As Long, Status As String)
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    
    For i = startRow To endRow
        If ws.Cells(i, 8).Value <> "Yes" Then  ' Don't overwrite successful downloads
            ws.Cells(i, 8).Value = Status  ' Column H: Processed
            If Status = "Processing" Then
                ws.Cells(i, 9).Value = Now  ' Column I: Process_Time
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
            
            ' Validate and update RIC_List
            If Not IsEmpty(premium) And IsNumeric(premium) And premium > 0 Then
                ' Successful download
                wsRIC.Cells(ricRow, 8).Value = "Yes"  ' Processed
                wsRIC.Cells(ricRow, 9).Value = Now     ' Process_Time
                wsRIC.Cells(ricRow, 10).Value = premium ' Premium
                
                If IsNumeric(iv) Then
                    wsRIC.Cells(ricRow, 11).Value = iv  ' IV
                    validationResult = ValidateIV(CDbl(iv), wsRIC.Cells(ricRow, 3).Value, _
                                                 GetSpotPrice(wsRIC.Cells(ricRow, 7).Value), wsRIC.Cells(ricRow, 2).Value)
                    wsRIC.Cells(ricRow, 13).Value = validationResult  ' Validation
                End If
                
                If IsNumeric(delta) Then
                    wsRIC.Cells(ricRow, 12).Value = delta  ' Delta
                End If
                
                ' Copy to staging if valid
                If validationResult = "OK" Or validationResult = "High" Then
                    ' Copy this single row to staging
                    CopyDataRowsToStaging wsCollection, i, 1
                End If
            Else
                ' Failed download
                wsRIC.Cells(ricRow, 8).Value = "Error"
                wsRIC.Cells(ricRow, 14).Value = "No data returned"  ' Error_Message
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
        If ws.Cells(i, 8).Value = "Yes" Then
            successCount = successCount + 1
        ElseIf ws.Cells(i, 8).Value = "Error" Then
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
    
    monthCode = GetMonthCode(Month(maturityDate), optionType)
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
' Keep remaining helper functions
' ============================================

Sub GenerateQualityReport()
    Dim ws As Worksheet
    Dim wsRIC As Worksheet
    Dim summaryRow As Long
    Dim totalProcessed As Long
    Dim totalSuccess As Long
    Dim totalErrors As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_QUALITY)
    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    
    ws.Cells.Clear
    
    ' Count statistics from RIC_List
    Dim lastRow As Long
    Dim i As Long
    lastRow = wsRIC.Cells(wsRIC.Rows.count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If wsRIC.Cells(i, 8).Value <> "No" And wsRIC.Cells(i, 8).Value <> "" Then
            totalProcessed = totalProcessed + 1
            If wsRIC.Cells(i, 8).Value = "Yes" Then
                totalSuccess = totalSuccess + 1
            ElseIf wsRIC.Cells(i, 8).Value = "Error" Then
                totalErrors = totalErrors + 1
            End If
        End If
    Next i
    
    ' Generate report
    ws.Range("A1").Value = "Option Data Quality Report"
    ws.Range("A2").Value = "Generated: " & Now
    ws.Range("A3").Value = "Root RIC: " & g_RootRIC
    
    summaryRow = 5
    ws.Cells(summaryRow, 1).Value = "Summary Statistics"
    ws.Cells(summaryRow + 1, 1).Value = "Total RICs:"
    ws.Cells(summaryRow + 1, 2).Value = lastRow - 1
    
    ws.Cells(summaryRow + 2, 1).Value = "Processed:"
    ws.Cells(summaryRow + 2, 2).Value = totalProcessed
    
    ws.Cells(summaryRow + 3, 1).Value = "Successful:"
    ws.Cells(summaryRow + 3, 2).Value = totalSuccess
    
    ws.Cells(summaryRow + 4, 1).Value = "Errors:"
    ws.Cells(summaryRow + 4, 2).Value = totalErrors
    
    ws.Cells(summaryRow + 5, 1).Value = "Success Rate:"
    If totalProcessed > 0 Then
        ws.Cells(summaryRow + 5, 2).Value = Format(totalSuccess / totalProcessed, "0.0%")
    End If
    
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Columns("A:B").AutoFit
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
        ' Named range not found, fall back to global variable
        GetSpotPrice = g_SpotPrice
        On Error GoTo 0
        Exit Function
    End If

    ' Scan for the underlying ticker (every 3rd column, same pattern as RefreshFutureUnderlyings)
    currentCol = startCol

    While currentCol <= 100  ' Reasonable upper limit
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
        End If
    Wend

    ' If we reach here, underlying not found or no valid price - fall back to global variable
    GetSpotPrice = g_SpotPrice

    On Error GoTo 0
End Function

Function GetBloombergTicker(underlyingRIC As String) As String
    ' Lookup Bloomberg ticker from RicBloomberg table on Config sheet
    ' Column 1 (N): LSEG underlying RIC (e.g., "ESZ4")
    ' Column 2 (O): Bloomberg ticker (e.g., "ES1 Index")
    Dim wsConfig As Worksheet
    Dim ricBloombergRange As Range
    Dim searchRow As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim foundRIC As String

    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets(SHEET_CONFIG)
    Set ricBloombergRange = wsConfig.Range("RicBloomberg")
    On Error GoTo 0

    ' If range not found, return input as fallback
    If ricBloombergRange Is Nothing Then
        GetBloombergTicker = underlyingRIC
        Exit Function
    End If

    ' Calculate data rows (skip header row 1)
    firstDataRow = ricBloombergRange.Row + 1
    lastDataRow = ricBloombergRange.Row + ricBloombergRange.Rows.count - 1

    ' Search for matching RIC in column 1 of range
    For searchRow = firstDataRow To lastDataRow
        foundRIC = Trim(CStr(wsConfig.Cells(searchRow, ricBloombergRange.Column).Value))

        If foundRIC = underlyingRIC Then
            ' Found match - return Bloomberg ticker from column 2
            GetBloombergTicker = Trim(CStr(wsConfig.Cells(searchRow, ricBloombergRange.Column + 1).Value))

            ' If Bloomberg ticker is empty, return input as fallback
            If GetBloombergTicker = "" Then
                GetBloombergTicker = underlyingRIC
            End If

            Exit Function
        End If
    Next searchRow

    ' Not found in table - return input RIC as fallback
    GetBloombergTicker = underlyingRIC
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

    ws.Range("A1:AB1").Font.Bold = True

    ' Format column A (Spot_Date) as YYYY-MM-DD hh:mm:ss for CSV export
    ws.Columns("A:A").NumberFormat = "yyyy-mm-dd hh:mm:ss"

    ' Format column D (Maturity) as YYYY-MM-DD hh:mm:ss for CSV export
    ws.Columns("D:D").NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Sub ClearStagingSheet()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_STAGING)

    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    ' Clear all data rows (keep header row)
    If lastRow > 1 Then
        ws.Rows("2:" & lastRow).ClearContents
    End If
End Sub

Sub SetupStagingSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STAGING)

    ' Set all column headers including new Greeks - matching CSV export requirements
    ws.Range("A1").Value = "Spot_Date"
    ws.Range("B1").Value = "Premium"
    ws.Range("C1").Value = "Ticker"
    ws.Range("D1").Value = "Maturity"
    ws.Range("E1").Value = "Interest_rate"
    ws.Range("F1").Value = "Spot"
    ws.Range("G1").Value = "Strike"
    ws.Range("H1").Value = "Type"
    ws.Range("I1").Value = "Implied_Volatility"
    ws.Range("J1").Value = "Delta"
    ws.Range("K1").Value = "Vega"
    ws.Range("L1").Value = "Gamma"
    ws.Range("M1").Value = "Theta"
    ws.Range("N1").Value = "Rho"
    ws.Range("O1").Value = "Lot_size"
    ws.Range("P1").Value = "Name"
    ws.Range("Q1").Value = "Reference"
    ws.Range("R1").Value = "ccy_pair"
    ws.Range("S1").Value = "Dividend"
    ws.Range("T1").Value = "DDELTA/DVOL"
    ws.Range("U1").Value = "DDELTA/DVOLDVOL"
    ws.Range("V1").Value = "DDELTA/DTIME"
    ws.Range("W1").Value = "DGAMMA/DSPOT"
    ws.Range("X1").Value = "DGAMMA/DVOL"
    ws.Range("Y1").Value = "DVEGA/DVOL"
    ws.Range("Z1").Value = "DVEGA/DVOLDVOL"

    ws.Range("A1:Z1").Font.Bold = True

    ' Format date columns to YYYY-MM-DD hh:mm:ss for CSV export
    ws.Columns("A:A").NumberFormat = "yyyy-mm-dd hh:mm:ss"  ' Spot_Date
    ws.Columns("D:D").NumberFormat = "yyyy-mm-dd hh:mm:ss"  ' Maturity

    ' Format numeric columns with sufficient decimal places for full precision
    ws.Columns("B:B").NumberFormat = "General"   ' Premium
    ws.Columns("E:E").NumberFormat = "General"   ' Interest_rate
    ws.Columns("F:F").NumberFormat = "General"   ' Spot
    ws.Columns("G:G").NumberFormat = "General"   ' Strike
    ws.Columns("I:N").NumberFormat = "General"   ' Greeks (IV, Delta, Vega, Gamma, Theta, Rho)
    ws.Columns("S:S").NumberFormat = "General"   ' Dividend
    ws.Columns("T:Z").NumberFormat = "General"   ' Higher-order Greeks
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

    While currentCol <= 100 ' Reasonable upper limit
        foundUnderlying = Trim(CStr(wsFuture.Cells(rowDownload, currentCol).Value))   ' Metadata column

        If foundUnderlying <> "" Then
            existingUnderlyings.Add foundUnderlying & " (Col " & (64 + currentCol) & ")", foundUnderlying
        End If

        currentCol = currentCol + 3  ' Move to next 3-column block

        ' Break if we find 3 consecutive empty blocks
        If wsFuture.Cells(2, currentCol).Value = "" And _
           wsFuture.Cells(2, currentCol + 3).Value = "" And _
           wsFuture.Cells(2, currentCol + 6).Value = "" Then
        End If
    Wend

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

Sub ExportToCSV()
    Dim stagingWs As Worksheet
    Dim csvPath As String
    Dim fileName As String

    Set stagingWs = ThisWorkbook.Worksheets(SHEET_STAGING)

    fileName = g_RootRIC & "_" & Format(Date, "yyyymm") & ".csv"
    csvPath = ThisWorkbook.Path & "\" & fileName

    stagingWs.Copy

    ActiveWorkbook.SaveAs fileName:=csvPath, FileFormat:=xlCSV
    ActiveWorkbook.Close False

    MsgBox "Data exported to: " & csvPath, vbInformation
End Sub

' Auto-save staging to CSV after each batch (silent, no popup)
Sub SaveStagingToCSV(Optional batchNumber As Long = 0)
    Dim stagingWs As Worksheet
    Dim csvPath As String
    Dim fileName As String
    Dim rowCount As Long

    On Error GoTo ErrorHandler

    Set stagingWs = ThisWorkbook.Worksheets(SHEET_STAGING)

    ' Check if staging has data (more than just header row)
    rowCount = stagingWs.Cells(stagingWs.Rows.count, 1).End(xlUp).Row
    If rowCount <= 1 Then Exit Sub

    ' Build filename with batch number if provided
    If batchNumber > 0 Then
        fileName = g_RootRIC & "_" & Format(Date, "yyyymmdd_HHmmss") & "_batch" & batchNumber & ".csv"
    Else
        fileName = g_RootRIC & "_" & Format(Date, "yyyymmdd_HHmmss") & ".csv"
    End If

    csvPath = ThisWorkbook.Path & "\" & fileName

    ' Save without prompts
    Application.DisplayAlerts = False
    stagingWs.Copy
    ActiveWorkbook.SaveAs fileName:=csvPath, FileFormat:=xlCSV
    ActiveWorkbook.Close False
    Application.DisplayAlerts = True

    ' Log to status bar instead of popup
    Application.StatusBar = "Auto-saved: " & fileName & " (" & rowCount - 1 & " rows)"

    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    Application.StatusBar = "Error saving CSV: " & Err.Description
End Sub


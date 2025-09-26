Attribute VB_Name = "OptionDownload"
' ============================================
' MODULE 1: Global Configuration and Types
' ============================================

Option Explicit

' Configuration Variables
Public g_UnderlyingTicker As String
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

' Sheet Names
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_RIC_LIST As String = "RIC_List"  ' Now used for progress tracking
Public Const SHEET_COLLECTION As String = "DataCollection"
Public Const SHEET_STAGING As String = "Staging"
Public Const SHEET_QUALITY As String = "QualityReport"
Public Const SHEET_FUTURE As String = "Future et co"

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
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Refreshing LSEG Future data..."
    
    DoEvents
    Application.Run "WorkspaceRefreshWorksheet", True, 120000, SHEET_FUTURE
    DoEvents
    
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox "Double check data in : " & SHEET_FUTURE, vbExclamation
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    MsgBox "Error refreshing data: " & Err.Description, vbExclamation
    Application.StatusBar = False
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
    
    ' Initialize
    InitializeWorkbook
    
    ' Load configuration
    If Not LoadConfiguration() Then
        MsgBox "Please complete configuration in Config sheet", vbExclamation
        Exit Sub
    End If
    
    ' Check if RIC_List has data
    If Not CheckRICListExists() Then
        MsgBox "Please run GenerateAllRICs first to create the RIC list!", vbExclamation
        Exit Sub
    End If
    
    ' Show summary
    response = MsgBox("Starting download for: " & g_UnderlyingTicker & vbNewLine & _
                     "Date Range: " & g_DateStart & " to " & g_DateEnd & vbNewLine & _
                     "RICs to process: " & CountUnprocessedRICs() & vbNewLine & _
                     vbNewLine & "Continue?", vbYesNo + vbQuestion)
    
    If response = vbNo Then Exit Sub
    
    ' Start batch processing
    ProcessAllBatchesFromRICList
    
    ' Generate final report
    GenerateQualityReport
    
    MsgBox "Process Complete! Check Quality Report for summary.", vbInformation
End Sub

Sub ProcessAllBatchesFromRICList()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim batchStart As Long
    Dim batchEnd As Long
    Dim continueProcess As Boolean
    Dim response As Integer
    
    Set ws = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    continueProcess = True
    currentRow = 2  ' Start after header
    
    ' Process in batches
    While currentRow <= lastRow And continueProcess
        ' Find batch of unprocessed RICs
        batchStart = FindNextUnprocessedRIC(currentRow)
        If batchStart = 0 Then Exit Sub  ' No more unprocessed RICs
        
        batchEnd = Application.Min(batchStart + g_BatchSize - 1, lastRow)
        
        ' Get batch info for display
        Dim batchMaturity As Date
        Dim batchType As String
        Dim batchStrikeMin As Double
        Dim batchStrikeMax As Double
        
        batchMaturity = ws.Cells(batchStart, 2).Value  ' Column B: Maturity
        batchType = ws.Cells(batchStart, 4).Value      ' Column D: Type
        batchStrikeMin = ws.Cells(batchStart, 3).Value ' Column C: Strike
        batchStrikeMax = ws.Cells(batchEnd, 3).Value   ' Column C: Strike
        
        ' Show batch details
        response = MsgBox("Process batch:" & vbNewLine & _
                         "Rows: " & batchStart & " to " & batchEnd & vbNewLine & _
                         "Maturity: " & Format(batchMaturity, "mmm-yyyy") & vbNewLine & _
                         "Type: " & batchType & vbNewLine & _
                         "Strikes: " & batchStrikeMin & " to " & batchStrikeMax & vbNewLine & _
                         "RICs: " & (batchEnd - batchStart + 1) & vbNewLine & _
                         vbNewLine & "Continue?", vbYesNo + vbQuestion, "Batch Processing")
        
        If response = vbNo Then
            continueProcess = False
            Exit Sub
        End If
        
        ' Mark batch as processing
        MarkBatchStatus batchStart, batchEnd, "Processing"
        
        ' Process the batch
        ProcessBatchFromRICList batchStart, batchEnd
        
        ' Update to next position
        currentRow = batchEnd + 1
    Wend
End Sub

Sub ProcessBatchFromRICList(startRow As Long, endRow As Long)
    Dim wsRIC As Worksheet
    Dim wsCollection As Worksheet
    Dim i As Long
    Dim ric As String
    Dim currentRow As Long
    Dim formulaCount As Long
    Dim successCount As Long
    Dim errorCount As Long
    Const ROW_SPACING As Long = 300  ' Space between formulas

    Set wsRIC = ThisWorkbook.Worksheets(SHEET_RIC_LIST)
    Set wsCollection = ThisWorkbook.Worksheets(SHEET_COLLECTION)

    ' Clear collection sheet
    ClearCollectionSheet

    ' Setup formulas for this batch with spacing
    currentRow = 2
    formulaCount = 0
    successCount = 0
    errorCount = 0

    For i = startRow To endRow
        ric = wsRIC.Cells(i, 1).Value  ' Column A: RIC

        ' Skip if already processed successfully
        If wsRIC.Cells(i, 8).Value = "Yes" Then  ' Column H: Processed
            GoTo NextRIC
        End If

        ' Calculate row position with spacing
        ' Every formula gets placed 300 rows apart
        currentRow = 2 + (formulaCount * ROW_SPACING)

        ' Setup formula in collection sheet
        wsCollection.Cells(currentRow, 1).Formula = BuildRHistoryFormula(ric, g_DateStart, g_DateEnd)

        ' Store metadata in same row
        wsCollection.Cells(currentRow, 7).Value = wsRIC.Cells(i, 3).Value  ' Strike
        wsCollection.Cells(currentRow, 8).Value = wsRIC.Cells(i, 4).Value  ' Type
        wsCollection.Cells(currentRow, 4).Value = wsRIC.Cells(i, 2).Value  ' Maturity

        ' Store RIC reference for tracking
        wsCollection.Cells(currentRow, 15).Value = i  ' Store row reference in RIC_List

        ' Also store the RIC itself for reference
        wsCollection.Cells(currentRow, 16).Value = ric

        ' Pre-populate Greek formulas for all 300 rows in this section
        ' This ensures formulas are ready before LSEG refresh
        PrePopulateGreekFormulas wsCollection, currentRow, ROW_SPACING, _
                                 wsRIC.Cells(i, 3).Value, _
                                 wsRIC.Cells(i, 4).Value, _
                                 wsRIC.Cells(i, 2).Value, _
                                 i  ' RIC row reference

        formulaCount = formulaCount + 1

NextRIC:
    Next i

    ' Only refresh if there's data to process
    If formulaCount > 0 Then
        ' Refresh LSEG data
        RefreshCollectionSheet

        ' Wait for refresh to complete
        Application.Wait Now + TimeValue("00:00:05")

        ' Process each formula result (they're spaced every 300 rows)
        ' Now we only need to copy rows with actual data
        Dim processRow As Long
        For i = 0 To formulaCount - 1
            processRow = 2 + (i * ROW_SPACING)

            ' Copy only rows that have LSEG data to staging
            CopyDataRowsToStaging wsCollection, processRow, ROW_SPACING
        Next i

        ' Validate and update RIC_List with results
        ValidateAndUpdateRICListWithSpacing wsCollection, formulaCount

        ' Calculate all formulas after everything is set
        Application.Calculate
    End If

    ' Show batch summary
    ShowBatchSummaryFromRICList startRow, endRow
End Sub

' New helper function to pre-populate Greek formulas for all 300 rows
Sub PrePopulateGreekFormulas(ws As Worksheet, startRow As Long, maxRows As Long, _
                             strike As Double, optType As String, maturity As Date, ricRowRef As Long)
    Dim i As Long
    Dim endRow As Long
    Dim callPutFlag As String
    Dim spot As Double
    Dim rate As Double
    Dim timeToExp As Double

    spot = GetSpotPrice()
    rate = GetRiskFreeRate()
    timeToExp = Application.Max((maturity - Date) / 365, 0.001)

    ' Data validation checks
    If timeToExp < 0 Then Exit Sub

    Dim moneyness As Double
    moneyness = strike / spot
    If moneyness <= 0 Or moneyness > 10 Then Exit Sub

    ' Convert P/C to c/p for functions
    If optType = "C" Then
        callPutFlag = "c"
    Else
        callPutFlag = "p"
    End If

    endRow = startRow + maxRows - 1

    For i = startRow To endRow
        ' Store metadata
        ws.Cells(i, 3).Value = g_UnderlyingTicker
        ws.Cells(i, 4).Value = maturity
        ws.Cells(i, 5).Value = rate
        ws.Cells(i, 6).Value = spot
        ws.Cells(i, 7).Value = strike
        ws.Cells(i, 8).Value = optType
        ws.Cells(i, 15).Value = ricRowRef
        ws.Cells(i, 17).Value = g_LotSize
        ws.Cells(i, 18).Value = ""
        ws.Cells(i, 19).Value = ""
        ws.Cells(i, 20).Value = g_Currency
        ws.Cells(i, 21).Value = 0

        ' IV - Column I (9)
        ws.Cells(i, 9).Formula = "=IF(B" & i & "="""",""""," & _
            "GBlackScholesImpVolBisection(""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,B" & i & "))"

        ' Delta - Column J (10)
        ws.Cells(i, 10).Formula = "=IF(B" & i & "="""",""""," & _
            "GBlackScholesNGreeks(""d"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & "))"

        ' Vega - Column K (11)
        ws.Cells(i, 11).Formula = "=IF(B" & i & "="""",""""," & _
            "GBlackScholesNGreeks(""v"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & "))"

        ' Gamma - Column L (12)
        ws.Cells(i, 12).Formula = "=IF(B" & i & "="""",""""," & _
            "GBlackScholesNGreeks(""g"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & "))"

        ' Theta - Column M (13)
        ws.Cells(i, 13).Formula = "=IF(B" & i & "="""",""""," & _
            "GBlackScholesNGreeks(""t"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & "))"

        ' Rho - Column N (14)
        ws.Cells(i, 14).Formula = "=IF(B" & i & "="""",""""," & _
            "GBlackScholesNGreeks(""r"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & "))"

        ' Speed - Column V (22)
        ws.Cells(i, 22).Formula = "=IF(B" & i & "="""",""""," & _
            "CGBlackScholes(""dvdv"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & ",J" & i & "))"

        ' DDELTA/DVOL - Column W (23)
        ws.Cells(i, 23).Formula = "=IF(B" & i & "="""",""""," & _
            "CGBlackScholes(""dddv"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & ",J" & i & "))"

        ' DDELTA/DVOLDVOL - Column X (24)
        ws.Cells(i, 24).Formula = "=IF(B" & i & "="""",""""," & _
            "CGBlackScholes(""dvv"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & ",J" & i & "))"

        ' Charm - Column Y (25)
        ws.Cells(i, 25).Formula = "=IF(B" & i & "="""",""""," & _
            "CGBlackScholes(""dt"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & ",J" & i & "))"

        ' DGamma/DSpot - Column Z (26)
        ws.Cells(i, 26).Formula = "=IF(B" & i & "="""",""""," & _
            "CGBlackScholes(""gps"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & ",J" & i & "))"

        ' Zomma - Column AA (27)
        ws.Cells(i, 27).Formula = "=IF(B" & i & "="""",""""," & _
            "CGBlackScholes(""gpv"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & ",J" & i & "))"

        ' Vomma - Column AB (28)
        ws.Cells(i, 28).Formula = "=IF(B" & i & "="""",""""," & _
            "CGBlackScholes(""dvdv"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & ",J" & i & "))"

        ' Ultima - Column AC (29)
        ws.Cells(i, 29).Formula = "=IF(B" & i & "="""",""""," & _
            "CGBlackScholes(""vvv"",""" & callPutFlag & """," & _
            spot & "," & strike & "," & timeToExp & "," & _
            rate & ",0,I" & i & ",J" & i & "))"
    Next i
End Sub


' New function to copy only rows with LSEG data to staging
Sub CopyDataRowsToStaging(ws As Worksheet, startRow As Long, maxRows As Long)
    Dim i As Long
    Dim endRow As Long
    Dim wsDest As Worksheet
    Dim nextRow As Long

    Set wsDest = ThisWorkbook.Worksheets(SHEET_STAGING)

    endRow = startRow + maxRows - 1

    ' Copy only rows that have premium data
    For i = startRow To endRow
        If Not IsEmpty(ws.Cells(i, 2).Value) And IsNumeric(ws.Cells(i, 2).Value) Then
            ' This row has LSEG data, copy it to staging with proper column mapping
            nextRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).Row + 1

            ' Map columns to staging sheet (matching CSV export format)
            wsDest.Cells(nextRow, 1).Value = ws.Cells(i, 1).Value   ' Spot_Date
            wsDest.Cells(nextRow, 2).Value = ws.Cells(i, 2).Value   ' Premium
            wsDest.Cells(nextRow, 3).Value = ws.Cells(i, 3).Value   ' Ticker
            wsDest.Cells(nextRow, 4).Value = ws.Cells(i, 4).Value   ' Maturity
            wsDest.Cells(nextRow, 5).Value = ws.Cells(i, 5).Value   ' Interest_rate
            wsDest.Cells(nextRow, 6).Value = ws.Cells(i, 6).Value   ' Spot
            wsDest.Cells(nextRow, 7).Value = ws.Cells(i, 7).Value   ' Strike
            wsDest.Cells(nextRow, 8).Value = ws.Cells(i, 8).Value   ' Type
            wsDest.Cells(nextRow, 9).Value = ws.Cells(i, 9).Value   ' Implied_Volatility
            wsDest.Cells(nextRow, 10).Value = ws.Cells(i, 10).Value ' Delta
            wsDest.Cells(nextRow, 11).Value = ws.Cells(i, 11).Value ' Vega
            wsDest.Cells(nextRow, 12).Value = ws.Cells(i, 12).Value ' Gamma
            wsDest.Cells(nextRow, 13).Value = ws.Cells(i, 13).Value ' Theta
            wsDest.Cells(nextRow, 14).Value = ws.Cells(i, 14).Value ' Rho
            wsDest.Cells(nextRow, 15).Value = ws.Cells(i, 17).Value ' Lot_size (from col Q)
            wsDest.Cells(nextRow, 16).Value = ws.Cells(i, 18).Value ' Name
            wsDest.Cells(nextRow, 17).Value = ws.Cells(i, 19).Value ' Reference
            wsDest.Cells(nextRow, 18).Value = ws.Cells(i, 20).Value ' ccy_pair
            wsDest.Cells(nextRow, 19).Value = ws.Cells(i, 21).Value ' Dividend
            wsDest.Cells(nextRow, 20).Value = ws.Cells(i, 22).Value ' DDELTA_DSPOT (from col V)
            wsDest.Cells(nextRow, 21).Value = ws.Cells(i, 23).Value ' DDELTA_DVOL (from col W)
            wsDest.Cells(nextRow, 22).Value = ws.Cells(i, 24).Value ' DDELTA_DVOLDVOL (from col X)
            wsDest.Cells(nextRow, 23).Value = ws.Cells(i, 25).Value ' DDELTA_DTIME (from col Y)
            wsDest.Cells(nextRow, 24).Value = ws.Cells(i, 26).Value ' DGAMMA_DSPOT (from col Z)
            wsDest.Cells(nextRow, 25).Value = ws.Cells(i, 27).Value ' DGAMMA_DVOL (from col AA)
            wsDest.Cells(nextRow, 26).Value = ws.Cells(i, 28).Value ' DVEGA_DVOL (from col AB)
            wsDest.Cells(nextRow, 27).Value = ws.Cells(i, 29).Value ' DVEGA_DVOLDVOL (from col AC)
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
                                                 GetSpotPrice(), wsRIC.Cells(ricRow, 2).Value)
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
            .Range("G1").Value = "Check Existence"
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
                                                 GetSpotPrice(), wsRIC.Cells(ricRow, 2).Value)
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
    
    MsgBox "Batch Complete!" & vbNewLine & vbNewLine & _
          "Rows processed: " & startRow & " to " & endRow & vbNewLine & _
          "Successful: " & successCount & vbNewLine & _
          "Errors: " & errorCount & vbNewLine & _
          "Skipped: " & (endRow - startRow + 1 - successCount - errorCount), _
          vbInformation, "Batch Summary"
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
                          """.Timestamp;.Close"",""START:" & startNum & _
                          " END:" & endNum & " INTERVAL:1D"")"
End Function

' ============================================
' Keep existing refresh and calculation functions
' ============================================

Sub RefreshCollectionSheet()
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Refreshing LSEG data..."
    
    DoEvents
    Application.Run "WorkspaceRefreshWorksheet", True, 120000, SHEET_COLLECTION
    DoEvents
    
    Application.Wait Now + TimeValue("00:00:02")
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    MsgBox "Error refreshing data: " & Err.Description, vbExclamation
    Application.StatusBar = False
End Sub

Sub CalculateGreeks(ws As Worksheet, startRow As Long, endRow As Long)
    Dim i As Long
    Dim spotDate As Variant
    Dim premium As Variant
    Dim strike As Double
    Dim spot As Double
    Dim rate As Double
    Dim maturity As Date
    Dim optType As String
    Dim timeToExp As Double
    
    spot = GetSpotPrice()
    rate = GetRiskFreeRate()
    
    For i = startRow To endRow
        spotDate = ws.Cells(i, 1).Value
        premium = ws.Cells(i, 2).Value
        strike = ws.Cells(i, 7).Value
        maturity = ws.Cells(i, 4).Value
        optType = ws.Cells(i, 8).Value
        
        If Not IsEmpty(premium) And IsNumeric(premium) Then
            timeToExp = Application.Max((maturity - Date) / 365, 0.001)
            
            ws.Cells(i, 5).Value = rate
            ws.Cells(i, 6).Value = spot
            
            ' Calculate IV
            ws.Cells(i, 9).Formula = "=IF(B" & i & "="""",""""," & _
                "GBlackScholesImpVolBisection(""" & optType & """," & _
                strike & "," & spot & "," & timeToExp & "," & _
                rate & ",0," & premium & "))"
            
            ' Calculate Delta
            ws.Cells(i, 10).Formula = "=IF(B" & i & "="""",""""," & _
                "GBlackScholesNGreeks(""Delta""," & strike & "," & _
                spot & "," & rate & "," & timeToExp & ",0,I" & i & "))"
        End If
    Next i
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
    ws.Range("A3").Value = "Underlying: " & g_UnderlyingTicker
    
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
    
    g_UnderlyingTicker = ws.Range("underlyingTicker").Value
    g_RootRIC = ws.Range("rootRIC").Value
    g_StrikeStep = ws.Range("steps").Value
    g_LotSize = ws.Range("lotSize").Value
    g_Currency = ws.Range("currency").Value
    g_DateStart = ws.Range("dateStart").Value
    g_DateEnd = ws.Range("dateEnd").Value
    g_BatchSize = ws.Range("batchSize").Value
    
    g_PutStrikeMin = ws.Range("minStrikePut").Value
    g_PutStrikeMax = ws.Range("maxStrikePut").Value
    g_CallStrikeMin = ws.Range("minStrikeCall").Value
    g_CallStrikeMax = ws.Range("maxStrikeCall").Value

    LoadConfiguration = True
    Exit Function
    
ErrorHandler:
    LoadConfiguration = False
End Function

Function GetSpotPrice() As Double
    On Error Resume Next
    GetSpotPrice = ThisWorkbook.Worksheets(SHEET_FUTURE).Range("B2").Value
    If GetSpotPrice = 0 Then GetSpotPrice = g_SpotPrice
    On Error GoTo 0
End Function

Function GetRiskFreeRate() As Double
    On Error Resume Next
    GetRiskFreeRate = ThisWorkbook.Worksheets(SHEET_FUTURE).Range("E2").Value
    If GetRiskFreeRate = 0 Then GetRiskFreeRate = 0.04
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
    ws.Range("V1").Value = "DDELTA_DSPOT"
    ws.Range("W1").Value = "DDELTA_DVOL"
    ws.Range("X1").Value = "DDELTA_DVOLDVOL"
    ws.Range("Y1").Value = "DDELTA_DTIME"
    ws.Range("Z1").Value = "DGAMMA_DSPOT"
    ws.Range("AA1").Value = "DGAMMA_DVOL"
    ws.Range("AB1").Value = "DVEGA_DVOL"
    ws.Range("AC1").Value = "DVEGA_DVOLDVOL"

    ws.Range("A1:AC1").Font.Bold = True
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
    ws.Range("T1").Value = "DDELTA_DSPOT"
    ws.Range("U1").Value = "DDELTA_DVOL"
    ws.Range("V1").Value = "DDELTA_DVOLDVOL"
    ws.Range("W1").Value = "DDELTA_DTIME"
    ws.Range("X1").Value = "DGAMMA_DSPOT"
    ws.Range("Y1").Value = "DGAMMA_DVOL"
    ws.Range("Z1").Value = "DVEGA_DVOL"
    ws.Range("AA1").Value = "DVEGA_DVOLDVOL"

    ws.Range("A1:AA1").Font.Bold = True
End Sub

Sub SetupQualitySheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_QUALITY)
    ws.Cells.Clear
End Sub

Sub ExportToCSV()
    Dim stagingWs As Worksheet
    Dim csvPath As String
    Dim fileName As String
    
    Set stagingWs = ThisWorkbook.Worksheets(SHEET_STAGING)
    
    fileName = g_UnderlyingTicker & "_" & Format(Date, "yyyymm") & ".csv"
    csvPath = ThisWorkbook.Path & "\" & fileName
    
    stagingWs.Copy
    
    ActiveWorkbook.SaveAs fileName:=csvPath, FileFormat:=xlCSV
    ActiveWorkbook.Close False
    
    MsgBox "Data exported to: " & csvPath, vbInformation
End Sub




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

' ============================================
' MAIN RIC GENERATION FUNCTION
' ============================================
Sub GenerateAllRICs()
    ' Main function to generate all RICs and display in new sheet
    Dim ricList As Collection
    Dim outputSheet As Worksheet
    Dim ricDict As Object  ' Dictionary instead of custom type
    Dim i As Long
    Dim lastRow As Long
    
    ' Generate all RICs
    Set ricList = BuildCompleteRICList()
    
    ' Check if sheet exists, create if it doesn't
    On Error Resume Next
    Set outputSheet = ThisWorkbook.Sheets(SHEET_RIC_LIST)
    On Error GoTo 0
    
    If outputSheet Is Nothing Then
        ' Create new sheet if it doesn't exist
        Set outputSheet = ThisWorkbook.Sheets.Add
        outputSheet.Name = SHEET_RIC_LIST
    Else
        ' Clear existing content in columns A to N
        lastRow = outputSheet.Cells(outputSheet.Rows.count, "A").End(xlUp).Row
        If lastRow > 0 Then
            outputSheet.Range("A1:N" & lastRow).Clear
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
        .Range("G1").Value = "Check Existence"
        .Range("H1").Value = "Processed"  ' New column for tracking processing status
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(200, 200, 200)
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
            ' Add RData formula to check if RIC exists
            .Cells(i, 7).Formula = "=@TR(A" & i & ",""UNDERLYING"")"
            ' Initialize Processed column as "No" or empty
            .Cells(i, 8).Value = "No"
        End With
        i = i + 1
    Next
    
    ' Format
    outputSheet.Columns("A:H").AutoFit  ' Updated to include column H
    outputSheet.Range("B:B").NumberFormat = "mm/dd/yyyy"
    outputSheet.Range("C:C").NumberFormat = "#,##0"
    
    ' Add conditional formatting to Processed column for visual feedback
    With outputSheet.Range("H2:H" & ricList.count + 1).FormatConditions
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
    
    
    ' Refresh the chain data
    DoEvents
    Application.Run "WorkspaceRefreshWorksheet", True, 120000, SHEET_RIC_LIST
    DoEvents
    
    MsgBox "Generated " & ricList.count & " RICs!" & vbNewLine & _
           "Check '" & SHEET_RIC_LIST & "' sheet for details." & vbNewLine & _
           "Column G will show maturity dates for valid RICs after LSEG refresh." & vbNewLine & _
           "Column H tracks processing status.", vbInformation
End Sub

' ============================================
' BUILD COMPLETE RIC LIST
' ============================================

Function BuildCompleteRICList() As Collection
    Dim ricList As New Collection
    Dim maturities As Collection
    Dim putStrikes As Collection
    Dim callStrikes As Collection
    Dim maturity As Variant
    Dim strike As Variant
    Dim ricInfo As Object  ' Dictionary
    
    Set maturities = GetMaturityDates()
    Set putStrikes = GetStrikeRange("PUT")
    Set callStrikes = GetStrikeRange("CALL")
    
    ' Generate PUT RICs
    For Each maturity In maturities
        For Each strike In putStrikes
            Set ricInfo = CreateRICInfo(CDate(maturity), CDbl(strike), "PUT")
            ricList.Add ricInfo
        Next strike
    Next maturity
    
    ' Generate CALL RICs
    For Each maturity In maturities
        For Each strike In callStrikes
            Set ricInfo = CreateRICInfo(CDate(maturity), CDbl(strike), "CALL")
            ricList.Add ricInfo
        Next strike
    Next maturity
    
    Set BuildCompleteRICList = ricList
End Function

' ============================================
' CREATE INDIVIDUAL RIC (Returns Dictionary)
' ============================================

Function CreateRICInfo(maturityDate As Date, strike As Double, optionType As String) As Object
    Dim rootRIC As String
    Dim monthCode As String
    Dim monthCodeCallForExpiredRIC As String
    Dim yearCode As String
    Dim strikeStr As String
    Dim ricDict As Object
    
    ' Create dictionary to hold RIC information
    Set ricDict = CreateObject("Scripting.Dictionary")
    
    ' Get values
    rootRIC = ThisWorkbook.Sheets(SHEET_CONFIG).Range("rootRIC").Value
    monthCode = GetMonthCodeFromTable(Month(maturityDate), optionType)
    monthCodeCallForExpiredRIC = GetMonthCodeFromTable(Month(maturityDate), "CALL")
    yearCode = Right(CStr(Year(maturityDate)), 2)
    strikeStr = FormatStrikeForRIC(strike)
    
    ' Populate dictionary
    ricDict.Add "FullRIC", BuildRICString(rootRIC, strikeStr, monthCode, yearCode, maturityDate, monthCodeCallForExpiredRIC)
    ricDict.Add "Maturity", maturityDate
    ricDict.Add "Strike", strike
    ricDict.Add "OptionType", optionType
    ricDict.Add "MonthCode", monthCode
    ricDict.Add "YearCode", yearCode
    
    Set CreateRICInfo = ricDict
End Function

' ============================================
' BUILD RIC STRING
' ============================================

Function BuildRICString(rootRIC As String, strikeStr As String, monthCode As String, yearCode As String, maturityDate As Date, monthCodeCallForExpiredRIC As String) As String
    ' Builds the complete RIC string
    ' Format example: 1EW7000T25 for current/future options
    '                 1EW7000T25^T25 for expired options
    
    ' Basic format
    BuildRICString = rootRIC & strikeStr & monthCode & yearCode
    
    ' Add ^{monthCode}{yearCode} suffix if maturity date is before today (expired option)
    If maturityDate < Date Then
        BuildRICString = BuildRICString & "^" & monthCodeCallForExpiredRIC & yearCode
    End If
End Function

' ============================================
' FORMAT STRIKE FOR RIC
' ============================================

Function FormatStrikeForRIC(strike As Double) As String
    Dim strikeStr As String
    
    If strike = Int(strike) Then
        strikeStr = CStr(Int(strike))
    Else
        strikeStr = Replace(CStr(strike), ".", "")
    End If
    
    FormatStrikeForRIC = strikeStr
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

' ============================================
' UTILITY FUNCTIONS FOR WORKING WITH RIC DICTIONARIES
' ============================================

Sub TestRICGeneration()
    ' Test function to generate a few sample RICs
    Dim testDate As Date
    Dim ricDict As Object
    
    ' Test date (use first maturity from Config)
    On Error Resume Next
    testDate = ThisWorkbook.Sheets(SHEET_CONFIG).Range("maturityDate").Cells(1, 1).Value
    On Error GoTo 0
    
    If Not IsDate(testDate) Then
        MsgBox "Please add a maturity date in Config sheet", vbExclamation
        Exit Sub
    End If
    
    ' Generate sample RICs
    Debug.Print "Sample RIC Generation Test:"
    Debug.Print "============================"
    
    ' Test PUT
    Set ricDict = CreateRICInfo(testDate, 6000, "PUT")
    Debug.Print "PUT 6000: " & ricDict("FullRIC")
    Debug.Print "  Maturity: " & ricDict("Maturity")
    Debug.Print "  Month Code: " & ricDict("MonthCode")
    
    Set ricDict = CreateRICInfo(testDate, 6500, "PUT")
    Debug.Print "PUT 6500: " & ricDict("FullRIC")
    
    ' Test CALL
    Set ricDict = CreateRICInfo(testDate, 7000, "CALL")
    Debug.Print "CALL 7000: " & ricDict("FullRIC")
    Debug.Print "  Maturity: " & ricDict("Maturity")
    Debug.Print "  Month Code: " & ricDict("MonthCode")
    
    Set ricDict = CreateRICInfo(testDate, 7500, "CALL")
    Debug.Print "CALL 7500: " & ricDict("FullRIC")
    
    MsgBox "Test complete! Check Immediate Window (Ctrl+G) for results", vbInformation
End Sub

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
' EXAMPLE: USE LOOKUP DICTIONARY
' ============================================

Sub ExampleUseLookupDictionary()
    Dim lookupDict As Object
    Dim searchKey As String
    Dim foundRIC As String
    
    ' Build lookup dictionary once
    Set lookupDict = BuildRICLookupDictionary()
    
    ' Example lookup
    searchKey = "7000_20251017_PUT"  ' Strike_YYYYMMDD_Type
    
    If lookupDict.Exists(searchKey) Then
        foundRIC = lookupDict(searchKey)
        Debug.Print "Found RIC: " & foundRIC
    Else
        Debug.Print "RIC not found for key: " & searchKey
    End If
End Sub

' ============================================
' DOWNLOAD FROM OPTION CHAIN
' ============================================

Sub DownloadFromChain()
    ' Downloads option chain from LSEG and populates RIC_List sheet
    Dim ws As Worksheet
    Dim chainSheet As Worksheet
    Dim ricListSheet As Worksheet
    Dim rootRIC As String
    Dim chainRIC As String
    Dim lastRow As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    ' Get root RIC from config
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    rootRIC = Trim(ws.Range("rootRIC").Value)

    If rootRIC = "" Then
        MsgBox "Please specify root RIC in Config sheet!", vbExclamation
        Exit Sub
    End If

    ' Create chain RIC for option chain download
    chainRIC = "0#" & rootRIC & "+"

    ' Create or clear chain download sheet
    Set chainSheet = CreateChainDownloadSheet()

    ' Setup chain download formula
    chainSheet.Range("A1").Value = "Chain RIC"
    chainSheet.Range("B1").Value = "Chain Data"
    chainSheet.Range("C1").Value = "Status"
    chainSheet.Range("A1:E1").Font.Bold = True

    ' Add chain download formula
    chainSheet.Range("A2").Value = chainRIC
    chainSheet.Range("B2").Formula = "=@TR(""" & chainRIC & """,""CF_NAME"",""CH=Fd RH=IN"")"
    chainSheet.Range("C2").Value = "Downloading..."

    Application.StatusBar = "Downloading option chain of chains for " & rootRIC & "..."

    ' Refresh the chain data
    DoEvents
    Application.Run "WorkspaceRefreshWorksheet", True, 120000, chainSheet.Name
    DoEvents

    ' Wait for refresh to complete
    Application.Wait Now + TimeValue("00:00:05")

    ' Check if data was downloaded
    If IsEmpty(chainSheet.Range("B3").Value) Or chainSheet.Range("B3").Value = "0" Then
        chainSheet.Range("C2").Value = "No data"
        MsgBox "No option chain data found for " & rootRIC & ". Please check if the root RIC is correct.", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

    chainSheet.Range("C2").Value = "Processing..."

    ' Process the downloaded chain data
    ProcessChainData chainSheet

    Application.StatusBar = False
    MsgBox "Option chain download complete! Check " & SHEET_RIC_LIST & " sheet for results.", vbInformation
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error in DownloadFromChain: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number, vbCritical
    If Not chainSheet Is Nothing Then
        chainSheet.Range("C2").Value = "Error: " & Err.Description
    End If
End Sub

' ============================================
' CREATE CHAIN DOWNLOAD SHEET
' ============================================

Function CreateChainDownloadSheet() As Worksheet
    Dim ws As Worksheet
    Dim sheetName As String

    sheetName = "ChainDownload"

    ' Delete existing sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create new sheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = sheetName

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
    DoEvents
    Application.Run "WorkspaceRefreshWorksheet", True, 120000, chainSheet.Name
    DoEvents

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
        .Range("G1").Value = "Check Existence"
        .Range("H1").Value = "Processed"
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(200, 200, 200)
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

    ' Find last row with data in this column
    lastRow = chainSheet.Cells(chainSheet.Rows.count, optionColumn).End(xlUp).Row

    ' Process each row in this column starting from row 2
    For i = 4 To lastRow
        optionDataText = Trim(CStr(chainSheet.Cells(i, optionColumn).Value))

        ' Skip if no data
        If optionDataText = "" Or optionDataText = "0" Then GoTo NextOptionRow
        
        ' Populate RIC_List row
        With ricListSheet
            .Cells(ricListRow, 1).Value = chainSheet.Cells(i, optionColumn).Value  ' RIC
            .Cells(ricListRow, 2).Value = chainSheet.Cells(i, optionColumn + 3).Value ' Maturity
            .Cells(ricListRow, 3).Value = chainSheet.Cells(i, optionColumn + 2).Value ' Strike
            .Cells(ricListRow, 4).Value = chainSheet.Cells(i, optionColumn + 4).Value ' Type
            .Cells(ricListRow, 5).Value = "n/a" ' Month Code
            .Cells(ricListRow, 6).Value = "n/a"  ' Year
            .Cells(ricListRow, 7).Value = chainSheet.Cells(i, optionColumn + 5).Value ' Underlying
            .Cells(ricListRow, 8).Value = "No"  ' Processed
        End With
        ricListRow = ricListRow + 1

NextOptionRow:
    Next i
End Sub


' ============================================
' FORMAT RIC LIST SHEET
' ============================================

Sub FormatRICListSheet(ricListSheet As Worksheet, ricListRow As Long)
    With ricListSheet
        .Columns("A:H").AutoFit
        .Range("B:B").NumberFormat = "mm/dd/yyyy"
        .Range("C:C").NumberFormat = "#,##0"

        ' Add conditional formatting
        If ricListRow > 2 Then
            With .Range("H2:H" & ricListRow - 1).FormatConditions
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







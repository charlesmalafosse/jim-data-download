Attribute VB_Name = "LSEGCalc"
Option Explicit


Sub RefreshLSEGWithTimeout(ws As Worksheet, Optional timeoutSeconds As Long = 120)
    Dim startTime As Double
    Dim originalCalcMode As XlCalculation
    
    startTime = Timer
    
    ' Store original calculation mode
    originalCalcMode = Application.Calculation
    
    Application.StatusBar = "Refreshing LSEG data for " & ws.Name & "..."
    DoEvents
    
    ' Clear any pending operations
    'Application.SendKeys "{ESC}"
    Application.Wait Now + TimeValue("0:00:01")
    DoEvents  ' Ensure events are processed after wait
    
    On Error GoTo RefreshError
    
    ' Set to manual calculation for the refresh
    Application.Calculation = xlCalculationManual
    
    ' Attempt the refresh
    Application.Run "WorkspaceRefreshWorksheet", True, timeoutSeconds * 1000, ws.Name

    ' Wait briefly for LSEG to start populating data
    Application.Wait Now + TimeValue("0:00:02")
    DoEvents

    ' Poll for completion using cell text inspection (more reliable than CalculationState)
    Dim readyCount As Long
    Dim totalCount As Long
    Dim checkCounter As Long
    checkCounter = 0

    Do While Not IsLSEGDataReady(ws, readyCount, totalCount, 10)
        DoEvents
        checkCounter = checkCounter + 1

        If Timer - startTime > timeoutSeconds Then
            Application.StatusBar = "Refresh timeout for " & ws.Name & " after " & Format(Timer - startTime, "0") & " seconds - forcing completion..."
            Application.SendKeys "{ESC}"
            DoEvents
            Exit Do
        End If

        ' Update status bar with elapsed time and progress
        Application.StatusBar = "Refreshing " & ws.Name & "... " & _
                               Format(Timer - startTime, "0") & "s - " & _
                               readyCount & " of " & totalCount & " cells ready"

        ' Check every 2 seconds to reduce CPU usage
        If checkCounter Mod 5 = 0 Then  ' Every 5th iteration
            Application.Wait Now + TimeValue("0:00:02")
        Else
            Application.Wait Now + TimeValue("0:00:01")  ' 1 second pause
        End If
    Loop
    
    ' Restore original calculation mode
    Application.Calculation = originalCalcMode
    
    ' Force recalculation of the worksheet
    'Application.StatusBar = "Calculating " & ws.Name & "..."
    'DoEvents
    'ws.Calculate
    
    Application.StatusBar = ws.Name & " refresh and calculation completed"
    Exit Sub
    
RefreshError:
    'Application.SendKeys "{ESC}"
    DoEvents
    
    Application.StatusBar = "Error during refresh of " & ws.Name & ": " & Err.Description
    
    ' Restore original calculation mode even on error
    Application.Calculation = originalCalcMode
    
    ' Still try to calculate the worksheet even if refresh had issues
    On Error Resume Next
    'ws.Calculate
    On Error GoTo 0
    
    Application.StatusBar = ws.Name & " refresh interrupted (calculation attempted)"
End Sub

' ============================================
' Helper function to check if LSEG data has finished downloading
' ============================================
Function IsLSEGDataReady(ws As Worksheet, Optional ByRef readyCount As Long, Optional ByRef totalCount As Long, Optional sampleSize As Long = 10) As Boolean
    Dim checkRow As Long
    Dim cellText As String
    Dim cellsReady As Long
    Dim cellsChecked As Long
    Dim lastRow As Long
    Dim checkInterval As Long
    Dim i As Long

    cellsReady = 0
    cellsChecked = 0

    ' Find last row with data in column A (date column from RHistory)
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    ' If no data or only header, consider it "ready" (nothing to wait for)
    If lastRow <= 1 Then
        readyCount = 0
        totalCount = 0
        IsLSEGDataReady = True
        Exit Function
    End If

    ' Calculate interval to sample cells evenly across the data range
    ' Check column B (Premium/Close price column) where LSEG data appears
    Dim rowsToCheck As Long
    rowsToCheck = Application.Min(sampleSize, lastRow - 1)

    If rowsToCheck <= 1 Then
        checkInterval = 1
    Else
        checkInterval = Application.Max(1, (lastRow - 2) \ (rowsToCheck - 1))
    End If

    ' Sample cells from beginning, middle, and end
    For i = 1 To rowsToCheck
        If i = 1 Then
            checkRow = 2  ' First data row
        ElseIf i = rowsToCheck Then
            checkRow = lastRow  ' Last data row
        Else
            checkRow = 2 + ((i - 1) * checkInterval)
        End If

        ' Don't check beyond last row
        If checkRow > lastRow Then checkRow = lastRow

        cellsChecked = cellsChecked + 1
        cellText = CStr(ws.Cells(checkRow, 2).Text)

        ' Check if cell is ready (no longer showing LSEG status messages)
        ' LSEG shows "Retrieving...", "#N/A Requesting Data...", or similar during download
        If InStr(1, cellText, "Retrieving", vbTextCompare) = 0 And _
           InStr(1, cellText, "Requesting", vbTextCompare) = 0 And _
           InStr(1, cellText, "Loading", vbTextCompare) = 0 Then
            cellsReady = cellsReady + 1
        End If
    Next i

    ' Return progress info
    readyCount = cellsReady
    totalCount = cellsChecked

    ' Consider ready if ALL checked cells are no longer retrieving
    IsLSEGDataReady = (cellsChecked > 0 And cellsReady = cellsChecked)
End Function




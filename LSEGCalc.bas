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
    
    ' Poll for completion with escape mechanism
    Do While Application.CalculationState <> xlDone
        DoEvents
        
        If Timer - startTime > timeoutSeconds Then
            Application.StatusBar = "Refresh timeout for " & ws.Name & " - forcing completion..."
            Application.SendKeys "{ESC}"
            DoEvents
            Exit Do
        End If
        
        ' Update status bar with elapsed time
        Application.StatusBar = "Refreshing " & ws.Name & "... " & _
                               Format(Timer - startTime, "0") & " seconds"
    Loop
    
    ' Restore original calculation mode
    Application.Calculation = originalCalcMode
    
    ' Force recalculation of the worksheet
    Application.StatusBar = "Calculating " & ws.Name & "..."
    DoEvents
    ws.Calculate
    
    Application.StatusBar = ws.Name & " refresh and calculation completed"
    Exit Sub
    
RefreshError:
    'Application.SendKeys "{ESC}"
    DoEvents
    
    ' Restore original calculation mode even on error
    Application.Calculation = originalCalcMode
    
    ' Still try to calculate the worksheet even if refresh had issues
    On Error Resume Next
    ws.Calculate
    On Error GoTo 0
    
    Application.StatusBar = ws.Name & " refresh interrupted (calculation attempted)"
End Sub


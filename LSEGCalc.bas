Attribute VB_Name = "LSEGCalc"
Option Explicit

Sub RefreshLSEGWithTimeout(ws As Worksheet, Optional timeoutSeconds As Long = 120)
    Dim startTime As Double
    
    startTime = Timer
    
    Application.StatusBar = "Refreshing LSEG data for " & ws.Name & "..."
    DoEvents
    
    ' Clear any pending operations
    Application.SendKeys "{ESC}"
    Application.Wait Now + TimeValue("0:00:01")
    
    On Error GoTo RefreshError
    
    ' Start refresh in a way we can interrupt
    Application.Calculation = xlCalculationManual
    
    ' Attempt the refresh
    Application.Run "WorkspaceRefreshWorksheet", True, timeoutSeconds * 1000, ws
    
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
    
    'Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = ws.Name & " refresh completed"
    Exit Sub
    
RefreshError:
    Application.SendKeys "{ESC}"
    'Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = ws.Name & " refresh interrupted"
End Sub


Attribute VB_Name = "SetupConfig"
' ============================================
' EXCEL WORKBOOK SETUP - RUN THIS ONCE
' This will create all sheets with proper formatting
' ============================================

' Sheet Name Constants
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_PROGRESS As String = "Progress"
Public Const SHEET_COLLECTION As String = "DataCollection"
Public Const SHEET_STAGING As String = "Staging"
Public Const SHEET_QUALITY As String = "QualityReport"
Public Const SHEET_FUTURE As String = "Future et co"
Public Const SHEET_MAIN As String = "Main"


Sub SetupCompleteWorkbook()
    Application.ScreenUpdating = False
    
    ' Create all sheets
    CreateConfigSheet
    CreateMainSheet
    CreateFutureSheet
    CreateProgressSheet
    CreateDataCollectionSheet
    CreateStagingSheet
    CreateQualityReportSheet
    CreateOptionDataSheets
    
    ' Set workbook properties
    ThisWorkbook.Sheets(SHEET_CONFIG).Activate
    
    Application.ScreenUpdating = True
    
    MsgBox "Workbook setup complete!" & vbNewLine & vbNewLine & _
           "Next steps:" & vbNewLine & _
           "1. Fill in Configuration sheet" & vbNewLine & _
           "2. Add maturity dates in Config sheet (starting row 25)" & vbNewLine & _
           "3. Update Future et co sheet with underlying data" & vbNewLine & _
           "4. Run 'MainDownloadProcess' to start downloading", vbInformation
End Sub

' ============================================
' CONFIG SHEET SETUP
' ============================================
Sub CreateConfigSheet()
    Dim ws As Worksheet
    
    ' Delete if exists and create new
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Config").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Config"
    
    ' Headers and labels
    With ws
        ' Main configuration
        .Range("A1").Value = "Configuration Parameter"
        .Range("B1").Value = "Value"
        .Range("C1").Value = "Description"
        
        .Range("A2").Value = "Underlying Ticker"
        .Range("B2").Value = "ES"
        .Range("C2").Value = "Bloomberg ticker"
        
        .Range("A3").Value = "Root RIC"
        .Range("B3").Value = "1EW"
        .Range("C3").Value = "LSEG root RIC for options"
        
        .Range("A4").Value = "Current Spot Price"
        .Range("B4").Value = 6500
        .Range("C4").Value = "Current underlying price"
        
        .Range("A5").Value = "Strike Step Size"
        .Range("B5").Value = 10
        .Range("C5").Value = "Step between strikes"
        
        .Range("A6").Value = "Lot Size"
        .Range("B6").Value = 50
        .Range("C6").Value = "Contract multiplier"
        
        .Range("A7").Value = "Currency"
        .Range("B7").Value = "USD"
        .Range("C7").Value = "Currency code"
        
        .Range("A8").Value = "Date Start"
        .Range("B8").Value = DateSerial(Year(Date), Month(Date) - 1, 1)
        .Range("C8").Value = "Start date for price data"
        
        .Range("A9").Value = "Date End"
        .Range("B9").Value = DateSerial(Year(Date), Month(Date), 0)
        .Range("C9").Value = "End date for price data"
        
        .Range("A10").Value = "Batch Size"
        .Range("B10").Value = 20
        .Range("C10").Value = "Number of strikes per batch"
        
        ' Strike ranges
        .Range("A12").Value = "Strike Ranges (Auto-calculated or override)"
        .Range("A12").Font.Bold = True
        
        .Range("A13").Value = "Put Strike Min"
        .Range("B13").Formula = "=ROUND(B4*0.5/B5,0)*B5"
        .Range("C13").Value = "Minimum put strike"
        
        .Range("A14").Value = "Put Strike Max"
        .Range("B14").Formula = "=ROUND(B4*0.9/B5,0)*B5"
        .Range("C14").Value = "Maximum put strike"
        
        .Range("A15").Value = "Call Strike Min"
        .Range("B15").Formula = "=ROUND(B4*0.9/B5,0)*B5"
        .Range("C15").Value = "Minimum call strike"
        
        .Range("A16").Value = "Call Strike Max"
        .Range("B16").Formula = "=ROUND(B4*1.5/B5,0)*B5"
        .Range("C16").Value = "Maximum call strike"
        
        ' Month codes reference - moved to columns E-G to avoid overlap
        .Range("E20").Value = "Month Codes Reference"
        .Range("E20").Font.Bold = True
        
        .Range("E21").Value = "Month"
        .Range("F21").Value = "Call Code"
        .Range("G21").Value = "Put Code"
        
        Dim i As Integer
        Dim callCodes As Variant
        Dim putCodes As Variant
        callCodes = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L")
        putCodes = Array("M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X")
        
        For i = 1 To 12
            .Cells(21 + i, 5).Value = i  ' Column E
            .Cells(21 + i, 6).Value = callCodes(i - 1)  ' Column F
            .Cells(21 + i, 7).Value = putCodes(i - 1)  ' Column G
        Next i
        
        ' Maturity dates section - now safe at A24
        .Range("A20").Value = "Maturity Dates to Process"
        .Range("A20").Font.Bold = True
        .Range("A21").Value = "Maturity Date"
        .Range("B21").Value = "Status"
        
        ' Add sample maturities (3rd Friday of next 3 months)
        For i = 1 To 3
            .Cells(21 + i, 1).Value = GetThirdFriday(DateAdd("m", i, Date))
            .Cells(21 + i, 2).Value = "Pending"
        Next i
        
        ' Formatting
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.Color = RGB(200, 200, 200)
        .Range("E21:G21").Font.Bold = True
        .Range("E21:G21").Interior.Color = RGB(220, 220, 220)
        .Range("A21:B21").Font.Bold = True
        .Range("A21:B21").Interior.Color = RGB(220, 220, 220)
        
        ' Column widths
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 30
        .Columns("E:G").ColumnWidth = 12
        
        ' Number formats
        .Range("B4").NumberFormat = "#,##0"
        .Range("B5").NumberFormat = "0"
        .Range("B6").NumberFormat = "0"
        .Range("B8:B9").NumberFormat = "mm/dd/yyyy"
        .Range("B13:B16").NumberFormat = "#,##0"
        .Range("A22:A40").NumberFormat = "mm/dd/yyyy"
        
        ' Borders
        .Range("A1:C16").Borders.LineStyle = xlContinuous
        .Range("E18:G31").Borders.LineStyle = xlContinuous
        .Range("A20:B40").Borders.LineStyle = xlContinuous
    End With
End Sub

' ============================================
' MAIN SHEET SETUP (Legacy compatibility)
' ============================================
Sub CreateMainSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Main").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Main"
    
    With ws
        ' Basic info that matches old structure
        .Range("A1").Value = "Parameter"
        .Range("B1").Value = "Value"
        
        .Range("A2").Value = "Bloomberg Ticker"
        .Range("B2").Value = "ES"
        
        .Range("A3").Value = "Root RIC"
        .Range("B3").Value = "1EW"
        
        .Range("A4").Value = "Current Maturity"
        .Range("B4").Formula = "=Config!A22"
        
        .Range("A5").Value = "Start Strike"
        .Range("B5").Value = 6000
        
        .Range("A8").Value = "Reference"
        .Range("B8").Value = "EMINI"
        
        .Range("A13").Value = "Start Date"
        .Range("B13").Formula = "=Config!B8"
        
        .Range("A14").Value = "End Date"
        .Range("B14").Formula = "=Config!B9"
        
        .Range("A15").Value = "Lot Size"
        .Range("B15").Formula = "=Config!B6"
        
        .Range("A16").Value = "Name"
        .Range("B16").Value = "S&P 500 E-mini"
        
        .Range("A17").Value = "Currency"
        .Range("B17").Formula = "=Config!B7"
        
        ' Format
        .Range("A1:B1").Font.Bold = True
        .Range("A1:B1").Interior.Color = RGB(200, 200, 200)
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 20
    End With
End Sub

' ============================================
' FUTURE ET CO SHEET SETUP
' ============================================
Sub CreateFutureSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Future et co").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Future et co"
    
    With ws
        ' Headers
        .Range("A1").Value = "Date"
        .Range("B1").Value = "Spot"
        .Range("C1").Value = "Dividend Yield"
        .Range("D1").Value = "Days to Expiry"
        .Range("E1").Value = "Risk Free Rate"
        
        ' Sample formulas for underlying data
        .Range("A2").Formula = "=TODAY()"
        
        ' LSEG formula for spot price (adjust RIC as needed)
        .Range("B2").Formula = "=@RHistory(""ES"",""TRDPRC_1"",""START:""&Config!B8&"" END:""&Config!B9&"" INTERVAL:1D"")"
        
        .Range("C2").Value = 0  ' No dividend for futures
        .Range("D2").Value = 30  ' Default days
        .Range("E2").Value = 0.045  ' Default 4.5% risk free rate
        
        ' Additional rows for different dates if needed
        Dim i As Integer
        For i = 3 To 32  ' One month of data
            .Cells(i, 1).Formula = "=TODAY()-" & (i - 2)
            .Cells(i, 2).Value = ""  ' Will be populated by LSEG
            .Cells(i, 3).Value = 0
            .Cells(i, 4).Value = 30
            .Cells(i, 5).Value = 0.045
        Next i
        
        ' Format
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(200, 200, 200)
        .Columns("A").ColumnWidth = 12
        .Columns("B:E").ColumnWidth = 15
        .Range("A:A").NumberFormat = "mm/dd/yyyy"
        .Range("B:B").NumberFormat = "#,##0.00"
        .Range("C:C").NumberFormat = "0.00%"
        .Range("E:E").NumberFormat = "0.00%"
        
        ' Borders
        .Range("A1:E32").Borders.LineStyle = xlContinuous
    End With
End Sub

' ============================================
' PROGRESS SHEET SETUP
' ============================================
Sub CreateProgressSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Progress").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Progress"
    
    With ws
        .Range("A1").Value = "Maturity"
        .Range("B1").Value = "Type"
        .Range("C1").Value = "Strike_Start"
        .Range("D1").Value = "Strike_End"
        .Range("E1").Value = "Status"
        .Range("F1").Value = "Records_Found"
        .Range("G1").Value = "Errors"
        .Range("H1").Value = "Last_Updated"
        
        ' Format
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(150, 200, 150)
        .Columns("A").NumberFormat = "mm/dd/yyyy"
        .Columns("C:D").NumberFormat = "#,##0"
        .Columns("H").NumberFormat = "mm/dd/yyyy hh:mm"
        .Columns("A:H").AutoFit
        
        ' Add conditional formatting for status
        With .Range("E:E").FormatConditions
            .Add Type:=xlTextString, String:="Complete", TextOperator:=xlContains
            .Item(.count).Interior.Color = RGB(200, 255, 200)
            
            .Add Type:=xlTextString, String:="Running", TextOperator:=xlContains
            .Item(.count).Interior.Color = RGB(255, 255, 200)
            
            .Add Type:=xlTextString, String:="Error", TextOperator:=xlContains
            .Item(.count).Interior.Color = RGB(255, 200, 200)
        End With
    End With
End Sub

' ============================================
' DATA COLLECTION SHEET SETUP
' ============================================
Sub CreateDataCollectionSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("DataCollection").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "DataCollection"
    
    With ws
        ' Headers matching your CSV format
        .Range("A1").Value = "Spot_Date"
        .Range("B1").Value = "Premium"
        .Range("C1").Value = "Ticker"
        .Range("D1").Value = "Maturity"
        .Range("E1").Value = "Interest_rate"
        .Range("F1").Value = "Spot"
        .Range("G1").Value = "Strike"
        .Range("H1").Value = "Type"
        .Range("I1").Value = "Implied_Volatility"
        .Range("J1").Value = "Delta"
        .Range("K1").Value = "Vega"
        .Range("L1").Value = "Gamma"
        .Range("M1").Value = "Theta"
        .Range("N1").Value = "Rho"
        .Range("O1").Value = "Lot_size"
        .Range("P1").Value = "Name"
        .Range("Q1").Value = "Reference"
        .Range("R1").Value = "ccy_pair"
        .Range("S1").Value = "Dividend"
        .Range("T1").Value = "DDELTA/DSPOT"
        .Range("U1").Value = "DDELTA/DVOL"
        .Range("V1").Value = "DDELTA/DVOLDVOL"
        .Range("W1").Value = "DDELTA/DTIME"
        .Range("X1").Value = "DGAMMA/DSPOT"
        .Range("Y1").Value = "DGAMMA/DVOL"
        .Range("Z1").Value = "DVEGA/DVOL"
        .Range("AA1").Value = "DVEGA/DVOLDVOL"
        
        ' Format headers
        .Range("A1:AA1").Font.Bold = True
        .Range("A1:AA1").Interior.Color = RGB(200, 200, 255)
        .Range("A1:AA1").WrapText = True
        .Rows(1).RowHeight = 30
        
        ' Column formats
        .Columns("A").NumberFormat = "mm/dd/yyyy"
        .Columns("B").NumberFormat = "#,##0.0000"
        .Columns("D").NumberFormat = "mm/dd/yyyy"
        .Columns("E").NumberFormat = "0.00%"
        .Columns("F:G").NumberFormat = "#,##0.00"
        .Columns("I:AA").NumberFormat = "0.0000"
        
        ' Set column widths
        .Columns("A:AA").ColumnWidth = 12
    End With
End Sub

' ============================================
' STAGING SHEET SETUP
' ============================================
Sub CreateStagingSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Staging").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Staging"
    
    ' Copy structure from DataCollection
    ThisWorkbook.Sheets("DataCollection").Range("A1:AA1").Copy
    ws.Range("A1").PasteSpecial xlPasteAll
    
    ' Add validation column
    ws.Range("AB1").Value = "Validation_Status"
    ws.Range("AB1").Font.Bold = True
    ws.Range("AB1").Interior.Color = RGB(200, 200, 255)
    
    ' Apply same formatting
    ws.Columns("A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B").NumberFormat = "#,##0.0000"
    ws.Columns("D").NumberFormat = "mm/dd/yyyy"
    ws.Columns("E").NumberFormat = "0.00%"
    ws.Columns("F:G").NumberFormat = "#,##0.00"
    ws.Columns("I:AA").NumberFormat = "0.0000"
    ws.Columns("A:AB").ColumnWidth = 12
End Sub

' ============================================
' QUALITY REPORT SHEET SETUP
' ============================================
Sub CreateQualityReportSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("QualityReport").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "QualityReport"
    
    With ws
        .Range("A1").Value = "Option Data Quality Report"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Report Generated:"
        .Range("B3").Formula = "=NOW()"
        .Range("B3").NumberFormat = "mm/dd/yyyy hh:mm"
        
        .Range("A4").Value = "Underlying:"
        .Range("B4").Formula = "=Config!B2"
        
        ' Summary section
        .Range("A6").Value = "SUMMARY STATISTICS"
        .Range("A6").Font.Bold = True
        .Range("A6").Font.Size = 12
        
        .Range("A7").Value = "Total Options Processed:"
        .Range("B7").Formula = "=COUNTA(Staging!A:A)-1"
        
        .Range("A8").Value = "Valid Data (OK):"
        .Range("B8").Formula = "=COUNTIF(Staging!AB:AB,""OK"")"
        
        .Range("A9").Value = "High IV Warnings:"
        .Range("B9").Formula = "=COUNTIF(Staging!AB:AB,""High"")"
        
        .Range("A10").Value = "Low IV Warnings:"
        .Range("B10").Formula = "=COUNTIF(Staging!AB:AB,""Low"")"
        
        .Range("A11").Value = "Missing Data:"
        .Range("B11").Formula = "=COUNTIF(Staging!AB:AB,""Missing"")"
        
        .Range("A12").Value = "Success Rate:"
        .Range("B12").Formula = "=IF(B7>0,B8/B7,0)"
        .Range("B12").NumberFormat = "0.00%"
        
        ' Coverage Matrix section
        .Range("A14").Value = "COVERAGE MATRIX"
        .Range("A14").Font.Bold = True
        .Range("A14").Font.Size = 12
        
        .Range("A15").Value = "Maturity"
        .Range("B15").Value = "Puts Found"
        .Range("C15").Value = "Calls Found"
        .Range("D15").Value = "Total"
        .Range("E15").Value = "ATM Coverage"
        
        .Range("A15:E15").Font.Bold = True
        .Range("A15:E15").Interior.Color = RGB(220, 220, 220)
        
        ' Issue Log section
        .Range("A25").Value = "ISSUE LOG"
        .Range("A25").Font.Bold = True
        .Range("A25").Font.Size = 12
        
        .Range("A26").Value = "Timestamp"
        .Range("B26").Value = "Maturity"
        .Range("C26").Value = "Strike"
        .Range("D26").Value = "Type"
        .Range("E26").Value = "Issue"
        
        .Range("A26:E26").Font.Bold = True
        .Range("A26:E26").Interior.Color = RGB(255, 220, 220)
        
        ' Format
        .Columns("A").ColumnWidth = 25
        .Columns("B:E").ColumnWidth = 15
        .Range("A7:B12").Borders.LineStyle = xlContinuous
    End With
End Sub

' ============================================
' OPTION DATA SHEETS (P1, P2, P3, C1, C2, C3)
' ============================================
Sub CreateOptionDataSheets()
    Dim sheetNames As Variant
    Dim i As Integer
    Dim ws As Worksheet
    
    sheetNames = Array("P1", "P2", "P3", "C1", "C2", "C3")
    
    For i = 0 To UBound(sheetNames)
        On Error Resume Next
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(sheetNames(i)).Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
        
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetNames(i)
        
        With ws
            ' Setup for LSEG formulas
            .Range("A1").Value = "LSEG Formula Area"
            .Range("A1").Font.Bold = True
            .Range("A1").Interior.Color = RGB(255, 255, 200)
            
            ' Configuration area
            .Range("H20").Value = "Configuration"
            .Range("H20").Font.Bold = True
            
            .Range("H21").Value = "Current RIC:"
            .Range("I21").Value = ""  ' Will be populated by VBA
            
            .Range("B21").Value = "Current Strike:"
            .Range("C21").Value = ""  ' Will be populated by VBA
            
            .Range("B4").Value = "Maturity:"
            .Range("C4").Value = ""  ' Will be populated by VBA
            
            ' Reference cells for formulas
            .Range("C8").Formula = "=Main!B8"  ' Reference
            .Range("C13").Formula = "=Main!B13"  ' Start date
            .Range("C14").Formula = "=Main!B14"  ' End date
            .Range("C15").Formula = "=Main!B15"  ' Lot size
            .Range("C16").Formula = "=Main!B16"  ' Name
            .Range("C17").Formula = "=Main!B17"  ' Currency
            
            ' Headers for data (starting at row 40)
            .Range("A40").Value = "Spot_Date"
            .Range("B40").Value = "Premium"
            .Range("C40").Value = "Ticker"
            .Range("D40").Value = "Maturity"
            .Range("E40").Value = "Interest_rate"
            .Range("F40").Value = "Spot"
            .Range("G40").Value = "Strike"
            .Range("H40").Value = "Type"
            .Range("I40").Value = "Implied_Volatility"
            .Range("J40").Value = "Delta"
            .Range("K40").Value = "Vega"
            .Range("L40").Value = "Gamma"
            .Range("M40").Value = "Theta"
            .Range("N40").Value = "Rho"
            
            ' Headers for higher order greeks
            .Range("U37").Value = "DDELTA/DSPOT"
            .Range("V37").Value = "DDELTA/DVOL"
            .Range("W37").Value = "DDELTA/DVOLDVOL"
            .Range("X37").Value = "DDELTA/DTIME"
            .Range("Y37").Value = "DGAMMA/DSPOT"
            .Range("Z37").Value = "DGAMMA/DVOL"
            .Range("AA37").Value = "DVEGA/DVOL"
            .Range("AB37").Value = "DVEGA/DVOLDVOL"
            
            .Range("A40:N40").Font.Bold = True
            .Range("A40:N40").Interior.Color = RGB(220, 220, 220)
            
            ' Setup areas for multiple formulas (every 300 rows as per your spec)
            Dim formulaRow As Long
            For formulaRow = 41 To 1241 Step 300
                .Cells(formulaRow, 1).Value = "' RHistory formula will go here"
                .Cells(formulaRow, 1).Interior.Color = RGB(255, 255, 220)
            Next formulaRow
            
            ' Format
            .Columns("A:N").ColumnWidth = 12
        End With
    Next i
End Sub

' ============================================
' HELPER FUNCTIONS
' ============================================

Function GetThirdFriday(monthDate As Date) As Date
    ' Get the third Friday of the month (typical option expiry)
    Dim firstDay As Date
    Dim firstFriday As Date
    Dim dayNum As Integer
    
    firstDay = DateSerial(Year(monthDate), Month(monthDate), 1)
    dayNum = Weekday(firstDay)
    
    ' Find first Friday
    If dayNum <= 6 Then
        firstFriday = firstDay + (6 - dayNum)
    Else
        firstFriday = firstDay + (13 - dayNum)
    End If
    
    ' Add two weeks to get third Friday
    GetThirdFriday = firstFriday + 14
End Function

' ============================================
' SAMPLE DATA POPULATION
' ============================================

Sub PopulateSampleData()
    ' This sub adds sample data to test the system
    Dim ws As Worksheet
    
    ' Add sample config data
    Set ws = ThisWorkbook.Sheets("Config")
    
    ' Sample underlying: E-mini S&P 500
    ws.Range("B2").Value = "ES"  ' Ticker
    ws.Range("B3").Value = "1EW"  ' Root RIC
    ws.Range("B4").Value = 6500   ' Spot price
    ws.Range("B5").Value = 10     ' Strike step
    ws.Range("B6").Value = 50     ' Lot size
    ws.Range("B7").Value = "USD"  ' Currency
    ws.Range("B10").Value = 20    ' Batch size
    
    ' Add sample spot data to Future sheet
    Set ws = ThisWorkbook.Sheets("Future et co")
    ws.Range("B2").Value = 6500   ' Current spot
    ws.Range("E2").Value = 0.045  ' Risk free rate
    
    MsgBox "Sample data populated!" & vbNewLine & _
           "You can now test the system with MainDownloadProcess", vbInformation
End Sub

' ============================================
' VALIDATION SETUP
' ============================================

Sub SetupDataValidation()
    Dim ws As Worksheet
    
    ' Add data validation to Config sheet
    Set ws = ThisWorkbook.Sheets("Config")
    
    ' Validation for currency
    With ws.Range("B7").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="USD,EUR,GBP,JPY,CHF,CAD,AUD"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' Validation for batch size
    With ws.Range("B10").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="5", Formula2:="100"
        .InputMessage = "Enter batch size between 5 and 100"
        .ErrorMessage = "Batch size must be between 5 and 100"
    End With
    
    ' Validation for Progress sheet status
    Set ws = ThisWorkbook.Sheets("Progress")
    With ws.Range("E:E").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="Pending,Running,Complete,Error,Skipped"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

' ============================================
' MAIN SETUP ENTRY POINT
' ============================================

Sub SETUP_NEW_WORKBOOK()
    Dim response As Integer
    
    response = MsgBox("This will create/recreate all sheets for the Option Download system." & vbNewLine & _
                     vbNewLine & "Existing sheets will be replaced!" & vbNewLine & _
                     vbNewLine & "Continue?", vbYesNo + vbExclamation, "Setup Workbook")
    
    If response = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Run complete setup
    SetupCompleteWorkbook
    
    ' Add data validation
    SetupDataValidation
    
    ' Optional: Add sample data
    response = MsgBox("Would you like to populate sample data for testing?", vbYesNo + vbQuestion)
    If response = vbYes Then
        PopulateSampleData
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Show instructions
    MsgBox "Setup Complete!" & vbNewLine & vbNewLine & _
           "NEXT STEPS:" & vbNewLine & _
           "1. Review and adjust configuration in 'Config' sheet" & vbNewLine & _
           "2. Ensure LSEG Excel Add-in is connected" & vbNewLine & _
           "3. Add/modify maturity dates in Config sheet (row 25+)" & vbNewLine & _
           "4. Run 'MainDownloadProcess' from the previous VBA code" & vbNewLine & _
           vbNewLine & _
           "The workbook is now ready for option data download!", vbInformation, "Setup Complete"
End Sub

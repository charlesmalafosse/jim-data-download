# LSEG VBA Non-Blocking Download System

## Table of Contents
1. [Problem Statement](#problem-statement)
2. [Architecture Overview](#architecture-overview)
3. [Core Components](#core-components)
4. [Implementation Details](#implementation-details)
5. [Usage Examples](#usage-examples)
6. [Key Techniques](#key-techniques)
7. [Best Practices](#best-practices)

---

## Problem Statement

### The Challenge
When downloading data from LSEG (London Stock Exchange Group) using their Excel VBA add-in, the following issues occur:

1. **Excel Freezing**: LSEG's `WorkspaceRefreshWorksheet` function can freeze Excel during data retrieval
2. **Blocking Operations**: VBA execution blocks while waiting for LSEG formulas (e.g., `RHistory`) to populate
3. **No Progress Feedback**: Users have no visibility into download progress or ability to cancel
4. **Timeout Issues**: Long-running downloads with no timeout mechanism
5. **Batch Processing**: Need to process hundreds of option RICs without freezing Excel

### The Solution
A multi-layered architecture that:
- Returns control to Excel during LSEG data retrieval
- Implements non-blocking asynchronous batch processing
- Provides timeout mechanisms and cancellation capability
- Uses state machine pattern for complex multi-phase operations
- Monitors data readiness without blocking

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                     APPLICATION LAYER                            │
│  (MainDownloadProcess, GenerateAllRICs, RefreshFutureSheet)    │
└────────────────────┬────────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────────┐
│                  BATCH PROCESSING LAYER                          │
│         (OnTime Chain - Asynchronous State Machine)             │
│                                                                  │
│  ┌──────────────┐    ┌───────────────┐    ┌─────────────────┐ │
│  │  Setup       │───▶│  Check        │───▶│  Process        │ │
│  │  Formulas    │    │  Refresh      │    │  Results        │ │
│  └──────────────┘    └───────────────┘    └─────────────────┘ │
│         │                  │ ▲                      │           │
│         │                  │ │ Polling Loop         │           │
│         │                  └─┘ (OnTime reschedule)  │           │
│         │                                            │           │
│         └────────────────────────────────────────────┴──────────┤
│                            ▼                                     │
│                    ┌──────────────┐                             │
│                    │  Trigger     │                             │
│                    │  Next/Done   │                             │
│                    └──────────────┘                             │
└────────────────────┬────────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────────┐
│                   CORE REFRESH LAYER                             │
│              (LSEGCalc.RefreshLSEGWithTimeout)                  │
│                                                                  │
│  • DoEvents for responsiveness                                  │
│  • Timeout monitoring with Timer                                │
│  • CalculationState polling                                     │
│  • ESC key cancellation                                         │
└────────────────────┬────────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────────┐
│                    LSEG ADD-IN LAYER                             │
│          (WorkspaceRefreshWorksheet, RHistory)                  │
└─────────────────────────────────────────────────────────────────┘
```

---

## Core Components

### 1. LSEGCalc.bas - Non-Blocking Refresh Function

**Purpose**: Wraps LSEG's `WorkspaceRefreshWorksheet` with timeout and non-blocking capabilities.

**File**: `LSEGCalc.bas`

```vba
Sub RefreshLSEGWithTimeout(ws As Worksheet, Optional timeoutSeconds As Long = 120)
    Dim startTime As Double
    Dim originalCalcMode As XlCalculation

    startTime = Timer
    originalCalcMode = Application.Calculation

    Application.StatusBar = "Refreshing LSEG data for " & ws.Name & "..."
    DoEvents

    ' Clear any pending operations and wait briefly
    Application.Wait Now + TimeValue("0:00:01")
    DoEvents

    On Error GoTo RefreshError

    ' Set to manual calculation for the refresh
    Application.Calculation = xlCalculationManual

    ' Attempt the refresh
    Application.Run "WorkspaceRefreshWorksheet", True, timeoutSeconds * 1000, ws.Name

    ' Poll for completion with escape mechanism
    Do While Application.CalculationState <> xlDone
        DoEvents  ' ← KEY: Return control to Excel

        ' Check timeout
        If Timer - startTime > timeoutSeconds Then
            Application.StatusBar = "Refresh timeout for " & ws.Name & " - forcing completion..."
            Application.SendKeys "{ESC}"  ' ← Force cancellation
            DoEvents
            Exit Do
        End If

        ' Update status bar with elapsed time
        Application.StatusBar = "Refreshing " & ws.Name & "... " & _
                               Format(Timer - startTime, "0") & " seconds"
    Loop

    ' Restore original calculation mode
    Application.Calculation = originalCalcMode

    Application.StatusBar = ws.Name & " refresh and calculation completed"
    Exit Sub

RefreshError:
    DoEvents
    Application.StatusBar = "Error during refresh of " & ws.Name & ": " & Err.Description
    Application.Calculation = originalCalcMode
    Application.StatusBar = ws.Name & " refresh interrupted (calculation attempted)"
End Sub
```

**Key Features**:
- `DoEvents` in polling loop - Returns control to Excel UI
- `Timer` function - Tracks elapsed time for timeout
- `Application.CalculationState` - Monitors if Excel is still calculating
- `Application.SendKeys "{ESC}"` - Force-cancels if timeout exceeded
- `Application.Calculation` mode management - Prevents unwanted recalcs
- Error handling with proper cleanup

---

### 2. OptionDownload.bas - Asynchronous Batch Processing

**Purpose**: Process hundreds of option RICs in batches without blocking Excel.

**File**: `OptionDownload.bas`

#### State Machine Enumeration

```vba
Public Enum BatchProcessState
    bpsIdle = 0
    bpsSetupFormulas = 1
    bpsRefreshing = 2
    bpsProcessingResults = 3
End Enum

' Global state variables
Public g_BatchState As BatchProcessState
Public g_BatchStartRow As Long
Public g_BatchEndRow As Long
Public g_FormulaCount As Long
Public g_StopRequested As Boolean
Public g_NextScheduledProc As String
Public g_RefreshCheckCount As Long
```

#### Phase 1: Setup Formulas

```vba
Sub ProcessBatch_SetupFormulas()
    Const ROW_SPACING As Long = 300

    ' Check stop flag
    If g_StopRequested Then
        ProcessBatch_Abort
        Exit Sub
    End If

    g_BatchState = bpsSetupFormulas

    ' Clear sheets
    ClearCollectionSheet
    ClearStagingSheet

    ' Setup RHistory formulas for batch
    Dim currentRow As Long
    g_FormulaCount = 0

    For i = g_BatchStartRow To g_BatchEndRow
        ric = wsRIC.Cells(i, 1).Value

        If wsRIC.Cells(i, 8).Value = "Yes" Then GoTo NextRIC

        currentRow = 2 + (g_FormulaCount * ROW_SPACING)

        ' Create RHistory formula
        wsCollection.Cells(currentRow, 1).Formula = BuildRHistoryFormula(ric, g_DateStart, g_DateEnd)

        ' Store metadata
        SetupRHistoryAndMetadata wsCollection, currentRow, ROW_SPACING, ...

        g_FormulaCount = g_FormulaCount + 1
NextRIC:
    Next i

    ' Only proceed if there's data
    If g_FormulaCount > 0 Then
        g_BatchState = bpsRefreshing
        g_RefreshCheckCount = 0

        ' Trigger LSEG refresh
        RefreshLSEGCollectionSheet  ' ← Uses LSEGCalc.RefreshLSEGWithTimeout

        ' Schedule check after 5 seconds (OnTime breaks execution chain)
        g_NextScheduledProc = "ProcessBatch_CheckRefresh"
        Application.OnTime Now + TimeValue("00:00:05"), g_NextScheduledProc
    Else
        ProcessBatch_TriggerNext
    End If
End Sub
```

**Key Points**:
- Formulas created but **not yet calculated**
- `Application.OnTime` schedules next phase **5 seconds later**
- VBA execution **ends here** - Excel regains control
- LSEG refreshes formulas asynchronously

---

#### Phase 2: Check Refresh Status

```vba
Sub ProcessBatch_CheckRefresh()
    ' Check stop flag
    If g_StopRequested Then
        ProcessBatch_Abort
        Exit Sub
    End If

    ' Check timeout (60 checks × 3 sec = 3 min timeout)
    g_RefreshCheckCount = g_RefreshCheckCount + 1
    If g_RefreshCheckCount > 60 Then
        MsgBox "LSEG refresh timeout for batch - proceeding anyway", vbExclamation
        ProcessBatch_ProcessResults
        Exit Sub
    End If

    Application.StatusBar = "Batch #" & g_BatchCounter & ": Checking refresh status (attempt " & g_RefreshCheckCount & ")..."

    ' Check if data ready
    If IsDataReady(wsCollection) Then
        ' Data ready, proceed to processing
        ProcessBatch_ProcessResults
    Else
        ' Still waiting, reschedule check in 3 seconds
        g_NextScheduledProc = "ProcessBatch_CheckRefresh"
        Application.OnTime Now + TimeValue("00:00:03"), g_NextScheduledProc
    End If
End Sub
```

**Key Points**:
- **Polling pattern** - Checks if LSEG data is ready
- Reschedules itself with `OnTime` if not ready (every 3 seconds)
- Timeout protection (3-minute max wait)
- Non-blocking - returns control between checks

---

#### Data Readiness Check

```vba
Function IsDataReady(ws As Worksheet) As Boolean
    Dim checkRow As Long
    Dim readyCount As Long
    Dim totalChecks As Long
    Dim cellText As String
    Const ROW_SPACING As Long = 300

    totalChecks = 0
    readyCount = 0

    ' Check first few formulas (max 5 samples)
    For checkRow = 2 To 2 + (g_FormulaCount * ROW_SPACING) Step ROW_SPACING
        totalChecks = totalChecks + 1

        cellText = CStr(ws.Cells(checkRow, 2).Text)

        ' Check if cell is ready (no longer shows "Retrieving...")
        If InStr(1, cellText, "Retrieving...", vbTextCompare) = 0 Then
            readyCount = readyCount + 1
        End If

        If totalChecks >= 5 Then Exit For
    Next

    ' Consider ready if ALL checked cells are no longer refreshing
    IsDataReady = (totalChecks > 0 And readyCount = totalChecks)
End Function
```

**Key Insight**: LSEG shows "Retrieving..." in cells while loading. We check cell `.Text` property to detect this.

---

#### Phase 3: Process Results

```vba
Sub ProcessBatch_ProcessResults()
    If g_StopRequested Then
        ProcessBatch_Abort
        Exit Sub
    End If

    g_BatchState = bpsProcessingResults

    ' Add Greek formulas to rows with data (after LSEG refresh)
    For i = 0 To g_FormulaCount - 1
        processRow = 2 + (i * ROW_SPACING)
        AddGreekFormulasToDataRows wsCollection, processRow, ROW_SPACING
    Next i

    ' Calculate Greeks
    wsCollection.Calculate

    ' Wait for calculation to complete
    Do While Application.CalculationState <> xlDone
        DoEvents  ' ← Non-blocking wait
        Application.Wait Now + TimeValue("00:00:01")
    Loop

    ' Validate and copy to staging
    ValidateAndUpdateRICListWithSpacing wsCollection, g_FormulaCount

    ' Save batch to CSV
    SaveStagingToCSV g_BatchCounter

    ' Auto-save workbook every 3 batches
    If g_BatchCounter Mod 3 = 0 Then
        ThisWorkbook.Save
    End If

    ShowBatchSummaryFromRICList g_BatchStartRow, g_BatchEndRow

    ' Trigger next batch
    ProcessBatch_TriggerNext
End Sub
```

**Key Points**:
- Greeks calculated **only after** LSEG data arrives
- Another non-blocking wait for calculation state
- Batch results saved incrementally (CSV per batch)
- Auto-save to prevent data loss

---

#### Phase 4: Trigger Next Batch

```vba
Sub ProcessBatch_TriggerNext()
    Dim nextStart As Long
    Dim nextEnd As Long

    ' Find next unprocessed batch
    nextStart = FindNextUnprocessedRIC(g_BatchEndRow + 1)

    If nextStart > 0 And nextStart <= lastRow Then
        ' Found next batch
        nextEnd = Application.Min(nextStart + g_BatchSize - 1, lastRow)
        g_BatchCounter = g_BatchCounter + 1

        g_BatchStartRow = nextStart
        g_BatchEndRow = nextEnd

        MarkBatchStatus nextStart, nextEnd, "Processing"

        ' Schedule next batch (2-second delay)
        g_NextScheduledProc = "ProcessBatch_SetupFormulas"
        Application.OnTime Now + TimeValue("00:00:02"), g_NextScheduledProc
    Else
        ' All done
        ProcessBatch_Complete
    End If
End Sub
```

**Key Points**:
- Finds next unprocessed batch from RIC_List tracking sheet
- Schedules next batch with `OnTime` (2-second delay)
- Loops back to Phase 1 for next batch
- Calls completion handler when done

---

#### Process Entry Point

```vba
Sub MainDownloadProcess()
    ' Initialize and validate configuration
    InitializeWorkbook
    LoadConfiguration
    CheckRICListExists
    CheckUnderlyings

    ' Initialize state
    g_BatchCounter = 0
    g_StopRequested = False
    g_BatchState = bpsIdle

    ' Find first unprocessed batch
    batchStart = FindNextUnprocessedRIC(2)
    batchEnd = Application.Min(batchStart + g_BatchSize - 1, lastRow)

    g_BatchStartRow = batchStart
    g_BatchEndRow = batchEnd
    g_BatchCounter = g_BatchCounter + 1

    ' Start the OnTime chain
    ProcessBatch_SetupFormulas

    ' VBA execution ends here - OnTime chain runs asynchronously
End Sub
```

**Critical**: MainDownloadProcess initiates the chain but does **not wait** for completion. The OnTime chain runs asynchronously.

---

### 3. RICconfiguration.bas - RIC Generation

**Purpose**: Generate list of option RICs and refresh underlying data.

**File**: `RICconfiguration.bas`

```vba
Sub GenerateAllRICs()
    ' Generate all option RICs
    Set ricList = BuildCompleteRICList()

    ' Create/clear RIC_List sheet
    Set outputSheet = ThisWorkbook.Sheets(SHEET_RIC_LIST)

    ' Output RICs with formulas
    For Each ric In ricList
        outputSheet.Cells(i, 1).Value = ricDict("FullRIC")
        outputSheet.Cells(i, 2).Value = ricDict("Maturity")
        ' ... more columns ...

        ' Add formula to check underlying (TR function)
        outputSheet.Cells(i, 7).Formula = "=@TR(A" & i & ",""UNDERLYING"")"
        outputSheet.Cells(i, 8).Value = "No"  ' Processing status
    Next

    ' Refresh LSEG data using non-blocking function
    RefreshLSEGWithTimeout wsRIC, 120  ' ← 2-minute timeout

    MsgBox "Generated " & ricList.Count & " RICs!"
End Sub
```

**Key Points**:
- Uses `RefreshLSEGWithTimeout` for non-blocking refresh
- TR formula validates RIC existence
- Processed column tracks download status

---

## Key Techniques

### 1. DoEvents - Return Control to Excel

**What**: VBA command that yields execution to Excel to process events.

**Why**: Prevents Excel from freezing during long-running operations.

**Where Used**:
- Inside polling loops (`RefreshLSEGWithTimeout:31`, `OptionDownload:562`)
- After status bar updates
- During calculation waits

**Example**:
```vba
Do While Application.CalculationState <> xlDone
    DoEvents  ' ← Let Excel process UI events, user clicks, etc.
    Application.Wait Now + TimeValue("00:00:01")
Loop
```

**Caution**: DoEvents allows users to trigger other VBA code. Protect against re-entrancy with state flags (`g_BatchState`).

---

### 2. Application.OnTime - Break Execution Chain

**What**: Schedules VBA procedure to run at a future time, then **ends current execution**.

**Why**:
- Breaks VBA execution chain, allowing LSEG async operations to complete
- Prevents VBA from blocking while LSEG populates formulas
- Enables true asynchronous batch processing

**Where Used**:
- Between batch processing phases (`OptionDownload:490`, `OptionDownload:528`, `OptionDownload:623`)

**Example**:
```vba
' Setup formulas (Phase 1)
wsCollection.Cells(row, 1).Formula = BuildRHistoryFormula(ric, ...)

' Trigger LSEG refresh
RefreshLSEGWithTimeout wsCollection, 60

' Schedule check in 5 seconds, then END EXECUTION
Application.OnTime Now + TimeValue("00:00:05"), "ProcessBatch_CheckRefresh"
' ← VBA ends here. Excel regains control. LSEG continues populating.
```

**Flow**:
```
VBA starts → Setup formulas → Trigger LSEG → Schedule OnTime → VBA ENDS
                                                    ↓
                                          (5 seconds later)
                                                    ↓
VBA resumes → Check if ready → If yes: process → If no: reschedule OnTime → VBA ENDS
```

---

### 3. Application.CalculationState - Monitor Calculation

**What**: Property indicating if Excel is currently calculating.

**Values**:
- `xlDone` - Calculation complete
- `xlCalculating` - Currently calculating
- `xlPending` - Calculation queued

**Where Used**:
- Polling loops to wait for calculation completion
- `RefreshLSEGWithTimeout:31`
- `OptionDownload:562`

**Example**:
```vba
ws.Calculate  ' Trigger calculation

' Wait for completion
Do While Application.CalculationState <> xlDone
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
Loop
```

---

### 4. Timer Function - Timeout Tracking

**What**: Returns seconds elapsed since midnight.

**Why**: Implement timeout protection for long-running operations.

**Where Used**:
- `RefreshLSEGWithTimeout:6` - Start time
- `RefreshLSEGWithTimeout:34` - Elapsed time check

**Example**:
```vba
Dim startTime As Double
startTime = Timer

Do While Application.CalculationState <> xlDone
    DoEvents

    If Timer - startTime > timeoutSeconds Then
        Application.SendKeys "{ESC}"  ' Force cancel
        Exit Do
    End If
Loop
```

**Edge Case**: Timer resets at midnight. For operations spanning midnight, use alternative like `Now` with date comparison.

---

### 5. Application.SendKeys "{ESC}" - Force Cancellation

**What**: Simulates pressing ESC key.

**Why**: Force-cancels LSEG operations that exceed timeout.

**Where Used**:
- `RefreshLSEGWithTimeout:36`

**Caution**:
- `SendKeys` is unreliable in some environments
- Currently commented out in production code (`LSEGCalc:18`, `LSEGCalc:58`)
- Use as last resort

---

### 6. Calculation Mode Management

**What**: Control when Excel recalculates formulas.

**Why**:
- Prevent unwanted recalculations during batch setup
- Improve performance by deferring calculation
- Control calculation timing

**Where Used**:
- `RefreshLSEGWithTimeout:12` - Store original mode
- `RefreshLSEGWithTimeout:25` - Set to manual
- `RefreshLSEGWithTimeout:47` - Restore original

**Example**:
```vba
Dim originalCalcMode As XlCalculation
originalCalcMode = Application.Calculation

Application.Calculation = xlCalculationManual
' ... do work without recalculating ...
Application.Calculation = originalCalcMode  ' Restore
```

---

### 7. State Machine Pattern

**What**: Track process state across asynchronous calls.

**Why**:
- OnTime breaks execution chain - need to track where we are
- Enable stop/abort capability
- Coordinate multiple async phases

**Where Used**:
- `OptionDownload:30-35` - BatchProcessState enum
- Global variables: `g_BatchState`, `g_BatchStartRow`, `g_BatchEndRow`, etc.

**States**:
```vba
Public Enum BatchProcessState
    bpsIdle = 0              ' No batch processing
    bpsSetupFormulas = 1     ' Creating RHistory formulas
    bpsRefreshing = 2        ' LSEG refreshing data
    bpsProcessingResults = 3 ' Calculating Greeks, validating
End Enum
```

**State Transitions**:
```
bpsIdle → bpsSetupFormulas → bpsRefreshing → bpsProcessingResults → bpsIdle (next batch)
```

---

### 8. Cell Text Inspection - Detect LSEG Status

**What**: Check cell's displayed text (not value) to detect LSEG loading status.

**Why**: LSEG shows "Retrieving..." in cell `.Text` while loading data.

**Where Used**:
- `IsDataReady` function (`OptionDownload:664`)

**Example**:
```vba
Dim cellText As String
cellText = CStr(ws.Cells(checkRow, 2).Text)  ' .Text shows formatted display

If InStr(1, cellText, "Retrieving...", vbTextCompare) = 0 Then
    ' Cell no longer shows "Retrieving..." - data is ready
    readyCount = readyCount + 1
End If
```

**Key Distinction**:
- `.Value` - Returns underlying value (may be error or empty during loading)
- `.Text` - Returns displayed text including "Retrieving..." message

---

## Usage Examples

### Example 1: Simple Non-Blocking Refresh

**Scenario**: Refresh a single worksheet with LSEG data.

```vba
Sub RefreshFutureSheet()
    Dim wsFuture As Worksheet
    Set wsFuture = ThisWorkbook.Worksheets("Future et co")

    ' Non-blocking refresh with 60-second timeout
    RefreshLSEGWithTimeout wsFuture, 60

    MsgBox "Refresh complete! Check data.", vbInformation
    Application.StatusBar = False
End Sub
```

**When to Use**:
- Single sheet with LSEG formulas
- User-initiated manual refresh
- No batch processing needed

---

### Example 2: Batch Processing with Progress Tracking

**Scenario**: Download hundreds of option prices in batches.

```vba
Sub MainDownloadProcess()
    ' 1. Initialize
    InitializeWorkbook
    LoadConfiguration

    ' 2. Validate prerequisites
    If Not CheckRICListExists() Then
        MsgBox "Please run GenerateAllRICs first!", vbExclamation
        Exit Sub
    End If

    If Not CheckUnderlyings() Then
        MsgBox "Please download underlying data first!", vbExclamation
        Exit Sub
    End If

    ' 3. Initialize batch state
    g_BatchCounter = 0
    g_StopRequested = False
    g_BatchState = bpsIdle

    ' 4. Find first unprocessed batch
    batchStart = FindNextUnprocessedRIC(2)
    batchEnd = Application.Min(batchStart + g_BatchSize - 1, lastRow)

    g_BatchStartRow = batchStart
    g_BatchEndRow = batchEnd
    g_BatchCounter = 1

    ' 5. Start async chain
    ProcessBatch_SetupFormulas

    ' EXECUTION ENDS HERE - OnTime chain handles the rest
End Sub
```

**User Experience**:
- Excel remains responsive throughout
- Status bar shows progress
- Can stop with `StopBatchProcessing()` macro
- Auto-saves every 3 batches

---

### Example 3: Add LSEG Data to New Project

**Scenario**: You have a new workbook and need to add LSEG data download capability.

**Step 1**: Import LSEGCalc module
```vba
' Create new module: LSEGCalc.bas
' Copy code from LSEGCalc.bas (lines 1-72)
```

**Step 2**: Add LSEG formulas to worksheet
```vba
Sub SetupMyFormulas()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MyDataSheet")

    ' Add RHistory formula
    ws.Range("A2").Formula = "=RHistory(""ESZ4"","".Timestamp;.Close"",""START:2024-01-01 END:2024-12-31 INTERVAL:1D"")"

    ' More formulas...
    ws.Range("A3").Formula = "=RHistory(""1EZ4"","".Timestamp;.Close"",""START:2024-01-01 END:2024-12-31 INTERVAL:1D"")"
End Sub
```

**Step 3**: Refresh with timeout protection
```vba
Sub RefreshMyData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MyDataSheet")

    ' Setup formulas
    SetupMyFormulas

    ' Non-blocking refresh with 2-minute timeout
    RefreshLSEGWithTimeout ws, 120

    MsgBox "Refresh complete!", vbInformation
    Application.StatusBar = False
End Sub
```

---

### Example 4: Polling Pattern for Custom Check

**Scenario**: You need to wait for LSEG data but have custom readiness criteria.

```vba
' Global variables
Public g_CheckCount As Long

Sub StartMyDownload()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MyData")

    ' Setup formulas and trigger refresh
    ws.Range("A2").Formula = "=RHistory(...)"
    RefreshLSEGWithTimeout ws, 60

    ' Initialize and schedule first check
    g_CheckCount = 0
    Application.OnTime Now + TimeValue("00:00:03"), "CheckMyDataReady"
End Sub

Sub CheckMyDataReady()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MyData")

    g_CheckCount = g_CheckCount + 1

    ' Timeout after 20 checks (1 minute)
    If g_CheckCount > 20 Then
        MsgBox "Timeout waiting for data!", vbExclamation
        Exit Sub
    End If

    ' Custom readiness check
    If IsMyDataReady(ws) Then
        ProcessMyData  ' Data is ready
    Else
        ' Not ready - reschedule check in 3 seconds
        Application.OnTime Now + TimeValue("00:00:03"), "CheckMyDataReady"
    End If
End Sub

Function IsMyDataReady(ws As Worksheet) As Boolean
    ' Custom logic - check if specific cells are populated
    Dim cellText As String
    cellText = CStr(ws.Range("A2").Text)

    ' Ready if not showing "Retrieving..." and has numeric value
    IsMyDataReady = (InStr(1, cellText, "Retrieving...", vbTextCompare) = 0) And _
                    IsNumeric(ws.Range("A2").Value)
End Function

Sub ProcessMyData()
    MsgBox "Data is ready - processing!", vbInformation
    ' ... process data ...
End Sub
```

---

## Best Practices

### 1. Always Use DoEvents in Wait Loops

**✅ Good**:
```vba
Do While Application.CalculationState <> xlDone
    DoEvents  ' Returns control to Excel
    Application.Wait Now + TimeValue("00:00:01")
Loop
```

**❌ Bad**:
```vba
Do While Application.CalculationState <> xlDone
    ' No DoEvents - Excel freezes!
Loop
```

---

### 2. Implement Timeout Protection

**✅ Good**:
```vba
Dim startTime As Double
Dim timeoutSeconds As Long
startTime = Timer
timeoutSeconds = 120

Do While Application.CalculationState <> xlDone
    DoEvents

    If Timer - startTime > timeoutSeconds Then
        MsgBox "Timeout - aborting", vbExclamation
        Exit Do
    End If
Loop
```

**❌ Bad**:
```vba
Do While Application.CalculationState <> xlDone
    DoEvents  ' Could loop forever!
Loop
```

---

### 3. Preserve and Restore Calculation Mode

**✅ Good**:
```vba
Dim originalCalcMode As XlCalculation
originalCalcMode = Application.Calculation

Application.Calculation = xlCalculationManual
' ... do work ...
Application.Calculation = originalCalcMode  ' Always restore
```

**❌ Bad**:
```vba
Application.Calculation = xlCalculationManual
' ... do work ...
' Forgot to restore - user's Excel now stuck in manual mode!
```

---

### 4. Use OnTime for Long Operations

**✅ Good** (for batch processing):
```vba
Sub Phase1()
    ' Setup formulas
    ws.Range("A1").Formula = "=RHistory(...)"

    RefreshLSEGWithTimeout ws, 60

    ' Schedule next phase - VBA ends, Excel regains control
    Application.OnTime Now + TimeValue("00:00:05"), "Phase2"
End Sub

Sub Phase2()
    ' Process results
    ProcessData
End Sub
```

**❌ Bad** (blocks Excel):
```vba
Sub AllInOne()
    ' Setup formulas
    ws.Range("A1").Formula = "=RHistory(...)"

    RefreshLSEGWithTimeout ws, 60

    ' Blocks VBA execution until complete
    Do While Not IsDataReady(ws)
        DoEvents
        Application.Wait Now + TimeValue("00:00:03")
    Loop

    ProcessData  ' Could be minutes later!
End Sub
```

---

### 5. Provide User Feedback

**✅ Good**:
```vba
Application.StatusBar = "Batch " & batchNum & ": Processing " & ricCount & " RICs..."
DoEvents  ' Update immediately

' Update periodically during loop
If i Mod 10 = 0 Then
    Application.StatusBar = "Batch " & batchNum & ": Processed " & i & " of " & total
End If

' Clear when done
Application.StatusBar = False
```

**❌ Bad**:
```vba
' No feedback - user has no idea what's happening
For i = 1 To 1000
    ProcessRIC i
Next i
```

---

### 6. Implement Stop Capability

**✅ Good**:
```vba
Public g_StopRequested As Boolean

Sub StartProcess()
    g_StopRequested = False
    ProcessBatch_Phase1
End Sub

Sub ProcessBatch_Phase1()
    If g_StopRequested Then
        Abort
        Exit Sub
    End If
    ' ... do work ...
End Sub

Sub StopProcessing()
    g_StopRequested = True
    MsgBox "Will stop after current phase", vbInformation
End Sub
```

**❌ Bad**:
```vba
Sub StartProcess()
    For i = 1 To 1000
        ProcessBatch i  ' No way to stop!
    Next i
End Sub
```

---

### 7. Save Incrementally for Long Processes

**✅ Good**:
```vba
' Save batch results to CSV after each batch
SaveBatchToCSV batchNum

' Auto-save workbook every N batches
If batchNum Mod 3 = 0 Then
    ThisWorkbook.Save
End If
```

**❌ Bad**:
```vba
' Process all 1000 batches...
For i = 1 To 1000
    ProcessBatch i
Next i

' Save once at end - crash = lose everything!
ThisWorkbook.Save
```

---

### 8. Check Cell .Text for LSEG Status

**✅ Good**:
```vba
Dim cellText As String
cellText = CStr(ws.Cells(row, 2).Text)

If InStr(1, cellText, "Retrieving...", vbTextCompare) = 0 Then
    ' Data is ready
End If
```

**❌ Bad**:
```vba
If Not IsEmpty(ws.Cells(row, 2).Value) Then
    ' May be True even while LSEG is still loading!
End If
```

---

### 9. Use State Machine for Complex Async Flows

**✅ Good**:
```vba
Public Enum ProcessState
    psIdle = 0
    psDownloading = 1
    psProcessing = 2
    psCompleting = 3
End Enum

Public g_State As ProcessState

Sub Phase1()
    g_State = psDownloading
    ' ... work ...
    Application.OnTime Now + TimeValue("00:00:05"), "Phase2"
End Sub

Sub Phase2()
    If g_State <> psDownloading Then
        MsgBox "Unexpected state!", vbCritical
        Exit Sub
    End If

    g_State = psProcessing
    ' ... work ...
End Sub
```

**❌ Bad**:
```vba
' No state tracking - hard to debug, no coordination
Sub Phase1()
    ' ... work ...
    Application.OnTime Now + TimeValue("00:00:05"), "Phase2"
End Sub

Sub Phase2()
    ' How do we know Phase1 completed? What if user ran this manually?
    ' ... work ...
End Sub
```

---

### 10. Validate Prerequisites Before Starting

**✅ Good**:
```vba
Sub MainProcess()
    ' Check all prerequisites
    If Not CheckRICListExists() Then
        MsgBox "Please generate RIC list first!", vbExclamation
        Exit Sub
    End If

    If Not CheckUnderlyingsAvailable() Then
        MsgBox "Please download underlyings first!", vbExclamation
        Exit Sub
    End If

    If Not ValidateConfiguration() Then
        MsgBox "Invalid configuration!", vbExclamation
        Exit Sub
    End If

    ' All checks passed - start process
    StartBatchProcessing
End Sub
```

**❌ Bad**:
```vba
Sub MainProcess()
    ' Just start - fail halfway through when data is missing!
    StartBatchProcessing
End Sub
```

---

## Common Pitfalls

### Pitfall 1: Re-entrancy with DoEvents

**Problem**: `DoEvents` allows users to trigger other macros, which can cause re-entrancy issues.

**Solution**: Use state flags to prevent re-entrancy.

```vba
Public g_IsProcessing As Boolean

Sub MyProcess()
    If g_IsProcessing Then
        MsgBox "Process already running!", vbExclamation
        Exit Sub
    End If

    g_IsProcessing = True

    ' ... work with DoEvents ...

    g_IsProcessing = False
End Sub
```

---

### Pitfall 2: Timer Rollover at Midnight

**Problem**: `Timer` resets to 0 at midnight, breaking timeout calculations.

**Solution**: Use `Now` for operations that might span midnight.

```vba
' Instead of Timer
Dim startTime As Double
startTime = Timer

' Use Now for midnight safety
Dim startTime As Date
startTime = Now

If DateDiff("s", startTime, Now) > timeoutSeconds Then
    ' Timeout
End If
```

---

### Pitfall 3: OnTime Doesn't Wait

**Problem**: Developers expect `OnTime` to pause execution, but it schedules and immediately continues.

**Example**:
```vba
Sub Wrong()
    Application.OnTime Now + TimeValue("00:00:05"), "Phase2"

    ' This runs IMMEDIATELY, not after 5 seconds!
    ProcessResults  ' ← BUG: Data not ready yet!
End Sub
```

**Solution**: Phase2 should do the processing.

```vba
Sub Correct()
    Application.OnTime Now + TimeValue("00:00:05"), "Phase2"
    ' Phase2 will run in 5 seconds
End Sub

Sub Phase2()
    ProcessResults  ' ← Runs 5 seconds later
End Sub
```

---

### Pitfall 4: Forgetting to Restore Settings

**Problem**: If error occurs, calculation mode or other settings not restored.

**Solution**: Use error handlers with cleanup.

```vba
Sub MyProcess()
    Dim originalCalcMode As XlCalculation
    originalCalcMode = Application.Calculation

    On Error GoTo CleanUp

    Application.Calculation = xlCalculationManual
    ' ... work ...

CleanUp:
    Application.Calculation = originalCalcMode
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbCritical
    End If
End Sub
```

---

## Troubleshooting

### Issue: Excel Still Freezing

**Possible Causes**:
1. Not using `DoEvents` in wait loops
2. Loop timeout too long
3. LSEG operation genuinely blocking

**Solution**:
- Verify `DoEvents` in all wait loops
- Reduce timeout or add more frequent `DoEvents` calls
- Break into smaller batches

---

### Issue: Data Not Populating

**Possible Causes**:
1. LSEG formulas not refreshing
2. `CalculationState` check too early
3. RIC codes invalid

**Solution**:
- Manually test LSEG formula in cell
- Increase delay before checking (5→10 seconds)
- Validate RIC codes with LSEG documentation

---

### Issue: Process Stops Mid-Batch

**Possible Causes**:
1. Error in OnTime callback
2. `g_StopRequested` set to True
3. Timeout exceeded

**Solution**:
- Check VBA error log
- Verify state variables (`g_BatchState`, `g_StopRequested`)
- Increase timeout values

---

### Issue: "Method OnTime Failed"

**Possible Causes**:
1. Target procedure name misspelled
2. Procedure not public
3. Too many OnTime calls scheduled

**Solution**:
- Verify procedure name in `g_NextScheduledProc` matches actual sub name
- Ensure procedure is `Public Sub`, not `Private Sub`
- Cancel previous OnTime before scheduling new one

---

## Summary

This LSEG VBA system solves Excel freezing during LSEG data downloads through:

1. **Non-blocking refresh** (`RefreshLSEGWithTimeout`) - DoEvents + timeout monitoring
2. **Asynchronous batch processing** - Application.OnTime breaks execution chain
3. **State machine pattern** - Track progress across async calls
4. **Polling for readiness** - Check cell text for "Retrieving..." status
5. **Calculation mode management** - Control when Excel recalculates
6. **Incremental saving** - Protect against data loss
7. **User feedback & control** - Status bar updates and stop capability

The architecture is applicable to any VBA project requiring:
- Long-running LSEG data downloads
- Batch processing of hundreds/thousands of items
- Non-blocking operations with progress tracking
- Timeout protection and cancellation capability

**Key Files**:
- `LSEGCalc.bas` - Core non-blocking refresh (lines 5-72)
- `OptionDownload.bas` - Batch processing orchestration (lines 376-693)
- `RICconfiguration.bas` - RIC generation with refresh (lines 21-112)

For new projects, start with `LSEGCalc.bas` and adapt the OnTime chain pattern as needed for your specific workflow.

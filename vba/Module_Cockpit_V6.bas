Attribute VB_Name = "Module_Cockpit_V6"
'Attribute VB_Name = "Module_COCKPIT_V6"
Option Explicit

'----------------------------------------------------------------------
' Module_COCKPIT_V6  --  10m vs 30 EMA Crossover Intraday Trading Engine
'
' Target Worksheet: "COCKPIT_V6"
' Status: ACTIVE -- Automated 10m Entry / 5m Exit Intraday Strategy
'
' Architecture:
'   Data Source: 100% from Sheet1 (Col DJ = 10m Entry Cross, Col AC = latest completed 5m Exit Cross, Col T = LTP)
'   Discovery:   10m EMA 10 vs 30 cross on forming bar (Col DJ = BULLISH or BEARISH)
'   Contra exit: A valid opposite 10m cross exits the open trade at current LTP.
'   RSI Filter:  Configurable bounds (BUY: 52 - 64, SELL: 38 - 47)
'   Stock ADX:   Entry requires Stock ADX >= 20 (Col CB); exit if 5m ADX falls below 18
'   Direction:   Direction gate removed (trades BUY & SELL independently based on stock signal)
'   RVOL Filter: RVOL floor filter removed (logged for telemetry; no volume floor blocking trades)
'   Ranking:     Candidates beyond free slots ranked by RVOL (primary) then ADX (tie-break)
'   Signals:     All qualifying signals are retained; AC selects live Apply vs Monitor-only
'   Entry signal: 10m forming-bar cross (09:25:00 to 14:55:00)
'   Entry order:  Resting limit at the nearest 5m EMA (10/30/50/100/200) on the
'                 pullback side of LTP, frozen at signal time; cancel after 3 subsequent
'                 completed 5m bars if unfilled; release the slot only after FYERS confirms cancel
'   Sizing:      Dynamic: live available capital / free slots / per-stock intraday margin (Python)
'   Stop Loss:   Fixed 1.0% from Entry Price (Entry * 0.99 for BUY, Entry * 1.01 for SELL)
'   Target:      Fixed 1.5% from Entry Price (Entry * 1.015 for BUY, Entry * 0.985 for SELL)
'   Exits: Broker-managed SL/TP for live rows; simulated local SL/TP for Monitor rows;
'     5m ADX < 18, opposite 10m cross, and completed 5m reverse cross exit at current LTP;
'     15:00 liquidation remains V6-managed for both.
'     (10m forming-candle cross remains strictly for ENTRY.)
'
' Broker Integration:
'   Dispatches acknowledged live orders via engine/fyers_order_manager_v6.py
'----------------------------------------------------------------------

' =====================================================================
'  CONFIGURATION & RISK CONSTANTS (ALL USER-CONFIGURABLE)
' =====================================================================
Public Const V6_ENABLED                 As Boolean = True
Public Const ENABLE_BROKER_EXECUTION    As Boolean = True   ' Set to True when ready for live FYERS orders

' --- Position Sizing & Slot Limits ---
Public Const MAX_CONCURRENT_POSITIONS   As Long = 3         ' Maximum concurrent active positions (also the sizing denominator)
Public Const SL_PCT                     As Double = 0.01    ' Fixed 1.0% Stop Loss
Public Const TGT_PCT                    As Double = 0.012   ' Fixed 1.5% Target (1.5% from LTP)

' --- Configurable RSI Filter Settings (Empirically Validated Sweet Spots) ---
Public Const RSI_FILTER_ENABLED         As Boolean = True   ' Toggle RSI filter ON/OFF
Public Const RSI_SOURCE_COL             As String = "DI"    ' Sheet1 Column for live forming 10m RSI
Public Const BUY_RSI_MIN                As Double = 40#     ' Minimum RSI for BUY entry (sweet spot: 52 - 64)
Public Const BUY_RSI_MAX                As Double = 60#     ' Maximum RSI for BUY entry (prevents buying into exhaustion)
Public Const SELL_RSI_MIN               As Double = 38#     ' Minimum RSI for SELL entry (sweet spot: 38 - 47)
Public Const SELL_RSI_MAX               As Double = 60#     ' Maximum RSI for SELL entry (prevents shorting into oversold capitulation)

' --- Stock ADX Filter (Stock Trend Conviction vs Consolidation Chop) ---
Public Const ADX_FILTER_ENABLED         As Boolean = True   ' Toggle Stock ADX filter ON/OFF
Public Const ADX_SOURCE_COL             As String = "CB"    ' Sheet1 Column for Stock ADX (Col CB = 14-period Wilder ADX)
Public Const STOCK_ADX_MIN              As Double = 20#     ' Minimum Stock ADX required for entry (ADX >= 20 = trending, < 20 = chop)

' --- Market Direction & RVOL Filters (REMOVED / DISABLED per user request) ---
Public Const MARKET_FILTER_ENABLED      As Boolean = False  ' Direction gate removed (trades BUY & SELL regardless of Nifty)
Public Const RVOL_FILTER_ENABLED        As Boolean = False  ' RVOL floor filter removed (no trades blocked by RVOL)
Public Const RVOL_SOURCE_COL            As String = "BP"    ' Sheet1 Column for RVOL (logged for telemetry)

' --- Candidate Selection & Tie-Breaking ---
Public Const RANK_CANDIDATES           As Boolean = True   ' Rank candidates beyond free slots (RVOL primary, ADX tie-break)

' --- Session Timing ---
Private Const EARLY_ENTRY_START         As String = "09:25:00" ' Trading window start: 09:25:00
Private Const LATE_ENTRY_CUTOFF         As String = "14:45:00" ' Trading window end: 14:55:00
Private Const FORCE_LIQUIDATION         As String = "15:05:00" ' Mandatory session close
Private Const DATA_START_ROW            As Long = 5
Private Const ENTRY_LIMIT_WAIT_BARS      As Long = 3

' =====================================================================
'  COCKPIT_V6 COLUMN LETTERS
' =====================================================================
Private Const CK_TICKER       As String = "A"
Private Const CK_ACTION       As String = "B"
Private Const CK_LTP          As String = "C"
Private Const CK_SECTOR       As String = "D"
Private Const CK_S3           As String = "E"
Private Const CK_S2           As String = "F"
Private Const CK_S1           As String = "G"
Private Const CK_PP           As String = "H"
Private Const CK_R1           As String = "I"
Private Const CK_R2           As String = "J"
Private Const CK_R3           As String = "K"
Private Const CK_RSI_5M       As String = "L"
Private Const CK_RSI_15M      As String = "M"
Private Const CK_RSI_1H       As String = "N"
Private Const CK_RSI_1D       As String = "O"
Private Const CK_SUPERTREND   As String = "P"
Private Const CK_CROSS_10_30  As String = "Q"
Private Const CK_CROSS_10_50  As String = "R"
Private Const CK_ENT          As String = "S"
Private Const CK_TGT          As String = "T"
Private Const CK_SL           As String = "U"
Private Const CK_CP           As String = "V"
Private Const CK_OPEN_TIME    As String = "W"
Private Const CK_CLOSE_TIME   As String = "X"
Private Const CK_COMENTS      As String = "AA"
Private Const CK_CANDLE_COUNT As String = "AB"     ' Bars held
Private Const CK_PEAK_PRICE   As String = "AD"     ' Best price seen
Private Const CK_BEST_MFE_PCT As String = "AE"     ' Peak favorable %
Private Const CK_LAST_BAR_TS  As String = "AI"     ' Bar timestamp
Private Const CK_EXEC_STATE  As String = "AJ"     ' Broker-linked V6 state
Private Const CK_ORDER_ID    As String = "AK"     ' FYERS order id
Private Const CK_FILLED_QTY  As String = "AL"     ' Actual broker filled quantity
Private Const CK_REQ_QTY     As String = "AM"     ' Requested quantity
Private Const CK_REQUEST_ID  As String = "AN"     ' Async dispatcher correlation id
Private Const CK_EXIT_FILL   As String = "AO"     ' Actual FYERS exit fill price
Private Const CK_REQUEST_TYPE As String = "AP"     ' ENTRY or EXIT request kind
Private Const CK_APPLY_TO_FYERS As String = "AC"   ' TRUE=live apply; FALSE=monitor-only
Private Const CK_ORIGIN       As String = "AQ"     ' V6, USER, PAPER, or MONITOR
Private Const CK_HAS_BROKER_SL As String = "AR"    ' FYERS currently reports attached SL
Private Const CK_HAS_BROKER_TP As String = "AS"    ' FYERS currently reports attached TP

' One broker snapshot is fetched per run and reused for reconciliation and slot counting.
Private gV6Snapshot As String
Private gV6SnapshotReady As Boolean
Private gV6LiveExitSubmittedThisRun As Boolean
Private gV6ErrorContext As String
Private gV6ReconciliationReady As Boolean


' =====================================================================
' Public Sub  Run_Cockpit_V6
' Main Entry Point -- Called natively by Excel Macro 'RunAll'
' =====================================================================
Public Sub Run_Cockpit_V6()
    If Not V6_ENABLED Then Exit Sub

    Dim currentTime As Date
    currentTime = TimeValue(Now)

    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Dim wsC6  As Worksheet: Set wsC6 = GetOrCreateCockpitV6(ThisWorkbook)

    EnsureCockpitV6Headers wsC6

    ' Fetch broker state once, then reuse the same snapshot for reconciliation and capacity.
    gV6SnapshotReady = False
    gV6Snapshot = vbNullString
    gV6LiveExitSubmittedThisRun = False
    gV6ReconciliationReady = False
    gV6ErrorContext = "STARTUP"
    If ENABLE_BROKER_EXECUTION Then
        gV6ErrorContext = "BROKER_SNAPSHOT"
        gV6SnapshotReady = GetV6BrokerSnapshot(gV6Snapshot)
        If gV6SnapshotReady Then
            gV6ErrorContext = "INITIAL_RECONCILIATION"
            ReconcileBrokerPositions wsC6, wsSrc, gV6Snapshot
            gV6ReconciliationReady = V6SnapshotFullyReconciled(wsC6, gV6Snapshot)
        Else
            wsC6.cells(2, 14).value = "WARNING: FYERS snapshot failed; live entries blocked; monitor-only signals may still be recorded."
            V6ReconcileLog "SNAPSHOT_FAILED; reconciliation skipped"
        End If
    Else
        gV6ReconciliationReady = True
    End If

    ' --- PHASE 1: MANAGE EXISTING OPEN POSITIONS (DEFENSE) ---
    gV6ErrorContext = "MANAGE_EXISTING_POSITIONS"
    ManageCockpitV6Trades wsC6, wsSrc, currentTime

    ' A live exit may free capacity during this run. Refresh once after dispatch,
    ' and count the slot only if FYERS now confirms the position is flat.
    If ENABLE_BROKER_EXECUTION And gV6LiveExitSubmittedThisRun Then
        Dim postExitSnapshot As String
        If GetV6BrokerSnapshot(postExitSnapshot) Then
            gV6Snapshot = postExitSnapshot
            gV6SnapshotReady = True
            ReconcileBrokerPositions wsC6, wsSrc, gV6Snapshot
            gV6ReconciliationReady = V6SnapshotFullyReconciled(wsC6, gV6Snapshot)
            V6ReconcileLog "POST_EXIT_SNAPSHOT refreshed before entry scan"
        Else
            V6ReconcileLog "POST_EXIT_SNAPSHOT failed; retain original broker counts and keep slots reserved"
        End If
    End If

    ' --- PHASE 2: SCAN FOR NEW TRADES (OFFENSE) ---
    If currentTime >= TimeValue(EARLY_ENTRY_START) And currentTime < TimeValue(LATE_ENTRY_CUTOFF) Then
        gV6ErrorContext = "SCAN_AND_DISPATCH"
        ScanForNewV6Trades wsC6, wsSrc
    End If

    ' N2 is a run-level status cell, so refresh it even after the entry window closes.
    UpdateV6PositionStatus wsC6

Cleanup:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Dim runErrNumber As Long: runErrNumber = Err.Number
    Dim runErrDescription As String: runErrDescription = Err.Description
    Dim runErrContext As String: runErrContext = gV6ErrorContext
    On Error Resume Next
    V6ReconcileLog "RUN_ERROR number=" & runErrNumber & " description=" & runErrDescription & _
        " context=" & runErrContext
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = True
    MsgBox "Error in Run_Cockpit_V6 (" & runErrNumber & ") at " & runErrContext & ": " & _
        runErrDescription, vbCritical
End Sub

Private Sub UpdateV6PositionStatus(wsC As Worksheet)
    Dim priorStatus As String
    priorStatus = Trim$(CStr(wsC.cells(2, 14).value))

    Dim sheetOpen As Long
    sheetOpen = CountOpenV6Trades(wsC)

    If Not ENABLE_BROKER_EXECUTION Then
        wsC.cells(2, 14).value = "Broker execution disabled | sheet open: " & sheetOpen
        Exit Sub
    End If

    If Not gV6SnapshotReady Then
        wsC.cells(2, 14).value = "WARNING: FYERS snapshot unavailable | sheet open: " & sheetOpen & ". Live entries blocked."
        Exit Sub
    End If

    Dim apiOpen As Long: apiOpen = CountSnapshotOpenPositions(gV6Snapshot)
    Dim pendingSymbols As Long: pendingSymbols = CountSnapshotPendingSymbols(gV6Snapshot)
    Dim brokerRows As Long: brokerRows = CountOpenBrokerRows(wsC)
    Dim monitorRows As Long: monitorRows = CountOpenMonitorRows(wsC)
    wsC.cells(2, 14).value = "FYERS positions=" & apiOpen & " | pending symbols=" & pendingSymbols & _
        " | broker rows=" & brokerRows & " | monitor rows=" & monitorRows & _
        " | reconciliation=" & IIf(gV6ReconciliationReady, "OK", "REVIEW")

    ' Keep an IP preflight failure visible while still showing the position counts.
    If Left$(UCase$(priorStatus), 14) = "ENTRY BLOCKED:" Then
        wsC.cells(2, 14).value = wsC.cells(2, 14).value & " | " & priorStatus
    End If
End Sub


' =====================================================================
' Private Sub  ManageCockpitV6Trades
' Evaluates monitor rows against local target/stop, and all rows against:
'   1. Opposite 10m cross (Sheet1 Col DJ), exited at the observed LTP
'   2. Completed 5m Reverse Cross (Sheet1 Col AC), exited at the observed LTP
'   3. 15:05 market liquidation
' Live rows leave attached target/stop handling to FYERS so terminal edits are respected.
' =====================================================================
Private Sub ManageCockpitV6Trades(wsC As Worksheet, wsSrc As Worksheet, currentTime As Date)
    Dim lastRow As Long
    lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    If lastRow < DATA_START_ROW Then Exit Sub

    Dim forceLiquidate As Boolean
    forceLiquidate = (currentTime >= TimeValue(FORCE_LIQUIDATION))

    Dim dispatchOk As Boolean
    Dim srcRow As Long, curLTP As Double, curADX As Double, adxAvailable As Boolean
    Dim cross5m As String, cross10m As String

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim ticker As String: ticker = Trim$(CStr(wsC.Range(CK_TICKER & r).value))
        srcRow = 0
        curLTP = 0
        curADX = 0
        adxAvailable = False
        cross5m = ""
        cross10m = ""
        If Len(ticker) = 0 Then GoTo NextTrade
        ' Check if trade is still OPEN (CloseTime in Col X is blank)
        Dim closeTimeVal As Variant: closeTimeVal = wsC.Range(CK_CLOSE_TIME & r).value
        If Not IsEmpty(closeTimeVal) And Len(Trim$(CStr(closeTimeVal))) > 0 Then
            GoTo NextTrade
        End If
        Dim execState As String: execState = UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & r).value)))
        If execState = "MONITOR_ENTRY_PENDING" Then
            ManageMonitorPendingEntry wsC, wsSrc, r, ticker
            GoTo NextTrade
        End If
        If ENABLE_BROKER_EXECUTION And V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) _
                And Not gV6SnapshotReady Then
            V6ReconcileLog "row=" & r & " ticker=" & ticker & " EXIT_CHECK_SKIPPED broker snapshot unavailable; no order sent"
            GoTo NextTrade
        End If
        If Len(execState) > 0 And execState <> "OPEN" Then
            GoTo NextTrade
        End If

        Dim action As String:   action = UCase$(Trim$(CStr(wsC.Range(CK_ACTION & r).value)))
        Dim entryPx As Double:  entryPx = CDbl(val(wsC.Range(CK_ENT & r).value))
        Dim tgtPx As Double:    tgtPx = CDbl(val(wsC.Range(CK_TGT & r).value))
        Dim slPx As Double:     slPx = CDbl(val(wsC.Range(CK_SL & r).value))
        If entryPx <= 0 Then GoTo NextTrade

        ' Fetch live telemetry from Sheet1 for this ticker
        srcRow = FindSheet1Row(wsSrc, ticker)
        If srcRow >= 4 Then
            curLTP = CDbl(val(wsSrc.cells(srcRow, "T").value))     ' Col T: LTP
            cross5m = UCase$(Trim$(CStr(wsSrc.cells(srcRow, "AC").value))) ' Col AC: latest completed 5m EMA cross
            cross10m = UCase$(Trim$(CStr(wsSrc.cells(srcRow, "DJ").value))) ' Col DJ: fresh forming 10m EMA cross
            Dim adxCell As Variant: adxCell = wsSrc.cells(srcRow, ADX_SOURCE_COL).value
            If Not IsError(adxCell) Then
                adxAvailable = (Len(Trim$(CStr(adxCell))) > 0 And IsNumeric(adxCell))
                If adxAvailable Then curADX = CDbl(adxCell) ' Col CB: 5m ADX(14)
            End If
        ElseIf UCase$(Trim$(CStr(wsC.Range(CK_ORIGIN & r).value))) = "USER" _
                Or UCase$(Trim$(CStr(wsC.Range(CK_ORIGIN & r).value))) = "MIXED" Then
            curLTP = SnapshotNumberForSymbol(gV6Snapshot, ticker, "ltp", 0#)
            If curLTP <= 0 Then GoTo NextTrade
        Else
            GoTo NextTrade
        End If
        If curLTP <= 0 Then GoTo NextTrade

        ' Update Current Price and LTP on sheet
        wsC.Range(CK_LTP & r).value = curLTP
        wsC.Range(CK_CP & r).value = curLTP

        ' Update bars held count
        Dim candleCount As Long: candleCount = CLng(val(wsC.Range(CK_CANDLE_COUNT & r).value)) + 1
        wsC.Range(CK_CANDLE_COUNT & r).value = candleCount

        ' Compute MFE and update Peak Price
        Dim curMfePct As Double
        If action = "BUY" Then
            curMfePct = (curLTP - entryPx) / entryPx
        Else
            curMfePct = (entryPx - curLTP) / entryPx
        End If

        Dim bestMfe As Double: bestMfe = val(wsC.Range(CK_BEST_MFE_PCT & r).value)
        If curMfePct > bestMfe Then
            wsC.Range(CK_BEST_MFE_PCT & r).value = curMfePct
            wsC.Range(CK_PEAK_PRICE & r).value = curLTP
        End If

        ' === CHECK EXIT CONDITIONS (TGT, SL, ADX, EMA CROSSES & 15:05) ===
        Dim exitFlag As Boolean: exitFlag = False
        Dim exitReason As String: exitReason = ""
        Dim exitPx As Double: exitPx = curLTP
        Dim exitOrderLTP As Double: exitOrderLTP = 0#
        Dim retryComment As String: retryComment = Trim$(CStr(wsC.Range(CK_COMENTS & r).value))
        Dim retryExit As Boolean: retryExit = (Left$(retryComment, 11) = "EXIT_RETRY|")

        ' A failed/unfilled LTP-limit exit remains latched and is retried at the
        ' latest LTP on each run until FYERS confirms the position is flat.
        If retryExit And Not forceLiquidate Then
            exitReason = Mid$(retryComment, 12)
            If InStr(1, exitReason, " | ", vbTextCompare) > 0 Then
                exitReason = Left$(exitReason, InStr(1, exitReason, " | ", vbTextCompare) - 1)
            End If
            exitFlag = True
            exitOrderLTP = curLTP
        End If

        ' 1. Mandatory 15:00:00 Market Close
        If retryExit And Not forceLiquidate Then
            ' Retry the latched contra/reverse-cross exit at the current LTP.
        ElseIf forceLiquidate Then
            exitFlag = True
            exitReason = "3_PM_EXIT (Session Liquidation)"

        ' Applied live trades use FYERS-attached SL/TP levels. Do not trigger against
        ' the local estimate: the user may have modified those levels in FYERS.
        ' Monitor-only (and paper) rows simulate their local target/stop levels.
        ElseIf V6UseLocalTakeProfit(wsC, r) _
                And tgtPx > 0 And action = "BUY" And curLTP >= tgtPx Then
            exitFlag = True
            exitPx = tgtPx
            exitReason = "TARGET (1.5% Hit)"
        ElseIf V6UseLocalTakeProfit(wsC, r) _
                And tgtPx > 0 And action = "SELL" And curLTP <= tgtPx Then
            exitFlag = True
            exitPx = tgtPx
            exitReason = "TARGET (1.5% Hit)"

        ' 3. Fixed 1.0% Stop-Loss Hit (monitor/paper rows only; live rows are broker-managed)
        ElseIf V6UseLocalStopLoss(wsC, r) _
                And action = "BUY" And curLTP <= slPx Then
            exitFlag = True
            exitPx = slPx
            exitReason = "STOP_LOSS (1.0% Hit)"
        ElseIf V6UseLocalStopLoss(wsC, r) _
                And action = "SELL" And curLTP >= slPx Then
            exitFlag = True
            exitPx = slPx
            exitReason = "STOP_LOSS (1.0% Hit)"

        ' 4. 5m ADX(14) weakening below 18: exit at observed LTP (both live and monitor rows).
        ElseIf adxAvailable And curADX < 18# Then
            exitFlag = True
            exitPx = curLTP
            exitOrderLTP = curLTP
            exitReason = "ADX_5M_BELOW_18"

        ' 5. Opposite 10m entry cross: close the existing position, never reverse it here.
        ElseIf action = "BUY" And cross10m = "BEARISH" Then
            exitFlag = True
            exitPx = curLTP
            exitOrderLTP = curLTP
            exitReason = "EMA_CROSS_10M_CONTRA"
        ElseIf action = "SELL" And cross10m = "BULLISH" Then
            exitFlag = True
            exitPx = curLTP
            exitOrderLTP = curLTP
            exitReason = "EMA_CROSS_10M_CONTRA"

        ' 6. Completed 5m Reverse EMA Cross (Col AC), exit at the current LTP.
        ElseIf action = "BUY" And cross5m = "BEARISH" Then
            exitFlag = True
            exitPx = curLTP
            exitOrderLTP = curLTP
            exitReason = "EMA_CROSS_5M (Reverse Exit)"
        ElseIf action = "SELL" And cross5m = "BULLISH" Then
            exitFlag = True
            exitPx = curLTP
            exitOrderLTP = curLTP
            exitReason = "EMA_CROSS_5M (Reverse Exit)"
        End If

        ' Apply Exit
        If exitFlag Then
            If exitOrderLTP > 0 Then
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " EXIT_TRIGGER reason=" & exitReason & _
                    " action=" & action & " cross10m=" & cross10m & " cross5m=" & cross5m & _
                    " ADX5m=" & IIf(adxAvailable, CStr(curADX), "NA") & " LTP=" & curLTP
            End If
            If ENABLE_BROKER_EXECUTION And V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) Then
                Dim rowQty As Long: rowQty = CLng(val(wsC.Range(CK_FILLED_QTY & r).value))
                If rowQty <= 0 Then rowQty = CLng(val(wsC.Range(CK_REQ_QTY & r).value))
                If rowQty > 0 Then
                    Dim exitReq As String
                    exitReq = NewV6RequestId("X", ticker)
                    wsC.Range(CK_EXEC_STATE & r).value = "EXIT_PENDING"
                    wsC.Range(CK_REQUEST_ID & r).value = exitReq
                    wsC.Range(CK_REQUEST_TYPE & r).value = "EXIT"
                    wsC.Range(CK_COMENTS & r).value = IIf(exitOrderLTP > 0, "EXIT_RETRY|", "") & _
                        exitReason & " | exit submitted; awaiting broker acknowledgement"
                    ' Exit at broker market price; retain exitOrderLTP only for the
                    ' trigger log / sheet fallback, not as an IOC limit instruction.
                    dispatchOk = DispatchBrokerOrder(IIf(action = "BUY", "SELL", "BUY"), ticker, rowQty, _
                        exitReason, exitReq, True)
                    If dispatchOk Then gV6LiveExitSubmittedThisRun = True
                    If Not dispatchOk Then
                        wsC.Range(CK_EXEC_STATE & r).value = "OPEN"
                        wsC.Range(CK_REQUEST_ID & r).ClearContents
                        wsC.Range(CK_COMENTS & r).value = IIf(exitOrderLTP > 0, "EXIT_RETRY|" & exitReason & " | ", "") & _
                            "Exit dispatcher failed to start; position remains open"
                    End If
                End If
            Else
                wsC.Range(CK_CP & r).value = exitPx
                wsC.Range(CK_CLOSE_TIME & r).value = Now
                If V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) Then
                    wsC.Range(CK_COMENTS & r).value = exitReason
                Else
                    wsC.Range(CK_COMENTS & r).value = "MONITOR ONLY | " & exitReason
                End If
                wsC.Range(CK_EXEC_STATE & r).value = "CLOSED"
            End If
        End If

NextTrade:
    Next r
End Sub


' Convert prior raw-cross boolean flags into simple filter outcomes.
Private Sub V6NormalizeRawCrossLog()
    On Error GoTo SafeExit

    Dim wsLog As Worksheet
    Dim lastRow As Long, r As Long
    Dim enabled As Boolean, thresholds As Variant
    Dim metric As Double, lowerBound As Double, upperBound As Double
    Set wsLog = ThisWorkbook.Worksheets("COCKPITV6_Log")
    lastRow = wsLog.cells(wsLog.Rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        If UCase$(Trim$(CStr(wsLog.cells(r, 19).value))) = "RAW_ENTRY_CROSS" Then
            If VarType(wsLog.cells(r, 22).value) = vbBoolean Then
                enabled = CBool(wsLog.cells(r, 22).value)
                thresholds = Split(CStr(wsLog.cells(r, 23).value), "-")
                If UBound(thresholds) = 1 And IsNumeric(thresholds(0)) And IsNumeric(thresholds(1)) _
                        And IsNumeric(wsLog.cells(r, 21).value) Then
                    lowerBound = CDbl(thresholds(0)): upperBound = CDbl(thresholds(1))
                    metric = CDbl(wsLog.cells(r, 21).value)
                    wsLog.cells(r, 22).value = IIf(metric >= lowerBound And metric <= upperBound, "PASS", "FAIL")
                Else
                    wsLog.cells(r, 22).value = "UNKNOWN"
                End If
            End If

            If VarType(wsLog.cells(r, 25).value) = vbBoolean Then
                enabled = CBool(wsLog.cells(r, 25).value)
                If IsNumeric(wsLog.cells(r, 24).value) And IsNumeric(wsLog.cells(r, 26).value) Then
                    wsLog.cells(r, 25).value = IIf(CDbl(wsLog.cells(r, 24).value) >= CDbl(wsLog.cells(r, 26).value), "PASS", "FAIL")
                Else
                    wsLog.cells(r, 25).value = "UNKNOWN"
                End If
            End If

            If VarType(wsLog.cells(r, 28).value) = vbBoolean Then
                enabled = CBool(wsLog.cells(r, 28).value)
                If enabled Then
                    wsLog.cells(r, 28).value = "FILTER_AND_RANK"
                Else
                    wsLog.cells(r, 28).value = IIf(RANK_CANDIDATES, "RANKING_INPUT", "TELEMETRY_ONLY")
                End If
            End If
            wsLog.Range(wsLog.cells(r, 30), wsLog.cells(r, 32)).ClearContents
        End If
    Next r

    wsLog.cells(1, 22).value = "RSIResult"
    wsLog.cells(1, 25).value = "ADXResult"
    wsLog.cells(1, 28).value = "RVOLRole"
    wsLog.Range("AD1:AF1").ClearContents
SafeExit:
End Sub

' Append each raw 10-minute entry cross with the values and gates used to qualify it.
' This is diagnostic only and does not change whether a candidate is traded.
Private Sub V6RawCrossSheetLog(ByVal sourceRow As Long, ByVal ticker As String, _
                               ByVal direction As String, ByVal ltp As Double, _
                               ByVal rsiVal As Double, ByVal adxVal As Double, ByVal rvolVal As Double, _
                               ByVal availableSlots As Long, ByVal qualified As Boolean, _
                               ByVal reason As String)
    On Error GoTo LogError

    Dim wsLog As Worksheet
    Dim NextRow As Long
    Dim errText As String
    Set wsLog = ThisWorkbook.Worksheets("COCKPITV6_Log")

    If Application.WorksheetFunction.CountA(wsLog.Range("A1:AC1")) = 0 Then
        wsLog.cells(1, 1).value = "RunDateTime"
        wsLog.cells(1, 2).value = "V6Row"
        wsLog.cells(1, 3).value = "Ticker"
        wsLog.cells(1, 4).value = "Action"
        wsLog.cells(1, 5).value = "EntryPx"
        wsLog.cells(1, 6).value = "TargetPx"
        wsLog.cells(1, 7).value = "StopPx"
        wsLog.cells(1, 8).value = "Route"
        wsLog.cells(1, 9).value = "ExecState"
        wsLog.cells(1, 10).value = "CloseDateTime"
        wsLog.cells(1, 11).value = "Sheet1Row"
        wsLog.cells(1, 12).value = "LTP"
        wsLog.cells(1, 13).value = "Cross5m_AC"
        wsLog.cells(1, 14).value = "FilledQty"
        wsLog.cells(1, 15).value = "RequestedQty"
        wsLog.cells(1, 16).value = "Decision"
        wsLog.cells(1, 17).value = "Reason"
        wsLog.cells(1, 18).value = "RequestID"
    End If
    ' The original decision logger may already have populated A:R. Add the
    ' raw-cross columns independently so an existing header row is upgraded.
    If Len(Trim$(CStr(wsLog.cells(1, 19).value))) = 0 Then
        wsLog.cells(1, 19).value = "RecordType"
        wsLog.cells(1, 20).value = "Cross10m_DJ"
        wsLog.cells(1, 21).value = "RSI"
        wsLog.cells(1, 22).value = "RSIResult"
        wsLog.cells(1, 23).value = "RSIRange"
        wsLog.cells(1, 24).value = "ADX"
        wsLog.cells(1, 25).value = "ADXResult"
        wsLog.cells(1, 26).value = "ADXMin"
        wsLog.cells(1, 27).value = "RVOL"
        wsLog.cells(1, 28).value = "RVOLRole"
        wsLog.cells(1, 29).value = "AvailableSlots"
    End If
    wsLog.cells(1, 20).value = "Cross10m_DJ"
    ' RSI and ADX are filter outcomes; RVOL is numeric ranking input, not an active gate.
    wsLog.cells(1, 22).value = "RSIResult"
    wsLog.cells(1, 25).value = "ADXResult"
    wsLog.cells(1, 28).value = "RVOLRole"

    NextRow = wsLog.cells(wsLog.Rows.count, 1).End(xlUp).row + 1
    If NextRow < 2 Then NextRow = 2
    wsLog.cells(NextRow, 1).value = Now
    wsLog.cells(NextRow, 3).value = ticker
    wsLog.cells(NextRow, 4).value = direction
    wsLog.cells(NextRow, 11).value = sourceRow
    wsLog.cells(NextRow, 12).value = ltp
    wsLog.cells(NextRow, 16).value = IIf(qualified, "QUALIFIED", "FILTERED")
    wsLog.cells(NextRow, 17).value = reason
    wsLog.cells(NextRow, 19).value = "RAW_ENTRY_CROSS"
    wsLog.cells(NextRow, 20).value = IIf(direction = "BUY", "BULLISH", "BEARISH")
    wsLog.cells(NextRow, 21).value = rsiVal
    If (direction = "BUY" And rsiVal >= BUY_RSI_MIN And rsiVal <= BUY_RSI_MAX) _
            Or (direction = "SELL" And rsiVal >= SELL_RSI_MIN And rsiVal <= SELL_RSI_MAX) Then
        wsLog.cells(NextRow, 22).value = "PASS"
    Else
        wsLog.cells(NextRow, 22).value = "FAIL"
    End If
    If direction = "BUY" Then
        wsLog.cells(NextRow, 23).value = CStr(BUY_RSI_MIN) & "-" & CStr(BUY_RSI_MAX)
    Else
        wsLog.cells(NextRow, 23).value = CStr(SELL_RSI_MIN) & "-" & CStr(SELL_RSI_MAX)
    End If
    wsLog.cells(NextRow, 24).value = adxVal
    If adxVal >= STOCK_ADX_MIN Then
        wsLog.cells(NextRow, 25).value = "PASS"
    Else
        wsLog.cells(NextRow, 25).value = "FAIL"
    End If
    wsLog.cells(NextRow, 26).value = STOCK_ADX_MIN
    wsLog.cells(NextRow, 27).value = rvolVal
    If RVOL_FILTER_ENABLED Then
        wsLog.cells(NextRow, 28).value = "FILTER_AND_RANK"
    Else
        wsLog.cells(NextRow, 28).value = IIf(RANK_CANDIDATES, "RANKING_INPUT", "TELEMETRY_ONLY")
    End If
    wsLog.cells(NextRow, 29).value = availableSlots
    wsLog.cells(NextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    Exit Sub

LogError:
    errText = Err.Description
    On Error Resume Next
    V6ReconcileLog "RAW_CROSS_LOG_FAILED row=" & sourceRow & " ticker=" & ticker & " error=" & errText
End Sub


' =====================================================================
' Private Sub  ScanForNewV6Trades
' Scans Sheet1 for fresh 10m EMA crosses (Col DJ), applies:
'   1. Configurable Sweet-Spot RSI Filter (BUY: 52-64, SELL: 38-47)
'   2. Stock's Own ADX Filter (Col CB >= 20 to avoid choppy consolidation)
'   (Direction Gate and RVOL Floor Filter REMOVED per user request)
'   Signals passing existing V6 filters are retained even when there are no live slots.
' =====================================================================
Private Sub ScanForNewV6Trades(wsC As Worksheet, wsSrc As Worksheet)
    ' 1. Count current open positions.
    '    Source of truth = FYERS positions/orderbook; fail closed if unavailable.
    Dim sheetOpen As Long: sheetOpen = CountOpenV6Trades(wsC)
    Dim apiOpen As Long: apiOpen = -1
    If gV6SnapshotReady Then apiOpen = CountSnapshotOpenPositions(gV6Snapshot)
    Dim pendingEntryCount As Long: pendingEntryCount = 0
    Dim unreflectedEntryCount As Long: unreflectedEntryCount = 0
    Dim unreflectedEntryDetails As String

    Dim openTradesCount As Long
    Dim canApplyLive As Boolean: canApplyLive = True
    wsC.cells(2, 14).ClearContents   ' clear prior API status note (N2)

    If ENABLE_BROKER_EXECUTION And apiOpen >= 0 Then
        pendingEntryCount = CountSnapshotPendingSymbols(gV6Snapshot)
        ' Reserve one slot for each broker pending symbol and for each V6 entry
        ' request not yet visible in the broker snapshot. This prevents a delayed
        ' asynchronous acknowledgement from making the next run over-admit.
        unreflectedEntryCount = CountUnreflectedEntryRequests(wsC, gV6Snapshot, unreflectedEntryDetails)
        openTradesCount = apiOpen + pendingEntryCount + unreflectedEntryCount
        V6ReconcileLog "SLOT_AUDIT api_open=" & apiOpen & " pending_symbols=" & pendingEntryCount & _
            " unreflected_entry_requests=" & unreflectedEntryCount & " counted_total=" & openTradesCount & _
            " max=" & MAX_CONCURRENT_POSITIONS & " rows=" & IIf(Len(unreflectedEntryDetails) > 0, unreflectedEntryDetails, "none")
        If Not gV6ReconciliationReady Or HasV6ExecutionMismatch(wsC) Then canApplyLive = False
    Else
        If ENABLE_BROKER_EXECUTION Then
            wsC.cells(2, 14).value = "WARNING: FYERS position snapshot unavailable. Live entries blocked; monitor-only signals continue."
            canApplyLive = False
            openTradesCount = 0
        Else
            openTradesCount = sheetOpen
        End If
    End If

    Dim availableSlots As Long: availableSlots = MAX_CONCURRENT_POSITIONS - openTradesCount
    If Not canApplyLive Then availableSlots = 0
    If availableSlots < 0 Then availableSlots = 0
    ' The external egress preflight is only needed if a live order could be admitted.
    Dim ipPreflightBlocked As Boolean: ipPreflightBlocked = False
    If ENABLE_BROKER_EXECUTION And canApplyLive And availableSlots > 0 Then
        If Not V6EntryIPReady(wsC) Then
            canApplyLive = False
            availableSlots = 0
            ipPreflightBlocked = True
        End If
    End If

    ' 2. Read Sheet1 in memory
    Dim srcLastRow As Long: srcLastRow = wsSrc.cells(wsSrc.Rows.count, 1).End(xlUp).row
    If srcLastRow < 4 Then Exit Sub

    Dim srcData As Variant
    srcData = wsSrc.Range("A1:ER" & srcLastRow).Value2

    Dim colLTP As Long:       colLTP = wsColToNum("T")
    Dim colEMA10 As Long: colEMA10 = wsColToNum("AS")
    Dim colEMA30 As Long: colEMA30 = wsColToNum("AT")
    Dim colEMA50 As Long: colEMA50 = wsColToNum("AU")
    Dim colEMA100 As Long: colEMA100 = wsColToNum("AV")
    Dim colEMA200 As Long: colEMA200 = wsColToNum("AW")
    Dim colCross10m As Long:  colCross10m = wsColToNum("DJ")
    Dim colRSI As Long:       colRSI = wsColToNum(RSI_SOURCE_COL)
    Dim colADX As Long:       colADX = wsColToNum(ADX_SOURCE_COL)
    Dim colRVOL As Long:      colRVOL = wsColToNum(RVOL_SOURCE_COL)

    ' Dynamic arrays to collect eligible candidate trades
    Dim candTicker() As String
    Dim candSrcRow() As Long
    Dim candAction() As String
    Dim candLTP() As Double
    Dim candEntryEMA() As Double
    Dim candEntryEMAPeriod() As Long
    Dim candRSI() As Double
    Dim candADX() As Double
    Dim candRVOL() As Double
    Dim candCount As Long: candCount = 0

    ' Convert legacy boolean filter flags in prior raw-cross rows to readable outcomes.
    V6NormalizeRawCrossLog

    Dim r As Long
    For r = 4 To srcLastRow
        Dim ticker As String: ticker = Trim$(CStr(srcData(r, 1)))
        If Len(ticker) = 0 Or IsIndexSymbol(ticker) Then GoTo NextStock

        ' Only Sheet1 DJ bullish/bearish events are raw entry crosses.
        Dim cross10m As String: cross10m = UCase$(Trim$(CStr(srcData(r, colCross10m))))
        Dim tradeDir As String: tradeDir = ""
        If cross10m = "BULLISH" Then
            tradeDir = "BUY"
        ElseIf cross10m = "BEARISH" Then
            tradeDir = "SELL"
        End If
        If Len(tradeDir) = 0 Then GoTo NextStock

        Dim ltp As Double: ltp = CDbl(val(srcData(r, colLTP)))
        Dim entryEMAPeriod As Long: entryEMAPeriod = 0
        Dim entryEMA As Double
        entryEMA = SelectNearestPullbackEMA(ltp, tradeDir, _
            CDbl(val(srcData(r, colEMA10))), CDbl(val(srcData(r, colEMA30))), _
            CDbl(val(srcData(r, colEMA50))), CDbl(val(srcData(r, colEMA100))), _
            CDbl(val(srcData(r, colEMA200))), entryEMAPeriod)
        Dim rsiVal As Double: rsiVal = CDbl(val(srcData(r, colRSI)))
        Dim adxVal As Double: adxVal = CDbl(val(srcData(r, colADX)))
        Dim rvolVal As Double: rvolVal = CDbl(val(srcData(r, colRVOL)))
        Dim eligible As Boolean: eligible = True
        Dim qualReason As String: qualReason = "Passed qualification filters"

        If ltp <= 0 Then
            eligible = False: qualReason = "Filtered: LTP missing or invalid"
        ElseIf entryEMA <= 0 Then
            eligible = False: qualReason = "Filtered: no 5m EMA (10/30/50/100/200) on valid pullback side of LTP"
        ElseIf IsTickerAlreadyOpen(wsC, ticker) Then
            eligible = False: qualReason = "Filtered: ticker already has an open V6 row"
        ElseIf ENABLE_BROKER_EXECUTION And gV6SnapshotReady And IsTickerOpenInBrokerSnapshot(gV6Snapshot, ticker) Then
            eligible = False: qualReason = "Filtered: ticker already has a FYERS position or pending order"
        ElseIf RSI_FILTER_ENABLED And ((tradeDir = "BUY" And (rsiVal < BUY_RSI_MIN Or rsiVal > BUY_RSI_MAX)) _
                Or (tradeDir = "SELL" And (rsiVal < SELL_RSI_MIN Or rsiVal > SELL_RSI_MAX))) Then
            eligible = False: qualReason = "Filtered: RSI outside configured range"
        ElseIf ADX_FILTER_ENABLED And adxVal < STOCK_ADX_MIN Then
            eligible = False: qualReason = "Filtered: ADX below configured minimum"
        ElseIf RVOL_FILTER_ENABLED And rvolVal < 0.5 Then
            eligible = False: qualReason = "Filtered: RVOL below 0.5"
        End If

        V6RawCrossSheetLog r, ticker, tradeDir, ltp, rsiVal, adxVal, rvolVal, availableSlots, eligible, qualReason
        If Not eligible Then GoTo NextStock

        ' Eligible candidate found -- add to candidate pool
        candCount = candCount + 1
        ReDim Preserve candTicker(1 To candCount)
        ReDim Preserve candSrcRow(1 To candCount)
        ReDim Preserve candAction(1 To candCount)
        ReDim Preserve candLTP(1 To candCount)
        ReDim Preserve candEntryEMA(1 To candCount)
        ReDim Preserve candEntryEMAPeriod(1 To candCount)
        ReDim Preserve candRSI(1 To candCount)
        ReDim Preserve candADX(1 To candCount)
        ReDim Preserve candRVOL(1 To candCount)

        candTicker(candCount) = ticker
        candSrcRow(candCount) = r
        candAction(candCount) = tradeDir
        candLTP(candCount) = ltp
        candEntryEMA(candCount) = entryEMA
        candEntryEMAPeriod(candCount) = entryEMAPeriod
        candRSI(candCount) = rsiVal
        candADX(candCount) = adxVal
        candRVOL(candCount) = rvolVal

NextStock:
    Next r

    If candCount = 0 Then Exit Sub

    ' 3. Rank candidates: RVOL (primary), ADX (tie-break), ticker (deterministic).
    '    Only the top candidates within available capacity can be applied live;
    '    all other qualifying candidates are retained as monitor-only rows.
    If candCount > 1 And RANK_CANDIDATES Then
        Dim i As Long, j As Long
        Dim tmpTicker As String, tmpAction As String
        Dim tmpSrcRow As Long, tmpLTP As Double, tmpEntryEMA As Double, tmpRSI As Double, tmpADX As Double, tmpRVOL As Double
        Dim tmpEntryEMAPeriod As Long
        Dim swapIt As Boolean
        For i = 1 To candCount - 1
            For j = i + 1 To candCount
                swapIt = False
                If candRVOL(j) > candRVOL(i) Then
                    swapIt = True
                ElseIf candRVOL(j) = candRVOL(i) Then
                    If candADX(j) > candADX(i) Then
                        swapIt = True
                    ElseIf candADX(j) = candADX(i) Then
                        If StrComp(candTicker(j), candTicker(i), vbTextCompare) < 0 Then swapIt = True
                    End If
                End If
                If swapIt Then
                    tmpTicker = candTicker(i): candTicker(i) = candTicker(j): candTicker(j) = tmpTicker
                    tmpSrcRow = candSrcRow(i): candSrcRow(i) = candSrcRow(j): candSrcRow(j) = tmpSrcRow
                    tmpAction = candAction(i): candAction(i) = candAction(j): candAction(j) = tmpAction
                    tmpLTP = candLTP(i): candLTP(i) = candLTP(j): candLTP(j) = tmpLTP
                    tmpEntryEMA = candEntryEMA(i): candEntryEMA(i) = candEntryEMA(j): candEntryEMA(j) = tmpEntryEMA
                    tmpEntryEMAPeriod = candEntryEMAPeriod(i): candEntryEMAPeriod(i) = candEntryEMAPeriod(j): candEntryEMAPeriod(j) = tmpEntryEMAPeriod
                    tmpRSI = candRSI(i): candRSI(i) = candRSI(j): candRSI(j) = tmpRSI
                    tmpADX = candADX(i): candADX(i) = candADX(j): candADX(j) = tmpADX
                    tmpRVOL = candRVOL(i): candRVOL(i) = candRVOL(j): candRVOL(j) = tmpRVOL
                End If
            Next j
        Next i
    End If

    ' 4. Admit top candidates into available slots
    Dim lastDestRow As Long
    lastDestRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    If lastDestRow < DATA_START_ROW Then lastDestRow = DATA_START_ROW - 1

    Dim destRow As Long: destRow = lastDestRow
    Dim numToAdmit As Long
    numToAdmit = IIf(candCount < availableSlots, candCount, availableSlots)

    ' 4a. Dynamic sizing (Python): live available capital / free slots / per-stock intraday margin.
    '     One synchronous call sizes ALL admitted candidates; returns 0 where it cannot size.
    Dim symList As String
    Dim k As Long
    For k = 1 To numToAdmit
        If Len(symList) > 0 Then symList = symList & ","
        symList = symList & candTicker(k) & ":" & candAction(k)
    Next k

    Dim qtyArr() As Long
    ReDim qtyArr(1 To candCount)
    For k = 1 To candCount: qtyArr(k) = 0: Next k
    If numToAdmit > 0 And ENABLE_BROKER_EXECUTION Then
        SizeCandidates symList, availableSlots, qtyArr, numToAdmit
        If numToAdmit < candCount Then ReDim Preserve qtyArr(1 To candCount)
    ElseIf numToAdmit > 0 Then
        ' Preserve existing non-broker/paper behavior when live execution is disabled.
        For k = 1 To numToAdmit: qtyArr(k) = 1: Next k
    End If

    For k = 1 To candCount
        Dim applyToFyers As Boolean
        applyToFyers = (k <= numToAdmit And qtyArr(k) > 0)
        Dim qty As Long: qty = qtyArr(k)
        If Not applyToFyers Then qty = 0 ' Monitor-only rows are virtual and do not need broker quantity.
        destRow = destRow + 1

        Dim slPx As Double
        Dim tgtPx As Double
        If candAction(k) = "BUY" Then
            slPx = Round(candEntryEMA(k) * (1# - SL_PCT), 2)
            tgtPx = Round(candEntryEMA(k) * (1# + TGT_PCT), 2)
        Else
            slPx = Round(candEntryEMA(k) * (1# + SL_PCT), 2)
            tgtPx = Round(candEntryEMA(k) * (1# - TGT_PCT), 2)
        End If

        WriteNewV6Row wsC, wsSrc, candSrcRow(k), destRow, candTicker(k), candAction(k), _
                      candLTP(k), candEntryEMA(k), tgtPx, slPx, qty, candRSI(k), candADX(k), candRVOL(k), candEntryEMAPeriod(k)
        wsC.Range(CK_APPLY_TO_FYERS & destRow).value = applyToFyers
        If applyToFyers And ENABLE_BROKER_EXECUTION Then
            wsC.Range(CK_ORIGIN & destRow).value = "V6"
        ElseIf applyToFyers Then
            wsC.Range(CK_ORIGIN & destRow).value = "PAPER"
        Else
            wsC.Range(CK_ORIGIN & destRow).value = "MONITOR"
        End If
        Dim decisionReason As String
        If applyToFyers Then
            decisionReason = "ADMITTED"
        ElseIf ENABLE_BROKER_EXECUTION And Not gV6SnapshotReady Then
            decisionReason = "BLOCKED: broker snapshot unavailable"
        ElseIf ENABLE_BROKER_EXECUTION And Not gV6ReconciliationReady Then
            decisionReason = "BLOCKED: unresolved broker reconciliation"
        ElseIf ipPreflightBlocked Then
            decisionReason = "BLOCKED: IP preflight failed"
        ElseIf ENABLE_BROKER_EXECUTION And availableSlots <= 0 And unreflectedEntryCount > 0 Then
            decisionReason = "BLOCKED: unreflected entry acknowledgements reserve capacity"
        ElseIf ENABLE_BROKER_EXECUTION And availableSlots <= 0 Then
            decisionReason = "MONITOR: no free broker slot"
        ElseIf k > numToAdmit Then
            decisionReason = "MONITOR: ranked outside current free-slot capacity"
        Else
            decisionReason = "MONITOR: live sizing returned zero"
        End If
        V6ReconcileLog "ENTRY_DECISION ticker=" & candTicker(k) & " side=" & candAction(k) & _
            " rank=" & k & "/" & candCount & " broker_open=" & openTradesCount & _
            " free_slots=" & availableSlots & " entry_ema=" & candEntryEMAPeriod(k) & _
            " entry_limit=" & candEntryEMA(k) & " qty=" & qty & " apply=" & applyToFyers & _
            " reason=" & decisionReason

        If applyToFyers And ENABLE_BROKER_EXECUTION Then
            Dim entryReq As String
            entryReq = NewV6RequestId("E", candTicker(k))
            wsC.Range(CK_EXEC_STATE & destRow).value = "SUBMITTING"
            wsC.Range(CK_REQUEST_ID & destRow).value = entryReq
            wsC.Range(CK_REQUEST_TYPE & destRow).value = "ENTRY"
            wsC.Range(CK_REQ_QTY & destRow).value = qty
            If Not DispatchBrokerOrder(candAction(k), candTicker(k), qty, "V6ENTRY", entryReq, False, SL_PCT * 100#, TGT_PCT * 100#, 0#, candEntryEMA(k)) Then
                wsC.Range(CK_EXEC_STATE & destRow).value = "REJECTED_NO_FILL"
                wsC.Range(CK_CLOSE_TIME & destRow).value = Now
                wsC.Range(CK_COMENTS & destRow).value = "Dispatcher failed to start; order was not sent"
                wsC.Range(CK_REQUEST_ID & destRow).ClearContents
            End If
        ElseIf applyToFyers Then
            wsC.Range(CK_EXEC_STATE & destRow).value = "MONITOR_ENTRY_PENDING"
            wsC.Range(CK_FILLED_QTY & destRow).value = qty
            wsC.Range(CK_REQ_QTY & destRow).value = qty
            wsC.Range(CK_COMENTS & destRow).value = "PAPER | 5m EMA" & candEntryEMAPeriod(k) & " limit; waiting up to 3 completed 5m bars"
        Else
            wsC.Range(CK_EXEC_STATE & destRow).value = "MONITOR_ENTRY_PENDING"
            wsC.Range(CK_FILLED_QTY & destRow).value = 0
            wsC.Range(CK_REQ_QTY & destRow).value = 0
            wsC.Range(CK_COMENTS & destRow).value = "MONITOR ONLY | 5m EMA" & candEntryEMAPeriod(k) & " limit; waiting up to 3 completed 5m bars"
            If k > numToAdmit Then
                If ENABLE_BROKER_EXECUTION And Not gV6SnapshotReady Then
                    wsC.Range(CK_COMENTS & destRow).value = "MONITOR ONLY | FYERS snapshot unavailable; live entry blocked"
                ElseIf ENABLE_BROKER_EXECUTION And Not gV6ReconciliationReady Then
                    wsC.Range(CK_COMENTS & destRow).value = "MONITOR ONLY | unresolved FYERS reconciliation; live entry blocked"
                ElseIf ipPreflightBlocked Then
                    wsC.Range(CK_COMENTS & destRow).value = "MONITOR ONLY | IP preflight failed; live entry blocked"
                ElseIf availableSlots <= 0 And unreflectedEntryCount > 0 Then
                    wsC.Range(CK_COMENTS & destRow).value = "MONITOR ONLY | unresolved entry acknowledgement(s) reserve capacity; see SLOT_AUDIT log"
                ElseIf availableSlots <= 0 Then
                    wsC.Range(CK_COMENTS & destRow).value = "MONITOR ONLY | no free FYERS slot"
                Else
                    wsC.Range(CK_COMENTS & destRow).value = "MONITOR ONLY | ranked outside current free-slot capacity"
                End If
            Else
                wsC.Range(CK_COMENTS & destRow).value = "MONITOR ONLY | live sizing unavailable; signal retained for lifecycle tracking"
            End If
        End If
    Next k

    ' Every qualified candidate is retained on this sheet. AC=TRUE applies to FYERS;
    ' AC=FALSE follows the same local lifecycle without using a live position slot.
End Sub


' =====================================================================
' Private Sub  WriteNewV6Row
' Formats and appends a newly admitted trade into COCKPIT_V6
' =====================================================================
Private Sub WriteNewV6Row(wsC As Worksheet, wsSrc As Worksheet, _
        srcRow As Long, destRow As Long, ticker As String, action As String, _
        ltp As Double, entryLimitPx As Double, tgtPx As Double, slPx As Double, qty As Long, _
        rsiVal As Double, adxVal As Double, rvolVal As Double, Optional entryEMAPeriod As Long = 0)

    wsC.cells(destRow, RangeRef(CK_TICKER)).value = ticker
    wsC.cells(destRow, RangeRef(CK_ACTION)).value = action
    wsC.cells(destRow, RangeRef(CK_LTP)).value = ltp

    ' Sector lookup formula
    wsC.cells(destRow, RangeRef(CK_SECTOR)).Formula = _
        "=IF(XLOOKUP(A" & destRow & ",Sheet1!A:A,Sheet1!BM:BM)<>0," & _
        "XLOOKUP(A" & destRow & ",Sheet1!A:A,Sheet1!BM:BM),"""")"

    ' Pivots S3 to R3
    wsC.cells(destRow, RangeRef(CK_S3)).value = wsSrc.cells(srcRow, "K").value
    wsC.cells(destRow, RangeRef(CK_S2)).value = wsSrc.cells(srcRow, "L").value
    wsC.cells(destRow, RangeRef(CK_S1)).value = wsSrc.cells(srcRow, "M").value
    wsC.cells(destRow, RangeRef(CK_PP)).value = wsSrc.cells(srcRow, "N").value
    wsC.cells(destRow, RangeRef(CK_R1)).value = wsSrc.cells(srcRow, "O").value
    wsC.cells(destRow, RangeRef(CK_R2)).value = wsSrc.cells(srcRow, "P").value
    wsC.cells(destRow, RangeRef(CK_R3)).value = wsSrc.cells(srcRow, "Q").value

    ' Multi-TF RSI
    wsC.cells(destRow, RangeRef(CK_RSI_5M)).value = wsSrc.cells(srcRow, "U").value
    wsC.cells(destRow, RangeRef(CK_RSI_15M)).value = wsSrc.cells(srcRow, "V").value
    wsC.cells(destRow, RangeRef(CK_RSI_1H)).value = wsSrc.cells(srcRow, "W").value
    wsC.cells(destRow, RangeRef(CK_RSI_1D)).value = wsSrc.cells(srcRow, "X").value

    wsC.cells(destRow, RangeRef(CK_SUPERTREND)).value = wsSrc.cells(srcRow, "BT").value
    wsC.cells(destRow, RangeRef(CK_CROSS_10_30)).value = wsSrc.cells(srcRow, "DJ").value ' 10m Cross from Col DJ (ENTRY)
    wsC.cells(destRow, RangeRef(CK_CROSS_10_50)).value = wsSrc.cells(srcRow, "AG").value

    ' Execution levels
    wsC.cells(destRow, RangeRef(CK_ENT)).value = entryLimitPx
    wsC.cells(destRow, RangeRef(CK_TGT)).value = tgtPx
    wsC.cells(destRow, RangeRef(CK_SL)).value = slPx
    wsC.cells(destRow, RangeRef(CK_CP)).value = ltp
    ' W begins as the row/signal recorded time. It is replaced with the fill or
    ' simulated-entry time when the row transitions into an open trade.
    wsC.cells(destRow, RangeRef(CK_OPEN_TIME)).value = Now

    ' Col Y and Z are workbook formula-owned; VBA must never write to them.

    Dim entryEMARef As String
    If entryEMAPeriod > 0 Then
        entryEMARef = "5m EMA" & entryEMAPeriod
    Else
        entryEMARef = "entry"
    End If
    wsC.cells(destRow, RangeRef(CK_COMENTS)).value = _
        "V6_ENTRY | 10m Cross | " & entryEMARef & " limit=" & Format(entryLimitPx, "0.00") & " | RSI=" & Format(rsiVal, "0.0") & " | ADX=" & Format(adxVal, "0.0") & " | RVOL=" & Format(rvolVal, "0.00") & " | Qty=" & qty & " | TGT=" & Format(tgtPx, "0.00") & " | SL=" & Format(slPx, "0.00")
    wsC.cells(destRow, RangeRef(CK_CANDLE_COUNT)).value = 0
    wsC.cells(destRow, RangeRef(CK_PEAK_PRICE)).value = ltp
    wsC.cells(destRow, RangeRef(CK_BEST_MFE_PCT)).value = 0#

    Dim entryBarStamp As Date
    entryBarStamp = DateValue(Now) + TimeSerial(Hour(Now), (Minute(Now) \ 5) * 5, 0)
    wsC.cells(destRow, RangeRef(CK_LAST_BAR_TS)).value = CDbl(entryBarStamp)
End Sub


' Choose the nearest 5m EMA on the pullback side of price at signal time.
' BUY: highest eligible EMA below LTP. SELL: lowest eligible EMA above LTP.
' The resulting limit is stored once in the row/order and is not re-anchored.
Private Function SelectNearestPullbackEMA(ByVal ltp As Double, ByVal action As String, _
        ByVal ema10 As Double, ByVal ema30 As Double, ByVal ema50 As Double, _
        ByVal ema100 As Double, ByVal ema200 As Double, ByRef selectedPeriod As Long) As Double
    Dim values As Variant: values = Array(ema10, ema30, ema50, ema100, ema200)
    Dim periods As Variant: periods = Array(10, 30, 50, 100, 200)
    Dim i As Long, candidate As Double, best As Double
    selectedPeriod = 0
    best = 0#

    For i = LBound(values) To UBound(values)
        candidate = CDbl(values(i))
        If candidate > 0# Then
            If UCase$(action) = "BUY" Then
                If candidate < ltp And candidate > best Then
                    best = candidate
                    selectedPeriod = CLng(periods(i))
                End If
            ElseIf UCase$(action) = "SELL" Then
                If candidate > ltp And (best = 0# Or candidate < best) Then
                    best = candidate
                    selectedPeriod = CLng(periods(i))
                End If
            End If
        End If
    Next i
    SelectNearestPullbackEMA = best
End Function


Private Sub ManageMonitorPendingEntry(wsC As Worksheet, wsSrc As Worksheet, ByVal rowNum As Long, ByVal ticker As String)
    Dim srcRow As Long: srcRow = FindSheet1Row(wsSrc, ticker)
    If srcRow < 4 Then Exit Sub

    ' Keep CP (Col V) current while a monitor-only entry waits for its limit price.
    Dim pendingLtpCell As Variant: pendingLtpCell = wsSrc.cells(srcRow, "T").value
    If Not IsError(pendingLtpCell) Then
        If IsNumeric(pendingLtpCell) Then
            If CDbl(pendingLtpCell) > 0 Then wsC.Range(CK_CP & rowNum).value = CDbl(pendingLtpCell)
        End If
    End If

    Dim barStamp As Date
    barStamp = DateValue(Now) + TimeSerial(Hour(Now), (Minute(Now) \ 5) * 5, 0)
    Dim lastStamp As Double: lastStamp = CDbl(val(wsC.Range(CK_LAST_BAR_TS & rowNum).value))
    If CDbl(barStamp) <= lastStamp Then Exit Sub

    Dim barsWaited As Long: barsWaited = CLng(val(wsC.Range(CK_CANDLE_COUNT & rowNum).value)) + 1
    wsC.Range(CK_CANDLE_COUNT & rowNum).value = barsWaited
    wsC.Range(CK_LAST_BAR_TS & rowNum).value = CDbl(barStamp)

    Dim entryPx As Double: entryPx = CDbl(val(wsC.Range(CK_ENT & rowNum).value))
    Dim action As String: action = UCase$(Trim$(CStr(wsC.Range(CK_ACTION & rowNum).value)))
    Dim barHigh As Double: barHigh = CDbl(val(wsSrc.cells(srcRow, "DR").value))
    Dim barLow As Double: barLow = CDbl(val(wsSrc.cells(srcRow, "DS").value))
    Dim touched As Boolean
    If action = "BUY" Then
        touched = (barLow > 0 And barLow <= entryPx)
    ElseIf action = "SELL" Then
        touched = (barHigh > 0 And barHigh >= entryPx)
    End If

    If touched Then
        wsC.Range(CK_EXEC_STATE & rowNum).value = "OPEN"
        wsC.Range(CK_OPEN_TIME & rowNum).value = Now
        wsC.Range(CK_CP & rowNum).value = CDbl(val(wsSrc.cells(srcRow, "T").value))
        wsC.Range(CK_CANDLE_COUNT & rowNum).value = 0
        wsC.Range(CK_PEAK_PRICE & rowNum).value = wsC.Range(CK_CP & rowNum).value
        wsC.Range(CK_BEST_MFE_PCT & rowNum).value = 0#
        wsC.Range(CK_COMENTS & rowNum).value = "MONITOR ONLY | selected 5m EMA limit touched; lifecycle started at limit price"
        Exit Sub
    End If

    If barsWaited >= ENTRY_LIMIT_WAIT_BARS Then
        wsC.Range(CK_EXEC_STATE & rowNum).value = "CANCELLED_NO_FILL"
        wsC.Range(CK_CLOSE_TIME & rowNum).value = Now
        wsC.Range(CK_COMENTS & rowNum).value = "MONITOR ONLY | selected 5m EMA was not reached within 3 completed 5m bars; cancelled"
    End If
End Sub


' =====================================================================
'  HELPER FUNCTIONS & LOOKUPS
' =====================================================================
Private Function CountOpenV6Trades(wsC As Worksheet) As Long
    Dim count As Long: count = 0
    Dim lastRow As Long
    lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    If lastRow < DATA_START_ROW Then CountOpenV6Trades = 0: Exit Function

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim ticker As String: ticker = Trim$(CStr(wsC.Range(CK_TICKER & r).value))
        Dim closeTime As Variant: closeTime = wsC.Range(CK_CLOSE_TIME & r).value
        If Len(ticker) > 0 And (IsEmpty(closeTime) Or Len(Trim$(CStr(closeTime))) = 0) Then
            If Not V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) Then GoTo NextCountRow
            Dim state As String: state = UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & r).value)))
            If state = "" Or state = "OPEN" Or state = "MONITOR_ENTRY_PENDING" Or state = "SUBMITTING" Or state = "ENTRY_CANCEL_PENDING" Or state = "EXIT_PENDING" Or state = "EXIT_ACK_WAIT" Or state = "MISMATCH" Then count = count + 1
        End If
NextCountRow:
    Next r
    CountOpenV6Trades = count
End Function


Private Function CountOpenBrokerRows(wsC As Worksheet) As Long
    CountOpenBrokerRows = CountOpenV6Trades(wsC)
End Function


Private Function CountOpenMonitorRows(wsC As Worksheet) As Long
    Dim lastRow As Long: lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim ticker As String: ticker = Trim$(CStr(wsC.Range(CK_TICKER & r).value))
        Dim closeVal As Variant: closeVal = wsC.Range(CK_CLOSE_TIME & r).value
        If Len(ticker) > 0 And (IsEmpty(closeVal) Or Len(Trim$(CStr(closeVal))) = 0) Then
            If Not V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) Then
                Dim state As String: state = UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & r).value)))
                If state = "MONITOR_ENTRY_PENDING" Or state = "OPEN" Then CountOpenMonitorRows = CountOpenMonitorRows + 1
            End If
        End If
    Next r
End Function


Private Function CountSnapshotPendingSymbols(ByVal snapshot As String) As Long
    Dim p As Long, marker As Long, qEnd As Long, qStart As Long, ticker As String
    Dim netQty As Double, pending As Double
    p = 1
    Do
        marker = InStr(p, snapshot, ":{""net_qty"":", vbTextCompare)
        If marker = 0 Then Exit Do
        qEnd = InStrRev(snapshot, """", marker - 1)
        qStart = InStrRev(snapshot, """", qEnd - 1)
        If qStart > 0 And qEnd > qStart Then
            ticker = Mid$(snapshot, qStart + 1, qEnd - qStart - 1)
            netQty = SnapshotNumberForSymbol(snapshot, ticker, "net_qty", 0#)
            pending = SnapshotNumberForSymbol(snapshot, ticker, "pending", 0#)
            If Left$(ticker, 2) <> "__" And netQty = 0 And pending > 0 Then
                CountSnapshotPendingSymbols = CountSnapshotPendingSymbols + 1
            End If
        End If
        p = marker + 1
    Loop
End Function


Private Function V6UseLocalTakeProfit(wsC As Worksheet, ByVal rowNum As Long) As Boolean
    If Not (ENABLE_BROKER_EXECUTION And V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & rowNum).value)) Then
        V6UseLocalTakeProfit = True
    Else
        V6UseLocalTakeProfit = Not CBoolSafe(wsC.Range(CK_HAS_BROKER_TP & rowNum).value)
    End If
End Function


Private Function V6UseLocalStopLoss(wsC As Worksheet, ByVal rowNum As Long) As Boolean
    If Not (ENABLE_BROKER_EXECUTION And V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & rowNum).value)) Then
        V6UseLocalStopLoss = True
    Else
        V6UseLocalStopLoss = Not CBoolSafe(wsC.Range(CK_HAS_BROKER_SL & rowNum).value)
    End If
End Function


Private Function CBoolSafe(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbBoolean Then
        CBoolSafe = CBool(value)
    Else
        CBoolSafe = (UCase$(Trim$(CStr(value))) = "TRUE" Or Trim$(CStr(value)) = "1")
    End If
End Function


Private Function CountPendingV6Entries(wsC As Worksheet, ByVal snapshot As String) As Long
    Dim lastRow As Long: lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    If lastRow < DATA_START_ROW Then Exit Function
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If Not V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) Then GoTo NextPendingRow
        Dim state As String: state = UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & r).value)))
        Dim reqType As String: reqType = UCase$(Trim$(CStr(wsC.Range(CK_REQUEST_TYPE & r).value)))
        If (state = "SUBMITTING" Or state = "ENTRY_CANCEL_PENDING") And reqType = "ENTRY" Then
            Dim ticker As String: ticker = Trim$(CStr(wsC.Range(CK_TICKER & r).value))
            If Len(ticker) > 0 And SnapshotNumberForSymbol(snapshot, ticker, "net_qty", 0#) = 0 _
                    And SnapshotNumberForSymbol(snapshot, ticker, "pending", 0#) > 0 Then
                CountPendingV6Entries = CountPendingV6Entries + 1
            End If
        End If
NextPendingRow:
    Next r
End Function


Private Function HasV6ExecutionMismatch(wsC As Worksheet) As Boolean
    Dim lastRow As Long: lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & r).value))) = "MISMATCH" Then
            HasV6ExecutionMismatch = True
            Exit Function
        End If
    Next r
End Function


Private Sub ReconcileBrokerPositions(wsC As Worksheet, wsSrc As Worksheet, ByVal snapshot As String)
    On Error GoTo ReconcileError
    Dim lastRow As Long: lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    V6ReconcileLog "BEGIN lastRow=" & lastRow
    Dim r As Long, ticker As String
    For r = DATA_START_ROW To lastRow
        On Error GoTo ReconcileError
        ticker = vbNullString
        ticker = UCase$(Trim$(CStr(wsC.Range(CK_TICKER & r).value)))
        If Len(ticker) = 0 Then GoTo NextRow
        gV6ErrorContext = "RECONCILE row=" & r & " ticker=" & ticker
        Dim state As String: state = UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & r).value)))
        Dim rowComment As String: rowComment = Trim$(CStr(wsC.Range(CK_COMENTS & r).value))
        Dim closeVal As Variant: closeVal = wsC.Range(CK_CLOSE_TIME & r).value
        Dim recoverRejectedAck As Boolean
        recoverRejectedAck = False
        If state = "REJECTED_NO_FILL" And Left$(UCase$(rowComment), 9) = "REJECTED:" Then
            If IsEmpty(closeVal) Or Len(Trim$(CStr(closeVal))) = 0 Then
                ' A user may clear CloseTime to request broker reconciliation again.
                recoverRejectedAck = True
            ElseIf IsDate(closeVal) Then
                recoverRejectedAck = (DateValue(CDate(closeVal)) = Date)
            End If
        End If
        If Not IsEmpty(closeVal) And Len(Trim$(CStr(closeVal))) > 0 And Not recoverRejectedAck Then GoTo NextRow
        ' Monitor-only rows are virtual and must never be reconciled against broker positions.
        If Not V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) Then
            V6ReconcileLog "row=" & r & " ticker=" & ticker & " SKIP monitor-only flag=" & CStr(wsC.Range(CK_APPLY_TO_FYERS & r).value)
            GoTo NextRow
        End If

        If state = "" Then
            state = "OPEN" ' Existing rows predate broker-linked state columns.
            wsC.Range(CK_EXEC_STATE & r).value = "OPEN"
            V6ReconcileLog "row=" & r & " ticker=" & ticker & " LEGACY_OPEN_STATE_INITIALIZED"
        End If
        If state <> "OPEN" And state <> "SUBMITTING" And state <> "ENTRY_CANCEL_PENDING" And state <> "EXIT_PENDING" And state <> "EXIT_ACK_WAIT" And state <> "MISMATCH" _
                And Not recoverRejectedAck Then
            V6ReconcileLog "row=" & r & " ticker=" & ticker & " SKIP state=" & state
            GoTo NextRow
        End If

        If state = "SUBMITTING" Or state = "ENTRY_CANCEL_PENDING" Then
            Dim pendingRequestType As Variant: pendingRequestType = wsC.Range(CK_REQUEST_TYPE & r).value
            If Not IsError(pendingRequestType) Then
                If UCase$(Trim$(CStr(pendingRequestType))) = "ENTRY" Then
                    Dim pendingEntrySrcRow As Long: pendingEntrySrcRow = FindSheet1Row(wsSrc, ticker)
                    If pendingEntrySrcRow >= 4 Then
                        Dim pendingEntryLtpCell As Variant: pendingEntryLtpCell = wsSrc.cells(pendingEntrySrcRow, "T").value
                        If Not IsError(pendingEntryLtpCell) Then
                            If IsNumeric(pendingEntryLtpCell) Then
                                If CDbl(pendingEntryLtpCell) > 0 Then wsC.Range(CK_CP & r).value = CDbl(pendingEntryLtpCell)
                            End If
                        End If
                    End If
                End If
            End If
        End If

        ' The live broker snapshot is authoritative. Capture it before applying an
        ' asynchronous result, because an old/reordered REJECTED acknowledgement
        ' must never close an Excel row while FYERS shows a real position.
        Dim netQty As Long: netQty = CLng(SnapshotNumberForSymbol(snapshot, ticker, "net_qty", 0#))
        Dim pending As Long: pending = CLng(SnapshotNumberForSymbol(snapshot, ticker, "pending", 0#))
        V6ReconcileLog "row=" & r & " ticker=" & ticker & " BROKER_SNAPSHOT netQty=" & netQty & " pending=" & pending & " state=" & state

        ' FYERS can return the exact saved order as Pending/Transit while its
        ' per-symbol pending count is absent from the position snapshot. Use the
        ' order ID as the tie-breaker so an accepted limit order is neither lost
        ' nor left waiting without its three-bar cancellation lifecycle.
        Dim reconcileReqType As String: reconcileReqType = UCase$(Trim$(CStr(wsC.Range(CK_REQUEST_TYPE & r).value)))
        Dim reconcileOrderId As String: reconcileOrderId = Trim$(CStr(wsC.Range(CK_ORDER_ID & r).value))
        Dim reconcileOrderStatus As Long: reconcileOrderStatus = -1
        Dim reconcileOrderFilled As Long: reconcileOrderFilled = 0
        If reconcileReqType = "ENTRY" And Len(reconcileOrderId) > 0 Then
            reconcileOrderStatus = SnapshotOrderStatus(snapshot, reconcileOrderId)
            reconcileOrderFilled = CLng(SnapshotOrderNumber(snapshot, reconcileOrderId, "filled_qty"))
            If netQty = 0 And reconcileOrderFilled = 0 _
                    And (reconcileOrderStatus = 1 Or reconcileOrderStatus = 5 Or reconcileOrderStatus = 7) Then
                ' The exact entry order is terminal and has no fill. Close this
                ' attempt even if another order for the same symbol is pending;
                ' that other broker order is still counted separately by symbol.
                wsC.Range(CK_EXEC_STATE & r).value = "CANCELLED_NO_FILL"
                wsC.Range(CK_CLOSE_TIME & r).value = Now
                wsC.Range(CK_REQUEST_ID & r).ClearContents
                wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                wsC.Range(CK_COMENTS & r).value = "FYERS orderbook confirms terminal no-fill status=" & reconcileOrderStatus
                V6ReconcileLog "row=" & r & " ticker=" & ticker & _
                    " ENTRY_TERMINAL_NO_FILL order=" & reconcileOrderId & _
                    " status=" & reconcileOrderStatus
                GoTo NextRow
            End If
            If netQty = 0 And pending = 0 And reconcileOrderFilled = 0 _
                    And (reconcileOrderStatus = 4 Or reconcileOrderStatus = 6) Then
                pending = 1
                V6ReconcileLog "row=" & r & " ticker=" & ticker & _
                    " ORDER_ID_PENDING status=" & reconcileOrderStatus & _
                    "; per-symbol pending count was zero, using order ID as authority"
            End If
        End If

        If recoverRejectedAck Then
            If netQty = 0 And pending = 0 Then
                If IsEmpty(closeVal) Or Len(Trim$(CStr(closeVal))) = 0 Then
                    wsC.Range(CK_CLOSE_TIME & r).value = Now
                End If
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " REJECTED_ACK_CONFIRMED broker flat; remains closed"
                GoTo NextRow
            ElseIf netQty = 0 And pending > 0 Then
                wsC.Range(CK_CLOSE_TIME & r).ClearContents
                wsC.Range(CK_EXEC_STATE & r).value = "MISMATCH"
                wsC.Range(CK_COMENTS & r).value = "REJECTED acknowledgement conflicts with pending FYERS order; review required"
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " REJECTED_ACK_CONFLICT pending=" & pending
                GoTo NextRow
            Else
                wsC.Range(CK_CLOSE_TIME & r).ClearContents
                wsC.Range(CK_EXEC_STATE & r).value = "OPEN"
                wsC.Range(CK_FILLED_QTY & r).value = Abs(netQty)
                Dim recoveredAvg As Double: recoveredAvg = SnapshotNumberForSymbol(snapshot, ticker, "avg_price", 0#)
                If recoveredAvg > 0 Then wsC.Range(CK_ENT & r).value = recoveredAvg
                wsC.Range(CK_COMENTS & r).value = "FYERS confirms open position; recovered from conflicting rejection acknowledgement"
                state = "OPEN"
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " RECOVERED_OPEN_FROM_BROKER netQty=" & netQty & " avg=" & recoveredAvg
            End If
        End If

        Dim wasSubmitting As Boolean: wasSubmitting = (state = "SUBMITTING")
        ApplyV6OrderResult wsC, r, state
        state = UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & r).value)))
        If state = "REJECTED_NO_FILL" Then
            If netQty <> 0 Then
                wsC.Range(CK_CLOSE_TIME & r).ClearContents
                wsC.Range(CK_EXEC_STATE & r).value = "OPEN"
                wsC.Range(CK_FILLED_QTY & r).value = Abs(netQty)
                Dim ackBrokerAvg As Double: ackBrokerAvg = SnapshotNumberForSymbol(snapshot, ticker, "avg_price", 0#)
                If ackBrokerAvg > 0 Then wsC.Range(CK_ENT & r).value = ackBrokerAvg
                wsC.Range(CK_COMENTS & r).value = "FYERS confirms open position; rejected acknowledgement did not match broker state"
                state = "OPEN"
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " REJECTION_OVERRIDDEN_BY_BROKER netQty=" & netQty & " avg=" & ackBrokerAvg
            ElseIf pending > 0 Then
                wsC.Range(CK_CLOSE_TIME & r).ClearContents
                wsC.Range(CK_EXEC_STATE & r).value = "MISMATCH"
                wsC.Range(CK_COMENTS & r).value = "Rejected acknowledgement conflicts with pending FYERS order; review required"
                state = "MISMATCH"
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " REJECTION_PENDING_CONFLICT pending=" & pending
            Else
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " REJECTION_CONFIRMED broker flat pending=0"
                GoTo NextRow
            End If
        End If
        If state <> "OPEN" And state <> "SUBMITTING" And state <> "ENTRY_CANCEL_PENDING" And state <> "EXIT_PENDING" And state <> "EXIT_ACK_WAIT" And state <> "MISMATCH" Then
            V6ReconcileLog "row=" & r & " ticker=" & ticker & " SKIP after acknowledgement state=" & state
            GoTo NextRow
        End If

        Dim rowQty As Long: rowQty = CLng(val(wsC.Range(CK_FILLED_QTY & r).value))
        Dim action As String: action = UCase$(Trim$(CStr(wsC.Range(CK_ACTION & r).value)))

        If netQty <> 0 Then
            Dim origin As String: origin = UCase$(Trim$(CStr(wsC.Range(CK_ORIGIN & r).value)))
            Dim rowOrderId As String: rowOrderId = Trim$(CStr(wsC.Range(CK_ORDER_ID & r).value))
            Dim rowRequestId As String: rowRequestId = Trim$(CStr(wsC.Range(CK_REQUEST_ID & r).value))
            Dim newlyAdopted As Boolean: newlyAdopted = False
            If origin = "" Then
                If Len(rowOrderId) = 0 And Len(rowRequestId) = 0 Then
                    origin = "USER": newlyAdopted = True
                Else
                    origin = "V6"
                End If
                wsC.Range(CK_ORIGIN & r).value = origin
            End If
            If (action = "BUY" And netQty < 0) Or (action = "SELL" And netQty > 0) Then
                action = IIf(netQty > 0, "BUY", "SELL")
                wsC.Range(CK_ACTION & r).value = action
                If origin = "V6" Then origin = "MIXED" Else origin = "USER"
                wsC.Range(CK_ORIGIN & r).value = origin
                newlyAdopted = True
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " BROKER_SIDE_SYNCED origin=" & origin & " action=" & action & " netQty=" & netQty
            End If
            If origin = "V6" And state = "OPEN" And rowQty <> Abs(netQty) Then
                origin = "MIXED"
                wsC.Range(CK_ORIGIN & r).value = origin
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " USER_QTY_ADJUSTMENT_MERGED fyersQty=" & Abs(netQty) & " sheetQty=" & rowQty
            End If
            If origin = "USER" Or origin = "MIXED" Then
                Dim userAvg As Double: userAvg = SnapshotNumberForSymbol(snapshot, ticker, "avg_price", 0#)
                If userAvg > 0 Then wsC.Range(CK_ENT & r).value = userAvg
                If newlyAdopted Or Len(Trim$(CStr(wsC.Range(CK_OPEN_TIME & r).value))) = 0 Then
                    wsC.Range(CK_OPEN_TIME & r).value = Now
                End If
            End If
            If state = "MISMATCH" Then
                state = "OPEN"
                wsC.Range(CK_EXEC_STATE & r).value = state
            End If
                If state = "SUBMITTING" Then
                    Dim brokerAvg As Double: brokerAvg = SnapshotNumberForSymbol(snapshot, ticker, "avg_price", 0#)
                    If brokerAvg > 0 Then
                        wsC.Range(CK_ENT & r).value = brokerAvg
                        If action = "BUY" Then
                            wsC.Range(CK_SL & r).value = Round(brokerAvg * (1# - SL_PCT), 2)
                            wsC.Range(CK_TGT & r).value = Round(brokerAvg * (1# + TGT_PCT), 2)
                        Else
                            wsC.Range(CK_SL & r).value = Round(brokerAvg * (1# + SL_PCT), 2)
                            wsC.Range(CK_TGT & r).value = Round(brokerAvg * (1# - TGT_PCT), 2)
                        End If
                    End If
                    wsC.Range(CK_OPEN_TIME & r).value = Now
                    wsC.Range(CK_CANDLE_COUNT & r).value = 0
                    wsC.Range(CK_LAST_BAR_TS & r).value = CDbl(DateValue(Now) + TimeSerial(Hour(Now), (Minute(Now) \ 5) * 5, 0))
                End If
                ' FYERS attached legs are returned as offsets/percentages by the
                ' position snapshot helper; it converts them to absolute prices.
                ' Missing/unknown values use -1 and leave the existing Excel cell alone.
                Dim brokerSLPrice As Double: brokerSLPrice = SnapshotNumberForSymbol(snapshot, ticker, "stop_loss_price", -1#)
                Dim brokerTPPrice As Double: brokerTPPrice = SnapshotNumberForSymbol(snapshot, ticker, "take_profit_price", -1#)
                ' If FYERS omits position-leg fields, fall back to the entry order's
                ' documented orderbook price fields, matched by the saved entry order ID.
                Dim entryOrderIdForLevels As String: entryOrderIdForLevels = Trim$(CStr(wsC.Range(CK_ORDER_ID & r).value))
                If brokerSLPrice < 0 And Len(entryOrderIdForLevels) > 0 Then
                    If SnapshotOrderNumber(snapshot, entryOrderIdForLevels, "has_stop_loss_price") > 0 Then
                        brokerSLPrice = SnapshotOrderNumber(snapshot, entryOrderIdForLevels, "stop_loss_price")
                    End If
                End If
                If brokerTPPrice < 0 And Len(entryOrderIdForLevels) > 0 Then
                    If SnapshotOrderNumber(snapshot, entryOrderIdForLevels, "has_take_profit_price") > 0 Then
                        brokerTPPrice = SnapshotOrderNumber(snapshot, entryOrderIdForLevels, "take_profit_price")
                    End If
                End If
                Dim hasBrokerSL As Boolean: hasBrokerSL = (SnapshotNumberForSymbol(snapshot, ticker, "has_stop_loss_price", 0#) > 0 Or brokerSLPrice > 0)
                Dim hasBrokerTP As Boolean: hasBrokerTP = (SnapshotNumberForSymbol(snapshot, ticker, "has_take_profit_price", 0#) > 0 Or brokerTPPrice > 0)
                wsC.Range(CK_HAS_BROKER_SL & r).value = hasBrokerSL
                wsC.Range(CK_HAS_BROKER_TP & r).value = hasBrokerTP
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " OPEN netQty=" & netQty & _
                    " state=" & state & " SL_API=" & brokerSLPrice & " TP_API=" & brokerTPPrice & _
                    " SL_before=" & wsC.Range(CK_SL & r).value & " TP_before=" & wsC.Range(CK_TGT & r).value
                ' Zero means the broker returned no attached level; never overwrite a valid sheet price with zero.
                If brokerSLPrice > 0 Then wsC.Range(CK_SL & r).value = brokerSLPrice
                If brokerTPPrice > 0 Then wsC.Range(CK_TGT & r).value = brokerTPPrice
                wsC.Range(CK_HAS_BROKER_SL & r).value = _
                    (SnapshotNumberForSymbol(snapshot, ticker, "has_stop_loss_price", 0#) > 0 Or brokerSLPrice > 0)
                wsC.Range(CK_HAS_BROKER_TP & r).value = _
                    (SnapshotNumberForSymbol(snapshot, ticker, "has_take_profit_price", 0#) > 0 Or brokerTPPrice > 0)
                If brokerSLPrice <= 0 Then
                    Dim fallbackEntry As Double: fallbackEntry = SnapshotNumberForSymbol(snapshot, ticker, "avg_price", 0#)
                    If fallbackEntry > 0 Then
                        If action = "BUY" Then
                            wsC.Range(CK_SL & r).value = Round(fallbackEntry * (1# - SL_PCT), 2)
                        Else
                            wsC.Range(CK_SL & r).value = Round(fallbackEntry * (1# + SL_PCT), 2)
                        End If
                    End If
                End If
                If brokerTPPrice <= 0 Then
                    If fallbackEntry <= 0 Then fallbackEntry = SnapshotNumberForSymbol(snapshot, ticker, "avg_price", 0#)
                    If fallbackEntry > 0 Then
                        If action = "BUY" Then
                            wsC.Range(CK_TGT & r).value = Round(fallbackEntry * (1# + TGT_PCT), 2)
                        Else
                            wsC.Range(CK_TGT & r).value = Round(fallbackEntry * (1# - TGT_PCT), 2)
                        End If
                    End If
                End If
                If Len(Trim$(CStr(wsC.Range(CK_OPEN_TIME & r).value))) = 0 Then wsC.Range(CK_OPEN_TIME & r).value = Now
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " LEVELS_WRITTEN SL=" & _
                    wsC.Range(CK_SL & r).value & " TP=" & wsC.Range(CK_TGT & r).value
                If state = "ENTRY_CANCEL_PENDING" And pending = 0 Then
                    wsC.Range(CK_EXEC_STATE & r).value = "OPEN"
                    wsC.Range(CK_REQUEST_ID & r).ClearContents
                    wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                    wsC.Range(CK_COMENTS & r).value = "FYERS entry partially filled; remaining quantity cancelled; open quantity reconciled"
                ElseIf state = "SUBMITTING" And pending > 0 And UCase$(Trim$(CStr(wsC.Range(CK_REQUEST_TYPE & r).value))) = "ENTRY" Then
                    If V6CancelPendingEntry(wsC, r, ticker) Then
                        wsC.Range(CK_EXEC_STATE & r).value = "ENTRY_CANCEL_PENDING"
                        state = "ENTRY_CANCEL_PENDING"
                        wsC.Range(CK_COMENTS & r).value = "FYERS entry partially filled; cancellation requested for unfilled remainder"
                    End If
                ElseIf state = "EXIT_PENDING" And pending = 0 Then
                    Dim exitOrderId As String: exitOrderId = Trim$(CStr(wsC.Range(CK_ORDER_ID & r).value))
                    Dim exitOrderStatus As Long: exitOrderStatus = SnapshotOrderStatus(snapshot, exitOrderId)
                    Dim exitFilled As Long: exitFilled = CLng(SnapshotOrderNumber(snapshot, exitOrderId, "filled_qty"))
                    If (exitOrderStatus = 1 Or exitOrderStatus = 5 Or exitOrderStatus = 7) And exitFilled = 0 Then
                        wsC.Range(CK_EXEC_STATE & r).value = "OPEN"
                        wsC.Range(CK_REQUEST_ID & r).ClearContents
                        wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                        If Left$(Trim$(CStr(wsC.Range(CK_COMENTS & r).value)), 11) = "EXIT_RETRY|" Then
                            wsC.Range(CK_COMENTS & r).value = Trim$(CStr(wsC.Range(CK_COMENTS & r).value)) & " | LTP limit not filled; retrying"
                        End If
                    ElseIf exitOrderStatus = 2 Then
                        wsC.Range(CK_EXEC_STATE & r).value = "EXIT_ACK_WAIT"
                        wsC.Range(CK_REQUEST_ID & r).ClearContents
                        wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                        wsC.Range(CK_COMENTS & r).value = "Exit order filled/part-filled; waiting for broker position snapshot to settle"
                    ElseIf exitFilled > 0 Then
                        ' IOC limit was partially filled and the remainder cancelled.
                        ' Retry the still-open broker quantity at the next observed LTP.
                        wsC.Range(CK_EXEC_STATE & r).value = "OPEN"
                        wsC.Range(CK_REQUEST_ID & r).ClearContents
                        wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                        If Left$(Trim$(CStr(wsC.Range(CK_COMENTS & r).value)), 11) = "EXIT_RETRY|" Then
                            wsC.Range(CK_COMENTS & r).value = Trim$(CStr(wsC.Range(CK_COMENTS & r).value)) & " | partial fill; retrying remainder"
                        End If
                    End If
                ElseIf state <> "EXIT_ACK_WAIT" And state <> "ENTRY_CANCEL_PENDING" And Not (state = "EXIT_PENDING" And pending > 0) Then
                    wsC.Range(CK_EXEC_STATE & r).value = "OPEN"
                End If
                wsC.Range(CK_FILLED_QTY & r).value = Abs(netQty)
                If rowQty <> Abs(netQty) Then
                    If Left$(Trim$(CStr(wsC.Range(CK_COMENTS & r).value)), 11) = "EXIT_RETRY|" Then
                        wsC.Range(CK_COMENTS & r).value = Trim$(CStr(wsC.Range(CK_COMENTS & r).value)) & _
                            " | remaining FYERS qty=" & Abs(netQty)
                    Else
                        wsC.Range(CK_COMENTS & r).value = "Reconciled filled quantity to FYERS net position"
                    End If
                End If
        ElseIf pending = 0 Then
            If state = "SUBMITTING" Then
                Dim brokerOrderId As String: brokerOrderId = Trim$(CStr(wsC.Range(CK_ORDER_ID & r).value))
                Dim orderStatus As Long: orderStatus = SnapshotOrderStatus(snapshot, brokerOrderId)
                Dim brokerFilled As Long: brokerFilled = CLng(SnapshotOrderNumber(snapshot, brokerOrderId, "filled_qty"))
                If (orderStatus = 1 Or orderStatus = 5 Or orderStatus = 7) And brokerFilled = 0 Then
                    wsC.Range(CK_EXEC_STATE & r).value = "CANCELLED_NO_FILL"
                    wsC.Range(CK_CLOSE_TIME & r).value = Now
                    wsC.Range(CK_REQUEST_ID & r).ClearContents
                    wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                    wsC.Range(CK_COMENTS & r).value = "FYERS entry order confirms no fill; terminal status=" & orderStatus
                    V6ReconcileLog "row=" & r & " ticker=" & ticker & " ENTRY_REJECTED orderStatus=" & orderStatus
                ElseIf orderStatus = 2 Or brokerFilled > 0 Then
                    wsC.Range(CK_EXEC_STATE & r).value = "MISMATCH"
                    wsC.Range(CK_COMENTS & r).value = "FYERS orderbook reports fill but net position is zero; manual reconciliation required"
                    V6ReconcileLog "row=" & r & " ticker=" & ticker & " FLAT_BUT_ORDERBOOK_FILLED status=" & orderStatus & " filled=" & brokerFilled
                Else
                    V6ReconcileLog "row=" & r & " ticker=" & ticker & " SUBMITTING flat; awaiting orderbook result status=" & orderStatus & " filled=" & brokerFilled
                End If
            ElseIf state = "ENTRY_CANCEL_PENDING" Then
                Dim cancelOrderId As String: cancelOrderId = Trim$(CStr(wsC.Range(CK_ORDER_ID & r).value))
                Dim cancelOrderStatus As Long: cancelOrderStatus = SnapshotOrderStatus(snapshot, cancelOrderId)
                Dim cancelFilled As Long: cancelFilled = CLng(SnapshotOrderNumber(snapshot, cancelOrderId, "filled_qty"))
                If (cancelOrderStatus = 1 Or cancelOrderStatus = 5 Or cancelOrderStatus = 7) And cancelFilled = 0 Then
                    wsC.Range(CK_EXEC_STATE & r).value = "CANCELLED_NO_FILL"
                    wsC.Range(CK_CLOSE_TIME & r).value = Now
                    wsC.Range(CK_REQUEST_ID & r).ClearContents
                    wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                    wsC.Range(CK_COMENTS & r).value = "FYERS entry order cancelled after 3 completed 5m bars with no fill"
                ElseIf cancelOrderStatus = 2 Or cancelFilled > 0 Then
                    wsC.Range(CK_EXEC_STATE & r).value = "MISMATCH"
                    wsC.Range(CK_COMENTS & r).value = "FYERS reports entry fill but position is flat; reconciliation required"
                Else
                    V6ReconcileLog "row=" & r & " ticker=" & ticker & " ENTRY_CANCEL_PENDING flat; awaiting terminal order status=" & cancelOrderStatus
                End If
            ElseIf state = "MISMATCH" Then
                ' A successful positions+orderbook snapshot confirms flat/no pending;
                ' retain the row as history but release its live slot.
                Dim mismatchSrcRow As Long: mismatchSrcRow = FindSheet1Row(wsSrc, ticker)
                Dim mismatchClosePx As Double: mismatchClosePx = 0#
                If mismatchSrcRow >= 4 Then mismatchClosePx = CDbl(val(wsSrc.cells(mismatchSrcRow, "T").value))
                If mismatchClosePx > 0 Then wsC.Range(CK_CP & r).value = mismatchClosePx
                wsC.Range(CK_CLOSE_TIME & r).value = Now
                wsC.Range(CK_EXEC_STATE & r).value = "CLOSED"
                wsC.Range(CK_REQUEST_ID & r).ClearContents
                wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                wsC.Range(CK_COMENTS & r).value = "BROKER_FLAT | Reconciled: FYERS confirms no open position or pending order"
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " MISMATCH_RESOLVED broker flat pending=0 CP=" & mismatchClosePx
            ElseIf wasSubmitting And state = "OPEN" Then
                wsC.Range(CK_EXEC_STATE & r).value = "MISMATCH"
                wsC.Range(CK_COMENTS & r).value = "FYERS acknowledged a fill but position snapshot is flat; manual reconciliation required"
            Else
                Dim srcRow As Long: srcRow = FindSheet1Row(wsSrc, ticker)
                Dim closePx As Double: closePx = CDbl(val(wsC.Range(CK_EXIT_FILL & r).value))
                If closePx <= 0 And srcRow >= 4 Then closePx = CDbl(val(wsSrc.cells(srcRow, "T").value))
                If closePx > 0 Then wsC.Range(CK_CP & r).value = closePx
                wsC.Range(CK_CLOSE_TIME & r).value = Now
                wsC.Range(CK_EXEC_STATE & r).value = "CLOSED"
                V6ReconcileLog "row=" & r & " ticker=" & ticker & " CLOSED broker flat pending=0 CP=" & closePx & " prior_state=" & state
                If state = "EXIT_PENDING" Then
                    wsC.Range(CK_COMENTS & r).value = Trim$(CStr(wsC.Range(CK_COMENTS & r).value)) & " | broker flat confirmed"
                Else
                    wsC.Range(CK_COMENTS & r).value = "BROKER_EXIT: FYERS net position is zero and no pending order remains"
                End If
            End If
        End If
        If netQty = 0 And pending > 0 Then
            V6ReconcileLog "row=" & r & " ticker=" & ticker & " NOT_CLOSED broker flat but pending=" & pending & " state=" & state
            If state = "ENTRY_CANCEL_PENDING" And UCase$(Trim$(CStr(wsC.Range(CK_REQUEST_TYPE & r).value))) = "ENTRY" Then
                ' A cancel request is not a cancellation confirmation. Retry while FYERS still shows the order pending.
                V6CancelPendingEntry wsC, r, ticker
            End If
            If state = "SUBMITTING" And UCase$(Trim$(CStr(wsC.Range(CK_REQUEST_TYPE & r).value))) = "ENTRY" Then
                Dim entryBarStamp As Date
                entryBarStamp = DateValue(Now) + TimeSerial(Hour(Now), (Minute(Now) \ 5) * 5, 0)
                Dim lastEntryStamp As Double: lastEntryStamp = CDbl(val(wsC.Range(CK_LAST_BAR_TS & r).value))
                If CDbl(entryBarStamp) > lastEntryStamp Then
                    Dim entryBarsWaited As Long: entryBarsWaited = CLng(val(wsC.Range(CK_CANDLE_COUNT & r).value)) + 1
                    wsC.Range(CK_CANDLE_COUNT & r).value = entryBarsWaited
                    wsC.Range(CK_LAST_BAR_TS & r).value = CDbl(entryBarStamp)
                    If entryBarsWaited >= ENTRY_LIMIT_WAIT_BARS Then
                        If V6CancelPendingEntry(wsC, r, ticker) Then
                            wsC.Range(CK_EXEC_STATE & r).value = "ENTRY_CANCEL_PENDING"
                            state = "ENTRY_CANCEL_PENDING"
                            wsC.Range(CK_COMENTS & r).value = "Selected 5m EMA limit not filled within 3 completed 5m bars; cancellation requested"
                            ' Confirm the broker cancellation in this same engine run where possible.
                            ' A cancel request alone is not enough to release the row or slot.
                            Dim afterCancelSnapshot As String
                            If GetV6BrokerSnapshot(afterCancelSnapshot) Then
                                gV6Snapshot = afterCancelSnapshot
                                gV6SnapshotReady = True
                                Dim afterCancelNetQty As Long
                                Dim afterCancelPending As Long
                                Dim afterCancelStatus As Long
                                Dim afterCancelFilled As Long
                                afterCancelNetQty = CLng(SnapshotNumberForSymbol(afterCancelSnapshot, ticker, "net_qty", 0#))
                                afterCancelPending = CLng(SnapshotNumberForSymbol(afterCancelSnapshot, ticker, "pending", 0#))
                                afterCancelStatus = SnapshotOrderStatus(afterCancelSnapshot, Trim$(CStr(wsC.Range(CK_ORDER_ID & r).value)))
                                afterCancelFilled = CLng(SnapshotOrderNumber(afterCancelSnapshot, _
                                    Trim$(CStr(wsC.Range(CK_ORDER_ID & r).value)), "filled_qty"))
                                V6ReconcileLog "row=" & r & " ticker=" & ticker & " POST_CANCEL_SNAPSHOT netQty=" & _
                                    afterCancelNetQty & " pending=" & afterCancelPending & " orderStatus=" & _
                                    afterCancelStatus & " filled=" & afterCancelFilled
                                If afterCancelNetQty = 0 And afterCancelPending = 0 And _
                                        (afterCancelStatus = 1 Or afterCancelStatus = 5 Or afterCancelStatus = 7) And _
                                        afterCancelFilled = 0 Then
                                    wsC.Range(CK_EXEC_STATE & r).value = "CANCELLED_NO_FILL"
                                    wsC.Range(CK_CLOSE_TIME & r).value = Now
                                    wsC.Range(CK_REQUEST_ID & r).ClearContents
                                    wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                                    wsC.Range(CK_COMENTS & r).value = _
                                        "FYERS confirmed entry cancelled after 3 completed 5m bars; no fill"
                                    state = "CANCELLED_NO_FILL"
                                    V6ReconcileLog "row=" & r & " ticker=" & ticker & " CANCELLED_NO_FILL confirmed same run"
                                End If
                            Else
                                V6ReconcileLog "row=" & r & " ticker=" & ticker & _
                                    " POST_CANCEL_SNAPSHOT unavailable; keep ENTRY_CANCEL_PENDING until next run"
                            End If
                        Else
                            wsC.Range(CK_COMENTS & r).value = "Selected 5m EMA limit age >=3 bars; cancellation failed/unconfirmed; slot remains reserved"
                        End If
                    End If
                End If
            End If
        End If
NextRow:
        On Error GoTo 0
    Next r
    AdoptUnmatchedBrokerPositions wsC, wsSrc, snapshot
    MarkDuplicateAppliedTickerRows wsC
    gV6ReconciliationReady = V6SnapshotFullyReconciled(wsC, snapshot)
    V6ReconcileLog "RECONCILIATION_RESULT ready=" & gV6ReconciliationReady & _
        " api_positions=" & CountSnapshotOpenPositions(snapshot) & _
        " pending_symbols=" & CountSnapshotPendingSymbols(snapshot) & _
        " broker_rows=" & CountOpenBrokerRows(wsC)
    V6ReconcileLog "END"
    Exit Sub

ReconcileError:
    Dim reconcileErrNumber As Long: reconcileErrNumber = Err.Number
    Dim reconcileErrDescription As String: reconcileErrDescription = Err.Description
    gV6ReconciliationReady = False
    On Error Resume Next
    If r >= DATA_START_ROW And r <= lastRow Then
        If Len(ticker) > 0 Then
            wsC.Range(CK_EXEC_STATE & r).value = "MISMATCH"
            wsC.Range(CK_COMENTS & r).value = "RECONCILE_ERROR | row held open; needs review"
        End If
        V6ReconcileLog "ROW_ERROR number=" & reconcileErrNumber & " description=" & reconcileErrDescription & _
            " row=" & r & " ticker=" & ticker
        Resume NextRow
    End If
    V6ReconcileLog "RECONCILE_ERROR number=" & reconcileErrNumber & " description=" & reconcileErrDescription
    On Error GoTo 0
    Err.Raise reconcileErrNumber, "ReconcileBrokerPositions", reconcileErrDescription
End Sub


Private Function V6SnapshotFullyReconciled(wsC As Worksheet, ByVal snapshot As String) As Boolean
    V6SnapshotFullyReconciled = Not HasV6ExecutionMismatch(wsC) _
        And Not HasUnmatchedBrokerPosition(wsC, snapshot)
End Function


Private Function HasUnmatchedBrokerPosition(wsC As Worksheet, ByVal snapshot As String) As Boolean
    Dim p As Long, marker As Long, qEnd As Long, qStart As Long, ticker As String
    Dim netQty As Long
    p = 1
    Do
        marker = InStr(p, snapshot, ":{""net_qty"":", vbTextCompare)
        If marker = 0 Then Exit Do
        qEnd = InStrRev(snapshot, """", marker - 1)
        qStart = InStrRev(snapshot, """", qEnd - 1)
        If qStart > 0 And qEnd > qStart Then
            ticker = UCase$(Mid$(snapshot, qStart + 1, qEnd - qStart - 1))
            If Left$(ticker, 2) <> "__" Then
                netQty = CLng(SnapshotNumberForSymbol(snapshot, ticker, "net_qty", 0#))
                If netQty <> 0 And FindActiveAppliedTickerRow(wsC, ticker) = 0 Then
                    HasUnmatchedBrokerPosition = True
                    V6ReconcileLog "UNMATCHED_BROKER_POSITION ticker=" & ticker & " netQty=" & netQty
                    Exit Function
                End If
            End If
        End If
        p = marker + 1
    Loop
End Function


Private Function CountUnreflectedEntryRequests(wsC As Worksheet, ByVal snapshot As String, _
                                               ByRef requestDetails As String) As Long
    Dim lastRow As Long: lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    Dim r As Long, k As Long, ticker As String, state As String, reqType As String
    Dim closeVal As Variant, priorCloseVal As Variant
    For r = DATA_START_ROW To lastRow
        If Not V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) Then GoTo NextRequestRow
        ' A closed history row must never reserve a live slot, even if its old
        ' execution state/request type was not cleared when it was closed.
        closeVal = wsC.Range(CK_CLOSE_TIME & r).value
        If IsError(closeVal) Then GoTo NextRequestRow
        If Not IsEmpty(closeVal) Then
            If Len(Trim$(CStr(closeVal))) > 0 Then GoTo NextRequestRow
        End If
        state = UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & r).value)))
        reqType = UCase$(Trim$(CStr(wsC.Range(CK_REQUEST_TYPE & r).value)))
        If (state <> "SUBMITTING" And state <> "ENTRY_CANCEL_PENDING") Or reqType <> "ENTRY" Then GoTo NextRequestRow
        ticker = UCase$(Trim$(CStr(wsC.Range(CK_TICKER & r).value)))
        If Len(ticker) = 0 Or IsTickerOpenInBrokerSnapshot(snapshot, ticker) Then GoTo NextRequestRow
        ' The broker pending-symbol count already reserves this order's slot.
        ' Do not add the same pending entry again as an unreflected request.
        If SnapshotNumberForSymbol(snapshot, ticker, "pending", 0#) > 0 Then GoTo NextRequestRow
        Dim brokerOrderId As String: brokerOrderId = Trim$(CStr(wsC.Range(CK_ORDER_ID & r).value))
        If Len(brokerOrderId) > 0 Then
            Dim brokerOrderStatus As Long: brokerOrderStatus = SnapshotOrderStatus(snapshot, brokerOrderId)
            Dim brokerFilledQty As Long: brokerFilledQty = CLng(SnapshotOrderNumber(snapshot, brokerOrderId, "filled_qty"))
            If (brokerOrderStatus = 1 Or brokerOrderStatus = 5 Or brokerOrderStatus = 7) And brokerFilledQty = 0 Then
                ' Reconciliation should normally have closed this row already. If
                ' not, don't reserve a slot for a terminal no-fill broker order.
                wsC.Range(CK_EXEC_STATE & r).value = "CANCELLED_NO_FILL"
                wsC.Range(CK_CLOSE_TIME & r).value = Now
                wsC.Range(CK_REQUEST_ID & r).ClearContents
                wsC.Range(CK_REQUEST_TYPE & r).ClearContents
                wsC.Range(CK_COMENTS & r).value = "FYERS orderbook confirms terminal no-fill status=" & brokerOrderStatus
                V6ReconcileLog "row=" & r & " ticker=" & ticker & _
                    " SLOT_RELEASED_BY_ORDER_ID status=" & brokerOrderStatus & " filled=0"
                GoTo NextRequestRow
            ElseIf brokerOrderStatus = 2 Or brokerFilledQty > 0 Then
                ' A fill with no matching net position is a real discrepancy; keep
                ' entries blocked for review rather than silently releasing capacity.
                wsC.Range(CK_EXEC_STATE & r).value = "MISMATCH"
                wsC.Range(CK_COMENTS & r).value = "FYERS orderbook reports a fill but no net position is visible; reconciliation required"
                V6ReconcileLog "row=" & r & " ticker=" & ticker & _
                    " SLOT_RECONCILIATION_MISMATCH order_status=" & brokerOrderStatus & _
                    " filled=" & brokerFilledQty
                GoTo NextRequestRow
            End If
        End If
        ' One slot per symbol even if duplicate rows exist; duplicates are separately
        ' marked MISMATCH and block live entries.
        For k = DATA_START_ROW To r - 1
            If V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & k).value) _
                    And UCase$(Trim$(CStr(wsC.Range(CK_TICKER & k).value))) = ticker Then
                priorCloseVal = wsC.Range(CK_CLOSE_TIME & k).value
                If Not IsError(priorCloseVal) Then
                    If IsEmpty(priorCloseVal) Or Len(Trim$(CStr(priorCloseVal))) = 0 Then
                        If (UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & k).value))) = "SUBMITTING" _
                                Or UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & k).value))) = "ENTRY_CANCEL_PENDING") _
                                And UCase$(Trim$(CStr(wsC.Range(CK_REQUEST_TYPE & k).value))) = "ENTRY" Then GoTo NextRequestRow
                    End If
                End If
            End If
        Next k
        CountUnreflectedEntryRequests = CountUnreflectedEntryRequests + 1
        If Len(requestDetails) > 0 Then requestDetails = requestDetails & ";"
        requestDetails = requestDetails & "row=" & r & ":ticker=" & ticker & ":state=" & state & _
            ":order=" & CStr(wsC.Range(CK_ORDER_ID & r).value) & _
            ":request=" & CStr(wsC.Range(CK_REQUEST_ID & r).value)
NextRequestRow:
    Next r
End Function


Private Sub AdoptUnmatchedBrokerPositions(wsC As Worksheet, wsSrc As Worksheet, ByVal snapshot As String)
    Dim p As Long, marker As Long, qEnd As Long, qStart As Long, ticker As String
    Dim netQty As Long, activeRow As Long
    p = 1
    Do
        marker = InStr(p, snapshot, ":{""net_qty"":", vbTextCompare)
        If marker = 0 Then Exit Do
        qEnd = InStrRev(snapshot, """", marker - 1)
        qStart = InStrRev(snapshot, """", qEnd - 1)
        If qStart > 0 And qEnd > qStart Then
            ticker = UCase$(Mid$(snapshot, qStart + 1, qEnd - qStart - 1))
            If Left$(ticker, 2) <> "__" Then
                netQty = CLng(SnapshotNumberForSymbol(snapshot, ticker, "net_qty", 0#))
                If netQty <> 0 Then
                    activeRow = FindActiveAppliedTickerRow(wsC, ticker)
                    If activeRow = 0 Then CreateUserBrokerRow wsC, wsSrc, snapshot, ticker, netQty
                End If
            End If
        End If
        p = marker + 1
    Loop
End Sub


Private Function FindActiveAppliedTickerRow(wsC As Worksheet, ByVal ticker As String) As Long
    Dim lastRow As Long: lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If UCase$(Trim$(CStr(wsC.Range(CK_TICKER & r).value))) = UCase$(ticker) _
                And V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) Then
            Dim closeVal As Variant: closeVal = wsC.Range(CK_CLOSE_TIME & r).value
            If IsEmpty(closeVal) Or Len(Trim$(CStr(closeVal))) = 0 Then
                Dim state As String: state = UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & r).value)))
                If state = "" Or state = "OPEN" Or state = "SUBMITTING" Or state = "ENTRY_CANCEL_PENDING" _
                        Or state = "EXIT_PENDING" Or state = "EXIT_ACK_WAIT" Or state = "MISMATCH" Then
                    FindActiveAppliedTickerRow = r
                    Exit Function
                End If
            End If
        End If
    Next r
End Function


Private Sub CreateUserBrokerRow(wsC As Worksheet, wsSrc As Worksheet, ByVal snapshot As String, _
                                ByVal ticker As String, ByVal netQty As Long)
    Dim destRow As Long: destRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row + 1
    If destRow < DATA_START_ROW Then destRow = DATA_START_ROW
    Dim srcRow As Long: srcRow = FindSheet1Row(wsSrc, ticker)
    Dim action As String: action = IIf(netQty > 0, "BUY", "SELL")
    Dim avgPx As Double: avgPx = SnapshotNumberForSymbol(snapshot, ticker, "avg_price", 0#)
    Dim liveLtp As Double: liveLtp = SnapshotNumberForSymbol(snapshot, ticker, "ltp", avgPx)
    Dim ltpCell As Variant
    If srcRow >= 4 Then
        ltpCell = wsSrc.cells(srcRow, "T").value
        If Not IsError(ltpCell) Then
            If IsNumeric(ltpCell) Then
                If CDbl(ltpCell) > 0 Then liveLtp = CDbl(ltpCell)
            End If
        End If
    End If
    Dim brokerSL As Double: brokerSL = SnapshotNumberForSymbol(snapshot, ticker, "stop_loss_price", -1#)
    Dim brokerTP As Double: brokerTP = SnapshotNumberForSymbol(snapshot, ticker, "take_profit_price", -1#)
    Dim hasSL As Boolean: hasSL = (SnapshotNumberForSymbol(snapshot, ticker, "has_stop_loss_price", 0#) > 0 And brokerSL > 0)
    Dim hasTP As Boolean: hasTP = (SnapshotNumberForSymbol(snapshot, ticker, "has_take_profit_price", 0#) > 0 And brokerTP > 0)
    Dim slPx As Double, tgtPx As Double
    If hasSL Then
        slPx = brokerSL
    ElseIf action = "BUY" Then
        slPx = Round(avgPx * (1# - SL_PCT), 2)
    Else
        slPx = Round(avgPx * (1# + SL_PCT), 2)
    End If
    If hasTP Then
        tgtPx = brokerTP
    ElseIf action = "BUY" Then
        tgtPx = Round(avgPx * (1# + TGT_PCT), 2)
    Else
        tgtPx = Round(avgPx * (1# - TGT_PCT), 2)
    End If

    If srcRow >= 4 Then
        WriteNewV6Row wsC, wsSrc, srcRow, destRow, ticker, action, liveLtp, avgPx, tgtPx, slPx, _
            Abs(netQty), SafeV6CellNumber(wsSrc, srcRow, "U", 0#), _
            SafeV6CellNumber(wsSrc, srcRow, ADX_SOURCE_COL, 0#), SafeV6CellNumber(wsSrc, srcRow, RVOL_SOURCE_COL, 0#)
    Else
        wsC.Range(CK_TICKER & destRow).value = ticker
        wsC.Range(CK_ACTION & destRow).value = action
        wsC.Range(CK_LTP & destRow).value = liveLtp
        wsC.Range(CK_ENT & destRow).value = avgPx
        wsC.Range(CK_TGT & destRow).value = tgtPx
        wsC.Range(CK_SL & destRow).value = slPx
        wsC.Range(CK_CP & destRow).value = liveLtp
    End If
    wsC.Range(CK_ENT & destRow).value = avgPx
    wsC.Range(CK_TGT & destRow).value = tgtPx
    wsC.Range(CK_SL & destRow).value = slPx
    wsC.Range(CK_CP & destRow).value = liveLtp
    wsC.Range(CK_OPEN_TIME & destRow).value = Now
    wsC.Range(CK_CLOSE_TIME & destRow).ClearContents
    wsC.Range(CK_APPLY_TO_FYERS & destRow).value = True
    wsC.Range(CK_ORIGIN & destRow).value = "USER"
    wsC.Range(CK_HAS_BROKER_SL & destRow).value = hasSL
    wsC.Range(CK_HAS_BROKER_TP & destRow).value = hasTP
    wsC.Range(CK_EXEC_STATE & destRow).value = "OPEN"
    wsC.Range(CK_FILLED_QTY & destRow).value = Abs(netQty)
    wsC.Range(CK_REQ_QTY & destRow).value = Abs(netQty)
    wsC.Range(CK_ORDER_ID & destRow).ClearContents
    wsC.Range(CK_REQUEST_ID & destRow).ClearContents
    wsC.Range(CK_REQUEST_TYPE & destRow).ClearContents
    wsC.Range(CK_CANDLE_COUNT & destRow).value = 0
    wsC.Range(CK_PEAK_PRICE & destRow).value = liveLtp
    wsC.Range(CK_BEST_MFE_PCT & destRow).value = 0#
    wsC.Range(CK_LAST_BAR_TS & destRow).value = CDbl(DateValue(Now) + TimeSerial(Hour(Now), (Minute(Now) \ 5) * 5, 0))
    wsC.Range(CK_COMENTS & destRow).value = "USER_OPEN | Adopted from FYERS snapshot; qty/side/entry reconciled. " & _
        IIf(hasSL, "FYERS SL active", "SL estimated locally") & " | " & IIf(hasTP, "FYERS TP active", "TP estimated locally")
    If srcRow < 4 Then
        wsC.Range(CK_COMENTS & destRow).value = CStr(wsC.Range(CK_COMENTS & destRow).value) & _
            " | Symbol absent from Sheet1; EMA/ADX exit checks unavailable"
    End If
    V6ReconcileLog "row=" & destRow & " ticker=" & ticker & " USER_POSITION_ADOPTED side=" & action & _
        " qty=" & Abs(netQty) & " avg=" & avgPx & " ltp=" & liveLtp & " hasSL=" & hasSL & " hasTP=" & hasTP
End Sub


Private Sub MarkDuplicateAppliedTickerRows(wsC As Worksheet)
    Dim lastRow As Long: lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    Dim r As Long, k As Long
    For r = DATA_START_ROW To lastRow
        Dim ticker As String: ticker = UCase$(Trim$(CStr(wsC.Range(CK_TICKER & r).value)))
        Dim closeVal As Variant: closeVal = wsC.Range(CK_CLOSE_TIME & r).value
        If Len(ticker) = 0 Or Not V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & r).value) _
                Or (Not IsEmpty(closeVal) And Len(Trim$(CStr(closeVal))) > 0) Then GoTo NextOwner
        For k = r + 1 To lastRow
            Dim laterClose As Variant: laterClose = wsC.Range(CK_CLOSE_TIME & k).value
            If UCase$(Trim$(CStr(wsC.Range(CK_TICKER & k).value))) = ticker _
                    And V6ApplyToFyers(wsC.Range(CK_APPLY_TO_FYERS & k).value) _
                    And (IsEmpty(laterClose) Or Len(Trim$(CStr(laterClose))) = 0) Then
                wsC.Range(CK_EXEC_STATE & k).value = "MISMATCH"
                wsC.Range(CK_COMENTS & k).value = "DUPLICATE_BROKER_ROW | FYERS nets this ticker; row " & r & " is the active record"
                V6ReconcileLog "row=" & k & " ticker=" & ticker & " DUPLICATE_BROKER_ROW owner_row=" & r
            End If
        Next k
NextOwner:
    Next r
End Sub


Private Function SafeV6CellNumber(ws As Worksheet, ByVal rowNum As Long, ByVal colRef As String, _
                                  Optional ByVal fallback As Double = 0#) As Double
    On Error GoTo SafeExit
    Dim value As Variant: value = ws.cells(rowNum, colRef).value
    If Not IsError(value) Then
        If IsNumeric(value) Then SafeV6CellNumber = CDbl(value): Exit Function
    End If
SafeExit:
    SafeV6CellNumber = fallback
End Function


Private Sub V6ReconcileLog(ByVal message As String)
    On Error Resume Next
    Dim fso As Object, ts As Object, logDir As String, logPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    logDir = ThisWorkbook.path & "\logs"
    If Not fso.FolderExists(logDir) Then fso.CreateFolder logDir
    logPath = logDir & "\cockpit_v6_reconcile.log"
    Set ts = fso.OpenTextFile(logPath, 8, True)
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    ts.Close
End Sub


Private Sub ApplyV6OrderResult(wsC As Worksheet, ByVal rowNum As Long, ByRef state As String)
    Dim requestId As String: requestId = Trim$(CStr(wsC.Range(CK_REQUEST_ID & rowNum).value))
    If Len(requestId) = 0 Then Exit Sub
    Dim path As String: path = ThisWorkbook.path & "\state\v6_order_results\" & requestId & ".json"
    Dim fso As Object, ts As Object, payload As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Or Not fso.FileExists(path) Then Exit Sub
    Set ts = fso.OpenTextFile(path, 1)
    payload = ts.ReadAll: ts.Close
    On Error GoTo 0
    On Error Resume Next
    fso.DeleteFile path, True
    On Error GoTo 0

    Dim status As String: status = UCase$(JsonStringValue(payload, "status"))
    Dim orderId As String: orderId = JsonStringValue(payload, "order_id")
    Dim fillQty As Long: fillQty = CLng(JsonPropertyNumber(payload, "filled_qty"))
    Dim fillPx As Double: fillPx = JsonPropertyNumber(payload, "fill_price")
    wsC.Range(CK_ORDER_ID & rowNum).value = orderId
    Dim requestType As String: requestType = UCase$(Trim$(CStr(wsC.Range(CK_REQUEST_TYPE & rowNum).value)))
    If requestType = "ENTRY" Or state = "SUBMITTING" Then
        If fillQty > 0 And (status = "FILLED" Or status = "PARTIAL") Then
            wsC.Range(CK_EXEC_STATE & rowNum).value = "OPEN"
            wsC.Range(CK_OPEN_TIME & rowNum).value = Now
            wsC.Range(CK_FILLED_QTY & rowNum).value = fillQty
            wsC.Range(CK_CANDLE_COUNT & rowNum).value = 0
            wsC.Range(CK_LAST_BAR_TS & rowNum).value = CDbl(DateValue(Now) + TimeSerial(Hour(Now), (Minute(Now) \ 5) * 5, 0))
            wsC.Range(CK_REQUEST_ID & rowNum).ClearContents
            wsC.Range(CK_REQUEST_TYPE & rowNum).ClearContents
            If fillPx > 0 Then
                wsC.Range(CK_ENT & rowNum).value = fillPx
                If UCase$(CStr(wsC.Range(CK_ACTION & rowNum).value)) = "BUY" Then
                    wsC.Range(CK_SL & rowNum).value = Round(fillPx * (1# - SL_PCT), 2)
                    wsC.Range(CK_TGT & rowNum).value = Round(fillPx * (1# + TGT_PCT), 2)
                Else
                    wsC.Range(CK_SL & rowNum).value = Round(fillPx * (1# + SL_PCT), 2)
                    wsC.Range(CK_TGT & rowNum).value = Round(fillPx * (1# - TGT_PCT), 2)
                End If
            End If
            wsC.Range(CK_COMENTS & rowNum).value = "FYERS " & status & " | order=" & orderId & " | filled=" & fillQty
        ElseIf status = "REJECTED" Or status = "IP_REJECTED" Or status = "REJECTED_NO_FILL" Or status = "BLOCKED_IP" Or status = "ERROR" Then
            wsC.Range(CK_EXEC_STATE & rowNum).value = "REJECTED_NO_FILL"
            wsC.Range(CK_CLOSE_TIME & rowNum).value = Now
            wsC.Range(CK_REQUEST_ID & rowNum).ClearContents
            wsC.Range(CK_REQUEST_TYPE & rowNum).ClearContents
            wsC.Range(CK_COMENTS & rowNum).value = status & ": " & JsonStringValue(payload, "message")
        End If
    ElseIf requestType = "EXIT" Or state = "EXIT_PENDING" Then
        If fillPx > 0 Then wsC.Range(CK_EXIT_FILL & rowNum).value = fillPx
        If status = "REJECTED" Or status = "REJECTED_NO_FILL" Or status = "IP_REJECTED" Or status = "ERROR" Or status = "BLOCKED_IP" Then
            wsC.Range(CK_EXEC_STATE & rowNum).value = "OPEN"
            wsC.Range(CK_REQUEST_ID & rowNum).ClearContents
            wsC.Range(CK_REQUEST_TYPE & rowNum).ClearContents
            If Left$(Trim$(CStr(wsC.Range(CK_COMENTS & rowNum).value)), 11) = "EXIT_RETRY|" Then
                wsC.Range(CK_COMENTS & rowNum).value = Trim$(CStr(wsC.Range(CK_COMENTS & rowNum).value)) & _
                    " | LTP limit not filled: " & JsonStringValue(payload, "message")
            Else
                wsC.Range(CK_COMENTS & rowNum).value = "EXIT FAILED: " & JsonStringValue(payload, "message")
            End If
        ElseIf status = "FILLED" Or status = "PARTIAL" Then
            wsC.Range(CK_REQUEST_ID & rowNum).ClearContents
            wsC.Range(CK_REQUEST_TYPE & rowNum).ClearContents
        End If
    End If
    state = UCase$(Trim$(CStr(wsC.Range(CK_EXEC_STATE & rowNum).value)))
End Sub


Private Function JsonPropertyNumber(ByVal payload As String, ByVal propertyName As String) As Double
    Dim p As Long
    p = InStr(1, payload, """" & propertyName & """:", vbTextCompare)
    If p > 0 Then JsonPropertyNumber = JsonNumberAfter(payload, p + Len("""" & propertyName & """:"))
End Function

Private Function IsTickerAlreadyOpen(wsC As Worksheet, ticker As String) As Boolean
    IsTickerAlreadyOpen = False
    Dim lastRow As Long
    lastRow = wsC.cells(wsC.Rows.count, RangeRef(CK_TICKER)).End(xlUp).row
    If lastRow < DATA_START_ROW Then Exit Function

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If StrComp(Trim$(CStr(wsC.Range(CK_TICKER & r).value)), ticker, vbTextCompare) = 0 Then
            Dim closeTime As Variant: closeTime = wsC.Range(CK_CLOSE_TIME & r).value
            If IsEmpty(closeTime) Or Len(Trim$(CStr(closeTime))) = 0 Then
                IsTickerAlreadyOpen = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Function IsTickerOpenInBrokerSnapshot(ByVal snapshot As String, ByVal ticker As String) As Boolean
    ' Manual FYERS positions and pending orders also reserve this symbol.
    If Len(snapshot) = 0 Or Len(ticker) = 0 Then Exit Function
    If SnapshotNumberForSymbol(snapshot, ticker, "net_qty", 0#) <> 0 Then
        IsTickerOpenInBrokerSnapshot = True
    ElseIf SnapshotNumberForSymbol(snapshot, ticker, "pending", 0#) > 0 Then
        IsTickerOpenInBrokerSnapshot = True
    End If
End Function

Private Function FindSheet1Row(wsSrc As Worksheet, ticker As String) As Long
    Dim lastRow As Long: lastRow = wsSrc.cells(wsSrc.Rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 4 To lastRow
        If StrComp(Trim$(CStr(wsSrc.cells(r, 1).value)), ticker, vbTextCompare) = 0 Then
            FindSheet1Row = r
            Exit Function
        End If
    Next r
    FindSheet1Row = 0
End Function

Private Function IsIndexSymbol(ByVal s As String) As Boolean
    s = UCase$(Trim$(s))
    If Left$(s, 3) = "NSE" Or Left$(s, 3) = "CNX" Or s = "NIFTY" Or s = "BANKNIFTY" Or Left$(s, 1) = "^" Then
        IsIndexSymbol = True
    Else
        IsIndexSymbol = False
    End If
End Function

Private Function RangeRef(ByVal colLetter As String) As Long
    RangeRef = wsColToNum(colLetter)
End Function

Private Function wsColToNum(ByVal colLetter As String) As Long
    Dim res As Long: res = 0
    Dim i As Long
    colLetter = UCase$(Trim$(colLetter))
    For i = 1 To Len(colLetter)
        res = res * 26 + (Asc(Mid$(colLetter, i, 1)) - Asc("A") + 1)
    Next i
    wsColToNum = res
End Function

Private Function GetOrCreateCockpitV6(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("COCKPIT_V6")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
        ws.Name = "COCKPIT_V6"
    End If
    Set GetOrCreateCockpitV6 = ws
End Function

Private Sub EnsureCockpitV6Headers(ws As Worksheet)
    If ws.cells(4, 1).value <> "Ticker" Then
        ws.cells(4, 1).value = "Ticker"
        ws.cells(4, 2).value = "Action"
        ws.cells(4, 3).value = "LTP"
        ws.cells(4, 4).value = "Sector"
        ws.cells(4, 17).value = "10 Vs 30"
        ws.cells(4, 19).value = "Ent"
        ws.cells(4, 20).value = "Tgt"
        ws.cells(4, 21).value = "SL"
        ws.cells(4, 22).value = "CP"
        ws.cells(4, 23).value = "Open Time"
        ws.cells(4, 24).value = "CloseTime"
        ws.cells(4, 26).value = "Bottomline"
        ws.cells(4, 27).value = "Coments"
        ws.cells(4, 28).value = "No_Of_Bars"
        ws.cells(4, 30).value = "Peak_Price"
        ws.cells(4, 31).value = "Best_MFE_Pct"
        ws.cells(4, 35).value = "Last_Bar_TS"
    End If
    ws.cells(4, 36).value = "Exec_State"
    ws.cells(4, 37).value = "FYERS_Order_ID"
    ws.cells(4, 38).value = "Filled_Qty"
    ws.cells(4, 39).value = "Requested_Qty"
    ws.cells(4, 40).value = "Request_ID"
        ws.cells(4, 41).value = "Exit_Fill_Px"
        ws.cells(4, 42).value = "Request_Type"
    ws.cells(4, 29).value = "Apply_To_FYERS"
    ws.cells(4, 43).value = "Trade_Origin"
    ws.cells(4, 44).value = "FYERS_Has_SL"
    ws.cells(4, 45).value = "FYERS_Has_TP"
End Sub

Private Function V6ApplyToFyers(ByVal flagValue As Variant) As Boolean
    ' Blank legacy rows remain real/applied trades. New monitor-only rows store FALSE.
    If IsError(flagValue) Then
        V6ApplyToFyers = False ' Invalid flag cells fail closed for live execution.
    ElseIf IsEmpty(flagValue) Or Len(Trim$(CStr(flagValue))) = 0 Then
        V6ApplyToFyers = True
    ElseIf VarType(flagValue) = vbBoolean Then
        V6ApplyToFyers = CBool(flagValue)
    Else
        Dim flagText As String: flagText = UCase$(Trim$(CStr(flagValue)))
        V6ApplyToFyers = (flagText = "TRUE" Or flagText = "A" Or flagText = "APPLY" Or flagText = "1")
    End If
End Function

' =====================================================================
' FYERS API -- POSITION/ORDER SNAPSHOT AND ENTRY IP PREFLIGHT
' Uses the staged Python manager so status parsing and SDK behavior stay aligned.
' =====================================================================
Private Function GetFyersOpenPositionCount() As Long
    Dim snapshot As String
    GetFyersOpenPositionCount = -1
    If Not GetV6BrokerSnapshot(snapshot) Then Exit Function
    GetFyersOpenPositionCount = CountSnapshotOpenPositions(snapshot)
End Function


Private Function GetV6BrokerSnapshot(ByRef snapshot As String) As Boolean
    Dim outputText As String
    GetV6BrokerSnapshot = False
    If Not RunV6Python("--snapshot", outputText, 30#) Then
        V6ReconcileLog "SNAPSHOT_FAILED; helper timed out or could not start; " & Left$(Replace(outputText, vbCrLf, " "), 180)
        Exit Function
    End If
    Dim p As Long: p = InStr(1, outputText, "V6SNAP|OK|", vbTextCompare)
    If p = 0 Then
        V6ReconcileLog "SNAPSHOT_FAILED; no success marker; " & Left$(Replace(outputText, vbCrLf, " "), 180)
        Exit Function
    End If
    snapshot = Mid$(outputText, p + Len("V6SNAP|OK|"))
    Dim outLines() As String: outLines = Split(snapshot, vbCrLf)
    snapshot = Trim$(outLines(0))
    If Len(snapshot) = 0 Then Exit Function
    GetV6BrokerSnapshot = True
End Function


Private Function RunV6Python(ByVal args As String, ByRef outputText As String, ByVal maxSeconds As Double) As Boolean
    On Error GoTo RunErr
    Dim pyExe As String: pyExe = "C:\Users\rkann\AppData\Local\Programs\Python\Python313\python.exe"
    If dir(pyExe) = "" Then pyExe = "python"
    Dim scriptPath As String: scriptPath = ThisWorkbook.path & "\engine\fyers_order_manager_v6.py"
    Dim runnerPath As String: runnerPath = ThisWorkbook.path & "\engine\v6_hidden_runner.py"
    Dim outputPath As String
    outputPath = Environ$("TEMP") & "\V6CLI_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 1000#)) & ".txt"
    Dim cmd As String
    cmd = """" & pyExe & """ """ & runnerPath & """ --timeout " & Format$(maxSeconds, "0") & _
          " --output """ & outputPath & """ -- """ & pyExe & """ """ & scriptPath & _
          """ --root """ & ThisWorkbook.path & """ " & args
    Dim wsh As Object, fso As Object, ts As Object, exitCode As Long
    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = ThisWorkbook.path
    exitCode = wsh.Run(cmd, 1, True) ' Keep the useful command window visible; wrapper enforces a hard timeout.
    Set fso = CreateObject("Scripting.FileSystemObject")
    outputText = ""
    If fso.FileExists(outputPath) Then
        Set ts = fso.OpenTextFile(outputPath, 1, False)
        outputText = ts.ReadAll
        ts.Close
        On Error Resume Next
        fso.DeleteFile outputPath, True
        On Error GoTo RunErr
    End If
    RunV6Python = (exitCode <> 124)
    Exit Function
RunErr:
    outputText = "V6HELPER|ERROR|" & Err.Description
    RunV6Python = False
End Function


Private Function V6EntryIPReady(wsC As Worksheet) As Boolean
    Dim outputText As String
    V6EntryIPReady = False
    If Not RunV6Python("--check-ip", outputText, 8#) Then
        wsC.cells(2, 14).value = "ENTRY BLOCKED: public IPv4 preflight unavailable."
        Exit Function
    End If
    If InStr(1, outputText, "V6IP|OK|", vbTextCompare) > 0 Then
        V6EntryIPReady = True
    Else
        Dim p As Long: p = InStr(1, outputText, "V6IP|BLOCK|", vbTextCompare)
        If p > 0 Then
            wsC.cells(2, 14).value = Left$(Trim$(Mid$(outputText, p)), 250)
        Else
            wsC.cells(2, 14).value = "ENTRY BLOCKED: IPv4 preflight failed."
        End If
    End If
End Function


Private Function CountSnapshotOpenPositions(ByVal snapshot As String) As Long
    Dim p As Long, key As Long, qty As Double, symStart As Long
    p = 1
    Do
        key = InStr(p, snapshot, """net_qty"":")
        If key = 0 Then Exit Do
        qty = JsonNumberAfter(snapshot, key + Len("""net_qty"":"))
        If qty <> 0 Then CountSnapshotOpenPositions = CountSnapshotOpenPositions + 1
        p = key + Len("""net_qty"":")
    Loop
End Function


Private Function SnapshotNumberForSymbol(ByVal snapshot As String, ByVal ticker As String, ByVal propertyName As String, Optional ByVal defaultValue As Double = 0#) As Double
    Dim p As Long, objStart As Long, objEnd As Long, prop As Long
    SnapshotNumberForSymbol = defaultValue
    ' Match the exact top-level symbol object. A plain ticker search can hit
    ' __orders__.symbol first and miss the symbol's pending count.
    p = InStr(1, snapshot, """" & UCase$(Trim$(ticker)) & """:{""net_qty"":", vbTextCompare)
    If p = 0 Then Exit Function
    objStart = InStr(p, snapshot, "{")
    objEnd = InStr(objStart, snapshot, "}")
    If objStart = 0 Or objEnd = 0 Then Exit Function
    prop = InStr(objStart, Mid$(snapshot, 1, objEnd), """" & propertyName & """:")
    If prop = 0 Then Exit Function
    SnapshotNumberForSymbol = JsonNumberAfter(snapshot, prop + Len("""" & propertyName & """:"))
End Function


Private Function SnapshotOrderNumber(ByVal snapshot As String, ByVal orderId As String, ByVal propertyName As String) As Double
    If Len(orderId) = 0 Then Exit Function
    Dim p As Long, objStart As Long, objEnd As Long, prop As Long
    p = InStr(1, snapshot, """" & orderId & """", vbTextCompare)
    If p = 0 Then Exit Function
    objStart = InStr(p, snapshot, "{")
    objEnd = InStr(objStart, snapshot, "}")
    If objStart = 0 Or objEnd = 0 Then Exit Function
    prop = InStr(objStart, Mid$(snapshot, 1, objEnd), """" & propertyName & """:")
    If prop > 0 Then SnapshotOrderNumber = JsonNumberAfter(snapshot, prop + Len("""" & propertyName & """:"))
End Function


Private Function SnapshotOrderStatus(ByVal snapshot As String, ByVal orderId As String) As Long
    SnapshotOrderStatus = -1
    If Len(orderId) > 0 Then SnapshotOrderStatus = CLng(SnapshotOrderNumber(snapshot, orderId, "status"))
End Function


Private Function V6CancelPendingEntry(wsC As Worksheet, ByVal rowNum As Long, ByVal ticker As String) As Boolean
    Dim orderId As String: orderId = Trim$(CStr(wsC.Range(CK_ORDER_ID & rowNum).value))
    If Len(orderId) = 0 Then Exit Function
    Dim outputText As String
    If Not RunV6Python("--cancel-order-id " & orderId, outputText, 8#) Then Exit Function
    V6CancelPendingEntry = (InStr(1, outputText, "V6CANCEL|CANCEL_REQUESTED|" & orderId & "|", vbTextCompare) > 0)
    V6ReconcileLog "row=" & rowNum & " ticker=" & ticker & " ENTRY_CANCEL_REQUEST order=" & orderId & _
        " accepted=" & V6CancelPendingEntry & " output=" & Replace(Trim$(outputText), vbCrLf, " ")
End Function


Private Function JsonNumberAfter(ByVal source As String, ByVal startAt As Long) As Double
    Dim i As Long, ch As String, token As String
    i = startAt
    Do While i <= Len(source) And (Mid$(source, i, 1) = " " Or Mid$(source, i, 1) = vbTab): i = i + 1: Loop
    Do While i <= Len(source)
        ch = Mid$(source, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "." Then
            token = token & ch: i = i + 1
        Else
            Exit Do
        End If
    Loop
    If Len(token) > 0 Then JsonNumberAfter = CDbl(val(token))
End Function


Private Function GetFyersAuthHeader() As String
    ' Reads fyers_tokens.json from the workbook folder, returns "app_id:access_token".
    Dim fso As Object, ts As Object, jsonText As String, path As String
    path = ThisWorkbook.path & "\fyers_tokens.json"
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If Not fso.FileExists(path) Then Exit Function
    Set ts = fso.OpenTextFile(path, 1)   ' 1 = ForReading
    jsonText = ts.ReadAll
    ts.Close
    On Error GoTo 0

    Dim appId As String, token As String
    appId = JsonStringValue(jsonText, "app_id")
    token = JsonStringValue(jsonText, "access_token")
    If Len(appId) = 0 Or Len(token) = 0 Then Exit Function
    GetFyersAuthHeader = appId & ":" & token
End Function


Private Function JsonStringValue(ByVal jsonText As String, ByVal key As String) As String
    ' Minimal extractor for a JSON string property. Null/numeric values return blank.
    Dim p As Long, c As Long, q1 As Long, q2 As Long, v As Long
    p = InStr(1, jsonText, """" & key & """", vbTextCompare)
    If p = 0 Then Exit Function
    c = InStr(p, jsonText, ":")
    If c = 0 Then Exit Function
    v = c + 1
    Do While v <= Len(jsonText) And (Mid$(jsonText, v, 1) = " " Or Mid$(jsonText, v, 1) = vbTab): v = v + 1: Loop
    If v > Len(jsonText) Or Mid$(jsonText, v, 1) <> """" Then Exit Function
    q1 = v
    q2 = InStr(q1 + 1, jsonText, """")
    If q2 = 0 Then Exit Function
    JsonStringValue = Mid$(jsonText, q1 + 1, q2 - q1 - 1)
End Function


Private Function CountNonZeroNetQty(ByVal jsonText As String) As Long
    ' Counts netPositions[] entries whose netQty is non-zero. -1 if not parseable.
    Dim startPos As Long, endPos As Long, arr As String
    startPos = InStr(1, jsonText, """netPositions""", vbTextCompare)
    If startPos = 0 Then CountNonZeroNetQty = -1: Exit Function
    startPos = InStr(startPos, jsonText, "[")
    If startPos = 0 Then CountNonZeroNetQty = -1: Exit Function
    endPos = InStr(startPos, jsonText, """overall""", vbTextCompare)
    If endPos = 0 Then endPos = Len(jsonText) + 1
    arr = Mid$(jsonText, startPos, endPos - startPos)

    Dim cnt As Long, posn As Long, c As Long, i As Long
    Dim numStr As String, ch As String
    cnt = 0: posn = 1
    Do
        posn = InStr(posn, arr, """netQty""", vbTextCompare)
        If posn = 0 Then Exit Do
        c = InStr(posn, arr, ":")
        If c = 0 Then Exit Do
        i = c + 1
        Do While i <= Len(arr) And Mid$(arr, i, 1) = " ": i = i + 1: Loop
        numStr = ""
        Do While i <= Len(arr)
            ch = Mid$(arr, i, 1)
            If ch = "-" Or (ch >= "0" And ch <= "9") Then
                numStr = numStr & ch
                i = i + 1
            Else
                Exit Do
            End If
        Loop
        If Len(numStr) > 0 Then
            If val(numStr) <> 0 Then cnt = cnt + 1
        End If
        posn = posn + 8
    Loop
    CountNonZeroNetQty = cnt
End Function


' =====================================================================
' Private Sub  SizeCandidates
' Dynamic sizing via engine/fyers_order_manager_v6.py --calc-qty.
' Python fetches live available capital + per-stock intraday margin and
' prints one "SIZE|TICKER=qty" line per candidate (in input order).
' qtys(i) = 0 when that candidate cannot be sized (skip it).
' =====================================================================
Private Sub SizeCandidates(ByVal symList As String, ByVal freeSlots As Long, ByRef qtys() As Long, ByVal n As Long)
    ReDim qtys(1 To n)
    Dim i As Long
    For i = 1 To n: qtys(i) = 0: Next i
    If n < 1 Or Len(symList) = 0 Or freeSlots < 1 Then Exit Sub

    Dim outTxt As String
    If Not RunV6Python("--calc-qty --slots " & freeSlots & " --symbols """ & symList & """", outTxt, 25#) Then Exit Sub

    ' Parse ONLY our marker lines (order preserved -> i-th marker = i-th candidate).
    Dim lines() As String, ln As String, p As Long, eq As Long
    Dim idx As Long: idx = 0
    lines = Split(Replace(outTxt, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        p = InStr(1, ln, "SIZE|", vbTextCompare)
        If p > 0 Then
            idx = idx + 1
            If idx > n Then Exit For
            eq = InStr(ln, "=")
            If eq > 0 Then qtys(idx) = CLng(val(Mid$(ln, eq + 1)))
        End If
    Next i
End Sub


' =====================================================================
' Private Function DispatchBrokerOrder
' Starts a hidden order process; request state is persisted and reconciled on later cycles.
' =====================================================================
Private Function DispatchBrokerOrder(ByVal action As String, ByVal ticker As String, ByVal qty As Long, _
        ByVal tag As String, ByVal requestId As String, Optional ByVal isExit As Boolean = False, _
        Optional ByVal slPct As Double = 0#, Optional ByVal tpPct As Double = 0#, _
        Optional ByVal estimatedPrice As Double = 0#, Optional ByVal entryLimitPrice As Double = 0#) As Boolean
    On Error GoTo DispatchErr
    DispatchBrokerOrder = False
    If qty <= 0 Or Len(requestId) = 0 Then Exit Function
    Dim pyExe As String: pyExe = "C:\Users\rkann\AppData\Local\Programs\Python\Python313\python.exe"
    If dir(pyExe) = "" Then pyExe = "python"
    Dim scriptPath As String: scriptPath = ThisWorkbook.path & "\engine\fyers_order_manager_v6.py"
    Dim extra As String: extra = ""
    If isExit Then extra = extra & " --exit"
    If isExit And estimatedPrice > 0 Then extra = extra & " --price " & Replace(Format(estimatedPrice, "0.00"), ",", ".")
    If Not isExit And entryLimitPrice > 0 Then extra = extra & " --limit-price " & Replace(Format(entryLimitPrice, "0.00"), ",", ".")
    If slPct > 0 Then extra = extra & " --sl-pct " & Replace(Format(slPct, "0.00"), ",", ".")
    If tpPct > 0 Then extra = extra & " --tp-pct " & Replace(Format(tpPct, "0.00"), ",", ".")
    Dim cmd As String
    cmd = """" & pyExe & """ """ & scriptPath & """ --root """ & ThisWorkbook.path & """ --symbol " & ticker & _
          " --action " & action & " --qty " & qty & " --tag """ & tag & """ --request-id " & requestId & extra

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = ThisWorkbook.path
    wsh.Run cmd, 0, False
    DispatchBrokerOrder = True
    Exit Function
DispatchErr:
    Debug.Print "V6 broker dispatcher launch failed: " & Err.Description
    DispatchBrokerOrder = False
End Function


Private Function NewV6RequestId(ByVal direction As String, ByVal ticker As String) As String
    NewV6RequestId = "V6" & direction & Format$(Now, "yymmddhhnnss") & _
                     Replace(UCase$(ticker), "-", "") & Format$(CLng((Timer - Int(Timer)) * 1000), "000")
End Function





Attribute VB_Name = "Module_Cockpit_V3"
'Attribute VB_Name = "Module_COCKPIT_V3"
Option Explicit

'----------------------------------------------------------------------
' Module_COCKPIT_V3  --  EMA10/EMA30 V3 Validated Architecture
'
' Target Worksheet: "COCKPIT_V3"
' Entry Rules:
'   1. Completed 5m EMA10/EMA30 Crossover (OVER for Long, UNDER for Short)
'   2. Trading Window: 09:20:00 to 15:00:00 (No trades on 09:15 open candle)
'   3. VWAP Alignment: Long > VWAP, Short < VWAP (Col BR)
'   4. Volume Confirmation: RVOL >= 1.0 (Col BP)
'   5. Progressive Pre-Cross Lookback: V_avg > 0 over current-day bars (Parsed from Col DM packet)
'   6. RSI Exhaustion Gate: Long <= 75, Short >= 25 (Col U)
'   7. Climax Overextension Gate: |LTP - EMA30| <= 1.5 * ATR
' Risk Management:
'   - Initial Stop Loss: 3-bar swing pivot capped at 1.0 ATR from Entry
'   - Kinetic Gate (Bar 3): If Gap_3 <= Gap_1, EJECT immediately at market
'   - Confirmed Runner: If Gap_3 > Gap_1, move SL to Break-Even (Entry Price)
'   - Trailing Runner: Price close crossing EMA30 against trade OR opposite EMA cross
'   - Hard Session Close: Mandatory liquidation at 15:05:00
'
' Formatting and layout 100% mirrors reference COCKPIT sheet.
'----------------------------------------------------------------------

' =====================================================================
'  TIME GATES & SESSION CONSTANTS
' =====================================================================
Private Const EARLY_ENTRY_START         As String = "09:20:00"
Private Const LATE_ENTRY_CUTOFF         As String = "15:00:00"
Private Const FORCE_LIQUIDATION         As String = "15:05:00"
Private Const DATA_START_ROW            As Long = 5
Private Const STRATEGY_NAME             As String = "EMA10_30_V3"

' =====================================================================
'  ENTRY FILTER THRESHOLDS
' =====================================================================
Private Const GATE_RVOL_MIN             As Double = 1#
Private Const GATE_RSI_BUY_MAX          As Double = 75#
Private Const GATE_RSI_SELL_MIN         As Double = 25#
Private Const GATE_MAX_CLIMAX_ATR       As Double = 1.5
Private Const GATE_MAX_RISK_ATR         As Double = 1#
Private Const TARGET_DEFAULT_ATR        As Double = 2#

' =====================================================================
'  COCKPIT_V3 COLUMN LETTERS
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
Private Const CK_OUTCOME      As String = "Y"      ' formula
Private Const CK_BOTTOMLINE   As String = "Z"      ' formula
Private Const CK_COMENTS      As String = "AA"
Private Const CK_CANDLE_COUNT As String = "AB"     ' execution cycles elapsed
Private Const CK_STRATEGY     As String = "AC"
Private Const CK_PEAK_PRICE   As String = "AD"     ' peak favorable price
Private Const CK_BEST_MFE_PCT As String = "AE"     ' peak favorable %
Private Const CK_LAST_MFE_BAR As String = "AF"     ' bar index of peak
Private Const CK_GAP_1        As String = "AG"     ' directional gap at Bar 1
Private Const CK_IS_RUNNER    As String = "AH"     ' 1 = confirmed runner, 0 = maturing

' Module-level caches
Private mRowLookup As Object
Private mActiveTrades As Object

' =====================================================================
' Public Sub  Run_Cockpit_V3
' Main Entry Point for COCKPIT_V3 Execution Engine
' =====================================================================
Public Sub Run_Cockpit_V3()
    Dim currentTime As Date
    currentTime = TimeValue(Now)

    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Dim wsSrc  As Worksheet: Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Dim wsC3   As Worksheet: Set wsC3 = GetOrCreateCockpitV3(ThisWorkbook)

    EnsureCockpitV3Headers wsC3
    Update_K3_Dashboard wsSrc, wsC3
    Set mRowLookup = BuildRowLookup(wsSrc)
    Set mActiveTrades = BuildActiveTradeLookup(wsC3)

    ' -- TIME GATE: Outside trading hours, only manage existing open trades --
    If currentTime < TimeValue(EARLY_ENTRY_START) Or currentTime >= TimeValue(LATE_ENTRY_CUTOFF) Then
        ManageCockpitV3Trades wsC3, wsSrc
        GoTo Cleanup
    End If

    ' -- Scan Sheet1 via in-memory 2D array --
    Dim srcLastRow As Long
    srcLastRow = wsSrc.cells(wsSrc.Rows.count, 1).End(xlUp).row
    If srcLastRow < 4 Then GoTo Cleanup

    Dim srcData As Variant
    srcData = wsSrc.Range("A1:DM" & srcLastRow).Value2

    Dim lastDestRow As Long
    lastDestRow = wsC3.cells(wsC3.Rows.count, 1).End(xlUp).row
    If lastDestRow < DATA_START_ROW Then lastDestRow = DATA_START_ROW - 1

    Dim r As Long, destRow As Long
    destRow = lastDestRow

    For r = 4 To srcLastRow
        Dim ticker As String: ticker = Trim$(CStr(srcData(r, 1)))
        If Len(ticker) = 0 Or Is_Index(ticker) Then GoTo NextStock

        ' 1. Read Crossover from Col AC (Col 29) in memory
        Dim crossStr As String: crossStr = UCase$(Trim$(CStr(srcData(r, 29))))
        Dim isBuyCross As Boolean: isBuyCross = False
        Dim isSellCross As Boolean: isSellCross = False

        If crossStr = "BULLISH" Then
            isBuyCross = True
        ElseIf crossStr = "BEARISH" Then
            isSellCross = True
        Else
            GoTo NextStock
        End If

        ' Prevent duplicate active entries for the same stock
        Dim tradeDir As String: tradeDir = IIf(isBuyCross, "BUY", "SELL")
        If mActiveTrades.Exists(ticker & "|" & tradeDir) Then GoTo NextStock

        ' 2. Fetch technicals from memory array
        Dim ltp As Double:   ltp = CDbl(val(srcData(r, 20)))     ' Col T
        Dim vwap As Double:  vwap = CDbl(val(srcData(r, 70)))    ' Col BR
        Dim rvol As Double:  rvol = CDbl(val(srcData(r, 68)))    ' Col BP
        Dim ema10 As Double: ema10 = CDbl(val(srcData(r, 45)))   ' Col AS
        Dim ema30 As Double: ema30 = CDbl(val(srcData(r, 46)))   ' Col AT
        Dim atr As Double:   atr = CDbl(val(srcData(r, 81)))     ' Col CC
        Dim rsi5m As Double: rsi5m = CDbl(val(srcData(r, 21)))   ' Col U

        If ltp <= 0 Or atr <= 0 Then GoTo NextStock

        ' 3. VWAP Alignment Filter
        If isBuyCross And ltp <= vwap Then GoTo NextStock
        If isSellCross And ltp >= vwap Then GoTo NextStock

        ' 4. Volume Floor Filter (RVOL >= 1.0) I have disabled this currently to see if the rest of the logic works
        'If rvol < GATE_RVOL_MIN Then GoTo NextStock

        ' 5. RSI Exhaustion Gate (Wilder / Brown Overbought/Oversold Floor)
        If isBuyCross And rsi5m > GATE_RSI_BUY_MAX Then GoTo NextStock
        If isSellCross And rsi5m < GATE_RSI_SELL_MIN Then GoTo NextStock

        ' 6. Climax Overextension Gate (|LTP - EMA30| <= 1.5 * ATR)
        If Abs(ltp - ema30) > (GATE_MAX_CLIMAX_ATR * atr) Then GoTo NextStock

        ' 7. Lookback Telemetry Packet & Velocity Gate (Col DM = Col 117)
        Dim dmPacket As String: dmPacket = Trim$(CStr(srcData(r, 117)))
        If Len(dmPacket) = 0 Then GoTo NextStock

        Dim parts() As String: parts = Split(dmPacket, "|")
        If UBound(parts) < 5 Then GoTo NextStock

        Dim g0 As Double, gk As Double, cATR As Double, k As Long
        g0 = CDbl(val(Split(parts(0), "=")(1)))
        cATR = CDbl(val(Split(parts(4), "=")(1)))
        k = CLng(val(Split(parts(5), "=")(1)))

        If k < 1 Or cATR <= 0 Then GoTo NextStock

        If k = 1 Then
            gk = CDbl(val(Split(parts(1), "=")(1)))
        ElseIf k = 2 Then
            gk = CDbl(val(Split(parts(2), "=")(1)))
        Else
            gk = CDbl(val(Split(parts(3), "=")(1)))
        End If

        ' Compute directional average velocity
        Dim vAvg As Double
        If isBuyCross Then
            vAvg = (g0 - gk) / (k * cATR)
        Else
            vAvg = ((-g0) - (-gk)) / (k * cATR)
        End If

        ' Reject if pre-cross gap was contracting or sluggish
        If vAvg <= 0 Then GoTo NextStock

        ' 8. Initial Stop Loss Calculation (3-bar swing pivot capped at 1.0 ATR)
        Dim swingLow As Double:  swingLow = CDbl(val(srcData(r, 94)))   ' Col CP
        Dim swingHigh As Double: swingHigh = CDbl(val(srcData(r, 93)))  ' Col CO
        Dim slPx As Double, tgtPx As Double

        If isBuyCross Then
            If swingLow > 0 And swingLow < ltp Then
                slPx = swingLow
            Else
                slPx = ltp - (atr * GATE_MAX_RISK_ATR)
            End If
            ' Cap risk at maximum 1.0 ATR
            If (ltp - slPx) > (atr * GATE_MAX_RISK_ATR) Then slPx = ltp - (atr * GATE_MAX_RISK_ATR)
            tgtPx = ltp + (atr * TARGET_DEFAULT_ATR)
        Else
            If swingHigh > 0 And swingHigh > ltp Then
                slPx = swingHigh
            Else
                slPx = ltp + (atr * GATE_MAX_RISK_ATR)
            End If
            ' Cap risk at maximum 1.0 ATR
            If (slPx - ltp) > (atr * GATE_MAX_RISK_ATR) Then slPx = ltp + (atr * GATE_MAX_RISK_ATR)
            tgtPx = ltp - (atr * TARGET_DEFAULT_ATR)
        End If

        ' Write new trade into COCKPIT_V3
        destRow = destRow + 1
        Dim dirGap As Double: dirGap = IIf(isBuyCross, (ema10 - ema30), (ema30 - ema10))
        WriteNewCockpitV3Row wsC3, wsSrc, r, destRow, ticker, tradeDir, ltp, slPx, tgtPx, dirGap, rvol, vwap, rsi5m, vAvg, dmPacket
        mActiveTrades(ticker & "|" & tradeDir) = True

NextStock:
    Next r

    ' Manage existing open positions
    ManageCockpitV3Trades wsC3, wsSrc

Cleanup:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = True
    MsgBox "Error in Run_Cockpit_V3: " & Err.Description, vbCritical
End Sub

' =====================================================================
' Private Sub  WriteNewCockpitV3Row
' Writes an admitted trade row into COCKPIT_V3 matching reference formatting
' =====================================================================
Private Sub WriteNewCockpitV3Row(wsC As Worksheet, wsSrc As Worksheet, _
        srcRow As Long, destRow As Long, ticker As String, action As String, _
        ltp As Double, slPx As Double, tgtPx As Double, dirGap As Double, _
        rvol As Double, vwap As Double, rsi5m As Double, vAvg As Double, dmPacket As String)

    wsC.cells(destRow, RangeRef(CK_TICKER)).value = ticker
    wsC.cells(destRow, RangeRef(CK_ACTION)).value = action
    wsC.cells(destRow, RangeRef(CK_LTP)).value = ltp
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
    wsC.cells(destRow, RangeRef(CK_CROSS_10_30)).value = wsSrc.cells(srcRow, "AC").value
    wsC.cells(destRow, RangeRef(CK_CROSS_10_50)).value = wsSrc.cells(srcRow, "AG").value

    ' Entry, Target, SL, CP, Open Time
    wsC.cells(destRow, RangeRef(CK_ENT)).value = ltp
    wsC.cells(destRow, RangeRef(CK_TGT)).value = tgtPx
    wsC.cells(destRow, RangeRef(CK_SL)).value = slPx
    wsC.cells(destRow, RangeRef(CK_CP)).value = ltp
    wsC.cells(destRow, RangeRef(CK_OPEN_TIME)).value = Now

    ' Standard Outcome & Bottomline Formulas
'    wsC.cells(destRow, RangeRef(CK_OUTCOME)).Formula = _
'        "=IF(ISBLANK(V" & destRow & "),"""",IF(B" & destRow & "=""BUY"",(V" & destRow & "-S" & destRow & ")/S" & destRow & ",(S" & destRow & "-V" & destRow & ")/S" & destRow & "))"
'    wsC.cells(destRow, RangeRef(CK_BOTTOMLINE)).Formula = _
'        "=IF(ISBLANK(X" & destRow & "),"""",IF(B" & destRow & "=""BUY"",(V" & destRow & "-S" & destRow & ")/S" & destRow & ",(S" & destRow & "-V" & destRow & ")/S" & destRow & "))"

    ' Comments with Full Audit Trail
    wsC.cells(destRow, RangeRef(CK_COMENTS)).value = _
        "ENTERED | Vavg=" & Format(vAvg, "0.00") & " RSI=" & Format(rsi5m, "0.0") & " | Telemetry: " & dmPacket

    wsC.cells(destRow, RangeRef(CK_CANDLE_COUNT)).value = 0
    wsC.cells(destRow, RangeRef(CK_STRATEGY)).value = STRATEGY_NAME
    wsC.cells(destRow, RangeRef(CK_PEAK_PRICE)).value = ltp
    wsC.cells(destRow, RangeRef(CK_BEST_MFE_PCT)).value = 0#
    wsC.cells(destRow, RangeRef(CK_LAST_MFE_BAR)).value = 0
    wsC.cells(destRow, RangeRef(CK_GAP_1)).value = dirGap
    wsC.cells(destRow, RangeRef(CK_IS_RUNNER)).value = 0

    ApplyCockpitColors wsC, destRow, (action = "BUY")
End Sub

' =====================================================================
' Private Sub  ManageCockpitV3Trades
' Manages active open trades: Kinetic Gate, Trailing EMA30, Break-even SL
' =====================================================================
Private Sub ManageCockpitV3Trades(wsC As Worksheet, wsSrc As Worksheet)
    Dim lastRow As Long
    lastRow = wsC.cells(wsC.Rows.count, 1).End(xlUp).row
    If lastRow < DATA_START_ROW Then Exit Sub

    Dim currentTime As Date: currentTime = TimeValue(Now)
    Dim forceLiquidate As Boolean: forceLiquidate = (currentTime >= TimeValue(FORCE_LIQUIDATION))

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim closeTime As Variant: closeTime = wsC.Range(CK_CLOSE_TIME & r).value
        If Not IsEmpty(closeTime) And Len(CStr(closeTime)) > 0 Then GoTo NextTrade

        Dim ticker As String:    ticker = CStr(wsC.Range(CK_TICKER & r).value)
        Dim entryPx As Double:   entryPx = val(wsC.Range(CK_ENT & r).value)
        Dim tgtPx As Double:     tgtPx = val(wsC.Range(CK_TGT & r).value)
        Dim slPx As Double:      slPx = val(wsC.Range(CK_SL & r).value)
        Dim action As String:    action = UCase(CStr(wsC.Range(CK_ACTION & r).value))
        Dim candleCount As Long: candleCount = val(wsC.Range(CK_CANDLE_COUNT & r).value)
        Dim gap1 As Double:      gap1 = val(wsC.Range(CK_GAP_1 & r).value)
        Dim isRunner As Long:    isRunner = val(wsC.Range(CK_IS_RUNNER & r).value)

        If Len(ticker) = 0 Or entryPx <= 0 Then GoTo NextTrade

        ' Update LTP / CP
        Dim curLTP As Double: curLTP = GetCurrentLTP(wsSrc, ticker, entryPx)
        wsC.Range(CK_CP & r).value = curLTP

        ' Increment cycle counter
        candleCount = candleCount + 1
        wsC.Range(CK_CANDLE_COUNT & r).value = candleCount

        ' Fetch current EMA values
        Dim curEMA10 As Double: curEMA10 = GetEMA10(wsSrc, ticker)
        Dim curEMA30 As Double: curEMA30 = GetEMA30(wsSrc, ticker)
        Dim curGap As Double
        curGap = IIf(action = "BUY", (curEMA10 - curEMA30), (curEMA30 - curEMA10))

        ' Record Bar 1 gap if not stored
        If candleCount = 1 And gap1 <= 0 Then
            gap1 = curGap
            wsC.Range(CK_GAP_1 & r).value = gap1
        End If

        ' Update peak MFE
        Dim curMfePct As Double
        curMfePct = IIf(action = "BUY", (curLTP - entryPx) / entryPx, (entryPx - curLTP) / entryPx)
        Dim bestMfe As Double: bestMfe = val(wsC.Range(CK_BEST_MFE_PCT & r).value)
        If curMfePct > bestMfe Then
            wsC.Range(CK_BEST_MFE_PCT & r).value = curMfePct
            wsC.Range(CK_PEAK_PRICE & r).value = curLTP
            wsC.Range(CK_LAST_MFE_BAR & r).value = candleCount
        End If

        ' --- EXIT EVALUATION ---
        Dim exitFlag As Boolean: exitFlag = False
        Dim exitReason As String: exitReason = ""

        ' 1. Mandatory Session Liquidation (15:05:00)
        If forceLiquidate Then
            exitFlag = True
            exitReason = "SESSION_CLOSE (15:05 Hard Exit)"

        ' 2. Initial Stop Loss Hit
        ElseIf action = "BUY" And curLTP <= slPx Then
            exitFlag = True
            exitReason = IIf(isRunner = 1, "STOP_LOSS (Break-Even Hit)", "STOP_LOSS (Initial SL Hit)")
        ElseIf action = "SELL" And curLTP >= slPx Then
            exitFlag = True
            exitReason = IIf(isRunner = 1, "STOP_LOSS (Break-Even Hit)", "STOP_LOSS (Initial SL Hit)")

        ' 3. Bar-3 Kinetic Gate (Launch Confirmation vs Eject)
        ElseIf candleCount = 3 And isRunner = 0 Then
            If curGap <= gap1 Then
                exitFlag = True
                exitReason = "KINETIC_EJECT (Bar 3 Gap Stalled: G3<=G1)"
            Else
                ' Confirmed Runner! Move SL to Break-Even
                isRunner = 1
                wsC.Range(CK_IS_RUNNER & r).value = 1
                wsC.Range(CK_SL & r).value = entryPx
                wsC.Range(CK_COMENTS & r).value = "CONFIRMED RUNNER | BE SL Activated | Gap3=" & Format(curGap, "0.00")
            End If

        ' 4. Runner Management (Post Bar-3)
        ElseIf isRunner = 1 Then
            ' A: Trailing EMA30 Close
            If action = "BUY" And curLTP < curEMA30 Then
                exitFlag = True
                exitReason = "TRAILING_EXIT (Price Closed Below EMA30)"
            ElseIf action = "SELL" And curLTP > curEMA30 Then
                exitFlag = True
                exitReason = "TRAILING_EXIT (Price Closed Above EMA30)"
            End If

            ' B: Opposite EMA Crossover
            Dim oppCross As String: oppCross = UCase$(Trim$(CStr(GetCross1030(wsSrc, ticker))))
            If action = "BUY" And oppCross = "BEARISH" Then
                exitFlag = True
                exitReason = "OPPOSITE_CROSS (EMA10 Under EMA30)"
            ElseIf action = "SELL" And oppCross = "BULLISH" Then
                exitFlag = True
                exitReason = "OPPOSITE_CROSS (EMA10 Over EMA30)"
            End If
        End If

        ' Execute Exit
        If exitFlag Then
            wsC.Range(CK_CLOSE_TIME & r).value = Now
            wsC.Range(CK_COMENTS & r).value = exitReason
            wsC.Range(CK_CP & r).value = curLTP
        End If

NextTrade:
    Next r
End Sub

' =====================================================================
' Helper Functions & State Management
' =====================================================================
Private Function GetOrCreateCockpitV3(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("COCKPIT_V3")
    On Error GoTo 0

    If ws Is Nothing Then
        wb.Sheets("COCKPIT").Copy After:=wb.Sheets("COCKPIT")
        Set ws = wb.ActiveSheet
        ws.Name = "COCKPIT_V3"

        ' Clear data rows
        Dim lr As Long: lr = ws.cells(ws.Rows.count, 1).End(xlUp).row
        If lr >= DATA_START_ROW Then
            ws.Range("A" & DATA_START_ROW & ":AH" & lr).ClearContents
            ws.Range("A" & DATA_START_ROW & ":AH" & lr).Interior.ColorIndex = xlNone
        End If
    End If
    Set GetOrCreateCockpitV3 = ws
End Function

Private Sub EnsureCockpitV3Headers(wsC As Worksheet)
    If Len(CStr(wsC.Range(CK_PEAK_PRICE & "4").value)) = 0 Then wsC.Range(CK_PEAK_PRICE & "4").value = "Peak_Price"
    If Len(CStr(wsC.Range(CK_BEST_MFE_PCT & "4").value)) = 0 Then wsC.Range(CK_BEST_MFE_PCT & "4").value = "Best_MFE_Pct"
    If Len(CStr(wsC.Range(CK_LAST_MFE_BAR & "4").value)) = 0 Then wsC.Range(CK_LAST_MFE_BAR & "4").value = "Last_MFE_Bar"
    If Len(CStr(wsC.Range(CK_GAP_1 & "4").value)) = 0 Then wsC.Range(CK_GAP_1 & "4").value = "Gap_1"
    If Len(CStr(wsC.Range(CK_IS_RUNNER & "4").value)) = 0 Then wsC.Range(CK_IS_RUNNER & "4").value = "Is_Runner"

    wsC.Columns(CK_PEAK_PRICE & ":" & CK_IS_RUNNER).Hidden = True
End Sub

Private Function BuildRowLookup(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim lastRow As Long: lastRow = ws.cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 4 Then Set BuildRowLookup = d: Exit Function

    Dim vals As Variant: vals = ws.Range("A4:A" & lastRow).value
    Dim i As Long
    For i = 1 To UBound(vals, 1)
        Dim t As String: t = Trim$(CStr(vals(i, 1)))
        If Len(t) > 0 And Not d.Exists(t) Then d(t) = i + 3
    Next i
    Set BuildRowLookup = d
End Function

Private Function BuildActiveTradeLookup(wsC As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim lastRow As Long: lastRow = wsC.cells(wsC.Rows.count, 1).End(xlUp).row
    If lastRow < DATA_START_ROW Then Set BuildActiveTradeLookup = d: Exit Function

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim closeVal As Variant: closeVal = wsC.Range(CK_CLOSE_TIME & r).value
        If IsEmpty(closeVal) Or Len(CStr(closeVal)) = 0 Then
            Dim ticker As String: ticker = Trim$(CStr(wsC.Range(CK_TICKER & r).value))
            Dim action As String: action = Trim$(CStr(wsC.Range(CK_ACTION & r).value))
            If Len(ticker) > 0 Then d(ticker & "|" & action) = True
        End If
    Next r
    Set BuildActiveTradeLookup = d
End Function

Private Function GetCurrentLTP(wsSrc As Worksheet, ByVal ticker As String, ByVal fallback As Double) As Double
    GetCurrentLTP = fallback
    If mRowLookup Is Nothing Then Exit Function
    If mRowLookup.Exists(ticker) Then
        Dim v As Double: v = CDbl(val(wsSrc.cells(CLng(mRowLookup(ticker)), "T").value))
        If v > 0 Then GetCurrentLTP = v
    End If
End Function

Private Function GetEMA10(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetEMA10 = 0
    If mRowLookup Is Nothing Then Exit Function
    If mRowLookup.Exists(ticker) Then
        GetEMA10 = CDbl(val(wsSrc.cells(CLng(mRowLookup(ticker)), "AS").value))
    End If
End Function

Private Function GetEMA30(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetEMA30 = 0
    If mRowLookup Is Nothing Then Exit Function
    If mRowLookup.Exists(ticker) Then
        GetEMA30 = CDbl(val(wsSrc.cells(CLng(mRowLookup(ticker)), "AT").value))
    End If
End Function

Private Function GetCross1030(wsSrc As Worksheet, ByVal ticker As String) As String
    GetCross1030 = ""
    If mRowLookup Is Nothing Then Exit Function
    If mRowLookup.Exists(ticker) Then
        GetCross1030 = CStr(wsSrc.cells(CLng(mRowLookup(ticker)), "AC").value)
    End If
End Function

Private Function Is_Index(t As String) As Boolean
    Dim u As String: u = UCase$(t)
    Is_Index = (InStr(u, "NIFTY") > 0 Or InStr(u, "BANKNIFTY") > 0 Or _
                InStr(u, "CNX") > 0 Or InStr(u, "MIDCAP") > 0 Or _
                InStr(u, "SMALLCAP") > 0 Or InStr(u, "LARGECAP") > 0)
End Function

Private Function RangeRef(ByVal letter As String) As Long
    Dim n As Long: n = 0
    Dim i As Long
    For i = 1 To Len(letter)
        n = n * 26 + (Asc(UCase$(Mid$(letter, i, 1))) - 64)
    Next i
    RangeRef = n
End Function

Private Sub ApplyCockpitColors(wsC As Worksheet, destRow As Long, ByVal isBuy As Boolean)
    Dim actionCell As Range: Set actionCell = wsC.cells(destRow, RangeRef(CK_ACTION))
    If isBuy Then
        actionCell.Interior.Color = RGB(220, 245, 220)
        actionCell.Font.Color = RGB(0, 100, 0)
        actionCell.Font.Bold = True
    Else
        actionCell.Interior.Color = RGB(255, 225, 225)
        actionCell.Font.Color = RGB(150, 0, 0)
        actionCell.Font.Bold = True
    End If
End Sub





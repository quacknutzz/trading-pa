Attribute VB_Name = "Module_Cockpit_V4"
'Attribute VB_Name = "Module_COCKPIT_V4"
Option Explicit

'----------------------------------------------------------------------
' Module_COCKPIT_V4  --  BTB V4.2b EMA-Aligned Structural Engine
'
' Target Worksheet: "COCKPIT_V4"
' Architecture:
'   Discovery:  Donchian 5-bar Breakout (col DF) OR 5m EMA10/30 Cross (col AC)
'   Quality:    Regime/sector alignment (col BN/BM) + grade (col DE)
'   Filters:    EMA10/30 directional alignment + RSI cap (25-75) + Climax gate (1.5 ATR)
'   Risk:       3-bar swing SL capped at 1.0 ATR, floored at 0.35%
'   Exits:      STRUCTURAL ONLY with BTB-gated expiry permission:
'                 1. Mandatory Session Liquidation (15:05:00)
'                 2. Stop Loss Hit (Initial swing SL)
'                 3. Target Hit (Fixed 2:1 R:R)
'                 4. Max holding expiry at 12 candles unless fresh BTB agrees
'                    and the trade does not carry failed-breakout telemetry
'               Non-structural checks (Kinetic, Trailing EMA30, Profit Retrace,
'               Dead Trade) are converted to non-spamming telemetry tags only.
' Data Provenance: 100% from Sheet1 (zero NI dependency)
' Layout: 100% mirrors COCKPIT_V3 reference sheet
'----------------------------------------------------------------------

' =====================================================================
'  TIME GATES & SESSION CONSTANTS
' =====================================================================
Private Const EARLY_ENTRY_START         As String = "09:20:00"
Private Const LATE_ENTRY_CUTOFF         As String = "15:00:00"
Private Const FORCE_LIQUIDATION         As String = "15:05:00"
Private Const DATA_START_ROW            As Long = 5
Private Const STRATEGY_NAME_DONCHIAN    As String = "V4.2b_Donchian"

' =====================================================================
'  DISCOVERY GATES (from V2)
' =====================================================================
Private Const GATE_RVOL_MIN             As Double = 1.25     ' Donchian breakout needs volume surge

' =====================================================================
'  QUALITY GATES
' =====================================================================
' Regime: Longs may work in bull/up/ranging contexts. Shorts require stock and sector downside alignment.
' Sector: Contra sectors are blocked for longs; shorts require TRENDING_BEAR or MILD_DOWN sector.
' Grade: A+, A, B pass. C rejected.

' =====================================================================
'  ENTRY FILTER THRESHOLDS (from V3)
' =====================================================================
Private Const GATE_RSI_BUY_MAX          As Double = 75#      ' Don't buy overbought exhaustion
Private Const GATE_RSI_SELL_MIN         As Double = 25#      ' Don't short oversold exhaustion
Private Const GATE_MAX_CLIMAX_ATR       As Double = 1.5      ' |LTP - EMA30| must be <= 1.5 ATR

' =====================================================================
'  RISK MANAGEMENT
' =====================================================================
Private Const GATE_MAX_RISK_ATR         As Double = 1#       ' SL capped at 1.0 ATR from entry
Private Const RISK_FLOOR_MIN_PCT        As Double = 0.0035   ' 0.35% minimum stop risk floor (from V2)
Private Const TARGET_RR_MULTIPLE        As Double = 2#       ' Target = 2.0 x initial_risk (adaptive R:R)

' =====================================================================
'  POSITION SIZING
' =====================================================================
Private Const SIZE_DEFAULT              As Double = 0.5      ' Default position size

' =====================================================================
'  EXIT MANAGEMENT & TELEMETRY CONSTANTS (V4.2b EMA-Aligned Structural)
' =====================================================================
Private Const MAX_CANDLES_EXIT          As Long = 12         ' Max holding expiry
Private Const MAX_TOTAL_CANDLES_EXIT    As Long = 24         ' Hard timeout for BTB-extended trades
Private Const BTB_EXTEND_SCORE_MIN      As Double = 60#      ' Fresh same-side BTB score required to extend
Private Const BTB_CONTRA_SCORE_MIN      As Double = 65#      ' Fresh opposite-side BTB score required to veto/exit
Private Const STAGNATION_BARS           As Long = 5          ' Stagnation telemetry threshold (diagnostics only)
Private Const STAGNATION_MIN_MFE_PCT    As Double = 0.0005   ' 0.05% minimum favorable excursion for telemetry
Private Const DEAD_TRADE_EXIT_MAX_PCT   As Double = -0.003   ' -0.3% adverse threshold for stagnation telemetry
Private Const PROFIT_RETRACE_FRAC       As Double = 0.4      ' 40% giveback from peak MFE for retrace telemetry
Private Const PROFIT_RETRACE_MIN_PCT    As Double = 0.002    ' Minimum 0.2% giveback for retrace telemetry tag
Private Const PROFIT_ARM_MIN_PCT        As Double = 0.004    ' Arm profit telemetry after +0.4% MFE
Private Const COOLDOWN_MINUTES_VAL      As Long = 30
Private Const MAX_ATTEMPTS_PER_TICKER   As Long = 2

' =====================================================================
'  COCKPIT_V4 COLUMN LETTERS (mirrors V3 reference exactly)
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
Private Const CK_MFE_BAR1     As String = "AG"     ' MFE pct at Bar 1 (for kinetic follow-through gate)
Private Const CK_IS_RUNNER    As String = "AH"     ' 1 = confirmed runner, 0 = maturing

' Sheet1 BTB shadow feature columns (Python-owned; VBA consumes only when DataQuality=0)
Private Const BTB_EXT_SIDE     As String = "DV"
Private Const BTB_EXT_DECISION As String = "DW"
Private Const BTB_EXT_SCORE    As String = "DX"
Private Const BTB_DATA_QUALITY As String = "DY"

' Module-level caches
Private mRowLookup As Object
Private mActiveTrades As Object
Private mCooldowns As Object
Private mTickerAttempts As Object


' =====================================================================
' Public Sub  Run_Cockpit_V4
' Main Entry Point for COCKPIT_V4 Unified Engine
' =====================================================================
Public Sub Run_Cockpit_V4()
    Dim currentTime As Date
    currentTime = TimeValue(Now)

    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Dim wsSrc  As Worksheet: Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Dim wsC4   As Worksheet: Set wsC4 = GetOrCreateCockpitV4(ThisWorkbook)

    EnsureCockpitV4Headers wsC4
    On Error Resume Next
    Update_K3_Dashboard wsSrc, wsC4
    On Error GoTo ErrorHandler
    Set mRowLookup = BuildRowLookup(wsSrc)
    Set mActiveTrades = BuildActiveTradeLookup(wsC4)
    Set mTickerAttempts = BuildTickerSessionAttemptLookup(wsC4)

    ' -- TIME GATE: Outside trading window, only manage existing open trades --
    If currentTime < TimeValue(EARLY_ENTRY_START) Or currentTime >= TimeValue(LATE_ENTRY_CUTOFF) Then
        ManageCockpitV4Trades wsC4, wsSrc
        GoTo Cleanup
    End If

    ' -- Scan Sheet1 via in-memory 2D array --
    Dim srcLastRow As Long
    srcLastRow = wsSrc.cells(wsSrc.Rows.count, 1).End(xlUp).row
    If srcLastRow < 4 Then GoTo Cleanup

    Dim srcData As Variant
    srcData = wsSrc.Range("A1:DF" & srcLastRow).Value2

    Dim lastDestRow As Long
    lastDestRow = wsC4.cells(wsC4.Rows.count, 1).End(xlUp).row
    If lastDestRow < DATA_START_ROW Then lastDestRow = DATA_START_ROW - 1

    Dim r As Long, destRow As Long
    destRow = lastDestRow

    For r = 4 To srcLastRow
        Dim ticker As String: ticker = Trim$(CStr(srcData(r, 1)))
        If Len(ticker) = 0 Or Is_Index(ticker) Then GoTo NextStock

        ' === DISCOVERY: Dual Trigger (Donchian Breakout col DF OR 5m EMA10/30 Cross col AC) ===
        Dim breakoutFlag As Long: breakoutFlag = CLng(val(srcData(r, 110)))
        Dim crossStr As String: crossStr = UCase$(Trim$(CStr(srcData(r, 29))))

        Dim boBuy As Boolean:  boBuy = (breakoutFlag = 1)
        Dim boSell As Boolean: boSell = (breakoutFlag = -1)
        Dim emaBuy As Boolean: emaBuy = (crossStr = "BULLISH")
        Dim emaSell As Boolean: emaSell = (crossStr = "BEARISH")

        Dim isBuy As Boolean
        Dim stratName As String
        Dim discovTag As String

        If (boBuy Or emaBuy) And Not (boSell Or emaSell) Then
            isBuy = True
            If boBuy And emaBuy Then
                stratName = "V4.2b_Dual_Long"
                discovTag = "Dual_BO_Cross"
            ElseIf boBuy Then
                stratName = "V4.2b_Donchian_Long"
                discovTag = "Donchian_BO"
            Else
                stratName = "V4.2b_EMACross_Long"
                discovTag = "EMA10_30_Cross"
            End If
        ElseIf (boSell Or emaSell) And Not (boBuy Or emaBuy) Then
            isBuy = False
            If boSell And emaSell Then
                stratName = "V4.2b_Dual_Short"
                discovTag = "Dual_BO_Cross"
            ElseIf boSell Then
                stratName = "V4.2b_Donchian_Short"
                discovTag = "Donchian_BO"
            Else
                stratName = "V4.2b_EMACross_Short"
                discovTag = "EMA10_30_Cross"
            End If
        Else
            GoTo NextStock  ' No discovery event, or conflicting signals (e.g. Bullish BO + Bearish Cross)
        End If

        ' === Volume Telemetry (captured for analytics, hard rejection gate removed) ===
        Dim rvol As Double: rvol = CDbl(val(srcData(r, 67)))

        ' === Fetch core technicals from memory array ===
        Dim ltp As Double:         ltp = CDbl(val(srcData(r, 20)))
        Dim ema10 As Double:       ema10 = CDbl(val(srcData(r, 45)))
        Dim ema30 As Double:       ema30 = CDbl(val(srcData(r, 46)))
        Dim atr As Double:         atr = CDbl(val(srcData(r, 81)))
        Dim rsi5m As Double:       rsi5m = CDbl(val(srcData(r, 21)))
        Dim stockRegime As String: stockRegime = UCase$(Trim$(CStr(srcData(r, 66))))
        Dim stockGrade As String:  stockGrade = UCase$(Trim$(CStr(srcData(r, 109))))

        If ltp <= 0 Or atr <= 0 Then GoTo NextStock

        ' === GATE 2b: EMA10/30 Directional Alignment ===
        ' Breakouts are admitted only when the fast/slow 5m trend state agrees with trade direction.
        If isBuy Then
            If ema10 <= ema30 Then GoTo NextStock
        Else
            If ema10 >= ema30 Then GoTo NextStock
        End If

        ' === GATE 3: Regime Filter (simplified Gemini) ===
        If isBuy Then
            ' Longs allowed in: TRENDING_BULL, MILD_UP, RANGING
            If stockRegime = "TRENDING_BEAR" Or stockRegime = "MILD_DOWN" Then GoTo NextStock
        Else
            ' Shorts now require explicit downside alignment. RANGING shorts were a material drag in live review.
            If Not (stockRegime = "TRENDING_BEAR" Or stockRegime = "MILD_DOWN") Then GoTo NextStock
        End If

        ' === GATE 4: Sector Contra Block (from Gemini) ===
        Dim secRegime As String: secRegime = GetSectorRegime(wsSrc, ticker)
        If isBuy Then
            If secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN" Then GoTo NextStock
        Else
            If Not (secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN") Then GoTo NextStock
        End If

        ' === GATE 5: Grade Filter ===
        If Not (stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B") Then GoTo NextStock

        ' === GATE 6: RSI Exhaustion Cap (from V3) ===
        If isBuy And rsi5m > GATE_RSI_BUY_MAX Then GoTo NextStock
        If Not isBuy And rsi5m < GATE_RSI_SELL_MIN Then GoTo NextStock

        ' === GATE 7: Climax Overextension (from V3) ===
        If Abs(ltp - ema30) > (GATE_MAX_CLIMAX_ATR * atr) Then GoTo NextStock

        ' === DUPLICATE / CHURN CHECK ===
        Dim tradeDir As String: tradeDir = IIf(isBuy, "BUY", "SELL")
        ' Keep at most one open trade per ticker. This blocks same-cycle opposite-side churn.
        If mActiveTrades.Exists(ticker & "|ANY") Then GoTo NextStock

        ' === SESSION ATTEMPT CAP ===
        Dim sessionAttempts As Long: sessionAttempts = 0
        If Not mTickerAttempts Is Nothing Then
            If mTickerAttempts.Exists(ticker) Then sessionAttempts = CLng(mTickerAttempts(ticker))
        End If
        If sessionAttempts >= MAX_ATTEMPTS_PER_TICKER Then GoTo NextStock

        ' === COOLDOWN CHECK ===
        If IsOnCooldown(ticker) Then GoTo NextStock

        ' === V4.2b BTB CONTRA VETO ===
        ' Fresh Python BTB features may veto a strong opposite-side setup.
        ' If BTB data is stale/bad/missing (DataQuality <> 0), this gate is ignored.
        If ShouldBTBOpposeTrade(wsSrc, ticker, tradeDir, BTB_CONTRA_SCORE_MIN) Then GoTo NextStock

        ' === ALL GATES PASSED — Calculate Risk & Write Trade ===

        ' Initial SL: 3-bar swing pivot capped at 1.0 ATR, floored at 0.35%
        Dim swingLow As Double:  swingLow = CDbl(val(srcData(r, 94)))   ' Col CP
        Dim swingHigh As Double: swingHigh = CDbl(val(srcData(r, 93)))  ' Col CO
        Dim slPx As Double, tgtPx As Double
        Dim initialRisk As Double

        If isBuy Then
            If swingLow > 0 And swingLow < ltp Then
                slPx = swingLow
            Else
                slPx = ltp - (atr * GATE_MAX_RISK_ATR)
            End If
            ' Cap risk at 1.0 ATR
            If (ltp - slPx) > (atr * GATE_MAX_RISK_ATR) Then slPx = ltp - (atr * GATE_MAX_RISK_ATR)
            initialRisk = ltp - slPx
        Else
            If swingHigh > 0 And swingHigh > ltp Then
                slPx = swingHigh
            Else
                slPx = ltp + (atr * GATE_MAX_RISK_ATR)
            End If
            ' Cap risk at 1.0 ATR
            If (slPx - ltp) > (atr * GATE_MAX_RISK_ATR) Then slPx = ltp + (atr * GATE_MAX_RISK_ATR)
            initialRisk = slPx - ltp
        End If

        ' Apply 0.35% risk floor with proportional size adjustment (from V2)
        Dim minRisk As Double: minRisk = ltp * RISK_FLOOR_MIN_PCT
        Dim finalSize As Double: finalSize = SIZE_DEFAULT

        If initialRisk > 0 And initialRisk < minRisk And minRisk > 0 Then
            finalSize = Round(SIZE_DEFAULT * (initialRisk / minRisk), 4)
            initialRisk = minRisk
            If isBuy Then slPx = ltp - initialRisk Else slPx = ltp + initialRisk
        End If
        If initialRisk <= 0 Then initialRisk = ltp * RISK_FLOOR_MIN_PCT

        ' Adaptive target: 2:1 R:R
        If isBuy Then
            tgtPx = ltp + (initialRisk * TARGET_RR_MULTIPLE)
        Else
            tgtPx = ltp - (initialRisk * TARGET_RR_MULTIPLE)
        End If

        ' EMA alignment scoring for comments (not a gate)
        Dim emaAligned As Boolean
        If isBuy Then emaAligned = (ema10 > ema30) Else emaAligned = (ema10 < ema30)

        ' Write trade (MFE_BAR1 starts at 0, updated at first management cycle)
        destRow = destRow + 1
        WriteNewCockpitV4Row wsC4, wsSrc, r, destRow, ticker, tradeDir, ltp, slPx, tgtPx, _
                             rvol, rsi5m, emaAligned, stockRegime, stockGrade, _
                             secRegime, initialRisk, finalSize, stratName, discovTag
        mActiveTrades(ticker & "|" & tradeDir) = True
        mActiveTrades(ticker & "|ANY") = True
        If Not mTickerAttempts Is Nothing Then
            If mTickerAttempts.Exists(ticker) Then
                mTickerAttempts(ticker) = mTickerAttempts(ticker) + 1
            Else
                mTickerAttempts(ticker) = 1
            End If
        End If

NextStock:
    Next r

    ' Manage existing open positions
    ManageCockpitV4Trades wsC4, wsSrc

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
    MsgBox "Error in Run_Cockpit_V4: " & Err.Description, vbCritical
End Sub


' =====================================================================
' Private Sub  WriteNewCockpitV4Row
' Writes an admitted trade row into COCKPIT_V4 matching V3 reference formatting
' =====================================================================
Private Sub WriteNewCockpitV4Row(wsC As Worksheet, wsSrc As Worksheet, _
        srcRow As Long, destRow As Long, ticker As String, action As String, _
        ltp As Double, slPx As Double, tgtPx As Double, _
        rvol As Double, rsi5m As Double, emaAligned As Boolean, _
        stockRegime As String, stockGrade As String, secRegime As String, _
        initialRisk As Double, finalSize As Double, _
        ByVal stratName As String, ByVal discovTag As String)

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
    'wsC.cells(destRow, RangeRef(CK_OUTCOME)).Formula = _
        "=IF(ISBLANK(V" & destRow & "),"""",IF(B" & destRow & "=""BUY"",(V" & destRow & "-S" & destRow & ")/S" & destRow & ",(S" & destRow & "-V" & destRow & ")/S" & destRow & "))"
    'wsC.cells(destRow, RangeRef(CK_BOTTOMLINE)).Formula = _
        "=IF(ISBLANK(X" & destRow & "),"""",IF(B" & destRow & "=""BUY"",(V" & destRow & "-S" & destRow & ")/S" & destRow & ",(S" & destRow & "-V" & destRow & ")/S" & destRow & "))"

    ' Comments with audit trail
    Dim emaTag As String: emaTag = IIf(emaAligned, "EMA_OK", "EMA_LAG")
    wsC.cells(destRow, RangeRef(CK_COMENTS)).value = _
        "V4.2b_ENTRY | " & discovTag & " | " & emaTag & _
        " | Reg=" & stockRegime & " Sec=" & secRegime & " Grd=" & stockGrade & _
        " | RVOL=" & Format(rvol, "0.00") & " RSI=" & Format(rsi5m, "0.0") & _
        " | Risk=" & Format(initialRisk, "0.00") & " Size=" & Format(finalSize, "0.00")

    wsC.cells(destRow, RangeRef(CK_CANDLE_COUNT)).value = 0
    wsC.cells(destRow, RangeRef(CK_STRATEGY)).value = stratName
    wsC.cells(destRow, RangeRef(CK_PEAK_PRICE)).value = ltp
    wsC.cells(destRow, RangeRef(CK_BEST_MFE_PCT)).value = 0#
    wsC.cells(destRow, RangeRef(CK_LAST_MFE_BAR)).value = 0
    wsC.cells(destRow, RangeRef(CK_MFE_BAR1)).value = 0#   ' Updated at first management cycle
    wsC.cells(destRow, RangeRef(CK_IS_RUNNER)).value = 0
    ApplyCockpitColors wsC, destRow, (action = "BUY")
End Sub


' =====================================================================
' Private Sub  ManageCockpitV4Trades
' Full lifecycle: Pure structural exits + adaptive expiry extension
' + non-exiting telemetry for kinetic/retrace/EMA/opposite-cross/stagnation states
' =====================================================================
Private Sub ManageCockpitV4Trades(wsC As Worksheet, wsSrc As Worksheet)
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
        Dim mfeBar1 As Double:   mfeBar1 = val(wsC.Range(CK_MFE_BAR1 & r).value)
        Dim isRunner As Long:    isRunner = val(wsC.Range(CK_IS_RUNNER & r).value)

        If Len(ticker) = 0 Or entryPx <= 0 Then GoTo NextTrade

        ' Update LTP / CP
        Dim curLTP As Double: curLTP = GetCurrentLTP(wsSrc, ticker, entryPx)
        wsC.Range(CK_CP & r).value = curLTP

        ' Increment cycle counter
        candleCount = candleCount + 1
        wsC.Range(CK_CANDLE_COUNT & r).value = candleCount

        ' Fetch current EMA value for EMA breach telemetry.
        Dim curEMA30 As Double: curEMA30 = GetEMA30(wsSrc, ticker)

        ' Compute current MFE %
        Dim curMfePct As Double
        curMfePct = IIf(action = "BUY", (curLTP - entryPx) / entryPx, (entryPx - curLTP) / entryPx)

        ' Record Bar 1 MFE for Kinetic follow-through gate
        If candleCount = 1 Then
            mfeBar1 = curMfePct
            wsC.Range(CK_MFE_BAR1 & r).value = mfeBar1
        End If

        ' Update peak MFE
        Dim bestMfe As Double: bestMfe = val(wsC.Range(CK_BEST_MFE_PCT & r).value)
        Dim lastMfeBar As Long: lastMfeBar = CLng(val(wsC.Range(CK_LAST_MFE_BAR & r).value))
        If curMfePct > bestMfe Then
            bestMfe = curMfePct
            lastMfeBar = candleCount
            wsC.Range(CK_BEST_MFE_PCT & r).value = bestMfe
            wsC.Range(CK_PEAK_PRICE & r).value = curLTP
            wsC.Range(CK_LAST_MFE_BAR & r).value = lastMfeBar
        End If

        ' --- EXIT EVALUATION (priority-ordered) ---
        Dim exitFlag As Boolean: exitFlag = False
        Dim exitReason As String: exitReason = ""

        ' 1. Mandatory Session Liquidation (15:05:00)
        If forceLiquidate Then
            exitFlag = True
            exitReason = "SESSION_CLOSE (15:05 Hard Exit)"

        ' 2. Stop Loss Hit
        ElseIf action = "BUY" And curLTP <= slPx Then
            exitFlag = True
            exitReason = IIf(slPx = entryPx, "STOP_LOSS (Break-Even Hit)", "STOP_LOSS (Initial SL Hit)")
        ElseIf action = "SELL" And curLTP >= slPx Then
            exitFlag = True
            exitReason = IIf(slPx = entryPx, "STOP_LOSS (Break-Even Hit)", "STOP_LOSS (Initial SL Hit)")

        ' 3. Target Hit
        ElseIf action = "BUY" And tgtPx > 0 And curLTP >= tgtPx Then
            exitFlag = True
            exitReason = "TARGET_HIT"
        ElseIf action = "SELL" And tgtPx > 0 And curLTP <= tgtPx Then
            exitFlag = True
            exitReason = "TARGET_HIT"

        ' 4. BTB contra failure during extension
        ElseIf candleCount > MAX_CANDLES_EXIT And ShouldBTBOpposeTrade(wsSrc, ticker, action, BTB_CONTRA_SCORE_MIN) Then
            exitFlag = True
            exitReason = "BTB_CONTRA_EXIT"

        ' 5. Hard timeout for BTB-extended trades
        ElseIf candleCount >= MAX_TOTAL_CANDLES_EXIT Then
            exitFlag = True
            exitReason = "BTB_EXTENSION_TIMEOUT (" & MAX_TOTAL_CANDLES_EXIT & " candles)"

        ' 6. Max Candle Expiry (12 candles / 60 minutes) unless fresh same-side BTB permits extension.
        '    Never extend failed-breakout telemetry: KINETIC_STAGNANT + EMA30_BREACH, or OPP_CROSS.
        ElseIf candleCount >= MAX_CANDLES_EXIT Then
            Dim expiryCmt As String: expiryCmt = CStr(wsC.Range(CK_COMENTS & r).value)
            If HasFailedBreakoutTelemetry(expiryCmt) Then
                exitFlag = True
                exitReason = "EXPIRED (" & MAX_CANDLES_EXIT & " candles; weak breakout telemetry)"
            ElseIf ShouldBTBExtendTrade(wsSrc, ticker, action) Then
                Dim extCmt As String: extCmt = CStr(wsC.Range(CK_COMENTS & r).value)
                If InStr(extCmt, "BTB_EXTENDED") = 0 Then
                    wsC.Range(CK_COMENTS & r).value = extCmt & " | BTB_EXTENDED"
                End If
            Else
            exitFlag = True
            exitReason = "EXPIRED (" & MAX_CANDLES_EXIT & " candles)"
            End If
        End If

        ' --- V4.1 PURE STRUCTURAL TELEMETRY (Non-Exiting Diagnostics) ---
        ' Evaluated strictly while trade remains active (skipped if structural exit triggered).
        ' Anti-Spam Guardrail: Each tag is appended at most once per trade.
        If Not exitFlag Then
            Dim curComment As String: curComment = CStr(wsC.Range(CK_COMENTS & r).value)

            ' T1. Kinetic Follow-Through Telemetry (Bar 3 / 4)
            ' Intentional: isRunner remains 0 for stagnant trades to enable downstream telemetry filtering of runner vs stagnant resolution.
            If (candleCount = 3 Or candleCount = 4) And isRunner = 0 Then
                If InStr(curComment, "KINETIC_") = 0 Then
                    If curMfePct <= mfeBar1 Or curMfePct <= 0 Then
                        curComment = curComment & " | KINETIC_STAGNANT (Bar " & candleCount & " MFE=" & _
                                     Format(curMfePct * 100, "0.00") & "% vs Bar1=" & Format(mfeBar1 * 100, "0.00") & "%)"
                    Else
                        isRunner = 1
                        wsC.Range(CK_IS_RUNNER & r).value = 1
                        curComment = curComment & " | KINETIC_RUNNER (Bar " & candleCount & " MFE=" & _
                                     Format(curMfePct * 100, "0.00") & "% vs Bar1=" & Format(mfeBar1 * 100, "0.00") & "%)"
                    End If
                    wsC.Range(CK_COMENTS & r).value = curComment
                End If
            End If

            ' T2. Profit Retrace Telemetry (Peak >= 0.50% with >= 40% giveback)
            If bestMfe >= PROFIT_ARM_MIN_PCT Then
                Dim givebackPct As Double: givebackPct = bestMfe - curMfePct
                Dim retraceTrigger As Double: retraceTrigger = bestMfe * PROFIT_RETRACE_FRAC
                If retraceTrigger < PROFIT_RETRACE_MIN_PCT Then retraceTrigger = PROFIT_RETRACE_MIN_PCT

                If curMfePct > 0 And givebackPct >= retraceTrigger Then
                    If InStr(curComment, "RETRACE_WARN") = 0 Then
                        curComment = curComment & " | RETRACE_WARN (Peak=" & Format(bestMfe * 100, "0.00") & _
                                     "%, Now=" & Format(curMfePct * 100, "0.00") & "%)"
                        wsC.Range(CK_COMENTS & r).value = curComment
                    End If
                End If
            End If

            ' T3. Trailing EMA30 Breach Telemetry
            If (action = "BUY" And curLTP < curEMA30) Or (action = "SELL" And curLTP > curEMA30) Then
                If InStr(curComment, "EMA30_BREACH") = 0 Then
                    curComment = curComment & " | EMA30_BREACH"
                    wsC.Range(CK_COMENTS & r).value = curComment
                End If
            End If

            ' T4. Opposite EMA Crossover Telemetry
            Dim oppCross As String: oppCross = UCase$(Trim$(CStr(GetCross1030(wsSrc, ticker))))
            If (action = "BUY" And oppCross = "BEARISH") Or (action = "SELL" And oppCross = "BULLISH") Then
                If InStr(curComment, "OPP_CROSS") = 0 Then
                    curComment = curComment & " | OPP_CROSS"
                    wsC.Range(CK_COMENTS & r).value = curComment
                End If
            End If

            ' T5. Dead Trade Stagnation Telemetry
            If candleCount >= STAGNATION_BARS Then
                If bestMfe < STAGNATION_MIN_MFE_PCT And curMfePct <= DEAD_TRADE_EXIT_MAX_PCT Then
                    If InStr(curComment, "DEAD_TRADE_WARN") = 0 Then
                        curComment = curComment & " | DEAD_TRADE_WARN (" & STAGNATION_BARS & " bars, MFE=" & _
                                     Format(bestMfe * 100, "0.00") & "%)"
                        wsC.Range(CK_COMENTS & r).value = curComment
                    End If
                End If
            End If
        End If

        ' Execute Exit
        If exitFlag Then
            wsC.Range(CK_CLOSE_TIME & r).value = Now
            wsC.Range(CK_CP & r).value = curLTP
            Dim prevComment As String: prevComment = CStr(wsC.Range(CK_COMENTS & r).value)
            wsC.Range(CK_COMENTS & r).value = prevComment & " -> " & exitReason
            RecordCooldown ticker
        End If

NextTrade:
    Next r
End Sub


' =====================================================================
' Helper Functions & State Management
' =====================================================================

Private Function GetOrCreateCockpitV4(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("COCKPIT_V4")
    On Error GoTo 0

    If ws Is Nothing Then
        ' Try to copy from COCKPIT_V3 first, then COCKPIT
        On Error Resume Next
        Dim srcSheet As Worksheet
        Set srcSheet = wb.Sheets("COCKPIT_V3")
        If srcSheet Is Nothing Then Set srcSheet = wb.Sheets("COCKPIT")
        On Error GoTo 0

        If Not srcSheet Is Nothing Then
            srcSheet.Copy After:=srcSheet
            Set ws = wb.ActiveSheet
            ws.Name = "COCKPIT_V4"

            ' Clear data rows but keep headers and formatting
            Dim lr As Long: lr = ws.cells(ws.Rows.count, 1).End(xlUp).row
            If lr >= DATA_START_ROW Then
                ws.Range("A" & DATA_START_ROW & ":AH" & lr).ClearContents
                ws.Range("A" & DATA_START_ROW & ":AH" & lr).Interior.ColorIndex = xlNone
            End If
        Else
            ' Fallback: create blank sheet
            Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
            ws.Name = "COCKPIT_V4"
        End If
    End If
    Set GetOrCreateCockpitV4 = ws
End Function

Private Sub EnsureCockpitV4Headers(wsC As Worksheet)
    Dim lastDataRow As Long
    lastDataRow = wsC.cells(wsC.Rows.count, 1).End(xlUp).row
    If lastDataRow < 4 Then lastDataRow = 4

    ' AI:BB are not used by V4.1. Clear visible extension/legacy analytics baggage once,
    ' then leave the columns hidden for audit cleanliness.
    If Len(CStr(wsC.Range("AI4").value)) > 0 Then
        wsC.Range("AI4:BB" & lastDataRow).ClearContents
        wsC.Range("AI4:BB" & lastDataRow).Interior.ColorIndex = xlNone
    End If

    wsC.Range(CK_PEAK_PRICE & "4").value = "Peak_Price"
    wsC.Range(CK_BEST_MFE_PCT & "4").value = "Best_MFE_Pct"
    wsC.Range(CK_LAST_MFE_BAR & "4").value = "Last_MFE_Bar"
    wsC.Range(CK_MFE_BAR1 & "4").value = "MFE_Bar1"
    wsC.Range(CK_IS_RUNNER & "4").value = "Is_Runner"
    wsC.Columns(CK_PEAK_PRICE & ":" & CK_IS_RUNNER).Hidden = True
    wsC.Columns("AI:BB").Hidden = True
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
            If Len(ticker) > 0 Then d(ticker & "|ANY") = True
        End If
    Next r
    Set BuildActiveTradeLookup = d
End Function

Private Function BuildTickerSessionAttemptLookup(wsC As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = wsC.cells(wsC.Rows.count, 1).End(xlUp).row
    If lastRow < DATA_START_ROW Then
        Set BuildTickerSessionAttemptLookup = dict
        Exit Function
    End If

    Dim todayDate As Date: todayDate = Date
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim ticker As String: ticker = Trim$(CStr(wsC.Range(CK_TICKER & r).value))
        If Len(ticker) > 0 Then
            Dim openTimeVal As String: openTimeVal = Trim$(CStr(wsC.Range(CK_OPEN_TIME & r).value))
            Dim isToday As Boolean: isToday = False

            If IsDate(openTimeVal) Then
                If DateValue(CDate(openTimeVal)) = todayDate Then isToday = True
            End If

            If isToday Then
                If dict.Exists(ticker) Then
                    dict(ticker) = dict(ticker) + 1
                Else
                    dict(ticker) = 1
                End If
            End If
        End If
    Next r

    Set BuildTickerSessionAttemptLookup = dict
End Function

Private Function GetCurrentLTP(wsSrc As Worksheet, ByVal ticker As String, ByVal fallback As Double) As Double
    GetCurrentLTP = fallback
    If mRowLookup Is Nothing Then Exit Function
    If mRowLookup.Exists(ticker) Then
        Dim v As Double: v = CDbl(val(wsSrc.cells(CLng(mRowLookup(ticker)), "T").value))
        If v > 0 Then GetCurrentLTP = v
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

Private Function GetBTBRow(wsSrc As Worksheet, ByVal ticker As String) As Long
    GetBTBRow = 0
    If Len(ticker) = 0 Then Exit Function
    If mRowLookup Is Nothing Then Exit Function
    If mRowLookup.Exists(ticker) Then GetBTBRow = CLng(mRowLookup(ticker))
End Function

Private Function CleanBTBSide(ByVal rawSide As String) As String
    Dim s As String: s = UCase$(Trim$(rawSide))
    CleanBTBSide = "NEUTRAL"
    If InStr(s, "BUY") > 0 Then
        CleanBTBSide = "BUY"
    ElseIf InStr(s, "SELL") > 0 Then
        CleanBTBSide = "SELL"
    End If
End Function

Private Function GetBTBDataQuality(wsSrc As Worksheet, ByVal ticker As String) As Long
    Dim rowNum As Long: rowNum = GetBTBRow(wsSrc, ticker)
    If rowNum <= 0 Then
        GetBTBDataQuality = -1
    Else
        GetBTBDataQuality = CLng(val(wsSrc.Range(BTB_DATA_QUALITY & rowNum).value))
    End If
End Function

Private Function GetBTBExtSide(wsSrc As Worksheet, ByVal ticker As String) As String
    Dim rowNum As Long: rowNum = GetBTBRow(wsSrc, ticker)
    If rowNum <= 0 Then
        GetBTBExtSide = "NEUTRAL"
    Else
        GetBTBExtSide = CleanBTBSide(CStr(wsSrc.Range(BTB_EXT_SIDE & rowNum).value))
    End If
End Function

Private Function GetBTBExtDecision(wsSrc As Worksheet, ByVal ticker As String) As String
    Dim rowNum As Long: rowNum = GetBTBRow(wsSrc, ticker)
    If rowNum <= 0 Then
        GetBTBExtDecision = "NEUTRAL"
    Else
        GetBTBExtDecision = UCase$(Trim$(CStr(wsSrc.Range(BTB_EXT_DECISION & rowNum).value)))
    End If
End Function

Private Function GetBTBExtScore(wsSrc As Worksheet, ByVal ticker As String) As Double
    Dim rowNum As Long: rowNum = GetBTBRow(wsSrc, ticker)
    If rowNum <= 0 Then
        GetBTBExtScore = 0#
    Else
        GetBTBExtScore = CDbl(val(wsSrc.Range(BTB_EXT_SCORE & rowNum).value))
    End If
End Function

Private Function ShouldBTBOpposeTrade(wsSrc As Worksheet, ByVal ticker As String, ByVal action As String, ByVal minScore As Double) As Boolean
    ShouldBTBOpposeTrade = False
    If GetBTBDataQuality(wsSrc, ticker) <> 0 Then Exit Function

    Dim btbSide As String: btbSide = GetBTBExtSide(wsSrc, ticker)
    Dim btbScore As Double: btbScore = GetBTBExtScore(wsSrc, ticker)
    If btbScore < minScore Then Exit Function

    action = UCase$(Trim$(action))
    If action = "BUY" And btbSide = "SELL" Then ShouldBTBOpposeTrade = True
    If action = "SELL" And btbSide = "BUY" Then ShouldBTBOpposeTrade = True
End Function

Private Function ShouldBTBExtendTrade(wsSrc As Worksheet, ByVal ticker As String, ByVal action As String) As Boolean
    ShouldBTBExtendTrade = False
    If GetBTBDataQuality(wsSrc, ticker) <> 0 Then Exit Function
    If GetBTBExtDecision(wsSrc, ticker) <> "EXTEND" Then Exit Function
    If GetBTBExtScore(wsSrc, ticker) < BTB_EXTEND_SCORE_MIN Then Exit Function

    Dim btbSide As String: btbSide = GetBTBExtSide(wsSrc, ticker)
    action = UCase$(Trim$(action))
    If action = "BUY" And btbSide = "BUY" Then ShouldBTBExtendTrade = True
    If action = "SELL" And btbSide = "SELL" Then ShouldBTBExtendTrade = True
End Function

Private Function HasFailedBreakoutTelemetry(ByVal comments As String) As Boolean
    Dim hasStagnant As Boolean
    Dim hasEmaBreach As Boolean
    Dim hasOppCross As Boolean

    hasStagnant = (InStr(1, comments, "KINETIC_STAGNANT", vbTextCompare) > 0)
    hasEmaBreach = (InStr(1, comments, "EMA30_BREACH", vbTextCompare) > 0)
    hasOppCross = (InStr(1, comments, "OPP_CROSS", vbTextCompare) > 0)

    HasFailedBreakoutTelemetry = ((hasStagnant And hasEmaBreach) Or hasOppCross)
End Function

Private Function GetSectorRegime(wsSrc As Worksheet, ByVal ticker As String) As String
    GetSectorRegime = ""
    If Len(ticker) = 0 Then Exit Function
    Dim sectorName As String: sectorName = ""
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then sectorName = CStr(wsSrc.cells(CLng(mRowLookup(ticker)), "BM").value)
    End If
    If Len(sectorName) > 0 Then
        If Not mRowLookup Is Nothing Then
            If mRowLookup.Exists(sectorName) Then
                GetSectorRegime = UCase$(CStr(wsSrc.cells(CLng(mRowLookup(sectorName)), "BN").value))
            End If
        End If
    End If
End Function

Private Function Is_Index(t As String) As Boolean
    Dim u As String: u = UCase$(t)
    Is_Index = (InStr(u, "NIFTY") > 0 Or InStr(u, "CNX") > 0 Or _
                InStr(u, "NSEBANK") > 0 Or InStr(u, "MIDCAP") > 0 Or _
                InStr(u, "SMALLCAP") > 0 Or InStr(u, "LARGECAP") > 0 Or _
                InStr(u, "NSEI") > 0 Or InStr(u, "BANKNIFTY") > 0)
End Function

Private Function IsOnCooldown(ticker As String) As Boolean
    IsOnCooldown = False
    If COOLDOWN_MINUTES_VAL <= 0 Then Exit Function
    If mCooldowns Is Nothing Then Set mCooldowns = CreateObject("Scripting.Dictionary")
    If Not mCooldowns.Exists(ticker) Then Exit Function
    If DateDiff("n", mCooldowns(ticker), Now) < COOLDOWN_MINUTES_VAL Then
        IsOnCooldown = True
    Else
        mCooldowns.Remove ticker
    End If
End Function

Private Sub RecordCooldown(ticker As String)
    If COOLDOWN_MINUTES_VAL <= 0 Then Exit Sub
    If mCooldowns Is Nothing Then Set mCooldowns = CreateObject("Scripting.Dictionary")
    mCooldowns(ticker) = Now
End Sub

Private Function RangeRef(ByVal letter As String) As Long
    Dim n As Long: n = 0
    Dim i As Long
    For i = 1 To Len(letter)
        n = n * 26 + (Asc(UCase$(Mid$(letter, i, 1))) - 64)
    Next i
    RangeRef = n
End Function

Private Sub ApplyCockpitColors(wsC As Worksheet, destRow As Long, ByVal isBuy As Boolean)
    ' Action cell (B)
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

    ' Ticker font color (A)
    Dim tickerCell As Range: Set tickerCell = wsC.cells(destRow, RangeRef(CK_TICKER))
    If isBuy Then
        tickerCell.Font.Color = RGB(0, 150, 0)
    Else
        tickerCell.Font.Color = RGB(180, 0, 0)
    End If

    ' RSI threshold coloring (L-O)
    Dim rsiCols As Variant: rsiCols = Array(CK_RSI_5M, CK_RSI_15M, CK_RSI_1H, CK_RSI_1D)
    Dim i As Long
    For i = LBound(rsiCols) To UBound(rsiCols)
        Dim rsiVal As Double: rsiVal = CDbl(val(wsC.cells(destRow, RangeRef(CStr(rsiCols(i)))).value))
        If rsiVal >= 80 Then
            wsC.cells(destRow, RangeRef(CStr(rsiCols(i)))).Interior.Color = RGB(210, 180, 140)  ' tan
        ElseIf (rsiVal >= 70 And rsiVal < 80) Or (rsiVal <= 30 And rsiVal >= 21) Then
            wsC.cells(destRow, RangeRef(CStr(rsiCols(i)))).Interior.Color = RGB(255, 255, 160)  ' yellow
        End If
    Next i

    ' EMA cross coloring (Q, R)
    Dim cross1030 As String: cross1030 = LCase$(CStr(wsC.cells(destRow, RangeRef(CK_CROSS_10_30)).value))
    If cross1030 = "bullish" Then
        wsC.cells(destRow, RangeRef(CK_CROSS_10_30)).Interior.Color = RGB(204, 255, 204)
        wsC.cells(destRow, RangeRef(CK_CROSS_10_30)).Font.Bold = True
    ElseIf cross1030 = "bearish" Then
        wsC.cells(destRow, RangeRef(CK_CROSS_10_30)).Interior.Color = RGB(255, 204, 204)
        wsC.cells(destRow, RangeRef(CK_CROSS_10_30)).Font.Bold = True
    End If

    Dim cross1050 As String: cross1050 = LCase$(CStr(wsC.cells(destRow, RangeRef(CK_CROSS_10_50)).value))
    If cross1050 = "bullish" Then
        wsC.cells(destRow, RangeRef(CK_CROSS_10_50)).Interior.Color = RGB(204, 255, 204)
        wsC.cells(destRow, RangeRef(CK_CROSS_10_50)).Font.Bold = True
    ElseIf cross1050 = "bearish" Then
        wsC.cells(destRow, RangeRef(CK_CROSS_10_50)).Interior.Color = RGB(255, 204, 204)
        wsC.cells(destRow, RangeRef(CK_CROSS_10_50)).Font.Bold = True
    End If

    ' SuperTrend coloring (P)
    Dim st As String: st = LCase$(CStr(wsC.cells(destRow, RangeRef(CK_SUPERTREND)).value))
    If InStr(st, "bullish") > 0 Or val(wsC.cells(destRow, RangeRef(CK_SUPERTREND)).value) > 0 Then
        wsC.cells(destRow, RangeRef(CK_SUPERTREND)).Interior.Color = RGB(204, 255, 204)
        wsC.cells(destRow, RangeRef(CK_SUPERTREND)).Font.Bold = True
    ElseIf InStr(st, "bearish") > 0 Then
        wsC.cells(destRow, RangeRef(CK_SUPERTREND)).Interior.Color = RGB(255, 204, 204)
        wsC.cells(destRow, RangeRef(CK_SUPERTREND)).Font.Bold = True
    End If
End Sub



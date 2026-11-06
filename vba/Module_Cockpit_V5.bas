Attribute VB_Name = "Module_Cockpit_V5"
'Attribute VB_Name = "Module_COCKPIT_V5"
Option Explicit

'----------------------------------------------------------------------
' Module_COCKPIT_V5  --  BTB V5.0 Donchian-Confirmed EMA/V3 Engine
'
' Target Worksheet: "COCKPIT_V5"
' Status: ACTIVE — Donchian breakout confirmed by directional EMA quality.
'
' Architecture (when activated):
'   Discovery:  Configurable combination of two independent trigger families:
'                 Family 1 (EMA Cross):  5m EMA10/30 cross with V5 or V3 gates
'                 Family 2 (Donchian):   5-bar session-reset high/low breakout
'   Combination Mode (COMBINATION_MODE constant):
'     "SCAFFOLD"          — No entries admitted (current default)
'     "DONCHIAN_CONFIRMED"— Donchian + at least one EMA gate variant passes
'     "UNION"             — Any single signal passes
'     "DONCHIAN_ONLY"     — Donchian branch standalone
'     "EMA_V5_ONLY"       — EMA cross + V5 gates standalone
'     "EMA_V3_ONLY"       — EMA cross + V3 gates standalone
'   Risk:       3-bar swing SL capped at 1.0 ATR, floored at 0.35%
'   Exits:      V5 Common Lifecycle (shared across all ablation tiers):
'                 1. Session Liquidation (15:05:00)
'                 2. Stop Loss Hit
'                 3. Target Hit (2:1 R:R)
'                 4. Trailing Breakeven (1.0 ATR MFE threshold)
'                 5. Profit Reversal (40% giveback from peak)
'                 6. Dead Trade (4 bars with < 0.3% MFE)
'                 7. 9-candle expiry
'   Telemetry:  Kinetic, EMA30 breach, opposite cross, retrace, stagnation
'               — diagnostic tags only, non-exiting.
'
' Data Provenance: 100% from Sheet1 (zero NI dependency)
' Layout: Mirrors COCKPIT_V4 reference sheet
'
'----------------------------------------------------------------------

' =====================================================================
'  COMBINATION MODE CONTROLLER
' =====================================================================
Public Const COMBINATION_MODE           As String = "DONCHIAN_CONFIRMED"
Private Const ENTRY_RULES_VALIDATED      As Boolean = True
' Options: "SCAFFOLD", "DONCHIAN_CONFIRMED", "UNION",
'          "DONCHIAN_ONLY", "EMA_V5_ONLY", "EMA_V3_ONLY"
'
' DONCHIAN_CONFIRMED requires a Donchian breakout, Donchian quality gates, and
' at least one EMA quality branch aligned with the breakout direction. A fresh
' EMA cross on the same scan is not required.

' =====================================================================
'  TIME GATES & SESSION CONSTANTS
' =====================================================================
Private Const EARLY_ENTRY_START         As String = "09:20:00"
Private Const LATE_ENTRY_CUTOFF         As String = "15:00:00"
Private Const FORCE_LIQUIDATION         As String = "15:05:00"
Private Const DATA_START_ROW            As Long = 5

' =====================================================================
'  ENTRY FILTER THRESHOLDS
' =====================================================================
' --- Donchian / V2 gates ---
Private Const GATE_RVOL_MIN             As Double = 1.25     ' Hard RVOL gate for Donchian breakouts
Private Const DONCHIAN_N                As Long = 5          ' 5-bar channel lookback

' --- V3 gates ---
Private Const GATE_RSI_BUY_MAX          As Double = 75#      ' RSI cap for BUY (V3)
Private Const GATE_RSI_SELL_MIN         As Double = 25#      ' RSI floor for SELL (V3)
Private Const GATE_MAX_CLIMAX_ATR       As Double = 1.5      ' |LTP - EMA30| <= 1.5 ATR (V3)
Private Const VELOCITY_LOOKBACK_K       As Long = 3          ' Pre-cross velocity lookback (V3)

' =====================================================================
'  RISK MANAGEMENT
' =====================================================================
Private Const GATE_MAX_RISK_ATR         As Double = 1#       ' SL capped at 1.0 ATR
Private Const RISK_FLOOR_MIN_PCT        As Double = 0.0035   ' 0.35% minimum stop risk floor
Private Const TARGET_RR_MULTIPLE        As Double = 2#       ' Target = 2.0 x initial_risk

' =====================================================================
'  POSITION SIZING
' =====================================================================
Private Const SIZE_DEFAULT              As Double = 0.5

' =====================================================================
'  EXIT MANAGEMENT (V5 Common Lifecycle)
' =====================================================================
Private Const MAX_CANDLES_EXIT          As Long = 9          ' 9-candle expiry (V5 standard)
Private Const TRAILING_BE_ATR_MULT      As Double = 1#       ' Breakeven threshold in ATR units
Private Const PROFIT_RETRACE_FRAC       As Double = 0.4      ' 40% giveback from peak MFE
Private Const PROFIT_RETRACE_MIN_PCT    As Double = 0.003    ' Minimum 0.3% giveback for exit
Private Const PROFIT_ARM_MIN_PCT        As Double = 0.006    ' Arm profit protection after +0.6% MFE
Private Const PROFIT_PROTECT_TGT_FRAC   As Double = 0.5      ' 50% of target distance
Private Const STAGNATION_BARS           As Long = 4          ' Dead trade after 4 bars
Private Const STAGNATION_MIN_MFE_PCT    As Double = 0.003    ' < 0.3% MFE = stagnant
Private Const DEAD_TRADE_EXIT_MAX_PCT   As Double = 0.001    ' Dead trade threshold
Private Const COOLDOWN_MINUTES_VAL      As Long = 30
Private Const MAX_ATTEMPTS_PER_TICKER   As Long = 2

' =====================================================================
'  COCKPIT_V5 COLUMN LETTERS (mirrors V4 layout exactly)
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
Private Const CK_CANDLE_COUNT As String = "AB"     ' bars held
Private Const CK_STRATEGY     As String = "AC"
Private Const CK_PEAK_PRICE   As String = "AD"     ' peak favorable price
Private Const CK_BEST_MFE_PCT As String = "AE"     ' peak favorable %
Private Const CK_LAST_MFE_BAR As String = "AF"     ' bar index of peak MFE
Private Const CK_MFE_BAR1     As String = "AG"     ' MFE at Bar 1
Private Const CK_IS_RUNNER    As String = "AH"     ' 1 = confirmed runner
Private Const CK_LAST_BAR_TS  As String = "AI"     ' last completed 5m bar timestamp (for bar counting)

' Module-level caches
Private mRowLookup As Object
Private mActiveTrades As Object
Private mCooldowns As Object
Private mTickerAttempts As Object


' =====================================================================
' Public Sub  Run_Cockpit_V5
' Main Entry Point for COCKPIT_V5 Research Scaffold
' =====================================================================
Public Sub Run_Cockpit_V5()
    Dim currentTime As Date
    currentTime = TimeValue(Now)

    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Dim wsSrc  As Worksheet: Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Dim wsC5   As Worksheet: Set wsC5 = GetOrCreateCockpitV5(ThisWorkbook)
    Dim wsAudit As Worksheet: Set wsAudit = GetOrCreateV5Funnel(ThisWorkbook)
    Dim auditRun As Date: auditRun = Now

    EnsureCockpitV5Headers wsC5
    On Error Resume Next
    Update_K3_Dashboard wsSrc, wsC5
    On Error GoTo ErrorHandler
    Set mRowLookup = BuildRowLookup(wsSrc)
    Set mActiveTrades = BuildActiveTradeLookup(wsC5)
    Set mTickerAttempts = BuildTickerSessionAttemptLookup(wsC5)

    ' -- TIME GATE: Outside trading window, only manage existing open trades --
    If currentTime < TimeValue(EARLY_ENTRY_START) Or currentTime >= TimeValue(LATE_ENTRY_CUTOFF) Then
        WriteV5FunnelRow wsAudit, auditRun, "*RUN*", "", "", "", "N/A", "N/A", "N/A", _
                         "NOT_EVALUATED", "SCAN_SKIPPED_TIME_GATE"
        ManageCockpitV5Trades wsC5, wsSrc
        GoTo Cleanup
    End If

    ' -- Disabled or explicitly scaffolded mode: manage open trades only --
    If Not ENTRY_RULES_VALIDATED Or UCase$(Trim$(COMBINATION_MODE)) = "SCAFFOLD" Then
        WriteV5FunnelRow wsAudit, auditRun, "*RUN*", "", "", "", "N/A", "N/A", "N/A", _
                         "NOT_EVALUATED", "SCAN_SKIPPED_MODE_DISABLED"
        ManageCockpitV5Trades wsC5, wsSrc
        GoTo Cleanup
    End If

    ' -- Scan Sheet1 via in-memory 2D array --
    Dim srcLastRow As Long
    srcLastRow = wsSrc.cells(wsSrc.Rows.count, 1).End(xlUp).row
    If srcLastRow < 4 Then GoTo Cleanup

    Dim srcData As Variant
    ' Range extends to DY (col 129) to cover DM telemetry packet (col 117, for V3 velocity)
    ' and BTB shadow features (DV-DY, cols 126-129). P2 fix for future gate porting.
    srcData = wsSrc.Range("A1:DY" & srcLastRow).Value2

    Dim lastDestRow As Long
    lastDestRow = wsC5.cells(wsC5.Rows.count, 1).End(xlUp).row
    If lastDestRow < DATA_START_ROW Then lastDestRow = DATA_START_ROW - 1

    Dim r As Long, destRow As Long
    destRow = lastDestRow

    For r = 4 To srcLastRow
        Dim auditCross As String, auditBreakout As String, auditEmaV5 As String
        Dim auditEmaV3 As String, auditDonchian As String, auditCombination As String
        Dim auditDecision As String, auditDirection As String
        Dim reasonEmaV5 As String, reasonEmaV3 As String, reasonDonchian As String
        auditCross = "": auditBreakout = "": auditEmaV5 = "N/A"
        auditEmaV3 = "N/A": auditDonchian = "N/A": auditCombination = "NOT_EVALUATED"
        auditDecision = "": auditDirection = "NEUTRAL"
        Dim ticker As String: ticker = Trim$(CStr(srcData(r, 1)))
        If Len(ticker) = 0 Or Is_Index(ticker) Then GoTo NextStock

        ' === FETCH CORE TECHNICALS FROM MEMORY ARRAY ===
        Dim ltp As Double:         ltp = CDbl(val(srcData(r, 20)))         ' Col T
        Dim ema10 As Double:       ema10 = CDbl(val(srcData(r, 45)))       ' Col AS
        Dim ema30 As Double:       ema30 = CDbl(val(srcData(r, 46)))       ' Col AT
        Dim atr As Double:         atr = CDbl(val(srcData(r, 81)))         ' Col CC
        Dim rsi5m As Double:       rsi5m = CDbl(val(srcData(r, 21)))       ' Col U
        Dim stockRegime As String: stockRegime = UCase$(Trim$(CStr(srcData(r, 66)))) ' Col BN
        Dim stockGrade As String:  stockGrade = UCase$(Trim$(CStr(srcData(r, 109)))) ' Col DE
        Dim rvol As Double:        rvol = CDbl(val(srcData(r, 67)))        ' Col BO
        Dim breakoutFlag As Long:  breakoutFlag = CLng(val(srcData(r, 110))) ' Col DF
        Dim crossStr As String:    crossStr = UCase$(Trim$(CStr(srcData(r, 29)))) ' Col AC

        auditCross = crossStr
        auditBreakout = CStr(breakoutFlag)
        If ltp <= 0 Or atr <= 0 Then
            auditDecision = "REJECT_INVALID_LTP_OR_ATR"
            GoTo NextStock
        End If

        ' === EVALUATE INDEPENDENT SIGNAL FAMILIES ===
        Dim emaDirection As String:      emaDirection = "NEUTRAL"
        Dim donchianDirection As String: donchianDirection = "NEUTRAL"
        Dim emaV5Pass As Boolean:        emaV5Pass = False
        Dim emaV3Pass As Boolean:        emaV3Pass = False
        Dim donchianPass As Boolean:     donchianPass = False

        ' --- Family 1: EMA Cross ---
        If crossStr = "BULLISH" Then
            emaDirection = "BUY"
        ElseIf crossStr = "BEARISH" Then
            emaDirection = "SELL"
        End If

        ' --- Family 2: Donchian Breakout ---
        If breakoutFlag = 1 Then
            donchianDirection = "BUY"
        ElseIf breakoutFlag = -1 Then
            donchianDirection = "SELL"
        End If

        ' Skip if no signal from either family
        auditDirection = emaDirection & "/" & donchianDirection
        If emaDirection = "NEUTRAL" And donchianDirection = "NEUTRAL" Then
            auditDecision = "REJECT_NO_TRIGGER"
            GoTo NextStock
        End If

        Dim secRegime As String: secRegime = GetSectorRegime(wsSrc, ticker)

        ' Resolve trade direction (must be unambiguous)
        Dim isBuy As Boolean
        Dim tradeDir As String
        If emaDirection <> "NEUTRAL" And donchianDirection <> "NEUTRAL" Then
            ' Both families fired: must agree on direction
            If emaDirection <> donchianDirection Then
                auditDecision = "REJECT_DIRECTION_CONFLICT"
                GoTo NextStock
            End If
            isBuy = (emaDirection = "BUY")
        ElseIf emaDirection <> "NEUTRAL" Then
            isBuy = (emaDirection = "BUY")
        Else
            isBuy = (donchianDirection = "BUY")
        End If
        tradeDir = IIf(isBuy, "BUY", "SELL")

        ' The Donchian-confirmed mode uses EMA alignment/quality as confirmation;
        ' requiring a fresh cross event here would make the two event triggers
        ' coincide on the same scan and reject Donchian-only candidates by design.
        Dim evaluateEmaQuality As Boolean
        evaluateEmaQuality = (emaDirection <> "NEUTRAL")
        If UCase$(Trim$(COMBINATION_MODE)) = "DONCHIAN_CONFIRMED" And _
           donchianDirection <> "NEUTRAL" Then evaluateEmaQuality = True

        ' === EVALUATE EMA QUALITY BRANCH ===
        If evaluateEmaQuality Then
            reasonEmaV5 = ""
            emaV5Pass = PassEmaV5Gates(isBuy, ltp, ema10, ema30, atr, rsi5m, _
                                       stockRegime, stockGrade, secRegime, reasonEmaV5)
            auditEmaV5 = IIf(emaV5Pass, "PASS", "FAIL:" & reasonEmaV5)
        End If

        ' === EVALUATE EMA V3 GATES ===
        Dim v3Velocity As Double: v3Velocity = 0#
        If evaluateEmaQuality Then
            reasonEmaV3 = ""
            emaV3Pass = PassV3Gates(srcData, r, isBuy, ltp, ema10, ema30, atr, rsi5m, v3Velocity, reasonEmaV3)
            auditEmaV3 = IIf(emaV3Pass, "PASS", "FAIL:" & reasonEmaV3)
        End If

        ' === EVALUATE DONCHIAN V2 GATES ===
        If donchianDirection <> "NEUTRAL" Then
            reasonDonchian = ""
            donchianPass = PassDonchianGates(isBuy, rvol, ema10, ema30, stockRegime, _
                                             stockGrade, secRegime, reasonDonchian)
            auditDonchian = IIf(donchianPass, "PASS", "FAIL:" & reasonDonchian)
        End If

        ' === APPLY COMBINATION RULE ===
        Dim admitted As Boolean: admitted = False
        Dim stratName As String: stratName = ""
        Dim discovTag As String: discovTag = ""

        Select Case UCase$(Trim$(COMBINATION_MODE))
            Case "DONCHIAN_CONFIRMED"
                ' Core confirmation: Donchian + at least one EMA gate variant
                If Not donchianPass Then
                    auditCombination = "FAIL_DONCHIAN_GATES"
                ElseIf Not (emaV5Pass Or emaV3Pass) Then
                    auditCombination = "FAIL_EMA_CONFIRMATION"
                Else
                    admitted = True
                    stratName = "V5_Confirmed"
                    discovTag = "D+" & IIf(emaV5Pass And emaV3Pass, "EV5+EV3", IIf(emaV5Pass, "EV5", "EV3"))
                End If
            Case "UNION"
                ' Research union: any signal passes
                If emaV5Pass Or emaV3Pass Or donchianPass Then
                    admitted = True
                    If donchianPass And (emaV5Pass Or emaV3Pass) Then
                        stratName = "V5_Dual": discovTag = "Union_Dual"
                    ElseIf donchianPass Then
                        stratName = "V5_Donchian": discovTag = "Union_D"
                    Else
                        stratName = "V5_EMA": discovTag = "Union_EMA"
                    End If
                End If
            Case "DONCHIAN_ONLY"
                If donchianPass Then
                    admitted = True: stratName = "V5_DOnly": discovTag = "Donchian_Standalone"
                End If
            Case "EMA_V5_ONLY"
                If emaV5Pass Then
                    admitted = True: stratName = "V5_EV5Only": discovTag = "EMA_V5_Standalone"
                End If
            Case "EMA_V3_ONLY"
                If emaV3Pass Then
                    admitted = True: stratName = "V5_EV3Only": discovTag = "EMA_V3_Standalone"
                End If
            Case Else
                ' SCAFFOLD or unknown: no entries
                admitted = False
        End Select

        If admitted Then
            auditCombination = "PASS"
        ElseIf auditCombination = "NOT_EVALUATED" Then
            auditCombination = "FAIL_" & UCase$(Trim$(COMBINATION_MODE))
        End If
        If Not admitted Then
            auditDecision = "REJECT_COMBINATION"
            GoTo NextStock
        End If

        ' === CHURN CHECKS ===
        If mActiveTrades.Exists(ticker & "|ANY") Then
            auditDecision = "REJECT_ACTIVE_TRADE"
            GoTo NextStock
        End If

        Dim sessionAttempts As Long: sessionAttempts = 0
        If Not mTickerAttempts Is Nothing Then
            If mTickerAttempts.Exists(ticker) Then sessionAttempts = CLng(mTickerAttempts(ticker))
        End If
        If sessionAttempts >= MAX_ATTEMPTS_PER_TICKER Then
            auditDecision = "REJECT_ATTEMPT_LIMIT"
            GoTo NextStock
        End If
        If IsOnCooldown(ticker) Then
            auditDecision = "REJECT_COOLDOWN"
            GoTo NextStock
        End If

        ' === CALCULATE RISK & WRITE TRADE ===
        Dim swingLow As Double:  swingLow = CDbl(val(srcData(r, 94)))   ' Col CP
        Dim swingHigh As Double: swingHigh = CDbl(val(srcData(r, 93)))  ' Col CO
        Dim slPx As Double, tgtPx As Double, initialRisk As Double

        If isBuy Then
            If swingLow > 0 And swingLow < ltp Then slPx = swingLow Else slPx = ltp - (atr * GATE_MAX_RISK_ATR)
            If (ltp - slPx) > (atr * GATE_MAX_RISK_ATR) Then slPx = ltp - (atr * GATE_MAX_RISK_ATR)
            initialRisk = ltp - slPx
        Else
            If swingHigh > 0 And swingHigh > ltp Then slPx = swingHigh Else slPx = ltp + (atr * GATE_MAX_RISK_ATR)
            If (slPx - ltp) > (atr * GATE_MAX_RISK_ATR) Then slPx = ltp + (atr * GATE_MAX_RISK_ATR)
            initialRisk = slPx - ltp
        End If

        Dim minRisk As Double: minRisk = ltp * RISK_FLOOR_MIN_PCT
        Dim finalSize As Double: finalSize = SIZE_DEFAULT

        If initialRisk > 0 And initialRisk < minRisk And minRisk > 0 Then
            finalSize = Round(SIZE_DEFAULT * (initialRisk / minRisk), 4)
            initialRisk = minRisk
            If isBuy Then slPx = ltp - initialRisk Else slPx = ltp + initialRisk
        End If
        If initialRisk <= 0 Then initialRisk = ltp * RISK_FLOOR_MIN_PCT

        If isBuy Then tgtPx = ltp + (initialRisk * TARGET_RR_MULTIPLE) Else tgtPx = ltp - (initialRisk * TARGET_RR_MULTIPLE)

        destRow = destRow + 1
        WriteNewCockpitV5Row wsC5, wsSrc, r, destRow, ticker, tradeDir, ltp, slPx, tgtPx, _
                             rvol, rsi5m, stockRegime, stockGrade, secRegime, _
                             initialRisk, finalSize, stratName, discovTag, _
                             emaV5Pass, emaV3Pass, donchianPass
        auditDecision = "SIGNAL_WRITTEN_ROW_" & CStr(destRow)

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
        If Len(ticker) > 0 And Not Is_Index(ticker) Then
            WriteV5FunnelRow wsAudit, auditRun, ticker, auditCross, auditBreakout, _
                             auditDirection, auditEmaV5, auditEmaV3, auditDonchian, _
                             auditCombination, auditDecision
        End If
    Next r

    ' Manage existing open positions
    ManageCockpitV5Trades wsC5, wsSrc

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
    MsgBox "Error in Run_Cockpit_V5: " & Err.Description, vbCritical
End Sub


' =====================================================================
' Private Function  PassEmaV5Gates
' EMA-cross quality gates used by the confirmation branch.
' =====================================================================
' Public market-signal entry mode uses this branch together with the V2 breakout
' gates. The source columns match COCKPIT_V4's production quality filters.
Private Function PassEmaV5Gates(isBuy As Boolean, ltp As Double, _
        ema10 As Double, ema30 As Double, atr As Double, rsi5m As Double, _
        stockRegime As String, stockGrade As String, secRegime As String, _
        ByRef failReason As String) As Boolean

    PassEmaV5Gates = False
    If ltp <= 0 Or atr <= 0 Then failReason = "BAD_LTP_ATR": Exit Function
    If Not IsKnownRegime(stockRegime) Then failReason = "UNKNOWN_STOCK_REGIME": Exit Function
    If Not IsKnownRegime(secRegime) Then failReason = "UNKNOWN_SECTOR_REGIME": Exit Function

    If isBuy Then
        If ema10 <= ema30 Then failReason = "EMA_NOT_BULL_ALIGNED": Exit Function
        If Not (stockRegime = "TRENDING_BULL" Or stockRegime = "MILD_UP" Or stockRegime = "RANGING") Then failReason = "STOCK_REGIME": Exit Function
        If secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN" Then failReason = "SECTOR_CONTRA": Exit Function
        If rsi5m > GATE_RSI_BUY_MAX Then failReason = "RSI_HIGH": Exit Function
    Else
        If ema10 >= ema30 Then failReason = "EMA_NOT_BEAR_ALIGNED": Exit Function
        If Not (stockRegime = "TRENDING_BEAR" Or stockRegime = "MILD_DOWN") Then failReason = "STOCK_REGIME": Exit Function
        If Not (secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN" Or secRegime = "RANGING") Then failReason = "SECTOR_CONTRA": Exit Function
        If rsi5m < GATE_RSI_SELL_MIN Then failReason = "RSI_LOW": Exit Function
    End If

    If Not (stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B") Then failReason = "GRADE": Exit Function
    If Abs(ltp - ema30) > (GATE_MAX_CLIMAX_ATR * atr) Then failReason = "CLIMAX": Exit Function
    failReason = ""
    PassEmaV5Gates = True
End Function


' =====================================================================
' Private Function  PassV3Gates
' V3: EMA alignment + VWAP + RSI + climax + DM telemetry velocity.
' =====================================================================
Private Function PassV3Gates(srcData As Variant, srcRow As Long, _
        isBuy As Boolean, ltp As Double, ema10 As Double, ema30 As Double, _
        atr As Double, rsi5m As Double, ByRef vAvg As Double, _
        ByRef failReason As String) As Boolean

    PassV3Gates = False
    vAvg = 0#
    If ltp <= 0 Or atr <= 0 Then failReason = "BAD_LTP_ATR": Exit Function

    If isBuy Then
        If ema10 <= ema30 Then failReason = "EMA_NOT_BULL_ALIGNED": Exit Function
        If rsi5m > GATE_RSI_BUY_MAX Then failReason = "RSI_HIGH": Exit Function
    Else
        If ema10 >= ema30 Then failReason = "EMA_NOT_BEAR_ALIGNED": Exit Function
        If rsi5m < GATE_RSI_SELL_MIN Then failReason = "RSI_LOW": Exit Function
    End If

    If Abs(ltp - ema30) > (GATE_MAX_CLIMAX_ATR * atr) Then failReason = "CLIMAX": Exit Function

    Dim vwap As Double: vwap = CDbl(val(srcData(srcRow, 70)))
    If vwap <= 0 Then failReason = "VWAP_MISSING": Exit Function
    If isBuy And ltp <= vwap Then failReason = "BELOW_VWAP": Exit Function
    If Not isBuy And ltp >= vwap Then failReason = "ABOVE_VWAP": Exit Function

    Dim packet As String: packet = Trim$(CStr(srcData(srcRow, 117)))
    If Len(packet) = 0 Then failReason = "DM_PACKET_MISSING": Exit Function
    Dim fields() As String: fields = Split(packet, "|")
    If UBound(fields) < 5 Then failReason = "DM_PACKET_SHORT": Exit Function

    Dim g0 As Double, g1 As Double, g2 As Double, g3 As Double, cATR As Double, kVal As Double
    If Not ParseTelemetryNumber(fields(0), g0) Then failReason = "DM_G0_INVALID": Exit Function
    If Not ParseTelemetryNumber(fields(1), g1) Then failReason = "DM_G1_INVALID": Exit Function
    If Not ParseTelemetryNumber(fields(2), g2) Then failReason = "DM_G2_INVALID": Exit Function
    If Not ParseTelemetryNumber(fields(3), g3) Then failReason = "DM_G3_INVALID": Exit Function
    If Not ParseTelemetryNumber(fields(4), cATR) Then failReason = "DM_ATR_INVALID": Exit Function
    If Not ParseTelemetryNumber(fields(5), kVal) Then failReason = "DM_K_INVALID": Exit Function
    If cATR <= 0 Or kVal < 1 Or kVal > VELOCITY_LOOKBACK_K Then failReason = "DM_ATR_OR_K_RANGE": Exit Function

    Dim gk As Double
    Select Case CLng(kVal)
        Case 1: gk = g1
        Case 2: gk = g2
        Case 3: gk = g3
        Case Else: failReason = "DM_K_UNSUPPORTED": Exit Function
    End Select

    If isBuy Then
        vAvg = (g0 - gk) / (CLng(kVal) * cATR)
    Else
        vAvg = (gk - g0) / (CLng(kVal) * cATR)
    End If
    If vAvg <= 0 Then failReason = "VELOCITY_NONPOSITIVE": Exit Function

    failReason = ""
    PassV3Gates = True
End Function


Private Function ParseTelemetryNumber(ByVal item As String, ByRef value As Double) As Boolean
    ParseTelemetryNumber = False
    Dim kv() As String: kv = Split(item, "=")
    If UBound(kv) < 1 Then Exit Function
    If Not IsNumeric(Trim$(kv(1))) Then Exit Function
    value = CDbl(Trim$(kv(1)))
    ParseTelemetryNumber = True
End Function


' =====================================================================
' Private Function  PassDonchianGates
' V2 gate chain: RVOL >= 1.25 + EMA alignment + regime + grade
' Replicates COCKPIT_V2.txt lines 404-451
' =====================================================================
Private Function PassDonchianGates(isBuy As Boolean, rvol As Double, _
        ema10 As Double, ema30 As Double, stockRegime As String, _
        stockGrade As String, secRegime As String, ByRef failReason As String) As Boolean

    PassDonchianGates = False
    If Not IsKnownRegime(stockRegime) Then failReason = "UNKNOWN_STOCK_REGIME": Exit Function
    If Not IsKnownRegime(secRegime) Then failReason = "UNKNOWN_SECTOR_REGIME": Exit Function

    ' 1. RVOL hard gate (V2's defining feature)
    If rvol < GATE_RVOL_MIN Then failReason = "RVOL_LOW": Exit Function

    ' 2. EMA trend alignment
    If isBuy Then
        If ema10 <= ema30 Then failReason = "EMA_NOT_BULL_ALIGNED": Exit Function
    Else
        If ema10 >= ema30 Then failReason = "EMA_NOT_BEAR_ALIGNED": Exit Function
    End If

    ' 3. Directional stock regime
    If isBuy Then
        If Not (stockRegime = "TRENDING_BULL" Or stockRegime = "MILD_UP" Or stockRegime = "RANGING") Then failReason = "STOCK_REGIME": Exit Function
    Else
        If Not (stockRegime = "TRENDING_BEAR" Or stockRegime = "MILD_DOWN") Then failReason = "STOCK_REGIME": Exit Function
    End If

    ' 4. Sector alignment
    If isBuy Then
        If secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN" Then failReason = "SECTOR_CONTRA": Exit Function
    Else
        If secRegime = "TRENDING_BULL" Or secRegime = "MILD_UP" Then failReason = "SECTOR_CONTRA": Exit Function
    End If

    ' 5. Grade filter used by the registered V2 study definition
    If Not (stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B") Then failReason = "GRADE": Exit Function

    failReason = ""
    PassDonchianGates = True
End Function


Private Function IsKnownRegime(ByVal regime As String) As Boolean
    Select Case UCase$(Trim$(regime))
        Case "TRENDING_BULL", "MILD_UP", "RANGING", "TRENDING_BEAR", "MILD_DOWN"
            IsKnownRegime = True
        Case Else
            IsKnownRegime = False
    End Select
End Function


' =====================================================================
' Private Sub  WriteNewCockpitV5Row
' Writes an admitted trade row into COCKPIT_V5 matching V4 formatting
' =====================================================================
Private Sub WriteNewCockpitV5Row(wsC As Worksheet, wsSrc As Worksheet, _
        srcRow As Long, destRow As Long, ticker As String, action As String, _
        ltp As Double, slPx As Double, tgtPx As Double, _
        rvol As Double, rsi5m As Double, _
        stockRegime As String, stockGrade As String, secRegime As String, _
        initialRisk As Double, finalSize As Double, _
        ByVal stratName As String, ByVal discovTag As String, _
        ByVal ev5Pass As Boolean, ByVal ev3Pass As Boolean, ByVal dPass As Boolean)

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

    ' Outcome & Bottomline formulas (P1 fix: without these the sheet shows no P&L)
    ' Outcome = live unrealized % (updates with CP). Bottomline = realized % (freezes at close).
'    wsC.cells(destRow, RangeRef(CK_OUTCOME)).Formula = _
'        "=IF(ISBLANK(V" & destRow & "),"""",IF(B" & destRow & "=""BUY"",(V" & destRow & "-S" & destRow & ")/S" & destRow & ",(S" & destRow & "-V" & destRow & ")/S" & destRow & "))"
'    wsC.cells(destRow, RangeRef(CK_BOTTOMLINE)).Formula = _
'        "=IF(ISBLANK(X" & destRow & "),"""",IF(B" & destRow & "=""BUY"",(V" & destRow & "-S" & destRow & ")/S" & destRow & ",(S" & destRow & "-V" & destRow & ")/S" & destRow & "))"

    ' Signal branch audit trail in comments
    Dim sigTag As String: sigTag = ""
    If ev5Pass Then sigTag = sigTag & "EV5 "
    If ev3Pass Then sigTag = sigTag & "EV3 "
    If dPass Then sigTag = sigTag & "D "
    sigTag = Trim$(sigTag)

    wsC.cells(destRow, RangeRef(CK_COMENTS)).value = _
        "V5_ENTRY | " & discovTag & " | Signals=[" & sigTag & "]" & _
        " | Mode=" & COMBINATION_MODE & _
        " | Reg=" & stockRegime & " Sec=" & secRegime & " Grd=" & stockGrade & _
        " | RVOL=" & Format(rvol, "0.00") & " RSI=" & Format(rsi5m, "0.0") & _
        " | Risk=" & Format(initialRisk, "0.00") & " Size=" & Format(finalSize, "0.00")

    wsC.cells(destRow, RangeRef(CK_CANDLE_COUNT)).value = 0
    wsC.cells(destRow, RangeRef(CK_STRATEGY)).value = stratName
    wsC.cells(destRow, RangeRef(CK_PEAK_PRICE)).value = ltp
    wsC.cells(destRow, RangeRef(CK_BEST_MFE_PCT)).value = 0#
    wsC.cells(destRow, RangeRef(CK_LAST_MFE_BAR)).value = 0
    wsC.cells(destRow, RangeRef(CK_MFE_BAR1)).value = 0#
    wsC.cells(destRow, RangeRef(CK_IS_RUNNER)).value = 0
    Dim entryBarStamp As Date
    entryBarStamp = DateValue(Now) + TimeSerial(Hour(Now), (Minute(Now) \ 5) * 5, 0)
    wsC.cells(destRow, RangeRef(CK_LAST_BAR_TS)).value = CDbl(entryBarStamp)
    ApplyCockpitColors wsC, destRow, (action = "BUY")
End Sub


' =====================================================================
' Private Sub  ManageCockpitV5Trades
' V5 Common Lifecycle: SL, TGT, Trailing BE, Profit Reversal,
' Dead Trade, 9-Candle Expiry, 15:05 Session Liquidation
' + Non-exiting telemetry tags (kinetic, EMA30 breach, retrace, stagnation)
'
' NOTE: VBA liquidates at 15:05 (live market safety margin to ensure
' order execution before exchange close). Python replay uses 15:20
' (historical backtest sees full session bars). This is an intentional
' difference — the VBA shadow observation is NOT a parity check against
' the Python study. Compare results qualitatively, not numerically.
' =====================================================================
Private Sub ManageCockpitV5Trades(wsC As Worksheet, wsSrc As Worksheet)
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

        ' Count each newly elapsed dated 5-minute clock bucket once.
        ' Sheet1 exposes no source-bar timestamp, so this is not feed-timestamp synchronized.
        Dim curBarTS As Date
        curBarTS = DateValue(Now) + TimeSerial(Hour(Now), (Minute(Now) \ 5) * 5, 0)
        Dim lastBarVal As Variant: lastBarVal = wsC.Range(CK_LAST_BAR_TS & r).value
        Dim isNewBar As Boolean: isNewBar = False

        If IsEmpty(lastBarVal) Or Len(CStr(lastBarVal)) = 0 Then
            ' First management cycle for this trade
            isNewBar = True
        Else
            If IsDate(lastBarVal) Then
                If CDbl(CDate(lastBarVal)) < CDbl(curBarTS) Then isNewBar = True
            Else
                isNewBar = True
            End If
        End If

        If isNewBar Then
            candleCount = candleCount + 1
            wsC.Range(CK_CANDLE_COUNT & r).value = candleCount
            wsC.Range(CK_LAST_BAR_TS & r).value = CDbl(curBarTS)
        End If

        ' Fetch current EMA30 for telemetry
        Dim curEMA30 As Double: curEMA30 = GetEMA30(wsSrc, ticker)

        ' Compute current MFE %
        Dim curMfePct As Double
        curMfePct = IIf(action = "BUY", (curLTP - entryPx) / entryPx, (entryPx - curLTP) / entryPx)

        ' Record Bar 1 MFE
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

        ' --- EXIT EVALUATION (priority-ordered V5 common lifecycle) ---
        Dim exitFlag As Boolean: exitFlag = False
        Dim exitReason As String: exitReason = ""

        ' 1. Mandatory Session Liquidation (15:05:00)
        If forceLiquidate Then
            exitFlag = True
            exitReason = "SESSION_CLOSE (15:05 Hard Exit)"

        ' 2. Stop Loss Hit
        ElseIf action = "BUY" And curLTP <= slPx Then
            exitFlag = True
            exitReason = IIf(slPx >= entryPx * 0.9999, "STOP_LOSS (Break-Even Hit)", "STOP_LOSS (Initial SL Hit)")
        ElseIf action = "SELL" And curLTP >= slPx Then
            exitFlag = True
            exitReason = IIf(slPx <= entryPx * 1.0001, "STOP_LOSS (Break-Even Hit)", "STOP_LOSS (Initial SL Hit)")

        ' 3. Target Hit
        ElseIf action = "BUY" And tgtPx > 0 And curLTP >= tgtPx Then
            exitFlag = True
            exitReason = "TARGET_HIT"
        ElseIf action = "SELL" And tgtPx > 0 And curLTP <= tgtPx Then
            exitFlag = True
            exitReason = "TARGET_HIT"
        End If

        ' 4. Trailing Breakeven Update (non-exiting, modifies SL)
        If Not exitFlag Then
            Dim curATR As Double: curATR = GetATR5m(wsSrc, ticker)
            Dim beThresh As Double
            If curATR > 0 And entryPx > 0 Then beThresh = (curATR * TRAILING_BE_ATR_MULT) / entryPx Else beThresh = 0.002
            If bestMfe >= beThresh Then
                If action = "BUY" Then
                    Dim bePxBuy As Double: bePxBuy = entryPx * 1.0005
                    If slPx < bePxBuy Then
                        wsC.Range(CK_SL & r).value = bePxBuy
                        slPx = bePxBuy
                    End If
                Else
                    Dim bePxSell As Double: bePxSell = entryPx * 0.9995
                    If slPx > bePxSell Then
                        wsC.Range(CK_SL & r).value = bePxSell
                        slPx = bePxSell
                    End If
                End If
            End If
        End If

        ' 4b. Immediate re-check: if BE update moved stop past current LTP, exit now (P2 fix)
        If Not exitFlag Then
            If action = "BUY" And curLTP <= slPx Then
                exitFlag = True
                exitReason = "STOP_LOSS (Break-Even Hit — immediate)"
            ElseIf action = "SELL" And curLTP >= slPx Then
                exitFlag = True
                exitReason = "STOP_LOSS (Break-Even Hit — immediate)"
            End If
        End If

        ' 5. Profit Reversal Exit
        If Not exitFlag Then
            Dim tgtDistPct As Double: tgtDistPct = Abs(tgtPx - entryPx) / entryPx
            Dim profitArm As Double: profitArm = PROFIT_ARM_MIN_PCT
            If tgtDistPct * PROFIT_PROTECT_TGT_FRAC > profitArm Then profitArm = tgtDistPct * PROFIT_PROTECT_TGT_FRAC

            If bestMfe >= profitArm Then
                Dim giveback As Double
                If action = "BUY" Then
                    giveback = bestMfe - ((curLTP - entryPx) / entryPx)
                Else
                    giveback = bestMfe - ((entryPx - curLTP) / entryPx)
                End If
                Dim retraceTrig As Double: retraceTrig = bestMfe * PROFIT_RETRACE_FRAC
                If retraceTrig < PROFIT_RETRACE_MIN_PCT Then retraceTrig = PROFIT_RETRACE_MIN_PCT
                If giveback >= retraceTrig Then
                    exitFlag = True
                    exitReason = "PROFIT_REVERSAL (Peak=" & Format(bestMfe * 100, "0.00") & "%)"
                End If
            End If
        End If

        ' 6. Dead Trade Cut & 9-Candle Expiry
        If Not exitFlag And candleCount >= STAGNATION_BARS And bestMfe < STAGNATION_MIN_MFE_PCT Then
            exitFlag = True
            exitReason = "DEAD_TRADE (" & candleCount & " bars, MFE=" & Format(bestMfe * 100, "0.00") & "%)"
        End If
        If Not exitFlag And candleCount >= MAX_CANDLES_EXIT Then
            exitFlag = True
            exitReason = "EXPIRED (" & MAX_CANDLES_EXIT & " candles)"
        End If

        ' --- V5 TELEMETRY (Non-Exiting Diagnostics) ---
        If Not exitFlag Then
            Dim curComment As String: curComment = CStr(wsC.Range(CK_COMENTS & r).value)

            ' T1. Kinetic Follow-Through (Bar 3-4)
            If (candleCount = 3 Or candleCount = 4) And isRunner = 0 Then
                If InStr(curComment, "KINETIC_") = 0 Then
                    If curMfePct <= mfeBar1 Or curMfePct <= 0 Then
                        curComment = curComment & " | KINETIC_STAGNANT"
                    Else
                        isRunner = 1
                        wsC.Range(CK_IS_RUNNER & r).value = 1
                        curComment = curComment & " | KINETIC_RUNNER"
                    End If
                    wsC.Range(CK_COMENTS & r).value = curComment
                End If
            End If

            ' T2. EMA30 Breach
            If (action = "BUY" And curLTP < curEMA30) Or (action = "SELL" And curLTP > curEMA30) Then
                If InStr(curComment, "EMA30_BREACH") = 0 Then
                    curComment = curComment & " | EMA30_BREACH"
                    wsC.Range(CK_COMENTS & r).value = curComment
                End If
            End If

            ' T3. Opposite EMA Crossover
            Dim oppCross As String: oppCross = UCase$(Trim$(CStr(GetCross1030(wsSrc, ticker))))
            If (action = "BUY" And oppCross = "BEARISH") Or (action = "SELL" And oppCross = "BULLISH") Then
                If InStr(curComment, "OPP_CROSS") = 0 Then
                    curComment = curComment & " | OPP_CROSS"
                    wsC.Range(CK_COMENTS & r).value = curComment
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

Private Function GetOrCreateCockpitV5(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("COCKPIT_V5")
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Dim srcSheet As Worksheet
        Set srcSheet = wb.Sheets("COCKPIT_V4")
        If srcSheet Is Nothing Then Set srcSheet = wb.Sheets("COCKPIT")
        On Error GoTo 0

        If Not srcSheet Is Nothing Then
            srcSheet.Copy After:=srcSheet
            Set ws = wb.ActiveSheet
            ws.Name = "COCKPIT_V5"

            Dim lr As Long: lr = ws.cells(ws.Rows.count, 1).End(xlUp).row
            If lr >= DATA_START_ROW Then
                ws.Range("A" & DATA_START_ROW & ":AI" & lr).ClearContents
                ws.Range("A" & DATA_START_ROW & ":AI" & lr).Interior.ColorIndex = xlNone
            End If
        Else
            Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
            ws.Name = "COCKPIT_V5"
        End If
    End If
    Set GetOrCreateCockpitV5 = ws
End Function

Private Function GetOrCreateV5Funnel(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("V5_Funnel")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
        ws.Name = "V5_Funnel"
    End If
    
    If Len(CStr(ws.cells(1, 1).value)) = 0 Then
        Dim headers(1 To 1, 1 To 10) As Variant
        headers(1, 1) = "Run_Timestamp": headers(1, 2) = "Ticker"
        headers(1, 3) = "EMA_Cross": headers(1, 4) = "Donchian_Flag"
        headers(1, 5) = "EMA_Direction/Donchian_Direction"
        headers(1, 6) = "EMA_V5_Gates": headers(1, 7) = "EMA_V3_Gates"
        headers(1, 8) = "Donchian_Gates": headers(1, 9) = "Combination"
        headers(1, 10) = "Decision"
        ws.cells(1, 1).Resize(1, 10).value = headers
        ws.Rows(1).Font.Bold = True
        ws.Columns(1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        ws.Columns("A:J").EntireColumn.AutoFit
    End If
    Set GetOrCreateV5Funnel = ws
End Function

Private Sub WriteV5FunnelRow(ws As Worksheet, runTime As Date, ticker As String, _
        crossValue As String, breakoutValue As String, directionValue As String, _
        emaV5Value As String, emaV3Value As String, donchianValue As String, _
        combinationValue As String, decisionValue As String)
    Dim NextRow As Long
    
    'commenting the rows below
    
'    nextRow = ws.cells(ws.Rows.count, 1).End(xlUp).row + 1
'    Dim rowData(1 To 1, 1 To 10) As Variant
'    rowData(1, 1) = runTime: rowData(1, 2) = ticker
'    rowData(1, 3) = crossValue: rowData(1, 4) = breakoutValue
'    rowData(1, 5) = directionValue: rowData(1, 6) = emaV5Value
'    rowData(1, 7) = emaV3Value: rowData(1, 8) = donchianValue
'    rowData(1, 9) = combinationValue: rowData(1, 10) = decisionValue
'    ws.cells(nextRow, 1).Resize(1, 10).value = rowData
End Sub

Private Sub EnsureCockpitV5Headers(wsC As Worksheet)
    Dim lastDataRow As Long
    lastDataRow = wsC.cells(wsC.Rows.count, 1).End(xlUp).row
    If lastDataRow < 4 Then lastDataRow = 4

    ' Clear legacy columns AJ:BB from V2/V4 analytics (AI is now LAST_BAR_TS, don't clear)
    If Len(CStr(wsC.Range("AJ4").value)) > 0 Then
        wsC.Range("AJ4:BB" & lastDataRow).ClearContents
        wsC.Range("AJ4:BB" & lastDataRow).Interior.ColorIndex = xlNone
    End If

    wsC.Range(CK_PEAK_PRICE & "4").value = "Peak_Price"
    wsC.Range(CK_BEST_MFE_PCT & "4").value = "Best_MFE_Pct"
    wsC.Range(CK_LAST_MFE_BAR & "4").value = "Last_MFE_Bar"
    wsC.Range(CK_MFE_BAR1 & "4").value = "MFE_Bar1"
    wsC.Range(CK_IS_RUNNER & "4").value = "Is_Runner"
    wsC.Range(CK_LAST_BAR_TS & "4").value = "Last_Bar_TS"
    wsC.Columns(CK_PEAK_PRICE & ":" & CK_LAST_BAR_TS).Hidden = True
    wsC.Columns("AJ:BB").Hidden = True
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

Private Function GetATR5m(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetATR5m = 0
    If mRowLookup Is Nothing Then Exit Function
    If mRowLookup.Exists(ticker) Then
        GetATR5m = CDbl(val(wsSrc.cells(CLng(mRowLookup(ticker)), "CC").value))
    End If
End Function

Private Function GetCross1030(wsSrc As Worksheet, ByVal ticker As String) As String
    GetCross1030 = ""
    If mRowLookup Is Nothing Then Exit Function
    If mRowLookup.Exists(ticker) Then
        GetCross1030 = CStr(wsSrc.cells(CLng(mRowLookup(ticker)), "AC").value)
    End If
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



Attribute VB_Name = "Module_Cockpit"
'Attribute VB_Name = "Module_COCKPIT"
Option Explicit

'----------------------------------------------------------------------
' Module_COCKPIT  --  BTB V5.00 (Clean Rewrite)
'
' Every gate is a named Boolean constant at the top. Nothing hardcoded.
' DEBUG_MODE = True writes rejection reasons to COCKPIT.
' Strategies:
'   1. Momentum Long (Strong)   — TRENDING_BULL + Grade A/A+
'   2. Momentum Long (Weak)     — TRENDING_BULL + Grade B
'   3. Momentum Short           — TRENDING_BEAR + Grade A/A+
'   4. Breakout Long            — RANGING + Breakout Flag=1 + Grade A/A+
'   5. Breakout Short           — RANGING + Breakout Flag=-1 + Grade A/A+
'   6. RSI Exhaustion Long      — ANY + ADX<30 + Multi-TF RSI oversold
'   7. RSI Exhaustion Short     — ANY + ADX<30 + Multi-TF RSI overbought
'   8. Momentum Long MILD       — MILD + cross Long + Sector TB/MILD gate
' Zero-clipboard design — all formatting is direct property assignment.
'----------------------------------------------------------------------

' =====================================================================
'  MASTER TOGGLES — set False to completely disable a feature
' =====================================================================
Private Const DEBUG_MODE                As Boolean = False   ' Set True to log rejections to Debug sheet
Private Const ENABLE_BREAKOUT           As Boolean = True
Private Const ENABLE_RSI_EXHAUSTION     As Boolean = True
Private Const ENABLE_MOMENTUM           As Boolean = True
Private Const ENABLE_MOMENTUM_MILD      As Boolean = True

' =====================================================================
'  REGIME / INDEX GATES
' =====================================================================
Private Const GATE_NIFTY_ENABLED        As Boolean = False    ' Block if Nifty regime check fails?
Private Const GATE_NIFTY_BLOCK_RANGING  As Boolean = True    ' Block ALL when Nifty=RANGING?
Private Const GATE_SECTOR_MILD_ENABLED  As Boolean = True    ' Sector TB/MILD gate for MILD strategy?

' =====================================================================
'  MOMENTUM STRATEGY GATES
' =====================================================================
Private Const GATE_MOM_ADX_ENABLED      As Boolean = True    ' ADX minimum check?
Private Const GATE_MOM_ADX_MIN          As Double = 20
Private Const GATE_MOM_VOL_ENABLED      As Boolean = True    ' Volume ratio check?
Private Const GATE_MOM_VOL_MIN          As Double = 0.5      ' NI-derived 2026: moderate volume + strong ER is more robust
Private Const GATE_MOM_ER10_ENABLED     As Boolean = True    ' Kaufman ER10 efficiency check?
Private Const GATE_MOM_ER10_MIN         As Double = 0.25     ' DECOUPLED ER10: 0.25 for Momentum trend pullbacks (empirically validated)
Private Const GATE_MOM_RSI_ENABLED      As Boolean = True    ' RSI range check?
Private Const GATE_MOM_RSI_BUY_MIN      As Double = 38
Private Const GATE_MOM_RSI_BUY_MAX      As Double = 80     ' LOGIC REVIEW: 55 killed trend continuation. RSI 50-70 is standard trend zone (Wilder/Brown)
Private Const GATE_MOM_RSI_SELL_MIN     As Double = 20     ' LOGIC REVIEW: 45 blocked bearish continuation. 30 is standard oversold boundary
Private Const GATE_MOM_RSI_SELL_MAX     As Double = 62
Private Const GATE_MOM_DIST_HI_ENABLED  As Boolean = True    ' Distance from day high?
Private Const GATE_MOM_DIST_HI_MIN      As Double = 1      ' LOGIC REVIEW: 3 ATR killed breakout follow-through and trend entries near highs
Private Const GATE_MOM_DIST_LO_ENABLED  As Boolean = True    ' Distance from day low?
Private Const GATE_MOM_DIST_LO_MIN      As Double = 1
Private Const GATE_MOM_10V30_ONLY       As Boolean = False   ' NI first-pass accepts 10/30 and 10/50 crosses

' =====================================================================
'  MILD MOMENTUM GATES (inherits momentum gates above)
' =====================================================================
Private Const MILD_MOM_ER10_MIN         As Double = 0.25     ' DECOUPLED ER10: 0.25 for MILD momentum (same as regular momentum)
Private Const MILD_MOM_SIZE             As Double = 0.25
Private Const MILD_MOM_TARGET_PCT       As Double = 0.005

' =====================================================================
'  BREAKOUT STRATEGY GATES
' =====================================================================
Private Const GATE_BO_VOL_ENABLED       As Boolean = True    ' Volume surge check?
Private Const GATE_BO_VOL_MIN           As Double = 1.25
Private Const GATE_BO_GRADE_ENABLED     As Boolean = True    ' FIX 4: Enabled — Grade C breakouts have no downside buffer, 9/14 entered trades were Grade C
Private Const GATE_BO_CONVICTION_ENABLED As Boolean = True   ' Breakout needs efficiency + structure, not volume alone
Private Const GATE_BO_ER10_MIN          As Double = 0.35     ' DECOUPLED ER10: 0.35 for Breakout Strategy (Impulse required)
Private Const GATE_RR_ER10_MIN          As Double = 0.25     ' DECOUPLED ER10: 0.25 for Range Rotation Strategy (Range bounds)
Private Const GATE_BO_ER10_PCTL_MIN     As Double = 0.5      ' Cross-sectional: active-universe median ER
Private Const GATE_BO_VOL_PCTL_MIN      As Double = 0.5      ' Cross-sectional: active-universe median volume ratio
Private Const BREAKOUT_N                As Long = 5
Private Const BREAKOUT_SIZE             As Double = 0.5

' =====================================================================
'  RSI EXHAUSTION STRATEGY GATES
' =====================================================================
Private Const GATE_RSI_EXH_ADX_ENABLED  As Boolean = True    ' ADX < 30 check?
Private Const RSI_EXH_ADX_MAX           As Double = 30
Private Const RSI_EXH_5M_OS             As Double = 20      ' Oversold: 5m RSI <= 20
Private Const RSI_EXH_15M_OS            As Double = 20      ' Oversold: 15m RSI <= 20
Private Const RSI_EXH_1H_OS             As Double = 35      ' Oversold: 1h RSI <= 20
Private Const RSI_EXH_5M_OB             As Double = 80      ' Overbought: 5m RSI >= 80
Private Const RSI_EXH_15M_OB            As Double = 80      ' Overbought: 15m RSI >= 80
Private Const RSI_EXH_1H_OB             As Double = 65      ' Overbought: 1h RSI >= 80
Private Const RSI_EXH_SIZE              As Double = 0.5

' =====================================================================
'  REGIME PERSISTENCE
' =====================================================================
Private Const MIN_REGIME_STABLE_BARS    As Long = 1         ' LOGIC REVIEW: 3 bars = 15 min delay, causes stale entries. Accept regime on first reading
Private Const GATE_RSI_SLOPE_ENABLED    As Boolean = False  ' LOGIC REVIEW: Conflicts with cross-is-the-turn-signal. Cross IS the momentum shift
Private Const GATE_RSI_SLOPE_THRESHOLD  As Double = 2       ' Min RSI drop/rise (points) to trigger block

' =====================================================================
'  SUPPORT / RESISTANCE LANDSCAPE GATE
' =====================================================================
Private Const GATE_PIVOT_ENABLED        As Boolean = False  ' LOGIC REVIEW: Pivot grid proximity is always triggered; not meaningful resistance detection
Private Const GATE_PIVOT_PROXIMITY_ATR  As Double = 0.5     ' Existing hard obstacle distance
Private Const GATE_SR_CONTEXT_ATR       As Double = 1#      ' Only judge room/risk inside next 1 ATR
Private Const GATE_SR_MIN_ROOM_RISK     As Double = 1#      ' Minimum structural room:risk sanity check

' =====================================================================
'  STAGNATION & TRAILING BREAKEVEN EXITS
' =====================================================================
Private Const STAGNATION_BARS           As Long = 5         ' FIX 3: Was 3 (15 min). ETERNAL cut at -0.06%, LAURUSLABS at -0.07% = spread noise. 5 bars = 25 min
Private Const STAGNATION_MIN_MFE_PCT    As Double = 0.0005  ' 0.05% min move threshold for stagnation
Private Const TRAILING_BE_ATR_MULT      As Double = 1#      ' Lock Breakeven (entry+fee) once MFE >= 1.0 x ATR
Private Const DEAD_TRADE_EXIT_MAX_PCT   As Double = -0.0005  ' FIX 3: Was -0.0002 (-0.02%). Only cut if actually -0.3% adverse, not bid-ask noise
Private Const PROFIT_PROTECT_MIN_PCT    As Double = 0.0025  ' Arm winner management after 0.25% profit
Private Const PROFIT_PROTECT_TGT_FRAC   As Double = 0.35    ' Or after 35% progress to target
Private Const PROFIT_STALL_BARS         As Long = 2         ' Book winners if no fresh MFE for this many cycles
Private Const PROFIT_RETRACE_FRAC       As Double = 0.35    ' Book winners after 35% giveback from peak MFE
Private Const PROFIT_RETRACE_MIN_PCT    As Double = 0.0015  ' Minimum 0.15% giveback before reversal exit

' =====================================================================
'  HIGH-CONVICTION ENTRY CONFIRMATION
' =====================================================================
Private Const GATE_DIRECTION_CONFIRM_ENABLED  As Boolean = True
Private Const GATE_SUPERTREND_CONFIRM_ENABLED As Boolean = False  ' LOGIC REVIEW: Redundant with regime+cross+ADX. Kills 30-40% of valid entries

' =====================================================================
'  RISK MANAGEMENT (Dynamic SL / Target)
' =====================================================================
Private Const USE_DYNAMIC_SL            As Boolean = True
Private Const SL_ATR_MOM_BULL           As Double = 1.5    ' Backtest-validated (was 2.0)
Private Const SL_ATR_MOM_BEAR           As Double = 1.5    ' Backtest-validated (was 2.0)
Private Const SL_ATR_MILD               As Double = 1#     ' Proportional to 0.5% target (was 2.0)
Private Const SL_ATR_BREAKOUT           As Double = 1.5
Private Const SL_ATR_RSI_EXH            As Double = 1#
Private Const TARGET_PCT_MOM_BULL       As Double = 0.015  ' Backtest-validated (was 0.02)
Private Const TARGET_PCT_MOM_BEAR       As Double = 0.015  ' Backtest-validated (was 0.02)
Private Const TARGET_PCT_MILD           As Double = 0.005
Private Const TARGET_PCT_BREAKOUT       As Double = 0.015
Private Const TARGET_PCT_RSI_EXH        As Double = 0.01

' =====================================================================
'  POSITION SIZES
' =====================================================================
Private Const SIZE_MOM_STRONG           As Double = 1#
Private Const SIZE_MOM_WEAK             As Double = 0.5
Private Const SIZE_MOM_BEAR             As Double = 0.5
Private Const SIZE_MILD                 As Double = 0.25
Private Const SIZE_BREAKOUT             As Double = 0.5
Private Const SIZE_RSI_EXH              As Double = 0.5
Private Const SIZE_RANGE_ROTATION       As Double = 0.25

' =====================================================================
'  SESSION / EXIT
' =====================================================================
Private Const LATE_ENTRY_CUTOFF         As String = "14:45:00"
Private Const FORCE_LIQUIDATION         As String = "15:00:00"
Private Const MAX_CANDLES_EXIT          As Long = 9
Private Const COOLDOWN_MINUTES_VAL      As Long = 30
Private Const DATA_START_ROW            As Long = 5

' --- NI (New_Logic_Intraday) COLUMN INDICES (1-based, within 29-col array) --
Private Const NI_TICKER       As Long = 1
Private Const NI_ACTION       As Long = 2
Private Const NI_LTP          As Long = 3
Private Const NI_SECTOR       As Long = 29       ' col AC
Private Const NI_S3           As Long = 6
Private Const NI_S2           As Long = 7
Private Const NI_S1           As Long = 8
Private Const NI_PP           As Long = 9
Private Const NI_R1           As Long = 10
Private Const NI_R2           As Long = 11
Private Const NI_R3           As Long = 12
Private Const NI_RSI_5M       As Long = 13       ' col M
Private Const NI_RSI_15M      As Long = 14       ' col N
Private Const NI_RSI_1H       As Long = 15       ' col O
Private Const NI_RSI_1D       As Long = 16       ' col P
Private Const NI_SUPERTREND   As Long = 17       ' col Q
Private Const NI_CROSS_10_30  As Long = 18       ' col R
Private Const NI_CROSS_10_50  As Long = 20       ' col T  (skip 19=col S,15m cross)

' --- COCKPIT COLUMN LETTERS (post-drop of 10EMA, everything shifted left) --
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
Private Const CK_OUTCOME      As String = "Y"      ' user formula — never written
Private Const CK_BOTTOMLINE   As String = "Z"      ' user formula — never written
Private Const CK_COMENTS      As String = "AA"
Private Const CK_CANDLE_COUNT As String = "AB"     ' number of execution cycles elapsed (Col AB)
Private Const CK_STRATEGY     As String = "AC"     ' Strategy name (Col AC)
Private Const CK_PEAK_PRICE   As String = "AD"     ' best favorable price seen while trade is open
Private Const CK_BEST_MFE_PCT As String = "AE"     ' best favorable excursion as decimal pct
Private Const CK_LAST_MFE_BAR As String = "AF"     ' last candle count where MFE improved

' --- Module-level state (persists across calls) ---
Private mRowLookup As Object
Private mActiveTrades As Object
Private mCooldowns As Object
Private mRegimeHistory As Object          ' Regime persistence: ticker -> "regime|count"
Private mPrevRSI As Object                ' Previous cycle RSI: ticker -> Double
Public mStatusArray As Variant            ' Status per NI row — read by NL_ID_Cumulative


' =====================================================================
' Public Sub  Run_Cockpit   (V5.00 Clean Rewrite)
' Every check is a named Boolean constant — nothing hardcoded.
' =====================================================================
Public Sub Run_Cockpit()
    Dim currentTime As Date
    currentTime = TimeValue(Now)

    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsNI   As Worksheet: Set wsNI = ThisWorkbook.Sheets("New_Logic_Intraday")
    Dim wsSrc  As Worksheet: Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Dim wsC    As Worksheet: Set wsC = ThisWorkbook.Sheets("COCKPIT")

    EnsureCockpitStateHeaders wsC
    On Error Resume Next
    Update_K3_Dashboard wsSrc, wsC
    On Error GoTo 0
    Set mRowLookup = BuildRowLookup(wsSrc)
    Set mActiveTrades = BuildActiveTradeLookup(wsC)

    ' -- Load NI data --
    Dim niLastRow As Long
    niLastRow = wsNI.cells(wsNI.Rows.count, 1).End(xlUp).row
    If niLastRow < 3 Then GoTo Cleanup

    Dim niData As Variant
    niData = wsNI.Range("A3:AC" & niLastRow).value

    ' -- TIME GATE: no new entries after cutoff --
    If currentTime >= TimeValue(LATE_ENTRY_CUTOFF) Then
        ManageCockpitTrades wsC, wsSrc
        GoTo Cleanup
    End If

    ' -- NIFTY GATE: check once per cycle --
    Dim niftyIsRanging As Boolean: niftyIsRanging = False
    If GATE_NIFTY_ENABLED Then
        niftyIsRanging = GetNiftyIsRanging(wsSrc)
    End If

    ' -- Find next empty row in COCKPIT --
    Dim lastDestRow As Long
    lastDestRow = wsC.cells(wsC.Rows.count, 1).End(xlUp).row
    If lastDestRow < DATA_START_ROW Then lastDestRow = DATA_START_ROW - 1

    ' ===== SCAN EVERY NI ROW =====
    Dim r As Long, destRow As Long, rejectReason As String
    destRow = lastDestRow
    ReDim mStatusArray(1 To UBound(niData, 1))

    For r = 1 To UBound(niData, 1)
        Dim entryTicker As String: entryTicker = SafeStr(niData(r, NI_TICKER))
        If Len(entryTicker) = 0 Then GoTo NextNI

        ' --- GATE 1: Block indices ---
        If Is_Index(entryTicker) Then
            rejectReason = "INDEX / Ticker=" & entryTicker: GoTo LogReject
        End If

        ' --- GATE 2: Nifty regime gate ---
        If niftyIsRanging And GATE_NIFTY_BLOCK_RANGING Then
            rejectReason = "NIFTY_RANGING": GoTo LogReject
        End If

        ' --- Fetch regime data for this stock ---
        Dim stockRegime As String: stockRegime = GetStableRegime(wsSrc, entryTicker)
        Dim stockGrade As String: stockGrade = GetStockGrade(wsSrc, entryTicker)
        Dim breakoutFlag As Long: breakoutFlag = GetBreakoutFlag(wsSrc, entryTicker)
        Dim stockADX As Double: stockADX = GetADX(wsSrc, entryTicker)
        Dim secRegime As String: secRegime = GetSectorRegime(wsSrc, entryTicker)

        ' --- Read EMA cross from NI ---
        Dim cross1030 As String: cross1030 = LCase$(SafeStr(niData(r, NI_CROSS_10_30)))
        Dim cross1050 As String: cross1050 = LCase$(SafeStr(niData(r, NI_CROSS_10_50)))
        Dim hasCross As Boolean: hasCross = False
        Dim isBuyCross As Boolean: isBuyCross = False

        If cross1030 = "over" Or cross1030 = "under" Then
            hasCross = True
            isBuyCross = (cross1030 = "over")
        ElseIf Not GATE_MOM_10V30_ONLY Then
            If cross1050 = "over" Or cross1050 = "under" Then
                hasCross = True
                isBuyCross = (cross1050 = "over")
            End If
        End If

        ' --- Build diagnostic context for all messages ---
        Dim diagCtx As String: diagCtx = " / Stock=" & stockRegime & " Sctr=" & secRegime & " Grd=" & stockGrade
        If hasCross Then
            diagCtx = diagCtx & " X=" & IIf(isBuyCross, "OVER", "UNDER")
        Else
            diagCtx = diagCtx & " X=NONE"
        End If

        ' ===== STRATEGY MATCHING =====
        Dim stratName As String: stratName = ""
        Dim stratIsBuy As Boolean: stratIsBuy = False
        Dim stratSize As Double: stratSize = 0
        Dim stratSlMult As Double: stratSlMult = 2#
        Dim stratTgtPct As Double: stratTgtPct = 0.02

        ' -- CONTRA TRADE GATE --
        If (stockRegime = "TRENDING_BULL" And (secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN")) Then
            rejectReason = "CONTRA" & diagCtx: GoTo LogReject
        End If
        If (stockRegime = "TRENDING_BEAR" And (secRegime = "TRENDING_BULL" Or secRegime = "MILD_UP")) Then
            rejectReason = "CONTRA" & diagCtx: GoTo LogReject
        End If
        If (stockRegime = "MILD_UP" And (secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN")) Then
            rejectReason = "CONTRA" & diagCtx: GoTo LogReject
        End If
        If (stockRegime = "MILD_DOWN" And (secRegime = "TRENDING_BULL" Or secRegime = "MILD_UP")) Then
            rejectReason = "CONTRA" & diagCtx: GoTo LogReject
        End If

        ' -- STRATEGY: BREAKOUT (RANGING regime + breakout flag) --
        If ENABLE_BREAKOUT And stockRegime = "RANGING" And breakoutFlag <> 0 Then
            Dim boGradeOk As Boolean: boGradeOk = True
            If GATE_BO_GRADE_ENABLED Then
                boGradeOk = (stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B")
            End If
            If boGradeOk And breakoutFlag = 1 Then
                stratName = "Breakout_Long": stratIsBuy = True
                stratSize = SIZE_BREAKOUT: stratSlMult = SL_ATR_BREAKOUT
                stratTgtPct = TARGET_PCT_BREAKOUT
            ElseIf boGradeOk And breakoutFlag = -1 Then
                stratName = "Breakout_Short": stratIsBuy = False
                stratSize = SIZE_BREAKOUT: stratSlMult = SL_ATR_BREAKOUT
                stratTgtPct = TARGET_PCT_BREAKOUT
            End If
        End If

        ' -- STRATEGY: RANGE ROTATION (RANGING regime + cross + NO breakout flag) --
        If stratName = "" And ENABLE_BREAKOUT And stockRegime = "RANGING" And breakoutFlag = 0 Then
            If hasCross Then
                If isBuyCross Then
                    ' FIX: Apply Sector Alignment to Range Rotation
                    If secRegime = "TRENDING_BULL" Or secRegime = "MILD_UP" Or secRegime = "RANGING" Then
                        stratName = "Range_Rotation_Long": stratIsBuy = True
                        stratSize = SIZE_RANGE_ROTATION: stratSlMult = SL_ATR_MILD
                        stratTgtPct = TARGET_PCT_MILD
                    End If
                Else
                    ' FIX: Apply Sector Alignment to Range Rotation
                    If secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN" Or secRegime = "RANGING" Then
                        stratName = "Range_Rotation_Short": stratIsBuy = False
                        stratSize = SIZE_RANGE_ROTATION: stratSlMult = SL_ATR_MILD
                        stratTgtPct = TARGET_PCT_MILD
                    End If
                End If
            End If
        End If

        ' -- STRATEGY: RSI EXHAUSTION (any regime, ADX gate required, aligned sector) --
        If stratName = "" And ENABLE_RSI_EXHAUSTION Then
            Dim adxOk As Boolean: adxOk = True
            If GATE_RSI_EXH_ADX_ENABLED Then
                adxOk = (stockADX < RSI_EXH_ADX_MAX)
            End If
            If adxOk Then
                Dim rsiSignal As String: rsiSignal = CheckRSIExhaustion(niData, r)
                If rsiSignal = "LONG" And secRegime <> "TRENDING_BEAR" And secRegime <> "MILD_DOWN" Then
                    stratName = "RSI_Exh_Long": stratIsBuy = True
                    stratSize = SIZE_RSI_EXH: stratSlMult = SL_ATR_RSI_EXH
                    stratTgtPct = TARGET_PCT_RSI_EXH
                ElseIf rsiSignal = "SHORT" And secRegime <> "TRENDING_BULL" And secRegime <> "MILD_UP" Then
                    stratName = "RSI_Exh_Short": stratIsBuy = False
                    stratSize = SIZE_RSI_EXH: stratSlMult = SL_ATR_RSI_EXH
                    stratTgtPct = TARGET_PCT_RSI_EXH
                End If
            End If
        End If

        ' -- STRATEGY: MOMENTUM (TRENDING_BULL/BEAR + cross matching) --
        If stratName = "" And ENABLE_MOMENTUM And hasCross Then
            If stockRegime = "TRENDING_BULL" And isBuyCross Then
                If secRegime = "TRENDING_BULL" Or secRegime = "MILD_UP" Or secRegime = "RANGING" Then
                    If stockGrade = "A+" Or stockGrade = "A" Then
                        stratName = "Momentum_Long_Strong": stratSize = SIZE_MOM_STRONG
                    ElseIf stockGrade = "B" Then
                        stratName = "Momentum_Long_Weak": stratSize = SIZE_MOM_WEAK
                    ElseIf stockGrade = "C" Then
                        stratName = "Momentum_Long_GradeC": stratSize = SIZE_MILD
                    End If
                    stratIsBuy = True: stratSlMult = SL_ATR_MOM_BULL
                    stratTgtPct = TARGET_PCT_MOM_BULL
                End If
            ElseIf stockRegime = "TRENDING_BEAR" And Not isBuyCross Then
                If secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN" Or secRegime = "RANGING" Then
                    stratName = "Momentum_Short": stratSize = SIZE_MOM_BEAR
                    stratIsBuy = False: stratSlMult = SL_ATR_MOM_BEAR
                    stratTgtPct = TARGET_PCT_MOM_BEAR
                End If
            End If
        End If

        ' -- STRATEGY: MILD (MILD_UP/MILD_DOWN + cross) --
        If stratName = "" And ENABLE_MOMENTUM_MILD And hasCross Then
            If stockRegime = "MILD_UP" And isBuyCross Then
                If secRegime = "TRENDING_BULL" Or secRegime = "MILD_UP" Or secRegime = "RANGING" Then
                    If stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B" Then
                        stratName = "Momentum_Long_MILD": stratSize = SIZE_MILD
                        stratIsBuy = True: stratSlMult = SL_ATR_MILD
                        stratTgtPct = TARGET_PCT_MILD
                    End If
                End If
            ElseIf stockRegime = "MILD_UP" And Not isBuyCross Then
                If secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN" Or secRegime = "RANGING" Then
                    If stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B" Then
                        stratName = "Momentum_Short_MILD": stratSize = SIZE_MILD
                        stratIsBuy = False: stratSlMult = SL_ATR_MILD
                        stratTgtPct = TARGET_PCT_MILD
                    End If
                End If
            ElseIf stockRegime = "MILD_DOWN" And isBuyCross Then
                If secRegime = "TRENDING_BULL" Or secRegime = "MILD_UP" Or secRegime = "RANGING" Then
                    If stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B" Then
                        stratName = "Momentum_Long_MILD": stratSize = SIZE_MILD
                        stratIsBuy = True: stratSlMult = SL_ATR_MILD
                        stratTgtPct = TARGET_PCT_MILD
                    End If
                End If
            ElseIf stockRegime = "MILD_DOWN" And Not isBuyCross Then
                If secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN" Or secRegime = "RANGING" Then
                    If stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B" Then
                        stratName = "Momentum_Short_MILD": stratSize = SIZE_MILD
                        stratIsBuy = False: stratSlMult = SL_ATR_MILD
                        stratTgtPct = TARGET_PCT_MILD
                    End If
                End If
            End If
        End If

        ' --- GATE: No strategy matched ---
        If stratName = "" Then
            If Not hasCross Then
                rejectReason = "NO_CROSS" & diagCtx
            Else
                rejectReason = "NO_STRATEGY" & diagCtx & " Flag=" & breakoutFlag
            End If
            GoTo LogReject
        End If

        ' --- GATE: Duplicate active trade ---
        If SymbolAlreadyActive(wsC, entryTicker) Then
            rejectReason = "DUPLICATE" & diagCtx: GoTo LogReject
        End If

        ' --- GATE: Cooldown ---
        If IsOnCooldown(entryTicker) Then
            rejectReason = "COOLDOWN" & diagCtx: GoTo LogReject
        End If

        ' --- STRATEGY-SPECIFIC GATES ---
        If Not PassStrategyGates(wsSrc, niData, r, entryTicker, stratName, stratIsBuy, _
                                 stockADX, secRegime, rejectReason) Then
            rejectReason = rejectReason & diagCtx
            GoTo LogReject
        End If

        ' --- GATE: Pivot proximity ---
        If GATE_PIVOT_ENABLED Then
            If Not PassPivotGate(niData, r, stratIsBuy, wsSrc, entryTicker, rejectReason) Then
                rejectReason = rejectReason & diagCtx
                GoTo LogReject
            End If
        End If

        ' --- ALL GATES PASSED: Write to COCKPIT ---
        destRow = destRow + 1
        WriteNewCockpitRow wsC, niData, r, destRow, wsSrc, stratIsBuy, stratName, stratSize, stratSlMult, stratTgtPct, stockRegime, stockGrade, secRegime
        If Not mActiveTrades Is Nothing Then mActiveTrades(entryTicker) = True
        mStatusArray(r) = "ENTERED: " & stratName & diagCtx
        GoTo NextNI

LogReject:
        mStatusArray(r) = "REJECTED: " & rejectReason
        If DEBUG_MODE Then LogDebugReject entryTicker, rejectReason

NextNI:
    Next r

    ' -- Update RSI snapshot --
    If mPrevRSI Is Nothing Then Set mPrevRSI = CreateObject("Scripting.Dictionary")
    Dim ri As Long
    For ri = 1 To UBound(niData, 1)
        Dim rsiTicker As String: rsiTicker = SafeStr(niData(ri, NI_TICKER))
        If Len(rsiTicker) > 0 Then mPrevRSI(rsiTicker) = SafeVal(niData(ri, NI_RSI_5M))
    Next ri

    ' -- Manage existing open trades --
    ManageCockpitTrades wsC, wsSrc

Cleanup:
    Set mRowLookup = Nothing
    Set mActiveTrades = Nothing
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Sub


' =====================================================================
' Private Function  BuildRowLookup
' Builds a Scripting.Dictionary mapping ticker->row number in Sheet1.
' Replaces Application.Match for O(1) lookups.
' =====================================================================
Private Function BuildRowLookup(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim src As Variant, i As Long
    Dim lastRow As Long
    lastRow = ws.cells(ws.Rows.count, "A").End(xlUp).row
    If lastRow < 1 Then Set BuildRowLookup = d: Exit Function
    src = ws.Range("A1:A" & lastRow).value
    For i = 1 To UBound(src, 1)
        Dim key As String: key = CStr(src(i, 1))
        If Len(key) > 0 Then
            d(key) = i
        End If
    Next i
    Set BuildRowLookup = d
End Function


' =====================================================================
' Private Sub  EnsureCockpitStateHeaders
' Adds hidden state fields used by adaptive exits. These are intentionally
' outside the user-facing COCKPIT table.
' =====================================================================
Private Sub EnsureCockpitStateHeaders(wsC As Worksheet)
    If Len(CStr(wsC.Range(CK_PEAK_PRICE & "4").value)) = 0 Then wsC.Range(CK_PEAK_PRICE & "4").value = "Peak_Price"
    If Len(CStr(wsC.Range(CK_BEST_MFE_PCT & "4").value)) = 0 Then wsC.Range(CK_BEST_MFE_PCT & "4").value = "Best_MFE_Pct"
    If Len(CStr(wsC.Range(CK_LAST_MFE_BAR & "4").value)) = 0 Then wsC.Range(CK_LAST_MFE_BAR & "4").value = "Last_MFE_Bar"

    wsC.Columns(CK_PEAK_PRICE & ":" & CK_LAST_MFE_BAR).Hidden = True
End Sub


' =====================================================================
' Private Function  BuildActiveTradeLookup
' Builds a Scripting.Dictionary of tickers currently active in COCKPIT
' =====================================================================
Private Function BuildActiveTradeLookup(wsC As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long
    lastRow = wsC.cells(wsC.Rows.count, 1).End(xlUp).row
    If lastRow < DATA_START_ROW Then Set BuildActiveTradeLookup = d: Exit Function

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim ticker As String: ticker = SafeStr(wsC.Range(CK_TICKER & r).value)
        Dim ct As String: ct = SafeStr(wsC.Range(CK_CLOSE_TIME & r).value)
        If Len(ticker) > 0 And Len(ct) = 0 Then
            d(ticker) = True
        End If
    Next r
    Set BuildActiveTradeLookup = d
End Function


' =====================================================================
' Private Function  SymbolAlreadyActive
' Returns True if ticker already has an open trade in COCKPIT (O(1) lookup)
' =====================================================================
Private Function SymbolAlreadyActive(wsC As Worksheet, ticker As String) As Boolean
    If mActiveTrades Is Nothing Then Set mActiveTrades = BuildActiveTradeLookup(wsC)
    SymbolAlreadyActive = mActiveTrades.Exists(ticker)
End Function


' =====================================================================
' Private Function  Is_Index
' Blocks index tickers from entering COCKPIT (same logic as Paper_Trades).
' =====================================================================
Private Function Is_Index(t As String) As Boolean
    Dim tUCase As String: tUCase = UCase(t)
    Is_Index = (InStr(tUCase, "NIFTY") > 0 Or InStr(tUCase, "CNX") > 0 Or _
                InStr(tUCase, "NSEBANK") > 0 Or InStr(tUCase, "MIDCAP") > 0 Or _
                InStr(tUCase, "SMALLCAP") > 0 Or InStr(tUCase, "NSEI") > 0 Or _
                InStr(tUCase, "LARGECAP") > 0)
End Function



' =====================================================================
' Private Function  PassStrategyGates
' Strategy-specific entry filters. Returns True if all gates pass.
' Updates rejectReason on failure.
' =====================================================================
Private Function PassStrategyGates(wsSrc As Worksheet, niData As Variant, _
        niIdx As Long, ticker As String, stratName As String, _
        isBuy As Boolean, adxVal As Double, ByVal secRegime As String, _
        ByRef rejectReason As String) As Boolean

    PassStrategyGates = False

    ' == RSI SLOPE CHECK (all momentum strategies) ==
    If GATE_RSI_SLOPE_ENABLED Then
        If stratName <> "Breakout_Long" And stratName <> "Breakout_Short" And _
           stratName <> "RSI_Exh_Long" And stratName <> "RSI_Exh_Short" Then
            If Not mPrevRSI Is Nothing Then
                If mPrevRSI.Exists(ticker) Then
                    Dim prevRsiSlope As Double: prevRsiSlope = CDbl(mPrevRSI(ticker))
                    Dim curRsiSlope As Double: curRsiSlope = val(niData(niIdx, NI_RSI_5M))
                    If isBuy And curRsiSlope < prevRsiSlope - GATE_RSI_SLOPE_THRESHOLD Then
                        rejectReason = "RSI_FALLING=" & Format(curRsiSlope, "0.0") & "<" & Format(prevRsiSlope, "0.0")
                        Exit Function
                    End If
                    If Not isBuy And curRsiSlope > prevRsiSlope + GATE_RSI_SLOPE_THRESHOLD Then
                        rejectReason = "RSI_RISING=" & Format(curRsiSlope, "0.0") & ">" & Format(prevRsiSlope, "0.0")
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    ' == DIRECTION CONFIRMATION ==
    If GATE_DIRECTION_CONFIRM_ENABLED Then
        If Not PassDirectionConfirmation(niData, niIdx, stratName, isBuy, rejectReason) Then Exit Function
    End If

    ' == HIGHER TIMEFRAME (HTF) RSI ALIGNMENT & EXHAUSTION GATE ==
    ' FIX: Prevents taking 5m continuation or range entries against heavily entrenched HTF momentum
    If stratName <> "RSI_Exh_Long" And stratName <> "RSI_Exh_Short" Then
        Dim htfRsi15m As Double: htfRsi15m = val(niData(niIdx, NI_RSI_15M))
        Dim htfRsi1H As Double:  htfRsi1H = val(niData(niIdx, NI_RSI_1H))

        If isBuy Then
            ' Block if HTF is heavily bearish (trend conflict)
            If (htfRsi15m > 0 And htfRsi15m < 40) Or (htfRsi1H > 0 And htfRsi1H < 40) Then
                rejectReason = "HTF_RSI_BEARISH(15m=" & htfRsi15m & ",1H=" & htfRsi1H & ")"
                Exit Function
            End If
            ' Block if HTF is already completely exhausted (overbought ceiling)
            If (htfRsi1H > 0 And htfRsi1H > 75) Then
                rejectReason = "HTF_RSI_EXHAUSTED(1H=" & htfRsi1H & ")"
                Exit Function
            End If
        Else
            ' Block if HTF is heavily bullish (trend conflict)
            If (htfRsi15m > 0 And htfRsi15m > 60) Or (htfRsi1H > 0 And htfRsi1H > 60) Then
                rejectReason = "HTF_RSI_BULLISH(15m=" & htfRsi15m & ",1H=" & htfRsi1H & ")"
                Exit Function
            End If
            ' Block if HTF is already completely exhausted (oversold floor)
            If (htfRsi1H > 0 And htfRsi1H < 35) Then
                rejectReason = "HTF_RSI_EXHAUSTED(1H=" & htfRsi1H & ")"
                Exit Function
            End If
        End If
    End If

    ' == ALL MOMENTUM VARIANTS AND RANGE ROTATION ==
    ' FIX: Range Rotation now shares spatial (Dist Hi/Lo) and RSI gates to prevent
    ' buying the top or shorting the bottom of a given range.
    If IsMomentumStrategy(stratName) Or IsRangeRotationStrategy(stratName) Then
        
        ' ER10 and ADX remain specific to Momentum
        If IsMomentumStrategy(stratName) Then
            If GATE_MOM_ER10_ENABLED Then
                Dim er10Val As Double: er10Val = GetStockER10(wsSrc, ticker)
                Dim er10Min As Double: er10Min = GATE_MOM_ER10_MIN
                If IsMildMomentumStrategy(stratName) Then er10Min = MILD_MOM_ER10_MIN
                If er10Val < er10Min Then
                    rejectReason = "ER10=" & Format(er10Val, "0.00") & "<" & er10Min: Exit Function
                End If
            End If

            If GATE_MOM_ADX_ENABLED Then
                If adxVal < GATE_MOM_ADX_MIN Then
                    rejectReason = "ADX<" & GATE_MOM_ADX_MIN: Exit Function
                End If
            End If
        End If

        If GATE_MOM_VOL_ENABLED Then
            If GetVolumeRatio(wsSrc, ticker) < GATE_MOM_VOL_MIN Then
                rejectReason = "VOL<" & GATE_MOM_VOL_MIN: Exit Function
            End If
        End If

        ' 5m RSI boundaries now apply to both to prevent overextension in ranges
        If GATE_MOM_RSI_ENABLED Then
            Dim rsi As Double: rsi = val(niData(niIdx, NI_RSI_5M))
            If isBuy Then
                If rsi < GATE_MOM_RSI_BUY_MIN Or rsi > GATE_MOM_RSI_BUY_MAX Then
                    rejectReason = "RSI=" & rsi: Exit Function
                End If
            Else
                If rsi < GATE_MOM_RSI_SELL_MIN Or rsi > GATE_MOM_RSI_SELL_MAX Then
                    rejectReason = "RSI=" & rsi: Exit Function
                End If
            End If
        End If

        ' Spatial distance checks now prevent shorting the floor / buying the ceiling of a range
        If isBuy Then
            If GATE_MOM_DIST_HI_ENABLED And GetDistDayHi(wsSrc, ticker) < GATE_MOM_DIST_HI_MIN Then
                rejectReason = "HI<" & GATE_MOM_DIST_HI_MIN: Exit Function
            End If
        Else
            If GATE_MOM_DIST_LO_ENABLED And GetDistDayLo(wsSrc, ticker) < GATE_MOM_DIST_LO_MIN Then
                rejectReason = "LO<" & GATE_MOM_DIST_LO_MIN: Exit Function
            End If
        End If

        ' Range Rotation specific cross-sectional conviction
        If IsRangeRotationStrategy(stratName) Then
            If Not PassRangeRotationConviction(wsSrc, ticker, rejectReason) Then Exit Function
        End If

    ' == BREAKOUT ==
    ElseIf stratName = "Breakout_Long" Or stratName = "Breakout_Short" Then
        If GATE_BO_VOL_ENABLED Then
            If GetVolumeRatio(wsSrc, ticker) < GATE_BO_VOL_MIN Then
                rejectReason = "BO_VOL<" & GATE_BO_VOL_MIN: Exit Function
            End If
        End If
        If GATE_BO_CONVICTION_ENABLED Then
            If Not PassBreakoutConviction(wsSrc, niData, niIdx, ticker, isBuy, rejectReason) Then Exit Function
        End If

    ' == RSI EXHAUSTION ==
    ' No additional gates — already filtered by ADX and RSI
    End If

    PassStrategyGates = True
End Function


' =====================================================================
' Private Function  PassBreakoutConviction
' Breakout is not allowed on volume alone. It must show directional
' efficiency and cross-sectional confirmation. The shared S/R landscape
' gate handles pivot location after strategy merit is established.
' =====================================================================
Private Function PassBreakoutConviction(wsSrc As Worksheet, niData As Variant, _
        ByVal niIdx As Long, ByVal ticker As String, ByVal isBuy As Boolean, _
        ByRef rejectReason As String) As Boolean

    PassBreakoutConviction = False

    Dim er10Val As Double: er10Val = GetStockER10Strict(wsSrc, ticker)
    If er10Val < GATE_BO_ER10_MIN Then
        rejectReason = "BO_ER10=" & Format(er10Val, "0.00") & "<" & GATE_BO_ER10_MIN
        Exit Function
    End If

    Dim erPct As Double
    erPct = GetActiveUniversePercentile(wsSrc, ticker, "DP")
    If erPct > 0 And erPct < GATE_BO_ER10_PCTL_MIN Then
        rejectReason = "BO_ER_PCTL=" & Format(erPct, "0.00")
        Exit Function
    End If

    Dim volPct As Double
    volPct = GetActiveUniversePercentile(wsSrc, ticker, "BO")
    If volPct > 0 And volPct < GATE_BO_VOL_PCTL_MIN Then
        rejectReason = "BO_VOL_PCTL=" & Format(volPct, "0.00")
        Exit Function
    End If

    PassBreakoutConviction = True
End Function


' =====================================================================
' Private Function  PassRangeRotationConviction
' Ranging + no breakout flag trades are allowed only when the cross has
' efficient movement, enough participation, and cross-sectional confirmation.
' =====================================================================
Private Function PassRangeRotationConviction(wsSrc As Worksheet, _
        ByVal ticker As String, ByRef rejectReason As String) As Boolean

    PassRangeRotationConviction = False

    Dim rrEr10Val As Double: rrEr10Val = GetStockER10Strict(wsSrc, ticker)
    If rrEr10Val < GATE_RR_ER10_MIN Then
        rejectReason = "RR_ER10=" & Format(rrEr10Val, "0.00") & "<" & GATE_RR_ER10_MIN
        Exit Function
    End If

    Dim rrVolRatio As Double: rrVolRatio = GetVolumeRatio(wsSrc, ticker)
    If rrVolRatio < GATE_MOM_VOL_MIN Then
        rejectReason = "RR_VOL=" & Format(rrVolRatio, "0.00") & "<" & GATE_MOM_VOL_MIN
        Exit Function
    End If

    Dim rrErPct As Double
    rrErPct = GetActiveUniversePercentile(wsSrc, ticker, "DP")
    If rrErPct = 0 Then
        rejectReason = "RR_ER_PCTL_MISSING"
        Exit Function
    End If
    If rrErPct > 0 And rrErPct < GATE_BO_ER10_PCTL_MIN Then
        rejectReason = "RR_ER_PCTL=" & Format(rrErPct, "0.00")
        Exit Function
    End If

    Dim rrVolPct As Double
    rrVolPct = GetActiveUniversePercentile(wsSrc, ticker, "BO")
    If rrVolPct = 0 Then
        rejectReason = "RR_VOL_PCTL_MISSING"
        Exit Function
    End If
    If rrVolPct > 0 And rrVolPct < GATE_BO_VOL_PCTL_MIN Then
        rejectReason = "RR_VOL_PCTL=" & Format(rrVolPct, "0.00")
        Exit Function
    End If

    PassRangeRotationConviction = True
End Function


' =====================================================================
' Private Function  PassNeutralSectorConfirmation
' Sector RANGING is neutral, not opposite. Neutral-sector momentum needs
' cross-sectional stock strength so the sector is not the only evidence.
' =====================================================================
Private Function PassNeutralSectorConfirmation(wsSrc As Worksheet, _
        ByVal ticker As String, ByRef rejectReason As String) As Boolean

    PassNeutralSectorConfirmation = False

    Dim nsErPct As Double
    nsErPct = GetActiveUniversePercentile(wsSrc, ticker, "DP")
    If nsErPct = 0 Then
        rejectReason = "NEUTRAL_ER_PCTL_MISSING"
        Exit Function
    End If
    If nsErPct > 0 And nsErPct < GATE_BO_ER10_PCTL_MIN Then
        rejectReason = "NEUTRAL_ER_PCTL=" & Format(nsErPct, "0.00")
        Exit Function
    End If

    Dim nsVolPct As Double
    nsVolPct = GetActiveUniversePercentile(wsSrc, ticker, "BO")
    If nsVolPct = 0 Then
        rejectReason = "NEUTRAL_VOL_PCTL_MISSING"
        Exit Function
    End If
    If nsVolPct > 0 And nsVolPct < GATE_BO_VOL_PCTL_MIN Then
        rejectReason = "NEUTRAL_VOL_PCTL=" & Format(nsVolPct, "0.00")
        Exit Function
    End If

    PassNeutralSectorConfirmation = True
End Function


' =====================================================================
' Private Function  IsMomentumStrategy
' Returns True for every momentum family variant, including Grade C and MILD.
' =====================================================================
Private Function IsMomentumStrategy(ByVal stratName As String) As Boolean
    IsMomentumStrategy = (Left$(stratName, 8) = "Momentum")
End Function


' =====================================================================
' Private Function  IsMildMomentumStrategy
' MILD momentum needs a stricter ER gate because broader 2026 tests were
' weaker after costs than regular momentum.
' =====================================================================
Private Function IsMildMomentumStrategy(ByVal stratName As String) As Boolean
    IsMildMomentumStrategy = (InStr(1, stratName, "_MILD", vbTextCompare) > 0)
End Function


' =====================================================================
' Private Function  IsRangeRotationStrategy
' Ranging stocks with an NI cross but no breakout flag are handled as
' smaller range-rotation trades, not as breakout or unexplained rejects.
' =====================================================================
Private Function IsRangeRotationStrategy(ByVal stratName As String) As Boolean
    IsRangeRotationStrategy = (Left$(stratName, 14) = "Range_Rotation")
End Function


' =====================================================================
' Private Function  IsRSIExhaustionStrategy
' RSI exhaustion is intentionally counter-trend, so trend-following
' confirmation gates do not apply to it.
' =====================================================================
Private Function IsRSIExhaustionStrategy(ByVal stratName As String) As Boolean
    IsRSIExhaustionStrategy = (stratName = "RSI_Exh_Long" Or stratName = "RSI_Exh_Short")
End Function


' =====================================================================
' Private Function  PassDirectionConfirmation
' EMA cross is only a trigger candidate. Entry also needs sheet action and,
' for trend-following setups, Supertrend to agree with trade direction.
' =====================================================================
Private Function PassDirectionConfirmation(niData As Variant, niIdx As Long, _
        ByVal stratName As String, ByVal isBuy As Boolean, _
        ByRef rejectReason As String) As Boolean

    PassDirectionConfirmation = True
    If IsRSIExhaustionStrategy(stratName) Then Exit Function

    Dim actionText As String
    actionText = LCase$(Trim$(CStr(niData(niIdx, NI_ACTION))))

    If isBuy Then
        If actionText <> "buy" Then
            rejectReason = "ACTION_NOT_BUY=" & actionText
            PassDirectionConfirmation = False
            Exit Function
        End If
    Else
        If actionText <> "sell" Then
            rejectReason = "ACTION_NOT_SELL=" & actionText
            PassDirectionConfirmation = False
            Exit Function
        End If
    End If

    If GATE_SUPERTREND_CONFIRM_ENABLED Then
        Dim stText As String
        stText = LCase$(Trim$(CStr(niData(niIdx, NI_SUPERTREND))))
        If isBuy And stText <> "bullish" Then
            rejectReason = "ST_NOT_BULLISH=" & stText
            PassDirectionConfirmation = False
            Exit Function
        End If
        If Not isBuy And stText <> "bearish" Then
            rejectReason = "ST_NOT_BEARISH=" & stText
            PassDirectionConfirmation = False
            Exit Function
        End If
    End If
End Function


' =====================================================================
' Private Function  PassPivotGate
' Evaluates the full daily pivot ladder from S3 through R3. NI remains the
' signal source; this gate rejects only when nearby S/R leaves poor room.
' =====================================================================
Private Function PassPivotGate(niData As Variant, niIdx As Long, _
        ByVal isBuy As Boolean, wsSrc As Worksheet, _
        ByVal ticker As String, ByRef rejectReason As String) As Boolean

    PassPivotGate = True   ' default: pass

    Dim atrVal As Double: atrVal = GetATR5m(wsSrc, ticker)
    If atrVal <= 0 Then Exit Function

    Dim ltp As Double: ltp = val(niData(niIdx, NI_LTP))
    If ltp <= 0 Then Exit Function

    Dim minDist As Double: minDist = atrVal * GATE_PIVOT_PROXIMITY_ATR
    Dim pp As Double:  pp = val(niData(niIdx, NI_PP))
    Dim r1 As Double:  r1 = val(niData(niIdx, NI_R1))
    Dim r2 As Double:  r2 = val(niData(niIdx, NI_R2))
    Dim r3 As Double:  r3 = val(niData(niIdx, NI_R3))
    Dim s1 As Double:  s1 = val(niData(niIdx, NI_S1))
    Dim s2 As Double:  s2 = val(niData(niIdx, NI_S2))
    Dim s3 As Double:  s3 = val(niData(niIdx, NI_S3))

    Dim levels(1 To 7) As Double
    levels(1) = s3
    levels(2) = s2
    levels(3) = s1
    levels(4) = pp
    levels(5) = r1
    levels(6) = r2
    levels(7) = r3

    Dim nearestAbove As Double: nearestAbove = 1E+15
    Dim nearestBelow As Double: nearestBelow = 0
    Dim i As Long, levelVal As Double
    For i = LBound(levels) To UBound(levels)
        levelVal = levels(i)
        If levelVal > 0 Then
            If levelVal > ltp And levelVal < nearestAbove Then nearestAbove = levelVal
            If levelVal < ltp And levelVal > nearestBelow Then nearestBelow = levelVal
        End If
    Next i

    Dim zoneLabel As String
    zoneLabel = PivotZoneLabel(ltp, s3, s2, s1, pp, r1, r2, r3)

    Dim forwardRoom As Double, structuralRisk As Double
    Dim roomAtr As Double, roomRisk As Double

    If isBuy Then
        If nearestAbove >= 1E+15 Then Exit Function

        forwardRoom = nearestAbove - ltp
        If nearestBelow > 0 Then
            structuralRisk = ltp - nearestBelow
        Else
            structuralRisk = atrVal
        End If
        If structuralRisk <= 0 Then structuralRisk = atrVal

        roomAtr = forwardRoom / atrVal
        roomRisk = forwardRoom / structuralRisk

        If forwardRoom < minDist Then
            rejectReason = "SR_RES_NEAR=" & Format(nearestAbove, "0.00") & _
                           " Zone=" & zoneLabel & _
                           " RoomATR=" & Format(roomAtr, "0.00")
            PassPivotGate = False
            Exit Function
        End If

        If roomAtr < GATE_SR_CONTEXT_ATR And roomRisk < GATE_SR_MIN_ROOM_RISK Then
            rejectReason = "SR_BUY_POOR_ROOM Zone=" & zoneLabel & _
                           " UpATR=" & Format(roomAtr, "0.00") & _
                           " RR=" & Format(roomRisk, "0.00")
            PassPivotGate = False
            Exit Function
        End If
    Else
        If nearestBelow <= 0 Then Exit Function

        forwardRoom = ltp - nearestBelow
        If nearestAbove < 1E+15 Then
            structuralRisk = nearestAbove - ltp
        Else
            structuralRisk = atrVal
        End If
        If structuralRisk <= 0 Then structuralRisk = atrVal

        roomAtr = forwardRoom / atrVal
        roomRisk = forwardRoom / structuralRisk

        If forwardRoom < minDist Then
            rejectReason = "SR_SUP_NEAR=" & Format(nearestBelow, "0.00") & _
                           " Zone=" & zoneLabel & _
                           " RoomATR=" & Format(roomAtr, "0.00")
            PassPivotGate = False
            Exit Function
        End If

        If roomAtr < GATE_SR_CONTEXT_ATR And roomRisk < GATE_SR_MIN_ROOM_RISK Then
            rejectReason = "SR_SELL_POOR_ROOM Zone=" & zoneLabel & _
                           " DnATR=" & Format(roomAtr, "0.00") & _
                           " RR=" & Format(roomRisk, "0.00")
            PassPivotGate = False
            Exit Function
        End If
    End If
End Function


' =====================================================================
' Private Function  PivotZoneLabel
' Returns the current S/R band for diagnostics only.
' =====================================================================
Private Function PivotZoneLabel(ByVal ltp As Double, _
        ByVal s3 As Double, ByVal s2 As Double, ByVal s1 As Double, _
        ByVal pp As Double, ByVal r1 As Double, ByVal r2 As Double, _
        ByVal r3 As Double) As String

    PivotZoneLabel = "PIVOT_UNKNOWN"
    If ltp <= 0 Then Exit Function

    If s3 > 0 And ltp < s3 Then PivotZoneLabel = "BELOW_S3": Exit Function
    If s3 > 0 And s2 > 0 And ltp >= s3 And ltp < s2 Then PivotZoneLabel = "S3_S2": Exit Function
    If s2 > 0 And s1 > 0 And ltp >= s2 And ltp < s1 Then PivotZoneLabel = "S2_S1": Exit Function
    If s1 > 0 And pp > 0 And ltp >= s1 And ltp < pp Then PivotZoneLabel = "S1_PP": Exit Function
    If pp > 0 And r1 > 0 And ltp >= pp And ltp < r1 Then PivotZoneLabel = "PP_R1": Exit Function
    If r1 > 0 And r2 > 0 And ltp >= r1 And ltp < r2 Then PivotZoneLabel = "R1_R2": Exit Function
    If r2 > 0 And r3 > 0 And ltp >= r2 And ltp < r3 Then PivotZoneLabel = "R2_R3": Exit Function
    If r3 > 0 And ltp >= r3 Then PivotZoneLabel = "ABOVE_R3": Exit Function
End Function


' =====================================================================
' Private Function  GetStockRegime
' Reads stock regime from Sheet1 col BN (5-minute regime with MILD split).
' Returns: "TRENDING_BULL", "TRENDING_BEAR", "MILD_UP", "MILD_DOWN", "RANGING", or ""
' =====================================================================
Private Function GetStockRegime(wsSrc As Worksheet, ByVal ticker As String) As String
    GetStockRegime = ""
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetStockRegime = UCase(CStr(wsSrc.cells(rowNum, "BN").value))
        Exit Function
    End If

    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetStockRegime = UCase(CStr(wsSrc.cells(CLng(m), "BN").value))
    On Error GoTo 0
End Function


' =====================================================================
' Private Function  GetStockER10
' Reads Kaufman Efficiency Ratio (10-bar) from Sheet1 col DP (ER_COL).
' Defaults to 1.0 (fail-open) if lookup fails.
' =====================================================================
Private Function GetStockER10(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetStockER10 = 1#   ' fail-open: allow trade if data unavailable
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetStockER10 = CDbl(val(wsSrc.cells(rowNum, "DP").value))
        If GetStockER10 <= 0 Then GetStockER10 = 1#
        Exit Function
    End If

    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then
        GetStockER10 = CDbl(wsSrc.cells(CLng(m), "DP").value)
        If GetStockER10 <= 0 Then GetStockER10 = 1#
    End If
    On Error GoTo 0
End Function


' =====================================================================
' Private Function  GetStockER10Strict
' Reads ER10 from Sheet1 col DP. Returns 0 when missing/invalid so
' conviction gates do not fail open on absent efficiency data.
' =====================================================================
Private Function GetStockER10Strict(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetStockER10Strict = 0
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If

    If rowNum = 0 Then
        On Error Resume Next
        Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
        If Not IsError(m) Then rowNum = CLng(m)
        On Error GoTo 0
    End If

    If rowNum > 0 Then
        GetStockER10Strict = CDbl(val(wsSrc.cells(rowNum, "DP").value))
        If GetStockER10Strict < 0 Then GetStockER10Strict = 0
    End If
End Function


' =====================================================================
' Private Function  GetActiveUniversePercentile
' Percentile rank of ticker's Sheet1 metric within active Sheet1 rows
' P1:R1. Returns 0 when target data is missing; returns -1 if the
' universe cannot be evaluated, allowing absolute validated gates to rule.
' =====================================================================
Private Function GetActiveUniversePercentile(wsSrc As Worksheet, _
        ByVal ticker As String, ByVal metricCol As String) As Double

    GetActiveUniversePercentile = -1
    If Len(ticker) = 0 Or Len(metricCol) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum = 0 Then
        On Error Resume Next
        Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
        If Not IsError(m) Then rowNum = CLng(m)
        On Error GoTo 0
    End If
    If rowNum = 0 Then
        GetActiveUniversePercentile = 0
        Exit Function
    End If

    Dim ownVal As Double: ownVal = CDbl(val(wsSrc.Range(metricCol & rowNum).value))
    If ownVal <= 0 Then
        GetActiveUniversePercentile = 0
        Exit Function
    End If

    Dim startRow As Long: startRow = CLng(val(wsSrc.Range("P1").value))
    Dim endRow As Long: endRow = CLng(val(wsSrc.Range("R1").value))
    If startRow <= 0 Or endRow < startRow Then Exit Function

    Dim countVals As Long: countVals = 0
    Dim countBelowOrEqual As Long: countBelowOrEqual = 0
    Dim r As Long, v As Double, sym As String
    For r = startRow To endRow
        sym = CStr(wsSrc.cells(r, "A").value)
        If Len(sym) > 0 And Not Is_Index(sym) Then
            v = CDbl(val(wsSrc.Range(metricCol & r).value))
            If v > 0 Then
                countVals = countVals + 1
                If v <= ownVal Then countBelowOrEqual = countBelowOrEqual + 1
            End If
        End If
    Next r

    If countVals > 0 Then GetActiveUniversePercentile = countBelowOrEqual / countVals
End Function


' =====================================================================
' Private Function  GetStableRegime
' Wraps GetStockRegime with a persistence filter. Only accepts a regime
' change after MIN_REGIME_STABLE_BARS consecutive identical readings.
' Uses module-level mRegimeHistory (persists across Run_Cockpit calls).
' =====================================================================
Private Function GetStableRegime(wsSrc As Worksheet, ByVal ticker As String) As String
    Dim rawRegime As String
    rawRegime = GetStockRegime(wsSrc, ticker)

    If mRegimeHistory Is Nothing Then
        Set mRegimeHistory = CreateObject("Scripting.Dictionary")
    End If

    If Len(rawRegime) = 0 Then
        GetStableRegime = rawRegime
        Exit Function
    End If

    If mRegimeHistory.Exists(ticker) Then
        Dim parts() As String: parts = Split(CStr(mRegimeHistory(ticker)), "|")
        Dim prevRegime As String: prevRegime = parts(0)
        Dim stableCount As Long: stableCount = CLng(parts(1))

        If rawRegime = prevRegime Then
            ' Same regime — increment counter
            stableCount = stableCount + 1
            mRegimeHistory(ticker) = rawRegime & "|" & stableCount
            GetStableRegime = rawRegime
        Else
            ' Regime changed — reset counter
            mRegimeHistory(ticker) = rawRegime & "|1"
            If stableCount >= MIN_REGIME_STABLE_BARS Then
                ' Old regime was confirmed stable — hold it until new one confirms
                GetStableRegime = prevRegime
            Else
                ' Old regime wasn't stable either — accept the new one
                GetStableRegime = rawRegime
            End If
        End If
    Else
        ' First reading for this ticker — accept as-is
        mRegimeHistory(ticker) = rawRegime & "|1"
        GetStableRegime = rawRegime
    End If
End Function


' =====================================================================
' Private Function  GetSectorRegime
' Reads sector name from Sheet1 col BM, then looks up sector ticker
' in col A and reads its regime from col BQ.
' Returns: "TRENDING_BULL", "TRENDING_BEAR", "MILD", "RANGING", or ""
' =====================================================================
Private Function GetSectorRegime(wsSrc As Worksheet, ByVal ticker As String) As String
    GetSectorRegime = ""
    If Len(ticker) = 0 Then Exit Function

    ' Get sector name from col BM
    Dim sectorName As String: sectorName = ""
    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        sectorName = CStr(wsSrc.cells(rowNum, "BM").value)
    Else
        On Error Resume Next
        Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
        If Not IsError(m) Then sectorName = CStr(wsSrc.cells(CLng(m), "BM").value)
        On Error GoTo 0
    End If

    If Len(sectorName) = 0 Then Exit Function

    ' Look up sector ticker in col A to get its 5m regime from BN
    If Not mRowLookup Is Nothing And mRowLookup.Exists(sectorName) Then
        GetSectorRegime = UCase(CStr(wsSrc.cells(CLng(mRowLookup(sectorName)), "BN").value))
        Exit Function
    End If
    On Error Resume Next
    m = Application.Match(sectorName, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetSectorRegime = UCase(CStr(wsSrc.cells(CLng(m), "BN").value))
    On Error GoTo 0
End Function


' =====================================================================
' Private Function  GetStockGrade
' Reads stock grade from Sheet1 col DE (SCORE_COL).
' Returns: "A+", "A", "B", "C", or ""
' =====================================================================
Private Function GetStockGrade(wsSrc As Worksheet, ByVal ticker As String) As String
    GetStockGrade = ""
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetStockGrade = UCase(CStr(wsSrc.cells(rowNum, "DE").value))
        Exit Function
    End If

    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetStockGrade = UCase(CStr(wsSrc.cells(CLng(m), "DE").value))
    On Error GoTo 0
End Function


' =====================================================================
' Private Function  GetBreakoutFlag
' Reads breakout flag from Sheet1 col DF (NARRATIVE_COL).
' Returns: 1 (Long breakout), -1 (Short breakout), 0 (none)
' =====================================================================
Private Function GetBreakoutFlag(wsSrc As Worksheet, ByVal ticker As String) As Long
    GetBreakoutFlag = 0
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetBreakoutFlag = CLng(val(wsSrc.cells(rowNum, "DF").value))
        Exit Function
    End If

    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetBreakoutFlag = CLng(val(wsSrc.cells(CLng(m), "DF").value))
    On Error GoTo 0
End Function


' =====================================================================
' Private Function  GetNiftyIsRanging
' Checks if NSEI (Nifty 50) regime is RANGING from Sheet1.
' Looks up NSEI row and reads col BQ (regime).
' =====================================================================
Private Function GetNiftyIsRanging(wsSrc As Worksheet) As Boolean
    GetNiftyIsRanging = False  ' fail-open: allow trades if data unavailable

    If mRowLookup Is Nothing Then Exit Function
    If Not mRowLookup.Exists("NSEI") Then Exit Function

    Dim niftyRow As Long: niftyRow = CLng(mRowLookup("NSEI"))
    If niftyRow <= 0 Then Exit Function

    Dim niftyRegime As String
    niftyRegime = UCase(CStr(wsSrc.cells(niftyRow, "BN").value))
    GetNiftyIsRanging = (niftyRegime = "RANGING")
End Function


' =====================================================================
' Private Function  CheckRSIExhaustion
' Multi-TF RSI exhaustion check (Strategies 6/7).
' Reads RSI from NI cols 13 (5m), 14 (15m), 15 (1h).
' Returns: "LONG", "SHORT", or ""
' =====================================================================
Private Function CheckRSIExhaustion(niData As Variant, r As Long) As String
    CheckRSIExhaustion = ""

    Dim rsi5m As Double:  rsi5m = SafeVal(niData(r, NI_RSI_5M))
    Dim rsi15m As Double: rsi15m = SafeVal(niData(r, NI_RSI_15M))
    Dim rsi1h As Double:  rsi1h = SafeVal(niData(r, NI_RSI_1H))

    If rsi5m <= RSI_EXH_5M_OS And rsi15m <= RSI_EXH_15M_OS And rsi1h <= RSI_EXH_1H_OS Then
        CheckRSIExhaustion = "LONG"
        Exit Function
    End If

    If rsi5m >= RSI_EXH_5M_OB And rsi15m >= RSI_EXH_15M_OB And rsi1h >= RSI_EXH_1H_OB Then
        CheckRSIExhaustion = "SHORT"
        Exit Function
    End If
End Function

' =====================================================================
' Private Sub  WriteNewCockpitRow
' Copies NI data + computes Ent/Tgt/SL into a fresh COCKPIT row.
' Uses mRowLookup for O(1) Sheet1 queries.
' =====================================================================
Private Sub WriteNewCockpitRow(wsC As Worksheet, niData As Variant, niIdx As Long, destRow As Long, _
                               wsSrc As Worksheet, isBuy As Boolean, stratName As String, _
                               sizeVal As Double, slMult As Double, tgtPct As Double, _
                               stockRegime As String, stockGrade As String, secRegime As String)

    Dim ticker As String: ticker = SafeStr(niData(niIdx, NI_TICKER))
    Dim ltp As Double: ltp = SafeVal(niData(niIdx, NI_LTP))
    
    ' Core Identifiers
    wsC.Range(CK_TICKER & destRow).value = ticker
    wsC.Range(CK_ACTION & destRow).value = IIf(isBuy, "BUY", "SELL")
    wsC.Range(CK_LTP & destRow).value = ltp
    wsC.Range(CK_SECTOR & destRow).value = SafeStr(niData(niIdx, NI_SECTOR))
    
    ' Populate S/R Levels
    wsC.Range(CK_S3 & destRow).value = SafeVal(niData(niIdx, NI_S3))
    wsC.Range(CK_S2 & destRow).value = SafeVal(niData(niIdx, NI_S2))
    wsC.Range(CK_S1 & destRow).value = SafeVal(niData(niIdx, NI_S1))
    wsC.Range(CK_PP & destRow).value = SafeVal(niData(niIdx, NI_PP))
    wsC.Range(CK_R1 & destRow).value = SafeVal(niData(niIdx, NI_R1))
    wsC.Range(CK_R2 & destRow).value = SafeVal(niData(niIdx, NI_R2))
    wsC.Range(CK_R3 & destRow).value = SafeVal(niData(niIdx, NI_R3))
    
    ' Populate Technical Indicators
    wsC.Range(CK_RSI_5M & destRow).value = SafeVal(niData(niIdx, NI_RSI_5M))
    wsC.Range(CK_RSI_15M & destRow).value = SafeVal(niData(niIdx, NI_RSI_15M))
    wsC.Range(CK_RSI_1H & destRow).value = SafeVal(niData(niIdx, NI_RSI_1H))
    wsC.Range(CK_RSI_1D & destRow).value = SafeVal(niData(niIdx, NI_RSI_1D))
    wsC.Range(CK_SUPERTREND & destRow).value = SafeStr(niData(niIdx, NI_SUPERTREND))
    wsC.Range(CK_CROSS_10_30 & destRow).value = SafeStr(niData(niIdx, NI_CROSS_10_30))
    wsC.Range(CK_CROSS_10_50 & destRow).value = SafeStr(niData(niIdx, NI_CROSS_10_50))

    ' Execution Data
    wsC.Range(CK_OPEN_TIME & destRow).value = Format(Now, "YYYY-MM-DD HH:MM:SS")
    wsC.Range(CK_ENT & destRow).value = ltp
    wsC.Range(CK_CP & destRow).value = ltp
    wsC.Range(CK_STRATEGY & destRow).value = stratName
    
    Dim diagCtx As String
    diagCtx = "Reg=" & stockRegime & " Grd=" & stockGrade & " Sec=" & secRegime & " Size=" & sizeVal
    wsC.Range(CK_COMENTS & destRow).value = diagCtx
    
    ' Target and Stop Loss Calculation
    Dim atr As Double: atr = GetATR5m(wsSrc, ticker)
    Dim slPx As Double, tgtPx As Double
    
    If isBuy Then
        slPx = ltp - IIf(atr > 0, atr * slMult, ltp * 0.005)
        tgtPx = ltp * (1 + tgtPct)
    Else
        slPx = ltp + IIf(atr > 0, atr * slMult, ltp * 0.005)
        tgtPx = ltp * (1 - tgtPct)
    End If
    
    wsC.Range(CK_SL & destRow).value = slPx
    wsC.Range(CK_TGT & destRow).value = tgtPx

    ' Initialize hidden dynamic exit tracking columns (set 0 so first trade management cycle becomes 1)
    wsC.Range(CK_CANDLE_COUNT & destRow).value = 0
    wsC.Range(CK_PEAK_PRICE & destRow).value = ltp
    wsC.Range(CK_BEST_MFE_PCT & destRow).value = 0
    wsC.Range(CK_LAST_MFE_BAR & destRow).value = 0

    ' Formatting
    ApplyCockpitColors wsC, destRow, isBuy, ltp, _
                       SafeVal(niData(niIdx, NI_S3)), SafeVal(niData(niIdx, NI_S2)), SafeVal(niData(niIdx, NI_S1)), _
                       SafeVal(niData(niIdx, NI_PP)), SafeVal(niData(niIdx, NI_R1)), SafeVal(niData(niIdx, NI_R2)), SafeVal(niData(niIdx, NI_R3)), _
                       SafeVal(niData(niIdx, NI_RSI_5M)), SafeVal(niData(niIdx, NI_RSI_15M)), SafeVal(niData(niIdx, NI_RSI_1H)), SafeVal(niData(niIdx, NI_RSI_1D)), _
                       SafeStr(niData(niIdx, NI_SUPERTREND)), SafeStr(niData(niIdx, NI_CROSS_10_30)), SafeStr(niData(niIdx, NI_CROSS_10_50))
End Sub
' =====================================================================
' FUNCTION: SafeGetNum (Complete)
' Prevents Type Mismatch crashes if source cells contain Excel errors.
' =====================================================================
Private Function SafeGetNum(ws As Worksheet, rowNum As Long, colName As String, fallback As Double) As Double
    Dim cellVal As Variant
    cellVal = ws.cells(rowNum, colName).value
    If IsError(cellVal) Then
        SafeGetNum = fallback
    ElseIf IsNumeric(cellVal) And Len(CStr(cellVal)) > 0 Then
        SafeGetNum = CDbl(cellVal)
    Else
        SafeGetNum = fallback
    End If
End Function

' =====================================================================
' Private Sub  ApplyCockpitColors
' Direct cell formatting — mirrors New_Logic_Intraday coloring.
' Zero clipboard usage.  All writes are Interior.Color / Font properties.
' =====================================================================
Private Sub ApplyCockpitColors(wsC As Worksheet, destRow As Long, _
        ByVal isBuy As Boolean, ByVal ltp As Double, _
        ByVal s3 As Double, ByVal s2 As Double, ByVal s1 As Double, _
        ByVal pp As Double, ByVal r1 As Double, ByVal r2 As Double, ByVal r3 As Double, _
        ByVal rsi5 As Double, ByVal rsi15 As Double, ByVal rsi1h As Double, ByVal rsi1d As Double, _
        ByVal superTrend As String, ByVal cross1030 As String, ByVal cross1050 As String)

    Dim cells As Range

    ' --- Ticker font (A) ---
    Set cells = wsC.Range(CK_TICKER & destRow)
    If isBuy Then
        cells.Font.Color = RGB(0, 150, 0)
    Else
        cells.Font.Color = RGB(180, 0, 0)
    End If

    ' --- Action (B) ---
    Set cells = wsC.Range(CK_ACTION & destRow)
    If isBuy Then
        cells.Interior.Color = RGB(204, 255, 204)
        cells.Font.Bold = True
    Else
        cells.Interior.Color = RGB(255, 204, 204)
        cells.Font.Bold = True
    End If

    ' --- LTP vs 10EMA (C) ---
    ' NI colors LTP green if LTP >= 10EMA, red if LTP < 10EMA.
    ' 10EMA is in NI col D (index 4) but not written to COCKPIT directly.
    ' Use the NI LTP value for self-reference — skip LTP coloring here
    ' since we don't have 10EMA available in the current scope.
    ' (LTP cell remains unformatted — matches NI only when 10EMA accessible.)

    ' --- Pivot levels (E-K): highlight critical level ---
    Dim pivots(1 To 7) As Double
    pivots(1) = s3: pivots(2) = s2: pivots(3) = s1: pivots(4) = pp
    pivots(5) = r1: pivots(6) = r2: pivots(7) = r3
    Dim pivotCols As Variant
    pivotCols = Array(CK_S3, CK_S2, CK_S1, CK_PP, CK_R1, CK_R2, CK_R3)

    If isBuy Then
        ' Find first pivot >= LTP scanning left to right (S2 upward)
        Dim jBuy As Long: jBuy = 0
        Dim j As Long
        For j = 2 To 7
            If pivots(j) >= ltp And pivots(j) > 0 Then
                jBuy = j: Exit For
            End If
        Next j
        If jBuy > 0 Then
            ' The level just crossed / being tested
            wsC.Range(pivotCols(jBuy - 1) & destRow).Interior.Color = RGB(255, 255, 160)  ' yellow
            ' Level just broken (support)
            If jBuy >= 2 Then
                wsC.Range(pivotCols(jBuy - 2) & destRow).Interior.Color = RGB(204, 255, 204)  ' green
            End If
            ' Level below (further support)
            If jBuy >= 3 Then
                wsC.Range(pivotCols(jBuy - 3) & destRow).Interior.Color = RGB(255, 204, 204)  ' red
            End If
        End If
    Else
        ' Find first pivot <= LTP scanning right to left (R3 downward)
        Dim jSell As Long: jSell = 0
        For j = 7 To 2 Step -1
            If pivots(j) <= ltp And pivots(j) > 0 Then
                jSell = j: Exit For
            End If
        Next j
        If jSell > 0 Then
            wsC.Range(pivotCols(jSell - 1) & destRow).Interior.Color = RGB(255, 255, 160)
            ' Level just above (resistance)
            If jSell <= 6 Then
                wsC.Range(pivotCols(jSell) & destRow).Interior.Color = RGB(204, 255, 204)
            End If
            ' Level above that
            If jSell <= 5 Then
                wsC.Range(pivotCols(jSell + 1) & destRow).Interior.Color = RGB(255, 204, 204)
            End If
        End If
    End If

    ' --- RSI (L-O): threshold coloring ---
    Dim rsiVals(1 To 4) As Double
    rsiVals(1) = rsi5: rsiVals(2) = rsi15: rsiVals(3) = rsi1h: rsiVals(4) = rsi1d
    Dim rsiCols As Variant
    rsiCols = Array(CK_RSI_5M, CK_RSI_15M, CK_RSI_1H, CK_RSI_1D)
    Dim i As Long
    For i = 1 To 4
        If rsiVals(i) >= 80 Then
            wsC.Range(rsiCols(i - 1) & destRow).Interior.Color = RGB(210, 180, 140)  ' tan
        ElseIf (rsiVals(i) >= 70 And rsiVals(i) < 80) Or _
               (rsiVals(i) <= 30 And rsiVals(i) >= 21) Then
            wsC.Range(rsiCols(i - 1) & destRow).Interior.Color = RGB(255, 255, 160)  ' yellow
        End If
    Next i

    ' --- Supertrend (P) ---
    Set cells = wsC.Range(CK_SUPERTREND & destRow)
    If LCase(superTrend) = "bullish" Then
        cells.Interior.Color = RGB(204, 255, 204)
        cells.Font.Bold = True
    ElseIf LCase(superTrend) = "bearish" Then
        cells.Interior.Color = RGB(255, 204, 204)
        cells.Font.Bold = True
    End If

    ' --- 10 Vs 30 (Q) ---
    Set cells = wsC.Range(CK_CROSS_10_30 & destRow)
    If LCase(cross1030) = "over" Then
        cells.Interior.Color = RGB(204, 255, 204): cells.Font.Bold = True
    ElseIf LCase(cross1030) = "under" Then
        cells.Interior.Color = RGB(255, 204, 204): cells.Font.Bold = True
    End If

    ' --- 10 vs 50 (R) ---
    Set cells = wsC.Range(CK_CROSS_10_50 & destRow)
    If LCase(cross1050) = "over" Then
        cells.Interior.Color = RGB(204, 255, 204): cells.Font.Bold = True
    ElseIf LCase(cross1050) = "under" Then
        cells.Interior.Color = RGB(255, 204, 204): cells.Font.Bold = True
    End If
End Sub





' =====================================================================
' Private Function  GetEMA30
' Looks up EMA 30 (5m) from Sheet1 column AT using module-level row cache.
' Falls back to Application.Match if cache miss.
' =====================================================================
Private Function GetEMA30(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetEMA30 = 0
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetEMA30 = CDbl(val(wsSrc.cells(rowNum, "AT").value))
        Exit Function
    End If

    ' Fallback: Match (should not happen if lookup is built correctly)
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetEMA30 = CDbl(wsSrc.cells(CLng(m), "AT").value)
    On Error GoTo 0
End Function


' =====================================================================
' Private Function  GetCurrentLTP
' Returns current LTP from Sheet1 col T using module-level row cache.
' =====================================================================
Private Function GetCurrentLTP(wsSrc As Worksheet, ByVal ticker As String, _
        ByVal fallback As Double) As Double

    GetCurrentLTP = fallback
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        Dim v As Double: v = CDbl(val(wsSrc.cells(rowNum, "T").value))
        If v > 0 Then GetCurrentLTP = v
        Exit Function
    End If

    ' Fallback: Match
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then
        v = CDbl(wsSrc.cells(CLng(m), "T").value)
        If v > 0 Then GetCurrentLTP = v
    End If
    On Error GoTo 0
End Function


' =====================================================================
' Private Function  GetADX
' Returns ADX from Sheet1 column CB using module-level row cache.
' Defaults to 99 (fail-open) if lookup fails — the trade is allowed.
' =====================================================================
Private Function GetADX(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetADX = 99    ' fail-open: if we can't find ADX, allow the trade
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetADX = CDbl(val(wsSrc.cells(rowNum, "CB").value))
        Exit Function
    End If

    ' Fallback: Match
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetADX = CDbl(wsSrc.cells(CLng(m), "CB").value)
    On Error GoTo 0
End Function


' =====================================================================
' Private Function  GetATR5m
' Looks up ATR_14_5m from Sheet1 column CC using module-level row cache.
' =====================================================================
Private Function GetATR5m(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetATR5m = 0
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetATR5m = SafeGetNum(wsSrc, rowNum, "CC", 0)
        Exit Function
    End If

    ' Fallback: Match
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetATR5m = SafeGetNum(wsSrc, CLng(m), "CC", 0)
    On Error GoTo 0
End Function



' =====================================================================
' Private Function  GetVolumeRatio
' Returns bar volume / 20-bar SMA from Sheet1 column BO (vol_ratio).
' Defaults to 1.0 (fail-open) if lookup fails.
' =====================================================================
Private Function GetVolumeRatio(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetVolumeRatio = 1#     ' fail-open: allow trade if data unavailable
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetVolumeRatio = CDbl(val(wsSrc.cells(rowNum, "BO").value))
        If GetVolumeRatio <= 0 Then GetVolumeRatio = 1#
        Exit Function
    End If

    ' Fallback: Match
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then
        GetVolumeRatio = CDbl(wsSrc.cells(CLng(m), "BO").value)
        If GetVolumeRatio <= 0 Then GetVolumeRatio = 1#
    End If
    On Error GoTo 0
End Function



' =====================================================================
' Private Function  GetDistDayHi
' Returns distance from daily high in ATR units (Sheet1 col DG).
' Defaults to 99 (fail-open) if lookup fails.
' =====================================================================
Private Function GetDistDayHi(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetDistDayHi = 99   ' fail-open: allow trade if data unavailable
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetDistDayHi = CDbl(val(wsSrc.cells(rowNum, "DG").value))
        If GetDistDayHi <= 0 Then GetDistDayHi = 99
        Exit Function
    End If

    ' Fallback: Match
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then
        GetDistDayHi = CDbl(wsSrc.cells(CLng(m), "DG").value)
        If GetDistDayHi <= 0 Then GetDistDayHi = 99
    End If
    On Error GoTo 0
End Function


' =====================================================================
' Private Function  GetDistDayLo
' Returns distance from daily low in ATR units (Sheet1 col DH).
' Defaults to 99 (fail-open) if lookup fails.
' =====================================================================
Private Function GetDistDayLo(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetDistDayLo = 99   ' fail-open: allow trade if data unavailable
    If Len(ticker) = 0 Then Exit Function

    Dim rowNum As Long: rowNum = 0
    If Not mRowLookup Is Nothing Then
        If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
    End If
    If rowNum > 0 Then
        GetDistDayLo = CDbl(val(wsSrc.cells(rowNum, "DH").value))
        If GetDistDayLo <= 0 Then GetDistDayLo = 99
        Exit Function
    End If

    ' Fallback: Match
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then
        GetDistDayLo = CDbl(wsSrc.cells(CLng(m), "DH").value)
        If GetDistDayLo <= 0 Then GetDistDayLo = 99
    End If
    On Error GoTo 0
End Function



' =====================================================================
' Private Function  IsOnCooldown
' Returns True if the ticker has been closed within COOLDOWN_MINUTES_VAL.
' Uses a static Dictionary persisting for the Excel session.
' =====================================================================
Private Function IsOnCooldown(ByVal ticker As String) As Boolean
    IsOnCooldown = False
    If COOLDOWN_MINUTES_VAL <= 0 Then Exit Function

    If mCooldowns Is Nothing Then Set mCooldowns = CreateObject("Scripting.Dictionary")

    If mCooldowns.Exists(ticker) Then
        Dim lastCloseTime As Date
        lastCloseTime = CDate(mCooldowns(ticker))
        If DateDiff("n", lastCloseTime, Now) < COOLDOWN_MINUTES_VAL Then
            IsOnCooldown = True
        Else
            mCooldowns.Remove ticker  ' Clean up expired
        End If
    End If
End Function


' =====================================================================
' Private Sub  RecordCooldown
' Records the current time as the last close time for the ticker.
' =====================================================================
Private Sub RecordCooldown(ByVal ticker As String)
    If COOLDOWN_MINUTES_VAL <= 0 Then Exit Sub

    If mCooldowns Is Nothing Then Set mCooldowns = CreateObject("Scripting.Dictionary")

    mCooldowns(ticker) = Now
End Sub


' =====================================================================
' Private Sub  ManageCockpitTrades
' Scans existing COCKPIT rows.
'   Active (X empty) ? update V(CP), check exit conditions
'   Closed (X has value) ? skip
' =====================================================================
Private Sub ManageCockpitTrades(wsC As Worksheet, wsSrc As Worksheet)
    Dim lastRow As Long
    lastRow = wsC.cells(wsC.Rows.count, 1).End(xlUp).row
    If lastRow < DATA_START_ROW Then Exit Sub

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        ' -- Skip closed trades (X has a value) --
        Dim closeTime As Variant
        closeTime = wsC.Range(CK_CLOSE_TIME & r).value
        If Not IsEmpty(closeTime) And Len(SafeStr(closeTime)) > 0 Then GoTo NextTrade

        ' -- Trade is open --
        Dim ticker As String:     ticker = SafeStr(wsC.Range(CK_TICKER & r).value)
        Dim entryPx As Double:    entryPx = SafeVal(wsC.Range(CK_ENT & r).value)
        Dim tgtPx As Double:      tgtPx = SafeVal(wsC.Range(CK_TGT & r).value)
        Dim slPx As Double:       slPx = SafeVal(wsC.Range(CK_SL & r).value)
        Dim openTime As Variant:  openTime = wsC.Range(CK_OPEN_TIME & r).value
        Dim action As String:     action = UCase$(SafeStr(wsC.Range(CK_ACTION & r).value))
        Dim stratName As String:  stratName = SafeStr(wsC.Range(CK_STRATEGY & r).value)

        If Len(ticker) = 0 Or entryPx <= 0 Then GoTo NextTrade

        ' -- Current LTP --
        Dim curLTP As Double: curLTP = GetCurrentLTP(wsSrc, ticker, entryPx)
        wsC.Range(CK_CP & r).value = curLTP

        ' -- Update Candle Count --
        Dim candleCount As Long
        candleCount = SafeVal(wsC.Range(CK_CANDLE_COUNT & r).value)
        candleCount = candleCount + 1
        wsC.Range(CK_CANDLE_COUNT & r).value = candleCount

        ' -- Track favorable excursion --
        Dim curMfePct As Double
        If action = "BUY" Then
            curMfePct = (curLTP - entryPx) / entryPx
        Else
            curMfePct = (entryPx - curLTP) / entryPx
        End If

        Dim bestMfePct As Double: bestMfePct = SafeVal(wsC.Range(CK_BEST_MFE_PCT & r).value)
        Dim peakPx As Double:     peakPx = SafeVal(wsC.Range(CK_PEAK_PRICE & r).value)
        Dim lastMfeBar As Long:   lastMfeBar = SafeVal(wsC.Range(CK_LAST_MFE_BAR & r).value)

        If peakPx <= 0 Then
            peakPx = entryPx
            wsC.Range(CK_PEAK_PRICE & r).value = peakPx
        End If

        If curMfePct > bestMfePct Then
            bestMfePct = curMfePct
            peakPx = curLTP
            lastMfeBar = candleCount
            wsC.Range(CK_BEST_MFE_PCT & r).value = bestMfePct
            wsC.Range(CK_PEAK_PRICE & r).value = peakPx
            wsC.Range(CK_LAST_MFE_BAR & r).value = lastMfeBar
        End If

        ' -- Breakeven threshold calculation --
        Dim atrBE As Double: atrBE = GetATR5m(wsSrc, ticker)
        Dim beThreshold As Double
        If atrBE > 0 And entryPx > 0 Then
            beThreshold = (atrBE * TRAILING_BE_ATR_MULT) / entryPx
        Else
            beThreshold = 0.002
        End If

        ' -- SL Sanity Correction: Only runs if SL is uninitialized or invalid BEFORE breakeven is locked --
        Dim isInvalidSL As Boolean: isInvalidSL = False
        If slPx <= 0 Then
            isInvalidSL = True
        ElseIf action = "BUY" And slPx >= entryPx And bestMfePct < beThreshold Then
            isInvalidSL = True
        ElseIf action = "SELL" And slPx <= entryPx And bestMfePct < beThreshold Then
            isInvalidSL = True
        End If

        If isInvalidSL Then
            Dim atrFallback As Double: atrFallback = GetATR5m(wsSrc, ticker)
            If atrFallback > 0 Then
                Dim atrMult As Double: atrMult = 1.5
                If stratName = "RSI_Exh_Long" Or stratName = "RSI_Exh_Short" Then
                    atrMult = SL_ATR_RSI_EXH
                ElseIf stratName = "Breakout_Long" Or stratName = "Breakout_Short" Then
                    atrMult = SL_ATR_BREAKOUT
                ElseIf IsMildMomentumStrategy(stratName) Then
                    atrMult = SL_ATR_MILD
                Else
                    atrMult = SL_ATR_MOM_BULL
                End If
                If action = "BUY" Then
                    slPx = entryPx - (atrFallback * atrMult)
                Else
                    slPx = entryPx + (atrFallback * atrMult)
                End If
                wsC.Range(CK_SL & r).value = slPx
            End If
        End If

        Dim targetDistPct As Double: targetDistPct = 0
        If tgtPx > 0 Then targetDistPct = Abs(tgtPx - entryPx) / entryPx

        Dim profitArmPct As Double: profitArmPct = PROFIT_PROTECT_MIN_PCT
        If targetDistPct * PROFIT_PROTECT_TGT_FRAC > profitArmPct Then
            profitArmPct = targetDistPct * PROFIT_PROTECT_TGT_FRAC
        End If

        Dim profitArmed As Boolean: profitArmed = (bestMfePct >= profitArmPct)

        ' -- Exit Evaluation --
        Dim exitFlag As Boolean: exitFlag = False
        Dim exitReason As String

        ' 1. Session Liquidation
        Dim currentTime As Date
        currentTime = TimeValue(Now)
        If currentTime >= TimeValue(FORCE_LIQUIDATION) Then
            exitFlag = True: exitReason = "Session Liquidation (" & FORCE_LIQUIDATION & ")"
        End If

        ' 2. Standard SL / Target
        If Not exitFlag Then
            If action = "BUY" Then
                If slPx > 0 And curLTP <= slPx Then
                    exitFlag = True: exitReason = "SL Hit"
                ElseIf tgtPx > 0 And curLTP >= tgtPx Then
                    exitFlag = True: exitReason = "Target Hit"
                End If
            Else
                If slPx > 0 And curLTP >= slPx Then
                    exitFlag = True: exitReason = "SL Hit"
                ElseIf tgtPx > 0 And curLTP <= tgtPx Then
                    exitFlag = True: exitReason = "Target Hit"
                End If
            End If
        End If

        ' 2a. Trailing Breakeven Lock
        If Not exitFlag Then
            If curMfePct >= beThreshold Then
                Dim bePx As Double
                If action = "BUY" Then
                    bePx = entryPx * 1.0005
                    If slPx < bePx Then
                        slPx = bePx
                        wsC.Range(CK_SL & r).value = slPx
                    End If
                Else
                    bePx = entryPx * 0.9995
                    If slPx > bePx Then
                        slPx = bePx
                        wsC.Range(CK_SL & r).value = slPx
                    End If
                End If
            End If
        End If

        ' 2b. Profit Protection
        If Not exitFlag And profitArmed Then
            Dim givebackPct As Double: givebackPct = bestMfePct - curMfePct
            Dim retraceTriggerPct As Double
            retraceTriggerPct = bestMfePct * PROFIT_RETRACE_FRAC
            If retraceTriggerPct < PROFIT_RETRACE_MIN_PCT Then retraceTriggerPct = PROFIT_RETRACE_MIN_PCT

            If curMfePct > 0 And givebackPct >= retraceTriggerPct Then
                exitFlag = True
                exitReason = "Profit Reversal (MFE=" & Format(bestMfePct * 100, "0.00") & "%, now=" & Format(curMfePct * 100, "0.00") & "%)"
            ElseIf curMfePct > 0 And (candleCount - lastMfeBar) >= PROFIT_STALL_BARS Then
                exitFlag = True
                exitReason = "Profit Stall (MFE=" & Format(bestMfePct * 100, "0.00") & "%, now=" & Format(curMfePct * 100, "0.00") & "%)"
            End If
        End If

        ' 2c. Dead Trade Cut
        If Not exitFlag And candleCount >= STAGNATION_BARS Then
            If bestMfePct < STAGNATION_MIN_MFE_PCT And curMfePct <= DEAD_TRADE_EXIT_MAX_PCT Then
                exitFlag = True: exitReason = "Dead Trade (" & STAGNATION_BARS & " bars, MFE=" & Format(bestMfePct * 100, "0.00") & "%, now=" & Format(curMfePct * 100, "0.00") & "%)"
            End If
        End If

        ' 3. RSI Exhaustion exit
        If Not exitFlag And (stratName = "RSI_Exh_Long" Or stratName = "RSI_Exh_Short") Then
            Dim rsi5mExit As Double: rsi5mExit = 0
            Dim rowNum As Long: rowNum = 0
            If Not mRowLookup Is Nothing Then
                If mRowLookup.Exists(ticker) Then rowNum = CLng(mRowLookup(ticker))
            End If
            If rowNum > 0 Then
                rsi5mExit = SafeVal(wsSrc.cells(rowNum, "M").value)
            End If
            If rsi5mExit > 0 Then
                If action = "BUY" And rsi5mExit >= 50 Then
                    exitFlag = True: exitReason = "RSI Centerline Cross (5m >= 50)"
                ElseIf action = "SELL" And rsi5mExit <= 50 Then
                    exitFlag = True: exitReason = "RSI Centerline Cross (5m <= 50)"
                End If
            End If
        End If

        ' 4. Expiry Exit
        If Not exitFlag And candleCount >= MAX_CANDLES_EXIT Then
            exitFlag = True: exitReason = "Expired (" & MAX_CANDLES_EXIT & " candles)"
        End If

        ' -- Execute exit --
        If exitFlag Then
            wsC.Range(CK_CP & r).value = curLTP
            wsC.Range(CK_CLOSE_TIME & r).value = Format(Now, "YYYY-MM-DD HH:MM:SS")
            Dim prevComments As String: prevComments = SafeStr(wsC.Range(CK_COMENTS & r).value)
            wsC.Range(CK_COMENTS & r).value = prevComments & " -> " & exitReason
            RecordCooldown ticker
        End If

NextTrade:
    Next r
End Sub


' =====================================================================
' Private Function  RangeRef
' Converts a column letter string to its 1-based integer column index.
' =====================================================================
Private Function RangeRef(ByVal letter As String) As Long
    RangeRef = Range(letter & "1").Column
End Function


' =====================================================================
' Private Sub  LogDebugReject
' Writes rejected ticker + reason to "Debug" sheet (not COCKPIT).
' Creates the sheet if it doesn't exist.
' =====================================================================
Private Sub LogDebugReject(ByVal ticker As String, ByVal reason As String)
    On Error Resume Next
    Dim wsDbg As Worksheet
    Set wsDbg = ThisWorkbook.Sheets("Debug")
    If wsDbg Is Nothing Then
        Set wsDbg = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsDbg.Name = "Debug"
        wsDbg.cells(1, 1) = "Ticker"
        wsDbg.cells(1, 2) = "Reason"
        wsDbg.cells(1, 3) = "Time"
    End If
    Dim nxt As Long
    nxt = wsDbg.cells(wsDbg.Rows.count, 1).End(xlUp).row + 1
    wsDbg.cells(nxt, 1) = ticker
    wsDbg.cells(nxt, 2) = reason
    wsDbg.cells(nxt, 3) = Format(Now, "HH:MM:SS")
    On Error GoTo 0
End Sub


' =====================================================================
' HELPER FUNCTIONS: SafeStr / SafeVal
' Prevents Type Mismatch (Error 13) crashes on sheet formula errors (#N/A, #VALUE!)
' =====================================================================
Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim$(CStr(v))
    End If
End Function

Private Function SafeVal(ByVal v As Variant, Optional ByVal fallback As Double = 0#) As Double
    If IsError(v) Or IsNull(v) Then
        SafeVal = fallback
    ElseIf IsNumeric(v) And Len(CStr(v)) > 0 Then
        SafeVal = CDbl(v)
    Else
        SafeVal = fallback
    End If
End Function



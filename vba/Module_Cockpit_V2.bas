Attribute VB_Name = "Module_Cockpit_V2"
'Attribute VB_Name = "Module_COCKPIT_V2"
Option Explicit

'----------------------------------------------------------------------
' Module_COCKPIT_V2  --  BTB V2.00 (Experimental Donchian+RVOL Discovery)
'
' Independent experimental engine running alongside V1 baseline.
' Target Worksheet: "COCKPIT_V2"
' Master Discovery: Donchian 5-bar Breakout + RVOL >= 1.25 + EMA10>EMA30
' Data Provenance: Reads 100% directly from "Sheet1" (Zero NI dependency)
' Downstream Parity: 100% downstream gates, exits, risk rules, and color coding match V1.
' Includes: Run_Cockpit_Compare with intuitive diagnostic reasons.
'----------------------------------------------------------------------

' =====================================================================
'  DISCOVERY MODULAR CONTROLLER
' =====================================================================
Public Const DISCOVERY_MODE             As String = "DONCHIAN_RVOL"   ' Options: "DONCHIAN_RVOL", "EMA_CROSS"

' =====================================================================
'  MASTER TOGGLES
' =====================================================================
Private Const DEBUG_MODE                As Boolean = False
Private Const ENABLE_BREAKOUT           As Boolean = True
Private Const ENABLE_MOMENTUM           As Boolean = True

' =====================================================================
'  BREAKOUT STRATEGY GATES
' =====================================================================
Private Const GATE_BO_VOL_ENABLED       As Boolean = True
Private Const GATE_BO_VOL_MIN           As Double = 1.25
Private Const GATE_MOM_ADX_MIN          As Double = 20#
Private Const GATE_MOM_VOL_MIN          As Double = 0.5
Private Const GATE_MOM_ER10_MIN         As Double = 0.25
Private Const BREAKOUT_N                As Long = 5

' RSI context gate: qualifies already-valid Donchian/Momentum entries.
Private Const GATE_RSI_STRONG_ER10_MIN  As Double = 0.35
Private Const GATE_RSI_STRONG_VOL_MIN   As Double = 1.75

' =====================================================================
'  DYNAMIC RISK & LIFECYCLE EXITS
' =====================================================================
Private Const SL_ATR_MOM_BULL           As Double = 1.5
Private Const SL_ATR_MOM_BEAR           As Double = 1.5
Private Const SL_ATR_BREAKOUT           As Double = 1.5

Private Const TARGET_PCT_MOM_BULL       As Double = 0.015
Private Const TARGET_PCT_MOM_BEAR       As Double = 0.015
Private Const TARGET_PCT_BREAKOUT       As Double = 0.015

Private Const SIZE_MOM_STRONG           As Double = 1#
Private Const SIZE_MOM_BEAR             As Double = 0.5
Private Const SIZE_BREAKOUT             As Double = 0.5

Private Const LATE_ENTRY_CUTOFF         As String = "15:00:00"
Private Const FORCE_LIQUIDATION         As String = "15:00:00"
Private Const MAX_CANDLES_EXIT          As Long = 9
Private Const COOLDOWN_MINUTES_VAL      As Long = 30
Private Const DATA_START_ROW            As Long = 5

Private Const STAGNATION_BARS           As Long = 5
Private Const STAGNATION_MIN_MFE_PCT    As Double = 0.0005
Private Const DEAD_TRADE_EXIT_MAX_PCT   As Double = -0.003
Private Const TRAILING_BE_ATR_MULT      As Double = 1#

' =====================================================================
'  ARCHITECTURAL CONTROLS (Approved V2 Production Upgrades)
' =====================================================================
Private Const MAX_ATTEMPTS_PER_TICKER    As Long = 2
Private Const ENABLE_CAT_DECAY_EXIT      As Boolean = True
Private Const RISK_FLOOR_MIN_PCT         As Double = 0.0035   ' 0.35% Minimum Stop Risk Floor

' --- COCKPIT_V2 COLUMN LETTERS --
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
Private Const CK_OUTCOME      As String = "Y"
Private Const CK_BOTTOMLINE   As String = "Z"
Private Const CK_COMENTS      As String = "AA"
Private Const CK_CANDLE_COUNT As String = "AB"
Private Const CK_STRATEGY     As String = "AC"
Private Const CK_PEAK_PRICE   As String = "AD"
Private Const CK_BEST_MFE_PCT As String = "AE"
Private Const CK_LAST_MFE_BAR As String = "AF"
Private Const CK_INITIAL_RISK As String = "AG"
Private Const CK_INITIAL_RISK_PCT As String = "AH"
Private Const CK_ENTRY_ATR    As String = "AI"
Private Const CK_STOP_DISTANCE As String = "AJ"
Private Const CK_POSITION_SIZE As String = "AK"
Private Const CK_PNL_R        As String = "AL"
Private Const CK_BEST_MFE_R   As String = "AM"
Private Const CK_WORST_MAE_PCT As String = "AN"
Private Const CK_WORST_MAE_R  As String = "AO"
Private Const CK_MFE_1_R      As String = "AP"
Private Const CK_MFE_2_R      As String = "AQ"
Private Const CK_MFE_3_R      As String = "AR"
Private Const CK_MFE_5_R      As String = "AS"
Private Const CK_BARS_TO_MFE  As String = "AT"
Private Const CK_LAST_MAE_BAR As String = "AU"
Private Const CK_INITIAL_CAT      As String = "AV"
Private Const CK_INITIAL_SCORE    As String = "AW"
Private Const CK_CURRENT_CAT      As String = "AX"
Private Const CK_CURRENT_SCORE    As String = "AY"
Private Const CK_CAT_TRANSITIONS  As String = "AZ"
Private Const CK_SCORE_CHANGES    As String = "BA"
Private Const CK_OVERRIDE_REASON  As String = "BB"

' Module-level lookup cache
Private mRowLookup As Object
Private mCooldowns As Object
Private mRejections As Object   ' Ticker -> Diagnostic Rejection Reason
Private mAuditRows As Collection ' Audit-only V2 funnel details for this run
Private mActiveTrades As Object ' Ticker|Dir -> True (O(1) Active Duplicate Cache)
Private mTickerAttempts As Object ' Ticker -> Long (O(1) Session Attempt Counter)
Private mRunTimestamp As String

' =====================================================================
' Public Sub  Run_Cockpit_V2
' Main Entry Point for COCKPIT_V2 Experimental Engine
' Reads 100% directly from Sheet1 (Zero NI dependency)
' =====================================================================
Public Sub Run_Cockpit_V2()
    Dim currentTime As Date
    currentTime = TimeValue(Now)

    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsSrc  As Worksheet: Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Dim wsC2   As Worksheet: Set wsC2 = ThisWorkbook.Sheets("COCKPIT_V2")

    EnsureCockpitV2StateHeaders wsC2
    Set mRowLookup = BuildRowLookup(wsSrc)
    Set mRejections = CreateObject("Scripting.Dictionary")
    Set mAuditRows = New Collection
    Set mActiveTrades = BuildActiveTradeLookup(wsC2)
    Set mTickerAttempts = BuildTickerSessionAttemptLookup(wsC2)
    mRunTimestamp = Format(Now, "YYYY-MM-DD HH:MM:SS")

    ' -- TIME GATE --
    If currentTime >= TimeValue(LATE_ENTRY_CUTOFF) Then
        RefreshOpenTradeConviction_V2 wsC2, wsSrc
        ManageCockpitTrades_V2 wsC2, wsSrc
        GoTo Cleanup
    End If

    ' -- Find next empty row in COCKPIT_V2 --
    Dim lastDestRow As Long
    lastDestRow = wsC2.cells(wsC2.Rows.count, 1).End(xlUp).row
    If lastDestRow < DATA_START_ROW Then lastDestRow = DATA_START_ROW - 1

    ' ===== SCAN SHEET1 ROWS VIA IN-MEMORY 2D ARRAY =====
    Dim srcLastRow As Long: srcLastRow = wsSrc.cells(wsSrc.Rows.count, 1).End(xlUp).row
    If srcLastRow < 4 Then GoTo Cleanup

    Dim srcData As Variant
    srcData = wsSrc.Range("A1:DF" & srcLastRow).Value2

    Dim r As Long, destRow As Long, rejectReason As String
    destRow = lastDestRow

    For r = 4 To srcLastRow
        Dim entryTicker As String: entryTicker = Trim$(CStr(srcData(r, 1)))
        If Len(entryTicker) = 0 Or Is_Index(entryTicker) Then GoTo NextStock

        ' Fetch indicators directly from memory array
        Dim ltp As Double:         ltp = CDbl(val(srcData(r, 20)))           ' Col T
        Dim stockRegime As String: stockRegime = UCase$(CStr(srcData(r, 66))) ' Col BN
        Dim stockGrade As String:  stockGrade = UCase$(CStr(srcData(r, 109))) ' Col DE
        Dim breakoutFlag As Long:  breakoutFlag = CLng(val(srcData(r, 110))) ' Col DF
        Dim stockADX As Double:    stockADX = CDbl(val(srcData(r, 80)))       ' Col CB
        Dim secRegime As String:   secRegime = GetSectorRegime(wsSrc, entryTicker)
        Dim volRatio As Double:    volRatio = CDbl(val(srcData(r, 67)))       ' Col BO
        Dim ema10 As Double:       ema10 = CDbl(val(srcData(r, 45)))          ' Col AS
        Dim ema30 As Double:       ema30 = CDbl(val(srcData(r, 46)))          ' Col AT

        If ltp <= 0 Then GoTo NextStock

        ' Determine Discovery Direction strictly from Donchian Flag
        Dim isBuy As Boolean, isShort As Boolean
        If breakoutFlag = 1 Then
            isBuy = True: isShort = False
        ElseIf breakoutFlag = -1 Then
            isBuy = False: isShort = True
        Else
            'AuditDecision_V2 wsSrc, r, entryTicker, "WAIT", stockRegime, stockGrade, secRegime, "", _
                             "NO_DONCHIAN_FLAG", "DF=" & breakoutFlag, breakoutFlag, volRatio, stockADX
            GoTo NextStock  ' No Donchian Breakout/Breakdown on this stock
        End If

        ' Candidate Conviction Cache for Overrides
        Dim candScore As Double: candScore = -1#
        Dim candCAT As String: candCAT = ""
        Dim isEmaOverride As Boolean: isEmaOverride = False
        Dim isGradeCOverride As Boolean: isGradeCOverride = False

        ' Evaluate Discovery Trigger (with EMA fail CAT1/CAT2 rescue)
        If Not PassDiscoveryTrigger_V2(wsSrc, r, entryTicker, isBuy, breakoutFlag, volRatio, ema10, ema30, rejectReason, isEmaOverride, candScore, candCAT) Then
            mRejections(entryTicker) = rejectReason
            AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, "", _
                             NormalizeV2Reason(rejectReason), rejectReason, breakoutFlag, volRatio, stockADX
            GoTo NextStock
        End If

        ' ===== STRATEGY MATCHING =====
        Dim stratName As String: stratName = ""
        Dim stratSize As Double: stratSize = 0
        Dim stratSlMult As Double: stratSlMult = 1.5
        Dim stratTgtPct As Double: stratTgtPct = 0.015

        ' Sector Contra Gate
        If (stockRegime = "TRENDING_BULL" And (secRegime = "TRENDING_BEAR" Or secRegime = "MILD_DOWN")) Then
            mRejections(entryTicker) = "SECTOR_CONTRA (Bull vs " & secRegime & ")"
            AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, stratName, _
                             "SECTOR_CONTRA", CStr(mRejections(entryTicker)), breakoutFlag, volRatio, stockADX
            GoTo NextStock
        End If
        If (stockRegime = "TRENDING_BEAR" And (secRegime = "TRENDING_BULL" Or secRegime = "MILD_UP")) Then
            mRejections(entryTicker) = "SECTOR_CONTRA (Bear vs " & secRegime & ")"
            AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, stratName, _
                             "SECTOR_CONTRA", CStr(mRejections(entryTicker)), breakoutFlag, volRatio, stockADX
            GoTo NextStock
        End If

        ' Strategy Router with Grade C Conviction-Aware Override (Rescue only CAT1/CAT2)
        If ENABLE_BREAKOUT And stockRegime = "RANGING" Then
            If (stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B") And isBuy Then
                stratName = "Breakout_Long": stratSize = SIZE_BREAKOUT: stratSlMult = SL_ATR_BREAKOUT: stratTgtPct = TARGET_PCT_BREAKOUT
            ElseIf (stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B") And isShort Then
                stratName = "Breakout_Short": stratSize = SIZE_BREAKOUT: stratSlMult = SL_ATR_BREAKOUT: stratTgtPct = TARGET_PCT_BREAKOUT
            End If
        End If

        If stratName = "" And ENABLE_MOMENTUM Then
            If (stockRegime = "TRENDING_BULL" Or stockRegime = "MILD_UP") And isBuy Then
                If stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B" Then
                    stratName = "Momentum_Long_Strong": stratSize = SIZE_MOM_STRONG: stratSlMult = SL_ATR_MOM_BULL: stratTgtPct = TARGET_PCT_MOM_BULL
                ElseIf stockGrade = "C" Then
                    If Len(candCAT) = 0 Then
                        Dim dStateG As String, r15G As Double, adxValG As Double, erValG As Double, rvValG As Double, dDistG As Double
                        CalculateConvictionScore_V2 wsSrc, r, isBuy, candScore, candCAT, dStateG, r15G, adxValG, erValG, rvValG, dDistG
                    End If
                    If candCAT = "CAT1" Or candCAT = "CAT2" Then
                        stratName = "Momentum_Long_Strong": stratSize = SIZE_MOM_STRONG: stratSlMult = SL_ATR_MOM_BULL: stratTgtPct = TARGET_PCT_MOM_BULL
                        isGradeCOverride = True
                    Else
                        mRejections(entryTicker) = "MOM_GRADE_C"
                        AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, "Momentum_Long_Strong", _
                                         "MOM_GRADE_C", "Grade C rejected [ShadowCAT=" & candCAT & "|S=" & Format(candScore, "0") & "]", breakoutFlag, volRatio, stockADX
                        GoTo NextStock
                    End If
                End If
            ElseIf (stockRegime = "TRENDING_BEAR" Or stockRegime = "MILD_DOWN") And isShort Then
                If stockGrade = "A+" Or stockGrade = "A" Or stockGrade = "B" Then
                    stratName = "Momentum_Short": stratSize = SIZE_MOM_BEAR: stratSlMult = SL_ATR_MOM_BEAR: stratTgtPct = TARGET_PCT_MOM_BEAR
                ElseIf stockGrade = "C" Then
                    If Len(candCAT) = 0 Then
                        Dim dStateGs As String, r15Gs As Double, adxValGs As Double, erValGs As Double, rvValGs As Double, dDistGs As Double
                        CalculateConvictionScore_V2 wsSrc, r, isBuy, candScore, candCAT, dStateGs, r15Gs, adxValGs, erValGs, rvValGs, dDistGs
                    End If
                    If candCAT = "CAT1" Or candCAT = "CAT2" Then
                        stratName = "Momentum_Short": stratSize = SIZE_MOM_BEAR: stratSlMult = SL_ATR_MOM_BEAR: stratTgtPct = TARGET_PCT_MOM_BEAR
                        isGradeCOverride = True
                    Else
                        mRejections(entryTicker) = "MOM_GRADE_C"
                        AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, "Momentum_Short", _
                                         "MOM_GRADE_C", "Grade C rejected [ShadowCAT=" & candCAT & "|S=" & Format(candScore, "0") & "]", breakoutFlag, volRatio, stockADX
                        GoTo NextStock
                    End If
                End If
            End If
        End If

        If stratName = "" Then
            mRejections(entryTicker) = "NO_STRATEGY_MATCH (Regime=" & stockRegime & " Grd=" & stockGrade & ")"
            AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, stratName, _
                             "NO_STRATEGY", CStr(mRejections(entryTicker)), breakoutFlag, volRatio, stockADX
            GoTo NextStock
        End If

        ' Active Duplicate Check
        If SymbolAlreadyActive(wsC2, entryTicker, IIf(isBuy, "BUY", "SELL")) Then
            mRejections(entryTicker) = "DUPLICATE_ACTIVE"
            AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, stratName, _
                             "DUPLICATE_ACTIVE", "Existing open trade in same direction", breakoutFlag, volRatio, stockADX
            GoTo NextStock
        End If

        ' Ticker Session Attempt Cap (Max 2 attempts per ticker per trading day across BUY + SELL combined)
        Dim sessionAttempts As Long: sessionAttempts = 0
        If Not mTickerAttempts Is Nothing Then
            If mTickerAttempts.Exists(entryTicker) Then sessionAttempts = CLng(mTickerAttempts(entryTicker))
        End If

        If sessionAttempts >= MAX_ATTEMPTS_PER_TICKER Then
            rejectReason = "TICKER_MAX_ATTEMPTS_REACHED (" & sessionAttempts & "/" & MAX_ATTEMPTS_PER_TICKER & ")"
            mRejections(entryTicker) = rejectReason
            AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, stratName, _
                             "TICKER_MAX_ATTEMPTS_REACHED", "Attempts=" & sessionAttempts, breakoutFlag, volRatio, stockADX
            GoTo NextStock
        End If
        If IsOnCooldown(entryTicker) Then
            mRejections(entryTicker) = "ON_COOLDOWN"
            AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, stratName, _
                             "ON_COOLDOWN", "Ticker inside cooldown window", breakoutFlag, volRatio, stockADX
            GoTo NextStock
        End If

        ' Downstream Strategy Gates
        If Not PassStrategyGates_V2(wsSrc, r, entryTicker, stratName, isBuy, stockADX, secRegime, rejectReason) Then
            mRejections(entryTicker) = rejectReason
            AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, stratName, _
                             NormalizeV2Reason(rejectReason), rejectReason, breakoutFlag, volRatio, stockADX
            GoTo NextStock
        End If

        Dim rsiStatus As String
        If Not PassRSIContextGate_V2(wsSrc, r, entryTicker, stratName, isBuy, stockRegime, volRatio, rejectReason, rsiStatus) Then
            mRejections(entryTicker) = rejectReason
            AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, stratName, _
                             NormalizeV2Reason(rejectReason), rejectReason, breakoutFlag, volRatio, stockADX
            GoTo NextStock
        End If

        ' Determine Override Reason for Audit
        Dim overrideReason As String
        If isEmaOverride And isGradeCOverride Then
            overrideReason = "EMA_AND_GRADE_C"
        ElseIf isEmaOverride Then
            overrideReason = "EMA_FAIL"
        ElseIf isGradeCOverride Then
            overrideReason = "GRADE_C"
        Else
            overrideReason = "NONE"
        End If

        ' Write Row to COCKPIT_V2 Directly From Sheet1
        destRow = destRow + 1
        WriteNewCockpitV2Row wsC2, wsSrc, r, destRow, isBuy, stratName, stratSize, stratSlMult, stratTgtPct, stockRegime, stockGrade, secRegime, volRatio, rsiStatus, overrideReason, candScore, candCAT
        If Not mActiveTrades Is Nothing Then mActiveTrades(entryTicker & "|" & IIf(isBuy, "BUY", "SELL")) = True
        If Not mTickerAttempts Is Nothing Then
            If mTickerAttempts.Exists(entryTicker) Then
                mTickerAttempts(entryTicker) = mTickerAttempts(entryTicker) + 1
            Else
                mTickerAttempts(entryTicker) = 1
            End If
        End If
        
        Dim auditPassDetail As String
        If overrideReason <> "NONE" Then
            auditPassDetail = "OVERRIDE=" & overrideReason & "|CAT=" & candCAT & "|S=" & Format(candScore, "0") & " (" & rsiStatus & ")"
        Else
            auditPassDetail = rsiStatus
        End If
        
        AuditDecision_V2 wsSrc, r, entryTicker, IIf(isBuy, "BUY", "SELL"), stockRegime, stockGrade, secRegime, stratName, _
                         "PASS", auditPassDetail, breakoutFlag, volRatio, stockADX

NextStock:
    Next r

    'WriteCockpitV2FunnelAudit

    ' Conviction Tracking: Refresh Conviction for all OPEN trades (Shadow Mode)
    RefreshOpenTradeConviction_V2 wsC2, wsSrc

    ' Lifecycle & Exit Management
    ManageCockpitTrades_V2 wsC2, wsSrc

Cleanup:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    Application.EnableEvents = True
End Sub


' =====================================================================
' Private Function  PassDiscoveryTrigger_V2
' =====================================================================
Private Function PassDiscoveryTrigger_V2(wsSrc As Worksheet, r As Long, ticker As String, _
        isBuy As Boolean, breakoutFlag As Long, volRatio As Double, _
        ema10 As Double, ema30 As Double, ByRef rejectReason As String, _
        ByRef isEmaOverride As Boolean, ByRef candScore As Double, ByRef candCAT As String) As Boolean

    PassDiscoveryTrigger_V2 = False
    isEmaOverride = False

    Select Case UCase$(Trim$(DISCOVERY_MODE))
        Case "DONCHIAN_RVOL"
            If isBuy Then
                If breakoutFlag <> 1 Then rejectReason = "No Donchian Breakout Flag (DF=" & breakoutFlag & ")": Exit Function
            Else
                If breakoutFlag <> -1 Then rejectReason = "No Donchian Breakdown Flag (DF=" & breakoutFlag & ")": Exit Function
            End If
            If volRatio < GATE_BO_VOL_MIN Then
                rejectReason = "Volume Ratio RVOL=" & Format(volRatio, "0.00") & " < " & GATE_BO_VOL_MIN: Exit Function
            End If

            ' EMA Trend Filter with Conviction-Aware Override (Rescue only CAT1/CAT2)
            Dim isEmaFail As Boolean
            If isBuy Then isEmaFail = (ema10 <= ema30) Else isEmaFail = (ema10 >= ema30)

            If isEmaFail Then
                Dim dState As String, r15 As Double, adxVal As Double, erVal As Double, rvVal As Double, dDist As Double
                CalculateConvictionScore_V2 wsSrc, r, isBuy, candScore, candCAT, dState, r15, adxVal, erVal, rvVal, dDist
                If candCAT = "CAT1" Or candCAT = "CAT2" Then
                    isEmaOverride = True
                Else
                    If isBuy Then
                        rejectReason = "EMA Trend Filter (EMA10 " & Format(ema10, "0.0") & " <= EMA30 " & Format(ema30, "0.0") & ") [ShadowCAT=" & candCAT & "|S=" & Format(candScore, "0") & "]"
                    Else
                        rejectReason = "EMA Trend Filter (EMA10 " & Format(ema10, "0.0") & " >= EMA30 " & Format(ema30, "0.0") & ") [ShadowCAT=" & candCAT & "|S=" & Format(candScore, "0") & "]"
                    End If
                    Exit Function
                End If
            End If

        Case "EMA_CROSS"
            Dim cross1030 As String: cross1030 = LCase(CStr(wsSrc.cells(r, "AC").value))
            If isBuy And cross1030 <> "over" Then rejectReason = "No 10/30 Over Cross": Exit Function
            If Not isBuy And cross1030 <> "under" Then rejectReason = "No 10/30 Under Cross": Exit Function

    End Select

    PassDiscoveryTrigger_V2 = True
End Function


' =====================================================================
' Private Sub  WriteNewCockpitV2Row
' Reads 100% directly from Sheet1 & applies exact V1 color coding
' =====================================================================
Private Sub WriteNewCockpitV2Row(wsC2 As Worksheet, wsSrc As Worksheet, srcRow As Long, _
        destRow As Long, isBuy As Boolean, stratName As String, _
        size As Double, slMult As Double, tgtPct As Double, _
        stockRegime As String, stockGrade As String, secRegime As String, volRatio As Double, _
        rsiStatus As String, Optional overrideReason As String = "NONE", _
        Optional preScore As Double = -1#, Optional preCAT As String = "")

    Dim ticker As String: ticker = CStr(wsSrc.cells(srcRow, "A").value)
    Dim ltp As Double: ltp = CDbl(val(wsSrc.cells(srcRow, "T").value))
    Dim atr5m As Double: atr5m = GetATR5m(wsSrc, ticker)

    ' Read Pivots S3..R3 directly from Sheet1 Cols K to Q
    Dim s3 As Double: s3 = CDbl(val(wsSrc.cells(srcRow, "K").value))
    Dim s2 As Double: s2 = CDbl(val(wsSrc.cells(srcRow, "L").value))
    Dim s1 As Double: s1 = CDbl(val(wsSrc.cells(srcRow, "M").value))
    Dim pp As Double: pp = CDbl(val(wsSrc.cells(srcRow, "N").value))
    Dim r1 As Double: r1 = CDbl(val(wsSrc.cells(srcRow, "O").value))
    Dim r2 As Double: r2 = CDbl(val(wsSrc.cells(srcRow, "P").value))
    Dim r3 As Double: r3 = CDbl(val(wsSrc.cells(srcRow, "Q").value))

    ' Read SuperTrend 5m directly from Sheet1 Col BT
    Dim superTrend As String: superTrend = CStr(wsSrc.cells(srcRow, "BT").value)
    Dim cross1030 As String: cross1030 = CStr(wsSrc.cells(srcRow, "AC").value)
    Dim cross1050 As String: cross1050 = CStr(wsSrc.cells(srcRow, "AG").value)

    ' Read RSI directly from Sheet1 Cols U (5m), V (15m), W (1h), X (1d)
    Dim rsi5 As Double: rsi5 = CDbl(val(wsSrc.cells(srcRow, "U").value))
    Dim rsi15 As Double: rsi15 = CDbl(val(wsSrc.cells(srcRow, "V").value))
    Dim rsi1h As Double: rsi1h = CDbl(val(wsSrc.cells(srcRow, "W").value))
    Dim rsi1d As Double: rsi1d = CDbl(val(wsSrc.cells(srcRow, "X").value))

    ' SL & Target Calculation with Risk-Budget-Preserving Sizing Floor
    Dim baseRisk As Double: baseRisk = 0#
    If atr5m > 0 Then baseRisk = atr5m * slMult

    Dim minRisk As Double: minRisk = ltp * RISK_FLOOR_MIN_PCT
    Dim initialRisk As Double
    Dim finalSize As Double
    Dim slPx As Double, tgtPx As Double

    If baseRisk > 0 Then
        If baseRisk < minRisk And minRisk > 0 Then
            ' Stop widened to 0.35% floor; position size scaled down proportionally to keep monetary risk constant
            initialRisk = minRisk
            finalSize = Round(size * (baseRisk / minRisk), 4)
        Else
            ' Normal ATR stop (>= 0.35%); position size unchanged
            initialRisk = baseRisk
            finalSize = size
        End If

        If isBuy Then
            slPx = ltp - initialRisk
            tgtPx = ltp * (1# + tgtPct)
        Else
            slPx = ltp + initialRisk
            tgtPx = ltp * (1# - tgtPct)
        End If
    Else
        ' Existing fallback when ATR is missing/zero (preserves existing behavior)
        If isBuy Then slPx = ltp * 0.995: tgtPx = ltp * 1.015 Else slPx = ltp * 1.005: tgtPx = ltp * 0.985
        If isBuy Then initialRisk = ltp - slPx Else initialRisk = slPx - ltp
        finalSize = size
    End If
    If initialRisk <= 0 Then initialRisk = 0

    ' Write Sheet Columns A to R (Directly Populated from Sheet1)
    wsC2.cells(destRow, RangeRef(CK_TICKER)) = ticker
    wsC2.cells(destRow, RangeRef(CK_ACTION)) = IIf(isBuy, "BUY", "SELL")
    wsC2.cells(destRow, RangeRef(CK_LTP)) = ltp
    wsC2.cells(destRow, RangeRef(CK_SECTOR)).Formula = _
        "=IF(XLOOKUP(A" & destRow & ",Sheet1!A:A,Sheet1!BM:BM)<>0," & _
        "XLOOKUP(A" & destRow & ",Sheet1!A:A,Sheet1!BM:BM),"""")"

    wsC2.cells(destRow, RangeRef(CK_S3)) = s3
    wsC2.cells(destRow, RangeRef(CK_S2)) = s2
    wsC2.cells(destRow, RangeRef(CK_S1)) = s1
    wsC2.cells(destRow, RangeRef(CK_PP)) = pp
    wsC2.cells(destRow, RangeRef(CK_R1)) = r1
    wsC2.cells(destRow, RangeRef(CK_R2)) = r2
    wsC2.cells(destRow, RangeRef(CK_R3)) = r3

    wsC2.cells(destRow, RangeRef(CK_RSI_5M)) = rsi5
    wsC2.cells(destRow, RangeRef(CK_RSI_15M)) = rsi15
    wsC2.cells(destRow, RangeRef(CK_RSI_1H)) = rsi1h
    wsC2.cells(destRow, RangeRef(CK_RSI_1D)) = rsi1d

    wsC2.cells(destRow, RangeRef(CK_SUPERTREND)) = superTrend
    wsC2.cells(destRow, RangeRef(CK_CROSS_10_30)) = cross1030
    wsC2.cells(destRow, RangeRef(CK_CROSS_10_50)) = cross1050

    ' Write Execution Columns S to AF
    wsC2.cells(destRow, RangeRef(CK_ENT)) = ltp
    wsC2.cells(destRow, RangeRef(CK_TGT)) = tgtPx
    wsC2.cells(destRow, RangeRef(CK_SL)) = slPx
    wsC2.cells(destRow, RangeRef(CK_CP)) = ltp
    wsC2.cells(destRow, RangeRef(CK_OPEN_TIME)) = Format(Now, "YYYY-MM-DD HH:MM:SS")

    ' Calculate Initial Conviction Score & Category (Shadow Mode)
    Dim initScore As Double, initCAT As String, initDonchState As String
    Dim initRsi15 As Double, initAdx As Double, initEr10 As Double, initRvol As Double, initDistAtr As Double
    If preScore >= 0 And Len(preCAT) > 0 Then
        initScore = preScore
        initCAT = preCAT
    Else
        CalculateConvictionScore_V2 wsSrc, srcRow, isBuy, initScore, initCAT, initDonchState, _
                                   initRsi15, initAdx, initEr10, initRvol, initDistAtr
    End If

    ' Metadata Column AA with visible formatted prefix
    Dim catPrefix As String
    catPrefix = "[" & initCAT & "|S=" & Format(initScore, "0") & "|INIT=" & initCAT & "|IS=" & Format(initScore, "0") & "] "
    Dim metaText As String
    If overrideReason <> "NONE" Then
        metaText = catPrefix & "[OVERRIDE: " & overrideReason & " | TrigPx: " & Format(ltp, "0.00") & " | VolRatio: " & Format(volRatio, "0.00") & " | ER10: " & Format(GetStockER10(wsSrc, ticker), "0.00") & " | RSI: " & rsiStatus & "]"
    Else
        metaText = catPrefix & "[DISCOVERY: " & DISCOVERY_MODE & " | TrigPx: " & Format(ltp, "0.00") & " | VolRatio: " & Format(volRatio, "0.00") & " | ER10: " & Format(GetStockER10(wsSrc, ticker), "0.00") & " | RSI: " & rsiStatus & "]"
    End If
    wsC2.cells(destRow, RangeRef(CK_COMENTS)) = metaText

    wsC2.cells(destRow, RangeRef(CK_CANDLE_COUNT)) = 0
    wsC2.cells(destRow, RangeRef(CK_STRATEGY)) = stratName

    wsC2.cells(destRow, RangeRef(CK_PEAK_PRICE)) = ltp
    wsC2.cells(destRow, RangeRef(CK_BEST_MFE_PCT)) = 0#
    wsC2.cells(destRow, RangeRef(CK_LAST_MFE_BAR)) = 0
    wsC2.cells(destRow, RangeRef(CK_INITIAL_RISK)) = initialRisk
    wsC2.cells(destRow, RangeRef(CK_INITIAL_RISK_PCT)) = IIf(ltp > 0, initialRisk / ltp, 0#)
    wsC2.cells(destRow, RangeRef(CK_ENTRY_ATR)) = atr5m
    wsC2.cells(destRow, RangeRef(CK_STOP_DISTANCE)) = initialRisk
    wsC2.cells(destRow, RangeRef(CK_POSITION_SIZE)) = finalSize
    wsC2.cells(destRow, RangeRef(CK_PNL_R)).Formula = "=IF(" & CK_INITIAL_RISK & destRow & ">0," & CK_BOTTOMLINE & destRow & "/" & CK_INITIAL_RISK & destRow & ","""")"
    wsC2.cells(destRow, RangeRef(CK_BEST_MFE_R)).Formula = "=IF(" & CK_INITIAL_RISK & destRow & ">0," & CK_ENT & destRow & "*" & CK_BEST_MFE_PCT & destRow & "/" & CK_INITIAL_RISK & destRow & ","""")"
    wsC2.cells(destRow, RangeRef(CK_WORST_MAE_PCT)) = 0#
    wsC2.cells(destRow, RangeRef(CK_WORST_MAE_R)).Formula = "=IF(" & CK_INITIAL_RISK & destRow & ">0," & CK_ENT & destRow & "*" & CK_WORST_MAE_PCT & destRow & "/" & CK_INITIAL_RISK & destRow & ","""")"
    wsC2.cells(destRow, RangeRef(CK_MFE_1_R)) = ""
    wsC2.cells(destRow, RangeRef(CK_MFE_2_R)) = ""
    wsC2.cells(destRow, RangeRef(CK_MFE_3_R)) = ""
    wsC2.cells(destRow, RangeRef(CK_MFE_5_R)) = ""
    wsC2.cells(destRow, RangeRef(CK_BARS_TO_MFE)) = 0
    wsC2.cells(destRow, RangeRef(CK_LAST_MAE_BAR)) = 0

    ' Write Authoritative Conviction State Columns AV..BA
    wsC2.cells(destRow, RangeRef(CK_INITIAL_CAT)) = initCAT
    wsC2.cells(destRow, RangeRef(CK_INITIAL_SCORE)) = initScore
    wsC2.cells(destRow, RangeRef(CK_CURRENT_CAT)) = initCAT
    wsC2.cells(destRow, RangeRef(CK_CURRENT_SCORE)) = initScore
    wsC2.cells(destRow, RangeRef(CK_CAT_TRANSITIONS)) = 0
    wsC2.cells(destRow, RangeRef(CK_SCORE_CHANGES)) = 0
    wsC2.cells(destRow, RangeRef(CK_OVERRIDE_REASON)) = overrideReason

    ' Highlight CAT cell (Col AX)
    ApplyCATColorCoding_V2 wsC2, destRow, initCAT

    ' ===== APPLY EXACT V1 CELL COLORING =====
    ApplyCockpitColors_V2 wsC2, destRow, isBuy, ltp, _
        s3, s2, s1, pp, r1, r2, r3, _
        rsi5, rsi15, rsi1h, rsi1d, _
        superTrend, cross1030, cross1050
End Sub


' =====================================================================
' Private Sub  ApplyCockpitColors_V2
' Direct cell formatting — mirrors V1 COCKPIT visual design exactly
' =====================================================================
Private Sub ApplyCockpitColors_V2(wsC2 As Worksheet, destRow As Long, _
        ByVal isBuy As Boolean, ByVal ltp As Double, _
        ByVal s3 As Double, ByVal s2 As Double, ByVal s1 As Double, _
        ByVal pp As Double, ByVal r1 As Double, ByVal r2 As Double, ByVal r3 As Double, _
        ByVal rsi5 As Double, ByVal rsi15 As Double, ByVal rsi1h As Double, ByVal rsi1d As Double, _
        ByVal superTrend As String, ByVal cross1030 As String, ByVal cross1050 As String)

    Dim cellRng As Range

    ' --- Ticker font (A) ---
    Set cellRng = wsC2.Range(CK_TICKER & destRow)
    If isBuy Then cellRng.Font.Color = RGB(0, 150, 0) Else cellRng.Font.Color = RGB(180, 0, 0)

    ' --- Action (B) ---
    Set cellRng = wsC2.Range(CK_ACTION & destRow)
    If isBuy Then
        cellRng.Interior.Color = RGB(204, 255, 204)
        cellRng.Font.Color = RGB(0, 100, 0)
        cellRng.Font.Bold = True
    Else
        cellRng.Interior.Color = RGB(255, 204, 204)
        cellRng.Font.Color = RGB(150, 0, 0)
        cellRng.Font.Bold = True
    End If

    ' --- RSI (L-O): threshold coloring ---
    Dim rsiVals(1 To 4) As Double
    rsiVals(1) = rsi5: rsiVals(2) = rsi15: rsiVals(3) = rsi1h: rsiVals(4) = rsi1d
    Dim rsiCols As Variant: rsiCols = Array(CK_RSI_5M, CK_RSI_15M, CK_RSI_1H, CK_RSI_1D)
    Dim i As Long
    For i = 1 To 4
        If rsiVals(i) >= 80 Then
            wsC2.Range(rsiCols(i - 1) & destRow).Interior.Color = RGB(210, 180, 140)  ' tan
        ElseIf (rsiVals(i) >= 70 And rsiVals(i) < 80) Or (rsiVals(i) <= 30 And rsiVals(i) >= 21) Then
            wsC2.Range(rsiCols(i - 1) & destRow).Interior.Color = RGB(255, 255, 160)  ' yellow
        End If
    Next i

    ' --- 10 vs 30 (Q) ---
    Set cellRng = wsC2.Range(CK_CROSS_10_30 & destRow)
    If LCase(cross1030) = "over" Then
        cellRng.Interior.Color = RGB(204, 255, 204): cellRng.Font.Bold = True
    ElseIf LCase(cross1030) = "under" Then
        cellRng.Interior.Color = RGB(255, 204, 204): cellRng.Font.Bold = True
    End If

    ' --- 10 vs 50 (R) ---
    Set cellRng = wsC2.Range(CK_CROSS_10_50 & destRow)
    If LCase(cross1050) = "over" Then
        cellRng.Interior.Color = RGB(204, 255, 204): cellRng.Font.Bold = True
    ElseIf LCase(cross1050) = "under" Then
        cellRng.Interior.Color = RGB(255, 204, 204): cellRng.Font.Bold = True
    End If
End Sub


' =====================================================================
' Private Sub  ManageCockpitTrades_V2
' Full Lifecycle & Exit Management matching V1 Production
' =====================================================================
Private Sub ManageCockpitTrades_V2(wsC2 As Worksheet, wsSrc As Worksheet)
    Dim lastRow As Long, r As Long
    lastRow = wsC2.cells(wsC2.Rows.count, 1).End(xlUp).row
    If lastRow < DATA_START_ROW Then Exit Sub

    For r = DATA_START_ROW To lastRow
        Dim closeTimeVal As Variant: closeTimeVal = wsC2.Range(CK_CLOSE_TIME & r).value
        If Not IsEmpty(closeTimeVal) And Len(CStr(closeTimeVal)) > 0 Then GoTo NextTrade

        Dim ticker As String: ticker = CStr(wsC2.Range(CK_TICKER & r).value)
        Dim action As String: action = UCase(CStr(wsC2.Range(CK_ACTION & r).value))
        Dim entryPx As Double: entryPx = CDbl(val(wsC2.Range(CK_ENT & r).value))
        Dim tgtPx As Double: tgtPx = CDbl(val(wsC2.Range(CK_TGT & r).value))
        Dim slPx As Double: slPx = CDbl(val(wsC2.Range(CK_SL & r).value))
        Dim initialRisk As Double: initialRisk = CDbl(val(wsC2.Range(CK_INITIAL_RISK & r).value))
        If initialRisk <= 0 And entryPx > 0 And slPx > 0 Then
            If action = "BUY" Then initialRisk = entryPx - slPx Else initialRisk = slPx - entryPx
            If initialRisk > 0 Then
                wsC2.Range(CK_INITIAL_RISK & r).value = initialRisk
                wsC2.Range(CK_INITIAL_RISK_PCT & r).value = initialRisk / entryPx
                wsC2.Range(CK_STOP_DISTANCE & r).value = initialRisk
            End If
        End If

        Dim candleCount As Long: candleCount = CLng(val(wsC2.Range(CK_CANDLE_COUNT & r).value)) + 1
        wsC2.Range(CK_CANDLE_COUNT & r).value = candleCount

        Dim curLTP As Double: curLTP = GetCurrentLTP(wsSrc, ticker, entryPx)
        wsC2.Range(CK_CP & r).value = curLTP

        Dim curMfePct As Double
        If action = "BUY" Then curMfePct = (curLTP - entryPx) / entryPx Else curMfePct = (entryPx - curLTP) / entryPx
        Dim curMaePct As Double
        If curMfePct < 0 Then curMaePct = -curMfePct Else curMaePct = 0#

        Dim bestMfePct As Double: bestMfePct = CDbl(val(wsC2.Range(CK_BEST_MFE_PCT & r).value))
        Dim lastMfeBar As Long: lastMfeBar = CLng(val(wsC2.Range(CK_LAST_MFE_BAR & r).value))

        If curMfePct > bestMfePct Then
            bestMfePct = curMfePct
            lastMfeBar = candleCount
            wsC2.Range(CK_BEST_MFE_PCT & r).value = bestMfePct
            wsC2.Range(CK_LAST_MFE_BAR & r).value = lastMfeBar
            wsC2.Range(CK_BARS_TO_MFE & r).value = lastMfeBar
        End If

        Dim worstMaePct As Double: worstMaePct = CDbl(val(wsC2.Range(CK_WORST_MAE_PCT & r).value))
        If curMaePct > worstMaePct Then
            wsC2.Range(CK_WORST_MAE_PCT & r).value = curMaePct
            wsC2.Range(CK_LAST_MAE_BAR & r).value = candleCount
        End If

        Dim bestMfeR As Double
        If initialRisk > 0 Then bestMfeR = (entryPx * bestMfePct) / initialRisk Else bestMfeR = 0#
        Select Case candleCount
            Case 1
                If Len(CStr(wsC2.Range(CK_MFE_1_R & r).value)) = 0 Then wsC2.Range(CK_MFE_1_R & r).value = bestMfeR
            Case 2
                If Len(CStr(wsC2.Range(CK_MFE_2_R & r).value)) = 0 Then wsC2.Range(CK_MFE_2_R & r).value = bestMfeR
            Case 3
                If Len(CStr(wsC2.Range(CK_MFE_3_R & r).value)) = 0 Then wsC2.Range(CK_MFE_3_R & r).value = bestMfeR
            Case 5
                If Len(CStr(wsC2.Range(CK_MFE_5_R & r).value)) = 0 Then wsC2.Range(CK_MFE_5_R & r).value = bestMfeR
        End Select

        Dim exitFlag As Boolean: exitFlag = False
        Dim exitReason As String

        If TimeValue(Now) >= TimeValue(FORCE_LIQUIDATION) Then
            exitFlag = True: exitReason = "Session Liquidation (" & FORCE_LIQUIDATION & ")"
        End If

        If Not exitFlag Then
            If action = "BUY" Then
                If slPx > 0 And curLTP <= slPx Then exitFlag = True: exitReason = "SL Hit"
                If tgtPx > 0 And curLTP >= tgtPx Then exitFlag = True: exitReason = "Target Hit"
            Else
                If slPx > 0 And curLTP >= slPx Then exitFlag = True: exitReason = "SL Hit"
                If tgtPx > 0 And curLTP <= tgtPx Then exitFlag = True: exitReason = "Target Hit"
            End If
        End If

        If Not exitFlag Then
            Dim atrBE As Double: atrBE = GetATR5m(wsSrc, ticker)
            Dim beThresh As Double
            If atrBE > 0 And entryPx > 0 Then beThresh = (atrBE * TRAILING_BE_ATR_MULT) / entryPx Else beThresh = 0.002
            If bestMfePct >= beThresh Then
                If action = "BUY" Then
                    If slPx < entryPx * 1.0005 Then wsC2.Range(CK_SL & r).value = entryPx * 1.0005
                Else
                    If slPx > entryPx * 0.9995 Then wsC2.Range(CK_SL & r).value = entryPx * 0.9995
                End If
            End If
        End If

        ' --- 4. CAT DECAY EXIT ---
        If Not exitFlag And ENABLE_CAT_DECAY_EXIT Then
            Dim initCAT As String: initCAT = Trim$(CStr(wsC2.Range(CK_INITIAL_CAT & r).value))
            Dim curCAT As String: curCAT = Trim$(CStr(wsC2.Range(CK_CURRENT_CAT & r).value))
            Dim isCatDecay As Boolean: isCatDecay = False

            If initCAT = "CAT1" Then
                If curCAT = "CAT3" Or curCAT = "CAT4" Or curCAT = "REJECT" Then isCatDecay = True
            ElseIf initCAT = "CAT2" Then
                If curCAT = "CAT4" Or curCAT = "REJECT" Then isCatDecay = True
            ElseIf initCAT = "CAT3" Then
                If curCAT = "REJECT" Then isCatDecay = True
            End If

            If isCatDecay Then
                exitFlag = True
                exitReason = "CAT_DECAY (" & initCAT & "->" & curCAT & ")"
            End If
        End If

        If Not exitFlag And candleCount >= STAGNATION_BARS Then
            If bestMfePct < STAGNATION_MIN_MFE_PCT And curMfePct <= DEAD_TRADE_EXIT_MAX_PCT Then
                exitFlag = True: exitReason = "Dead Trade (" & STAGNATION_BARS & " bars)"
            End If
        End If

        If Not exitFlag And candleCount >= MAX_CANDLES_EXIT Then
            exitFlag = True: exitReason = "Expired (" & MAX_CANDLES_EXIT & " candles)"
        End If

        If exitFlag Then
            wsC2.Range(CK_CP & r).value = curLTP
            wsC2.Range(CK_CLOSE_TIME & r).value = Format(Now, "YYYY-MM-DD HH:MM:SS")
            wsC2.Range(CK_COMENTS & r).value = CStr(wsC2.Range(CK_COMENTS & r).value) & " -> " & exitReason
            RecordCooldown ticker
        End If

NextTrade:
    Next r
End Sub


' =====================================================================
' Public Sub  Run_Cockpit_Compare
' =====================================================================
Public Sub Run_Cockpit_Compare()
    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsC1 As Worksheet: Set wsC1 = ThisWorkbook.Sheets("COCKPIT")
    Dim wsC2 As Worksheet: Set wsC2 = ThisWorkbook.Sheets("COCKPIT_V2")
    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Sheets("Sheet1")

    Dim wsComp As Worksheet
    On Error Resume Next
    Set wsComp = ThisWorkbook.Sheets("COCKPIT_COMPARE")
    If wsComp Is Nothing Then
        Set wsComp = ThisWorkbook.Sheets.Add(After:=wsC2)
        wsComp.Name = "COCKPIT_COMPARE"
    End If
    On Error GoTo 0

    wsComp.Range("A1:I1").value = Array("Ticker", "Open Time", "Classification", "V1 Action", "V2 Action", "V1 Strategy", "V2 Strategy", "V1 Status / Outcome", "V2 Diagnostic / Rejection Reason")
    wsComp.Range("A1:I1").Font.Bold = True
    wsComp.Range("A1:I1").Interior.Color = RGB(220, 230, 241)

    Dim r1Last As Long: r1Last = wsC1.cells(wsC1.Rows.count, 1).End(xlUp).row
    Dim r2Last As Long: r2Last = wsC2.cells(wsC2.Rows.count, 1).End(xlUp).row

    Dim v1Dict As Object: Set v1Dict = CreateObject("Scripting.Dictionary")
    Dim v2Dict As Object: Set v2Dict = CreateObject("Scripting.Dictionary")

    Dim r As Long, key As String, tkr As String, tme As String, act As String, strat As String, cmt As String

    For r = DATA_START_ROW To r1Last
        tkr = CStr(wsC1.Range(CK_TICKER & r).value)
        tme = CStr(wsC1.Range(CK_OPEN_TIME & r).value)
        act = CStr(wsC1.Range(CK_ACTION & r).value)
        strat = CStr(wsC1.Range(CK_STRATEGY & r).value)
        cmt = CStr(wsC1.Range(CK_COMENTS & r).value)
        If Len(tkr) > 0 Then
            key = tkr & "|" & Left(tme, 16)
            v1Dict(key) = Array(tkr, tme, act, strat, cmt)
        End If
    Next r

    For r = DATA_START_ROW To r2Last
        tkr = CStr(wsC2.Range(CK_TICKER & r).value)
        tme = CStr(wsC2.Range(CK_OPEN_TIME & r).value)
        act = CStr(wsC2.Range(CK_ACTION & r).value)
        strat = CStr(wsC2.Range(CK_STRATEGY & r).value)
        cmt = CStr(wsC2.Range(CK_COMENTS & r).value)
        If Len(tkr) > 0 Then
            key = tkr & "|" & Left(tme, 16)
            v2Dict(key) = Array(tkr, tme, act, strat, cmt)
        End If
    Next r

    Dim outRow As Long: outRow = 0
    Dim k As Variant, item1 As Variant, item2 As Variant, cls As String, v2Reason As String

    Dim allKeys As Object: Set allKeys = CreateObject("Scripting.Dictionary")
    For Each k In v1Dict.Keys: allKeys(k) = True: Next k
    For Each k In v2Dict.Keys: allKeys(k) = True: Next k

    Dim totalKeys As Long: totalKeys = allKeys.count
    Dim outArr As Variant, clsArr As Variant
    If totalKeys > 0 Then
        ReDim outArr(1 To totalKeys, 1 To 9)
        ReDim clsArr(1 To totalKeys)
    End If

    For Each k In allKeys.Keys
        Dim hasV1 As Boolean: hasV1 = v1Dict.Exists(k)
        Dim hasV2 As Boolean: hasV2 = v2Dict.Exists(k)
        outRow = outRow + 1
        If hasV1 Then item1 = v1Dict(k)
        If hasV2 Then item2 = v2Dict(k)

        If hasV1 And hasV2 Then
            cls = "BOTH"
            outArr(outRow, 1) = item1(0): outArr(outRow, 2) = item1(1): outArr(outRow, 3) = cls
            outArr(outRow, 4) = item1(2): outArr(outRow, 5) = item2(2)
            outArr(outRow, 6) = item1(3): outArr(outRow, 7) = item2(3)
            outArr(outRow, 8) = item1(4): outArr(outRow, 9) = item2(4)
        ElseIf hasV1 Then
            cls = "V1_ONLY"
            tkr = CStr(item1(0))
            v2Reason = "REJECTED: No Donchian Breakout Flag (DF=0)"
            If Not mRejections Is Nothing Then
                If mRejections.Exists(tkr) Then v2Reason = "REJECTED: " & CStr(mRejections(tkr))
            End If
            outArr(outRow, 1) = item1(0): outArr(outRow, 2) = item1(1): outArr(outRow, 3) = cls
            outArr(outRow, 4) = item1(2): outArr(outRow, 5) = "WAIT"
            outArr(outRow, 6) = item1(3): outArr(outRow, 7) = "-"
            outArr(outRow, 8) = item1(4): outArr(outRow, 9) = v2Reason
        Else
            cls = "V2_ONLY"
            outArr(outRow, 1) = item2(0): outArr(outRow, 2) = item2(1): outArr(outRow, 3) = cls
            outArr(outRow, 4) = "WAIT": outArr(outRow, 5) = item2(2)
            outArr(outRow, 6) = "-": outArr(outRow, 7) = item2(3)
            outArr(outRow, 8) = "MISSED: No EMA 10/30 Crossover on this bar": outArr(outRow, 9) = item2(4)
        End If
        clsArr(outRow) = cls
    Next k

    If totalKeys > 0 Then wsComp.Range("A2:I" & (totalKeys + 1)).value = outArr

    If totalKeys > 0 Then
        Dim runStart As Long: runStart = 2
        Dim curRow As Long
        For curRow = 3 To totalKeys + 1
            If clsArr(curRow - 1) <> clsArr(runStart - 1) Then
                ColorCompareRun wsComp, runStart, curRow - 1, clsArr(runStart - 1)
                runStart = curRow
            End If
        Next curRow
        ColorCompareRun wsComp, runStart, totalKeys + 1, clsArr(runStart - 1)
    End If

    If totalKeys <= 20 Then wsComp.Columns("A:I").AutoFit
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
End Sub


' =====================================================================
' Audit-only V2 funnel logging. These routines must never feed decisions.
' =====================================================================
Private Sub AuditDecision_V2(wsSrc As Worksheet, srcRow As Long, ticker As String, _
        side As String, stockRegime As String, stockGrade As String, secRegime As String, _
        stratName As String, finalReason As String, reasonDetail As String, _
        breakoutFlag As Long, volRatio As Double, adxVal As Double)

    If mAuditRows Is Nothing Then Set mAuditRows = New Collection
    If Len(mRunTimestamp) = 0 Then mRunTimestamp = Format(Now, "YYYY-MM-DD HH:MM:SS")

    Dim er10Val As Double: er10Val = GetStockER10(wsSrc, ticker)
    Dim rsi5 As Double: rsi5 = CDbl(val(wsSrc.cells(srcRow, "U").value))
    Dim donchianAuditPacket As String: donchianAuditPacket = CStr(wsSrc.cells(srcRow, "BX").value)
    Dim donchianAuditState As String: donchianAuditState = CStr(wsSrc.cells(srcRow, "BY").value)

    mAuditRows.Add Array(mRunTimestamp, ticker, side, stockRegime, secRegime, stockGrade, _
                         stratName, finalReason, reasonDetail, breakoutFlag, volRatio, _
                         adxVal, er10Val, rsi5, donchianAuditPacket, donchianAuditState)
End Sub

Private Sub WriteCockpitV2FunnelAudit()
    On Error GoTo AuditFail

    If mAuditRows Is Nothing Then Exit Sub
    If mAuditRows.count = 0 Then Exit Sub

    Dim wsSummary As Worksheet: Set wsSummary = GetOrCreateSheet_V2("COCKPIT_V2_FUNNEL")
    Dim wsDetail As Worksheet: Set wsDetail = GetOrCreateSheet_V2("COCKPIT_V2_AUDIT")

    If Len(CStr(wsSummary.Range("A1").value)) = 0 Then
        wsSummary.Range("A1:F1").value = Array("RunTimestamp", "Reason", "Count", "% of Universe", "Top Tickers", "Run Rows")
        wsSummary.Range("A1:F1").Font.Bold = True
        wsSummary.Range("A1:F1").Interior.Color = RGB(220, 230, 241)
    End If

    wsDetail.Range("A1:P1").value = Array("RunTimestamp", "Ticker", "Side", "Regime", "SectorRegime", _
        "Grade", "Strategy", "FinalReason", "ReasonDetail", "DonchianFlag", "VolRatio", "ADX", "ER10", "RSI5", _
        "DonchianEventPacket", "DonchianEventState")
    wsDetail.Range("A1:P1").Font.Bold = True
    wsDetail.Range("A1:P1").Interior.Color = RGB(220, 230, 241)

    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")
    Dim topTickers As Object: Set topTickers = CreateObject("Scripting.Dictionary")
    Dim i As Long, item As Variant, reason As String, ticker As String

    For i = 1 To mAuditRows.count
        item = mAuditRows(i)
        reason = CStr(item(7))
        ticker = CStr(item(1))

        If Not counts.Exists(reason) Then counts(reason) = 0
        counts(reason) = CLng(counts(reason)) + 1
        AppendTopTicker_V2 topTickers, reason, ticker
    Next i

    Dim detailStart As Long: detailStart = wsDetail.cells(wsDetail.Rows.count, 1).End(xlUp).row + 1
    Dim detailArr() As Variant
    ReDim detailArr(1 To mAuditRows.count, 1 To 16)

    Dim c As Long
    For i = 1 To mAuditRows.count
        item = mAuditRows(i)
        For c = 1 To 16
            detailArr(i, c) = item(c - 1)
        Next c
    Next i
    wsDetail.Range("A" & detailStart & ":P" & (detailStart + mAuditRows.count - 1)).value = detailArr

    Dim reasonOrder As Variant
    reasonOrder = Array("NO_DONCHIAN_FLAG", "DISCOVERY_VOL_FAIL", "DISCOVERY_EMA_FAIL", _
        "SECTOR_CONTRA", "MOM_GRADE_C", "NO_STRATEGY", "DUPLICATE_ACTIVE", "ON_COOLDOWN", _
        "ADX_FAIL", "MOM_VOL_FAIL", "ER_FAIL", "RSI_MISSING", "RSI_WEAK", _
        "RSI_EXTENDED_WEAK_Q", "PASS", "OTHER")

    Dim summaryStart As Long: summaryStart = wsSummary.cells(wsSummary.Rows.count, 1).End(xlUp).row + 1
    Dim summaryArr() As Variant
    ReDim summaryArr(1 To UBound(reasonOrder) + 1, 1 To 6)

    Dim cnt As Long
    For i = LBound(reasonOrder) To UBound(reasonOrder)
        reason = CStr(reasonOrder(i))
        cnt = 0
        If counts.Exists(reason) Then cnt = CLng(counts(reason))

        summaryArr(i + 1, 1) = mRunTimestamp
        summaryArr(i + 1, 2) = reason
        summaryArr(i + 1, 3) = cnt
        summaryArr(i + 1, 4) = cnt / mAuditRows.count
        If topTickers.Exists(reason) Then summaryArr(i + 1, 5) = topTickers(reason) Else summaryArr(i + 1, 5) = "-"
        summaryArr(i + 1, 6) = mAuditRows.count
    Next i
    wsSummary.Range("A" & summaryStart & ":F" & (summaryStart + UBound(reasonOrder))).value = summaryArr
    ' Performance: AutoFit only on initial header creation, never across thousands of historical rows
    If summaryStart <= 2 Then wsSummary.Columns("A:F").AutoFit
    If detailStart <= 2 Then wsDetail.Columns("A:P").AutoFit
    Exit Sub

AuditFail:
    Debug.Print "COCKPIT_V2 audit logging failed: " & Err.Number & " - " & Err.Description
End Sub

Private Function GetOrCreateSheet_V2(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet_V2 = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet_V2 Is Nothing Then
        Set GetOrCreateSheet_V2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        GetOrCreateSheet_V2.Name = sheetName
    End If
End Function

Private Sub AppendTopTicker_V2(topTickers As Object, reason As String, ticker As String)
    If Len(ticker) = 0 Then Exit Sub

    If Not topTickers.Exists(reason) Then
        topTickers(reason) = ticker
    ElseIf UBound(Split(CStr(topTickers(reason)), ", ")) < 11 Then
        topTickers(reason) = CStr(topTickers(reason)) & ", " & ticker
    End If
End Sub

Private Function NormalizeV2Reason(ByVal reasonText As String) As String
    Dim s As String: s = UCase$(reasonText)

    If InStr(s, "DONCHIAN") > 0 Then NormalizeV2Reason = "NO_DONCHIAN_FLAG": Exit Function
    If InStr(s, "VOLUME RATIO RVOL") > 0 Then NormalizeV2Reason = "DISCOVERY_VOL_FAIL": Exit Function
    If InStr(s, "EMA TREND FILTER") > 0 Then NormalizeV2Reason = "DISCOVERY_EMA_FAIL": Exit Function
    If InStr(s, "SECTOR_CONTRA") > 0 Then NormalizeV2Reason = "SECTOR_CONTRA": Exit Function
    If InStr(s, "MOM_GRADE_C") > 0 Then NormalizeV2Reason = "MOM_GRADE_C": Exit Function
    If InStr(s, "NO_STRATEGY") > 0 Then NormalizeV2Reason = "NO_STRATEGY": Exit Function
    If InStr(s, "DUPLICATE_ACTIVE") > 0 Then NormalizeV2Reason = "DUPLICATE_ACTIVE": Exit Function
    If InStr(s, "ON_COOLDOWN") > 0 Then NormalizeV2Reason = "ON_COOLDOWN": Exit Function
    If InStr(s, "RSI_MISSING") > 0 Then NormalizeV2Reason = "RSI_MISSING": Exit Function
    If InStr(s, "RSI_EXTENDED_WEAK_Q") > 0 Then NormalizeV2Reason = "RSI_EXTENDED_WEAK_Q": Exit Function
    If InStr(s, "RSI_WEAK") > 0 Then NormalizeV2Reason = "RSI_WEAK": Exit Function
    If InStr(s, "ADX=") > 0 Then NormalizeV2Reason = "ADX_FAIL": Exit Function
    If InStr(s, "VOLUME RATIO <") > 0 Then NormalizeV2Reason = "MOM_VOL_FAIL": Exit Function
    If InStr(s, "ER10") > 0 Then NormalizeV2Reason = "ER_FAIL": Exit Function

    NormalizeV2Reason = "OTHER"
End Function


' Helper Functions
Private Sub EnsureCockpitV2StateHeaders(wsC2 As Worksheet)
    If Len(CStr(wsC2.Range(CK_PEAK_PRICE & "4").value)) = 0 Then wsC2.Range(CK_PEAK_PRICE & "4").value = "Peak_Price"
    If Len(CStr(wsC2.Range(CK_BEST_MFE_PCT & "4").value)) = 0 Then wsC2.Range(CK_BEST_MFE_PCT & "4").value = "Best_MFE_Pct"
    If Len(CStr(wsC2.Range(CK_LAST_MFE_BAR & "4").value)) = 0 Then wsC2.Range(CK_LAST_MFE_BAR & "4").value = "Last_MFE_Bar"
    wsC2.Range(CK_INITIAL_RISK & "4").value = "InitialRisk"
    wsC2.Range(CK_INITIAL_RISK_PCT & "4").value = "InitialRiskPct"
    wsC2.Range(CK_ENTRY_ATR & "4").value = "EntryATR"
    wsC2.Range(CK_STOP_DISTANCE & "4").value = "StopDistance"
    wsC2.Range(CK_POSITION_SIZE & "4").value = "PositionSize"
    wsC2.Range(CK_PNL_R & "4").value = "PnL_R"
    wsC2.Range(CK_BEST_MFE_R & "4").value = "MFE_R"
    wsC2.Range(CK_WORST_MAE_PCT & "4").value = "MAE_Pct"
    wsC2.Range(CK_WORST_MAE_R & "4").value = "MAE_R"
    wsC2.Range(CK_MFE_1_R & "4").value = "MFE_1_R"
    wsC2.Range(CK_MFE_2_R & "4").value = "MFE_2_R"
    wsC2.Range(CK_MFE_3_R & "4").value = "MFE_3_R"
    wsC2.Range(CK_MFE_5_R & "4").value = "MFE_5_R"
    wsC2.Range(CK_BARS_TO_MFE & "4").value = "BarsToMFE"
    wsC2.Range(CK_LAST_MAE_BAR & "4").value = "Last_MAE_Bar"
    wsC2.Range(CK_INITIAL_CAT & "4").value = "Initial_CAT"
    wsC2.Range(CK_INITIAL_SCORE & "4").value = "Initial_Score"
    wsC2.Range(CK_CURRENT_CAT & "4").value = "Current_CAT"
    wsC2.Range(CK_CURRENT_SCORE & "4").value = "Current_Score"
    wsC2.Range(CK_CAT_TRANSITIONS & "4").value = "CAT_Transitions"
    wsC2.Range(CK_SCORE_CHANGES & "4").value = "Score_Changes"
    wsC2.Range(CK_OVERRIDE_REASON & "4").value = "Override_Reason"
    wsC2.Columns(CK_PEAK_PRICE & ":" & CK_OVERRIDE_REASON).Hidden = True
End Sub

Private Function BuildRowLookup(wsSrc As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, r As Long, t As String
    lastRow = wsSrc.cells(wsSrc.Rows.count, 1).End(xlUp).row
    For r = 4 To lastRow
        t = Trim$(CStr(wsSrc.cells(r, 1).value))
        If Len(t) > 0 Then
            If Not dict.Exists(t) Then dict(t) = r
        End If
    Next r
    Set BuildRowLookup = dict
End Function

Private Function GetCurrentLTP(wsSrc As Worksheet, ByVal ticker As String, ByVal fallback As Double) As Double
    GetCurrentLTP = fallback
    If Len(ticker) = 0 Then Exit Function
    If Not mRowLookup Is Nothing Then If mRowLookup.Exists(ticker) Then GetCurrentLTP = CDbl(val(wsSrc.cells(CLng(mRowLookup(ticker)), "T").value)): Exit Function
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetCurrentLTP = CDbl(wsSrc.cells(CLng(m), "T").value)
    On Error GoTo 0
End Function

Private Function GetATR5m(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetATR5m = 0
    If Len(ticker) = 0 Then Exit Function
    If Not mRowLookup Is Nothing Then If mRowLookup.Exists(ticker) Then GetATR5m = CDbl(val(wsSrc.cells(CLng(mRowLookup(ticker)), "CC").value)): Exit Function
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetATR5m = CDbl(wsSrc.cells(CLng(m), "CC").value)
    On Error GoTo 0
End Function

Private Function GetStockER10(wsSrc As Worksheet, ByVal ticker As String) As Double
    GetStockER10 = 0.5
    If Len(ticker) = 0 Then Exit Function
    If Not mRowLookup Is Nothing Then If mRowLookup.Exists(ticker) Then GetStockER10 = CDbl(val(wsSrc.cells(CLng(mRowLookup(ticker)), "DP").value)): Exit Function
    On Error Resume Next
    Dim m As Variant: m = Application.Match(ticker, wsSrc.Range("A:A"), 0)
    If Not IsError(m) Then GetStockER10 = CDbl(wsSrc.cells(CLng(m), "DP").value)
    On Error GoTo 0
End Function

Private Function GetSectorRegime(wsSrc As Worksheet, ByVal ticker As String) As String
    GetSectorRegime = ""
    If Len(ticker) = 0 Then Exit Function
    Dim sectorName As String
    If Not mRowLookup Is Nothing Then If mRowLookup.Exists(ticker) Then sectorName = CStr(wsSrc.cells(CLng(mRowLookup(ticker)), "BM").value)
    If Len(sectorName) > 0 Then If mRowLookup.Exists(sectorName) Then GetSectorRegime = UCase(CStr(wsSrc.cells(CLng(mRowLookup(sectorName)), "BN").value)): Exit Function
End Function

Private Function SymbolAlreadyActive(wsC2 As Worksheet, ticker As String, dir As String) As Boolean
    If Not mActiveTrades Is Nothing Then
        SymbolAlreadyActive = mActiveTrades.Exists(ticker & "|" & UCase$(dir))
        Exit Function
    End If
    Dim lastRow As Long, r As Long
    lastRow = wsC2.cells(wsC2.Rows.count, 1).End(xlUp).row
    For r = DATA_START_ROW To lastRow
        If CStr(wsC2.Range(CK_TICKER & r).value) = ticker Then
            If UCase$(CStr(wsC2.Range(CK_ACTION & r).value)) = UCase$(dir) Then
                If IsEmpty(wsC2.Range(CK_CLOSE_TIME & r).value) Or Len(CStr(wsC2.Range(CK_CLOSE_TIME & r).value)) = 0 Then
                    SymbolAlreadyActive = True: Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function BuildActiveTradeLookup(wsC2 As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, r As Long
    lastRow = wsC2.cells(wsC2.Rows.count, 1).End(xlUp).row
    If lastRow >= DATA_START_ROW Then
        For r = DATA_START_ROW To lastRow
            Dim closeVal As Variant: closeVal = wsC2.Range(CK_CLOSE_TIME & r).value
            If IsEmpty(closeVal) Or Len(CStr(closeVal)) = 0 Then
                Dim t As String: t = Trim$(CStr(wsC2.Range(CK_TICKER & r).value))
                Dim a As String: a = UCase$(Trim$(CStr(wsC2.Range(CK_ACTION & r).value)))
                If Len(t) > 0 And Len(a) > 0 Then dict(t & "|" & a) = True
            End If
        Next r
    End If
    Set BuildActiveTradeLookup = dict
End Function

Private Function IsOnCooldown(ticker As String) As Boolean
    If mCooldowns Is Nothing Then Set mCooldowns = CreateObject("Scripting.Dictionary")
    If Not mCooldowns.Exists(ticker) Then Exit Function
    If DateDiff("n", mCooldowns(ticker), Now) < COOLDOWN_MINUTES_VAL Then IsOnCooldown = True Else mCooldowns.Remove ticker
End Function

Private Sub RecordCooldown(ticker As String)
    If mCooldowns Is Nothing Then Set mCooldowns = CreateObject("Scripting.Dictionary")
    mCooldowns(ticker) = Now
End Sub

Private Function Is_Index(t As String) As Boolean
    Is_Index = (InStr(t, "NIFTY") > 0 Or InStr(t, "CNX") > 0 Or InStr(t, "NSEBANK") > 0 Or InStr(t, "MIDCAP") > 0 Or InStr(t, "SMALLCAP") > 0 Or InStr(t, "NSEI") > 0)
End Function

Private Function PassStrategyGates_V2(wsSrc As Worksheet, srcRow As Long, ticker As String, stratName As String, isBuy As Boolean, adxVal As Double, secRegime As String, ByRef rejectReason As String) As Boolean
    PassStrategyGates_V2 = False
    If adxVal < GATE_MOM_ADX_MIN Then rejectReason = "ADX=" & Format(adxVal, "0.0") & " < " & GATE_MOM_ADX_MIN: Exit Function
    If CDbl(val(wsSrc.cells(srcRow, "BO").value)) < GATE_MOM_VOL_MIN Then rejectReason = "Volume Ratio < " & GATE_MOM_VOL_MIN: Exit Function
    If GetStockER10(wsSrc, ticker) < GATE_MOM_ER10_MIN Then rejectReason = "ER10 < " & GATE_MOM_ER10_MIN: Exit Function
    PassStrategyGates_V2 = True
End Function

Private Function PassRSIContextGate_V2(wsSrc As Worksheet, srcRow As Long, ticker As String, _
        stratName As String, isBuy As Boolean, stockRegime As String, volRatio As Double, _
        ByRef rejectReason As String, ByRef rsiStatus As String) As Boolean

    PassRSIContextGate_V2 = False
    rsiStatus = "RSI_PASS"

    Dim rsi5 As Double: rsi5 = CDbl(val(wsSrc.cells(srcRow, "U").value))
    If rsi5 <= 0 Then rejectReason = "RSI_MISSING": Exit Function

    Dim minRsi As Double, maxRsi As Double
    Dim extHigh As Double, extLow As Double
    Dim isExtended As Boolean

    If isBuy Then
        Select Case stockRegime
            Case "TRENDING_BULL"
                minRsi = 50: extHigh = 80
            Case "MILD_UP"
                minRsi = 45: extHigh = 70
            Case "RANGING"
                minRsi = 50: extHigh = 75
            Case Else
                minRsi = 50: extHigh = 80
        End Select

        If rsi5 < minRsi Then
            rejectReason = "RSI_WEAK (RSI5=" & Format(rsi5, "0.0") & " < " & minRsi & ", Regime=" & stockRegime & ")"
            Exit Function
        End If
        isExtended = (rsi5 >= extHigh)
    Else
        Select Case stockRegime
            Case "TRENDING_BEAR"
                maxRsi = 50: extLow = 20
            Case "MILD_DOWN"
                maxRsi = 55: extLow = 30
            Case "RANGING"
                maxRsi = 50: extLow = 25
            Case Else
                maxRsi = 50: extLow = 20
        End Select

        If rsi5 > maxRsi Then
            rejectReason = "RSI_WEAK (RSI5=" & Format(rsi5, "0.0") & " > " & maxRsi & ", Regime=" & stockRegime & ")"
            Exit Function
        End If
        isExtended = (rsi5 <= extLow)
    End If

    If isExtended Then
        Dim er10Val As Double: er10Val = GetStockER10(wsSrc, ticker)
        If er10Val < GATE_RSI_STRONG_ER10_MIN Or volRatio < GATE_RSI_STRONG_VOL_MIN Then
            rejectReason = "RSI_EXTENDED_WEAK_Q (RSI5=" & Format(rsi5, "0.0") & _
                           ", ER10=" & Format(er10Val, "0.00") & _
                           ", VolRatio=" & Format(volRatio, "0.00") & ")"
            Exit Function
        End If
        rsiStatus = "RSI_EXTENDED_ALLOWED"
    End If

    PassRSIContextGate_V2 = True
End Function

Private Function RangeRef(ByVal letter As String) As Long
    RangeRef = Range(letter & "1").Column
End Function

Private Sub ColorCompareRun(wsComp As Worksheet, startRow As Long, endRow As Long, ByVal cls As String)
    If startRow > endRow Then Exit Sub
    Dim colr As Long
    Select Case cls
        Case "BOTH": colr = RGB(235, 241, 222)
        Case "V1_ONLY": colr = RGB(253, 233, 217)
        Case Else: colr = RGB(220, 230, 241)
    End Select
    wsComp.Range("A" & startRow & ":I" & endRow).Interior.Color = colr
End Sub


' =====================================================================
' Private Sub CalculateConvictionScore_V2
' Evaluates frozen multi-factor conviction score (0-100) and assigns CAT.
' Same exact logic used for new trades and open trade 5-minute refreshes.
' =====================================================================
Private Sub CalculateConvictionScore_V2(wsSrc As Worksheet, srcRow As Long, isBuy As Boolean, _
        ByRef outScore As Double, ByRef outCAT As String, ByRef outState As String, _
        ByRef outRsi15 As Double, ByRef outAdx As Double, ByRef outEr10 As Double, _
        ByRef outRvol As Double, ByRef outDistAtr As Double)

    Dim ticker As String: ticker = CStr(wsSrc.cells(srcRow, "A").value)
    Dim ltp As Double: ltp = CDbl(val(wsSrc.cells(srcRow, "T").value))
    Dim atr5m As Double: atr5m = CDbl(val(wsSrc.cells(srcRow, "CC").value))

    outRsi15 = CDbl(val(wsSrc.cells(srcRow, "V").value))
    outAdx = CDbl(val(wsSrc.cells(srcRow, "CB").value))
    outEr10 = CDbl(val(wsSrc.cells(srcRow, "DP").value))
    outRvol = CDbl(val(wsSrc.cells(srcRow, "BO").value))
    If atr5m <= 0 Then atr5m = GetATR5m(wsSrc, ticker)
    If outEr10 <= 0 Then outEr10 = 0.5

    Dim openPx As Double: openPx = CDbl(val(wsSrc.cells(srcRow, "G").value))
    Dim highPx As Double: highPx = CDbl(val(wsSrc.cells(srcRow, "H").value))
    Dim lowPx As Double: lowPx = CDbl(val(wsSrc.cells(srcRow, "I").value))
    Dim closePx As Double: closePx = CDbl(val(wsSrc.cells(srcRow, "J").value))
    If openPx <= 0 Then openPx = ltp
    If highPx <= 0 Then highPx = ltp
    If lowPx <= 0 Then lowPx = ltp
    If closePx <= 0 Then closePx = ltp

    Dim bxPacket As String: bxPacket = CStr(wsSrc.cells(srcRow, "BX").value)
    outState = CStr(wsSrc.cells(srcRow, "BY").value)
    If Len(outState) = 0 Then outState = "DONCHIAN_EVENT"

    ' Parse age, valid, and dist_atr from BX packet if present
    Dim evtAge As Long: evtAge = 0
    Dim evtValid As Long: evtValid = 1
    outDistAtr = 0.35

    If Len(bxPacket) > 0 Then
        Dim parts() As String: parts = Split(bxPacket, "|")
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            Dim kv() As String: kv = Split(parts(i), "=")
            If UBound(kv) >= 1 Then
                Select Case LCase$(Trim$(kv(0)))
                    Case "age": evtAge = CLng(val(kv(1)))
                    Case "valid": evtValid = CLng(val(kv(1)))
                    Case "dist_atr": outDistAtr = Abs(CDbl(val(kv(1))))
                End Select
            End If
        Next i
    End If

    If outDistAtr <= 0 And atr5m > 0 Then
        outDistAtr = Abs(ltp - CDbl(val(wsSrc.cells(srcRow, "E").value))) / atr5m
    End If

    Dim score As Double: score = 0#

    ' Step 1: Freshness (BX/BY parity)
    If evtAge = 0 And evtValid = 1 Then
        score = score + 15#
    Else
        score = score - 10#
    End If

    If isBuy Then
        ' Hard Rejection Gate
        If outRsi15 < 50# Or outEr10 < 0.25 Then
            outScore = -999#
            outCAT = "REJECT"
            Exit Sub
        End If

        ' RSI 15m
        If outRsi15 >= 80# Then
            score = score + 25#
        ElseIf outRsi15 >= 70# Then
            score = score + 20#
        ElseIf outRsi15 >= 60# Then
            score = score + 10#
        End If

        ' ADX
        If outAdx >= 35# Then
            score = score + 15#
        ElseIf outAdx >= 25# Then
            score = score + 5#
        End If

        ' ER10 Sweet Spot
        If outEr10 >= 0.25 And outEr10 <= 0.45 Then
            score = score + 15#
        ElseIf outEr10 > 0.45 And outEr10 <= 0.6 Then
            score = score + 8#
        ElseIf outEr10 > 0.6 And outEr10 <= 0.75 Then
            score = score + 2#
        End If

        ' Breakout Extension
        If outDistAtr >= 0.25 And outDistAtr <= 0.75 Then
            score = score + 15#
        ElseIf (outDistAtr >= 0.1 And outDistAtr < 0.25) Or (outDistAtr > 0.75 And outDistAtr <= 1#) Then
            score = score + 8#
        End If

        ' Candle Quality
        Dim cRange As Double: cRange = highPx - lowPx
        If cRange > 0 Then
            Dim cBody As Double: cBody = Abs(closePx - openPx)
            Dim bodyPct As Double: bodyPct = cBody / cRange * 100#
            Dim closeLoc As Double: closeLoc = (closePx - lowPx) / cRange * 100#
            If bodyPct >= 60# And closeLoc >= 85# Then score = score + 10#
        End If

        ' RVOL
        If outRvol >= 1.8 And outRvol <= 3.5 Then score = score + 10#

        ' Timing (Afternoon Expansion)
        If TimeValue(Format(Now, "hh:mm:ss")) >= TimeValue("13:30:00") Then score = score + 10#

    Else ' SELL
        ' Hard Rejection Gate
        If outRsi15 > 50# Or outEr10 < 0.25 Then
            outScore = -999#
            outCAT = "REJECT"
            Exit Sub
        End If

        ' RSI 15m Bearish
        If outRsi15 <= 25# Then
            score = score + 25#
        ElseIf outRsi15 <= 35# Then
            score = score + 20#
        ElseIf outRsi15 <= 45# Then
            score = score + 10#
        End If

        ' ADX
        If outAdx >= 35# Then
            score = score + 15#
        ElseIf outAdx >= 25# Then
            score = score + 5#
        End If

        ' ER10
        If outEr10 >= 0.25 And outEr10 <= 0.45 Then
            score = score + 15#
        ElseIf outEr10 > 0.45 And outEr10 <= 0.6 Then
            score = score + 8#
        End If

        ' Breakdown Extension
        If outDistAtr >= 0.25 And outDistAtr <= 0.75 Then
            score = score + 15#
        ElseIf (outDistAtr >= 0.1 And outDistAtr < 0.25) Or (outDistAtr > 0.75 And outDistAtr <= 1#) Then
            score = score + 8#
        End If

        ' Candle Quality Bearish
        Dim cRangeSell As Double: cRangeSell = highPx - lowPx
        If cRangeSell > 0 Then
            Dim cBodySell As Double: cBodySell = Abs(closePx - openPx)
            Dim bodyPctSell As Double: bodyPctSell = cBodySell / cRangeSell * 100#
            Dim closeLocSell As Double: closeLocSell = (highPx - closePx) / cRangeSell * 100#
            If bodyPctSell >= 60# And closeLocSell >= 85# Then score = score + 10#
        End If

        ' RVOL
        If outRvol >= 1.8 And outRvol <= 3.5 Then score = score + 10#

        ' Timing
        If TimeValue(Format(Now, "hh:mm:ss")) >= TimeValue("13:30:00") Then score = score + 10#
    End If

    ' Cap score to max 100 if > 100, or keep 0
    If score < 0 Then
        outScore = 0#
    ElseIf score > 100 Then
        outScore = 100#
    Else
        outScore = score
    End If

    ' Category Mapping
    If outScore >= 70# Then
        outCAT = "CAT1"
    ElseIf outScore >= 55# Then
        outCAT = "CAT2"
    ElseIf outScore >= 40# Then
        outCAT = "CAT3"
    Else
        outCAT = "CAT4"
    End If
End Sub

' =====================================================================
' Private Sub RefreshOpenTradeConviction_V2
' Refreshes conviction for all OPEN trades every 5 minutes.
' Maintains Initial CAT/Score, updates Current CAT/Score, and logs transitions.
' Does NOT create duplicate trades, close, or resize any trade.
' =====================================================================
Private Sub RefreshOpenTradeConviction_V2(wsC2 As Worksheet, wsSrc As Worksheet)
    Dim lastRow As Long, r As Long
    lastRow = wsC2.cells(wsC2.Rows.count, 1).End(xlUp).row
    If lastRow < DATA_START_ROW Then Exit Sub

    For r = DATA_START_ROW To lastRow
        Dim closeTimeVal As Variant: closeTimeVal = wsC2.Range(CK_CLOSE_TIME & r).value
        If Not IsEmpty(closeTimeVal) And Len(CStr(closeTimeVal)) > 0 Then GoTo NextOpenTrade

        Dim ticker As String: ticker = CStr(wsC2.Range(CK_TICKER & r).value)
        If Len(ticker) = 0 Then GoTo NextOpenTrade

        Dim isBuy As Boolean: isBuy = (UCase$(CStr(wsC2.Range(CK_ACTION & r).value)) = "BUY")
        Dim openTime As String: openTime = CStr(wsC2.Range(CK_OPEN_TIME & r).value)
        Dim stratName As String: stratName = CStr(wsC2.Range(CK_STRATEGY & r).value)
        Dim candleCount As Long: candleCount = CLng(val(wsC2.Range(CK_CANDLE_COUNT & r).value))

        Dim initCAT As String: initCAT = CStr(wsC2.Range(CK_INITIAL_CAT & r).value)
        Dim initScore As Double: initScore = CDbl(val(wsC2.Range(CK_INITIAL_SCORE & r).value))
        Dim prevCAT As String: prevCAT = CStr(wsC2.Range(CK_CURRENT_CAT & r).value)
        Dim prevScore As Double: prevScore = CDbl(val(wsC2.Range(CK_CURRENT_SCORE & r).value))
        Dim catTransitions As Long: catTransitions = CLng(val(wsC2.Range(CK_CAT_TRANSITIONS & r).value))
        Dim scoreChanges As Long: scoreChanges = CLng(val(wsC2.Range(CK_SCORE_CHANGES & r).value))

        If mRowLookup Is Nothing Then Set mRowLookup = BuildRowLookup(wsSrc)
        If Not mRowLookup.Exists(ticker) Then GoTo NextOpenTrade

        Dim srcRow As Long: srcRow = CLng(mRowLookup(ticker))
        If srcRow <= 0 Then GoTo NextOpenTrade

        Dim curScore As Double, curCAT As String, curDonchState As String
        Dim curRsi15 As Double, curADX As Double, curEr10 As Double, curRvol As Double, curDistAtr As Double
        CalculateConvictionScore_V2 wsSrc, srcRow, isBuy, curScore, curCAT, curDonchState, _
                                   curRsi15, curADX, curEr10, curRvol, curDistAtr

        ' Initialize empty fields if legacy trade
        If Len(initCAT) = 0 Then
            initCAT = curCAT
            initScore = curScore
            prevCAT = curCAT
            prevScore = curScore
            wsC2.Range(CK_INITIAL_CAT & r).value = initCAT
            wsC2.Range(CK_INITIAL_SCORE & r).value = initScore
        End If

        ' Update Current State Columns
        wsC2.Range(CK_CURRENT_CAT & r).value = curCAT
        wsC2.Range(CK_CURRENT_SCORE & r).value = curScore

        ' Track CAT Transitions (Only increment when CAT actually changes)
        Dim catChanged As Boolean: catChanged = (curCAT <> prevCAT)
        If catChanged Then
            catTransitions = catTransitions + 1
            wsC2.Range(CK_CAT_TRANSITIONS & r).value = catTransitions
        End If

        ' Track Score Changes (Increment when score shifts by >= 5 points)
        Dim scoreChanged As Boolean: scoreChanged = (Abs(curScore - prevScore) >= 5#)
        If scoreChanged Then
            scoreChanges = scoreChanges + 1
            wsC2.Range(CK_SCORE_CHANGES & r).value = scoreChanges
        End If

        ' Update Visible Comment Prefix & re-apply CAT character formatting
        Dim curComment As String: curComment = CStr(wsC2.Range(CK_COMENTS & r).value)
        Dim newPrefix As String
        newPrefix = "[" & curCAT & "|S=" & Format(curScore, "0") & "|INIT=" & initCAT & "|IS=" & Format(initScore, "0") & "] "

        Dim updatedComment As String
        If Left$(curComment, 1) = "[" And InStr(curComment, "]") > 0 Then
            Dim closeBracketPos As Long: closeBracketPos = InStr(curComment, "]")
            updatedComment = newPrefix & Trim$(Mid$(curComment, closeBracketPos + 1))
        Else
            updatedComment = newPrefix & curComment
        End If

        If updatedComment <> curComment Then
            wsC2.Range(CK_COMENTS & r).value = updatedComment
            ApplyCATColorCoding_V2 wsC2, r, curCAT
        ElseIf catChanged Then
            ApplyCATColorCoding_V2 wsC2, r, curCAT
        End If

        ' Audit Log: Record every CAT transition (and material score changes >= 10 pts)
'        If catChanged Or Abs(curScore - prevScore) >= 10# Then
'            LogConvictionTransition_V2 wsSrc, ticker, openTime, initCAT, prevCAT, curCAT, _
'                                      initScore, prevScore, curScore, candleCount, _
'                                      IIf(isBuy, "BUY", "SELL"), stratName, _
'                                      curRsi15, curAdx, curEr10, curRvol, curDonchState
'        End If

NextOpenTrade:
    Next r
End Sub

' =====================================================================
' Private Sub LogConvictionTransition_V2
' Records conviction trajectory shifts to COCKPIT_V2_TRANSITIONS sheet.
' =====================================================================
Private Sub LogConvictionTransition_V2(wsSrc As Worksheet, ticker As String, openTime As String, _
        initCAT As String, prevCAT As String, curCAT As String, _
        initScore As Double, prevScore As Double, curScore As Double, barsOpen As Long, _
        side As String, strategy As String, rsi15 As Double, adxVal As Double, er10Val As Double, _
        rvolVal As Double, donchState As String)

    On Error Resume Next
    Dim wsTrans As Worksheet
    Set wsTrans = GetOrCreateSheet_V2("COCKPIT_V2_TRANSITIONS")
    If wsTrans Is Nothing Then Exit Sub

    Dim NextRow As Long
    NextRow = wsTrans.cells(wsTrans.Rows.count, 1).End(xlUp).row + 1
    If NextRow < 5 Then
        NextRow = 5
        wsTrans.cells(4, 1).value = "Timestamp"
        wsTrans.cells(4, 2).value = "Ticker"
        wsTrans.cells(4, 3).value = "OpenTime"
        wsTrans.cells(4, 4).value = "InitCAT"
        wsTrans.cells(4, 5).value = "PrevCAT"
        wsTrans.cells(4, 6).value = "CurCAT"
        wsTrans.cells(4, 7).value = "InitScore"
        wsTrans.cells(4, 8).value = "PrevScore"
        wsTrans.cells(4, 9).value = "CurScore"
        wsTrans.cells(4, 10).value = "BarsOpen"
        wsTrans.cells(4, 11).value = "Side"
        wsTrans.cells(4, 12).value = "Strategy"
        wsTrans.cells(4, 13).value = "RSI15"
        wsTrans.cells(4, 14).value = "ADX"
        wsTrans.cells(4, 15).value = "ER10"
        wsTrans.cells(4, 16).value = "RVOL"
        wsTrans.cells(4, 17).value = "DonchianState"
        wsTrans.Range("A4:Q4").Font.Bold = True
        wsTrans.Range("A4:Q4").Interior.Color = RGB(220, 230, 241)
    End If

    wsTrans.cells(NextRow, 1).value = Format(Now, "YYYY-MM-DD HH:MM:SS")
    wsTrans.cells(NextRow, 2).value = ticker
    wsTrans.cells(NextRow, 3).value = openTime
    wsTrans.cells(NextRow, 4).value = initCAT
    wsTrans.cells(NextRow, 5).value = prevCAT
    wsTrans.cells(NextRow, 6).value = curCAT
    wsTrans.cells(NextRow, 7).value = initScore
    wsTrans.cells(NextRow, 8).value = prevScore
    wsTrans.cells(NextRow, 9).value = curScore
    wsTrans.cells(NextRow, 10).value = barsOpen
    wsTrans.cells(NextRow, 11).value = side
    wsTrans.cells(NextRow, 12).value = strategy
    wsTrans.cells(NextRow, 13).value = rsi15
    wsTrans.cells(NextRow, 14).value = adxVal
    wsTrans.cells(NextRow, 15).value = er10Val
    wsTrans.cells(NextRow, 16).value = rvolVal
    wsTrans.cells(NextRow, 17).value = donchState

    ' Color coding for transitions
    If curCAT = "CAT1" Then
        wsTrans.Range("A" & NextRow & ":Q" & NextRow).Interior.Color = RGB(235, 247, 235)
    ElseIf curCAT = "CAT4" Or curCAT = "REJECT" Then
        wsTrans.Range("A" & NextRow & ":Q" & NextRow).Interior.Color = RGB(255, 240, 240)
    End If
End Sub

' =====================================================================
' Private Sub ApplyCATColorCoding_V2
' Formats the comment cell with distinct conviction color coding.
' =====================================================================
Private Sub ApplyCATColorCoding_V2(wsC2 As Worksheet, r As Long, catVal As String)
    On Error Resume Next
    Dim cellRng As Range
    Set cellRng = wsC2.Range(CK_COMENTS & r)
    
    Dim txt As String: txt = CStr(cellRng.value)
    If Len(txt) = 0 Then Exit Sub
    
    ' Neutralize background and set base font to clean black
    cellRng.Interior.Pattern = xlNone
    cellRng.Font.Color = RGB(0, 0, 0)
    cellRng.Font.Bold = False
    
    ' Locate the primary CAT tag within the comment string [CAT...
    Dim cats As Variant
    'Dim catVal As String
    Dim catPos As Long
    Dim startPos As Long
    Dim i As Long
    
    cats = Array("CAT1", "CAT2", "CAT3", "CAT4", "REJECT")
    
    For i = LBound(cats) To UBound(cats)
        catVal = cats(i)
        startPos = 1
        
        Do
            catPos = InStr(startPos, txt, catVal)
            If catPos = 0 Then Exit Do
            
            With cellRng.Characters(Start:=catPos, Length:=Len(catVal)).Font
                .Bold = True
                
                Select Case catVal
                    Case "CAT1"
                        .Color = RGB(0, 128, 0)        ' Green
                    Case "CAT2"
                        .Color = RGB(255, 191, 0)      ' Amber
                    Case "CAT3"
                        .Color = RGB(255, 102, 102)    ' Light Red
                    Case "CAT4"
                        .Color = RGB(220, 20, 60)      ' Red
                    Case "REJECT"
                        .Color = RGB(204, 0, 0)        ' Dark Red
                End Select
            End With
            
            startPos = catPos + Len(catVal)
        Loop
    Next i
End Sub




' =====================================================================
' Private Function BuildTickerSessionAttemptLookup
' Builds session-scoped dictionary: Ticker -> Count of attempts TODAY.
' Counts both OPEN and CLOSED trades across BUY + SELL combined.
' Rejects rows with missing or non-today date timestamps.
' =====================================================================
Private Function BuildTickerSessionAttemptLookup(wsC2 As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = wsC2.cells(wsC2.Rows.count, 1).End(xlUp).row
    If lastRow < DATA_START_ROW Then
        Set BuildTickerSessionAttemptLookup = dict
        Exit Function
    End If

    Dim todayDate As Date: todayDate = Date
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim ticker As String: ticker = Trim$(CStr(wsC2.Range(CK_TICKER & r).value))
        If Len(ticker) > 0 Then
            Dim openTimeVal As String: openTimeVal = Trim$(CStr(wsC2.Range(CK_OPEN_TIME & r).value))
            Dim isToday As Boolean: isToday = False

            If IsDate(openTimeVal) Then
                If DateValue(CDate(openTimeVal)) = todayDate Then
                    isToday = True
                End If
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



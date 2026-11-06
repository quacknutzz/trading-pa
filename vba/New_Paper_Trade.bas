Attribute VB_Name = "New_Paper_Trade"
'==========================================================
' PAPER TRADE LIFECYCLE
'
' PENDING
'    -> waits until latest candle trades through Entry
' OPEN
'    -> only execution manager may set this
' CLOSED / CANCELLED
'
'==========================================================

Option Explicit

' =========================================================================================
' CONFIGURATION & CONSTANTS
' =========================================================================================
Const SH_MASTER As String = "AI_Master"
Const SH_RAW As String = "Sheet1"
Const SH_INTRADAY As String = "New_Logic_Intraday"
Const SH_TRADES As String = "Paper_Trades"
Const SH_BULLETIN As String = "Bulletin"
Const SH_AUDIT As String = "Paper_Trade_Audit"

' --- Tweakable Parameters ---
Const ATR_STOP_MULT As Double = 2#
Const LOG_SKIPPED_TRADES As Boolean = False

' --- Column Maps (AI_Master - fixed absolute columns) ---
Const C_LTP As String = "C"
Const C_RVOL As String = "S"
Const C_ATR As String = "Y"
Const C_SECTOR As String = "L"
Const C_SECTOR_SCORE As String = "N"
Const C_SECTOR_REGIME As String = "O"
Const C_SIGNAL_CODE As String = "D"
Const C_ACTIONABILITY As String = "F"
Const C_CONFIDENCE As String = "H"
Const C_GRADE As String = "I"
Const C_REASON As String = "K"
Const C_ENTRY As String = "AB"
Const C_TARGET As String = "AC"
Const C_STOP As String = "AD"
Const C_SETUP As String = "J"
Const C_SUPERTREND As String = "W"
Const C_VWAP As String = "U"
Const C_SPIKE As String = "V"
Const C_INST_SUMMARY As String = "AI"

' --- Raw Sheet1 Columns (dashboard / bulletin / index context only) ---
Const RAW_LTP As String = "T"
Const RAW_CHANGE_PTS As String = "R"
Const RAW_ATR As String = "CC"
Const RAW_SUPERTREND As String = "BT"
Const RAW_PIVOT As String = "N"
Const RAW_R1 As String = "O"
Const RAW_S1 As String = "M"
Const RAW_S2 As String = "L"
Const RAW_R2 As String = "P"
Const RAW_EMA20_5M As String = "AT"
Const RAW_INST_SUMMARY As String = "DQ"

' --- Paper Trades Map ---
Const P_TICKER As Long = 1
Const P_DIR As Long = 2
Const P_DATE As Long = 3
Const P_ENTRY As Long = 4
Const P_TARGET As Long = 5
Const P_SL As Long = 6
Const P_STATUS As Long = 7
Const P_SIG_TIME As Long = 8
Const P_OPEN_TIME As Long = 9
Const P_CLOSE_TIME As Long = 10
Const P_COMM As Long = 11
Const P_ACTION As Long = 12
Const P_LTP As Long = 13
Const P_COACH As Long = 14

' --- Start Rows ---
Const ROW_START_M As Long = 4
Const ROW_START_I As Long = 3
Const ROW_START_P As Long = 5

' --- Custom Types ---
Type MacroState
    niftyTrend As String
    sectorTrend As String
End Type

Private gVbaScannedRows As Long
Private gVbaEligibleRows As Long
Private gVbaSkippedZeroLevels As Long
Private gVbaSkippedActionability As Long
Private gVbaInsertedRows As Long

Private Sub Reset_VBA_Audit_Counters()
    gVbaScannedRows = 0
    gVbaEligibleRows = 0
    gVbaSkippedZeroLevels = 0
    gVbaSkippedActionability = 0
    gVbaInsertedRows = 0
End Sub

Private Function Ensure_Audit_Sheet() As Worksheet
    On Error Resume Next
    Set Ensure_Audit_Sheet = ThisWorkbook.Worksheets(SH_AUDIT)
    On Error GoTo 0
    If Ensure_Audit_Sheet Is Nothing Then
        Set Ensure_Audit_Sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        Ensure_Audit_Sheet.Name = SH_AUDIT
        Ensure_Audit_Sheet.Range("A1:G1").value = Array("Timestamp", "Ticker", "Stage", "Actionability", "SignalCode", "Inserted", "Reason")
        Ensure_Audit_Sheet.Rows(1).Font.Bold = True
    End If
End Function

Private Sub Audit_Paper_Decision(ByVal ticker As String, ByVal stageName As String, ByVal actionability As String, ByVal sigCode As String, ByVal inserted As Boolean, ByVal reasonText As String)
    If LOG_SKIPPED_TRADES = False Then Exit Sub
    On Error Resume Next
    Dim wsA As Worksheet: Set wsA = Ensure_Audit_Sheet()
    If wsA Is Nothing Then Exit Sub
    Dim nr As Long: nr = wsA.cells(wsA.Rows.count, 1).End(xlUp).row + 1
    wsA.cells(nr, 1).value = Now
    wsA.cells(nr, 2).value = ticker
    wsA.cells(nr, 3).value = stageName
    wsA.cells(nr, 4).value = actionability
    wsA.cells(nr, 5).value = sigCode
    wsA.cells(nr, 6).value = IIf(inserted, "YES", "NO")
    wsA.cells(nr, 7).value = Left(reasonText, 240)
End Sub

Private Sub Write_VBA_Audit_Summary()
    If LOG_SKIPPED_TRADES = False Then Exit Sub
    On Error Resume Next
    Dim wsA As Worksheet: Set wsA = Ensure_Audit_Sheet()
    If wsA Is Nothing Then Exit Sub
    Dim nr As Long: nr = wsA.cells(wsA.Rows.count, 1).End(xlUp).row + 1
    wsA.cells(nr, 1).value = Now
    wsA.cells(nr, 2).value = "ALL"
    wsA.cells(nr, 3).value = "SUMMARY"
    wsA.cells(nr, 4).value = "ENTER_ROWS=" & gVbaEligibleRows
    wsA.cells(nr, 5).value = "SCANNED=" & gVbaScannedRows
    wsA.cells(nr, 6).value = "INSERTED=" & gVbaInsertedRows
    wsA.cells(nr, 7).value = "VBA scanned rows: " & gVbaScannedRows & _
                               " | VBA eligible rows actionability=ENTER: " & gVbaEligibleRows & _
                               " | VBA skipped zero levels: " & gVbaSkippedZeroLevels & _
                               " | VBA skipped actionability: " & gVbaSkippedActionability & _
                               " | VBA inserted: " & gVbaInsertedRows
End Sub

' =========================================================================================
' NARRATIVE HELPERS (Deterministic, Trader-Readable)
' =========================================================================================
Private Function Extract_Tag(ByVal s As String, ByVal key As String) As String
    Dim p As Long: p = InStr(1, s, key, vbTextCompare)
    If p = 0 Then Extract_Tag = "": Exit Function
    Dim startPos As Long: startPos = p + Len(key)
    Dim tail As String: tail = Mid$(s, startPos)
    Dim cut As Long
    cut = InStr(1, tail, "|")
    If cut = 0 Then cut = InStr(1, tail, "||")
    If cut > 0 Then tail = Left$(tail, cut - 1)
    tail = Trim$(tail)
    ' Some tags are followed by parentheses or extra tokens (e.g., "Flow:Bullish (MinConf:...)").
    If InStr(1, key, "Exec:", vbTextCompare) = 0 _
       And InStr(1, key, "Setup:", vbTextCompare) = 0 _
       And InStr(1, key, "Pattern:", vbTextCompare) = 0 _
       And InStr(1, key, "Risk:", vbTextCompare) = 0 _
       And InStr(1, key, "Hint:", vbTextCompare) = 0 _
       And InStr(1, key, "SECTOR_", vbTextCompare) = 0 Then
        Dim p2 As Long
        p2 = InStr(1, tail, "(")
        If p2 > 0 Then tail = Trim$(Left$(tail, p2 - 1))
        p2 = InStr(1, tail, " ")
        If p2 > 0 Then tail = Trim$(Left$(tail, p2 - 1))
    End If
    Extract_Tag = Trim$(tail)
End Function

Private Function Extract_Last_Number(ByVal s As String, ByVal key As String) As String
    Dim p As Long, lastP As Long: lastP = 0
    p = InStr(1, s, key, vbTextCompare)
    Do While p > 0
        lastP = p
        p = InStr(p + 1, s, key, vbTextCompare)
    Loop
    If lastP = 0 Then Extract_Last_Number = "": Exit Function
    Dim startPos As Long: startPos = lastP + Len(key)
    Dim tail As String: tail = Mid$(s, startPos)
    Dim out As String: out = ""
    Dim i As Long, ch As String
    For i = 1 To Len(tail)
        ch = Mid$(tail, i, 1)
        If (ch Like "[0-9]") Or ch = "." Or ch = "-" Then
            out = out & ch
        Else
            Exit For
        End If
    Next i
    Extract_Last_Number = Trim$(out)
End Function

Private Function Extract_Lifecycle_From_Cockpit(ByVal commVal As String) As String
    Dim normalized As String: normalized = Replace(commVal, vbCr, "")
    Dim lines() As String: lines = Split(normalized, vbLf)
    If UBound(lines) >= 1 Then
        Dim parts() As String: parts = Split(lines(1), "|")
        If UBound(parts) >= 0 Then
            Extract_Lifecycle_From_Cockpit = UCase(Trim(parts(0)))
            Exit Function
        End If
    End If
    Extract_Lifecycle_From_Cockpit = ""
End Function

Private Function Extract_Execution_From_Cockpit(ByVal commVal As String) As String
    Dim normalized As String: normalized = Replace(commVal, vbCr, "")
    Dim lines() As String: lines = Split(normalized, vbLf)
    If UBound(lines) >= 1 Then
        Dim parts() As String: parts = Split(lines(1), "|")
        If UBound(parts) >= 3 Then
            Extract_Execution_From_Cockpit = UCase(Trim(parts(3)))
            Exit Function
        End If
    End If
    Extract_Execution_From_Cockpit = ""
End Function


Private Function Derive_Pattern_Label(ByVal setup As String, ByVal reasonStr As String, ByVal scienceStr As String, ByVal pState As String) As String
    Dim blob As String
    blob = UCase$(setup & " | " & reasonStr & " | " & scienceStr & " | " & pState)
    If InStr(blob, "DECAY") > 0 Or InStr(blob, "EXIT") > 0 Then
        Derive_Pattern_Label = "Decay Exit"
    ElseIf InStr(blob, "BREAKOUT") > 0 Then
        Derive_Pattern_Label = "Momentum Breakout"
    ElseIf InStr(blob, "BREAKDOWN") > 0 Then
        Derive_Pattern_Label = "Breakdown"
    ElseIf InStr(blob, "PULLBACK") > 0 Then
        Derive_Pattern_Label = "Pullback"
    ElseIf InStr(blob, "REVERSAL") > 0 Or InStr(blob, "MEAN_REVERSION") > 0 Then
        Derive_Pattern_Label = "Reversal"
    ElseIf InStr(blob, "RECOVERY") > 0 Then
        Derive_Pattern_Label = "Recovery"
    ElseIf InStr(blob, "MOMENTUM") > 0 Or InStr(blob, "TREND") > 0 Or InStr(blob, "TOP_BUY") > 0 Or InStr(blob, "TOP_SELL") > 0 Then
        Derive_Pattern_Label = "Momentum"
    ElseIf InStr(blob, "WATCH") > 0 Then
        Derive_Pattern_Label = "Watch Setup"
    Else
        Derive_Pattern_Label = "Scanning"
    End If
End Function

Private Function Derive_Risk_Label(ByVal reasonStr As String, ByVal setup As String, ByVal execTag As String, ByVal netRR As String, ByVal mStatus As String) As String
    Dim blob As String
    blob = UCase$(reasonStr & " | " & setup & " | " & execTag & " | " & mStatus)
    If InStr(blob, "CAP") > 0 Or InStr(blob, "DIVERSIFICATION") > 0 Then
        Derive_Risk_Label = "Diversification Cap"
    ElseIf InStr(blob, "EXHAUST") > 0 Or InStr(blob, "STRETCH") > 0 Or InStr(blob, "OVEREXTENDED") > 0 Then
        Derive_Risk_Label = "Exhausted"
    ElseIf InStr(blob, "NET_RR_TOO_LOW") > 0 Then
        Derive_Risk_Label = "Poor Reward"
    ElseIf InStr(blob, "NO_TRIGGER") > 0 Or InStr(blob, "NOT_TRIGGERED") > 0 Or InStr(blob, "BAD_ENTRY") > 0 Then
        Derive_Risk_Label = "No Trigger"
    ElseIf InStr(blob, "WEAK_TREND") > 0 Then
        Derive_Risk_Label = "Weak Trend"
    ElseIf InStr(blob, "FILTERED") > 0 Then
        Derive_Risk_Label = "Filtered"
    ElseIf InStr(blob, "DECAY") > 0 Or InStr(blob, "WITHDRAWN") > 0 Then
        Derive_Risk_Label = "Thesis Decay"
    ElseIf val(netRR) > 0 And val(netRR) < 0.5 Then
        Derive_Risk_Label = "Poor Reward"
    Else
        Derive_Risk_Label = "Manage Risk"
    End If
End Function

Private Function Derive_Corrective_Hint(ByVal actionability As String, ByVal pDir As String, ByVal patternLabel As String, ByVal riskLabel As String) As String
    Dim act As String
    act = UCase$(Trim$(actionability))
    If act = "REVIEW_EXIT" Then
        Derive_Corrective_Hint = "Review exit now"
    ElseIf riskLabel = "Diversification Cap" Then
        Derive_Corrective_Hint = "Lower priority; keep on watch"
    ElseIf riskLabel = "Exhausted" Then
        Derive_Corrective_Hint = "Do not chase"
    ElseIf riskLabel = "Poor Reward" Then
        Derive_Corrective_Hint = "Wait for better entry"
    ElseIf riskLabel = "No Trigger" Then
        Derive_Corrective_Hint = "Wait for trigger"
    ElseIf riskLabel = "Weak Trend" Then
        Derive_Corrective_Hint = "Watch only"
    ElseIf act = "ENTER_LONG" Or act = "ENTER_SHORT" Then
        If patternLabel = "Momentum Breakout" Then
            Derive_Corrective_Hint = "Act only on follow-through"
        ElseIf patternLabel = "Reversal" Then
            Derive_Corrective_Hint = "Use starter size"
        Else
            Derive_Corrective_Hint = "Enter with defined risk"
        End If
    ElseIf act = "WATCHLIST" Or act = "PENDING_LONG" Or act = "PENDING_SHORT" Then
        Derive_Corrective_Hint = "Stay ready; no chase"
    ElseIf act = "MANAGE_POSITION" Then
        Derive_Corrective_Hint = "Trail risk, do not add"
    Else
        Derive_Corrective_Hint = "Stay selective"
    End If
End Function

Private Function Derive_Execution_Style(ByVal pPattern As String, ByVal pState As String) As String
    Dim pat As String: pat = UCase(Trim(pPattern))
    Dim st As String: st = UCase(Trim(pState))
    
    If st = "BROKEN" Then
        Derive_Execution_Style = "Failed"
    ElseIf InStr(pat, "BREAKOUT") > 0 Or InStr(pat, "MOMENTUM") > 0 Then
        Derive_Execution_Style = "Breakout"
    ElseIf InStr(pat, "PULLBACK") > 0 Or InStr(pat, "RECOVERY") > 0 Then
        Derive_Execution_Style = "Pullback"
    ElseIf InStr(pat, "REVERSAL") > 0 Or InStr(pat, "MEAN_REVERSION") > 0 Then
        Derive_Execution_Style = "Reversal"
    ElseIf InStr(pat, "BREAKDOWN") > 0 Then
        Derive_Execution_Style = "Breakdown"
    ElseIf InStr(pat, "DECAY") > 0 Or InStr(pat, "FADING") > 0 Then
        Derive_Execution_Style = "Fading"
    Else
        Derive_Execution_Style = "Breakout"
    End If
End Function

Private Function Derive_Signal_Book(ByVal reasonStr As String, ByVal setup As String) As String
    Dim book As String
    book = Extract_Tag(reasonStr, "Book:")
    If book = "" Then book = Extract_Tag(setup, "Book:")
    If book = "" Then book = Extract_Tag(reasonStr, "Bucket:")
    If book = "" Then book = Extract_Tag(setup, "Bucket:")
    
    book = UCase(Trim(book))
    If InStr(book, "EVENT") > 0 Then
        Derive_Signal_Book = "Event"
    ElseIf InStr(book, "TREND") > 0 Then
        Derive_Signal_Book = "Trend"
    ElseIf InStr(book, "REVERSAL") > 0 Or InStr(book, "REVERT") > 0 Then
        Derive_Signal_Book = "Reversal"
    Else
        If InStr(UCase(setup & reasonStr), "EVENT") > 0 Then
            Derive_Signal_Book = "Event"
        ElseIf InStr(UCase(setup & reasonStr), "REVERSAL") > 0 Then
            Derive_Signal_Book = "Reversal"
        Else
            Derive_Signal_Book = "Trend"
        End If
    End If
End Function

Private Function Format_Sector_Code(ByVal sector As String) As String
    Dim s As String: s = UCase(Trim(sector))
    s = Replace(s, "NIFTY ", "")
    s = Replace(s, " ", "")
    If s = "" Then
        Format_Sector_Code = "CNX"
    ElseIf Left(s, 3) = "CNX" Then
        Format_Sector_Code = s
    Else
        Format_Sector_Code = "CNX" & s
    End If
End Function

Private Function Derive_AI_Assessment_Tags(ByVal pDir As String, ByVal rvol As Double, ByVal niftyTrend As String, ByVal sectorTrend As String, ByVal rrVal As Double, ByVal execStyle As String, ByVal stateStr As String, ByVal spikeVal As String, ByVal cfChecks As String, ByVal secScore As Double) As String
    Dim tags(1 To 12) As String
    Dim count As Long: count = 0
    Dim d As String: d = UCase(Trim(pDir))
    
    ' --- 1. Sector Assessment ---
    Dim secT As String: secT = UCase(Trim(sectorTrend))
    If secT = "BULLISH" And d = "BUY" Then
        If secScore > 7 Then
            count = count + 1: tags(count) = "Sector Leader"
        Else
            count = count + 1: tags(count) = "Strong Sector"
        End If
    ElseIf secT = "BEARISH" And d = "SELL" Then
        If secScore < -7 Then
            count = count + 1: tags(count) = "Sector Leader"
        Else
            count = count + 1: tags(count) = "Strong Sector"
        End If
    ElseIf secT = "BEARISH" And d = "BUY" Then
        count = count + 1: tags(count) = "Weak Sector"
    ElseIf secT = "BULLISH" And d = "SELL" Then
        count = count + 1: tags(count) = "Weak Sector"
    End If
    
    ' --- 2. Nifty/Trend Assessment ---
    Dim nifT As String: nifT = UCase(Trim(niftyTrend))
    If nifT = "BULLISH" And d = "BUY" Then
        count = count + 1: tags(count) = "Clean Trend"
    ElseIf nifT = "BEARISH" And d = "SELL" Then
        count = count + 1: tags(count) = "Clean Trend"
    ElseIf nifT = "BEARISH" And d = "BUY" Then
        count = count + 1: tags(count) = "Counter Trend"
    ElseIf nifT = "BULLISH" And d = "SELL" Then
        count = count + 1: tags(count) = "Counter Trend"
    End If
    
    ' --- 3. RVOL Assessment ---
    If rvol > 1.5 Then
        count = count + 1: tags(count) = "High RVOL"
    ElseIf rvol > 0 And rvol < 1# Then
        count = count + 1: tags(count) = "Low RVOL"
    End If
    
    ' --- 4. RR Assessment ---
    If rrVal >= 2# Then
        count = count + 1: tags(count) = "Good RR"
    ElseIf rrVal > 0 And rrVal < 1.3 Then
        count = count + 1: tags(count) = "Low RR"
    End If
    
    ' --- 5. Flow / Momentum / Style ---
    If val(cfChecks) >= 2.5 Then
        count = count + 1: tags(count) = "Strong Flow"
    End If
    
    If execStyle = "Breakout" Then
        count = count + 1: tags(count) = "Breakout"
    ElseIf execStyle = "Pullback" Then
        count = count + 1: tags(count) = "Momentum"
    End If
    
    ' --- 6. Spike / Cautions ---
    If val(spikeVal) > 2# Then
        count = count + 1: tags(count) = "Spike"
    End If
    
    If stateStr = "WEAKENING" Then
        count = count + 1: tags(count) = "Weakening"
        count = count + 1: tags(count) = "Don't Add"
    End If
    
    ' Fallback if no tags derived
    If count = 0 Then
        count = count + 1: tags(count) = "Momentum"
    End If
    
    ' Limit to maximum 4 tags
    Dim limit As Long: limit = IIf(count > 4, 4, count)
    Dim finalStr As String: finalStr = tags(1)
    Dim i As Long
    For i = 2 To limit
        finalStr = finalStr & " | " & tags(i)
    Next i
    
    Derive_AI_Assessment_Tags = finalStr
End Function

Private Function Build_Action_Headline(ByVal actionability As String, ByVal pDir As String) As String
    Dim act As String
    act = UCase$(Trim$(actionability))
    Select Case act
        Case "ENTER_LONG": Build_Action_Headline = "BUY NOW"
        Case "ENTER_SHORT": Build_Action_Headline = "SELL NOW"
        Case "PENDING_LONG": Build_Action_Headline = "SETUP BUY"
        Case "PENDING_SHORT": Build_Action_Headline = "SETUP SELL"
        Case "REVIEW_EXIT": Build_Action_Headline = "EXIT NOW"
        Case "MANAGE_POSITION": Build_Action_Headline = "HOLD"
        Case "WATCHLIST"
            If UCase$(pDir) = "SELL" Then
                Build_Action_Headline = "SETUP SELL"
            Else
                Build_Action_Headline = "SETUP BUY"
            End If
        Case Else
            If UCase$(pDir) = "SELL" Then
                Build_Action_Headline = "SELL NOW"
            ElseIf UCase$(pDir) = "BUY" Then
                Build_Action_Headline = "BUY NOW"
            Else
                Build_Action_Headline = "HOLD"
            End If
    End Select
End Function

' =========================================================================================
' MAIN ORCHESTRATOR
' =========================================================================================
Sub Run_Paper_Trader()
    On Error GoTo ErrorHandler
    
    Dim wsRaw As Worksheet, wsM As Worksheet, wsI As Worksheet, wsP As Worksheet
    Set wsRaw = Sheets(SH_RAW)
    Set wsM = Sheets(SH_MASTER)
    Set wsI = Sheets(SH_INTRADAY)
    Set wsP = Sheets(SH_TRADES)

    If wsRaw Is Nothing Or wsM Is Nothing Or wsI Is Nothing Or wsP Is Nothing Then
        ' SILENT LOG: Avoid UI-blocking in background loop
        If Not wsP Is Nothing Then wsP.Range("Z1").value = "Critical Sheets Missing! | " & Now
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Refresh derived signal layer before VBA scans it.
    wsM.Calculate

    ' 1. DASHBOARD UPDATE (K3 Logic - Preserved)
    Update_K3_Dashboard wsRaw, wsP

    ' 2. MANAGE EXISTING TRADES (SIGNAL ? OPEN ? CLOSED state machine)
    Update_Existing_Records wsP, wsM
    Manage_Open_Trades wsP, wsM

    ' 3. SCAN NEW SIGNALS
    Process_New_Signals wsI, wsM, wsP

    ' 4. VISUALS & HUD
    ColorIndexValues
    Populate_Bulletin_Sheet wsRaw
    wsP.Range("Z1").value = "Heartbeat: " & Now

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    ' SILENT LOGGING (Prevents UI Freeze)
    If Not wsP Is Nothing Then
        wsP.Range("Z1").value = "CRASH: " & Err.Description & " | " & Now
    End If
    Resume CleanExit
End Sub

' =========================================================================================
' LOGIC: SCANNING (Union of Master & Intraday)
' =========================================================================================
Sub Process_New_Signals(wsI As Worksheet, wsM As Worksheet, wsP As Worksheet)
    Dim processedSymbols As New Collection
    Dim ticker As String, r As Long
    Reset_VBA_Audit_Counters
    
    ' 1. Start with Python Brain Signals
    Dim lastM As Long: lastM = Get_True_Last_Row(wsM)
    For r = ROW_START_M To lastM
        ticker = Trim(wsM.cells(r, 1).value)
        If ticker <> "" Then
            If Not Is_Index(ticker) Then gVbaScannedRows = gVbaScannedRows + 1
            Dim sigCode As String: sigCode = UCase(Trim(wsM.cells(r, C_SIGNAL_CODE).value))
            Dim actionability As String: actionability = UCase(Trim(wsM.cells(r, C_ACTIONABILITY).value))
            If actionability = "ENTER_LONG" Or actionability = "ENTER_SHORT" Then gVbaEligibleRows = gVbaEligibleRows + 1
            If Not Is_Index(ticker) And sigCode <> "" And _
               (actionability = "ENTER_LONG" Or actionability = "ENTER_SHORT" Or _
                actionability = "PENDING_LONG" Or actionability = "PENDING_SHORT" Or _
                actionability = "ACTIONABLE" Or _
                actionability = "REVIEW_EXIT" Or actionability = "MANAGE_POSITION") Then
                Process_Single_Symbol wsM, wsP, ticker, processedSymbols
            End If
        End If
    Next r
    
    ' 2. Add Intraday Crosses
    Dim lastI As Long: lastI = Get_True_Last_Row(wsI)
    For r = ROW_START_I To lastI
        ticker = Trim(wsI.cells(r, 1).value)
        Process_Single_Symbol wsM, wsP, ticker, processedSymbols
    Next r
    
    ' 3. Refresh Active Trades (Enforce Sync)
    Dim lastP As Long: lastP = Get_True_Last_Row(wsP)
    For r = ROW_START_P To lastP
        Dim st As String: st = UCase(Trim(wsP.cells(r, P_STATUS).value))
        If st = "SIGNAL" Or st = "OPEN" Or st = "PENDING" Then
            ticker = Trim(wsP.cells(r, P_TICKER).value)
            Process_Single_Symbol wsM, wsP, ticker, processedSymbols
        End If
    Next r
    Write_VBA_Audit_Summary
End Sub

Sub Process_Single_Symbol(wsM As Worksheet, wsP As Worksheet, ticker As String, processedSymbols As Collection)
    If ticker = "" Or Is_Index(ticker) Then Exit Sub
    
    On Error Resume Next
    processedSymbols.Add ticker, ticker
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    On Error GoTo 0
    
    Dim pRow As Long: pRow = Find_Active_Trade_Row(wsP, ticker)
    
    If pRow > 0 Then
        Scan_Signal_From_Master wsM, wsP, ticker, pRow
    Else
        If Check_Cooldown(wsP, ticker) = False Then
            Scan_Signal_From_Master wsM, wsP, ticker
        End If
    End If
End Sub

Sub Scan_Signal_From_Master(wsM As Worksheet, wsP As Worksheet, ticker As String, Optional pRow As Long = 0)
    Dim mRow As Long: mRow = Find_Row(wsM, ticker)
    If mRow = 0 Then Exit Sub

    ' --- READ INPUTS ---
    Dim sigCode As String: sigCode = UCase(Trim(wsM.cells(mRow, C_SIGNAL_CODE).value))
    Dim actionability As String: actionability = UCase(Trim(wsM.cells(mRow, C_ACTIONABILITY).value))
    
    ' [ML-FIRST] Fallback: recover exact action from rich payload when F is coarse/stale.
    ' AI_Master F can still be ACTIONABLE/WATCH/IGNORE while J contains Action:ENTER_LONG etc.
    If actionability = "WATCH" Or actionability = "IGNORE" Or actionability = "" Or actionability = "ACTIONABLE" Then
        Dim reasonEarly As String: reasonEarly = UCase(Trim(wsM.cells(mRow, C_REASON).Text))
        Dim setupEarly As String: setupEarly = UCase(Trim(wsM.cells(mRow, C_SETUP).Text))
        If InStr(reasonEarly, "[ENTER_LONG]") > 0 Then actionability = "ENTER_LONG"
        If InStr(reasonEarly, "[ENTER_SHORT]") > 0 Then actionability = "ENTER_SHORT"
        If InStr(reasonEarly, "[PENDING_LONG]") > 0 Then actionability = "PENDING_LONG"
        If InStr(reasonEarly, "[PENDING_SHORT]") > 0 Then actionability = "PENDING_SHORT"
        If InStr(setupEarly, "ACTION:ENTER_LONG") > 0 Then actionability = "ENTER_LONG"
        If InStr(setupEarly, "ACTION:ENTER_SHORT") > 0 Then actionability = "ENTER_SHORT"
        If InStr(setupEarly, "ACTION:PENDING_LONG") > 0 Then actionability = "PENDING_LONG"
        If InStr(setupEarly, "ACTION:PENDING_SHORT") > 0 Then actionability = "PENDING_SHORT"
        If InStr(setupEarly, "ACTION:WATCHLIST") > 0 Then actionability = "WATCHLIST"
        If InStr(setupEarly, "ACTION:REVIEW_EXIT") > 0 Then actionability = "REVIEW_EXIT"
        If InStr(setupEarly, "ACTION:MANAGE_POSITION") > 0 Then actionability = "MANAGE_POSITION"
        If InStr(setupEarly, "ACTION:NO_ACTION") > 0 Then actionability = "NO_ACTION"
    End If
    
    Dim aiConf As Double: aiConf = val(wsM.cells(mRow, C_CONFIDENCE).value)
    ' [V5.1] aiConf is now Raw ATR Expected Value (e.g. 0.88 = 0.88x ATR move)
    
    Dim aiGrade As String: aiGrade = Trim(wsM.cells(mRow, C_GRADE).value)
    Dim setup As String: setup = wsM.cells(mRow, C_SETUP).value
    
    Dim entryPx As Double: entryPx = val(wsM.cells(mRow, C_ENTRY).value)
    Dim slPx As Double: slPx = val(wsM.cells(mRow, C_STOP).value)
    Dim tgtPx As Double: tgtPx = val(wsM.cells(mRow, C_TARGET).value)
    
    Dim ltp As Double: ltp = wsM.cells(mRow, C_LTP).value
    Dim rvol As Double: rvol = val(wsM.cells(mRow, C_RVOL).value)
    Dim atr As Double: atr = val(wsM.cells(mRow, C_ATR).value)
    If atr <= 0 Then atr = ltp * 0.005
    Dim sName As String: sName = wsM.cells(mRow, C_SECTOR).value

    If pRow = 0 And actionability <> "ENTER_LONG" And actionability <> "ENTER_SHORT" And _
       actionability <> "PENDING_LONG" And actionability <> "PENDING_SHORT" And _
       actionability <> "ACTIONABLE" And _
       actionability <> "REVIEW_EXIT" Then
        gVbaSkippedActionability = gVbaSkippedActionability + 1
        Audit_Paper_Decision ticker, "SKIP_ACTIONABILITY", actionability, sigCode, False, wsM.cells(mRow, C_REASON).Text
        Exit Sub
    End If
    
    ' --- FILTER OUT EMPTY NOISE SIGNALS (but allow ENTER rows to recover below) ---
    If entryPx = 0 And tgtPx = 0 And slPx = 0 And pRow = 0 And _
       actionability <> "ENTER_LONG" And actionability <> "ENTER_SHORT" Then
        gVbaSkippedZeroLevels = gVbaSkippedZeroLevels + 1
        Audit_Paper_Decision ticker, "SKIP_LEVELS", actionability, sigCode, False, wsM.cells(mRow, C_REASON).Text
        Exit Sub
    End If
    
    Dim mState As MacroState: mState = Get_Macro_Context(wsM, sName)
    
    Dim dir As String, isStealthFilled As Boolean
    isStealthFilled = (InStr(sigCode, "_FILLED") > 0)
    If actionability = "ENTER_LONG" Or actionability = "PENDING_LONG" Then
        dir = "BUY"
    ElseIf actionability = "ENTER_SHORT" Or actionability = "PENDING_SHORT" Then
        dir = "SELL"
    ElseIf InStr(sigCode, "BUY") > 0 Then
        dir = "BUY"
    ElseIf InStr(sigCode, "SELL") > 0 Then
        dir = "SELL"
    Else
        dir = "WAIT"
    End If

    ' Keep the paper-trade contract coherent even if AI_Master D is stale.
    If actionability = "ENTER_LONG" Then sigCode = "BUY_NOW"
    If actionability = "ENTER_SHORT" Then sigCode = "SELL_NOW"
    If actionability = "PENDING_LONG" And (sigCode = "" Or sigCode = "NO_TRADE" Or sigCode = "WAIT" Or sigCode = "STALE") Then sigCode = "BUY_PENDING"
    If actionability = "PENDING_SHORT" And (sigCode = "" Or sigCode = "NO_TRADE" Or sigCode = "WAIT" Or sigCode = "STALE") Then sigCode = "SELL_PENDING"

    ' [VBA RECOVERY] Some AI_Master rows arrive with ENTER_* action but stale D/zero AB:AD.
    ' Paper trade insertion needs levels, so synthesize a conservative ATR bracket from LTP.
    If entryPx = 0 And tgtPx = 0 And slPx = 0 And _
       (actionability = "ENTER_LONG" Or actionability = "ENTER_SHORT") And ltp > 0 Then
        entryPx = ltp
        If atr <= 0 Then atr = ltp * 0.005
        If dir = "SELL" Then
            slPx = entryPx + (atr * ATR_STOP_MULT)
            tgtPx = entryPx - (atr * ATR_STOP_MULT * 2)
        Else
            slPx = entryPx - (atr * ATR_STOP_MULT)
            tgtPx = entryPx + (atr * ATR_STOP_MULT * 2)
        End If
        Audit_Paper_Decision ticker, "RECOVER_LEVELS", actionability, sigCode, False, "VBA fallback from LTP/ATR because AB:AD were zero"
    End If

    If entryPx > 0 And tgtPx > 0 And slPx > 0 And _
       (actionability = "ENTER_LONG" Or actionability = "ENTER_SHORT" Or actionability = "PENDING_LONG" Or actionability = "PENDING_SHORT") Then
        Dim badGeometry As Boolean: badGeometry = False
        If dir = "BUY" Then badGeometry = Not (slPx < entryPx And entryPx < tgtPx)
        If dir = "SELL" Then badGeometry = Not (tgtPx < entryPx And entryPx < slPx)
        If badGeometry And ltp > 0 Then
            entryPx = ltp
            If atr <= 0 Then atr = ltp * 0.005
            If dir = "SELL" Then
                slPx = entryPx + (atr * ATR_STOP_MULT)
                tgtPx = entryPx - (atr * ATR_STOP_MULT * 2)
            Else
                slPx = entryPx - (atr * ATR_STOP_MULT)
                tgtPx = entryPx + (atr * ATR_STOP_MULT * 2)
            End If
            Audit_Paper_Decision ticker, "REPAIR_LEVEL_GEOMETRY", actionability, sigCode, False, "VBA repaired invalid entry/target/stop direction"
        End If
    End If
    
    ' Determine existing status early (for OPEN trades we must narrate stored levels).
    Dim existingStatus As String: existingStatus = ""
    If pRow > 0 Then existingStatus = UCase(Trim(wsP.cells(pRow, P_STATUS).value))

    ' If trade is OPEN, narrate using Paper_Trades levels (source-of-truth) because
    ' AI_Master may legitimately show zero levels in MANAGE/NO_ACTION states.
    If pRow > 0 And existingStatus = "OPEN" Then
        entryPx = val(wsP.cells(pRow, P_ENTRY).value)
        tgtPx = val(wsP.cells(pRow, P_TARGET).value)
        slPx = val(wsP.cells(pRow, P_SL).value)
    End If

    ' --- BUILD COMMENT (Trader-readable, stable schema) ---
    Dim stVal As Double: stVal = val(wsM.cells(mRow, C_SUPERTREND).value)
    Dim stPos As String: stPos = IIf(ltp > stVal And stVal > 0, "Above", "Below")

    Dim spikeVal As String: spikeVal = wsM.cells(mRow, C_SPIKE).Text
    If spikeVal = "" Then spikeVal = "0"

    Dim reasonStr As String: reasonStr = wsM.cells(mRow, C_REASON).Text
    
    ' --- EXTRACT CORE TAGS (Robust cross-column fallback) ---
    Dim pState As String: pState = Extract_Tag(reasonStr, "State:")
    If pState = "" Then pState = Extract_Tag(setup, "State:")
    
    Dim pAction As String: pAction = Extract_Tag(reasonStr, "Action:")
    If pAction = "" Then pAction = Extract_Tag(setup, "Action:")
    
    Dim pBucket As String: pBucket = Extract_Tag(reasonStr, "Bucket:")
    If pBucket = "" Then pBucket = Extract_Tag(setup, "Bucket:")
    
    Dim pConv As String: pConv = Extract_Tag(reasonStr, "Conv:")
    If pConv = "" Then pConv = Extract_Tag(setup, "Conv:")
    
    Dim pMin As String: pMin = Extract_Tag(reasonStr, "MinConf:")
    If pMin = "" Then pMin = Extract_Tag(setup, "MinConf:")
    
    ' Determine Direction (BUY/SELL)
    Dim pDir As String: pDir = "BUY"
    If actionability = "ENTER_SHORT" Or actionability = "PENDING_SHORT" Then
        pDir = "SELL"
    ElseIf actionability = "ENTER_LONG" Or actionability = "PENDING_LONG" Then
        pDir = "BUY"
    ElseIf InStr(1, actionability, "SHORT", vbTextCompare) > 0 Then
        pDir = "SELL"
    ElseIf InStr(1, actionability, "LONG", vbTextCompare) > 0 Then
        pDir = "BUY"
    ElseIf InStr(1, setup & reasonStr, "SHORT", vbTextCompare) > 0 Or InStr(1, setup & reasonStr, "SELL", vbTextCompare) > 0 Then
        pDir = "SELL"
    End If
    
    Dim pSector As String: pSector = Trim(wsM.cells(mRow, C_SECTOR).value) ' L

    ' [V5.2] Metrics extraction for Line 2
    Dim pRR As String: pRR = Extract_Tag(reasonStr, "RR:")
    If pRR = "" Then pRR = Extract_Tag(setup, "RR:")
    
    Dim pNetRR As String: pNetRR = Extract_Tag(reasonStr, "NetRR:")
    If pNetRR = "" Then pNetRR = Extract_Tag(setup, "NetRR:")
    
    Dim pRisk As String: pRisk = Extract_Tag(reasonStr, "Risk/Share:")
    If pRisk = "" Then pRisk = Extract_Tag(setup, "Risk/Share:")
    
    Dim pExec As String: pExec = Extract_Tag(reasonStr, "Exec:")
    If pExec = "" Then pExec = Extract_Tag(setup, "Exec:")

    Dim pPattern As String: pPattern = Extract_Tag(reasonStr, "Pattern:")
    If pPattern = "" Then pPattern = Extract_Tag(setup, "Pattern:")
    
    Dim pRiskLabel As String: pRiskLabel = Extract_Tag(reasonStr, "Risk:")
    If pRiskLabel = "" Then pRiskLabel = Extract_Tag(setup, "Risk:")
    
    Dim pHint As String: pHint = Extract_Tag(reasonStr, "Hint:")
    If pHint = "" Then pHint = Extract_Tag(setup, "Hint:")

    ' Compute RR and ATR geometry
    Dim rrVal As Double: rrVal = 0
    Dim denominator As Double: denominator = Abs(entryPx - slPx)
    If denominator > 0.0001 Then rrVal = Abs(tgtPx - entryPx) / denominator
    Dim stopAtr As Double: stopAtr = 0
    If atr > 0 And denominator > 0 Then stopAtr = denominator / atr

    Dim rrStr As String
    If pRR <> "" Then
        rrStr = " | RR=" & pRR
    Else
        rrStr = " | RR=" & Format(rrVal, "0.00")
    End If
    If pNetRR <> "" Then rrStr = rrStr & " | NetRR=" & pNetRR
    If pRisk <> "" Then rrStr = rrStr & " | Risk/Share=" & pRisk
    If pExec <> "" Then rrStr = rrStr & " | Exec=" & pExec

    ' Handle Diversification/Capping Status
    Dim mStatus As String: mStatus = ""
    If InStr(reasonStr, "[DIVERSIFICATION] Capped") > 0 Then mStatus = "CAP"
    If InStr(reasonStr, "[GATE_NOTE]") > 0 Then mStatus = "FILTERED"
    
    ' Extract Science triggers (everything after Checks:)
    Dim cfChecks As String: cfChecks = Extract_Tag(reasonStr, "Checks:")
    If cfChecks = "" Then cfChecks = Extract_Tag(setup, "Checks:")
    
    Dim cfScience As String: cfScience = ""
    ' Find the blob containing the Science tags
    Dim blob As String: blob = IIf(InStr(setup, "Checks:") > 0, setup, reasonStr)
    Dim checksPos As Long: checksPos = InStr(1, blob, "Checks:", vbTextCompare)
    If checksPos > 0 Then
        Dim pipePos As Long: pipePos = InStr(checksPos, blob, " | ")
        If pipePos > 0 Then
            Dim rawScience As String: rawScience = Trim(Mid(blob, pipePos + 3))
            ' Strip out the detail section (everything after ||)
            Dim dPos As Long: dPos = InStr(1, rawScience, "||")
            If dPos > 0 Then rawScience = Trim(Left(rawScience, dPos - 1))
            cfScience = rawScience
            cfScience = Replace(cfScience, " | ", ", ")
        End If
    End If

    If pPattern = "" Then pPattern = Derive_Pattern_Label(setup, reasonStr, cfScience, pState)
    If pRiskLabel = "" Then pRiskLabel = Derive_Risk_Label(reasonStr, setup, pExec, pNetRR, mStatus)
    If pHint = "" Then pHint = Derive_Corrective_Hint(actionability, pDir, pPattern, pRiskLabel)

    Dim actionHeadline As String: actionHeadline = Build_Action_Headline(actionability, pDir)

    ' 1. Derive Signal Book (Event/Trend/Reversal)
    Dim signalBook As String: signalBook = Derive_Signal_Book(reasonStr, setup)

    ' 2. Determine Cat1/Cat2/Cat3
    Dim catBucket As String
    If aiGrade = "A" Or aiGrade = "A+" Then
        catBucket = "Cat-1"
    ElseIf aiGrade = "B" Then
        catBucket = "Cat-2"
    Else
        catBucket = "Cat-3"
    End If

    ' 3. Grade
    Dim gradeStr As String
    If aiGrade <> "" Then
        gradeStr = "Grade " & aiGrade
    Else
        gradeStr = "Grade B"
    End If

    Dim line1 As String
    line1 = actionHeadline & " | " & signalBook & " | " & catBucket & " | " & gradeStr

    ' --- Line 2: Current State ---
    Dim stateStr As String
    stateStr = UCase(Trim(pState))
    If stateStr = "" Then
        If existingStatus = "OPEN" Then
            stateStr = "ACTIVE"
        ElseIf existingStatus = "SIGNAL" Or existingStatus = "PENDING" Then
            stateStr = "MATURING"
        Else
            stateStr = "ACTIVE"
        End If
    End If

    Dim sectorCode As String: sectorCode = Format_Sector_Code(pSector)

    Dim rvolStr As String
    rvolStr = "RVOL " & Format(rvol, "0.0") & "x"

    Dim execStyle As String: execStyle = Derive_Execution_Style(pPattern, stateStr)

    Dim line2 As String
    line2 = stateStr & " | " & sectorCode & " | " & rvolStr & " | " & execStyle

    ' --- Line 3: AI Assessment ---
    Dim secScore As Double: secScore = val(wsM.cells(mRow, C_SECTOR_SCORE).value)
    Dim line3 As String
    line3 = Derive_AI_Assessment_Tags(pDir, rvol, mState.niftyTrend, mState.sectorTrend, rrVal, execStyle, stateStr, spikeVal, cfChecks, secScore)

    ' Assemble cockpit text for Column K (P_COMM)
    Dim comm As String
    comm = line1 & vbNewLine & line2 & vbNewLine & line3
    
    ' --- EXECUTE OR SYNC ---------------------------------------------------------
If pRow > 0 Then

    '==========================================================================
    ' Synchronize order parameters while waiting for execution.
    ' Never overwrite an OPEN trade's frozen execution values.
    '==========================================================================
    If existingStatus <> "OPEN" Then

        If entryPx > 0 Then
            wsP.cells(pRow, P_ENTRY).value = entryPx
            wsP.cells(pRow, P_TARGET).value = tgtPx
            wsP.cells(pRow, P_SL).value = slPx
        End If

    End If


    '==========================================================================
    ' Always keep live market fields current.
    '==========================================================================
    wsP.cells(pRow, P_LTP).value = ltp
    wsP.cells(pRow, P_COMM).value = comm


    '==========================================================================
    ' IMPORTANT:
    '
    ' This routine NEVER transitions a trade to OPEN.
    '
    ' ENTER_LONG / ENTER_SHORT simply indicate that the order is eligible
    ' for execution.
    '
    ' Manage_Single_Trade() is the ONLY routine allowed to perform:
    '
    '     PENDING  --->  OPEN
    '
    ' after verifying that the latest candle actually traded through
    ' the Entry price.
    '==========================================================================
    If (actionability = "ENTER_LONG" _
        Or actionability = "ENTER_SHORT" _
        Or isStealthFilled) _
    And existingStatus = "SIGNAL" Then

        wsP.cells(pRow, P_STATUS).value = "PENDING"
        wsP.cells(pRow, P_ACTION).value = "PENDING " & dir

        existingStatus = "PENDING"

    End If


    Format_Single_Row wsP, pRow

End If


If dir = "WAIT" Or sigCode = "STALE" Then

    ' Conviction Locking in Manage_Single_Trade()
    ' decides whether to cancel.
    If pRow = 0 Then
        Audit_Paper_Decision ticker, _
                             "SKIP_WAIT_OR_STALE", _
                             actionability, _
                             sigCode, _
                             False, _
                             reasonStr
    End If

    Exit Sub

End If
    
    ' --- SKIP BLOCKED SIGNALS (no actionable levels) ---
    If entryPx = 0 And tgtPx = 0 And slPx = 0 Then
        If pRow = 0 Then gVbaSkippedZeroLevels = gVbaSkippedZeroLevels + 1
        If pRow = 0 Then Audit_Paper_Decision ticker, "SKIP_BLOCKED_LEVELS", actionability, sigCode, False, reasonStr
        Exit Sub
    End If
    
    If pRow = 0 And entryPx > 0 Then
        Dim newRow As Long: newRow = Find_First_Empty_Row(wsP)
        Dim newStatus As String: newStatus = "SIGNAL"
        Dim newAction As String: newAction = "SIGNAL " & dir
        Dim openNow As Boolean: openNow = False
        If actionability = "ENTER_LONG" Or actionability = "ENTER_SHORT" Then
            newStatus = "PENDING"
            newAction = "PENDING " & dir
            openNow = False
        End If
        Update_Row wsP, newRow, ticker, dir, entryPx, tgtPx, slPx, newStatus, newAction, comm, ltp, True, openNow
        gVbaInsertedRows = gVbaInsertedRows + 1
        Audit_Paper_Decision ticker, IIf(openNow, "INSERT_OPEN", "INSERT_SIGNAL"), actionability, sigCode, True, reasonStr
    ElseIf pRow > 0 Then
        Audit_Paper_Decision ticker, "SYNC_ACTIVE", actionability, sigCode, False, reasonStr
    End If
End Sub

Sub Update_Existing_Records(wsP As Worksheet, wsM As Worksheet)
    Dim lastP As Long: lastP = Get_True_Last_Row(wsP)
    If lastP < ROW_START_P Then Exit Sub
    Dim r As Long
    For r = ROW_START_P To lastP
        Dim st As String: st = UCase(Trim(wsP.cells(r, P_STATUS).value))
        If st = "SIGNAL" Or st = "OPEN" Or st = "PENDING" Then
             Dim ticker As String: ticker = Trim(wsP.cells(r, P_TICKER).value)
             Dim mRow As Long: mRow = Find_Row(wsM, ticker)
             If mRow > 0 Then
                  ' Call Scan_Signal_From_Master to rebuild comments + update status
                  Scan_Signal_From_Master wsM, wsP, ticker, r
             End If
        End If
    Next r
End Sub

' [FIX] Manage state machine for active trades: SIGNAL ? OPEN ? CLOSED
Sub Manage_Open_Trades(wsP As Worksheet, wsM As Worksheet)
    Dim lastP As Long: lastP = Get_True_Last_Row(wsP)
    If lastP < ROW_START_P Then Exit Sub
    Dim r As Long
    For r = ROW_START_P To lastP
        Dim st As String: st = UCase(Trim(wsP.cells(r, P_STATUS).value))
        If st = "SIGNAL" Or st = "OPEN" Or st = "PENDING" Then
            Dim ticker As String: ticker = Trim(wsP.cells(r, P_TICKER).value)
            Dim ltp As Double: ltp = wsP.cells(r, P_LTP).value
            If ltp <= 0 Then
                Dim mRow As Long: mRow = Find_Row(wsM, ticker)
                If mRow > 0 Then ltp = wsM.cells(mRow, C_LTP).value
            End If
            Dim atr As Double: atr = 0
            Dim mRow2 As Long: mRow2 = Find_Row(wsM, ticker)
            If mRow2 > 0 Then atr = val(wsM.cells(mRow2, C_ATR).value)
            If atr <= 0 And ltp > 0 Then atr = ltp * 0.005
            Dim pySig As String: pySig = ""
            Dim aiConf As Double: aiConf = 0
            Dim aiGrade As String: aiGrade = ""
            If mRow2 > 0 Then
                pySig = UCase(Trim(wsM.cells(mRow2, C_SIGNAL_CODE).value))
                aiConf = val(wsM.cells(mRow2, C_CONFIDENCE).value)
                aiGrade = Trim(wsM.cells(mRow2, C_GRADE).value)
            End If
            Manage_Single_Trade wsP, r, ltp, atr, pySig, aiConf, aiGrade
        End If
    Next r
End Sub

Sub Manage_Single_Trade(wsP As Worksheet, r As Long, ltp As Double, atr As Double, Optional pySig As String = "", Optional aiConf As Double = 0, Optional aiGrade As String = "")
    Dim status As String: status = UCase(Trim(wsP.cells(r, P_STATUS).value))
    Dim act As String: act = UCase(wsP.cells(r, P_ACTION).value)
    Dim dir As String: dir = IIf(InStr(act, "BUY") > 0, "BUY", "SELL")
    Dim sh1 As Worksheet
    On Error Resume Next
    Set sh1 = ThisWorkbook.Sheets(SH_RAW)
    On Error GoTo 0
    wsP.cells(r, P_LTP).value = ltp
    
    If (status = "SIGNAL" Or status = "PENDING") And (pySig = "STALE" Or pySig = "WAIT" Or pySig = "EXIT_WARN") Then
        ' --- CONVICTION LOCKING ---
        ' Hold the status PENDING/SIGNAL for at least 3 cycles (15 mins) ignoring single-bar WAIT drops
        Dim minElapsed As Double
        minElapsed = (Now - Date - val(wsP.cells(r, P_SIG_TIME).value)) * 1440 ' Convert days to minutes
        If minElapsed >= 15 Then
            wsP.cells(r, P_STATUS).value = "CANCELLED"
            wsP.cells(r, P_CLOSE_TIME).value = Format(Now, "hh:mm")
            Exit Sub
        End If
        ' Allow it to fall through to check for Entry Fill below despite being in WAIT state!
    End If
    
    If status = "SIGNAL" Or status = "PENDING" Then
        ' [OHLCV VALIDATION] Check if entry price was actually touched in the 5-min candle.
        ' Read OHLCV from Sheet1 (BV=High, BW=Low, BX=Close, T=LTP).
        Dim tickerSig As String: tickerSig = Trim(wsP.cells(r, P_TICKER).value)
        Dim sigEntry As Double: sigEntry = val(wsP.cells(r, P_ENTRY).value)
        Dim sigATR As Double: sigATR = atr
        If sigATR <= 0 Then sigATR = sigEntry * 0.01  ' fallback 1%
        
        If sh1 Is Nothing Then Exit Sub
        Dim sigRow As Variant: sigRow = Application.Match(tickerSig, sh1.Range("A:A"), 0)
        
        If Not IsError(sigRow) Then
            Dim barHigh As Double: barHigh = val(sh1.cells(sigRow, "DR").value)
            Dim barLow As Double: barLow = val(sh1.cells(sigRow, "DS").value)
            Dim barClose As Double: barClose = val(sh1.cells(sigRow, "DT").value)
            
            ' Check if entry price is within current candle range. If touched, fill.
            ' If not touched, keep PENDING — decay timer cancels stale signals later.
            If barHigh > 0 And barLow > 0 Then
                If sigEntry >= barLow And sigEntry <= barHigh Then

                    If (dir = "BUY" And ltp >= sigEntry) Or _
                       (dir = "SELL" And sigEntry >= ltp) Then
                
                        wsP.cells(r, P_ENTRY).value = sigEntry
                        wsP.cells(r, P_STATUS).value = "OPEN"
                        wsP.cells(r, P_ACTION).value = "HOLD " & dir
                
                        If wsP.cells(r, P_OPEN_TIME).value = "" Then
                            wsP.cells(r, P_OPEN_TIME).value = (Now - Date)
                        End If
                
                    End If
                
                End If
                                ' Otherwise: keep PENDING — price hasn't been touched yet.
                ' Decay timer at line 785 cancels after 25 min if still unfilled.
            End If
        End If
        
        ' Signal expiry: if still SIGNAL/PENDING after 5 cycles (25 min), cancel as stale
        Dim sigAge As Double
        sigAge = (Now - Date - val(wsP.cells(r, P_SIG_TIME).value)) * 1440
        If sigAge >= 25 And (status = "SIGNAL" Or status = "PENDING") Then
            wsP.cells(r, P_STATUS).value = "CANCELLED"
            wsP.cells(r, P_ACTION).value = "CANCEL: signal expired (" & Format(sigAge, "0") & " min)"
            wsP.cells(r, P_CLOSE_TIME).value = Format(Now, "hh:mm")
            Exit Sub
        End If
    End If
    
    ' Re-read status (may have changed above)
    status = UCase(Trim(wsP.cells(r, P_STATUS).value))
    
    If status = "OPEN" Then
        If pySig = "EXIT_NOW" Then
            wsP.cells(r, P_STATUS).value = "CLOSED"
            wsP.cells(r, P_ACTION).value = "EXIT: SNIPER"
            wsP.cells(r, P_CLOSE_TIME).value = Format(Now, "hh:mm")
            Exit Sub
        End If
        If pySig = "EXIT_WARN" Then
            wsP.cells(r, P_STATUS).value = "CLOSED"
            wsP.cells(r, P_ACTION).value = "EXIT: AI WITHDRAWN"
            wsP.cells(r, P_CLOSE_TIME).value = Format(Now, "hh:mm")
            Exit Sub
        End If
        
        Dim sl As Double: sl = val(wsP.cells(r, P_SL).value)
        Dim tp As Double: tp = val(wsP.cells(r, P_TARGET).value)
        Dim en As Double: en = val(wsP.cells(r, P_ENTRY).value)
        
        ' [OHLCV ENHANCED] Read bar candle for SL/Target validation
        Dim tickerOpen As String: tickerOpen = Trim(wsP.cells(r, P_TICKER).value)
        If sh1 Is Nothing Then GoTo CoachOnly
        Dim ohRow As Variant: ohRow = Application.Match(tickerOpen, sh1.Range("A:A"), 0)
        Dim barH As Double: barH = 0: Dim barL As Double: barL = 0
        If Not IsError(ohRow) Then
            barH = val(sh1.cells(ohRow, "DR").value)
            barL = val(sh1.cells(ohRow, "DS").value)
        End If
        
        ' --- PARTIAL R:R CALCULATION (1.0 R) ---
        Dim partialTp As Double
        Dim risk As Double: risk = Abs(en - sl)
        If dir = "BUY" Then partialTp = en + risk Else partialTp = en - risk
        
        ' 1. Check Stop Loss (ltp OR bar Low/High touched SL)
        Dim slHit As Boolean: slHit = False
        If dir = "BUY" Then
            slHit = (ltp <= sl) Or (barL > 0 And barL <= sl)
        Else
            slHit = (ltp >= sl) Or (barH > 0 And barH >= sl)
        End If
        If slHit Then
            wsP.cells(r, P_STATUS).value = "CLOSED"
            wsP.cells(r, P_ACTION).value = "EXIT LOSS"
            wsP.cells(r, P_CLOSE_TIME).value = Format(Now, "hh:mm")
        
        ' 2. Check Final Target (ltp OR bar Low/High touched target)
        ElseIf (dir = "BUY" And (ltp >= tp Or (barH > 0 And barH >= tp))) Or _
               (dir = "SELL" And (ltp <= tp Or (barL > 0 And barL <= tp))) Then
            wsP.cells(r, P_STATUS).value = "CLOSED"
            wsP.cells(r, P_ACTION).value = "EXIT WIN"
            wsP.cells(r, P_CLOSE_TIME).value = Format(Now, "hh:mm")
            
        ' 3. Check Partial Target (1R)
        ElseIf InStr(1, act, "EXIT 50%", vbTextCompare) = 0 Then
            If (dir = "BUY" And ltp >= partialTp) Or (dir = "SELL" And ltp <= partialTp) Then
                wsP.cells(r, P_ACTION).value = "EXIT 50%: 1R Reached"
            End If
        End If
    End If
    
    ' Refresh status if it was closed
CoachOnly:
    status = UCase(Trim(wsP.cells(r, P_STATUS).value))
    Update_Trade_Coach wsP, r, ltp, dir, status, aiConf, aiGrade
End Sub

' =========================================================================================
' HELPER FUNCTIONS
' =========================================================================================
Function Get_Macro_Context(wsM As Worksheet, sName As String) As MacroState
    Dim res As MacroState: res.niftyTrend = "NEUTRAL": res.sectorTrend = "NEUTRAL"
    Dim nRow As Long: nRow = Find_Row(wsM, "NSEI")
    If nRow > 0 Then
        Dim nLTP As Double: nLTP = wsM.cells(nRow, C_LTP).value
        Dim nSt As Double: nSt = wsM.cells(nRow, C_SUPERTREND).value
        res.niftyTrend = IIf(nLTP > nSt, "BULLISH", "BEARISH")
    End If
    Dim sRow As Long: sRow = Find_Row(wsM, sName)
    If sRow > 0 Then
        Dim sLtp As Double: sLtp = wsM.cells(sRow, C_LTP).value
        Dim sSt As Double: sSt = wsM.cells(sRow, C_SUPERTREND).value
        res.sectorTrend = IIf(sLtp > sSt, "BULLISH", "BEARISH")
    End If
    Get_Macro_Context = res
End Function

Function Check_Cooldown(wsP As Worksheet, ticker As String) As Boolean
    Dim rng As Range: Set rng = wsP.Columns(P_TICKER).Find(ticker, LookIn:=xlValues, SearchDirection:=xlPrevious)
    If Not rng Is Nothing Then
        If rng.row >= ROW_START_P Then
            Dim st As String: st = UCase(wsP.cells(rng.row, P_STATUS).value)
            If st = "SIGNAL" Or st = "OPEN" Or st = "PENDING" Then Check_Cooldown = True: Exit Function
        End If
    End If
    Check_Cooldown = False
End Function

Function Get_True_Last_Row(ws As Worksheet) As Long
    Dim rng As Range: Set rng = ws.Columns(1).Find("*", SearchDirection:=xlPrevious)
    If rng Is Nothing Then Get_True_Last_Row = 1 Else Get_True_Last_Row = rng.row
End Function

Function Find_First_Empty_Row(ws As Worksheet) As Long
    Find_First_Empty_Row = Get_True_Last_Row(ws) + 1
    If Find_First_Empty_Row < ROW_START_P Then Find_First_Empty_Row = ROW_START_P
End Function

Function Find_Active_Trade_Row(ws As Worksheet, ticker As String) As Long
    Dim rng As Range: Set rng = ws.Columns(P_TICKER).Find(ticker, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
    If Not rng Is Nothing Then
        If rng.row >= ROW_START_P Then
            Dim st As String: st = UCase(ws.cells(rng.row, P_STATUS).value)
            If st = "SIGNAL" Or st = "OPEN" Or st = "PENDING" Then Find_Active_Trade_Row = rng.row: Exit Function
        End If
    End If
    Find_Active_Trade_Row = 0
End Function

Function Find_Row(ws As Worksheet, searchVal As String) As Long
    Dim res As Variant: res = Application.Match(searchVal, ws.Columns(1), 0)
    If IsError(res) Then res = Application.Match(searchVal & ".NS", ws.Columns(1), 0)
    If IsError(res) Then Find_Row = 0 Else Find_Row = CLng(res)
End Function

Function Is_Index(t As String) As Boolean
    Is_Index = (InStr(t, "NIFTY") > 0 Or InStr(t, "CNX") > 0 Or InStr(t, "NSEBANK") > 0 Or InStr(t, "MIDCAP") > 0 Or InStr(t, "SMALLCAP") > 0 Or InStr(t, "NSEI") > 0)
End Function

Sub Update_Row(ws As Worksheet, r As Long, t As String, d As String, e As Double, tg As Double, sl As Double, st As String, act As String, cm As String, l As Double, Optional resetSignalTime As Boolean = False, Optional openNow As Boolean = False)
    With ws
        .cells(r, P_TICKER).value = t: .cells(r, P_DIR).value = d
        If resetSignalTime Or .cells(r, P_DATE).value = "" Then .cells(r, P_DATE).value = Date
        If resetSignalTime Or .cells(r, P_SIG_TIME).value = "" Then .cells(r, P_SIG_TIME).value = (Now - Date)
        .cells(r, P_ENTRY).value = e: .cells(r, P_TARGET).value = tg: .cells(r, P_SL).value = sl
        .cells(r, P_STATUS).value = st: .cells(r, P_ACTION).value = act: .cells(r, P_COMM).value = cm: .cells(r, P_LTP).value = l
        
        ' Clear residual data in case this is a recycled row
        .cells(r, P_OPEN_TIME).value = ""
        .cells(r, P_CLOSE_TIME).value = ""
        If openNow Then .cells(r, P_OPEN_TIME).value = (Now - Date)
    End With
    Format_Single_Row ws, r
End Sub

Function Calc_Pct_Change(ByVal CurrentPrice As Double, ByVal PointsChange As Double) As Double
    If CurrentPrice = 0 Then Exit Function
    Dim Prev As Double: Prev = CurrentPrice - PointsChange
    If Prev = 0 Then Exit Function
    Calc_Pct_Change = (PointsChange / Prev) * 100
End Function

' =========================================================================================
' PRESERVATION ZONES (K3 Dashboard & Population)
' =========================================================================================
Sub Update_K3_Dashboard(wsM As Worksheet, wsP As Worksheet)
    On Error Resume Next
    Dim k3 As Range: Set k3 = wsP.Range("K3")
    Dim nRow As Long: nRow = Find_Row(wsM, "NSEI")
    If nRow = 0 Then nRow = Find_Row(wsM, "NIFTY 50")
    If nRow = 0 Then k3.value = "NIFTY NOT FOUND": Exit Sub
    
    Dim ltp As Double: ltp = wsM.cells(nRow, RAW_LTP).value
    Dim chg As Double: chg = wsM.cells(nRow, RAW_CHANGE_PTS).value
    Dim st As Double: st = wsM.cells(nRow, RAW_SUPERTREND).value
    Dim ema20 As Double: ema20 = val(wsM.cells(nRow, RAW_EMA20_5M).value)
    
    ' --- LIVE ADX & INDICATORS ---
    Dim adxVal As Double: adxVal = val(wsM.cells(nRow, "CB").value)
    Dim atrVal As Double: atrVal = val(wsM.cells(nRow, RAW_ATR).value)
    
    ' --- LEVELS ---
    Dim piv As Double: piv = wsM.cells(nRow, RAW_PIVOT).value
    Dim r1 As Double: r1 = wsM.cells(nRow, RAW_R1).value
    Dim s1 As Double: s1 = wsM.cells(nRow, RAW_S1).value
    Dim s2 As Double: s2 = wsM.cells(nRow, RAW_S2).value
    Dim r2 As Double: r2 = wsM.cells(nRow, RAW_R2).value
    
    ' --- PYTHON CONTEXT READ ---
    Dim nScore As String: nScore = wsM.cells(nRow, "BV").Text
    Dim nRegime As String: nRegime = wsM.cells(nRow, "BW").Text
    Dim nSentVal As Double: nSentVal = val(wsM.cells(nRow, "DN").value)
    
    If ltp = 0 Then k3.value = "NIFTY DATA LOADING...": Exit Sub
    Dim pct As Double: pct = Calc_Pct_Change(ltp, chg)
    
    ' --- ENHANCED BEHAVIORAL LOGIC ---
    Dim trendDesc As String, behavior As String, whatNext As String, trendColor As Long
    Dim emaDist As Double: emaDist = 0: If ema20 > 0 Then emaDist = ((ltp - ema20) / ema20) * 100
    
    Dim drift As String: drift = IIf(chg >= 0, "Drifting Up", "Drifting Down")
    Dim stretchState As String: stretchState = IIf(Abs(emaDist) > 0.4, "Stretched", "Stable")
    Dim atrStretch As String: atrStretch = ""
    If atrVal > 0 Then atrStretch = " [" & Format(Abs(ltp - ema20) / atrVal, "0.0") & "x ATR]"
    
    behavior = drift & ". Currently " & Format(Abs(emaDist), "0.00") & "% " & _
               IIf(emaDist >= 0, "above", "below") & " 20-EMA (" & stretchState & ")" & atrStretch & "."
    
    If ltp > st Then
        trendDesc = IIf(ltp > ema20, "STRONG BULLISH", "WEAK BULLISH")
        trendColor = RGB(0, 128, 0)
        If stretchState = "Stretched" Then
            whatNext = "Expect 'Mean Reversion' Pullback to " & Format(ema20, "0")
        Else
            whatNext = "Follow Momentum. Next major hurdle: " & Format(r1, "0")
        End If
    Else
        trendDesc = IIf(ltp < ema20, "STRONG BEARISH", "WEAK BEARISH")
        trendColor = RGB(180, 0, 0)
        If stretchState = "Stretched" Then
            whatNext = "Expect 'Dead-Cat' Bounce to " & Format(ema20, "0")
        Else
            whatNext = "Follow Downside. Next major support: " & Format(s1, "0")
        End If
    End If
    
    ' --- LEVEL CLUSTERING & TACTICAL SPACE ---
    Dim nearSup As Double, nearRes As Double, supName As String, resName As String
    If ltp > piv Then nearSup = piv: supName = "Pivot" Else nearSup = s1: supName = "S1"
    If ltp < s1 Then nearSup = s2: supName = "S2"
    If ltp < piv Then nearRes = piv: resName = "Pivot" Else nearRes = r1: resName = "R1"
    If ltp > r1 Then nearRes = r2: resName = "R2"
    
    Dim room As String: room = "Room to " & resName & ": +" & Format(nearRes - ltp, "0") & " pts [" & Format(nearRes, "0") & "] | Risk to " & supName & ": -" & Format(ltp - nearSup, "0") & " pts [" & Format(nearSup, "0") & "]"
    
    Dim clusterMsg As String: clusterMsg = ""
    If Abs(piv - s1) < 15 Then clusterMsg = " [Support Cluster: Pivot+S1]"
    If Abs(piv - r1) < 15 Then clusterMsg = " [Resistance Cluster: Pivot+R1]"
    
    ' --- CONVICTION & SENTIMENT GUARD ---
    Dim adxLabel As String
    If adxVal < 20 Then
        adxLabel = " (Grind/Range)"
    ElseIf adxVal < 30 Then
        adxLabel = " (Trending)"
    Else
        adxLabel = " (Impulsive!)"
    End If
    
    Dim sentiWarning As String: sentiWarning = ""
    If (chg > 0 And nSentVal < 4#) Or (chg < 0 And nSentVal > 6#) Then sentiWarning = " " & ChrW(9888) & " SENTI DIVERGENCE"
    
    ' --- SESSION PHASE ---
    Dim sessLabel As String, h As Integer, m As Integer, ct As Double
    h = Hour(Now): m = Minute(Now): ct = h + (m / 60)
    If ct < 10 Then
        sessLabel = "OPENING DRIVE"
    ElseIf ct < 13 Then
        sessLabel = "MID-DAY GRIND"
    ElseIf ct < 14.75 Then
        sessLabel = "AFTERNOON PUSH"
    Else
        sessLabel = "CLOSING PRESSURE"
    End If

    ' --- BUILD TACTICAL COCKPIT ---
    Dim dash As String
    dash = "NIFTY 50: " & Format(ltp, "#,##0.00") & " (" & IIf(chg > 0, "+", "") & Format(pct, "0.00") & "%) | ADX: " & Format(adxVal, "0") & adxLabel & vbNewLine & _
           "STATUS: " & trendDesc & " (AI: " & nScore & ") | REGIME: " & nRegime & " | SENTI: " & Format(nSentVal, "0.0") & sentiWarning & vbNewLine & _
           "SESSION: " & sessLabel & " | BEHAVIOR: " & behavior & vbNewLine & _
           "TACTICAL: " & room & clusterMsg

    ' --- [NEW] APPEND INSTITUTIONAL SUMMARY & PARSE SIZE ---
    Dim instSum As String: instSum = wsM.cells(nRow, RAW_INST_SUMMARY).Text
    
    ' [V7.1] Parse Capital Multiplier for HUD display
    Dim sizeVal As String: sizeVal = "1.00"
    If InStr(instSum, "SIZE=") > 0 Then
        On Error Resume Next
        sizeVal = Split(instSum, "SIZE=")(1)
        sizeVal = Trim(Replace(sizeVal, "|", ""))
        On Error GoTo 0
    End If

    instSum = Replace(instSum, "Institutional Summary for this Scan:", "")
    ' Normalize all line breaks
    instSum = Replace(instSum, vbCrLf, vbLf)
    instSum = Replace(instSum, vbCr, vbLf)
    
    ' Collapse ANY double blank lines
    Do While InStr(instSum, vbLf & vbLf) > 0
        instSum = Replace(instSum, vbLf & vbLf, vbLf)
    Loop
    
    ' Convert back
    instSum = Replace(instSum, vbLf, vbCrLf)
    
    instSum = Replace(instSum, "Verdict", "VERDICT")
    instSum = Replace(instSum, "Action", "ACTION")
    instSum = Trim(instSum)
    
    If instSum <> "" Then
        ' Prepend Size to the summary block for clarity
        dash = dash & vbNewLine & "CAPITAL SIZE: " & sizeVal & "x" & vbNewLine & instSum
    End If
    With k3
        .value = dash
        .WrapText = True
        .Font.Bold = False
        .Font.Color = vbBlack
        ' Apply bolding and colors
        Highlight_Text k3, "STATUS:", RGB(0, 0, 0), True
        Highlight_Text k3, "REGIME:", RGB(0, 0, 0), True
        Highlight_Text k3, "TACTICAL:", RGB(0, 0, 0), True
        Highlight_Text k3, "SENTI:", RGB(0, 0, 0), True
        Highlight_Text k3, "BEHAVIOR:", RGB(0, 0, 0), True
        Highlight_Text k3, "LEVELS:", RGB(0, 0, 0), True
        Highlight_Text k3, "SESSION:", RGB(0, 0, 0), True
        Highlight_Text k3, "Verdict:", RGB(0, 0, 128), True
        Highlight_Text k3, "Action:", RGB(160, 0, 0), True
        Highlight_Text k3, "Institutional Summary for this Scan:", RGB(0, 0, 0), True
        Highlight_Text k3, trendDesc, trendColor, True
        If Len(sentiWarning) > 0 Then Highlight_Text k3, sentiWarning, RGB(180, 0, 0), True
        If Abs(ltp - nearSup) < 12 Then Highlight_Text k3, "LEVELS:", RGB(0, 128, 0), True
    End With
    On Error GoTo 0
End Sub

Sub Highlight_Text(rng As Range, txt As String, clr As Long, Optional bld As Boolean = False)
    Dim p As Long: p = InStr(1, rng.value, txt, vbTextCompare)
    If p > 0 Then
        rng.Characters(p, Len(txt)).Font.Color = clr
        If bld Then rng.Characters(p, Len(txt)).Font.Bold = True
    End If
End Sub

Private Sub Highlight_Cell_Substring(cell As Range, txt As String, clr As Long, bld As Boolean, startIdx As Long)
    If Len(txt) = 0 Then Exit Sub
    Dim p As Long: p = startIdx
    Do
        p = InStr(p, cell.value, txt, vbTextCompare)
        If p = 0 Then Exit Do
        With cell.Characters(p, Len(txt)).Font
            .Color = clr
            .Bold = bld
        End With
        p = p + Len(txt)
    Loop
End Sub


Sub Format_Single_Row(ws As Worksheet, r As Long)
    On Error Resume Next
    With ws.cells(r, P_ACTION)
        .Font.Bold = True
        Dim actTxt As String: actTxt = UCase(.value)
        If InStr(actTxt, "HOLD") > 0 Or InStr(actTxt, "PENDING") > 0 Or InStr(actTxt, "OPEN") > 0 Then
            .Font.Color = RGB(0, 0, 255)
        ElseIf InStr(actTxt, "BUY") > 0 Or InStr(actTxt, "WIN") > 0 Then
            .Font.Color = RGB(0, 128, 0)
        ElseIf InStr(actTxt, "SELL") > 0 Or InStr(actTxt, "LOSS") > 0 Then
            .Font.Color = RGB(192, 0, 0)
        ElseIf InStr(actTxt, "WATCHLIST") > 0 Then
            .Font.Color = RGB(105, 105, 105)
        Else
            .Font.Color = RGB(128, 128, 128)
        End If
    End With
    
    Dim c As Range: Set c = ws.cells(r, P_COMM)
    c.Font.Color = vbBlack: c.Font.Bold = False
    
    ' --- Line 1 Action Headlines & Categories ---
    Highlight_Text c, "BUY NOW", RGB(0, 128, 0), True
    Highlight_Text c, "SELL NOW", RGB(192, 0, 0), True
    Highlight_Text c, "EXIT NOW", RGB(192, 0, 0), True
    Highlight_Text c, "SETUP BUY", RGB(0, 102, 204), True
    Highlight_Text c, "SETUP SELL", RGB(0, 102, 204), True
    Highlight_Text c, "HOLD", RGB(0, 70, 200), True
    
    ' Signal Books (Key words - bold black)
    Highlight_Text c, "Event", RGB(0, 0, 0), True
    Highlight_Text c, "Trend", RGB(0, 0, 0), True
    Highlight_Text c, "Reversal", RGB(0, 0, 0), True
    
    ' Category Buckets
    Highlight_Text c, "[Cat-1]", RGB(0, 128, 0), True
    Highlight_Text c, "[Cat-2]", RGB(0, 0, 200), True
    Highlight_Text c, "[Cat-3]", RGB(120, 120, 120), False
    
    ' Grades
    Highlight_Text c, "Grade A", RGB(0, 128, 0), True
    Highlight_Text c, "Grade B", RGB(0, 0, 200), True
    Highlight_Text c, "Grade C", RGB(120, 120, 120), False
    Highlight_Text c, "Grade D", RGB(192, 0, 0), True
    Highlight_Text c, "Grade None", RGB(120, 120, 120), False
    
    ' --- Line 2 Lifecycle states ---
    Highlight_Text c, "ACTIVE", RGB(0, 128, 0), True
    Highlight_Text c, "MATURING", RGB(0, 70, 200), True
    Highlight_Text c, "WEAKENING", RGB(200, 110, 0), True
    Highlight_Text c, "BROKEN", RGB(192, 0, 0), True
    
    ' Execution Styles
    Highlight_Text c, "Breakout", RGB(0, 128, 0), True
    Highlight_Text c, "Pullback", RGB(0, 102, 204), True
    Highlight_Text c, "Reversal", RGB(200, 110, 0), True
    Highlight_Text c, "Breakdown", RGB(192, 0, 0), True
    Highlight_Text c, "Fading", RGB(200, 110, 0), True
    Highlight_Text c, "Failed", RGB(192, 0, 0), True
    
    ' --- Line 3 AI Assessment tags (Strengths & Cautions) ---
    Highlight_Text c, "Strong Sector", RGB(0, 128, 0), True
    Highlight_Text c, "Sector Leader", RGB(0, 128, 0), True
    Highlight_Text c, "Strong Flow", RGB(0, 128, 0), True
    Highlight_Text c, "High RVOL", RGB(0, 128, 0), True
    Highlight_Text c, "Clean Trend", RGB(0, 128, 0), True
    Highlight_Text c, "Good RR", RGB(0, 128, 0), True
    Highlight_Text c, "Momentum", RGB(0, 128, 0), True
    
    Highlight_Text c, "Weak Sector", RGB(192, 0, 0), True
    Highlight_Text c, "Counter Trend", RGB(200, 110, 0), True
    Highlight_Text c, "Low RVOL", RGB(192, 0, 0), True
    Highlight_Text c, "Extended", RGB(200, 110, 0), True
    Highlight_Text c, "Low RR", RGB(192, 0, 0), True
    Highlight_Text c, "Spike", RGB(192, 0, 0), True
    Highlight_Text c, "Don't Add", RGB(192, 0, 0), True
    Highlight_Text c, "Weakening", RGB(200, 110, 0), True
    
    ' Dynamic CNX sector code highlighting (Bold Dark Black)
    Dim pCnx As Long: pCnx = 1
    Do
        pCnx = InStr(pCnx, c.value, "CNX", vbTextCompare)
        If pCnx = 0 Then Exit Do
        Dim pSpace As Long: pSpace = InStr(pCnx, c.value, " ")
        Dim pPipe As Long: pPipe = InStr(pCnx, c.value, "|")
        Dim pNL As Long: pNL = InStr(pCnx, c.value, vbCr)
        Dim pLF As Long: pLF = InStr(pCnx, c.value, vbLf)
        
        Dim pEndWord As Long: pEndWord = Len(c.value) + 1
        If pSpace > 0 And pSpace < pEndWord Then pEndWord = pSpace
        If pPipe > 0 And pPipe < pEndWord Then pEndWord = pPipe
        If pNL > 0 And pNL < pEndWord Then pEndWord = pNL
        If pLF > 0 And pLF < pEndWord Then pEndWord = pLF
        
        Dim cnxWord As String: cnxWord = Trim(Mid(c.value, pCnx, pEndWord - pCnx))
        If Len(cnxWord) > 3 Then
            With c.Characters(pCnx, Len(cnxWord)).Font
                .Bold = True
                .Color = RGB(0, 0, 0)
            End With
        End If
        pCnx = pCnx + 3
    Loop
    
    ' Dynamic RVOL formatting
    Dim pRvol As Long: pRvol = InStr(1, c.value, "RVOL ", vbTextCompare)
    If pRvol > 0 Then
        c.Characters(pRvol, 4).Font.Bold = True
        c.Characters(pRvol, 4).Font.Color = RGB(0, 0, 0)
        
        Dim pEndRvol As Long: pEndRvol = InStr(pRvol + 5, c.value, " |", vbTextCompare)
        If pEndRvol = 0 Then pEndRvol = Len(c.value) + 1
        Dim rvolValStr As String: rvolValStr = Trim(Mid(c.value, pRvol + 5, pEndRvol - pRvol - 5))
        Dim rvolNumStr As String: rvolNumStr = Replace(rvolValStr, "x", "")
        If IsNumeric(rvolNumStr) Then
            Dim rvolVal As Double: rvolVal = val(rvolNumStr)
            Dim rvolClr As Long: rvolClr = RGB(0, 0, 0)
            If rvolVal >= 2# Then
                rvolClr = RGB(0, 128, 0) ' Green
            ElseIf rvolVal < 1# Then
                rvolClr = RGB(200, 110, 0) ' Orange
            End If
            With c.Characters(pRvol + 5, Len(rvolValStr)).Font
                .Bold = True
                .Color = rvolClr
            End With
        End If
    End If
    
    ' Backwards compatibility and helper formats
    Highlight_Text c, "[HOT]", RGB(255, 69, 0), True
    Highlight_Text c, "[STATUS: CAP]", RGB(220, 20, 60), True
    Highlight_Text c, "[STATUS: FILTERED]", RGB(220, 20, 60), True
    Highlight_Text c, "[CONV:", RGB(128, 0, 128), True
    Highlight_Text c, ">> LOCKED", RGB(160, 0, 0), True
    
    On Error GoTo 0
End Sub

Sub ColorIndexValues()
    Dim ws As Worksheet: Set ws = Sheets(SH_TRADES)
    Dim rng As Range: Set rng = ws.Range("K2")
    ws.Range("K2").value = ws.Range("K1").value
    Dim cell As Range: For Each cell In rng
        If cell.value <> "" Then
            Dim txt As String: txt = cell.value: Dim startPos As Long: startPos = 1: Dim openP As Long, closeP As Long
            Do
                openP = InStr(startPos, txt, "- "): If openP = 0 Then Exit Do
                closeP = InStr(openP, txt, "("): If closeP = 0 Then Exit Do
                Dim numStr As String: numStr = Mid(txt, openP + 1, closeP - openP - 1)
                If IsNumeric(numStr) Then
                    Dim v As Double: v = CDbl(numStr): Dim clr As Long: If Mid(txt, closeP + 1, 1) = "-" Then clr = RGB(200, 0, 0) Else clr = RGB(0, 128, 0)
                    cell.Characters(openP + 1, Len(numStr)).Font.Color = clr: cell.Characters(openP + 1, Len(numStr)).Font.Bold = True
                End If
                startPos = closeP + 1
            Loop
        End If
    Next cell
End Sub

' --- STANDALONE BULLETIN SHEET POPULATION ---
Sub Populate_Bulletin_Sheet(wsM As Worksheet)
    On Error Resume Next
    Dim wsB As Worksheet: Set wsB = Sheets(SH_BULLETIN)
    If wsB Is Nothing Then Exit Sub
    
    ' 1. Clear existing data
    Dim lastRowB As Long: lastRowB = wsB.cells(wsB.Rows.count, 1).End(xlUp).row
    If lastRowB >= 4 Then wsB.Range("A4:C" & lastRowB).ClearContents
    
    Dim maxRows As Long: maxRows = wsM.cells(wsM.Rows.count, 1).End(xlUp).row - ROW_START_M + 1
    If maxRows < 1 Then Exit Sub
    
    Dim outData() As Variant
    ReDim outData(1 To maxRows, 1 To 3)
    
    ' 2. Populate ALL symbols in Sheet1 row order (including indices)
    Dim r As Long, outIdx As Long: outIdx = 1
    For r = ROW_START_M To wsM.cells(wsM.Rows.count, 1).End(xlUp).row
        Dim ticker As String: ticker = Trim(wsM.cells(r, 1).value)
        If ticker = "" Then GoTo NextR
        
        ' --- READ INPUTS ---
        Dim ltp As Double: ltp = wsM.cells(r, RAW_LTP).value
        Dim chg As Double: chg = wsM.cells(r, RAW_CHANGE_PTS).value
        Dim st As Double: st = wsM.cells(r, RAW_SUPERTREND).value
        Dim ema20 As Double: ema20 = val(wsM.cells(r, RAW_EMA20_5M).value)
        Dim adxVal As Double: adxVal = val(wsM.cells(r, "CB").value)
        Dim score As String: score = wsM.cells(r, "DE").Text
        Dim regime As String: regime = wsM.cells(r, "BQ").Text
        
        Dim piv As Double: piv = wsM.cells(r, RAW_PIVOT).value
        Dim r1 As Double: r1 = wsM.cells(r, RAW_R1).value
        Dim s1 As Double: s1 = wsM.cells(r, RAW_S1).value
        
        Dim pct As Double: pct = Calc_Pct_Change(ltp, chg)
        Dim emaDist As Double: emaDist = 0: If ema20 > 0 Then emaDist = ((ltp - ema20) / ema20) * 100
        Dim trendDesc As String: trendDesc = IIf(ltp > st, IIf(ltp > ema20, "STRONG BULLISH", "WEAK BULLISH"), IIf(ltp < ema20, "STRONG BEARISH", "WEAK BEARISH"))
        
        Dim nearSup As Double, nearRes As Double, supName As String, resName As String
        If ltp > piv Then nearSup = piv: supName = "Pivot" Else nearSup = s1: supName = "S1"
        If ltp < piv Then nearRes = piv: resName = "Pivot" Else nearRes = r1: resName = "R1"

        ' --- Build status text (same format for indices and stocks) ---
        Dim statusText As String
        
        ' Check if this is an index for GLOBAL CONTEXT label
        If Is_Index(ticker) Then
            Dim nSentVal As Double: nSentVal = val(wsM.cells(r, "DN").value)
            statusText = "GLOBAL CONTEXT | " & ticker & ": " & Format(ltp, "#,##0.00") & " (" & IIf(chg > 0, "+", "") & Format(pct, "0.00") & "%) | ADX: " & Format(adxVal, "0") & vbNewLine & _
                         "STATUS: " & trendDesc & " (AI: " & score & ") | REGIME: " & regime & " | SENTI: " & Format(nSentVal, "0.0")
        Else
            statusText = ticker & ": " & Format(ltp, "#,##0.00") & " (" & IIf(chg > 0, "+", "") & Format(pct, "0.00") & "%) | ADX: " & Format(adxVal, "0") & vbNewLine & _
                         "STATUS: " & trendDesc & " (AI: " & score & ") | REGIME: " & regime & vbNewLine & _
                         "BEHAVIOR: Currently " & Format(Abs(emaDist), "0.00") & "% " & IIf(emaDist >= 0, "above", "below") & " 20-EMA." & vbNewLine & _
                         "TACTICAL: " & resName & " [+" & Format(Abs(nearRes - ltp), "0") & "] | " & supName & " [-" & Format(Abs(ltp - nearSup), "0") & "]"
        End If

        Dim summary As String: summary = wsM.cells(r, RAW_INST_SUMMARY).Text
        summary = Replace(summary, "Institutional Summary for this Scan:", "")
        summary = Replace(summary, "Verdict", "VERDICT"): summary = Replace(summary, "Action", "ACTION")
        summary = Trim(summary)
        
        ' Normalize all line breaks
        summary = Replace(summary, vbCrLf, vbLf)
        summary = Replace(summary, vbCr, vbLf)

        ' Collapse ANY double blank lines
        Do While InStr(summary, vbLf & vbLf) > 0
            summary = Replace(summary, vbLf & vbLf, vbLf)
        Loop
        
        ' Convert back
        summary = Replace(summary, vbLf, vbCrLf)

        ' Add to array
        outData(outIdx, 1) = ticker
        outData(outIdx, 2) = statusText
        outData(outIdx, 3) = summary
        outIdx = outIdx + 1
NextR:
    Next r
    
    ' 3. Bulk Write Array to Excel
    If outIdx > 1 Then
        wsB.Range("A4").Resize(outIdx - 1, 3).value = outData
        With wsB.Range("B4:C" & (3 + outIdx - 1))
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
End Sub

' =========================================================================================
' COLUMN N: REAL-TIME TRADE COACH
' =========================================================================================
Sub Update_Trade_Coach(wsP As Worksheet, r As Long, ltp As Double, dir As String, status As String, Optional aiConf As Double = 0, Optional aiGrade As String = "")
    Dim cCell As Range: Set cCell = wsP.cells(r, P_COACH)
    Dim act As String: act = UCase(Trim(wsP.cells(r, P_ACTION).value))
    
    If status = "CLOSED" Or status = "CANCELLED" Then
        cCell.value = "FINISHED | " & status & vbLf & "WHEN: N/A | HOW: N/A"
        cCell.Font.Color = RGB(128, 128, 128)
        cCell.Font.Bold = False
        cCell.HorizontalAlignment = xlCenter
        cCell.VerticalAlignment = xlCenter
        cCell.WrapText = True
        Exit Sub
    End If
    
    If status <> "OPEN" And status <> "SIGNAL" And status <> "PENDING" Then Exit Sub
    
    Dim en As Double: en = val(wsP.cells(r, P_ENTRY).value)
    Dim sl As Double: sl = val(wsP.cells(r, P_SL).value)
    Dim tp As Double: tp = val(wsP.cells(r, P_TARGET).value)
    
    If en = 0 Or sl = 0 Or tp = 0 Then Exit Sub
    
    Dim risk As Double: risk = Abs(en - sl)
    If risk < 0.0001 Then Exit Sub
    
    Dim rLtp As Double
    If dir = "BUY" Then
        rLtp = (ltp - en) / risk
    Else
        rLtp = (en - ltp) / risk
    End If
    
    ' Check bars elapsed for Dead Money
    Dim minElapsed As Double: minElapsed = 0
    Dim openTimeStr As String: openTimeStr = wsP.cells(r, P_OPEN_TIME).Text
    If status = "OPEN" And openTimeStr <> "" And IsDate(openTimeStr) Then
        minElapsed = (Now - Date - TimeValue(openTimeStr)) * 1440
        If minElapsed < 0 Then minElapsed = minElapsed + 1440 ' Cross midnight fix
    End If
    
    ' Analyze Momentum Death
    Dim prevMsg As String: prevMsg = cCell.value
    Dim wasFlying As Boolean: wasFlying = (InStr(1, prevMsg, "FLYING", vbTextCompare) > 0 Or _
                                           InStr(1, prevMsg, "BOOK 50%", vbTextCompare) > 0 Or _
                                           InStr(1, prevMsg, "+1R", vbTextCompare) > 0)
    
    ' Extract lifecycle state and execution style from P_COMM (Column K)
    Dim commVal As String: commVal = wsP.cells(r, P_COMM).value
    Dim pState As String: pState = Extract_Lifecycle_From_Cockpit(commVal)
    Dim execStyle As String: execStyle = Extract_Execution_From_Cockpit(commVal)
    
    Dim actionStr As String: actionStr = ""
    Dim reasonStr As String: reasonStr = ""
    Dim whenStr As String: whenStr = ""
    Dim howStr As String: howStr = ""
    
    Dim suggestedTrail As Double: suggestedTrail = en
    If dir = "BUY" Then
        If rLtp >= 1.4 Then
            suggestedTrail = en + risk
        ElseIf rLtp >= 1.2 Then
            suggestedTrail = en + 0.5 * risk
        End If
    Else
        If rLtp >= 1.4 Then
            suggestedTrail = en - risk
        ElseIf rLtp >= 1.2 Then
            suggestedTrail = en - 0.5 * risk
        End If
    End If
    
    ' Assign Scenarios
    If wasFlying And rLtp < 0.6 And rLtp > -0.5 Then
        actionStr = "EXIT NOW"
        reasonStr = "Momentum Fading"
        whenStr = "Immediately"
        howStr = "Market Exit at " & Format(ltp, "0.00")
    ElseIf status = "OPEN" And (aiGrade = "D" Or aiGrade = "NONE") And aiConf > 0 And Abs(rLtp) > 0.3 Then
        actionStr = "EXIT NOW"
        reasonStr = "AI Fade"
        whenStr = "Next bar"
        howStr = "Market Exit at " & Format(ltp, "0.00")
    ElseIf pState = "BROKEN" Or rLtp <= -0.8 Then
        actionStr = "EXIT NOW"
        If pState = "BROKEN" Then
            reasonStr = "Thesis Broken"
            whenStr = "Next bar"
        Else
            reasonStr = "Stop Breached"
            whenStr = "Immediately"
        End If
        howStr = "Market Exit at " & Format(ltp, "0.00")
    ElseIf rLtp >= 1# And InStr(1, wsP.cells(r, P_ACTION).value, "50%", vbTextCompare) = 0 Then
        actionStr = "BOOK 50%"
        reasonStr = "Reached +1.0R"
        whenStr = "Immediately"
        howStr = "Sell 50% at " & Format(ltp, "0.00") & " & SL to Cost " & Format(en, "0.00")
    ElseIf rLtp >= 1.2 Then
        actionStr = "TRAIL STOP"
        reasonStr = "Reached +" & Format(rLtp, "0.1") & "R"
        whenStr = "Immediately"
        howStr = "Trail SL to " & Format(suggestedTrail, "0.00")
    ElseIf minElapsed >= 90 And rLtp > -0.2 And rLtp < 0.2 Then
        actionStr = "EXIT NOW"
        reasonStr = "Dead Money (90m)"
        whenStr = "Immediately"
        howStr = "Market Exit at " & Format(ltp, "0.00") & " (Scratch)"
    ElseIf pState = "WEAKENING" Then
        actionStr = "HOLD"
        reasonStr = "Momentum Fading"
        whenStr = "In Play"
        howStr = "Stop at " & Format(sl, "0.00") & " & Don't Add"
    ElseIf pState = "MATURING" Then
        actionStr = "HOLD"
        reasonStr = "Thesis Healthy (Maturing)"
        whenStr = "In Play"
        howStr = "Stop at " & Format(sl, "0.00")
    Else
        If status = "SIGNAL" Or status = "PENDING" Then
            actionStr = "WAIT FOR ENTRY"
            If InStr(execStyle, "BREAKOUT") > 0 Then
                reasonStr = "Breakout Style"
            ElseIf InStr(execStyle, "PULLBACK") > 0 Then
                reasonStr = "Pullback Style"
            Else
                reasonStr = "Entry Pending"
            End If
            whenStr = "Touch " & Format(en, "0.00")
            howStr = "Limit Order at " & Format(en, "0.00")
        Else
            actionStr = "HOLD"
            reasonStr = "Thesis Healthy"
            whenStr = "In Play"
            howStr = "Stop at " & Format(sl, "0.00")
        End If
    End If
    
    ' Output exactly 2 lines
    Dim cellText As String
    cellText = actionStr & " | " & reasonStr & vbLf & "WHEN: " & whenStr & " | HOW: " & howStr
    cCell.value = cellText
    
    ' Set default formatting for the entire cell: regular font and soft dark gray
    cCell.Font.Bold = False
    cCell.Font.Color = RGB(120, 120, 120) ' Soft dark gray
    cCell.HorizontalAlignment = xlCenter
    cCell.VerticalAlignment = xlCenter
    cCell.WrapText = True
    
    ' Format Action command (the part before the first '|')
    Dim actionClr As Long
    If actionStr = "EXIT NOW" Then
        actionClr = RGB(192, 0, 0) ' Red
    ElseIf actionStr = "BOOK 50%" Or actionStr = "TRAIL STOP" Then
        actionClr = RGB(0, 128, 0) ' Green
    ElseIf actionStr = "HOLD" Then
        If InStr(reasonStr, "Fading") > 0 Or InStr(reasonStr, "Don't Add") > 0 Then
            actionClr = RGB(200, 110, 0) ' Orange
        Else
            actionClr = RGB(0, 128, 0) ' Green
        End If
    ElseIf actionStr = "WAIT FOR ENTRY" Then
        actionClr = RGB(105, 105, 105) ' Gray
    Else
        actionClr = RGB(0, 0, 0) ' Black
    End If
    
    ' Bold the Action part
    Dim pPipe As Long: pPipe = InStr(1, cellText, "|")
    If pPipe > 0 Then
        With cCell.Characters(1, pPipe - 1).Font
            .Bold = True
            .Color = actionClr
        End With
    End If
    
    ' Format reason text (rest of line 1 and line 2)
    ' 1. Key words -> Bold Dark Black (RGB 0,0,0)
    Highlight_Cell_Substring cCell, "Thesis", RGB(0, 0, 0), True, pPipe
    Highlight_Cell_Substring cCell, "Maturing", RGB(0, 0, 0), True, pPipe
    Highlight_Cell_Substring cCell, "Waiting", RGB(0, 0, 0), True, pPipe
    Highlight_Cell_Substring cCell, "breakout", RGB(0, 0, 0), True, pPipe
    Highlight_Cell_Substring cCell, "pullback", RGB(0, 0, 0), True, pPipe
    Highlight_Cell_Substring cCell, "entry", RGB(0, 0, 0), True, pPipe
    Highlight_Cell_Substring cCell, "Healthy", RGB(0, 0, 0), True, pPipe
    
    ' 2. Attention/action phrases -> Bold and colored
    Highlight_Cell_Substring cCell, "Thesis Broken", RGB(192, 0, 0), True, pPipe
    Highlight_Cell_Substring cCell, "Stop Breached", RGB(192, 0, 0), True, pPipe
    Highlight_Cell_Substring cCell, "AI Fade", RGB(192, 0, 0), True, pPipe
    
    ' Highlight the prices and trailing stop values
    Highlight_Cell_Substring cCell, Format(en, "0.00"), RGB(0, 0, 0), True, pPipe
    Highlight_Cell_Substring cCell, Format(sl, "0.00"), RGB(192, 0, 0), True, pPipe ' Red for SL
    Highlight_Cell_Substring cCell, Format(ltp, "0.00"), RGB(0, 102, 204), True, pPipe ' Blue for LTP
    If suggestedTrail <> en Then
        Highlight_Cell_Substring cCell, Format(suggestedTrail, "0.00"), RGB(0, 128, 0), True, pPipe ' Green for trailing SL
    End If
    
    ' Highlight cautions
    Highlight_Cell_Substring cCell, "Momentum Fading", RGB(200, 110, 0), True, pPipe
    Highlight_Cell_Substring cCell, "Don't Add", RGB(200, 110, 0), True, pPipe
    Highlight_Cell_Substring cCell, "Dead Money", RGB(200, 110, 0), True, pPipe
    
    ' 3. Dynamic R-value highlighting (e.g. Reached +1.0R, Reached +1.2R)
    Dim pReached As Long: pReached = InStr(pPipe, cellText, "Reached +", vbTextCompare)
    If pReached >= pPipe Then
        Dim pREnd As Long: pREnd = InStr(pReached, cellText, "R", vbTextCompare)
        If pREnd > pReached Then
            With cCell.Characters(pReached, pREnd - pReached + 1).Font
                .Bold = True
                .Color = RGB(0, 128, 0) ' Green
            End With
        End If
    End If
End Sub


' --- NEW: CONFLUENCE SCIENCE TRANSLATOR ---
Function Translate_Science_Tags(rawTags As String) As String
    If rawTags = "" Then
        Translate_Science_Tags = ""
        Exit Function
    End If
    
    Dim tagArray() As String
    tagArray = Split(rawTags, ",")
    
    Dim finalStr As String
    finalStr = ""
    
    Dim i As Long
    Dim tag As String
    Dim cleanTag As String
    Dim explanation As String
    
    For i = 0 To UBound(tagArray)
        tag = Trim(tagArray(i))
        cleanTag = ""
        explanation = ""
        
        If tag <> "" Then
            ' Skip redundant labels explicitly so they don't print in the Science section
            If InStr(1, tag, "RR:", vbTextCompare) = 1 Or _
               InStr(1, tag, "NetRR:", vbTextCompare) = 1 Or _
               InStr(1, tag, "Risk/Share:", vbTextCompare) = 1 Or _
               InStr(1, tag, "Exec:", vbTextCompare) = 1 Or _
               InStr(1, tag, "SECTOR_", vbTextCompare) = 1 Or _
               InStr(1, tag, "Setup:", vbTextCompare) = 1 Then
                ' Do nothing
            ' Trend Layer
            ElseIf tag = "D_EMA_BULL" Then
                cleanTag = "Daily EMAs Aligned": explanation = "Macro 10>20>50 strong tailwind"
            ElseIf tag = "D_EMA_BEAR" Then
                cleanTag = "Daily EMAs Aligned": explanation = "Macro 10<20<50 strong headwind"
                'End If
            ElseIf tag = "D_ST_GREEN" Or tag = "D_ST_RED" Then
                cleanTag = "Daily SuperTrend": explanation = "Core trend path of least resistance"
            ElseIf InStr(tag, "D_RSI_BULL") > 0 Or InStr(tag, "D_RSI_BEAR") > 0 Then
                cleanTag = "Daily RSI": explanation = "Macro momentum supports the trade"
            ElseIf tag = "1H_CONFIRMS_BULL" Or tag = "1H_CONFIRMS_BEAR" Then
                cleanTag = "1H Trend Sync": explanation = "Hourly mirrors the daily trend"
            
            ' Structure Layer
            ElseIf tag = "VWAP_ABOVE" Then
                cleanTag = "Above VWAP": explanation = "Aligned with institutional volume flow"
            ElseIf tag = "VWAP_BELOW" Then
                cleanTag = "Below VWAP": explanation = "Aligned with institutional volume flow"
            ElseIf tag = "15M_ST_GREEN" Or tag = "15M_ST_RED" Then
                cleanTag = "15M SuperTrend": explanation = "Intraday structure is intact"
            ElseIf tag = "15M_EMA_BULL" Or tag = "15M_EMA_BEAR" Then
                cleanTag = "15M EMA Cross": explanation = "Fresh intraday structural shift"
            
            ' Trigger Layer
            ElseIf tag = "NEAR_SUPPORT" Or tag = "NEAR_RESISTANCE" Then
                cleanTag = "Near Pivot/SR": explanation = "Risk defined at a major floor/ceiling"
            ElseIf tag = "ADX_BULL_TREND" Or tag = "ADX_BEAR_TREND" Then
                cleanTag = "ADX Impulsive": explanation = "High-velocity momentum confirmed"
            ElseIf tag = "BB_TOP" Or tag = "BB_BOTTOM" Then
                cleanTag = "Volatility Squeeze": explanation = "Explosive volume expansion breakout"
            
            ' RSI parsing
            ElseIf InStr(tag, "RSI_5M_ZONE(") > 0 Or InStr(tag, "RSI_5M_ZONE_SHORT(") > 0 Then
                Dim rVal As String: rVal = Split(Split(tag, "(")(1), ")")(0)
                cleanTag = "5M Action Zone": explanation = "RSI at " & rVal & ", plenty of room to run"
            ElseIf InStr(tag, "RSI_5M_OVERSOLD(") > 0 Or InStr(tag, "RSI_5M_OVERBOUGHT(") > 0 Then
                Dim rValEx As String: rValEx = Split(Split(tag, "(")(1), ")")(0)
                cleanTag = "5M RSI Extreme": explanation = "RSI at " & rValEx & ", sharp mean-reversion snapback"
            
            ' Candlestick parsing
            ElseIf Left(tag, 3) = "CS_" Then
                Dim patName As String: patName = Mid(tag, 4)
                cleanTag = "Candle (" & patName & ")"
                If InStr(patName, "Marubozu") > 0 Then
                    explanation = "Strong opening to closing momentum"
                ElseIf InStr(patName, "Hammer") > 0 Or InStr(patName, "Pin Bar") > 0 Then
                    explanation = "Sharp rejection of lower prices"
                ElseIf InStr(patName, "Shooting Star") > 0 Then
                    explanation = "Sharp rejection of higher prices"
                ElseIf InStr(patName, "Engulfing") > 0 Then
                    explanation = "Complete reversal of prior candle sentiment"
                ElseIf InStr(patName, "Harami") > 0 Or InStr(patName, "Inside") > 0 Then
                    explanation = "Volatility contraction before breakout"
                ElseIf InStr(patName, "Morning Star") > 0 Or InStr(patName, "Evening Star") > 0 Then
                    explanation = "Classic 3-bar reversal sequence"
                Else
                    explanation = "High-probability price action confirmation"
                End If
            End If
            
            If cleanTag <> "" Then
                finalStr = finalStr & vbNewLine & " - " & cleanTag & ": " & explanation
            ElseIf InStr(1, tag, "RR:", vbTextCompare) <> 1 And InStr(1, tag, "NetRR:", vbTextCompare) <> 1 And InStr(1, tag, "Risk/Share:", vbTextCompare) <> 1 And InStr(1, tag, "Exec:", vbTextCompare) <> 1 And InStr(1, tag, "SECTOR_", vbTextCompare) <> 1 And InStr(1, tag, "Setup:", vbTextCompare) <> 1 Then
                finalStr = finalStr & vbNewLine & " - " & tag
            End If
        End If
    Next i
    
    Translate_Science_Tags = finalStr
End Function




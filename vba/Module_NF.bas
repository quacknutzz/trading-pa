Attribute VB_Name = "Module_NF"
' ---------------- Bias Narratives & Gate (5m / 15m / 1h) ----------------
Option Explicit

' ====== USER SETTINGS - change header names here if different ======
Private Const H_EMA10_5 As String = "EMA10_5m"
Private Const H_EMA30_5 As String = "EMA30_5m"
Private Const H_EMA10_15 As String = "EMA10_15m"
Private Const H_EMA30_15 As String = "EMA30_15m"
Private Const H_EMA10_1H As String = "EMA10_1h"
Private Const H_EMA30_1H As String = "EMA30_1h"

' If True, 1h disagreement is advisory only (does not auto-block). If False, 1h disagreement blocks trade.
Private Const ONEH_AS_ADVISORY As Boolean = True

' Minimum consecutive bars to consider "holding" for narrative clarity
Private Const MIN_HOLD_BARS As Long = 1

Public Sub NF_ComputeBiasNarratives()
    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets("Sheet1")
    Dim wsOut As Worksheet: Set wsOut = ThisWorkbook.Worksheets("NF_Outputs")
    Dim lastRow As Long
    
    ' Validate NF_Outputs exists
    On Error GoTo ErrSheet
    lastRow = wsOut.cells(wsOut.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "NF_Outputs looks empty — run NF_BuildOutputs_Step1 first.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Read Sheet1 headers to find EMA columns (fall back to names if not found)
    Dim hdrCol As Object: Set hdrCol = BuildHeaderMap(wsSrc)
    If hdrCol Is Nothing Then
        MsgBox "Cannot read headers from Sheet1 (row1).", vbExclamation: Exit Sub
    End If
    
    ' Validate EMA headers exist
    Dim needed As Variant: needed = Array(H_EMA10_5, H_EMA30_5, H_EMA10_15, H_EMA30_15, H_EMA10_1H, H_EMA30_1H)
    Dim i As Long
    For i = LBound(needed) To UBound(needed)
        If Not hdrCol.Exists(UCase(needed(i))) Then
            MsgBox "Header '" & needed(i) & "' not found in Sheet1. Please adjust constants or your headers.", vbCritical
            Exit Sub
        End If
    Next i
    
    ' Bulk-read required columns from Sheet1 into array (we need all EMA columns)
    Dim lastSrcRow As Long: lastSrcRow = wsSrc.cells(wsSrc.Rows.count, "A").End(xlUp).row
    Dim srcRange As Variant
    srcRange = wsSrc.Range("A2:" & wsSrc.cells(lastSrcRow, 1).Address(False, False)).Resize(lastSrcRow - 1, 1).Value2 ' placeholder read to ensure sheet exists (we will use header map to read columns)
    ' Build arrays per needed EMA column
    Dim arr10_5() As Variant, arr30_5() As Variant
    Dim arr10_15() As Variant, arr30_15() As Variant
    Dim arr10_1h() As Variant, arr30_1h() As Variant
    arr10_5 = wsSrc.Range(wsSrc.cells(2, hdrCol(UCase(H_EMA10_5))), wsSrc.cells(lastSrcRow, hdrCol(UCase(H_EMA10_5)))).Value2
    arr30_5 = wsSrc.Range(wsSrc.cells(2, hdrCol(UCase(H_EMA30_5))), wsSrc.cells(lastSrcRow, hdrCol(UCase(H_EMA30_5)))).Value2
    arr10_15 = wsSrc.Range(wsSrc.cells(2, hdrCol(UCase(H_EMA10_15))), wsSrc.cells(lastSrcRow, hdrCol(UCase(H_EMA10_15)))).Value2
    arr30_15 = wsSrc.Range(wsSrc.cells(2, hdrCol(UCase(H_EMA30_15))), wsSrc.cells(lastSrcRow, hdrCol(UCase(H_EMA30_15)))).Value2
    arr10_1h = wsSrc.Range(wsSrc.cells(2, hdrCol(UCase(H_EMA10_1H))), wsSrc.cells(lastSrcRow, hdrCol(UCase(H_EMA10_1H)))).Value2
    arr30_1h = wsSrc.Range(wsSrc.cells(2, hdrCol(UCase(H_EMA30_1H))), wsSrc.cells(lastSrcRow, hdrCol(UCase(H_EMA30_1H)))).Value2
    
    ' Now read NF_Outputs symbol column to align rows (we assume Sheet1 and NF_Outputs row order match symbols)
    Dim outLast As Long: outLast = wsOut.cells(wsOut.Rows.count, "A").End(xlUp).row
    Dim symbols As Variant: symbols = wsOut.Range("A2:A" & outLast).Value2
    
    ' Prepare output arrays
    Dim outBias5() As Variant, outBias15() As Variant, outBias1h() As Variant
    Dim outGate() As Variant, outNarr() As Variant
    ReDim outBias5(1 To outLast - 1, 1 To 1)
    ReDim outBias15(1 To outLast - 1, 1 To 1)
    ReDim outBias1h(1 To outLast - 1, 1 To 1)
    ReDim outGate(1 To outLast - 1, 1 To 1)
    ReDim outNarr(1 To outLast - 1, 1 To 1)
    
    ' We'll map output rows to Sheet1 rows by position: assume row 2 in NF_Outputs corresponds to row 2 in Sheet1 etc.
    Dim rOut As Long, rSrc As Long
    For rOut = 1 To outLast - 1
        rSrc = rOut ' mapping 1-to-1: row2->index1
    
        Dim b5 As String, b15 As String, b1h As String
        b5 = NF_BiasFromEMAs(arr10_5(rSrc, 1), arr30_5(rSrc, 1))
        b15 = NF_BiasFromEMAs(arr10_15(rSrc, 1), arr30_15(rSrc, 1))
        b1h = NF_BiasFromEMAs(arr10_1h(rSrc, 1), arr30_1h(rSrc, 1))
        
        outBias5(rOut, 1) = b5
        outBias15(rOut, 1) = b15
        outBias1h(rOut, 1) = b1h
        
        ' count consecutive holding bars for narrative
        Dim hold5 As Long: hold5 = NF_CountConsecutive(wsSrc, hdrCol(UCase(H_EMA10_5)), hdrCol(UCase(H_EMA30_5)), lastSrcRow, rSrc + 1, b5)
        Dim hold15 As Long: hold15 = NF_CountConsecutive(wsSrc, hdrCol(UCase(H_EMA10_15)), hdrCol(UCase(H_EMA30_15)), lastSrcRow, rSrc + 1, b15)
        Dim hold1h As Long: hold1h = NF_CountConsecutive(wsSrc, hdrCol(UCase(H_EMA10_1H)), hdrCol(UCase(H_EMA30_1H)), lastSrcRow, rSrc + 1, b1h)
        
        ' Build gate & narrative
        Dim gate As String, narr As String
        gate = NF_DecideGate(b5, b15, b1h)
        narr = NF_BuildNarrative(b5, hold5, b15, hold15, b1h, hold1h, gate)
        
        outGate(rOut, 1) = gate
        outNarr(rOut, 1) = narr
    Next rOut
    
    ' Write outputs in one go into NF_Outputs columns M:Q (or adjust columns as you prefer)
    ' We'll look for existing header columns; if not present add them at the end.
    Dim cBias5 As Long: cBias5 = NF_EnsureOutCol(wsOut, "Bias_5m")
    Dim cBias15 As Long: cBias15 = NF_EnsureOutCol(wsOut, "Bias_15m")
    Dim cBias1h As Long: cBias1h = NF_EnsureOutCol(wsOut, "Bias_1h")
    Dim cGate As Long: cGate = NF_EnsureOutCol(wsOut, "BiasGate")
    Dim cNarr As Long: cNarr = NF_EnsureOutCol(wsOut, "BiasNarrative")
    
    wsOut.Range(wsOut.cells(2, cBias5), wsOut.cells(1 + UBound(outBias5, 1), cBias5)).value = outBias5
    wsOut.Range(wsOut.cells(2, cBias15), wsOut.cells(1 + UBound(outBias15, 1), cBias15)).value = outBias15
    wsOut.Range(wsOut.cells(2, cBias1h), wsOut.cells(1 + UBound(outBias1h, 1), cBias1h)).value = outBias1h
    wsOut.Range(wsOut.cells(2, cGate), wsOut.cells(1 + UBound(outGate, 1), cGate)).value = outGate
    wsOut.Range(wsOut.cells(2, cNarr), wsOut.cells(1 + UBound(outNarr, 1), cNarr)).value = outNarr
    
    MsgBox "Bias narratives computed and written to NF_Outputs.", vbInformation
    Exit Sub

ErrSheet:
    MsgBox "NF_Outputs or Sheet1 not found. Create NF_Outputs and ensure Sheet1 has EMA columns.", vbCritical
End Sub

' ---------------- helper functions ----------------
Private Function BuildHeaderMap(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = ws.cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        Dim h As String: h = Trim(UCase(CStr(ws.cells(1, c).value)))
        If h <> "" Then d(h) = c
    Next c
    Set BuildHeaderMap = d
End Function

'=============================================================
' Return Bull/Bear/Neutral from two EMA values (fast vs slow)
'=============================================================
Public Function NF_BiasFromEMAs(emaFast As Variant, emaSlow As Variant) As String
    If Not IsNumeric(emaFast) Or Not IsNumeric(emaSlow) Then
        NF_BiasFromEMAs = "Neutral"
    ElseIf CDbl(emaFast) > CDbl(emaSlow) Then
        NF_BiasFromEMAs = "Bull"
    ElseIf CDbl(emaFast) < CDbl(emaSlow) Then
        NF_BiasFromEMAs = "Bear"
    Else
        NF_BiasFromEMAs = "Neutral"
    End If
End Function

' Count consecutive bars (backwards) from rowIndex in Sheet1 where EMA10>EMA30 matches targetBias
' rowIndex is the actual sheet row (2..n)
Private Function NF_CountConsecutive(ws As Worksheet, col10 As Long, col30 As Long, lastRow As Long, rowIndex As Long, targetBias As String) As Long
    Dim cnt As Long: cnt = 0
    Dim r As Long
    For r = rowIndex To 2 Step -1
        Dim v10 As Variant: v10 = ws.cells(r, col10).Value2
        Dim v30 As Variant: v30 = ws.cells(r, col30).Value2
        Dim b As String: b = NF_BiasFromEMAs(v10, v30)
        If b = targetBias Then
            cnt = cnt + 1
        Else
            Exit For
        End If
    Next r
    NF_CountConsecutive = cnt
End Function

'--- Decide gate based on OHLCV or EMA status in arr()
'-----------------------------------------------------------------
' Decides if a row has enough OHLCV data to be considered valid
' Works directly on the srcArr array (from Sheet1)
'-----------------------------------------------------------------
'=============================================================
' Decide whether row passes "gate" based on 5m/15m/1h biases (robust to "Bullish"/"Bull"/"bearish", etc.)
Public Function NF_DecideGate(bias5 As Variant, bias15 As Variant, bias1h As Variant) As Boolean
    Dim bullCount As Long, bearCount As Long
    Dim s As String

    bullCount = 0: bearCount = 0

    If Not IsError(bias5) And Not IsEmpty(bias5) Then
        s = LCase$(CStr(bias5))
        If InStr(s, "bull") > 0 Then bullCount = bullCount + 1
        If InStr(s, "bear") > 0 Then bearCount = bearCount + 1
    End If

    If Not IsError(bias15) And Not IsEmpty(bias15) Then
        s = LCase$(CStr(bias15))
        If InStr(s, "bull") > 0 Then bullCount = bullCount + 1
        If InStr(s, "bear") > 0 Then bearCount = bearCount + 1
    End If

    If Not IsError(bias1h) And Not IsEmpty(bias1h) Then
        s = LCase$(CStr(bias1h))
        If InStr(s, "bull") > 0 Then bullCount = bullCount + 1
        If InStr(s, "bear") > 0 Then bearCount = bearCount + 1
    End If

    ' Pass gate if at least 2 TFs agree in the same direction
    NF_DecideGate = (bullCount >= 2 Or bearCount >= 2)
End Function
Private Function NF_BuildNarrative(b5 As String, h5 As Long, b15 As String, h15 As Long, b1h As String, h1h As Long, gate As String) As String
    Dim s As String
    s = ""
    s = s & "5m: " & b5 & " (holding " & h5 & " bar" & IIf(h5 = 1, "", "s") & "); "
    s = s & "15m: " & b15 & " (holding " & h15 & " bar" & IIf(h15 = 1, "", "s") & "); "
    s = s & "1h: " & b1h & " (holding " & h1h & " bar" & IIf(h1h = 1, "", "s") & "). "
    Select Case gate
        Case "PASS"
            s = "PASS — " & s & "Aligned across 5m/15m" & IIf(b1h = b15, " and 1h.", " (1h advisory).")
        Case "ADVISORY_BLOCK"
            s = "ADVISORY_BLOCK — " & s & "15m agrees with 5m but 1h conflicts (use discretion)."
        Case "BLOCK_1h"
            s = "BLOCK — " & s & "15m and 5m align but 1h conflicts (blocking)."
        Case "BLOCK_5m_CONFLICT"
            s = "BLOCK — " & s & "5m conflicts with 15m; wait for 5m to realign."
        Case "BLOCK_NEUTRAL_15m"
            s = "BLOCK — " & s & "15m neutral; skip trades."
        Case Else
            s = gate & " — " & s
    End Select
    NF_BuildNarrative = s
End Function

Private Function NF_EnsureOutCol(ws As Worksheet, header As String) As Long
    Dim lastCol As Long: lastCol = ws.cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If Trim(UCase(ws.cells(1, c).value)) = Trim(UCase(header)) Then
            NF_EnsureOutCol = c: Exit Function
        End If
    Next c
    ws.cells(1, lastCol + 1).value = header
    NF_EnsureOutCol = lastCol + 1
End Function


'-----------------------------------------------------------------
' Scores one row combining EMA bias, MACD, ATR, VWAP, Vol Spike
' srcArr = Sheet1 array (OHLCV + indicators)
' tradeArr = Trades sheet array (MACD etc)
' srcRow = row index in srcArr
' tradeRow = row index in tradeArr
' Returns: Variant() = {Score, Narrative}
'-----------------------------------------------------------------
Public Function NF_ScoreRow(srcArr As Variant, tradeArr As Variant, _
                            srcRow As Long, tradeRow As Long, _
                            bias5 As Variant, bias15 As Variant, bias1h As Variant) As Variant
    Dim score As Long: score = 0
    Dim supList As String, oppList As String, macdList As String
    Dim narr As String
    Dim s As String
    Dim isBull As Boolean, isBear As Boolean

    ' --- EMA Bias contribution ---
    s = LCase$(CStr(bias5))
    isBull = (InStr(s, "bull") > 0)
    isBear = (InStr(s, "bear") > 0)
    If isBull Then
        score = score + 1: supList = supList & "EMA 5m, "
    ElseIf isBear Then
        score = score - 1: oppList = oppList & "EMA 5m, "
    End If

    s = LCase$(CStr(bias15))
    isBull = (InStr(s, "bull") > 0)
    isBear = (InStr(s, "bear") > 0)
    If isBull Then
        score = score + 2: supList = supList & "EMA 15m, "
    ElseIf isBear Then
        score = score - 2: oppList = oppList & "EMA 15m, "
    End If

    s = LCase$(CStr(bias1h))
    isBull = (InStr(s, "bull") > 0)
    isBear = (InStr(s, "bear") > 0)
    If isBull Then
        score = score + 3: supList = supList & "EMA 1h, "
    ElseIf isBear Then
        score = score - 3: oppList = oppList & "EMA 1h, "
    End If

'    ' --- ATR (cols CC=81, CD=82) ---
'    On Error Resume Next
'    If IsNumeric(srcArr(srcRow, 81)) Then score = score + 1: supList = supList & "ATR5m, "
'    If IsNumeric(srcArr(srcRow, 82)) Then score = score + 1: supList = supList & "ATR15m, "
'    On Error GoTo 0


    ' --- ATR (cols CC=81, CD=82) ---
    On Error Resume Next
    If IsNumeric(srcArr(srcRow, 81)) Then
        If InStr(1, LCase$(bias5), "bull") > 0 Then
            score = score + 1: supList = supList & "ATR5m, "
        ElseIf InStr(1, LCase$(bias5), "bear") > 0 Then
            score = score - 1: oppList = oppList & "ATR5m, "
        End If
    End If
    
    If IsNumeric(srcArr(srcRow, 82)) Then
        If InStr(1, LCase$(bias15), "bull") > 0 Then
            score = score + 1: supList = supList & "ATR15m, "
        ElseIf InStr(1, LCase$(bias15), "bear") > 0 Then
            score = score - 1: oppList = oppList & "ATR15m, "
        End If
    End If
    On Error GoTo 0


    ' --- VWAP (BR=70) ---
    If UCase$(CStr(srcArr(srcRow, 70))) = "Y" Then score = score + 1: supList = supList & "VWAP, "

    ' --- Vol Spike (BS=71) ---
    If UCase$(CStr(srcArr(srcRow, 71))) = "Y" Then score = score + 1: supList = supList & "VolSpike, "

    ' --- MACD (Trades AF:AI = cols 31–34) ---
    If Not IsEmpty(tradeArr) And tradeRow > 0 Then
        Dim tfArr As Variant: tfArr = Array("5m", "15m", "1h", "1d")
        Dim j As Long, v As String, tf As String, supports As Boolean, refBias As Variant, status As String

        For j = 0 To 3
            v = UCase$(CStr(tradeArr(tradeRow, 31 + j)))  ' AF=31
            tf = tfArr(j)
            status = "Neutral"

            If v = "OVER" Or v = "UNDER" Then
                ' Pick a reference bias (prefer 15m, then 1h, else 5m)
                refBias = bias15
                If InStr(LCase$(CStr(refBias)), "neutral") > 0 Then refBias = bias1h
                If InStr(LCase$(CStr(refBias)), "neutral") > 0 Then refBias = bias5

                refBias = LCase$(CStr(refBias))
                supports = False

                If InStr(refBias, "bull") > 0 Then
                    supports = (v = "OVER")
                ElseIf InStr(refBias, "bear") > 0 Then
                    supports = (v = "UNDER")
                End If

'                If supports Then
'                    score = score + 1
'                    supList = supList & "MACD " & tf & " (aligned), "
'                    status = "Aligned"
'                Else
'                    score = score - 1
'                    oppList = oppList & "MACD " & tf & " (divergent), "
'                    status = "Divergent"
'                End If
                
                If supports Then
                    score = score + 2
                    supList = supList & "MACD " & tf & " (aligned), "
                    status = "Aligned"
                Else
                    score = score - 2
                    oppList = oppList & "MACD " & tf & " (divergent), "
                    status = "Divergent"
                End If


            End If

            macdList = macdList & "MACD " & tf & "=" & status & "; "
        Next j
    End If

    ' --- Clean up lists ---
    If Len(supList) > 2 Then supList = Left$(supList, Len(supList) - 2)
    If Len(oppList) > 2 Then oppList = Left$(oppList, Len(oppList) - 2)
    If Len(macdList) > 2 Then macdList = Left$(macdList, Len(macdList) - 2)

    ' --- EMA alignment bonus/penalty ---
    If bias5 = "Bullish" And bias15 = "Bullish" And bias1h = "Bullish" Then
        score = score + 2
        supList = supList & "EMA Align Bullish, "
    ElseIf bias5 = "Bearish" And bias15 = "Bearish" And bias1h = "Bearish" Then
        score = score - 2
        oppList = oppList & "EMA Align Bearish, "
    End If


    ' --- Narrative ---
    narr = "Bias " & CStr(bias5) & "/" & CStr(bias15) & "/" & CStr(bias1h)
    narr = narr & " | MACD: [" & macdList & "]"
    narr = narr & " | Support: [" & supList & "]"
    If Len(oppList) > 0 Then narr = narr & " | Oppose: [" & oppList & "]"
    narr = narr & " | Score=" & CStr(score)

    Dim res(1 To 2) As Variant
    res(1) = score
    res(2) = narr
    NF_ScoreRow = res
End Function

'==================================================
' Helper to evaluate MACD per timeframe
'==================================================
Private Sub NF_AddMACDSupport(tf As String, val As String, bias As String, _
                               ByRef sup As Long, ByRef opp As Long, _
                               ByRef supList As String, ByRef oppList As String)

    Dim isSupport As Boolean, meaning As String
    Dim v As String: v = Trim(UCase$(CStr(val)))
    
    ' If MACD value is not "OVER" or "UNDER", treat as neutral — skip it.
    If v <> "OVER" And v <> "UNDER" Then Exit Sub
    
    ' Decide if this TF supports the bias
    If bias = "Bullish" Then
        isSupport = (v = "OVER")
    ElseIf bias = "Bearish" Then
        isSupport = (v = "UNDER")
    Else
        ' If bias neutral, treat MACD as neutral
        Exit Sub
    End If
    
    ' Map TF to meaning
    Select Case LCase$(tf)
        Case "5m": meaning = "short-term"
        Case "15m": meaning = "intra-day"
        Case "1h": meaning = "hourly trend"
        Case "1d": meaning = "long-term"
        Case Else: meaning = "other"
    End Select
    
    ' Append to lists
    If isSupport Then
        sup = sup + 1
        If supList <> "" Then supList = supList & "; "
        supList = supList & tf & " supportive (" & meaning & " aligned)"
    Else
        opp = opp + 1
        If oppList <> "" Then oppList = oppList & "; "
        oppList = oppList & tf & " opposing (" & meaning & " divergence)"
    End If
End Sub
'=============================================================
' Get bias for a single timeframe ("5m","15m","1h") from srcArr
' Looks up EMA10_xx and EMA30_xx by header name; returns Neutral if missing
' srcArr is A1:DDlastRow with row 1 = headers; rowIdx is the *sheet* row index
'=============================================================
Public Function NF_GetBiasForTF(ws As Worksheet, srcArr As Variant, rowIdx As Long, tf As String) As String
    Dim hdr As Object: Set hdr = BuildHeaderMap(ws)
    Dim cFast As Long, cSlow As Long

    Select Case LCase$(tf)
        Case "5m"
            cFast = NF_ColByHdr(hdr, "EMA10_5m")
            cSlow = NF_ColByHdr(hdr, "EMA30_5m")
        Case "15m"
            cFast = NF_ColByHdr(hdr, "EMA10_15m")
            cSlow = NF_ColByHdr(hdr, "EMA30_15m")
        Case "1h", "60m"
            cFast = NF_ColByHdr(hdr, "EMA10_1h")
            cSlow = NF_ColByHdr(hdr, "EMA30_1h")
        Case Else
            NF_GetBiasForTF = "Neutral"
            Exit Function
    End Select

    If cFast = 0 Or cSlow = 0 Then
        NF_GetBiasForTF = "Neutral"
    Else
        ' srcArr uses sheet coordinates because it was loaded from A1:DD...
        NF_GetBiasForTF = NF_BiasFromEMAs(srcArr(rowIdx, cFast), srcArr(rowIdx, cSlow))
    End If
End Function

Private Function NF_EvaluateMACD(wsTrades As Worksheet, tradeRow As Long, _
                                 ByRef supportive As Collection, ByRef opposing As Collection) As String
    Dim tf As Variant, col As Long, v As String
    Dim summary As String
    
    If tradeRow = 0 Then Exit Function
    
    tf = Array("5m", "15m", "1h", "1d")
    
    For col = 0 To 3
        v = wsTrades.cells(tradeRow, 31 + col).value  ' AF:AI
        If LCase(v) Like "*over*" Then
            supportive.Add "MACD " & tf(col)
        ElseIf LCase(v) Like "*under*" Then
            opposing.Add "MACD " & tf(col)
        End If
    Next col
    
    summary = "MACD checked (" & Join(tf, "/") & ")"
    NF_EvaluateMACD = summary
End Function


'-----------------------------------------------------------------
' Finds the row index of a symbol in the Trades array
' tradeArr = 2D array from Trades sheet (cols A:AI)
' symbol = string to search
' Returns row index (1-based within tradeArr), or 0 if not found
'-----------------------------------------------------------------
Public Function NF_FindTradeRow(tradeArr As Variant, vSymbol As Variant) As Long
    Dim r As Long, sym As String, v As Variant
    
    NF_FindTradeRow = 0
    If IsEmpty(tradeArr) Then Exit Function
    
    If IsError(vSymbol) Then Exit Function
    If Len(Trim$(CStr(vSymbol))) = 0 Then Exit Function
    sym = Trim$(CStr(vSymbol))
    
    For r = LBound(tradeArr, 1) To UBound(tradeArr, 1)
        v = tradeArr(r, 1) ' symbol col in Trades/New_Trades
        If Not IsError(v) Then
            If Trim$(CStr(v)) = sym Then
                NF_FindTradeRow = r
                Exit Function
            End If
        End If
    Next r
End Function


Public Function NF_LoadTradeArray(Optional ByRef whichSheet As String = "") As Variant
    Dim wsN As Worksheet, wsT As Worksheet
    Dim lrN As Long, lrT As Long
    
    On Error Resume Next
    Set wsN = ThisWorkbook.Sheets("New_Trades")
    Set wsT = ThisWorkbook.Sheets("Trades")
    On Error GoTo 0
    
    whichSheet = ""
    
    If Not wsN Is Nothing Then
        lrN = wsN.cells(wsN.Rows.count, "A").End(xlUp).row
        If lrN >= 2 Then
            NF_LoadTradeArray = wsN.Range("A2:AI" & lrN).value ' row 2 down
            whichSheet = "New_Trades"
            Exit Function
        End If
    End If
    
    If Not wsT Is Nothing Then
        lrT = wsT.cells(wsT.Rows.count, "A").End(xlUp).row
        If lrT >= 3 Then
            NF_LoadTradeArray = wsT.Range("A3:AI" & lrT).value ' row 3 down
            whichSheet = "Trades"
            Exit Function
        End If
    End If
    
    ' If neither has data, return Empty
    NF_LoadTradeArray = Empty
End Function


'=============================================================
' Convenience: get all three biases via NF_GetBiasForTF
'=============================================================
Public Sub NF_GetBiasTripletFromArr(ws As Worksheet, srcArr As Variant, rowIdx As Long, _
                                    ByRef bias5 As Variant, ByRef bias15 As Variant, ByRef bias1h As Variant)
    Dim ema10 As Double, ema30 As Double
    
    ' --- 5m bias (AS=col 45, AT=col 46)
    If IsNumeric(srcArr(rowIdx, 45)) And IsNumeric(srcArr(rowIdx, 46)) Then
        ema10 = srcArr(rowIdx, 45)
        ema30 = srcArr(rowIdx, 46)
        If ema10 > ema30 Then
            bias5 = "Bullish"
        ElseIf ema10 < ema30 Then
            bias5 = "Bearish"
        Else
            bias5 = "Neutral"
        End If
    Else
        bias5 = "Neutral"
    End If
    
    ' --- 15m bias (AX=col 50, AY=col 51)
    If IsNumeric(srcArr(rowIdx, 50)) And IsNumeric(srcArr(rowIdx, 51)) Then
        ema10 = srcArr(rowIdx, 50)
        ema30 = srcArr(rowIdx, 51)
        If ema10 > ema30 Then
            bias15 = "Bullish"
        ElseIf ema10 < ema30 Then
            bias15 = "Bearish"
        Else
            bias15 = "Neutral"
        End If
    Else
        bias15 = "Neutral"
    End If
    
    ' --- 1h bias (BC=col 55, BD=col 56)
    If IsNumeric(srcArr(rowIdx, 55)) And IsNumeric(srcArr(rowIdx, 56)) Then
        ema10 = srcArr(rowIdx, 55)
        ema30 = srcArr(rowIdx, 56)
        If ema10 > ema30 Then
            bias1h = "Bullish"
        ElseIf ema10 < ema30 Then
            bias1h = "Bearish"
        Else
            bias1h = "Neutral"
        End If
    Else
        bias1h = "Neutral"
    End If
End Sub


'=============================================================
' Map a header name (case-insensitive) to its column index.
' Requires your existing BuildHeaderMap(ws) to return a Scripting.Dictionary
' whose keys are UCase(headerText) and values are 1-based column numbers.
'=============================================================
Private Function NF_ColByHdr(hdr As Object, headerText As String) As Long
    Dim k As String
    k = UCase$(headerText)
    If Not hdr Is Nothing Then
        If hdr.Exists(k) Then
            NF_ColByHdr = hdr(k)
            Exit Function
        End If
    End If
    NF_ColByHdr = 0
End Function



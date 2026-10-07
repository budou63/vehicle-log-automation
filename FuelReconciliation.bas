Attribute VB_Name = "FuelReconciliation"
Option Explicit

Private Const GAS_HEADER As String = "ガソリン給油の有無 [給油の場合はその他に数値入力（?）]"
Private Const OUTPUT_MARKER As String = "【給油照合管理】"

Public Sub 給油記録を照合()
    Dim ticketWS As Worksheet, logWS As Worksheet
    Dim tickets As Collection, logs As Collection
    Dim statuses As Object, matches As Object, usedTickets As Object, usedLogs As Object
    Dim reservedSameDayTickets As Object, reservedSameDayLogs As Object
    Dim gasCol As Long

    On Error GoTo Failed
    Set ticketWS = ThisWorkbook.Worksheets("給油記録")
    Set logWS = ThisWorkbook.Worksheets("記録")
    gasCol = HeaderColumn(logWS, GAS_HEADER)
    If gasCol = 0 Then
        MsgBox "記録シートのガソリン給油列が見つかりません。", vbExclamation
        Exit Sub
    End If
    If Not PrepareTicketOutput(ticketWS) Then Exit Sub

    Set statuses = CreateObject("Scripting.Dictionary")
    Set tickets = ReadTickets(ticketWS, statuses)
    Set logs = ReadLogs(logWS, gasCol)
    Set matches = CreateObject("Scripting.Dictionary")
    Set usedTickets = CreateObject("Scripting.Dictionary")
    Set usedLogs = CreateObject("Scripting.Dictionary")
    Set reservedSameDayTickets = CreateObject("Scripting.Dictionary")
    Set reservedSameDayLogs = CreateObject("Scripting.Dictionary")

    ResolveStage tickets, logs, usedTickets, usedLogs, matches, reservedSameDayTickets, reservedSameDayLogs, True
    ReserveAmbiguousSameDay tickets, logs, usedTickets, usedLogs, reservedSameDayTickets, reservedSameDayLogs
    ResolveStage tickets, logs, usedTickets, usedLogs, matches, reservedSameDayTickets, reservedSameDayLogs, False
    SetRemainingStatuses tickets, logs, statuses, usedTickets, usedLogs, matches, reservedSameDayTickets, reservedSameDayLogs
    If Not ResetManagedTicketOutput(ticketWS) Then Exit Sub
    WriteTicketResults ticketWS, tickets, statuses, matches

    MsgBox "給油記録の照合が完了しました。", vbInformation
    Exit Sub
Failed:
    MsgBox "給油照合を中止しました。" & Err.Description, vbExclamation
End Sub

Private Function PrepareTicketOutput(ws As Worksheet) As Boolean
    Dim r As Long, lastRow As Long, activeRow As Boolean, hasOutput As Boolean

    lastRow = LastTicketInputRow(ws)
    If Not HeadersAreSafe(ws) Then
        MsgBox "給油記録シートのD:H見出しに管理外の値があります。上書きしないため中止しました。", vbExclamation
        Exit Function
    End If

    For r = 2 To lastRow
        activeRow = TicketRowHasInput(ws, r)
        hasOutput = Application.WorksheetFunction.CountA(ws.Range("D" & r & ":H" & r)) > 0
        If activeRow And hasOutput And Trim$(CStr(ws.Cells(r, "H").Value)) <> OUTPUT_MARKER Then
            MsgBox "給油記録の入力行に、このマクロ以外のD:H既存値があります。上書きしないため中止しました。", vbExclamation
            Exit Function
        End If
    Next r
    PrepareTicketOutput = True
End Function

Private Function ResetManagedTicketOutput(ws As Worksheet) As Boolean
    Dim r As Long, lastRow As Long
    lastRow = LastTicketSheetRow(ws)
    For r = 2 To lastRow
        If Not IsError(ws.Cells(r, "H").Value) Then
            If Trim$(CStr(ws.Cells(r, "H").Value)) = OUTPUT_MARKER Then ws.Range("D" & r & ":H" & r).ClearContents
        End If
    Next r
    ws.Range("D1:H1").Value = Array("照合結果", "実給油日", "LoGo回答番号", "備考", "管理マーカー")
    ResetManagedTicketOutput = True
End Function

Private Function LastTicketSheetRow(ws As Worksheet) As Long
    Dim c As Long, candidate As Long
    For c = 1 To 8
        candidate = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If candidate > LastTicketSheetRow Then LastTicketSheetRow = candidate
    Next c
End Function

Private Function HeadersAreSafe(ws As Worksheet) As Boolean
    Dim expected As Variant, c As Long, currentValue As String
    expected = Array("照合結果", "実給油日", "LoGo回答番号", "備考", "管理マーカー")
    For c = 0 To 4
        currentValue = Trim$(CStr(ws.Cells(1, "D").Offset(0, c).Value))
        If Len(currentValue) > 0 And currentValue <> expected(c) Then Exit Function
    Next c
    HeadersAreSafe = True
End Function

Private Function LastTicketInputRow(ws As Worksheet) As Long
    Dim c As Long, candidate As Long
    For c = 1 To 3
        candidate = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If candidate > LastTicketInputRow Then LastTicketInputRow = candidate
    Next c
End Function

Private Function TicketRowHasInput(ws As Worksheet, ByVal rowNo As Long) As Boolean
    TicketRowHasInput = Application.WorksheetFunction.CountA(ws.Range("A" & rowNo & ":C" & rowNo)) > 0
End Function

Private Function ReadTickets(ws As Worksheet, statuses As Object) As Collection
    Dim result As New Collection, r As Long, lastRow As Long, amount As Double, car As String
    lastRow = LastTicketInputRow(ws)
    For r = 2 To lastRow
        If TicketRowHasInput(ws, r) Then
            If IsError(ws.Cells(r, "A").Value) Or IsError(ws.Cells(r, "B").Value) Or IsError(ws.Cells(r, "C").Value) Then
                statuses.Add CStr(r), "入力値要確認"
            Else
                car = Trim$(CStr(ws.Cells(r, "B").Value))
                If Not IsDate(ws.Cells(r, "A").Value) Or Len(car) = 0 Or Not PositiveAmount(ws.Cells(r, "C").Value, amount) Then
                    statuses.Add CStr(r), "入力値要確認"
                Else
                    result.Add Array(CStr(r), CDate(ws.Cells(r, "A").Value), car, amount)
                End If
            End If
        End If
    Next r
    Set ReadTickets = result
End Function

Private Function ReadLogs(ws As Worksheet, ByVal gasCol As Long) As Collection
    Dim answerCounts As Object, result As New Collection
    Dim r As Long, lastRow As Long, answer As String, car As String, amount As Double

    Set answerCounts = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        If Not IsError(ws.Cells(r, "A").Value) Then
            answer = Trim$(CStr(ws.Cells(r, "A").Value))
            If Len(answer) > 0 Then
                If answerCounts.Exists(answer) Then
                    answerCounts(answer) = CLng(answerCounts(answer)) + 1
                Else
                    answerCounts.Add answer, 1
                End If
            End If
        End If
    Next r

    For r = 2 To lastRow
        If Not IsError(ws.Cells(r, "A").Value) And Not IsError(ws.Cells(r, "J").Value) And Not IsError(ws.Cells(r, "K").Value) And Not IsError(ws.Cells(r, gasCol).Value) Then
            answer = Trim$(CStr(ws.Cells(r, "A").Value))
            car = Trim$(CStr(ws.Cells(r, "J").Value))
            If Len(answer) > 0 And Len(car) > 0 Then
                If IsDate(ws.Cells(r, "K").Value) And PositiveAmount(ws.Cells(r, gasCol).Value, amount) Then
                    If CLng(answerCounts(answer)) = 1 Then result.Add Array(answer, CDate(ws.Cells(r, "K").Value), car, amount)
                End If
            End If
        End If
    Next r
    Set ReadLogs = result
End Function

Private Sub ResolveStage(tickets As Collection, logs As Collection, usedTickets As Object, usedLogs As Object, matches As Object, reservedTickets As Object, reservedLogs As Object, ByVal sameDay As Boolean)
    Dim proposals As New Collection, ticket As Variant, logItem As Variant, ticketKey As String, answer As String
    Dim changed As Boolean, proposal As Variant

    Do
        Set proposals = New Collection
        For Each ticket In tickets
            ticketKey = CStr(ticket(0))
            If Not usedTickets.Exists(ticketKey) And (sameDay Or Not reservedTickets.Exists(ticketKey)) Then
                If TicketCandidateCount(ticket, logs, usedLogs, reservedLogs, sameDay) = 1 Then
                    logItem = OnlyTicketCandidate(ticket, logs, usedLogs, reservedLogs, sameDay)
                    answer = CStr(logItem(0))
                    If LogCandidateCount(logItem, tickets, usedTickets, reservedTickets, sameDay) = 1 Then proposals.Add Array(ticketKey, answer, CDate(logItem(1)))
                End If
            End If
        Next ticket

        changed = (proposals.Count > 0)
        For Each proposal In proposals
            ticketKey = CStr(proposal(0)): answer = CStr(proposal(1))
            If Not usedTickets.Exists(ticketKey) And Not usedLogs.Exists(answer) Then
                usedTickets.Add ticketKey, True
                usedLogs.Add answer, True
                matches.Add ticketKey, Array(answer, sameDay, CDate(proposal(2)))
            End If
        Next proposal
    Loop While changed
End Sub

Private Function TicketCandidateCount(ticket As Variant, logs As Collection, usedLogs As Object, reservedLogs As Object, ByVal sameDay As Boolean) As Long
    Dim logItem As Variant
    For Each logItem In logs
        If Not usedLogs.Exists(CStr(logItem(0))) And (sameDay Or Not reservedLogs.Exists(CStr(logItem(0)))) Then
            If IsCandidate(ticket, logItem, sameDay) Then TicketCandidateCount = TicketCandidateCount + 1
        End If
    Next logItem
End Function

Private Function OnlyTicketCandidate(ticket As Variant, logs As Collection, usedLogs As Object, reservedLogs As Object, ByVal sameDay As Boolean) As Variant
    Dim logItem As Variant
    For Each logItem In logs
        If Not usedLogs.Exists(CStr(logItem(0))) And (sameDay Or Not reservedLogs.Exists(CStr(logItem(0)))) Then
            If IsCandidate(ticket, logItem, sameDay) Then
                OnlyTicketCandidate = logItem
                Exit Function
            End If
        End If
    Next logItem
End Function

Private Function LogCandidateCount(logItem As Variant, tickets As Collection, usedTickets As Object, reservedTickets As Object, ByVal sameDay As Boolean) As Long
    Dim ticket As Variant
    For Each ticket In tickets
        If Not usedTickets.Exists(CStr(ticket(0))) And (sameDay Or Not reservedTickets.Exists(CStr(ticket(0)))) Then
            If IsCandidate(ticket, logItem, sameDay) Then LogCandidateCount = LogCandidateCount + 1
        End If
    Next ticket
End Function

Private Function IsCandidate(ticket As Variant, logItem As Variant, ByVal sameDay As Boolean) As Boolean
    If CStr(ticket(2)) <> CStr(logItem(2)) Then Exit Function
    If CDbl(ticket(3)) <> CDbl(logItem(3)) Then Exit Function
    If sameDay Then
        IsCandidate = (CDate(ticket(1)) = CDate(logItem(1)))
    Else
        IsCandidate = (CDate(logItem(1)) > CDate(ticket(1)))
    End If
End Function

Private Sub ReserveAmbiguousSameDay(tickets As Collection, logs As Collection, usedTickets As Object, usedLogs As Object, reservedTickets As Object, reservedLogs As Object)
    Dim ticket As Variant, logItem As Variant, ticketKey As String, answer As String
    For Each ticket In tickets
        ticketKey = CStr(ticket(0))
        If Not usedTickets.Exists(ticketKey) Then
            For Each logItem In logs
                answer = CStr(logItem(0))
                If Not usedLogs.Exists(answer) Then
                    If IsCandidate(ticket, logItem, True) Then
                        If Not reservedTickets.Exists(ticketKey) Then reservedTickets.Add ticketKey, True
                        If Not reservedLogs.Exists(answer) Then reservedLogs.Add answer, True
                    End If
                End If
            Next logItem
        End If
    Next ticket
End Sub

Private Sub SetRemainingStatuses(tickets As Collection, logs As Collection, statuses As Object, usedTickets As Object, usedLogs As Object, matches As Object, reservedTickets As Object, reservedLogs As Object)
    Dim ticket As Variant, ticketKey As String
    For Each ticket In tickets
        ticketKey = CStr(ticket(0))
        If Not matches.Exists(ticketKey) Then
            If reservedTickets.Exists(ticketKey) Then
                statuses.Add ticketKey, "候補複数・要確認"
            ElseIf TicketCandidateCount(ticket, logs, usedLogs, reservedLogs, False) > 0 Or TicketCandidateCount(ticket, logs, usedLogs, reservedLogs, True) > 0 Then
                statuses.Add ticketKey, "候補複数・要確認"
            Else
                statuses.Add ticketKey, "未照合"
            End If
        End If
    Next ticket
End Sub

Private Sub WriteTicketResults(ws As Worksheet, tickets As Collection, statuses As Object, matches As Object)
    Dim ticket As Variant, ticketKey As String, matchInfo As Variant, statusText As String
    Dim allRows As Object, key As Variant
    Set allRows = CreateObject("Scripting.Dictionary")
    For Each ticket In tickets: allRows.Add CStr(ticket(0)), True: Next ticket
    For Each key In statuses.Keys: If Not allRows.Exists(CStr(key)) Then allRows.Add CStr(key), True

    For Each key In allRows.Keys
        ticketKey = CStr(key)
        If matches.Exists(ticketKey) Then
            matchInfo = matches(ticketKey)
            statusText = IIf(CBool(matchInfo(1)), "同日一致", "後日一致")
        Else
            statusText = statuses(ticketKey)
        End If
        ws.Cells(CLng(ticketKey), "D").Value = statusText
        ws.Cells(CLng(ticketKey), "H").Value = OUTPUT_MARKER
        If matches.Exists(ticketKey) Then
            ws.Cells(CLng(ticketKey), "E").Value = CDate(matchInfo(2))
            ws.Cells(CLng(ticketKey), "F").Value = CStr(matchInfo(0))
            ws.Cells(CLng(ticketKey), "G").Value = IIf(CBool(matchInfo(1)), "", "発行日後の給油")
        ElseIf statusText = "候補複数・要確認" Then
            ws.Cells(CLng(ticketKey), "G").Value = "同条件の候補が複数あります"
        ElseIf statusText = "入力値要確認" Then
            ws.Cells(CLng(ticketKey), "G").Value = "日付・車両番号・給油量を確認してください"
        End If
    Next key
End Sub

Private Function HeaderColumn(ws As Worksheet, text As String) As Long
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If Trim$(CStr(ws.Cells(1, c).Value)) = text Then HeaderColumn = c: Exit Function
    Next c
End Function

Private Function PositiveAmount(v As Variant, ByRef n As Double) As Boolean
    If IsError(v) Or Not IsNumeric(v) Then Exit Function
    n = CDbl(v): PositiveAmount = (n > 0#)
End Function

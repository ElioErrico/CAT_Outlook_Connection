Attribute VB_Name = "modMarkdownHelpers"
'==========================
' Modulo: modMarkdownHelpers
'==========================
Option Explicit

' --------- API PUBBLICHE (usate dal Core) ---------

Public Function BuildMessageFromSelection(ByVal rng As Word.Range) As String
    Dim out As String
    Dim curStart As Long
    Dim nextT As Word.Table
    Dim txtRng As Word.Range

    If rng.Tables.Count = 0 Then
        out = SanitizeForApi(NormalizeParagraphs(rng.text))
        BuildMessageFromSelection = out
        Exit Function
    End If

    curStart = rng.Start
    Do
        Set nextT = FindNextTableInRange(rng, curStart)
        If nextT Is Nothing Then
            If curStart < rng.End Then
                Set txtRng = rng.Document.Range(curStart, rng.End)
                out = out & SanitizeForApi(NormalizeParagraphs(txtRng.text))
            End If
            Exit Do
        End If

        ' testo prima della tabella
        If nextT.Range.Start > curStart Then
            Set txtRng = rng.Document.Range(curStart, nextT.Range.Start)
            out = out & SanitizeForApi(NormalizeParagraphs(txtRng.text))
            If Len(out) > 0 And Right$(out, 1) <> vbLf Then out = out & vbLf
        End If

        ' tabella in Markdown (separata da righe vuote)
        If Len(out) > 0 And Right$(out, 1) <> vbLf Then out = out & vbLf
        out = out & TableToMarkdown(nextT) & vbLf
        curStart = nextT.Range.End
        If curStart < rng.End Then out = out & vbLf
    Loop

    BuildMessageFromSelection = Trim$(out)
End Function

Public Sub InsertMarkdownInline(ByVal sel As Word.Selection, _
                                ByVal s As String, _
                                Optional ByVal defaultBold As Boolean = False, _
                                Optional ByVal defaultItalic As Boolean = False)
    Dim i As Long
    Dim buf As String
    Dim boldState As Boolean, italicState As Boolean

    s = NormalizeToLf(s)
    boldState = False: italicState = False
    i = 1

    Do While i <= Len(s)
        If Mid$(s, i, 1) = "\" Then
            If i < Len(s) Then
                buf = buf & Mid$(s, i + 1, 1)
                i = i + 2
            Else
                i = i + 1
            End If

        ElseIf i <= Len(s) - 2 And (Mid$(s, i, 3) = "***" Or Mid$(s, i, 3) = "___") Then
            FlushRun sel, buf, defaultBold Or boldState, defaultItalic Or italicState
            boldState = Not boldState
            italicState = Not italicState
            i = i + 3

        ElseIf i <= Len(s) - 1 And (Mid$(s, i, 2) = "**" Or Mid$(s, i, 2) = "__") Then
            FlushRun sel, buf, defaultBold Or boldState, defaultItalic Or italicState
            boldState = Not boldState
            i = i + 2

        ElseIf Mid$(s, i, 1) = "*" Or Mid$(s, i, 1) = "_" Then
            FlushRun sel, buf, defaultBold Or boldState, defaultItalic Or italicState
            italicState = Not italicState
            i = i + 1

        ElseIf Mid$(s, i, 1) = vbLf Then
            FlushRun sel, buf, defaultBold Or boldState, defaultItalic Or italicState
            sel.TypeParagraph
            i = i + 1

        Else
            buf = buf & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop

    FlushRun sel, buf, defaultBold Or boldState, defaultItalic Or italicState
End Sub

Public Sub EnsureParagraphBeforeInsertion(ByVal sel As Word.Selection)
    Dim lastChar As String
    lastChar = ""
    On Error Resume Next
    lastChar = Right$(sel.Range.text, 1)
    On Error GoTo 0
    If lastChar <> vbCr Then sel.TypeParagraph
End Sub

Public Function GetSelectedMarkdownTableText(ByVal rng As Word.Range) As String
    Dim tx As String, lines() As String, i As Long, out As String
    tx = rng.text
    tx = Replace(tx, vbCrLf, vbLf)
    tx = Replace(tx, vbCr, vbLf)

    lines = Split(tx, vbLf)
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(lines(i))
        If ln <> "```" And LCase$(ln) <> "```markdown" And ln <> "" Then
            out = out & ln & vbLf
        End If
    Next i
    If Right$(out, 1) = vbLf Then out = Left$(out, Len(out) - 1)
    GetSelectedMarkdownTableText = out
End Function

Public Sub ConvertMarkdownToWord(ByVal markdown As String, _
                                 ByVal targetRange As Word.Range, _
                                 ByVal sel As Word.Selection)
    ' Wrapper comodo
    CreateAndInsertWordTableFromMarkdown markdown, targetRange, sel
End Sub

Public Sub CreateAndInsertWordTableFromMarkdown(ByVal markdown As String, _
                                                ByVal targetRange As Word.Range, _
                                                ByVal sel As Word.Selection)
    Dim rawLines() As String, lines() As String
    Dim i As Long, rowCount As Long, colCount As Long
    Dim sepLineIndex As Long, actualRow As Long
    Dim rowVals() As String
    Dim wordTable As Word.Table
    Dim r As Word.Range, cellText As String
    Dim c As Long

    rawLines = Split(NormalizeToLf(RemoveFences(markdown)), vbLf)
    lines = FilterNonEmpty(rawLines)

    rowCount = UBound(lines) - LBound(lines) + 1
    If rowCount < 2 Then Err.Raise vbObjectError + 7001, , "Tabella Markdown non valida."

    colCount = CountMarkdownColumns(lines(0))
    If colCount < 1 Then Err.Raise vbObjectError + 7002, , "Impossibile determinare il numero di colonne."

    sepLineIndex = FindSeparatorIndex(lines, colCount)
    If sepLineIndex = -1 Then Err.Raise vbObjectError + 7003, , "Riga separatrice Markdown non trovata."

    targetRange.text = ""
    Set wordTable = targetRange.Document.Tables.Add(Range:=targetRange, NumRows:=(rowCount - 1), NumColumns:=colCount)

    actualRow = 1
    For i = 0 To UBound(lines)
        If i <> sepLineIndex Then
            rowVals = SplitMarkdownRow(lines(i), colCount)
            For c = 1 To colCount
                Set r = wordTable.Cell(actualRow, c).Range
                r.End = r.End - 1
                r.text = ""
                r.Select               ' Sposto la Selection su questa cella
                cellText = rowVals(c - 1)
                InsertMarkdownInline sel, cellText, (actualRow = 1), False
            Next c
            actualRow = actualRow + 1
        End If
    Next i

    ApplyNiceTableFormatting wordTable

    ' Posiziona il caret dopo la tabella
    Dim afterTbl As Word.Range
    Set afterTbl = wordTable.Range
    afterTbl.Collapse wdCollapseEnd
    sel.SetRange Start:=afterTbl.End, End:=afterTbl.End
End Sub

Public Function ExtractFirstMarkdownTableBlock(ByVal src As String, _
                                               ByRef preText As String, _
                                               ByRef tableBlock As String, _
                                               ByRef postText As String) As Boolean
    Dim lines() As String, i As Long, j As Long, k As Long
    Dim hdr As String, sep As String, colCount As Long
    Dim startIdx As Long, sepIdx As Long, endIdx As Long

    lines = Split(src, vbLf)
    startIdx = -1: sepIdx = -1: endIdx = -1

    For i = LBound(lines) To UBound(lines) - 1
        hdr = Trim$(lines(i))
        If IsFenceLine(hdr) Or hdr = "" Then GoTo NextI
        If CountPipes(hdr) >= 2 Then
            For j = i + 1 To UBound(lines)
                sep = Trim$(lines(j))
                If sep <> "" And Not IsFenceLine(sep) Then
                    colCount = CountMarkdownColumns(hdr)
                    If colCount > 0 And (IsSeparatorLine(sep, colCount) Or IsSeparatorLike(sep)) Then
                        startIdx = i: sepIdx = j
                    End If
                    Exit For
                End If
            Next j
            If startIdx <> -1 Then Exit For
        End If
NextI:
    Next i

    If startIdx = -1 Then
        preText = src: tableBlock = "": postText = ""
        ExtractFirstMarkdownTableBlock = False
        Exit Function
    End If

    endIdx = sepIdx
    For k = sepIdx + 1 To UBound(lines)
        If Trim$(lines(k)) = "" Or IsFenceLine(Trim$(lines(k))) Then Exit For
        If InStr(1, lines(k), "|") = 0 Then Exit For
        endIdx = k
    Next k

    preText = JoinSubArray(lines, LBound(lines), startIdx - 1, vbLf)
    tableBlock = JoinSubArray(lines, startIdx, endIdx, vbLf)
    postText = ""
    If endIdx + 1 <= UBound(lines) Then
        postText = JoinSubArray(lines, endIdx + 1, UBound(lines), vbLf)
    End If

    ExtractFirstMarkdownTableBlock = True
End Function

' --------- HELPER INTERNI ---------

Public Sub FlushRun(ByVal sel As Word.Selection, ByRef buf As String, ByVal isBold As Boolean, ByVal isItalic As Boolean)
    Dim startPos As Long
    Dim r As Word.Range
    If Len(buf) = 0 Then Exit Sub

    startPos = sel.Range.Start
    sel.TypeText buf
    Set r = sel.Document.Range(Start:=startPos, End:=sel.Range.Start)
    r.Font.Bold = isBold
    r.Font.Italic = isItalic

    buf = vbNullString
End Sub

Public Function NormalizeToLf(ByVal tx As String) As String
    tx = Replace(tx, vbCrLf, vbLf)
    tx = Replace(tx, vbCr, vbLf)
    NormalizeToLf = tx
End Function

Public Function RemoveFences(ByVal tx As String) As String
    Dim out As String, lines() As String, i As Long, ln As String
    lines = Split(NormalizeToLf(tx), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Not IsFenceLine(ln) Then out = out & lines(i) & vbLf
    Next i
    If Right$(out, 1) = vbLf Then out = Left$(out, Len(out) - 1)
    RemoveFences = out
End Function

Public Function FilterNonEmpty(arr() As String) As String()
    Dim tmp() As String, i As Long, v As String, k As Long
    ReDim tmp(0 To 0)
    For i = LBound(arr) To UBound(arr)
        v = Trim$(arr(i))
        If v <> "" Then
            If tmp(0) = "" Then
                tmp(0) = v
            Else
                ReDim Preserve tmp(0 To UBound(tmp) + 1)
                tmp(UBound(tmp)) = v
            End If
        End If
    Next i
    FilterNonEmpty = tmp
End Function

Public Function FindSeparatorIndex(lines() As String, ByVal colCount As Long) As Long
    Dim i As Long
    For i = 1 To UBound(lines)
        If IsSeparatorLine(lines(i), colCount) Or IsSeparatorLike(lines(i)) Then
            FindSeparatorIndex = i
            Exit Function
        End If
    Next i
    FindSeparatorIndex = -1
End Function

Public Function CountMarkdownColumns(ByVal rowText As String) As Long
    Dim parts() As String, i As Long, cnt As Long
    parts = Split(rowText, "|")
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then cnt = cnt + 1
    Next i
    CountMarkdownColumns = cnt
End Function

Public Function IsSeparatorLine(ByVal rowText As String, ByVal colCount As Long) As Boolean
    Dim vals() As String, i As Long, tok As String, seen As Long
    vals = Split(rowText, "|")
    For i = LBound(vals) To UBound(vals)
        tok = Trim$(vals(i))
        If tok <> "" Then
            seen = seen + 1
            If Not IsSeparatorToken(tok) Then
                IsSeparatorLine = False
                Exit Function
            End If
        End If
    Next i
    IsSeparatorLine = (seen = colCount)
End Function

Public Function IsSeparatorLike(ByVal rowText As String) As Boolean
    IsSeparatorLike = (InStr(1, rowText, "---") > 0)
End Function

Public Function IsSeparatorToken(ByVal tok As String) As Boolean
    Dim c As Long, ch As String
    Dim hyphenCount As Long

    hyphenCount = Len(Replace$(tok, ":", ""))
    hyphenCount = hyphenCount - Len(Replace$(Replace$(tok, ":", ""), "-", "")) ' num '-' min. 3

    If hyphenCount >= 3 Then
        For c = 1 To Len(tok)
            ch = Mid$(tok, c, 1)
            If ch <> "-" And ch <> ":" Then
                IsSeparatorToken = False
                Exit Function
            End If
        Next c
        IsSeparatorToken = True
    Else
        IsSeparatorToken = False
    End If
End Function

Public Function SplitMarkdownRow(ByVal rowText As String, ByVal colCount As Long) As String()
    Dim parts() As String, tmp() As String
    Dim i As Long, v As String, k As Long

    parts = Split(rowText, "|")
    ReDim tmp(0 To colCount - 1)
    k = 0
    For i = LBound(parts) To UBound(parts)
        v = Trim$(parts(i))
        If v <> "" Then
            If k <= UBound(tmp) Then
                tmp(k) = v
                k = k + 1
            End If
        End If
    Next i
    For i = k To UBound(tmp)
        tmp(i) = ""
    Next i
    SplitMarkdownRow = tmp
End Function

Public Sub ApplyNiceTableFormatting(ByVal tbl As Word.Table)
    Dim i As Long, j As Long
    On Error Resume Next
    tbl.Style = "Tabella griglia"
    If Err.Number <> 0 Then
        Err.Clear: tbl.Style = "Table Grid"
    End If
    On Error GoTo 0

    With tbl
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Range.Font.Name = "Calibri"
        .Range.Font.Size = 11
        .Borders.OutsideLineStyle = wdLineStyleSingle
        .Borders.InsideLineStyle = wdLineStyleSingle

        For j = 1 To .Columns.Count
            With .Cell(1, j).Range
                .Font.Bold = True
                .Shading.BackgroundPatternColor = RGB(240, 240, 240)
            End With
        Next j

        For i = 2 To .rows.Count
            If i Mod 2 = 0 Then
                For j = 1 To .Columns.Count
                    .Cell(i, j).Range.Shading.BackgroundPatternColor = RGB(249, 249, 249)
                Next j
            End If
        Next i

        .AutoFitBehavior wdAutoFitContent
    End With
End Sub

Public Function JoinSubArray(arr() As String, ByVal a As Long, ByVal b As Long, ByVal sep As String) As String
    Dim i As Long, s As String
    If b < a Or a < LBound(arr) Or b > UBound(arr) Then Exit Function
    For i = a To b
        s = s & arr(i)
        If i < b Then s = s & sep
    Next i
    JoinSubArray = s
End Function

Public Function IsFenceLine(ByVal ln As String) As Boolean
    ln = LCase$(Trim$(ln))
    IsFenceLine = (ln = "```" Or ln = "```markdown")
End Function

Public Function CountPipes(ByVal s As String) As Long
    CountPipes = (Len(s) - Len(Replace$(s, "|", "")))
End Function

Public Function TableToMarkdown(ByVal tbl As Word.Table) As String
    Dim r As Long, c As Long, cols As Long, rows As Long
    Dim line As String, sep As String, md As String

    rows = tbl.rows.Count
    cols = tbl.Columns.Count
    If cols = 0 Or rows = 0 Then Exit Function

    ' Header = prima riga
    line = ""
    For c = 1 To cols
        line = line & "| " & EscapePipes(Trim$(GetCellPlainText(tbl.Cell(1, c)))) & " "
    Next c
    line = line & "|"
    md = line & vbLf

    ' Separatore con allineamento derivato dall'header
    sep = ""
    For c = 1 To cols
        Select Case tbl.Cell(1, c).Range.ParagraphFormat.Alignment
            Case wdAlignParagraphCenter: sep = sep & "|:" & String$(3, "-") & ":"
            Case wdAlignParagraphRight:  sep = sep & "| " & String$(3, "-") & ":"
            Case Else:                   sep = sep & "| " & String$(3, "-") & " "
        End Select
        sep = sep & " "
    Next c
    sep = sep & "|"
    md = md & sep & vbLf

    ' Dati
    For r = 2 To rows
        line = ""
        For c = 1 To cols
            line = line & "| " & EscapePipes(Trim$(GetCellPlainText(tbl.Cell(r, c)))) & " "
        Next c
        line = line & "|"
        md = md & line & vbLf
    Next r

    TableToMarkdown = SanitizeForApi(md)
End Function

Public Function GetCellPlainText(ByVal cel As Word.Cell) As String
    Dim tx As String
    tx = cel.Range.text
    If Len(tx) >= 2 Then
        tx = Replace(tx, Chr(13) & Chr(7), "")
    End If
    tx = Replace(tx, vbCr, " / ")
    tx = Replace(tx, Chr(11), " / ")
    tx = Replace(tx, vbTab, " ")
    GetCellPlainText = tx
End Function

Public Function EscapePipes(ByVal s As String) As String
    EscapePipes = Replace(s, "|", "\|")
End Function

Public Function NormalizeParagraphs(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, Chr(11), vbLf) ' manual line break
    NormalizeParagraphs = s
End Function

Public Function SanitizeControlChars(ByVal s As String, Optional ByVal replaceWith As String = "") As String
    Dim i As Long, ch As String, code As Long, out As String
    If LenB(s) = 0 Then
        SanitizeControlChars = s
        Exit Function
    End If
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If (code = 9 Or code = 10 Or code = 13 Or code >= 32) Then
            out = out & ch
        Else
            If Len(replaceWith) > 0 Then out = out & replaceWith
        End If
    Next i
    SanitizeControlChars = out
End Function

Public Function SanitizeForApi(ByVal s As String) As String
    Dim t As String
    t = s
    t = NormalizeParagraphs(t)
    t = Replace(t, vbTab, "    ")
    t = Replace(t, ChrW(160), " ")
    t = RemoveCharsByCodepoints(t, Array( _
        8203, 8204, 8205, _
        8234, 8235, 8236, 8237, 8238, _
        8298, 8299, 8300, 8301, 8302, 8303 _
    ))
    t = SanitizeControlChars(t, "")
    t = TrimLineEnds(t)
    t = CollapseBlankLines(t, 2)
    SanitizeForApi = t
End Function

Public Function RemoveCharsByCodepoints(ByVal s As String, ByVal codes As Variant) As String
    Dim i As Long, ch As String, cp As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        cp = AscW(ch)
        If Not InArrayLong(cp, codes) Then out = out & ch
    Next
    RemoveCharsByCodepoints = out
End Function

Public Function InArrayLong(ByVal v As Long, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If v = CLng(arr(i)) Then InArrayLong = True: Exit Function
    Next
End Function

Public Function TrimLineEnds(ByVal s As String) As String
    Dim lines() As String, i As Long, out As String
    lines = Split(s, vbLf)
    For i = LBound(lines) To UBound(lines)
        lines(i) = RTrim$(Replace(lines(i), vbTab, " "))
    Next
    out = Join(lines, vbLf)
    TrimLineEnds = out
End Function

Public Function CollapseBlankLines(ByVal s As String, Optional ByVal maxConsecutive As Long = 2) As String
    Dim lines() As String, i As Long, emptyRun As Long, out As String
    lines = Split(s, vbLf)
    emptyRun = 0
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) = 0 Then
            emptyRun = emptyRun + 1
            If emptyRun <= maxConsecutive Then out = out & vbLf
        Else
            emptyRun = 0
            If Len(out) > 0 Then out = out & lines(i) Else out = lines(i)
            If i < UBound(lines) Then out = out & vbLf
        End If
    Next
    Do While Left$(out, 1) = vbLf: out = Mid$(out, 2): Loop
    Do While Right$(out, 1) = vbLf And Right$(out, 2) = vbLf & vbLf: out = Left$(out, Len(out) - 1): Loop
    CollapseBlankLines = out
End Function

Public Function FindNextTableInRange(ByVal rng As Word.Range, ByVal fromPos As Long) As Word.Table
    Dim t As Word.Table, best As Word.Table, bestStart As Long
    bestStart = 0
    For Each t In rng.Tables
        If (fromPos >= t.Range.Start And fromPos < t.Range.End) Then
            Set FindNextTableInRange = t
            Exit Function
        ElseIf t.Range.Start >= fromPos Then
            If best Is Nothing Or t.Range.Start < bestStart Then
                Set best = t
                bestStart = t.Range.Start
            End If
        End If
    Next t
    If Not best Is Nothing Then Set FindNextTableInRange = best
End Function

' ===== Formatting Helpers =====

' Forza il formato HTML del messaggio in composizione (se non lo è già)
Public Sub EnsureHtmlBody(ByVal insp As Outlook.Inspector)
    Dim itm As Object
    If insp Is Nothing Then Exit Sub
    On Error Resume Next
    Set itm = insp.CurrentItem
    If Not itm Is Nothing Then
        If itm.BodyFormat = olFormatPlain Or itm.BodyFormat = olFormatUnspecified Then
            itm.BodyFormat = olFormatHTML
        End If
    End If
    On Error GoTo 0
End Sub

' Applica una formattazione "grigia" SOLO al testo (escludendo il paragrafo finale)
Public Sub GrayOutRange(ByVal r As Word.Range)
    On Error Resume Next
    Dim work As Word.Range
    Set work = r.Duplicate

    ' Se l'ultimo carattere è un paragrafo, escludilo dal range da formattare
    If work.Characters.Count > 0 Then
        If AscW(Right$(work.text, 1)) = 13 Then ' vbCr / ¶
            work.MoveEnd wdCharacter, -1
        End If
    End If

    ' 1) Colore font grigio
    work.Font.Color = RGB(128, 128, 128)
    work.Font.ColorIndex = wdGray50

    ' 2) No evidenziazione "gialla"
    work.HighlightColorIndex = wdNoHighlight

    ' 3) Sfondo tenue come fallback (sul testo, non sul paragrafo)
    work.Shading.Texture = wdTextureNone
    work.Shading.ForegroundPatternColor = wdColorAutomatic
    work.Shading.BackgroundPatternColor = RGB(240, 240, 240)
    On Error GoTo 0
End Sub





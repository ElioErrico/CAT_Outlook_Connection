Attribute VB_Name = "modCheshireCatCore"
'==========================
' Modulo: modCheshireCatCore
'==========================
Option Explicit

' ---- Binder per usare la Selection del WordEditor in Outlook ----
Private mSel As Word.Selection

' Chiamala da modOutlookCheshireCat prima di usare funzioni che scrivono
Public Sub CCAT_BindSelection(ByVal wSel As Word.Selection)
    Set mSel = wSel
End Sub

' Ritorna la Selection corrente:
' - se in Outlook: quella passata via bind (mSel)
' - se in Word standalone: fallback a Word.Application.Selection
Public Function CurSel() As Word.Selection
    If Not mSel Is Nothing Then
        Set CurSel = mSel
        Exit Function
    End If
    On Error Resume Next
    Set CurSel = Word.Application.Selection
    On Error GoTo 0
    If CurSel Is Nothing Then
        Err.Raise vbObjectError + 9001, , "Selection Word non inizializzata (chiama CCAT_BindSelection)."
    End If
End Function

' ========= INVIO TESTO E INSERIMENTO RISPOSTA =========
Public Sub InviaTestoAChat()
    Dim sel As Word.Selection
    Dim selectedText As String
    Dim response As String
    Dim startRange As Word.Range

    Set sel = CurSel()
    If sel.Range.Start = sel.Range.End Then
        MsgBox "Nessun testo selezionato", vbExclamation
        Exit Sub
    End If

    ' Costruisce payload pulito: testo + eventuali tabelle convertite in Markdown
    selectedText = BuildMessageFromSelection(sel.Range)
    If Len(Trim$(selectedText)) = 0 Then
        MsgBox "La selezione è vuota dopo la normalizzazione.", vbExclamation
        Exit Sub
    End If

    ' Chiama la tua API (modCheshireCatApi)
    response = CheshireCat_Chat(selectedText)
    If Len(response) = 0 Then
        MsgBox "Risposta vuota dall'API.", vbExclamation
        Exit Sub
    End If

    ' Punto di inserimento dopo la selezione corrente
    Set startRange = sel.Range
    startRange.Collapse Direction:=wdCollapseEnd
    startRange.Select
    sel.TypeParagraph

    ' Inserisce subito testo + tabelle formattate
    InsertAIResponseWithMarkdownTables response
End Sub

Public Sub CancellaCronologiaChat()
    Dim jwtToken As String
    Dim success As Boolean

    jwtToken = GetJWToken() ' <-- dal modulo API
    If Left$(jwtToken, 6) = "Errore" Or Len(jwtToken) = 0 Then
        MsgBox "Errore durante il recupero del token: " & jwtToken, vbCritical
        Exit Sub
    End If

    success = ClearChatHistory(jwtToken) ' <-- dal modulo API
    If success Then
        MsgBox "Cronologia cancellata con successo!"
    Else
        MsgBox "Errore durante la cancellazione della cronologia."
    End If
End Sub

' ========= CONVERSIONE / INSERIMENTO TABELLE MARKDOWN =========

' Inserisce testo e converte TUTTE le tabelle Markdown trovate (2a, 3a, ...).
Public Sub InsertAIResponseWithMarkdownTables(ByVal response As String)
    Dim rest As String
    Dim pre As String, tbl As String, post_ As String
    Dim hasTable As Boolean
    Dim safety As Long

    rest = NormalizeToLf(Replace(response, "\n", vbLf)) ' normalizza fine riga

    safety = 0 ' guardia anti-loop
    Do
        hasTable = ExtractFirstMarkdownTableBlock(rest, pre, tbl, post_)

        If hasTable = False Then
            If Len(pre) > 0 Then InsertMarkdownInline CurSel(), pre, False, False
            Exit Do
        End If

        ' 1) Testo prima della tabella corrente
        If Len(pre) > 0 Then
            InsertMarkdownInline CurSel(), pre, False, False
        End If

        ' 2) Tabella corrente
        EnsureParagraphBeforeInsertion CurSel()
        CreateAndInsertWordTableFromMarkdown tbl, CurSel().Range, CurSel()

        ' 3) Se c'è ancora del testo dopo, aggiungi una riga di separazione
        rest = post_
        If Len(rest) > 0 Then CurSel().TypeParagraph

        safety = safety + 1
        If safety > 20 Then Exit Do ' estrema sicurezza
    Loop
End Sub

' === Strumento manuale: seleziona tabella Markdown e convertila in tabella Word
Public Sub ConvertiTabellaMarkdown()
    Dim selRng As Word.Range
    Dim md As String

    If CurSel().Range.Start = CurSel().Range.End Then
        MsgBox "Nessuna selezione.", vbExclamation
        Exit Sub
    End If

    Set selRng = CurSel().Range
    md = Trim(GetSelectedMarkdownTableText(selRng))
    If Len(md) = 0 Then
        MsgBox "Seleziona il testo della tabella Markdown (righe con '|' ).", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH
    ConvertMarkdownToWord md, selRng, CurSel()
    MsgBox "Tabella convertita con successo!", vbInformation
    Exit Sub
EH:
    MsgBox "Errore durante la conversione: " & Err.Description, vbCritical
End Sub


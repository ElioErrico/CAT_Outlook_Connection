Attribute VB_Name = "modCheshireCatCore"
'==========================
' Modulo: modCheshireCatCore
'==========================
Option Explicit

' ---- Binder per usare la Selection del WordEditor in Outlook/Word ----
Private mSel As Word.Selection
' Punto di inserimento preferito (facoltativo): se impostato, la risposta verrà inserita qui
Private mInsertAt As Word.Range
' Messaggio forzato (costruito altrove: contesto+thread+richiesta+lingua)
Private g_ForcedMessage As String

' ================== BINDINGS & STATE ==================
' Chiamala da modOutlookCheshireCat prima di usare funzioni che scrivono
Public Sub CCAT_BindSelection(ByVal wSel As Word.Selection)
    Set mSel = wSel
End Sub

' (Opzionale) Imposta un punto di inserimento esplicito dove scrivere la risposta
Public Sub CCAT_SetInsertionRange(ByVal r As Word.Range)
    If r Is Nothing Then Exit Sub
    Set mInsertAt = r.Duplicate
End Sub

' Imposta il messaggio da inviare direttamente (one-shot)
Public Sub CCAT_SetForcedMessage(ByVal msg As String)
    g_ForcedMessage = Trim$(msg)
End Sub

' (Facoltativa) Pulisce lo stato
Public Sub CCAT_ClearBindings()
    Set mSel = Nothing
    Set mInsertAt = Nothing
    g_ForcedMessage = ""
End Sub

' Ritorna la Selection corrente:
' - se in Outlook/Word con bind: quella passata via bind (mSel)
' - altrimenti fallback a Word.Application.Selection
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

' Wrapper: mantiene retro-compatibilità con chiamate esistenti
Public Function BuildMessageFromSelection(ByVal r As Word.Range) As String
    BuildMessageFromSelection = modMarkdownHelpers.BuildMessageFromSelection(r)
End Function

' ========= INVIO TESTO E INSERIMENTO RISPOSTA =========
Public Sub InviaTestoAChat()
    Dim sel As Word.Selection
    Dim messageToSend As String
    Dim response As String
    Dim target As Word.Range
    Dim needLineBreak As Boolean

    Set sel = CurSel()

    ' Se non ho un messaggio forzato, richiedo una selezione non vuota
    If Len(g_ForcedMessage) = 0 Then
        If sel.Range.Start = sel.Range.End Then
            MsgBox "Nessun testo selezionato.", vbExclamation
            Exit Sub
        End If
    End If

    ' 1) Costruisci il payload: usa quello forzato se presente, altrimenti normalizza la selezione
    If Len(g_ForcedMessage) > 0 Then
        messageToSend = g_ForcedMessage
        g_ForcedMessage = "" ' one-shot, azzera dopo l'uso
    Else
        messageToSend = modMarkdownHelpers.BuildMessageFromSelection(sel.Range)
        If Len(Trim$(messageToSend)) = 0 Then
            MsgBox "La selezione è vuota dopo la normalizzazione.", vbExclamation
            Exit Sub
        End If
    End If

    ' 2) Chiama l'API (modCheshireCatApi)
    response = modCheshireCatApi.CheshireCat_Chat(messageToSend)
    If Len(response) = 0 Then
        MsgBox "Risposta vuota dall'API.", vbExclamation
        Exit Sub
    End If

    ' 3) Determina un punto di inserimento FUORI dal range grigiato
    If Not mInsertAt Is Nothing Then
        Set target = mInsertAt.Duplicate
        target.Collapse wdCollapseEnd
        needLineBreak = True
        Set mInsertAt = Nothing ' evita riutilizzi involontari
    Else
        Set target = sel.Range.Duplicate
        target.Collapse wdCollapseEnd
        needLineBreak = True
    End If

    ' 4) Inserisci un invio e resetta i formati per evitare eredità
    If needLineBreak Then
        target.InsertParagraphAfter
        target.Collapse wdCollapseEnd
    End If
    With target
        .ParagraphFormat.Reset
        .Font.Reset
        .HighlightColorIndex = wdNoHighlight
        .Shading.Texture = wdTextureNone
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColor = wdColorAutomatic
    End With

    ' 5) Sposta la Selection sul punto di inserimento e scrivi la risposta con gestione Markdown/tabelle
    CurSel().SetRange Start:=target.Start, End:=target.End
    InsertAIResponseWithMarkdownTables response
End Sub

Public Sub CancellaCronologiaChat()
    Dim jwtToken As String
    Dim success As Boolean

    jwtToken = modCheshireCatApi.GetJWToken()
    If Left$(jwtToken, 6) = "Errore" Or Len(jwtToken) = 0 Then
        MsgBox "Errore durante il recupero del token: " & jwtToken, vbCritical
        Exit Sub
    End If

    success = modCheshireCatApi.ClearChatHistory(jwtToken)
    If success Then
        MsgBox "Cronologia cancellata con successo!", vbInformation
    Else
        MsgBox "Errore durante la cancellazione della cronologia.", vbExclamation
    End If
End Sub

' ========= CONVERSIONE / INSERIMENTO TABELLE MARKDOWN =========

' Inserisce testo e converte TUTTE le tabelle Markdown trovate (2a, 3a, ...).
' Usa sempre la Selection corrente (spostata in 'InviaTestoAChat' sul punto target).
Public Sub InsertAIResponseWithMarkdownTables(ByVal response As String)
    Dim rest As String
    Dim pre As String, tbl As String, post_ As String
    Dim hasTable As Boolean
    Dim safety As Long

    ' Normalizza fine riga e gestisci "\n" letterali
    rest = modMarkdownHelpers.NormalizeToLf(Replace(response, "\n", vbLf))

    safety = 0 ' guardia anti-loop
    Do
        hasTable = modMarkdownHelpers.ExtractFirstMarkdownTableBlock(rest, pre, tbl, post_)
        If hasTable = False Then
            If Len(pre) > 0 Then modMarkdownHelpers.InsertMarkdownInline CurSel(), pre, False, False
            Exit Do
        End If

        ' 1) Testo prima della tabella corrente
        If Len(pre) > 0 Then
            modMarkdownHelpers.InsertMarkdownInline CurSel(), pre, False, False
        End If

        ' 2) Tabella corrente (assicura un paragrafo "pulito" prima)
        modMarkdownHelpers.EnsureParagraphBeforeInsertion CurSel()
        modMarkdownHelpers.CreateAndInsertWordTableFromMarkdown tbl, CurSel().Range, CurSel()

        ' 3) Continua con il resto; se c'è altro testo, separa con un invio
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
    md = Trim(modMarkdownHelpers.GetSelectedMarkdownTableText(selRng))
    If Len(md) = 0 Then
        MsgBox "Seleziona il testo della tabella Markdown (righe con '|' ).", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH
    modMarkdownHelpers.ConvertMarkdownToWord md, selRng, CurSel()
    MsgBox "Tabella convertita con successo!", vbInformation
    Exit Sub
EH:
    MsgBox "Errore durante la conversione: " & Err.Description, vbCritical
End Sub


Attribute VB_Name = "modOutlookCheshireCat"
'==========================
' Modulo: modOutlookCheshireCat
'==========================
Option Explicit

' --- Prende la Selection del WordEditor della finestra di composizione ---
Private Function GetComposeSelection() As Word.Selection
    Dim insp As Outlook.Inspector
    On Error Resume Next
    Set insp = Application.ActiveInspector
    On Error GoTo 0
    If insp Is Nothing Then Exit Function
    If insp.EditorType <> olEditorWord Then Exit Function

    Dim wdDoc As Word.Document
    Set wdDoc = insp.WordEditor
    Set GetComposeSelection = wdDoc.Application.Selection
End Function

Public Sub CCAT_InserisciRispostaDaSelezione_Outlook()
    Dim insp As Outlook.Inspector
    Dim sel As Word.Selection
    Dim rGray As Word.Range
    Dim insertAt As Word.Range

    ' 1) Inspector + forza HTML
    On Error Resume Next
    Set insp = Application.ActiveInspector
    On Error GoTo 0
    If insp Is Nothing Then
        MsgBox "Apri una finestra di composizione con editor Word.", vbExclamation
        Exit Sub
    End If
    EnsureHtmlBody insp

    ' 2) Selection valida
    Set sel = GetComposeSelection()
    If sel Is Nothing Then
        MsgBox "Apri una finestra di composizione con editor Word.", vbExclamation
        Exit Sub
    End If
    If sel.Range.Start = sel.Range.End Then
        MsgBox "Nessun testo selezionato.", vbExclamation
        Exit Sub
    End If

    ' 3) Duplica la selezione e grigia SOLO il testo (non il paragrafo)
    Set rGray = sel.Range.Duplicate
    GrayOutRange rGray

    ' 4) Prepara un punto di inserimento DOPO il range grigio, con formattazione pulita
    Set insertAt = rGray.Duplicate
    insertAt.Collapse wdCollapseEnd
    insertAt.ParagraphFormat.Reset
    insertAt.Font.Reset
    insertAt.HighlightColorIndex = wdNoHighlight
    insertAt.Shading.Texture = wdTextureNone
    insertAt.Shading.ForegroundPatternColor = wdColorAutomatic
    insertAt.Shading.BackgroundPatternColor = wdColorAutomatic

    ' (facoltativo) vai a capo prima della risposta
    insertAt.InsertParagraphAfter
    insertAt.Collapse wdCollapseEnd

    ' 5) Binda la selezione al core e passa il punto di inserimento
    '    => Aggiungi in modCheshireCatCore una property/argomento per il target range
    modCheshireCatCore.CCAT_BindSelection sel
    modCheshireCatCore.CCAT_SetInsertionRange insertAt   ' <-- aggiungi questa API nel core
    modCheshireCatCore.InviaTestoAChat
End Sub




' ===== Cancella cronologia remota =====
Public Sub CCAT_CancellaCronologia_Outlook()
    Dim jwt As String
    jwt = GetJWToken()
    If Left$(jwt, 6) = "Errore" Or Len(jwt) = 0 Then
        MsgBox "Errore token: " & jwt, vbCritical
        Exit Sub
    End If
    If ClearChatHistory(jwt) Then
        MsgBox "Cronologia cancellata.", vbInformation
    Else
        MsgBox "Errore durante la cancellazione.", vbExclamation
    End If
End Sub



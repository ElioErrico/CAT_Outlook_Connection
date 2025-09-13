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

' ===== Invia SELEZIONE + contesto (thread completo) al backend e inserisce la risposta =====
Public Sub CCAT_InserisciRispostaDaSelezione_Outlook()
    Dim insp As Outlook.Inspector
    Dim sel As Word.Selection
    Dim selectedText As String
    Dim response As String
    Dim startRange As Word.Range
    Dim rGray As Word.Range

    ' 1) Recupera l'inspector e forza HTML (per garantire la formattazione)
    On Error Resume Next
    Set insp = Application.ActiveInspector
    On Error GoTo 0
    If insp Is Nothing Then
        MsgBox "Apri una finestra di composizione con editor Word.", vbExclamation
        Exit Sub
    End If
    modMarkdownHelpers.EnsureHtmlBody insp

    ' 2) (Ri)ottieni la Selection dal WordEditor
    Set sel = GetComposeSelection()
    If sel Is Nothing Then
        MsgBox "Apri una finestra di composizione con editor Word.", vbExclamation
        Exit Sub
    End If
    If sel.Range.Start = sel.Range.End Then
        MsgBox "Nessun testo selezionato.", vbExclamation
        Exit Sub
    End If

    ' 3) Colora in grigio il testo selezionato (senza alterare il contenuto)
    Set rGray = sel.Range.Duplicate
    modMarkdownHelpers.GrayOutRange rGray

    ' 4) Bind al core e invia
    modCheshireCatCore.CCAT_BindSelection sel
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



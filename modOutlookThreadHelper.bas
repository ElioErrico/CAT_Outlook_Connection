Attribute VB_Name = "modOutlookThreadHelper"
' ==========================
' Modulo: modOutlookThreadHelper
' ==========================
Option Explicit

' Limiti di default (puoi modificarli)
Public Const THREAD_MAX_ITEMS As Long = 12
Public Const THREAD_MAX_CHARS_PER_ITEM As Long = 2500

' API principale: restituisce il testo dello scambio mail
Public Function BuildMailExchange(ByVal insp As Outlook.Inspector, _
                                  Optional ByVal maxItems As Long = THREAD_MAX_ITEMS, _
                                  Optional ByVal maxChars As Long = THREAD_MAX_CHARS_PER_ITEM) As String
    On Error GoTo EH
    Dim mi As Outlook.MailItem
    Dim conv As Outlook.Conversation
    Dim tbl As Outlook.Table
    Dim rw As Outlook.Row
    Dim sb As String
    Dim n As Long

    If insp Is Nothing Then Exit Function
    If insp.EditorType <> olEditorWord Then Exit Function

    Set mi = TryGetMailItem(insp.CurrentItem)
    If mi Is Nothing Then Exit Function

    Set conv = mi.GetConversation
    If conv Is Nothing Then
        ' Fallback: solo il messaggio corrente
        BuildMailExchange = FormatOneItem(mi, maxChars)
        Exit Function
    End If

    Set tbl = conv.GetTable
    If tbl Is Nothing Then
        BuildMailExchange = FormatOneItem(mi, maxChars)
        Exit Function
    End If

    ' Ordina cronologicamente
    tbl.Sort "[ReceivedTime]", olAscending

    Do Until tbl.EndOfTable
        Set rw = tbl.GetNextRow
        Dim entryId As String
        entryId = rw("EntryID")

        Dim it As Object
        Set it = Application.Session.GetItemFromID(entryId)
        If Not it Is Nothing Then
            If TypeOf it Is Outlook.MailItem Then
                sb = sb & FormatOneItem(it, maxChars) & vbCrLf & String(60, "-") & vbCrLf
                n = n + 1
                If n >= maxItems Then Exit Do
            End If
        End If
    Loop

    BuildMailExchange = sb
    Exit Function
EH:
    ' In caso di errore, restituisci almeno il corrente
    On Error Resume Next
    If Not mi Is Nothing Then
        BuildMailExchange = FormatOneItem(mi, maxChars)
    End If
End Function

' --- Helpers ---

Private Function TryGetMailItem(ByVal it As Object) As Outlook.MailItem
    On Error Resume Next
    If TypeOf it Is Outlook.MailItem Then
        Set TryGetMailItem = it
    End If
End Function

Private Function FormatOneItem(ByVal m As Outlook.MailItem, ByVal maxChars As Long) As String
    Dim hdr As String
    Dim txt As String

    hdr = "From: " & m.SenderName & " <" & GetSenderSmtp(m) & ">" & vbCrLf & _
          "To: " & NullIfEmpty(m.To) & IIf(Len(m.CC) > 0, "  |  Cc: " & m.CC, "") & vbCrLf & _
          "Date: " & Format$(IIf(m.ReceivedTime = #1/1/4501#, m.SentOn, m.ReceivedTime), "yyyy-mm-dd hh:nn") & _
          "  |  Subject: " & m.Subject

    txt = NormalizeText(m.body)
    If maxChars > 0 And Len(txt) > maxChars Then
        txt = Left$(txt, maxChars) & "..."
    End If

    FormatOneItem = hdr & vbCrLf & txt
End Function

Private Function GetSenderSmtp(ByVal m As Outlook.MailItem) As String
    On Error Resume Next
    If m.SenderEmailType = "EX" Then
        Dim pa As Outlook.PropertyAccessor
        Set pa = m.PropertyAccessor
        ' PR_SENDER_EMAIL_ADDRESS
        GetSenderSmtp = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001E")
        If Len(GetSenderSmtp) = 0 Then
            Dim exu As Outlook.ExchangeUser
            Set exu = m.Sender.GetExchangeUser
            If Not exu Is Nothing Then GetSenderSmtp = exu.PrimarySmtpAddress
        End If
    Else
        GetSenderSmtp = m.SenderEmailAddress
    End If
    If Len(GetSenderSmtp) = 0 Then GetSenderSmtp = m.SenderName
End Function

Private Function NormalizeText(ByVal s As String) As String
    ' Pulisce un minimo il testo
    s = Replace$(s, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    s = Replace$(s, Chr$(160), " ")
    NormalizeText = Trim$(s)
End Function

Private Function NullIfEmpty(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then
        NullIfEmpty = "(n/d)"
    Else
        NullIfEmpty = s
    End If
End Function



Attribute VB_Name = "modCheshireCatApi"
' ==========================
' Modulo: modCheshireCatApi
' ==========================
Option Explicit

' =========== CONFIG ===========
Public Const DEFAULT_URL As String = "http://localhost:1865"
Public Const DEFAULT_USERNAME As String = "admin"
Public Const DEFAULT_PASSWORD As String = "admin"

' ===== TIMEOUTS (ms) =====
Private Const TO_RESOLVE  As Long = 5000
Private Const TO_CONNECT  As Long = 15000
Private Const TO_SEND     As Long = 60000
Private Const TO_RECEIVE  As Long = 300000

' ====== Helpers JSON ======
Private Function ExtractJsonValue(jsonText As String, key As String) As String
    Dim startPos As Long, valueStart As Long, endPos As Long
    startPos = InStr(1, jsonText, """" & key & """:""")
    If startPos = 0 Then Exit Function
    valueStart = startPos + Len("""" & key & """:""")
    endPos = InStr(valueStart, jsonText, """")
    If endPos > valueStart Then ExtractJsonValue = Mid$(jsonText, valueStart, endPos - valueStart)
End Function

Private Function EscapeJsonString(text As String) As String
    Dim s As String
    s = Replace(text, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJsonString = s
End Function

' ========== HTTP / AUTH ==========
Public Function GetJWToken() As String
    Dim http As Object, url As String, body As String

    url = DEFAULT_URL & "/auth/token"
    body = "{""username"":""" & EscapeJsonString(DEFAULT_USERNAME) & """,""password"":""" & EscapeJsonString(DEFAULT_PASSWORD) & """}"

    On Error GoTo ErrH
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    With http
        .Open "POST", url, False
        .setTimeouts TO_RESOLVE, TO_CONNECT, TO_SEND, TO_RECEIVE
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Connection", "Keep-Alive"
        .Send body
        If .Status = 200 Then
            GetJWToken = ExtractJsonValue(.responseText, "access_token")
            If Len(GetJWToken) = 0 Then GetJWToken = "Token non trovato nella risposta"
        Else
            GetJWToken = "Errore HTTP: " & .Status & " - " & .StatusText
        End If
    End With
    Exit Function
ErrH:
    GetJWToken = "Errore: " & Err.Number & " - " & Err.Description
End Function

Public Function CheshireCat_Chat(messageText As String) As String
    Dim http As Object, url As String, body As String, jwt As String, content As String

    jwt = GetJWToken()
    If Left$(jwt, 5) = "Error" Or Left$(jwt, 6) = "Errore" Then
        CheshireCat_Chat = jwt
        Exit Function
    End If

    url = DEFAULT_URL & "/message"
    body = "{""text"":""" & EscapeJsonString(messageText) & """}"

    On Error GoTo ErrH
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    With http
        .Open "POST", url, False
        .setTimeouts TO_RESOLVE, TO_CONNECT, TO_SEND, TO_RECEIVE
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & jwt
        .setRequestHeader "Connection", "Keep-Alive"
        .Send body
        If .Status <> 200 Then
            CheshireCat_Chat = "Errore HTTP: " & .Status & " - " & .StatusText
            Exit Function
        End If
    End With

    content = ExtractJsonValue(http.responseText, "content")
    CheshireCat_Chat = IIf(Len(content) > 0, content, "Campo 'content' non trovato nella risposta")
    Exit Function
ErrH:
    CheshireCat_Chat = "Errore: " & Err.Number & " - " & Err.Description
End Function

Public Function ClearChatHistory(ByVal jwtToken As String) As Boolean
    Dim http As Object, url As String
    url = DEFAULT_URL & "/memory/conversation_history"
    On Error GoTo ErrH
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    With http
        .Open "DELETE", url, False
        .setTimeouts TO_RESOLVE, TO_CONNECT, TO_SEND, TO_RECEIVE
        .setRequestHeader "Authorization", "Bearer " & jwtToken
        .Send
        ClearChatHistory = (.Status = 200)
    End With
    Exit Function
ErrH:
    ClearChatHistory = False
End Function

' (Opzionale) Classificatore semplice
Public Function CheshireCat_Classify(sentence As String, labels As Variant) As String
    Dim labels_list As String, i As Long, prompt As String
    If TypeName(labels) = "String" Then
        labels_list = "- " & Replace(labels, ",", vbNewLine & "- ")
    ElseIf IsArray(labels) Then
        For i = LBound(labels) To UBound(labels)
            labels_list = labels_list & "- " & labels(i) & vbNewLine
        Next i
    End If
    prompt = "Classify this sentence:" & vbNewLine & """" & sentence & """" & vbNewLine & vbNewLine & _
             "Allowed classes are:" & vbNewLine & labels_list & vbNewLine & vbNewLine & "Just output the class, nothing else."
    CheshireCat_Classify = CheshireCat_Chat(prompt)
End Function



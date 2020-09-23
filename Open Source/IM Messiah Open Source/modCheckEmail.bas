Attribute VB_Name = "modCheckEmail"
Public Function CheckEmail(clPassport As String, clPassword As String, logintime As String, auth As String, compose As Integer, toemail As String)
sl = DateToTimestamp(Now()) - logintime
creds = MSNEncryptPw(auth & sl & clPassword)
On Error GoTo errormsg
Open "login.txt" For Input As 1
    txt = Input(LOF(1), 1)
Close #1
For x = 0 To 9
    txtline = Split(txt, vbCrLf)(x)
    full = full + txtline + vbCrLf
Next
    Dim startsting As String
    startstring = "   <input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34)
    txtline = startstring & "login" & Chr(34) & " value=" & Chr(34) & Split(clPassport, "@")(0) & Chr(34) & ">"
    full = full + txtline + vbCrLf 'username
    txtline = startstring & "username" & Chr(34) & " value=" & Chr(34) & clPassport & Chr(34) & ">"
    full = full + txtline + vbCrLf 'email address
    full = full + Split(txt, vbCrLf)(12) + vbCrLf 'sid
    full = full + Split(txt, vbCrLf)(13) + vbCrLf 'kv
    full = full + Split(txt, vbCrLf)(14) + vbCrLf 'id
    txtline = startstring & "sl" & Chr(34) & " value=" & Chr(34) & sl & Chr(34) & ">"
    full = full + txtline + vbCrLf 'sl
    If compose <> 1 Then
        full = full + Split(txt, vbCrLf)(16) + vbCrLf 'rru
    ElseIf compose = 1 Then
        txtline = startstring & "rru" & Chr(34) & " value=" & Chr(34) & "/cgi-bin/compose?mailto=1&to=" & toemail & Chr(34) & ">"
        full = full + txtline + vbCrLf 'rru
    End If
    txtline = startstring & "auth" & Chr(34) & " value=" & Chr(34) & auth & Chr(34) & ">"
    full = full + txtline + vbCrLf 'auth
    txtline = startstring & "creds" & Chr(34) & " value=" & Chr(34) & creds & Chr(34) & ">"
    full = full + txtline + vbCrLf 'creds
    full = full + Split(txt, vbCrLf)(19) + vbCrLf 'svc
    full = full + Split(txt, vbCrLf)(20) + vbCrLf 'js
    full = full + Split(txt, vbCrLf)(21) + vbCrLf '</form>
    full = full + Split(txt, vbCrLf)(22) + vbCrLf '</body>
    full = full + Split(txt, vbCrLf)(23) '</html>
On Error GoTo Create
    FileSystem.Kill "login.htm"
Create:
    Open "login.htm" For Output As #1
    Print #1, full
    Close #1
    Shell "explorer login.htm"
    frmContacts.tmrerasehtm.Enabled = True
    Exit Function
errormsg:
    MsgBox (Err.Description)
End Function

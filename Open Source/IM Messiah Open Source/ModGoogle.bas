Attribute VB_Name = "ModGoogle"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public sURL As String
Public strPacket As String
Public strNext As String
Public strPrev As String

Public Function GetStringBetween(ByVal Str As String, ByVal str1 As String, ByVal str2 As String, Optional ByVal st As Long = 0) As String
    On Error Resume Next
    Dim s1, s2, s, l As Long
    Dim foundstr As String
    
    s1 = InStr(st + 1, Str, str1, vbTextCompare)
    s2 = InStr(s1 + 1, Str, str2, vbTextCompare)
    
    If s1 = 0 Or s2 = 0 Or IsNull(s1) Or IsNull(s2) Then
        foundstr = Str
    Else
        s = s1 + Len(str1)
        l = s2 - s
        foundstr = Mid(Str, s, l)
    End If
    
    GetStringBetween = foundstr
End Function

Sub OpenURL(URL As String)
    ShellExecute hwnd, "open", URL, vbNullString, vbNullString, conSwNormal
End Sub

Public Function CleanUp(sData As String)

    If InStr(sData, LCase("&amp;")) Then
        sData = Replace(sData, LCase("&amp;"), "&")
    End If
    If InStr(sData, LCase("&quot;")) Then
        sData = Replace(sData, LCase("&quot;"), Chr(34))
    End If
    If InStr(sData, LCase("&nbsp;")) Then
        sData = Replace(sData, LCase("&nbsp;"), " ")
    End If
    If InStr(sData, LCase("&copy;")) Then
        sData = Replace(sData, LCase("&copy;"), "©")
    End If
    If InStr(sData, LCase("&trade;")) Then
        sData = Replace(sData, LCase("&trade;"), "™")
    End If
    If InStr(sData, "<b>") Then
        sData = Replace(sData, "<b>", "")
        sData = Replace(sData, "</b>", "")
    End If

    If InStr(sData, "</a>") Then
        sData = Replace(sData, "</a>", "")
    End If
    If InStr(sData, "<a href=") Then
        Temp$ = GetStringBetween(sData, "<a href=", ">")
        sData = Replace(sData, "<a href=", "")
        sData = Replace(sData, ">", "")
        sData = Replace(sData, Temp$, "")
        sData = Temp$ & " - " & sData
    End If
    
    CleanUp = sData
End Function

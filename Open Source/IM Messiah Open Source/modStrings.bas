Attribute VB_Name = "modStrings"
Public strSID As String, strCKI As String, strMIP As String, strMaster As String
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

Public Function URLEncode(strData As String) As String

Dim intCount As Integer
Dim strBuffer As String
Dim strReturn As String

strReturn = strData
    For intCount = 1 To Len(strData)
        strBuffer = Mid(strData, intCount, 1)
        If Not strBuffer Like "[a-z,A-Z,0-9]" Then
            strReturn = Replace(strReturn, strBuffer, "%" & Hex(Asc(strBuffer)))
        End If
    Next intCount
    URLEncode = strReturn
End Function

Public Function URLDecode(strInput As String) As String
Dim strCodedChar  As String
Dim intBeginBy As Integer
intBeginBy = 1
Begin:
For bp1 = intBeginBy To Len(strInput)
    If Mid(strInput, bp1, 1) = "%" Then
        strCodedChar = Mid(strInput, bp1 + 1, 1) & Mid(strInput, bp1 + 2, 1)
    On Error GoTo nextthing
        strInput = Left(strInput, bp1 - 1) & Chr(Val("&H" & strCodedChar)) & Right(strInput, Len(strInput) - bp1 - 2)
        intBeginBy = bp1
        DoEvents
        GoTo Begin
    End If
Next bp1
nextthing:
URLDecode = strInput
End Function

Public Function bgrhex2rgb(code) As String
  Dim newcode As String
  newcode = String(6 - Len(code), "0") & code
  bgrhex2rgb = RGB(Val("&H" & Right(newcode, 2)), Val("&H" & Mid(newcode, 3, 2)), Val("&H" & Left(newcode, 2)))
End Function

Public Function GetBetween(Str As String, Optional dStart As String, Optional dEnd As String, Optional Length As Long) As String
    Dim x1 As Long, x2 As Long
    
    'Start?
    x1 = IIf(dStart = "", 1, InStr(1, LCase$(Str), LCase$(dStart)) + Len(dStart))
    
    'Rip the string :0
    If x1 > 0 Then
        If dEnd = "" Then
            GetBetween = Mid$(Str, x1)
        Else
            x2 = InStr(x1, LCase$(Str), LCase$(dEnd)) - x1
            If x2 > 0 Then
                GetBetween = Mid$(Str, x1, x2)
            Else
                GetBetween = "n/f"
            End If
        End If
    Else
        GetBetween = "n/f"
    End If
    
    'Length?
    If Length > 0 And GetBetween <> "n/f" Then GetBetween = Left$(GetBetween, Length)
End Function

Public Function HexToBin(ByVal Data As String)
    Dim DataOut As String, x As Long, sHex As String
    For x = 1 To Len(Data) Step 2
        sHex = Mid$(Data, x, 2)
        DataOut = DataOut & Chr(Val("&H" & sHex))
    Next
    HexToBin = DataOut
End Function


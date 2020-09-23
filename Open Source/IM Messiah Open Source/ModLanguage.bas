Attribute VB_Name = "ModLanguage"
Private Type Lang
    Phrases() As String
    Name As String
End Type

Public OtherLang() As Lang
Dim Langs As Integer

Public Function GetPhrase(LangID, ID)
    If LangID = 0 Then
        LangID = 1
    End If
    If UBound(OtherLang(LangID).Phrases()) >= ID Then
        GetPhrase = OtherLang(LangID).Phrases(ID)
    Else
        GetPhrase = OtherLang(LangID).Phrases(1)
    End If
End Function

Public Function ShowLangs()
    Dim StrOutput As String
    For x = 1 To Langs
        StrOutput = StrOutput & vbNewLine & x & "." & OtherLang(x).Name
    Next
    ShowLangs = StrOutput
End Function

Public Sub LoadLang()
    ReDim OtherLang(0) As Lang
    Dim StrData As String
    Dim LangText() As String
    Open App.Path & "\data\Lang.txt" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, StrData
            If Left(StrData, 1) = "A" Then
                Langs = Langs + 1
                ReDim Preserve OtherLang(Langs) As Lang
                OtherLang(Langs).Name = Mid(StrData, 2)
                ReDim OtherLang(Langs).Phrases(0) As String
            ElseIf Left(StrData, 1) = "B" Then
                LangText() = Split(Mid$(StrData, 2), "||")
                For x = 1 To UBound(LangText())
                    ReDim Preserve OtherLang(x).Phrases(UBound(OtherLang(x).Phrases()) + 1) As String
                    OtherLang(x).Phrases(UBound(OtherLang(x).Phrases())) = LangText(x)
                Next
            End If
        Loop Until EOF(1)
    End If
    Close
End Sub

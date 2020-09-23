Attribute VB_Name = "ModPreferences"
Private Type People
    Email As String
    Name As String
    Language As Integer
    Warnings As Integer
End Type

Public Prefs() As People
Dim PrefCount As Integer

Sub LoadPref()
    ReDim Prefs(0) As People
    PrefCount = 0
    If Dir(App.Path & "\data\Pref.txt") <> "Pref.txt" Then
        Exit Sub
    End If
    Dim StrData As String
    Open App.Path & "\data\Pref.txt" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, StrData
            If Left(StrData, 1) = "A" Then
                PrefCount = PrefCount + 1
                ReDim Preserve Prefs(PrefCount) As People
                Prefs(PrefCount).Email = UCase(Mid(StrData, 2))
            ElseIf Left(StrData, 1) = "B" Then
                Prefs(PrefCount).Name = Mid(StrData, 2)
            ElseIf Left(StrData, 1) = "C" Then
                Prefs(PrefCount).Language = Val(Trim(Mid(StrData, 2)))
            ElseIf Left(StrData, 1) = "D" Then
                Prefs(PrefCount).Warnings = Val(Trim(Mid(StrData, 2)))
            End If
        Loop Until EOF(1)
    End If
    Close
End Sub

Sub SavePrefs()
    If PrefCount = 0 Then
        Exit Sub
    End If
    Open App.Path & "\data\Pref.txt" For Output As #1
    For x = 1 To PrefCount
        Print #1, "A" & Prefs(x).Email
        Print #1, "B" & Prefs(x).Name
        Print #1, "C" & Prefs(x).Language
    Next
    Close
End Sub

Function GotPref(Email) As Boolean
    If PrefCount = 0 Then
        GotPref = False
        Exit Function
    End If
    GotPref = False
    For x = 1 To PrefCount
        If Prefs(x).Email = UCase(Email) Then
            GotPref = True
            Exit Function
        End If
    Next
End Function

Function PrefEmail(Name) As String
    PrefEmail = ""
    If PrefCount = 0 Then
        Exit Function
    End If
    For x = 1 To PrefCount
        If UCase(Prefs(x).Name) = UCase(Name) Then
            PrefEmail = Prefs(x).Email
            Exit Function
        End If
    Next
End Function

Function PrefGetName(Email) As String
    PrefGetName = Email
    If PrefCount = 0 Then
        Exit Function
    End If
    For x = 1 To PrefCount
        If Prefs(x).Email = UCase(Email) Then
            If Prefs(x).Name <> "" Then
                PrefGetName = Prefs(x).Name
            End If
            Exit Function
        End If
    Next
End Function


Function PrefName(Email) As String
    If PrefCount = 0 Then
        PrefName = ""
        Exit Function
    End If
    PrefName = ""
    For x = 1 To PrefCount
        If Prefs(x).Email = UCase(Email) Then
            PrefName = Prefs(x).Name
            Exit Function
        End If
    Next
End Function

Function PrefIndex(Email)
    For x = 1 To PrefCount
        If Prefs(x).Email = UCase(Email) Then
            PrefIndex = x
            Exit Function
        End If
    Next
End Function

Function NewPrefName(Email, Name) As Boolean
    If PrefCount = 0 Then
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = Name
        Prefs(PrefCount).Language = 1
        Prefs(PrefCount).Warnings = 0
        NewPrefName = True
        SavePrefs
        Exit Function
    End If
    For x = 1 To PrefCount
        If UCase(Prefs(x).Name) = UCase(Name) Then
            If Prefs(x).Email <> UCase(Email) Then
                NewPrefName = False
                Exit Function
            End If
        End If
    Next
    If GotPref(Email) = True Then
        Prefs(PrefIndex(Email)).Name = Name
        NewPrefName = True
    Else
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = Name
        Prefs(PrefCount).Language = 1
        Prefs(PrefCount).Warnings = 0
        NewPrefName = True
    End If
    SavePrefs
End Function

Function GetLang(Email)
    If PrefCount = 0 Then
        GetLang = 1
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = ""
        Prefs(PrefCount).Language = 1
        Prefs(PrefCount).Warnings = 0
        SavePrefs
        Exit Function
    End If
    If GotPref(Email) = True Then
        GetLang = Prefs(PrefIndex(Email)).Language
    Else
        GetLang = 1
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = ""
        Prefs(PrefCount).Language = 1
        Prefs(PrefCount).Warnings = 0
        SavePrefs
    End If
    
End Function

Sub SaveLang(Email, Lang)
    If PrefCount = 0 Then
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = ""
        Prefs(PrefCount).Language = Lang
        Prefs(PrefCount).Warnings = 0
        SavePrefs
        Exit Sub
    End If
    If GotPref(Email) = True Then
        Prefs(PrefIndex(Email)).Language = Lang
    Else
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = ""
        Prefs(PrefCount).Language = Lang
        Prefs(PrefCount).Warnings = 0
    End If
    SavePrefs
End Sub

Sub WarnPerson(Email)
    If PrefCount = 0 Then
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = ""
        Prefs(PrefCount).Language = 1
        Prefs(PrefCount).Warnings = 1
        SavePrefs
        Exit Sub
    End If
    If GotPref(Email) = True Then
        Prefs(PrefIndex(Email)).Warnings = Prefs(PrefIndex(Email)).Warnings + 1
        If Prefs(PrefIndex(Email)).Warnings = 3 Then
            Prefs(PrefIndex(Email)).Warnings = 0
            AddBan Message
        End If
    Else
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = ""
        Prefs(PrefCount).Language = 1
        Prefs(PrefCount).Warnings = 1
    End If
    SavePrefs
End Sub

Function GetWarnings(Email)
    If PrefCount = 0 Then
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = ""
        Prefs(PrefCount).Language = 1
        Prefs(PrefCount).Warnings = 0
        SavePrefs
        GetWarnings = 0
        Exit Function
    End If
    If GotPref(Email) = True Then
        GetWarnings = Prefs(PrefIndex(Email)).Warnings
    Else
        PrefCount = PrefCount + 1
        ReDim Preserve Prefs(PrefCount) As People
        Prefs(PrefCount).Email = UCase(Email)
        Prefs(PrefCount).Name = ""
        Prefs(PrefCount).Language = 1
        Prefs(PrefCount).Warnings = 0
        GetWarnings = 0
        SavePrefs
    End If
End Function


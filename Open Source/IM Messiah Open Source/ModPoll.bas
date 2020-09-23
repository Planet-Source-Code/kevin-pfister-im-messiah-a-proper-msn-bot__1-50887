Attribute VB_Name = "ModPoll"

Public PollTitle As String
Public PollPosts() As String
Public PollVotes() As Integer
Public PollVoted() As String

Sub LoadPoll()
    ReDim PollPosts(0) As String
    ReDim PollVotes(0) As Integer
    ReDim PollVoted(0) As String
    FEnt = 0
    If Dir(App.Path & "\data\Poll.txt") <> "Poll.txt" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\Poll.txt" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            If Left(strData, 1) = "A" Then
                PollTitle = Mid(strData, 2)
            ElseIf Left(strData, 1) = "B" Then
                ReDim Preserve PollPosts(UBound(PollPosts()) + 1) As String
                ReDim Preserve PollVotes(UBound(PollVotes()) + 1) As Integer
                PollPosts(UBound(PollPosts())) = Mid(strData, 2)
            ElseIf Left(strData, 1) = "C" Then
                PollVotes(UBound(PollVotes())) = Val(Mid(strData, 2))
            ElseIf Left(strData, 1) = "D" Then
                ReDim Preserve PollVoted(UBound(PollVoted()) + 1) As String
                PollVoted(UBound(PollVoted())) = UCase(Mid(strData, 2))
            End If
        Loop Until EOF(1)
    End If
    Close
End Sub

Function GetPollCode()
    If Dir(App.Path & "\data\Poll.txt") <> "Poll.txt" Then
        GetPollCode = ""
        Exit Function
    End If
    Dim strData As String
    Open App.Path & "\data\Poll.txt" For Input As #1
    If LOF(1) <> 0 Then
        GetPollCode = Input(LOF(1), 1)
    End If
    Close
End Function

Sub SavePollCode(Code)
    Dim strData As String
    Open App.Path & "\data\Poll.txt" For Output As #1
    Print #1, Code
    Close
End Sub

Function HasVoted(Email) As Boolean
    HasVoted = False
    For X = 1 To UBound(PollVoted())
        If PollVoted(X) = UCase(Email) Then
            HasVoted = True
            Exit Function
        End If
    Next
End Function

Sub SavePoll()
    Open App.Path & "\data\Poll.txt" For Output As #1
        If UBound(PollPosts()) <> 0 Then
            Print #1, "A" & PollTitle
            For X = 1 To UBound(PollPosts())
                Print #1, "B" & PollPosts(X)
                Print #1, "C" & PollVotes(X)
            Next
            For X = 1 To UBound(PollVoted())
                Print #1, "D" & PollVoted(X)
            Next
        End If
    Close
End Sub

Function ReturnPoll()
    Dim strData As String
    strData = PollTitle & vbNewLine
    For X = 1 To UBound(PollPosts())
        strData = strData & vbNewLine & X & "." & PollPosts(X) & "  [" & PollVotes(X) & " Votes]"
    Next
    ReturnPoll = strData
End Function

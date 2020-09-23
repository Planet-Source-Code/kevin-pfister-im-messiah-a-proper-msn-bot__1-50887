Attribute VB_Name = "ModAI"
Private Type AIQuestion
    Question As String
    Answers() As String
End Type

Public AIResponses() As AIQuestion
Public Greetings() As String

Dim Entries As Integer

Sub LoadAI()
    ReDim AIResponses(0) As AIQuestion
    ReDim Greetings(0) As String
    Entries = 0
    If Dir(App.Path & "\data\AI.txt") <> "AI.txt" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\AI.txt" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            If Left(strData, 1) = "A" Then
                ReDim Preserve Greetings(UBound(Greetings()) + 1) As String
                Greetings(UBound(Greetings())) = Mid(strData, 2)
            ElseIf Left(strData, 1) = "B" Then
                Entries = Entries + 1
                ReDim Preserve AIResponses(Entries) As AIQuestion
                AIResponses(Entries).Question = Mid(strData, 2)
                ReDim AIResponses(Entries).Answers(0) As String
            ElseIf Left(strData, 1) = "C" Then
                ReDim Preserve AIResponses(Entries).Answers(UBound(AIResponses(Entries).Answers()) + 1) As String
                AIResponses(Entries).Answers(UBound(AIResponses(Entries).Answers())) = Mid(strData, 2)
            ElseIf Left(strData, 1) = "D" Then
                Entries = Entries + 1
                ReDim Preserve AIResponses(Entries) As AIQuestion
                AIResponses(Entries).Question = "DD~" & Mid(strData, 2)
                ReDim AIResponses(Entries).Answers(0) As String
            ElseIf Left(strData, 1) = "E" Then
                Entries = Entries + 1
                ReDim Preserve AIResponses(Entries) As AIQuestion
                AIResponses(Entries).Question = "EE~" & Mid(strData, 2)
                ReDim AIResponses(Entries).Answers(0) As String
            End If
        Loop Until EOF(1)
    End If
    Close
End Sub

Sub AddToAI(Question, Answer)
    If Right(Question, 1) = "?" Then
        Question = Mid(Question, 1, Len(Question) - 1)
    End If
    If UBound(AIResponses()) = 0 Then
        'No Entries
        Entries = Entries + 1
        ReDim AIResponses(Entries) As AIQuestion
        AIResponses(Entries).Question = Question
        ReDim AIResponses(Entries).Answers(1) As String
        AIResponses(Entries).Answers(1) = Answer
    Else
        Dim Found As Boolean
        Dim Found1 As Boolean
        Dim X As Integer
        Dim Y As Integer
        For X = 1 To UBound(AIResponses())
            If UCase$(AIResponses(X).Question) = UCase$(Question) Then
                Found = True
                Found1 = False
                For Y = 1 To UBound(AIResponses(X).Answers())
                    If UCase$(AIResponses(X).Answers(Y)) = UCase$(Answer) Then
                        Found1 = True
                    End If
                Next
                If Found1 = False Then
                    ReDim Preserve AIResponses(X).Answers(UBound(AIResponses(X).Answers()) + 1) As String
                    AIResponses(X).Answers(UBound(AIResponses(X).Answers())) = Answer
                End If
            End If
        Next
        If Found = False Then
            Entries = Entries + 1
            ReDim Preserve AIResponses(Entries) As AIQuestion
            AIResponses(Entries).Question = Question
            ReDim Preserve AIResponses(Entries).Answers(1) As String
            AIResponses(Entries).Answers(1) = Answer
        End If
    End If
End Sub

Function GetAIResponse(Question)
    Dim AskQ As String
    Dim IncWords() As String
    Dim X As Integer
    Dim Y As Integer
    Dim Found As Boolean
    
    If Right$(Question, 1) = "?" Then
        Question = Mid$(Question, 1, Len(Question) - 1)
    End If
    Question = Trim$(UCase$(Question))
    If UBound(AIResponses()) <> 0 Then
        For X = 1 To UBound(AIResponses())
            AskQ = UCase$(AIResponses(X).Question)
            If Right$(AIResponses(X).Question, 1) = "?" Then
                AskQ = Mid(AskQ, 1, Len(AskQ) - 1)
            End If
            
            If Left$(AskQ, 3) = "DD~" Then
                AskQ = Mid$(AskQ, 4)
                IncWords() = Split(AskQ, " ")
                Found = False
                Question = " " & Question
                For Y = 1 To UBound(IncWords())
                    If InStr(1, Question, IncWords(Y)) = 0 Then
                        Found = True
                        Exit For
                    End If
                Next
                If Found = False Then
                    If UBound(AIResponses(X).Answers()) = 1 Then
                        GetAIResponse = AIResponses(X).Answers(1)
                    Else
                        GetAIResponse = AIResponses(X).Answers(Rnd * UBound(AIResponses(X).Answers()) + 1)
                    End If
                    Exit Function
                End If
            ElseIf Left$(AskQ, 3) = "EE~" Then
                AskQ = Mid$(AskQ, 4)
                IncWords() = Split(AskQ, "||")
                Found = False
                For Y = 1 To UBound(IncWords())
                    If Right(IncWords(Y), 1) = "?" Then
                        If Question = Mid(IncWords(Y), 1, Len(IncWords(Y)) - 1) Then
                            Found = True
                            Exit For
                        End If
                    Else
                        If Question = IncWords(Y) Then
                            Found = True
                        End If
                    End If
                Next
                'Debug.Print AskQ
                If Found = True Then
                    If UBound(AIResponses(X).Answers()) = 1 Then
                        GetAIResponse = AIResponses(X).Answers(1)
                    Else
                        GetAIResponse = AIResponses(X).Answers(Rnd * UBound(AIResponses(X).Answers()) + 1)
                    End If
                    Exit Function
                End If
            Else
                If AskQ = Question Then
                    If UBound(AIResponses(X).Answers()) = 1 Then
                        GetAIResponse = AIResponses(X).Answers(1)
                    Else
                        GetAIResponse = AIResponses(X).Answers(Rnd * UBound(AIResponses(X).Answers()) + 1)
                    End If
                    Exit Function
                End If
            End If
        Next
    Else
        GetAIResponse = ""
    End If
End Function

Sub SaveAI()
    Open App.Path & "\data\AI.txt" For Output As #1
        If UBound(Greetings()) <> 0 Then
            For X = 1 To UBound(Greetings())
                Print #1, "A" & Greetings(X)
            Next
        End If
        Print #1, " "
        If UBound(AIResponses()) <> 0 Then
            For X = 1 To UBound(AIResponses())
                If Left$(AIResponses(X).Question, 3) = "DD~" Then
                    Print #1, "D" & Mid(AIResponses(X).Question, 4)
                ElseIf Left$(AIResponses(X).Question, 3) = "EE~" Then
                    Print #1, "E" & Mid(AIResponses(X).Question, 4)
                Else
                    Print #1, "B" & AIResponses(X).Question
                End If
                For Y = 1 To UBound(AIResponses(X).Answers())
                    Print #1, "C" & AIResponses(X).Answers(Y)
                Next
                Print #1, " "
            Next
        End If
    Close
End Sub

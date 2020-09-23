Attribute VB_Name = "ModForum"
Public Type Forum
    Title As String
    Posts() As String
End Type

Public Forums() As Forum
Dim FEnt As Integer

Sub LoadForum()
    ReDim Forums(0) As Forum
    FEnt = 0
    If Dir(App.Path & "\data\Forum.txt") <> "Forum.txt" Then
        Exit Sub
    End If
    Dim StrData As String
    Open App.Path & "\data\Forum.txt" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, StrData
            If Left(StrData, 1) = "A" Then
                FEnt = FEnt + 1
                ReDim Preserve Forums(FEnt) As Forum
                Forums(FEnt).Title = Mid(StrData, 2)
                ReDim Forums(FEnt).Posts(0) As String
            ElseIf Left(StrData, 1) = "B" Then
                ReDim Preserve Forums(FEnt).Posts(UBound(Forums(FEnt).Posts()) + 1) As String
                Forums(FEnt).Posts(UBound(Forums(FEnt).Posts())) = Mid(StrData, 2)
            End If
        Loop Until EOF(1)
    End If
    Close
End Sub

Sub AddFTopic(TopicName, Post)
    FEnt = FEnt + 1
    ReDim Preserve Forums(FEnt) As Forum
    Forums(FEnt).Title = TopicName
    ReDim Forums(FEnt).Posts(1) As String
    Forums(FEnt).Posts(1) = Post
    Call SaveForum
End Sub

Sub AddFPost(TopicNo, Post)
    ReDim Preserve Forums(TopicNo).Posts(UBound(Forums(TopicNo).Posts()) + 1) As String
    Forums(TopicNo).Posts(UBound(Forums(TopicNo).Posts())) = Post
    Call SaveForum
End Sub

Sub SaveForum()
    Open App.Path & "\data\Forum.txt" For Output As #1
        If UBound(Forums()) <> 0 Then
            For X = 1 To UBound(Forums())
                Print #1, "A" & Forums(X).Title
                For Y = 1 To UBound(Forums(X).Posts())
                    Print #1, "B" & Forums(X).Posts(Y)
                Next
            Next
        End If
    Close
End Sub

Function ReturnTopics()
    Dim StrData As String
    For X = 1 To FEnt
        StrData = StrData & vbNewLine & X & "." & Forums(X).Title
    Next
    ReturnTopics = StrData
End Function

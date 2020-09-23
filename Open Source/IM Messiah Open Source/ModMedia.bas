Attribute VB_Name = "ModMedia"
Public Type Review
    Title As String
    Posts() As String
End Type

Public Reviews() As Review
Dim REnt As Integer

Public TrailerName() As String
Public TrailerLink() As String

Sub LoadReviews()
    ReDim Reviews(0) As Review
    REnt = 0
    If Dir(App.Path & "\data\Reviews.txt") <> "Reviews.txt" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\Reviews.txt" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            If Left(strData, 1) = "A" Then
                REnt = REnt + 1
                ReDim Preserve Reviews(REnt) As Review
                Reviews(REnt).Title = Mid(strData, 2)
                ReDim Reviews(REnt).Posts(0) As String
            ElseIf Left(strData, 1) = "B" Then
                ReDim Preserve Reviews(REnt).Posts(UBound(Reviews(REnt).Posts()) + 1) As String
                Reviews(REnt).Posts(UBound(Reviews(REnt).Posts())) = Mid(strData, 2)
            End If
        Loop Until EOF(1)
    End If
    Close
End Sub

Sub AddFilmT(TopicName, Post)
    REnt = REnt + 1
    ReDim Preserve Reviews(REnt) As Review
    Reviews(REnt).Title = TopicName
    ReDim Reviews(REnt).Posts(1) As String
    Reviews(REnt).Posts(1) = Post
    Call SaveReviews
End Sub

Sub AddFilmR(TopicNo, Post)
    ReDim Preserve Reviews(TopicNo).Posts(UBound(Reviews(TopicNo).Posts()) + 1) As String
    Reviews(TopicNo).Posts(UBound(Reviews(TopicNo).Posts())) = Post
    Call SaveReviews
End Sub

Sub SaveReviews()
    Open App.Path & "\data\Reviews.txt" For Output As #1
        If UBound(Reviews()) <> 0 Then
            For x = 1 To UBound(Reviews())
                Print #1, "A" & Reviews(x).Title
                For Y = 1 To UBound(Reviews(x).Posts())
                    Print #1, "B" & Reviews(x).Posts(Y)
                Next
            Next
        End If
    Close
End Sub

Function ReturnReviews()
    Dim strData As String
    For x = 1 To REnt
        strData = strData & vbNewLine & x & "." & Reviews(x).Title
    Next
    ReturnReviews = strData
End Function

Sub SaveTrailors(strData)
    ReDim TrailerName(0) As String
    ReDim TrailerLink(0) As String
    strData = Replace(strData, "&amp;", "&")
    Do
        StartP = InStr(1, strData, "<a href=")
        If StartP <> 0 Then
            ReDim Preserve TrailerName(UBound(TrailerName()) + 1) As String
            ReDim Preserve TrailerLink(UBound(TrailerLink()) + 1) As String
            StartP = StartP + Len("<a href=" & """")
            EndP = InStr(StartP, strData, ">") - 1
            FilmLink = Mid(strData, StartP, EndP - StartP)
            TrailerLink(UBound(TrailerLink())) = "http://www.apple.com/trailers/" & FilmLink
            StartP = EndP + 2
            EndP = InStr(StartP, strData, "<")
            FilmName = Mid(strData, StartP, EndP - StartP)
            TrailerName(UBound(TrailerName())) = FilmName
            strData = Mid(strData, EndP + 1)
        End If
    Loop Until InStr(1, strData, "<a href") = 0
End Sub

Function GetTrailors()
    Dim strData As String
    For x = 1 To UBound(TrailerName())
        strData = strData & vbNewLine & x & ": " & TrailerName(x)
    Next
    GetTrailors = strData
End Function

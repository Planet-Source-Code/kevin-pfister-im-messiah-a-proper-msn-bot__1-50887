Attribute VB_Name = "ModMain"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Integer


Public Const Capacity = 50

Public Admins() As String
Public Users() As String
Public Quotes() As String
Public WhoSaid() As String
Public BanList() As String
Public Suggestions() As String
Public Jokes() As String
Public JokesA() As String
Public FunnySites() As String
Public Nicks() As String
Public OldNews() As String
Public IMUsers() As String

Public BotName() As String
Public BotAddr() As String
Public BotDesc() As String

Public HangmanWords() As String

Public CurNews As String
Public StartTime As Long

Public Convos As Integer
Public OpenConvos As Integer
Public MsgOut As Integer
Public MsgIn As Integer

Public NewContacts As Integer
Public Announce As String

Public RegUser As Integer

Public IDEDebug As Boolean

Public BotUpdates As String

Public EditorsNote As String

Public ChatOn As Boolean
Public ChatNo As Integer

Public MeWannaChat As Boolean

Public UsedMe() As String

Public Everyone() As String

Sub SaveBotUpdates(Updates)
    Open App.Path & "\data\BotUpdates.txt" For Output As #1
    Print #1, Updates
    Close
End Sub

Sub UpdateStats()
    'Doesn't work in the opensource Bot
    
    'This Sub would have sent details to my webserver to be shown on the stats page
End Sub

Function WebAdmins()
    Dim strData As String
    If UBound(Admins()) = 0 Then
        strData = "    <p style=" & """" & "margin-top: 0; margin-bottom: 0" & """" & "><font size=" & """" & "2" & """" & ">There is currently no Admins</font></p>"
    Else
        For X = 1 To UBound(Admins())
            strData = strData & "    <p style=" & """" & "margin-top: 0; margin-bottom: 0" & """" & "><font size=" & """" & "2" & """" & ">" & PrefName(Admins(X)) & "</font></p>"
        Next
    End If
End Function

Function WebMods()
    Dim strData As String
    If UBound(Users()) = 0 Then
        strData = "    <p style=" & """" & "margin-top: 0; margin-bottom: 0" & """" & "><font size=" & """" & "2" & """" & ">There is currently no Mods</font></p>"
    Else
        For X = 1 To UBound(Users())
            strData = strData & "    <p style=" & """" & "margin-top: 0; margin-bottom: 0" & """" & "><font size=" & """" & "2" & """" & ">" & PrefName(Users(X)) & "</font></p>"
        Next
    End If
End Function

Public Sub LoadEdNote()
    If Dir(App.Path & "\data\EditorNote.txt") <> "EditorNote.txt" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\EditorNote.txt" For Input As #1
    If LOF(1) <> 0 Then
        strData = Input(LOF(1), 1)
        EditorsNote = strData
    End If
    Close
End Sub

Public Sub SaveEdNote()
    Open App.Path & "\data\EditorNote.txt" For Output As #1
    Print #1, EditorsNote
    Close
End Sub

Public Sub LoadUserBots()
    ReDim BotName(0) As String
    ReDim BotAddr(0) As String
    ReDim BotDesc(0) As String
    If Dir(App.Path & "\data\UserBots.dat") <> "UserBots.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\UserBots.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve BotName(UBound(BotName()) + 1) As String
            BotName(UBound(BotName())) = strData
            
            Line Input #1, strData
            ReDim Preserve BotAddr(UBound(BotAddr()) + 1) As String
            BotAddr(UBound(BotAddr())) = strData
            
            Line Input #1, strData
            ReDim Preserve BotDesc(UBound(BotDesc()) + 1) As String
            BotDesc(UBound(BotDesc())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub AddUserBots(Name, Adr, Desc)
    ReDim Preserve BotName(UBound(BotName()) + 1) As String
    BotName(UBound(BotName())) = Name
    
    ReDim Preserve BotAddr(UBound(BotAddr()) + 1) As String
    BotAddr(UBound(BotAddr())) = Adr

    ReDim Preserve BotDesc(UBound(BotDesc()) + 1) As String
    BotDesc(UBound(BotDesc())) = Desc

    Open App.Path & "\data\UserBots.dat" For Append As #1
        Print #1, Name
        Print #1, Adr
        Print #1, Desc
    Close
End Sub

Public Sub LoadHangman()
    ReDim HangmanWords(0) As String
    If Dir(App.Path & "\data\HangmanWords.dat") <> "HangmanWords.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\HangmanWords.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve HangmanWords(UBound(HangmanWords()) + 1) As String
            HangmanWords(UBound(HangmanWords())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub ClearLog()
    Open App.Path & "\data\Log.txt" For Output As #1
    Close
End Sub

Public Sub AddToLog(Text)
    Open App.Path & "\data\Log.txt" For Append As #1
    Print #1, ""
    Print #1, Time & " : " & Text
    Close
End Sub

Public Function ViewLog()
    If Dir(App.Path & "\data\Log.txt") <> "Log.txt" Then
        Exit Function
    End If
    Dim strData As String
    Open App.Path & "\data\Log.txt" For Input As #1
    If LOF(1) <> 0 Then
        strData = Input(LOF(1), 1)
        ViewLog = strData
    End If
    Close
End Function

Public Sub ApplyMod(Reason, Name)
    Open App.Path & "\data\Apply.txt" For Append As #1
    Print #1, ""
    Print #1, Time
    Print #1, "---------------------------"
    Print #1, Reason
    Print #1, Name
    Close
End Sub


Public Sub loadBotUD()
    If Dir(App.Path & "\data\BotUpdates.txt") <> "BotUpdates.txt" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\BotUpdates.txt" For Input As #1
    If LOF(1) <> 0 Then
        BotUpdates = Input(LOF(1), 1)
    End If
    Close
End Sub

Public Sub SayToAll(Text)
    Announce = Text
    frmMain.TellAll
End Sub

Public Sub AddNews(News)
    Dim Exists As Boolean
    Exists = False
    For Z = 1 To UBound(OldNews())
        If OldNews(Z) = News Then
            Exists = True
        End If
    Next
    If Exists = False Then
        ReDim Preserve OldNews(UBound(OldNews()) + 1) As String
        OldNews(UBound(OldNews())) = News
        Open App.Path & "\data\OldNews.dat" For Output As #1
        For Z = 1 To UBound(OldNews())
            Print #1, "<>" & OldNews(Z);
        Next
        Close
    End If
End Sub

Public Sub loadNews()
    ReDim OldNews(0) As String
    If Dir(App.Path & "\data\OldNews.dat") <> "OldNews.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\OldNews.dat" For Input As #1
    If LOF(1) <> 0 Then
        strData = Input(LOF(1), 1)
        OldNews() = Split(strData, "<>")
    End If
    Close
End Sub

Public Sub loadGenSets()
    IDEDebug = GetSetting("IM Messiah", "General", "IDEDEBUG", True)
    ReDim Everyone(0) As String
End Sub

Public Sub LoadNicks()
    ReDim Nicks(0) As String
    If Dir(App.Path & "\data\Nicks.dat") <> "Nicks.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\Nicks.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve Nicks(UBound(Nicks()) + 1) As String
            Nicks(UBound(Nicks())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub AddNick(Nick)
    Dim Exists As Boolean
    Exists = False
    For Z = 1 To UBound(Nicks())
        If Nicks(Z) = Nick Then
            Exists = True
        End If
    Next
    If Exists = False Then
        ReDim Preserve Nicks(UBound(Nicks()) + 1) As String
        Nicks(UBound(Nicks())) = Nick
        Open App.Path & "\data\Nicks.dat" For Output As #1
        For Z = 1 To UBound(Nicks())
            Print #1, Nicks(Z)
        Next
        Close
    End If
End Sub

Public Function RandomNick() As String
    If UBound(Nicks()) <> 0 Then
        Num = Int(Rnd * UBound(Nicks())) + 1
        RandomNick = Nicks(Num)
    End If
End Function

Public Sub LoadSites()
    ReDim FunnySites(0) As String
    If Dir(App.Path & "\data\FunnySites.dat") <> "FunnySites.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\FunnySites.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve FunnySites(UBound(FunnySites()) + 1) As String
            FunnySites(UBound(FunnySites())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub AddSite(Site)
    Dim Exists As Boolean
    Exists = False
    For Z = 1 To UBound(FunnySites())
        If FunnySites(Z) = Site Then
            Exists = True
        End If
    Next
    If Exists = False Then
        ReDim Preserve FunnySites(UBound(FunnySites()) + 1) As String
        FunnySites(UBound(FunnySites())) = Site
        Open App.Path & "\data\FunnySites.dat" For Output As #1
        For Z = 1 To UBound(FunnySites())
            Print #1, FunnySites(Z)
        Next
        Close
    End If
End Sub

Public Sub LoadJokes()
    ReDim Jokes(0) As String
    ReDim JokesA(0) As String
    If Dir(App.Path & "\data\Jokes.dat") <> "Jokes.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\Jokes.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve Jokes(UBound(Jokes()) + 1) As String
            Jokes(UBound(Jokes())) = strData
            Line Input #1, strData
            ReDim Preserve JokesA(UBound(JokesA()) + 1) As String
            JokesA(UBound(JokesA())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub AddJoke(Joke, Answer)
    Dim Exists As Boolean
    Exists = False
    For Z = 1 To UBound(Jokes())
        If Jokes(Z) = Joke Then
            Exists = True
        End If
    Next
    If Exists = False Then
        ReDim Preserve Jokes(UBound(Jokes()) + 1) As String
        Jokes(UBound(Jokes())) = Joke
        ReDim Preserve JokesA(UBound(JokesA()) + 1) As String
        JokesA(UBound(JokesA())) = Answer
        Open App.Path & "\data\Jokes.dat" For Output As #1
        For Z = 1 To UBound(Jokes())
            Print #1, Jokes(Z)
            Print #1, JokesA(Z)
        Next
        Close
    End If
End Sub

Public Function RandomJoke() As String
    If UBound(Jokes()) <> 0 Then
        Num = Int(Rnd * UBound(Jokes())) + 1
        RandomJoke = Jokes(Num) & vbNewLine & vbNewLine & "..." & JokesA(Num)
    End If
End Function

Public Function ViewJoke(Index) As String
    ViewJoke = Jokes(Index) & vbNewLine & vbNewLine & "..." & JokesA(Index)
End Function

Public Sub StartupVars()
    ReDim IMUsers(0) As String
    ReDim UsedMe(0) As String
    StartTime = Timer
    MeWannaChat = True
End Sub

Public Sub AddBotUser(Email)
    If UBound(UsedMe()) = 0 Then
        ReDim Preserve UsedMe(UBound(UsedMe()) + 1) As String
        UsedMe(UBound(UsedMe())) = Email
    Else
        Dim Exists As Boolean
        Exists = False
        For Z = 1 To UBound(UsedMe())
            If UCase(UsedMe(Z)) = UCase(Email) Then
                Exists = True
            End If
        Next
        If Exists = False Then
            ReDim Preserve UsedMe(UBound(UsedMe()) + 1) As String
            UsedMe(UBound(UsedMe())) = Email
        End If
    End If
End Sub

Public Sub ClearSuggestions()
    ReDim Suggestions(0) As String
    Open App.Path & "\data\Suggestions.dat" For Output As #1
    Close
End Sub

Public Sub AddSuggestion(Suggestion)
    Dim Exists As Boolean
    Exists = False
    For Z = 1 To UBound(Suggestions())
        If Suggestions(Z) = Suggestion Then
            Exists = True
        End If
    Next
    If Exists = False Then
        ReDim Preserve Suggestions(UBound(Suggestions()) + 1) As String
        Suggestions(UBound(Suggestions())) = Suggestion
        Open App.Path & "\data\Suggestions.dat" For Output As #1
        For Z = 1 To UBound(Suggestions())
            Print #1, Suggestions(Z)
        Next
        Close
    End If
End Sub

Public Sub LoadSuggestions()
    ReDim Suggestions(0) As String
    If Dir(App.Path & "\data\Suggestions.dat") <> "Suggestions.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\Suggestions.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve Suggestions(UBound(Suggestions()) + 1) As String
            Suggestions(UBound(Suggestions())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub LoadQuotes()
    ReDim Quotes(0) As String
    ReDim WhoSaid(0) As String
    If Dir(App.Path & "\data\Quotes.dat") <> "Quotes.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\Quotes.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve Quotes(UBound(Quotes()) + 1) As String
            Quotes(UBound(Quotes())) = strData
            Line Input #1, strData
            ReDim Preserve WhoSaid(UBound(WhoSaid()) + 1) As String
            WhoSaid(UBound(WhoSaid())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub AddQuote(Name, Quote)
    Dim Exists As Boolean
    Exists = False
    For Z = 1 To UBound(Quotes())
        If Quotes(Z) = Quote Then
            Exists = True
        End If
    Next
    If Exists = False Then
        ReDim Preserve Quotes(UBound(Quotes()) + 1) As String
        Quotes(UBound(Quotes())) = Quote
        ReDim Preserve WhoSaid(UBound(WhoSaid()) + 1) As String
        WhoSaid(UBound(WhoSaid())) = Name
        Open App.Path & "\data\Quotes.dat" For Output As #1
        For Z = 1 To UBound(Quotes())
            Print #1, Quotes(Z)
            Print #1, WhoSaid(Z)
        Next
        Close
    End If
End Sub

Public Function RandomQuote() As String
    If UBound(Quotes()) <> 0 Then
        Num = Int(Rnd * UBound(Quotes())) + 1
        RandomQuote = Quotes(Num) & vbNewLine & "Said by" & vbNewLine & WhoSaid(Num)
    End If
End Function

Public Function ViewQuote(Index) As String
    ViewQuote = Quotes(Index) & vbNewLine & "Said by" & vbNewLine & WhoSaid(Index)
End Function

Public Sub LoadAdmins()
    ReDim Admins(0) As String
    If Dir(App.Path & "\data\Admins.dat") <> "Admins.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\Admins.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve Admins(UBound(Admins()) + 1) As String
            Admins(UBound(Admins())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub LoadUsers()
    ReDim Users(0) As String
    If Dir(App.Path & "\data\Users.dat") <> "Users.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\Users.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve Users(UBound(Users()) + 1) As String
            Users(UBound(Users())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub LoadBans()
    ReDim BanList(0) As String
    If Dir(App.Path & "\data\BanList.dat") <> "BanList.dat" Then
        Exit Sub
    End If
    Dim strData As String
    Open App.Path & "\data\BanList.dat" For Input As #1
    If LOF(1) <> 0 Then
        Do
            Line Input #1, strData
            ReDim Preserve BanList(UBound(BanList()) + 1) As String
            BanList(UBound(BanList())) = strData
        Loop Until EOF(1)
    End If
    Close
End Sub

Public Sub AddBan(Name)
    Dim Exists As Boolean
    Exists = False
    For Z = 1 To UBound(BanList())
        If BanList(Z) = Name Then
            Exists = True
        End If
    Next
    If Exists = False Then
        ReDim Preserve BanList(UBound(BanList()) + 1) As String
        BanList(UBound(BanList())) = Name
        Open App.Path & "\data\BanList.dat" For Output As #1
        For Z = 1 To UBound(BanList())
            Print #1, BanList(Z)
        Next
        Close
    End If
End Sub

Public Sub DeleteBan(Name)
    Dim temp() As String
    ReDim temp(0) As String
    If UBound(BanList()) <> 0 Then
        For Z = 1 To UBound(BanList())
            If BanList(Z) <> Name Then
                ReDim Preserve temp(UBound(temp()) + 1) As String
                temp(UBound(temp())) = BanList(Z)
            End If
        Next
    End If
    ReDim BanList(UBound(temp())) As String
    If UBound(temp()) <> 0 Then
        For Z = 1 To UBound(temp())
            BanList(Z) = temp(Z)
        Next
    End If
    Open App.Path & "\data\BanList.dat" For Output As #1
    For Z = 1 To UBound(BanList())
        Print #1, BanList(Z)
    Next
    Close
End Sub

Public Function IsBanned(Name) As Boolean
    Name = UCase(Name)
    If UBound(BanList()) = 0 Then
        IsBanned = False
        Exit Function
    End If
    Dim Found As Boolean
    Found = False
    For Z = 1 To UBound(BanList())
        If Name = UCase(BanList(Z)) Then
            Found = True
        End If
    Next
    IsBanned = Found
End Function

Public Sub AddUser(Name)
    Dim Exists As Boolean
    Exists = False
    For Z = 1 To UBound(Users())
        If Users(Z) = Name Then
            Exists = True
        End If
    Next
    If Exists = False Then
        ReDim Preserve Users(UBound(Users()) + 1) As String
        Users(UBound(Users())) = Name
        Open App.Path & "\data\Users.dat" For Output As #1
        For Z = 1 To UBound(Users())
            Print #1, Users(Z)
        Next
        Close
    End If
End Sub

Public Sub DeleteUser(Name)
    Dim temp() As String
    ReDim temp(0) As String
    If UBound(Users()) <> 0 Then
        For Z = 1 To UBound(Users())
            If Users(Z) <> Name Then
                ReDim Preserve temp(UBound(temp()) + 1) As String
                temp(UBound(temp())) = Users(Z)
            End If
        Next
    End If
    ReDim Users(UBound(temp())) As String
    If UBound(temp()) <> 0 Then
        For Z = 1 To UBound(temp())
            Users(Z) = temp(Z)
        Next
    End If
    Open App.Path & "\data\Users.dat" For Output As #1
    For Z = 1 To UBound(Users())
        Print #1, Users(Z)
    Next
    Close
End Sub

Public Sub AddAdmin(Name)
    Dim Exists As Boolean
    Exists = False
    For Z = 1 To UBound(Admins())
        If Admins(Z) = Name Then
            Exists = True
        End If
    Next
    If Exists = False Then
        ReDim Preserve Admins(UBound(Admins()) + 1) As String
        Admins(UBound(Admins())) = Name
        Open App.Path & "\data\Admins.dat" For Output As #1
        For Z = 1 To UBound(Admins())
            Print #1, Admins(Z)
        Next
        Close
    End If
End Sub

Public Function IsAdmin(Name) As Boolean
    Name = UCase(Name)
    If UBound(Admins()) = 0 Then
        IsAdmin = False
        Exit Function
    End If
    Dim Found As Boolean
    Found = False
    For Z = 1 To UBound(Admins())
        If Name = UCase(Admins(Z)) Then
            Found = True
        End If
    Next
    IsAdmin = Found
End Function

Public Function IsUser(Name) As Boolean
    Name = UCase(Name)
    If UBound(Users()) = 0 Then
        IsUser = False
        Exit Function
    End If
    Dim Found As Boolean
    Found = False
    For Z = 1 To UBound(Users())
        If Name = UCase(Users(Z)) Then
            Found = True
        End If
    Next
    IsUser = Found
End Function

Public Sub ContactUs(Text, Name)
    Open App.Path & "\data\ContactUs.txt" For Append As #1
    Print #1, ""
    Print #1, "---------------------------"
    Print #1, "Time:"
    Print #1, Text
    Print #1, Name
    Print #1, "---------------------------"
    Close
End Sub

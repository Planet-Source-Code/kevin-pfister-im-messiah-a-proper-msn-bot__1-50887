VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMSG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instant Message"
   ClientHeight    =   3945
   ClientLeft      =   3090
   ClientTop       =   3090
   ClientWidth     =   6000
   Icon            =   "frmMSG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   6000
   Begin VB.CommandButton cmdsend 
      Caption         =   "Send"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox txtHist 
      Height          =   3045
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   5371
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMSG.frx":0CCA
   End
   Begin VB.TextBox txtMSG 
      Height          =   405
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3240
      Width           =   4380
   End
   Begin MSWinsockLib.Winsock sckMSG 
      Left            =   1080
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1863
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5775
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myTrID1 As Integer
Dim strHeader As String
Dim myname As String
Public strSID, strMIP, strCKI As String
Public buddy As String
Public clientside As Integer 'whether or not the client initiated convo
Dim buddyfont As String, buddycolor As String, buddystyle As String

Dim ReturnVar As Integer

Dim QuoteSaid As String
Dim JokeLine As String
Dim HiLowNo As Integer

Public InConvo As Boolean

Dim OthelloGrid(1 To 8, 1 To 8)         'This is the Grid on which the calcualations are done
Dim SelectableSquares(1 To 8, 1 To 8)   'Used to hold possible squares to move to
Dim checking(1 To 8, 1 To 8)            'Used to hold possible squares to move to
Dim ahead(1 To 8, 1 To 8)               'Used to hold possible user/computer moves
Dim SquarePotent As Long                'This holds the Number of counters placed
Dim GoodShot As Long                    'To check if a shot is valid
Dim Win As Long                         'To check if the shot was successful

Dim Wipeout(1 To 3, 1 To 3) As Boolean
Dim WipeMoves As Integer

Dim AIon As Boolean
Public HearAnnounce As Boolean

Dim HWord As String
Dim Tried() As String

Dim AIAdding As String
Dim TopName As String
Dim CurTopic As Integer
Dim CurFilm As Integer
Dim FilmName As String
Dim SckSend As Integer

Dim AddBot1 As String
Dim AddBot2 As String

Dim UserLang As Integer

Public ImChat As Boolean

Sub AddMeToChat(Name)
    sndstr "CAL " & myTrID1 & " " & Name & vbCrLf
    frmMain.IncrementTrailID
    DoEvents
    SendMSG "Welcome " & PrefGetName(Name) & " to the chat room", 7
End Sub

Function DrawOthello() As String
    Dim strData As String
    strData = OthelloSquare(1, 1) & "|" & OthelloSquare(2, 1) & "|" & OthelloSquare(3, 1) & "|" & OthelloSquare(4, 1) & "|" & OthelloSquare(5, 1) & "|" & OthelloSquare(6, 1) & "|" & OthelloSquare(7, 1) & "|" & OthelloSquare(8, 1)
    strData = strData & vbNewLine & "-+-+-+-+-+-+-+-"
    strData = strData & vbNewLine & OthelloSquare(1, 2) & "|" & OthelloSquare(2, 2) & "|" & OthelloSquare(3, 2) & "|" & OthelloSquare(4, 2) & "|" & OthelloSquare(5, 2) & "|" & OthelloSquare(6, 2) & "|" & OthelloSquare(7, 2) & "|" & OthelloSquare(8, 2)
    strData = strData & vbNewLine & "-+-+-+-+-+-+-+-"
    strData = strData & vbNewLine & OthelloSquare(1, 3) & "|" & OthelloSquare(2, 3) & "|" & OthelloSquare(3, 3) & "|" & OthelloSquare(4, 3) & "|" & OthelloSquare(5, 3) & "|" & OthelloSquare(6, 3) & "|" & OthelloSquare(7, 3) & "|" & OthelloSquare(8, 3)
    strData = strData & vbNewLine & "-+-+-+-+-+-+-+-"
    strData = strData & vbNewLine & OthelloSquare(1, 4) & "|" & OthelloSquare(2, 4) & "|" & OthelloSquare(3, 4) & "|" & OthelloSquare(4, 4) & "|" & OthelloSquare(5, 4) & "|" & OthelloSquare(6, 4) & "|" & OthelloSquare(7, 4) & "|" & OthelloSquare(8, 4)
    strData = strData & vbNewLine & "-+-+-+-+-+-+-+-"
    strData = strData & vbNewLine & OthelloSquare(1, 5) & "|" & OthelloSquare(2, 5) & "|" & OthelloSquare(3, 5) & "|" & OthelloSquare(4, 5) & "|" & OthelloSquare(5, 5) & "|" & OthelloSquare(6, 5) & "|" & OthelloSquare(7, 5) & "|" & OthelloSquare(8, 5)
    strData = strData & vbNewLine & "-+-+-+-+-+-+-+-"
    strData = strData & vbNewLine & OthelloSquare(1, 6) & "|" & OthelloSquare(2, 6) & "|" & OthelloSquare(3, 6) & "|" & OthelloSquare(4, 6) & "|" & OthelloSquare(5, 6) & "|" & OthelloSquare(6, 6) & "|" & OthelloSquare(7, 6) & "|" & OthelloSquare(8, 6)
    strData = strData & vbNewLine & "-+-+-+-+-+-+-+-"
    strData = strData & vbNewLine & OthelloSquare(1, 7) & "|" & OthelloSquare(2, 7) & "|" & OthelloSquare(3, 7) & "|" & OthelloSquare(4, 7) & "|" & OthelloSquare(5, 7) & "|" & OthelloSquare(6, 7) & "|" & OthelloSquare(7, 7) & "|" & OthelloSquare(8, 7)
    strData = strData & vbNewLine & "-+-+-+-+-+-+-+-"
    strData = strData & vbNewLine & OthelloSquare(1, 8) & "|" & OthelloSquare(2, 8) & "|" & OthelloSquare(3, 8) & "|" & OthelloSquare(4, 8) & "|" & OthelloSquare(5, 8) & "|" & OthelloSquare(6, 8) & "|" & OthelloSquare(7, 8) & "|" & OthelloSquare(8, 8)
    DrawOthello = strData
End Function

Function OthelloSquare(X, Y) As String
    If OthelloGrid(X, Y) = 2 Then
        OthelloSquare = "8"
    ElseIf OthelloGrid(X, Y) = 1 Then
        OthelloSquare = "0"
    Else
        OthelloSquare = "_"
    End If
End Function

Function ShowWipeout() As String
    Dim strData As String
    strData = WipeSquare(1, 1) & "|" & WipeSquare(2, 1) & "|" & WipeSquare(3, 1)
    strData = strData & vbNewLine & "-+-+-"
    strData = strData & vbNewLine & WipeSquare(1, 2) & "|" & WipeSquare(2, 2) & "|" & WipeSquare(3, 2)
    strData = strData & vbNewLine & "-+-+-"
    strData = strData & vbNewLine & WipeSquare(1, 3) & "|" & WipeSquare(2, 3) & "|" & WipeSquare(3, 3)
    ShowWipeout = strData
End Function

Function WipeSquare(X, Y) As String
    If Wipeout(X, Y) = True Then
        WipeSquare = "#"
    Else
        WipeSquare = "0"
    End If
End Function

Private Sub cmdsend_Click()
    If sckMSG.State = sckConnected Then
        If Left(txtMSG.Text, 1) = "\" Then
            ProcessCommand UCase(Mid(txtMSG.Text, 2))
        ElseIf txtMSG.Text <> "" Then
            SendMSG txtMSG.Text
            
            greytext (myname + " says: " & vbCrLf)
            txtHist.SelText = txtMSG.Text & vbCrLf
            
            txtMSG.Text = ""
            txtMSG.SetFocus
            SendKeys "{BACKSPACE}"
        End If
    Else
        txtMSG.SetFocus
        SendKeys "{BACKSPACE}"
        frmMain.Messenger.SendData "XFR " & frmMain.intTrailid & " SB" & vbCrLf
        frmMain.buddyconnect = buddy
    End If
End Sub

Private Sub Form_Load()
    myname = frmContacts.lblName.Caption
    sckMSG.Connect strMIP
    Convos = Convos + 1
    OpenConvos = OpenConvos + 1
    InConvo = True
    HearAnnounce = True
    AIon = False
    UserLang = GetLang(buddy)
    ImChat = False
End Sub

Private Sub sndstr(Message As String)
    If sckMSG.State = 7 Then
        sckMSG.SendData Message
    End If
    myTrID
    If IDEDebug = True Then
        Debug.Print Message & vbCrLf
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    sckMSG.Close
    AIon = False
    If InConvo = True Then
        OpenConvos = OpenConvos - 1
        InConvo = False
    End If
    buddy = ""
    If ImChat = True Then
        ImChat = False
        ChatOn = False
    End If
End Sub

Private Sub sckMSG_Connect()
  If clientside = 1 Then
    sndstr "USR " & myTrID1 & " " & frmMain.txtUsername.Text & " " & strCKI & vbCrLf
  Else
    sndstr "ANS " & myTrID1 & " " & frmMain.txtUsername.Text & " " & strCKI & " " & strSID & vbCrLf
  End If
End Sub

Private Function greytext(Text As String)
    txtHist.SelStart = Len(txtHist.Text)
    txtHist.SelColor = RGB(150, 150, 150)
    txtHist.SelText = Text
    txtHist.SelStart = Len(txtHist.Text)
    txtHist.SelColor = RGB(0, 0, 0)
End Function

Private Function myTrID() As Integer
    myTrID = myTrID + 1
End Function

Private Sub sckMSG_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    sckMSG.GetData strData
    If IDEDebug = True Then
        Debug.Print strData & vbCrLf
    End If
    frmMain.txtOutput = frmMain.txtOutput & strData & vbCrLf
    Select Case Left(strData, 3)
        Case "MSG":
            Dim strMessage As String
            fnMessage strData
        Case "USR":
            If Split(strData, " ")(2) = "OK" Then
                sndstr "CAL " & myTrID1 & " " & buddy & vbCrLf
                myTrID
            End If
        Case "JOI":
            Me.Show
            greytext (buddy & " has joined the chat." & vbCrLf)
            cmdsend_Click
            Me.Caption = "Instant Message - " & buddy
        Case "BYE":
            greytext (buddy & " has left the chat." & vbCrLf)
            If ImChat = False Then
                sckMSG.Close
                If InConvo = True Then
                    OpenConvos = OpenConvos - 1
                    InConvo = False
                End If
                Unload Me
            End If
        Case "IRO":
            If InConvo = False Then
                OpenConvos = OpenConvos + 1
                InConvo = True
            End If
            AddBotUser buddy
            If IsBanned(buddy) = True Then
                If ImChat = False Then
                    SendMSG GetPhrase(UserLang, 25)
                    Unload Me
                End If
            Else
                If IsAdmin(buddy) = True Then
                    SendMSG GetPhrase(UserLang, 26) & vbNewLine & GetPhrase(UserLang, 27) & vbNewLine & GetPhrase(UserLang, 28) & vbNewLine & GetPhrase(UserLang, 29), 8
                ElseIf IsUser(buddy) = True Then
                    SendMSG GetPhrase(UserLang, 26) & vbNewLine & GetPhrase(UserLang, 27) & vbNewLine & GetPhrase(UserLang, 28) & vbNewLine & GetPhrase(UserLang, 30), 1
                Else
                    SendMSG GetPhrase(UserLang, 26) & vbNewLine & GetPhrase(UserLang, 27) & vbNewLine & GetPhrase(UserLang, 28) & vbNewLine & GetPhrase(UserLang, 31), 7
                End If
                If ImChat = False Then
                    If EditorsNote <> "" Then
                        SendMSG "Welcome to IM messiah" & vbNewLine & EditorsNote
                    End If
                End If
                UserLang = GetLang(buddy)
            End If
    End Select
End Sub

Private Function decodefont(fontline As String)
'X-MMS-IM-Format: FN=Microsoft%20Sans%20Serif; EF=; CO=ff; CS=0; PF=22
'buddyfont As String, buddycolor As String, buddystyle As String

    If UBound(Split(fontline, " ")) = 5 Then 'Making sure it contains all of the elements of a font string
        buddycolor = Split(fontline, " ")(3)
        buddycolor = Right(buddycolor, Len(buddycolor) - 3)
        buddycolor = Left(buddycolor, Len(buddycolor) - 1)
        txtHist.SelColor = bgrhex2rgb(buddycolor)
        
        buddystyle = Split(fontline, " ")(2)
        buddystyle = Right(buddystyle, Len(buddystyle) - 3)
        buddystyle = Left(buddystyle, Len(buddystyle) - 1)
        If InStr(buddystyle, "B") > 0 Then txtHist.SelBold = True Else: txtHist.SelBold = False
        If InStr(buddystyle, "I") > 0 Then txtHist.SelItalic = True Else: txtHist.SelItalic = False
        If InStr(buddystyle, "S") > 0 Then txtHist.SelStrikeThru = True Else: txtHist.SelStrikeThru = False
        If InStr(buddystyle, "U") > 0 Then txtHist.SelUnderline = True Else: txtHist.SelUnderline = False
        
        buddyfont = Split(fontline, " ")(1)
        buddyfont = Right(buddyfont, Len(buddyfont) - 3)
        buddyfont = Left(buddyfont, Len(buddyfont) - 1)
        txtHist.SelFontName = URLDecode(buddyfont)
        
    End If
End Function

Private Function fnMessage(Data As String)
    On Error Resume Next
    Dim strUname As String
    Dim strFname As String
    Dim Message As String
    Dim styleline As String
    
    strUname = Split(Data, " ")(1)
    strFname = Split(Data, " ")(2)
    strFname = URLDecode(strFname)
    
    If Split(Split(Data, vbCrLf)(3), " ")(0) = "TypingUser:" Then
        lblInfo.Caption = buddy & " is typing a message..."
    End If

    
    If Split(Data, vbCrLf)(5) = "" Then
        If UBound(Split(Data, vbCrLf)) > 5 Then Message = Replace(Split(Data, vbCrLf)(6), vbCrLf, " ")
    Else
        Message = Replace(Split(Data, vbCrLf)(5), vbCrLf, " ")
    End If
    
    Dim fndmsg As Integer
    fndmsg = InStr(1, Message, "MSG " & buddy)
    If fndmsg > 0 Then Message = Mid(Message, 1, fndmsg)
    Message = Message & " "
    
    If Message <> " " Then
        greytext (strFname + " says: " & vbCrLf)
        
        styleline = Split(Data, vbCrLf)(3)
        decodefont (styleline)
        txtHist.SelText = Message & vbCrLf
        
        lblInfo.Caption = "Last message received at " & Time & "."
        If Me.Visible = False Then
            Me.Show
        End If
        MsgIn = MsgIn + 1
        If ImChat = False Then
            Call ProcessTxt(Message)
        Else
            If Message = "\EXITCHAT" Then
                If IsAdmin(buddy) = False Then
                    SendMSG GetPhrase(UserLang, 2) 'Must be admin
                Else
                    Call frmMain.ExitChats
                    frmMain.TellMod "Chatroom Stopped By " & PrefGetName(buddy)
                End If
            End If
        End If
    End If
    
End Function

Sub ProcessTxt(Message)
    If Mid$(Message, 1, 1) = "\" Then
        Call ProcessCommand(UCase(Trim$(Mid$(Message, 2))))
    Else
        Call ProcessReturn(Trim$(Message))
    End If
End Sub

Sub ProcessCommand(Command)
    Dim Z As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Output As String
    Dim XString As String
    
    ReturnVar = 0
    Select Case Command
    
    Case "ADDMOD"
        'Add Mod
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ReturnVar = 4
            SendMSG GetPhrase(UserLang, 3) & vbNewLine & GetPhrase(UserLang, 4)
        End If
    
    Case "DELMOD"
        'Delete Mod
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ReturnVar = 5
            SendMSG GetPhrase(UserLang, 5) & vbNewLine & GetPhrase(UserLang, 6)
        End If
    
    Case "CRASH"
        'Crash User
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ReturnVar = 7
            SendMSG GetPhrase(UserLang, 7) & vbNewLine & GetPhrase(UserLang, 8)
        End If
    
    Case "FLOOD"
        'Flood User
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ReturnVar = 8
            SendMSG "Flood User:" & vbNewLine & GetPhrase(UserLang, 9)
        End If
    
    Case "POKE"
        'Poke User
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            ReturnVar = 9
            SendMSG GetPhrase(UserLang, 10) & vbNewLine & GetPhrase(UserLang, 11)
        Else
            SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
        End If
        
    Case "DUSERS"
        'Display Users
        If UBound(Users()) <> 0 Then
            Dim UseTxt As String
            For Z = 1 To UBound(Users())
                UseTxt = UseTxt & vbNewLine & Users(Z)
            Next
            SendMSG "Mods:" & UseTxt
        Else
            SendMSG "There are currently No Mods Registered"
        End If
    
    Case "DADMINS"
        'Display Admins
        If UBound(Admins()) <> 0 Then
            Dim AdTxt As String
            For Z = 1 To UBound(Admins())
                AdTxt = AdTxt & vbNewLine & Admins(Z)
            Next
            SendMSG "Admins:" & AdTxt
        Else
            SendMSG "There are currently No Admins Registered"
        End If
    
    Case "POPUP"
        'Popup Message
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ReturnVar = 10
            SendMSG "Popup Message:" & vbNewLine & GetPhrase(UserLang, 12)
        End If
    
    Case "CHGNAME"
        'Change Name
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ReturnVar = 11
            SendMSG GetPhrase(UserLang, 13) & vbNewLine & GetPhrase(UserLang, 14)
        End If
    
    Case "SOCON"
        'Number of open Convos
        SendMSG "Number of Open Converstations:" & vbNewLine & OpenConvos
    
    Case "STCON"
        'Total Number of Convos
        SendMSG "Total Number of Converstations:" & vbNewLine & Convos
    
    Case "NQUOTE1"
        'New Quote
        ReturnVar = 14
        SendMSG "New Quote:" & vbNewLine & "Please Enter Quote"
    
    Case "NQUOTE2"
        'New Quote
        ReturnVar = 15
        SendMSG "Please Enter name of who said it"
    
    Case "VQUOTE"
        'View Quotes
        If UBound(Quotes()) = 0 Then
            SendMSG "There are no quotes currently stored"
        Else
            ReturnVar = 16
            SendMSG "View Quotes:" & vbNewLine & "1 to " & UBound(Quotes())
        End If
    
    Case "RQUOTE"
        'Display random quotes
        If UBound(Quotes()) = 0 Then
            SendMSG "There are no quotes currently stored"
        Else
            SendMSG "Random Quote:" & vbNewLine & RandomQuote
        End If
    
    Case "ABOUT"
        'About
        SendMSG "(co)About: \about" & vbNewLine & "Created by Kevin Pfister" & vbNewLine & "Part of Digital Existance Programming (DEP)" & vbNewLine & "Email: Yet_Another_Idiot@Hotmail.com" & vbNewLine & "Web: www.dep.zion.me.uk" & vbNewLine & vbNewLine & "Special thanks go to:" & vbNewLine & "Louis Dron - Help with Translation" & vbNewLine & vbNewLine & "Type \Contact to send us a message" & vbNewLine & vbNewLine & "Please Donate to us to help maintain the server and to constaintly add to Messiah, just go to my site and click donate"
    
    Case "BANUSER"
        'Ban User
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ReturnVar = 17
            SendMSG "Ban User:" & vbNewLine & "Please type in the name of the user you wish to Ban"
        End If
    
    Case "UNBANUSER"
        'Unban user
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ReturnVar = 18
            SendMSG "UnBan User:" & vbNewLine & "Please type in the name of the user you wish to UnBan"
        End If
    
    Case "BANLIST"
        'Display BanList
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            If UBound(BanList()) <> 0 Then
                Dim BanTxt As String
                For Z = 1 To UBound(BanList())
                    BanTxt = BanTxt & vbNewLine & BanList(Z)
                Next
                SendMSG "Banned Users:" & BanTxt
            Else
                SendMSG GetPhrase(UserLang, 15)
            End If
        End If
    Case "ASUG"
        'Add Suggestion
        ReturnVar = 21
        SendMSG GetPhrase(UserLang, 16)
    
    Case "VSUG"
        'View Suggestions
        If UBound(Suggestions()) = 0 Then
            SendMSG GetPhrase(UserLang, 17)
        Else
            ReturnVar = 20
            SendMSG GetPhrase(UserLang, 18) & vbNewLine & "1 to " & UBound(Suggestions())
        End If
    
    Case "CSUG"
        'Clear Suggestions
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ClearSuggestions
            SendMSG GetPhrase(UserLang, 19)
        End If
    
    Case "INVITE"
        'Invite User
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            ReturnVar = 22
            SendMSG GetPhrase(UserLang, 20)
        Else
            SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
        End If
    
    Case "ANNOUNCE"
        'Announce Message
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            ReturnVar = 23
            SendMSG "Please enter the message you wish to announce:"
        Else
            SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
        End If

    Case "GENSTAT"
        'General Stats
        SendMSG "General Stats:" & vbNewLine & "Registered: " & RegUser & vbNewLine & "Admins: " & UBound(Admins()) & vbNewLine & "Mods: " & UBound(Users()) & vbNewLine & "Online for " & Int((Timer - StartTime) / 60) & " minutes" & vbNewLine & OpenConvos & " Open Converstations" & vbNewLine & Convos & " Total Converstations" & vbNewLine & "Messages In " & MsgIn & vbNewLine & "Messages Out " & MsgOut & vbNewLine & "New Contacts since login " & NewContacts & vbNewLine & vbNewLine & "More detailed Stats can be found on the main site"

    Case "USERSTAT"
        'User Stats
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            ReturnVar = 25
            SendMSG "Please enter the name of the person you wish to have stats on:"
        Else
            SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
        End If
        
    Case "ASITE"
        'Add Funny Site
        ReturnVar = 29
        SendMSG "Address of funny site you wish to add:"
        
    Case "VJOKE"
        'View Joke
        If UBound(Jokes()) = 0 Then
            SendMSG "There are no Jokes currently stored"
        Else
            ReturnVar = 30
            SendMSG "View Jokes:" & vbNewLine & "1 to " & UBound(Jokes())
        End If
    
    Case "AJOKE"
        'Add Joke
        ReturnVar = 31
        SendMSG "Please Enter your Joke, then enter your Punchline:"
    
    Case "AJOKE1"
        'Add Joke
        ReturnVar = 32
        SendMSG "Please Enter your Punchline:"
    
    Case "RJOKE"
        If UBound(Jokes()) = 0 Then
            SendMSG "There are no jokes currently stored"
        Else
            SendMSG RandomJoke
        End If
    Case "CONVOS"
        'Open Convos
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            SendMSG "Open Convos:" & frmMain.OpenConvos
        End If
    
    Case "RSS"
        'Paper Scissors Stone
        ReturnVar = 33
        SendMSG "Welcome to Paper Scissors Stone" & vbNewLine & "Choose either:" & vbNewLine & "1.Paper" & vbNewLine & "2.Scissors" & vbNewLine & "3.Stone"
    
    Case "AMSN"
        'Add MSN Nickname
        ReturnVar = 35
        SendMSG "Please Enter the MSN Nickname:"
    
    Case "VMSN"
        'View MSN Nickname
        If UBound(Nicks()) = 0 Then
            SendMSG "There are no MSN NickNames currently stored"
        Else
            ReturnVar = 36
            SendMSG "MSN NickNames:" & vbNewLine & "1 to " & UBound(Nicks())
        End If

    Case "RMSN"
        'Random MSN Nickname
        If UBound(Nicks()) = 0 Then
            SendMSG "There are no MSN NickNames currently stored"
        Else
            ReturnVar = 20
            SendMSG "Random MSN NickName:" & vbNewLine & RandomNick
        End If
        
    Case "HILOW"
        'Hi Low Game
        ReturnVar = 37
        HiLowNo = Int(Rnd * 100) + 1
        SendMSG "Welcome to Hi-Low, Guess the number between 1 and 100"
        
    Case "WIPEOUT"
        'Wipeout
        ReturnVar = 38
        SendMSG "Welcome to Wipeout" & vbNewLine & "Your aim is to remove the hashes by flipping squares, which inturn flip the surrounding squares"
        SendMSG "Square Values" & vbNewLine & " 1|2|3" & vbNewLine & " *-+-+-" & vbNewLine & " 4|5|6" & vbNewLine & " *-+-+-" & vbNewLine & " 7|8|9" & vbNewLine & "Type \exit to exit"
        For X = 1 To 3
            For Y = 1 To 3
                If Rnd < 0.5 Then
                    Wipeout(X, Y) = True
                End If
            Next
        Next
        WipeMoves = 0
        SendMSG ShowWipeout & vbNewLine & vbNewLine & "Choice:"
        
    Case "OTHELLO"
        'Othello
        ReturnVar = 39
        SendMSG "Welcome to Othello" & vbNewLine & "Your aim is to have the most 8's on screen at the end of the game" & vbNewLine & "Play by entering the relevent co-ordinates" & vbNewLine & "eg. 5,5"
        For X = 1 To 8
            For Y = 1 To 8
                OthelloGrid(X, Y) = 0
            Next
        Next
        OthelloGrid(4, 4) = 1   'Add a White
        OthelloGrid(5, 4) = 2   'Add a Black
        OthelloGrid(4, 5) = 2   'Add a Black
        OthelloGrid(5, 5) = 1   'Add a White
        
        SendMSG DrawOthello & vbNewLine & vbNewLine & "Choice:"
        
    Case "LEAVE"
        'Bots leaves the conversation
        SendMSG "Goodbye"
        Unload Me
        
    Case "ARCNEWS"
         'Archive News
        If UBound(OldNews()) = 0 Then
            SendMSG "There are no News currently stored"
        Else
            ReturnVar = 41
            SendMSG "Old News:" & vbNewLine & "1 to " & UBound(OldNews())
        End If
        
    Case "VAC"
        'View all IM Messiah Commands
        SendMSG "All IM Messiah Commands:" & vbNewLine & "\ADDMOD  - Add Mod" & vbNewLine & "\DELMOD - Delete Mod" & vbNewLine & "\CRASH - Crash User" & vbNewLine & "\FLOOD - Flood User" & vbNewLine & "\POKE - Poke User" & vbNewLine & "\DUSERS - Display Users" & vbNewLine & "\DADMINS - Display Admins" & vbNewLine & "\POPUP - Popup Message" & vbNewLine & "\CHGNAME - Change Name" & vbNewLine & "\SOCON - Number of open Convos" & vbNewLine & "\STCON - Total Number of Convos" & vbNewLine & "\NQUOTE1 - New Quote" & vbNewLine & "\VQUOTE - View Quotes" & vbNewLine & "\RQUOTE - Display random quotes"
        SendMSG "\ABOUT - About" & vbNewLine & "\BANUSER - Ban User" & vbNewLine & "\UNBANUSER - Unban user" & vbNewLine & "\BANLIST - Display BanList" & vbNewLine & "\ASUG - Add Suggestion" & vbNewLine & "\VSUG - View Suggestions" & vbNewLine & "\CSUG - Clear Suggestions" & vbNewLine & "\INVITE - Invite User" & vbNewLine & "\ANNOUNCE - Announce Message" & vbNewLine & "\GENSTAT - General Stats" & vbNewLine & "\USERSTAT - User Stats" & vbNewLine & "\ASITE - Add Funny Site" & vbNewLine & "\VJOKE - View Joke" & vbNewLine & "\AJOKE - Add Joke"
        SendMSG "\CONVOS - Open Convos" & vbNewLine & "\RSS - Paper Scissors Stone" & vbNewLine & "\AMSN - Add MSN Nickname" & vbNewLine & "\VMSN - View MSN Nickname" & vbNewLine & "\RMSN - Random MSN Nickname" & vbNewLine & "\HILOW - Hi Low Game" & vbNewLine & "\WIPEOUT - Wipeout" & vbNewLine & "\OTHELLO - Othello" & vbNewLine & "\LEAVE - Bots leaves the convo" & vbNewLine & "\ARCNEWS - Archive News" & vbNewLine & "\VAC - View all Commands" & vbNewLine & "\SOCKETS - View Sockets" & vbNewLine & "\CURSOC - Current Socket" & vbNewLine & "\BOBAI - Switches AI on or off"
        SendMSG "\USRCHT - Chat with other users" & vbNewLine & "\ATOPIC - Add topic to Forum" & vbNewLine & "\VTOPIC - View Topic in Forum" & vbNewLine & "\GOOGLE - Perform Google Search" & vbNewLine & "\KILLSCK - Kills a socket" & vbNewLine & "\FREESCK - No of Free socket" & vbNewLine & "\GETTIME - Shows local Time" & vbNewLine & "\STATUS - Change Bots Status" & vbNewLine & "\APPLY - Apply for Mod" & vbNewLine & "\UPDATES - Shows Recent Bot Updates" & vbNewLine & "\SHWNANN - Turn On,Off Annoucements" & vbNewLine & "\HANGMAN - Play Hangman" & vbNewLine & "\ABOUTBOB - Displays info about AI" & vbNewLine & "\ADDTOBOB - Add responses"
        SendMSG "\ADDBOT - Add User Bot" & vbNewLine & "\VIEWBOTS - View User Bots" & vbNewLine & "\WEBDEF - Web Definition for a word" & vbNewLine & "\WEBMEMO - Store Memo on Site" & vbNewLine & "\EMAILMEMO - Send Memo as email" & vbNewLine & "\STRMEMO - Store Memo on bot" & vbNewLine & "\VIEWMEMO - View Stored Memos" & vbNewLine & "\LANG - Change Bots Language" & vbNewLine & "\REVSTR - Reverse String" & vbNewLine & "\JBLSTR - Jumbles String" & vbNewLine & "\UCSTR - Turns to upper case" & vbNewLine & "\LCSTR - Turns to lower case" & vbNewLine & "\POLL - View Current Poll" & vbNewLine & "\VPOLL - Vote for a poll"
        
        SendMSG "Menu Commands:" & vbNewLine & "\MENU" & vbNewLine & "\ADMIN" & vbNewLine & "\USER" & vbNewLine & "\MESSIAH" & vbNewLine & "\STATS" & vbNewLine & "\QUOTES" & vbNewLine & "\OTHER" & vbNewLine & "\FUN" & vbNewLine & "\FSITES" & vbNewLine & "\GAMES" & vbNewLine & "\JOKES" & vbNewLine & "\BOT" & vbNewLine & "\NEWS" & vbNewLine & "\CHAT" & vbNewLine & "\FORUM" & vbNewLine & "\SUG" & vbNewLine & "\MSNNAME" & vbNewLine & "\USRBOT"
    
    Case "SOCKETS"
        'View Sockets
        
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            SendMSG "Sockets:" & frmMain.AllSocket
        Else
            SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
        End If
        
    Case "CURSOC"
        'View Current Socket
        
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            SendMSG "Current Socket:" & frmMain.CurSocket(buddy)
        Else
            SendMSG "You have to be an Mod or higher to have access to these commands"
        End If
        
    Case "USRCHT"
        'User Chat
        If MeWannaChat = False Then
            SendMSG "Chat is currently Disabled by Admins"
        Else
            SendMSG "Terms of using the Chat:" & vbNewLine & "1.No Swearing" & vbNewLine & "2.No Spamming or Advertising" & vbNewLine & "3.No Use of bots" & vbNewLine & vbNewLine & "Messiah Commands have been disabled in the chat to prevent flooding"
            
            If ChatOn = False Then
                SendMSG "No Chats Started, this will become the chat room"
                ImChat = True
                ChatOn = True
            Else
                SendMSG "You shall be connected to a chat shortly, please wait"
                frmMain.AddPersonToChat buddy
            End If
        End If
        
    Case "BOBAI"
        'Bob AI System
        AIon = Not AIon
        If AIon = True Then
            SendMSG "AI Bot is now on"
            SendMSG "AI Bot is still in beta stages and so may not fully work"
        Else
            SendMSG "AI Bot is now off"
        End If
        
    Case "ATOPIC"
        'Add Topic to Forum
        ReturnVar = 71
        SendMSG "New Topic:" & vbNewLine & "Please Enter the Name of the Topic"
        
    Case "ATOPIC1"
        'Add Topic to Forum
        ReturnVar = 72
        SendMSG "Enter the Post:"
    Case "VTOPIC"
        'View Topic in Forum
        If UBound(Forums()) > 0 Then
            ReturnVar = 70
            SendMSG "Forum Topics:" & ReturnTopics
            SendMSG "Your Selection:"
        Else
            SendMSG "No Topics currently stored"
        End If
        
    Case "GOOGLE"
        'Get a google Search
        SendMSG "Google Search:" & vbNewLine & "Search for:"
        ReturnVar = 45
        
        
    Case "KILLSCK"
        'Kill a socket (End Convo)
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            ReturnVar = 48
            SendMSG "Kill Socket:" & vbNewLine & "Please Enter the Index of the socket you wish to kill"
        Else
            SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
        End If
        
    Case "FREESCK"
        'Number of Free Sockets
        SendMSG "Number of Free Sockets:" & frmMain.sckCount
       
    Case "STATUS"
        'Change the bots status
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            ReturnVar = 49
            SendMSG "Bots Status:" & vbNewLine & "Please Enter the Index of the Status you would like"
        Else
            SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
        End If
    
    Case "APPLY"
        'Apply to be a Mod
        ReturnVar = 50
        SendMSG "Apply to be a Mod:" & vbNewLine & "Please enter why you believe you should be a Mod, a decision will be sent back to you as soon as possible"
    
    Case "UPDATES"
        'Show Bot Updates
        SendMSG BotUpdates
        
    Case "GETTIME"
        'Show Local Time
        SendMSG Time
    
    Case "SHWANN"
        'Show Announcements
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            HearAnnounce = Not HearAnnounce
            If HearAnnounce = True Then
                SendMSG "Announcements will now be shown"
            Else
                SendMSG "Announcements will now not be shown"
            End If
        Else
            SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
        End If
    Case "BOTLOG"
        'Bot Log
        SendMSG "Bot Log:" & vbNewLine & ViewLog
    Case "HANGMAN"
        'Hangman
        SendMSG "Welcome to Hangman" & vbNewLine & "Type \exit to end the game"
        If UBound(HangmanWords()) = 0 Then
            SendMSG "Sorry there is no words currently in the memory"
        Else
            ReturnVar = 51
            HWord = HangmanWords(Int(Rnd * UBound(HangmanWords())) + 1)
            ReDim Tried(0) As String
            SendMSG ShowWord & vbNewLine & vbNewLine & "Your selection:"
        End If
    Case "ADDTOBOB"
        'Add To Bob
        SendMSG "Add To AI:" & vbNewLine & "Please type the message you want Bob to reply to:"
        ReturnVar = 53
    Case "ADDTOBOB1"
        'Add To Bob
        SendMSG "Please type in the response:"
        ReturnVar = 54
    Case "BOBSTATS"
        'Stats about bob
        Dim BobsReply As Integer
        BobsReply = 0
        If UBound(AIResponses()) <> 0 Then
            For X = 1 To UBound(AIResponses())
                For Y = 1 To UBound(AIResponses(X).Answers())
                    BobsReply = BobsReply + 1
                Next
            Next
        End If
    
        SendMSG "Bobs Stats:" & vbNewLine & "Keywords:" & UBound(AIResponses()) & vbNewLine & "Responses:" & BobsReply
    Case "ABOUTBOB"
        'About Bob
        SendMSG "Bob AI System" & vbNewLine & "Created by Kevin Pfister" & vbNewLine & "Version 3.2"
    Case "ADDBOT"
        'Add a user Bot
        ReturnVar = 56
        SendMSG "Adding your MSN Bot:" & vbNewLine & "Please Enter its name:"
    Case "ADDBOT1"
        ReturnVar = 57
        SendMSG "Please Enter its email Address:"
    Case "ADDBOT2"
        ReturnVar = 58
        SendMSG "Please Enter a brief Description of it:"
    Case "VIEWBOTS"
        SendMSG "View User Bots:"
        If UBound(HangmanWords()) = 0 Then
            SendMSG "Sorry No User Bots have been added"
        Else
            ReturnVar = 59
            SendMSG "1 to " & UBound(BotName()) & vbNewLine & vbNewLine & "Your selection:"
        End If
    Case "WEBDEF"
        'Web Definition for a word
        SendMSG "Please Enter the word you would like the definition of:"
        SendMSG "Service is currently offline"
    Case "WEBMEMO"
        'Store Memo on site
        ReturnVar = 62
        SendMSG "Please Enter the Memo:"
    Case "EMAILMEMO"
        'Email Memo
        SendMSG "Please Enter the Memo:"
        SendMSG "Service is currently offline"
    Case "STRMEMO"
        'Store Memo
        ReturnVar = 63
        SendMSG "Please Enter the Memo:"
    Case "VIEWMEMO"
        'View a stored memo
        If Dir(App.Path & "\Data\Memo\" & buddy & ".txt") <> buddy & ".txt" Then
            SendMSG "Sorry, a Memo was not found under your email address", 4
        Else
            Open App.Path & "\Data\Memo\" & buddy & ".txt" For Input As #1
            XString = Input(LOF(1), 1)
            Close
            SendMSG "Message Stored:" & vbNewLine & XString, 8
        End If
    Case "DISPPIC"
        'frmMain.msgsend ("CHG " & frmMain.intTrailid & " HDN" & vbCrLf)
        Call CreateMSNObject
        frmMain.msgsend "CHG " & frmMain.intTrailid & " NLN 536870948 " & URLEncode(MSNObject) & vbCrLf
    Case "LANG"
        'Change Language
        ReturnVar = 61
        SendMSG "Languages:" & ShowLangs
    Case "REVSTR"
        'Reverse String
        ReturnVar = 66
        SendMSG "Please Enter the string you would like to reserve:"
    Case "JBLSTR"
        'Jumble String
        ReturnVar = 67
        SendMSG "Please Enter the string you would like to jumble:"
    Case "UCSTR"
        'Upper Case String
        ReturnVar = 68
        SendMSG "Please Enter the string you would like to convert to Upper Case:"
    Case "LCSTR"
        'Lower Case String
        ReturnVar = 69
        SendMSG "Please Enter the string you would like to convert to Lower Case:"
    
    Case "POLL"
        'View Poll
        If UBound(PollPosts()) = 0 Then
            SendMSG "No Poll is currently being run"
        Else
            SendMSG "Poll:" & vbNewLine & ReturnPoll & vbNewLine & vbNewLine & "To Vote Type \VPOLL"
        End If
    Case "USERNAME"
        'Alter your Username
        ReturnVar = 76
        SendMSG "Please Enter the name you would like to be prefixed as:"
    Case "MSGTOSCK"
        'Send Msg To Socket
        ReturnVar = 77
        SendMSG "Please Enter the Socket Index you wish to Send to:"
    Case "MSGTOSCK1"
        'Send Msg To Socket
        ReturnVar = 78
        SendMSG "Please Enter the Message you would like to send:"
    Case "8BALL"
        '8 Ball
        ReturnVar = 79
        SendMSG "What would you like the 8 Ball to predict?"
    Case "EXITCHAT"
        'Exit Chat
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            Call frmMain.ExitChats
            frmMain.TellMod "Chatroom Stopped By " & PrefGetName(buddy)
        End If
    Case "ASCII"
        'Ascii Art
        SendMSG "Please Configure your viewing screen just to see the following text"
        
        SendMSG "0000000000000000000000000000000000000000" & vbNewLine & "0000000000000000000000000000000000000000" & vbNewLine & "0000000000000000000000000000000000000000" & vbNewLine & "0000000000000000000000000000000000000000" & vbNewLine & "0000000000000000000000000000000000000000" & vbNewLine & "0000000000000000000000000000000000000000" & vbNewLine & "0000000000000000000000000000000000000000" & vbNewLine & "0000000000000000000000000000000000000000" & vbNewLine & "0000000000000000000000000000000000000000" & vbNewLine & "0000000000000000000000000000000000000000"
    Case "WARN"
        'Warn User
        If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
            ReturnVar = 80
            SendMSG "Warn User:" & vbNewLine & "Please Enter the Name of the user you wish to Warn:"
        Else
            SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
        End If
    Case "CONTACT"
        'Contact us with a message
        ReturnVar = 81
        SendMSG "Please Enter the message you would like to send use, A reply will be sent back as soon as possible"
    Case "UPDATEWEB"
        'Update Web Stats
        UpdateStats
    Case "TRAILERS"
        'Film Trailors
        If UBound(TrailerName()) = 0 Then
            SendMSG "Newest Trailers:" & vbNewLine & "There are no Trailers Stored"
        Else
            ReturnVar = 89
            SendMSG "Newest Trailers:" & vbNewLine & ModMedia.GetTrailors & vbNewLine & vbNewLine & "Please Enter your Selection:"
        End If
    Case "TVSHOWS"
    
    Case "DOWNLOADS"
        'Downloads
        
    Case "AREVIEW"
        'Add Film Review
        ReturnVar = 84
        SendMSG "New Film Review:" & vbNewLine & "Please Enter the Name of the Film"
        
    Case "AREVIEW1"
        'Add Review
        ReturnVar = 85
        SendMSG "Enter the Review:"
    Case "VREVIEW"
        'View Film Reviews
        If UBound(Reviews()) > 0 Then
            ReturnVar = 86
            SendMSG "Film Reviews:" & ReturnReviews
            SendMSG "Your Selection:"
        Else
            SendMSG "No Film Reviews currently stored"
        End If
    Case "USED"
        'people that have used Messiah
        If UBound(UsedMe()) = 0 Then
            SendMSG "Currently No One has used Messiah"
        Else
            Output = ""
            For Z = 1 To UBound(UsedMe())
                Output = Output & vbNewLine & PrefName(UsedMe(Z))
            Next
            SendMSG "Users:" & Output
        End If
    Case "AADMIN"
        'Add Admin
        If IsAdmin(buddy) = False Then
            SendMSG GetPhrase(UserLang, 2) 'Must be admin
        Else
            ReturnVar = 90
            SendMSG "Please Enter 'Super Admin' Password:"
        End If
    Case Else
        If Mid$(Command, 1, 5) = "VPOLL" Then
            If UBound(PollPosts()) = 0 Then
                SendMSG "No Poll is currently being run"
            Else
                ReturnVar = 75
                If Len(Command) <> 5 Then
                    ProcessReturn Trim$(Mid$(Command, 6))
                Else
                    SendMSG "Poll:" & vbNewLine & ReturnPoll & vbNewLine & "Please Enter the Index of the Poll Item you wish to vote for"
                End If
            End If
        End If
        If Mid$(Command, 1, 7) = "REVIEWS" Then
            'Forum
            ReturnVar = 83
            If Len(Command) <> 7 Then
                ProcessReturn Trim$(Mid$(Command, 8))
            Else
                SendMSG "(*)Reviews: \reviews <Num>" & vbNewLine & "1.Add Review" & vbNewLine & "2.View Reviews" & vbNewLine & "3.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 5) = "MEDIA" Then
            'Media Stuff
            ReturnVar = 82
            If Len(Command) <> 5 Then
                ProcessReturn Trim$(Mid$(Command, 6))
            Else
                SendMSG "(~)Media: \media <Num>" & vbNewLine & "1.Film Reviews" & vbNewLine & "2.Film Trailers" & vbNewLine & "3.TV Shows" & vbNewLine & "4.Downloads" & vbNewLine & "5.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 6) = "STRING" Then
            'String Commands
            ReturnVar = 64
            If Len(Command) <> 6 Then
                ProcessReturn Trim$(Mid$(Command, 7))
            Else
                SendMSG "(*)String: \string <Num>" & vbNewLine & "1.Reserve String" & vbNewLine & "2.Jumble Text" & vbNewLine & "3.Upper Case Text" & vbNewLine & "4.Lower Case Text" & vbNewLine & "5.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 4) = "PREF" Then
            'Preferences Commands
            ReturnVar = 65
            If Len(Command) <> 4 Then
                ProcessReturn Trim$(Mid$(Command, 5))
            Else
                SendMSG "(*)Preferences: \pref <Num>" & vbNewLine & "1.Language" & vbNewLine & "2.Username" & vbNewLine & "3.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 4) = "MENU" Then
            'MENU Command
            ReturnVar = 1
            If Len(Command) <> 4 Then
                ProcessReturn Trim$(Mid$(Command, 5))
            Else
                SendMSG GetPhrase(UserLang, 32) & vbNewLine & GetPhrase(UserLang, 33) & vbNewLine & GetPhrase(UserLang, 34) & vbNewLine & "3.Other Functions" & vbNewLine & "4.Statistics" & vbNewLine & "5.Chat" & vbNewLine & "6.Fun Stuff" & vbNewLine & "7.Internet" & vbNewLine & "8.Preferences" & vbNewLine & "9.View All Commands" & vbNewLine & "10.About" & vbNewLine & vbNewLine & "(co): http://dep.zion.me.uk/IM", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 4) = "CHAT" Then
            'CHAT Command
            ReturnVar = 42
            If Len(Command) <> 4 Then
                ProcessReturn Trim$(Mid$(Command, 5))
            Else
                SendMSG "(*)Chat: \chat <Num>" & vbNewLine & "1.Chat to other users" & vbNewLine & "2.Forums" & vbNewLine & "3.AI Bot (Bob)" & vbNewLine & "4.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 3) = "BOB" Then
            'BOB Command
            ReturnVar = 52
            If Len(Command) <> 3 Then
                ProcessReturn Trim$(Mid$(Command, 4))
            Else
                SendMSG "(*)Bob: \Bob <Num>" & vbNewLine & "1.Switch On/Off" & vbNewLine & "2.Add Responce" & vbNewLine & "3.Stats" & vbNewLine & "4.About" & vbNewLine & "5.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 5) = "ADMIN" Then
            'Admin Commands
            If IsAdmin(buddy) = False Then
                SendMSG GetPhrase(UserLang, 2) 'Must be admin
            Else
                ReturnVar = 2
                If Len(Command) <> 5 Then
                    ProcessReturn Trim$(Mid$(Command, 6))
                Else
                    SendMSG "(*)Admin Commands: \admin <Num>" & vbNewLine & "1.Add Mod" & vbNewLine & "2.Delete Mod" & vbNewLine & "3.Ban User" & vbNewLine & "4.UnBan User" & vbNewLine & "5.Ban List" & vbNewLine & "6.Messiah Functions" & vbNewLine & "7.Bot" & vbNewLine & "8.Convos" & vbNewLine & "9.Back", 8
                End If
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 5) = "FORUM" Then
            'Forum
            ReturnVar = 43
            If Len(Command) <> 5 Then
                ProcessReturn Trim$(Mid$(Command, 6))
            Else
                SendMSG "(*)Forum: \forum <Num>" & vbNewLine & "1.Add Topic" & vbNewLine & "2.View Topics" & vbNewLine & "3.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 3) = "MOD" Then
            'Mod Commands
            If IsUser(buddy) = True Or IsAdmin(buddy) = True Then
                ReturnVar = 3
                If Len(Command) <> 3 Then
                    ProcessReturn Trim$(Mid$(Command, 4))
                Else
                    SendMSG "(*)Mod Commands: \mod <Num>" & vbNewLine & "1.Poke User" & vbNewLine & "2.Invite User" & vbNewLine & "3.Announce" & vbNewLine & "4.User Stats" & vbNewLine & "5.Current Socket" & vbNewLine & "6.Show Announcements" & vbNewLine & "7.Warn User" & vbNewLine & "8.Back", 8
                End If
            Else
                SendMSG GetPhrase(UserLang, 1) 'Must be Mod or Higher
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 7) = "MESSIAH" Then
            'Messiah Functions
            If IsAdmin(buddy) = False Then
                SendMSG GetPhrase(UserLang, 2) 'Must be admin
            Else
                ReturnVar = 6
                If Len(Command) <> 7 Then
                    ProcessReturn Trim$(Mid$(Command, 8))
                Else
                    SendMSG "(6)Messiah: \messiah <Num>" & vbNewLine & "Warning: Overuse of these commands can result in ban or removal of your admin rights" & vbNewLine & "1.Crash User" & vbNewLine & "2.Flood User" & vbNewLine & "3.Poke User" & vbNewLine & "4.Back", 8
                End If
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 5) = "STATS" Then
            'Change Name
            ReturnVar = 12
            If Len(Command) <> 5 Then
                ProcessReturn Trim$(Mid$(Command, 6))
            Else
                SendMSG "(*)Statistics: \stats <Num>" & vbNewLine & "1.General Stats" & vbNewLine & "2.Open Converstations" & vbNewLine & "3.Total Convo's" & vbNewLine & "4.Total Convo's" & vbNewLine & "5.Admins" & vbNewLine & "6.Mods" & vbNewLine & "7.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 6) = "QUOTES" Then
            'MENU Command
            ReturnVar = 13
            If Len(Command) <> 6 Then
                ProcessReturn Trim$(Mid$(Command, 7))
            Else
                SendMSG "(*)Quotes: \quotes <Num>" & vbNewLine & "1.New Quote" & vbNewLine & "2.View Quotes" & vbNewLine & "3.Random Quote" & vbNewLine & "4.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 3) = "SUG" Then
            'Fun Stuff
            ReturnVar = 46
            If Len(Command) <> 3 Then
                ProcessReturn Trim$(Mid$(Command, 4))
            Else
                SendMSG "(*)Suggestions: \sug <Num>" & vbNewLine & "1.Add Suggestion" & vbNewLine & "2.View Suggestions" & vbNewLine & "3.Clear Suggestions" & vbNewLine & "4.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 7) = "MSNNAME" Then
            'msnname
            ReturnVar = 47
            If Len(Command) <> 7 Then
                ProcessReturn Trim$(Mid$(Command, 8))
            Else
                SendMSG "(*)MSN Nickames: \msnname <Num>" & vbNewLine & "1.Add MSN Name" & vbNewLine & "2.View MSN Names" & vbNewLine & "3.Random MSN Name" & vbNewLine & "4.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 5) = "OTHER" Then
            'Other Commands
            ReturnVar = 19
            If Len(Command) <> 5 Then
                ProcessReturn Trim$(Mid$(Command, 6))
            Else
                SendMSG "(E)Other: \other <Num>" & vbNewLine & "1.Suggestions" & vbNewLine & "2.MSN Nicknames" & vbNewLine & "3.Quotes" & vbNewLine & "4.Apply for Mod" & vbNewLine & "5.Get Local Time" & vbNewLine & "6.Other MSN Bots" & vbNewLine & "7.String Functions" & vbNewLine & "8.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 4) = "MEMO" Then
            'MEMO Command
            ReturnVar = 60
            If Len(Command) <> 4 Then
                ProcessReturn Trim$(Mid$(Command, 5))
            Else
                SendMSG "(*)Memo: \memo <Num>" & vbNewLine & "1.Web Memo(Stores on Site)" & vbNewLine & "2.Email Memo(Sends memo)" & vbNewLine & "3.Bot Memo(Saves to bot)" & vbNewLine & "4.View Memos" & vbNewLine & "5.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 6) = "USRBOT" Then
            'Other Users Bots
            ReturnVar = 55
            If Len(Command) <> 6 Then
                ProcessReturn Trim$(Mid$(Command, 7))
            Else
                SendMSG "(E)Users MSN Bots: \usrbot <Num>" & vbNewLine & "These are bots made by users, Add them to test them out" & vbNewLine & "1.Add Your MSN Bot" & vbNewLine & "2.View MSN Bots" & vbNewLine & "3.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 3) = "FUN" Then
            'Fun Stuff
            ReturnVar = 24
            If Len(Command) <> 3 Then
                ProcessReturn Trim$(Mid$(Command, 4))
            Else
                SendMSG "(*)Fun Stuff: \fun <Num>" & vbNewLine & "1.Fun Sites" & vbNewLine & "2.Games" & vbNewLine & "3.Jokes" & vbNewLine & "4.Poll" & vbNewLine & "5.Media" & vbNewLine & "6.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 6) = "FSITES" Then
            'Fun Sites
            ReturnVar = 26
            If Len(Command) <> 6 Then
                ProcessReturn Trim$(Mid$(Command, 7))
            Else
                SendMSG "(*)Fun Sites: \fsites <Num>" & vbNewLine & "1.List" & vbNewLine & "2.Add to List" & vbNewLine & "3.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 5) = "GAMES" Then
            'Games
            ReturnVar = 27
            If Len(Command) <> 5 Then
                ProcessReturn Trim$(Mid$(Command, 6))
            Else
                SendMSG "(*)Games: \games <Num>" & vbNewLine & "1.Paper Scissors Stone" & vbNewLine & "2.Hi - Low" & vbNewLine & "3.Wipeout" & vbNewLine & "4.Othello" & vbNewLine & "5.Hangman" & vbNewLine & "6.8 Ball" & vbNewLine & "7.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 5) = "JOKES" Then
            'Jokes
            ReturnVar = 28
            If Len(Command) <> 5 Then
                ProcessReturn Trim$(Mid$(Command, 6))
            Else
                SendMSG "(*)Jokes: \jokes <Num>" & vbNewLine & "1.Add Joke" & vbNewLine & "2.View Joke" & vbNewLine & "3.Random Joke" & vbNewLine & "4.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 3) = "BOT" Then
            'Bot
            If IsAdmin(buddy) = False Then
                SendMSG GetPhrase(UserLang, 2) 'Must be admin
            Else
                ReturnVar = 34
                If Len(Command) <> 3 Then
                    ProcessReturn Trim$(Mid$(Command, 4))
                Else
                    SendMSG "(*)Bot: \bot <Num>" & vbNewLine & "1.Popup" & vbNewLine & "2.Change Name" & vbNewLine & "3.Sockets in use" & vbNewLine & "4.Bot Log" & vbNewLine & "5.Kill Socket" & vbNewLine & "6.Free Sockets" & vbNewLine & "7.Status" & vbNewLine & "8.Message to Socket" & vbNewLine & "9.Back", 8
                End If
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 4) = "NEWS" Then
            'News
            ReturnVar = 40
            If Len(Command) <> 4 Then
                ProcessReturn Trim$(Mid$(Command, 5))
            Else
                SendMSG "(*)News: \news <Num>" & vbNewLine & "Uses BBC News Service" & vbNewLine & "1.Current News" & vbNewLine & "2.Archive" & vbNewLine & "3.Show Bot Updates" & vbNewLine & "4.Back", 8
            End If
            Exit Sub
        End If
        If Mid$(Command, 1, 3) = "INT" Then
            'Internet
            ReturnVar = 44
            If Len(Command) <> 3 Then
                ProcessReturn Trim$(Mid$(Command, 4))
            Else
                SendMSG "(*)Internet: \int <Num>" & vbNewLine & "1.News" & vbNewLine & "2.Search Google" & vbNewLine & "3.Word Definition" & vbNewLine & "4.Memo Service" & vbNewLine & "5.Back", 8
            End If
            Exit Sub
        End If
        
    End Select
    
End Sub

Sub ProcessReturn(Message)
    Dim Z As Integer
    Dim Message1 As String
    If ReturnVar = 0 Then
        If AIon = True Then
            If GetAIResponse(Message) <> "" Then
                SendMSG "(M)" & GetAIResponse(Message), 7
            End If
        End If
        Exit Sub
    End If
    
    Select Case ReturnVar
    
    Case 1
        'MENU Command
        ReturnVar = 0
        If Message = "1" Then
            'Admin Commands
            ProcessCommand "ADMIN"
        ElseIf Message = "2" Then
            'Mod Commands
            ProcessCommand "MOD"
        ElseIf Message = "3" Then
            'Other Commands
            ProcessCommand "OTHER"
        ElseIf Message = "4" Then
            'Statistics
            ProcessCommand "STATS"
        ElseIf Message = "5" Then
            'Chat
            ProcessCommand "CHAT"
        ElseIf Message = "6" Then
            'Fun Stuff
            ProcessCommand "FUN"
        ElseIf Message = "7" Then
            'Internet
            ProcessCommand "INT"
        ElseIf Message = "8" Then
            'Preferences
            ProcessCommand "PREF"
        ElseIf Message = "9" Then
            'VAC
            ProcessCommand "VAC"
        ElseIf Message = "10" Then
            'About
            ProcessCommand "ABOUT"
        End If

    Case 2
        'Admin Commands
        ReturnVar = 0
        If Message = "1" Then
            'Add User
            ProcessCommand "ADDMOD"
        ElseIf Message = "2" Then
            'Delete User
            ProcessCommand "DELMOD"
        ElseIf Message = "3" Then
            'Ban User
            ProcessCommand "BANUSER"
        ElseIf Message = "4" Then
            'UnBan User
            ProcessCommand "UNBANUSER"
        ElseIf Message = "5" Then
            'Ban List
            ProcessCommand "BANLIST"
        ElseIf Message = "6" Then
            'Messiah Functions
            ProcessCommand "MESSIAH"
        ElseIf Message = "7" Then
            'Bot
            ProcessCommand "BOT"
        ElseIf Message = "8" Then
            'Open Convos
            ProcessCommand "CONVOS"
        ElseIf Message = "9" Then
            'Back
            ProcessCommand "MENU"
        End If

    Case 3
        'Mod Commands
        ReturnVar = 0
        If Message = "1" Then
            'POKE User
            ProcessCommand "POKE"
        ElseIf Message = "2" Then
            'Invite User
            ProcessCommand "INVITE"
        ElseIf Message = "3" Then
            'Announce Message
            ProcessCommand "ANNOUNCE"
        ElseIf Message = "4" Then
            'User Stats
            ProcessCommand "USERSTAT"
        ElseIf Message = "5" Then
            'Current Socket
            ProcessCommand "CURSOC"
        ElseIf Message = "6" Then
            'Show Announcements
            ProcessCommand "SHWANN"
        ElseIf Message = "7" Then
            'Warn User
            ProcessCommand "WARN"
        ElseIf Message = "8" Then
            'Back
            ProcessCommand "MENU"
        End If

    Case 4
        'Add a Mod
        ReturnVar = 0
        If InStr(1, Message, "@") <> 0 Then
            AddUser Message
            SendMSG "Mod Added"
            frmMain.TellMod "Mod (" & PrefGetName(Message) & ") Added by " & PrefGetName(buddy)
        Else
            SendMSG "Not a valid email address", 4
        End If

    Case 5
        'Remove a Mod
        ReturnVar = 0
        If InStr(1, Message, "@") <> 0 Then
            DeleteUser Message
            SendMSG "Mod Removed"
            frmMain.TellMod "Mod (" & PrefGetName(Message) & ") Removed by " & PrefGetName(buddy)
        Else
            SendMSG "Not a valid email address", 4
        End If

    Case 6
        'Messiah Functions
        ReturnVar = 0
        If Message = "1" Then
            'Crash User
            ProcessCommand "CRASH"
        ElseIf Message = "2" Then
            'Flood User
            ProcessCommand "FLOOD"
        ElseIf Message = "3" Then
            'POKE User
            ProcessCommand "POKE"
        ElseIf Message = "4" Then
            'Back
            ProcessCommand "ADMIN"
        End If

    Case 7
        'Crash a user
        ReturnVar = 0
        If InStr(1, Message, "@") <> 0 Then
            DoEvents
            frmMain.Messenger.SendData "XFR " & frmMain.intTrailid & " SB" & vbCrLf
            frmMain.buddyconnect = Message
            Dim SM As Long
            SM = Timer
            Do
                DoEvents
            Loop Until Timer - SM > 1
            For Z = 1 To 100
                DoEvents
                Call frmMain.TellUser(Message, ":@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@")
            Next
            SendMSG "User has been Crashed"
            frmMain.TellMod PrefGetName(buddy) & " Is crashing " & PrefGetName(Message)
        Else
            SendMSG "Not a valid email address", 4
        End If
        
    Case 8
        'Flood
        ReturnVar = 0
        If InStr(1, Message, "@") <> 0 Then
            frmMain.Messenger.SendData "XFR " & frmMain.intTrailid & " SB" & vbCrLf
            frmMain.buddyconnect = Message
            For Z = 1 To 20
                DoEvents
                Call frmMain.TellUser(Message, "You are getting flooded by IM Messiah")
            Next
            SendMSG "User has been flood"
            frmMain.TellMod PrefGetName(buddy) & " Is flooding " & PrefGetName(Message)
       
        Else
            SendMSG "Not a valid email address", 4
        End If
        
    Case 9
        'Poke a user
        ReturnVar = 0
        If InStr(1, Message, "@") > 0 Then
            frmMain.Messenger.SendData "XFR " & frmMain.intTrailid & " SB" & vbCrLf
            frmMain.buddyconnect = Message
            Call frmMain.TellUser(Message, "Someone has poked you.... *poke*")
            SendMSG "User has been poked"
        Else
            SendMSG "Not a valid email address", 4
        End If
        
    Case 10
        'Popup Name
        ReturnVar = 0
        If Message <> "EXIT" Then
            Message = Replace(Message, " ", "%20")
            Message1 = Replace(frmContacts.lblName.Caption, " ", "%20")
            DoEvents
            frmMain.msgsend ("CHG " & frmMain.intTrailid & " HDN" & vbCrLf)
            Call frmMain.IncrementTrailID
            frmMain.msgsend ("REA " & frmMain.intTrailid & " " & frmMain.txtUsername.Text & " IM%20Messiah:%20" & Message & vbCrLf)
            Call frmMain.IncrementTrailID
            frmMain.msgsend ("CHG " & frmMain.intTrailid & " NLN" & vbCrLf)
            Call frmMain.IncrementTrailID
            frmMain.msgsend ("REA " & frmMain.intTrailid & " " & frmMain.txtUsername.Text & " " & Message1 & vbCrLf)
        End If
        
    Case 11
        'Change Name
        ReturnVar = 0
        If Message <> "EXIT" Then
            Message1 = Replace(Message, " ", "%20")
            frmMain.msgsend ("CHG " & frmMain.intTrailid & " HDN" & vbCrLf)
            Call frmMain.IncrementTrailID
            If UCase(Message) = "/NORM" Then
                frmMain.msgsend ("REA " & frmMain.intTrailid & " " & frmMain.txtUsername.Text & " IM%20Messiah:%20Type%20\menu%20to%20access%20features " & vbCrLf)
                frmContacts.lblName.Caption = "IM Messiah: Type \menu to access features"
            Else
                frmMain.msgsend ("REA " & frmMain.intTrailid & " " & frmMain.txtUsername.Text & " " & Message1 & vbCrLf)
                frmContacts.lblName.Caption = Message
            End If
            Call frmMain.IncrementTrailID
            frmMain.msgsend ("CHG " & frmMain.intTrailid & " NLN" & vbCrLf)
        End If
        
    Case 12
        'Statistics
        ReturnVar = 0
        If Message = "1" Then
            'General Stats
            ProcessCommand "GENSTAT"
        ElseIf Message = "2" Then
            'Open Converstations
            ProcessCommand "SOCON"
        ElseIf Message = "3" Then
            'Total Converstations
            ProcessCommand "STCON"
        ElseIf Message = "4" Then
            'Total Total Convos
        ElseIf Message = "5" Then
            'Display Admins
            ProcessCommand "DADMINS"
        ElseIf Message = "6" Then
            'Display Users
            ProcessCommand "DUSERS"
        ElseIf Message = "7" Then
            'Back
            ProcessCommand "MENU"
        End If
        
    Case 13
        'Quotes
        ReturnVar = 0
        If Message = "1" Then
            'New Quote
            ProcessCommand "NQUOTE1"
        ElseIf Message = "2" Then
            'View Quote
            ProcessCommand "VQUOTE"
        ElseIf Message = "3" Then
            'Random Quote
            ProcessCommand "RQUOTE"
        ElseIf Message = "4" Then
            'Back
            ProcessCommand "MENU"
        End If
        
    Case 14
        'Quote
        ReturnVar = 0
        QuoteSaid = Message
        ProcessCommand "NQUOTE2"
        
    Case 15
        'Person who said quote
        ReturnVar = 0
        If Message = "" Then
            Call AddQuote("Unknown", QuoteSaid)
        Else
            Call AddQuote(Message, QuoteSaid)
        End If
        SendMSG "Quote added"
        
    Case 16
        'Person who said quote
        ReturnVar = 0
        SendMSG ViewQuote(Val(Message))
        
    Case 17
        'Ban User
        ReturnVar = 0
        If InStr(1, Message, "@") <> 0 Then
            AddBan Message
            SendMSG "User Banned"
            frmMain.TellMod "User (" & PrefGetName(Message) & ") Banned by " & PrefGetName(buddy)
        Else
            SendMSG "Not a valid email address", 4
        End If
        
    Case 18
        'Unban User
        ReturnVar = 0
        If InStr(1, Message, "@") <> 0 Then
            DeleteBan Message
            SendMSG "User UnBanned"
            frmMain.TellMod "User (" & PrefGetName(Message) & ") UnBanned by " & PrefGetName(buddy)
        Else
            SendMSG "Not a valid email address", 4
        End If
        
    Case 19
        'Other Commands
        ReturnVar = 0
        If Message = "1" Then
            'Suggestions
            ProcessCommand "SUG"
        ElseIf Message = "2" Then
            'Add MSN Name
            ProcessCommand "MSNNAME"
        ElseIf Message = "3" Then
            'Quotes
            ProcessCommand "QUOTES"
        ElseIf Message = "4" Then
            'Apply for a mod
            ProcessCommand "APPLY"
        ElseIf Message = "5" Then
            'Local Time
            ProcessCommand "GETTIME"
        ElseIf Message = "6" Then
            'Other Users Bots
            ProcessCommand "USRBOT"
        ElseIf Message = "7" Then
            'String Functions
            ProcessCommand "STRING"
        ElseIf Message = "8" Then
            'Back
            ProcessCommand "MENU"
        End If
        
    Case 20
        'View Suggestion
        ReturnVar = 0
        SendMSG Suggestions(Val(Message))
        
    Case 21
        'Add Suggestion
        ReturnVar = 0
        AddSuggestion Message
        SendMSG "Suggestion Added"
        
    Case 22
        'Invite user
        ReturnVar = 0
        If InStr(1, Message, "@") <> 0 Then
            sndstr "CAL " & myTrID1 & " " & Message & vbCrLf
            frmMain.IncrementTrailID
            DoEvents
            SendMSG "User Added"
        Else
            SendMSG "Not a valid email address", 4
        End If
        
    Case 23
        'Announce
        ReturnVar = 0
        SendMSG "Message Announced"
        DoEvents
        SayToAll Message
    
    Case 24
        'Fun Stuff
        ReturnVar = 0
        If Message = "1" Then
            'Fun Sites
            ProcessCommand "FSITES"
        ElseIf Message = "2" Then
            'Games
            ProcessCommand "GAMES"
        ElseIf Message = "3" Then
            'Jokes
            ProcessCommand "JOKES"
        ElseIf Message = "4" Then
            'Poll
            ProcessCommand "POLL"
        ElseIf Message = "5" Then
            'Media
            ProcessCommand "MEDIA"
        ElseIf Message = "6" Then
            'Back
            ProcessCommand "MENU"
        End If

    Case 25
        'User Stats
        ReturnVar = 0
        If InStr(1, Message, "@") <> 0 Then
            SendMSG "User Stats:" & vbNewLine & "Name:" & PrefGetName(Message) & "Admin: " & IsAdmin(Message) & vbNewLine & "Mod: " & IsUser(Message) & vbNewLine & "Banned: " & IsBanned(Name) & vbNewLine & "Warnings:" & GetWarnings(Message)
        Else
            SendMSG "Not a valid email address", 4
        End If
    
    Case 26
        'Fun Sites
        ReturnVar = 0
        If Message = "1" Then
            'List Sites
            Dim FSites As String
            If UBound(FunnySites()) = 0 Then
                SendMSG "There are no funny sites currently stored"
            Else
                For Z = 1 To UBound(FunnySites())
                    FSites = FSites & vbNewLine & FunnySites(Z)
                Next
                SendMSG "Funny Sites:" & FSites
            End If
        ElseIf Message = "2" Then
            'Add Sites
            ProcessCommand "ASITE"
        ElseIf Message = "3" Then
            'Back
            ProcessCommand "FUN"
        End If

    Case 27
        'Games
        ReturnVar = 0
        If Message = "1" Then
            'Rock Scissors Stone
            ProcessCommand "RSS"
        ElseIf Message = "2" Then
            'HiLow
            ProcessCommand "HILOW"
        ElseIf Message = "3" Then
            'Wipeout
            ProcessCommand "WIPEOUT"
        ElseIf Message = "4" Then
            'Othello
            ProcessCommand "OTHELLO"
        ElseIf Message = "5" Then
            'Hangman
            ProcessCommand "HANGMAN"
        ElseIf Message = "6" Then
            '8Ball
            ProcessCommand "8BALL"
        ElseIf Message = "7" Then
            'Back
            ProcessCommand "FUN"
        End If
    Case 28
        'Jokes
        ReturnVar = 0
        If Message = "1" Then
            'Add Joke
            ProcessCommand "AJOKE"
        ElseIf Message = "2" Then
            'View Joke
            ProcessCommand "VJOKE"
        ElseIf Message = "3" Then
            'Random Joke
            ProcessCommand "RJOKE"
        ElseIf Message = "4" Then
            'Back
            ProcessCommand "FUN"
        End If
    
    Case 29
        'Add Funny Site
        ReturnVar = 0
        AddSite Message
        SendMSG "Funny Site added"
    Case 30
        'View Joke
        ReturnVar = 0
        SendMSG ViewJoke(Val(Message))

    Case 31
        'Add Joke
        ReturnVar = 0
        JokeLine = Message
        ProcessCommand "AJOKE1"
        
    Case 32
        'Add Joke
        ReturnVar = 0
        AddJoke JokeLine, Message
        SendMSG "Joke Added"
    Case 33
        'Paper Scissors Stone
        ReturnVar = 0
        Dim ComChoice As Integer
        ComChoice = Int(Rnd * 3) + 1
        Dim ComsRSS As String
        If ComChoice = 1 Then
            ComsRSS = "Paper"
        ElseIf ComChoice = 2 Then
            ComsRSS = "Scissors"
        ElseIf ComChoice = 3 Then
            ComsRSS = "Stone"
        End If
        If Message = "1" Then
            'Paper
            If ComChoice = 1 Then
                SendMSG "Players Choice: Paper" & vbNewLine & "Bots Choice: Paper" & vbNewLine & "Player Draws"
            ElseIf ComChoice = 2 Then
                SendMSG "Players Choice: Paper" & vbNewLine & "Bots Choice: Scissors" & vbNewLine & "Player Loses"
            ElseIf ComChoice = 3 Then
                SendMSG "Players Choice: Paper" & vbNewLine & "Bots Choice: Stone" & vbNewLine & "Player Wins"
            End If
        ElseIf Message = "2" Then
            'Scissors
            If ComChoice = 1 Then
                SendMSG "Players Choice: Scissors" & vbNewLine & "Bots Choice: Paper" & vbNewLine & "Player Wins"
            ElseIf ComChoice = 2 Then
                SendMSG "Players Choice: Scissors" & vbNewLine & "Bots Choice: Scissors" & vbNewLine & "Player Draws"
            ElseIf ComChoice = 3 Then
                SendMSG "Players Choice: Scissors" & vbNewLine & "Bots Choice: Stone" & vbNewLine & "Player Loses"
            End If
        ElseIf Message = "3" Then
            'Stone
            If ComChoice = 1 Then
                SendMSG "Players Choice: Stone" & vbNewLine & "Bots Choice: Paper" & vbNewLine & "Player Loses"
            ElseIf ComChoice = 2 Then
                SendMSG "Players Choice: Stone" & vbNewLine & "Bots Choice: Scissors" & vbNewLine & "Player Wins"
            ElseIf ComChoice = 3 Then
                SendMSG "Players Choice: Stone" & vbNewLine & "Bots Choice: Stone" & vbNewLine & "Player Draws"
            End If
        End If
        
    Case 34
        'Bot
        ReturnVar = 0
        If Message = "1" Then
            'Popup
            ProcessCommand "POPUP"
        ElseIf Message = "2" Then
            'Messiah Functions
            ProcessCommand "CHGNAME"
        ElseIf Message = "3" Then
            'Sockets
            ProcessCommand "SOCKETS"
        ElseIf Message = "4" Then
            'Bot Log
            ProcessCommand "BOTLOG"
        ElseIf Message = "5" Then
            'Kill
            ProcessCommand "KILLSCK"
        ElseIf Message = "6" Then
            'Free
            ProcessCommand "FREESCK"
        ElseIf Message = "7" Then
            'Status
            ProcessCommand "STATUS"
        ElseIf Message = "8" Then
            'Message To Socket
            ProcessCommand "MSGTOSCK"
        ElseIf Message = "9" Then
            'Back
            ProcessCommand "ADMIN"
        End If
        
    Case 35
        'Add MSN Nick
        ReturnVar = 0
        AddNick Message
        
    Case 36
        'View MSN Nick
        ReturnVar = 0
        SendMSG "MSN Nickname:" & vbNewLine & Nicks(Val(Message))
        
    Case 37
        'HiLow
        If Val(Message) = HiLowNo Then
            ReturnVar = 0
            SendMSG "You have guessed the correct number"
        ElseIf Val(Message) > HiLowNo Then
            SendMSG "You have guessed too High"
        ElseIf Val(Message) < HiLowNo Then
            SendMSG "You have guessed too Low"
        End If
        
    Case 38
        If Val(Message) > 0 And Val(Message) < 10 Then
            If Message = "1" Then
                Wipeout(1, 1) = Not Wipeout(1, 1)
                Wipeout(2, 1) = Not Wipeout(2, 1)
                Wipeout(1, 2) = Not Wipeout(1, 2)
            ElseIf Message = "2" Then
                Wipeout(1, 1) = Not Wipeout(1, 1)
                Wipeout(2, 1) = Not Wipeout(2, 1)
                Wipeout(3, 1) = Not Wipeout(3, 1)
                Wipeout(2, 2) = Not Wipeout(2, 2)
            ElseIf Message = "3" Then
                Wipeout(3, 1) = Not Wipeout(3, 1)
                Wipeout(2, 1) = Not Wipeout(2, 1)
                Wipeout(3, 2) = Not Wipeout(3, 2)
            ElseIf Message = "4" Then
                Wipeout(1, 1) = Not Wipeout(1, 1)
                Wipeout(1, 2) = Not Wipeout(1, 2)
                Wipeout(2, 2) = Not Wipeout(2, 2)
                Wipeout(1, 3) = Not Wipeout(1, 3)
            ElseIf Message = "5" Then
                Wipeout(2, 1) = Not Wipeout(2, 1)
                Wipeout(1, 2) = Not Wipeout(1, 2)
                Wipeout(2, 2) = Not Wipeout(2, 2)
                Wipeout(3, 2) = Not Wipeout(3, 2)
                Wipeout(2, 3) = Not Wipeout(2, 3)
            ElseIf Message = "6" Then
                Wipeout(3, 1) = Not Wipeout(3, 1)
                Wipeout(3, 2) = Not Wipeout(3, 2)
                Wipeout(3, 3) = Not Wipeout(3, 3)
                Wipeout(2, 2) = Not Wipeout(2, 2)
            ElseIf Message = "7" Then
                Wipeout(1, 3) = Not Wipeout(1, 3)
                Wipeout(2, 3) = Not Wipeout(2, 3)
                Wipeout(1, 2) = Not Wipeout(1, 2)
            ElseIf Message = "8" Then
                Wipeout(1, 3) = Not Wipeout(1, 3)
                Wipeout(2, 3) = Not Wipeout(2, 3)
                Wipeout(3, 3) = Not Wipeout(3, 3)
                Wipeout(2, 2) = Not Wipeout(2, 2)
            ElseIf Message = "9" Then
                Wipeout(3, 3) = Not Wipeout(3, 3)
                Wipeout(2, 3) = Not Wipeout(2, 3)
                Wipeout(3, 2) = Not Wipeout(3, 2)
            End If
            Dim BlackS As Boolean
            Dim X As Integer
            Dim Y As Integer
            BlackS = False
            For X = 1 To 3
                For Y = 1 To 3
                    If Wipeout(X, Y) = True Then
                        BlackS = True
                    End If
                Next
            Next
            WipeMoves = WipeMoves + 1
            If BlackS = True Then
                SendMSG ShowWipeout & vbNewLine & vbNewLine & "Choice:"
            Else
                SendMSG ShowWipeout & vbNewLine & vbNewLine & "You have won in " & WipeMoves & " moves"
            End If
        Else
            ReturnVar = 0
        End If
    
    Case 39
        'Othello
        X = Val(Mid(Message, 1, InStr(1, Message, ",") - 1))
        Y = Val(Mid(Message, InStr(1, Message, ",") + 1))
        If OthelloGrid(X, Y) = 0 Then
            Call GameEngine(X, Y, 1, 2)   'Call placing sub
            If GoodShot = 1 Then
                If CheckWin(2) = False Then
                    SendMSG "After Users go..." & vbNewLine & DrawOthello & vbNewLine & vbNewLine & "Computer Thinking...", 6
                    Call Comp   'Compute computers go
                Else
                    ReturnVar = 0
                End If
            Else
                SendMSG "Invalid Move, please try again or type \exit to leave", 4
            End If
        Else
            SendMSG "Invalid Move, please try again or type \exit to leave", 4
        End If
    
    Case 40
        'News
        ReturnVar = 0
        If Message = "1" Then
            'Current News
            SendMSG "Current News:" & vbNewLine & CurNews
        ElseIf Message = "2" Then
            'Archive News
            ProcessCommand "ARCNEWS"
        ElseIf Message = "3" Then
            'Bot Updates
            ProcessCommand "UPDATES"
        ElseIf Message = "4" Then
            'Back
            ProcessCommand "INT"
        End If

    Case 41
        'View Old News
        ReturnVar = 0
        SendMSG "Old News:" & vbNewLine & OldNews(Val(Message))
    
    Case 42
        'Chat
        ReturnVar = 0
        If Message = "1" Then
            'Chat to other Users
            ProcessCommand "USRCHT"
        ElseIf Message = "2" Then
            'Forum
            ProcessCommand "FORUM"
        ElseIf Message = "3" Then
            'AI Bot
            ProcessCommand "BOB"
        ElseIf Message = "4" Then
            'Back
            ProcessCommand "MENU"
        End If

    Case 43
        'Forum
        ReturnVar = 0
        If Message = "1" Then
            'Add Topic
            ProcessCommand "ATOPIC"
        ElseIf Message = "2" Then
            'View Topics
            ProcessCommand "VTOPIC"
        ElseIf Message = "3" Then
            'Back
            ProcessCommand "CHAT"
        End If
    Case 44
        'Internet
        ReturnVar = 0
        If Message = "1" Then
            'News
            ProcessCommand "NEWS"
        ElseIf Message = "2" Then
            'Search Google
            ProcessCommand "GOOGLE"
        ElseIf Message = "3" Then
            'Web Defintion
            ProcessCommand "WEBDEF"
        ElseIf Message = "4" Then
            'Memo
            ProcessCommand "MEMO"
        ElseIf Message = "5" Then
            'Back
            ProcessCommand "MENU"
        End If
    Case 45
        'Google Search
        ReturnVar = 0
        SendMSG "Google Results:" & frmMain.GoogleSearch(Message)
    Case 46
        'Suggestions
        ReturnVar = 0
        If Message = "1" Then
            'Add Suggestion
            ProcessCommand "ASUG"
        ElseIf Message = "2" Then
            'View Suggestions
            ProcessCommand "VSUG"
        ElseIf Message = "3" Then
            'Clear
            ProcessCommand "CSUG"
        ElseIf Message = "4" Then
            'Back
            ProcessCommand "OTHER"
        End If
    Case 47
        'MSN Names
        ReturnVar = 0
        If Message = "1" Then
            'Add MSN Name
            ProcessCommand "AMSN"
        ElseIf Message = "2" Then
            'View MSN Names
            ProcessCommand "VMSN"
        ElseIf Message = "3" Then
            'Random MSN Name
            ProcessCommand "RMSN"
        ElseIf Message = "4" Then
            'Back
            ProcessCommand "OTHER"
        End If
    Case 48
        'Kill Socket
        ReturnVar = 0
        If Val(Message) <= 0 Or Val(Message) > Capacity Then
            SendMSG "Value out of bounds"
        Else
            frmMain.KillSck Val(Message)
            SendMSG "Socket Killed"
            frmMain.TellMod PrefGetName(buddy) & " killed Socket" & Message
        End If
    Case 49
        ReturnVar = 0
        Select Case Val(Message)
            Case 1
                frmMain.msgsend ("CHG " & frmMain.intTrailid & " NLN" & vbCrLf)
            Case 2
                frmMain.msgsend "CHG " & frmMain.intTrailid & " BSY" & vbCrLf
            Case 3
                frmMain.msgsend "CHG " & frmMain.intTrailid & " BRB" & vbCrLf
            Case 4
                frmMain.msgsend "CHG " & frmMain.intTrailid & " AWY" & vbCrLf
            Case 5
                frmMain.msgsend "CHG " & frmMain.intTrailid & " PHN" & vbCrLf
            Case 6
                frmMain.msgsend "CHG " & frmMain.intTrailid & " LUN" & vbCrLf
            Case 7
                frmMain.msgsend "CHG " & frmMain.intTrailid & " HDN" & vbCrLf
        End Select
        
        For X = 1 To 7
            FrmMessiah.mnuStatus(X).Checked = False
        Next
        FrmMessiah.mnuStatus(Val(Message)).Checked = True
        SendMSG "Status Changed"
        frmMain.TellMod PrefGetName(buddy) & " has changed IM Messiah's Status"
    Case 50
        'Apply for Mod
        ReturnVar = 0
        Call ApplyMod(Message, buddy)
        SendMSG "Request as been sent"
        frmMain.TellMod PrefGetName(buddy) & " has applied to be a mod"
    Case 51
        'Hangman
        Dim Found As Boolean
        If UBound(Tried()) = 0 Then
            ReDim Preserve Tried(UBound(Tried()) + 1) As String
            Tried(UBound(Tried())) = UCase(Message)
        Else
            Found = False
            For Y = 1 To UBound(Tried())
                If UCase(Tried(Y)) = UCase(Message) Then
                    Found = True
                End If
            Next
            If Found = True Then
                SendMSG "You have already used that Letter"
            Else
                ReDim Preserve Tried(UBound(Tried()) + 1) As String
                Tried(UBound(Tried())) = UCase(Message)
            End If
        End If
        If ShowWord = HWord Then
            SendMSG "You Win" & vbNewLine & vbNewLine & ShowWord
            ReturnVar = 0
        Else
            SendMSG ShowWord & vbNewLine & vbNewLine & "Your selection:"
        End If
    Case 52
        'Bob AI
        ReturnVar = 0
        If Message = "1" Then
            'Turn On/off
            AIon = Not AIon
            If AIon = True Then
                SendMSG "AI Bot is now on"
                SendMSG "AI Bot is still in beta stages and so may not fully work"
            Else
                SendMSG "AI Bot is now off"
            End If
        ElseIf Message = "2" Then
            'Add Response
            ProcessCommand "ADDTOBOB"
        ElseIf Message = "3" Then
            'Stats
            ProcessCommand "BOBSTATS"
        ElseIf Message = "4" Then
            'About
            ProcessCommand "ABOUTBOB"
        ElseIf Message = "5" Then
            'Back
            ProcessCommand "CHAT"
        End If
    Case 53
        ReturnVar = 0
        AIAdding = Message
        ProcessCommand "ADDTOBOB1"
    Case 54
        ReturnVar = 0
        Call AddToAI(AIAdding, Message)
        SendMSG "Response Added"
    Case 55
        'Other Bots
        ReturnVar = 0
        If Message = "1" Then
            'Add User Bot
            ProcessCommand "ADDBOT"
        ElseIf Message = "2" Then
            'View User Bots
            ProcessCommand "VIEWBOTS"
        ElseIf Message = "3" Then
            'Back
            ProcessCommand "BOBSTATS"
        End If
    Case 56
        AddBot1 = Message
        ProcessCommand "ADDBOT1"
    Case 57
        AddBot2 = Message
        ProcessCommand "ADDBOT2"
    Case 58
        ReturnVar = 0
        AddUserBots AddBot1, AddBot2, Message
        SendMSG "User Bot has been added", 4
    Case 59
        'View Bot
        ReturnVar = 0
        If Val(Message) <= 0 Or Val(Message) > UBound(BotName()) Then
            SendMSG "Invalid Choice", 4
        Else
            SendMSG "Name: " & BotName(Val(Message)) & vbNewLine & "Address: " & BotAddr(Val(Message)) & vbNewLine & "Description:" & vbNewLine & BotDesc(Val(Message))
        End If
    Case 60
        'Memo
        ReturnVar = 0
        If Message = "1" Then
            'Web Memo
            ProcessCommand "WEBMEMO"
        ElseIf Message = "2" Then
            'Email Memo
            ProcessCommand "EMAILMEMO"
        ElseIf Message = "3" Then
            'Store Memo
            ProcessCommand "STRMEMO"
        ElseIf Message = "4" Then
            'View Memo
            ProcessCommand "VIEWMEMO"
        ElseIf Message = "5" Then
            'Back
            ProcessCommand "OTHER"
        End If
    Case 61
        'Select Language
        If Val(Message) <= 0 Or Val(Message) > UBound(OtherLang()) Then
            SendMSG "Invalid Selection"
        Else
            UserLang = Val(Message)
            Call SaveLang(buddy, UserLang)
            SendMSG OtherLang(UserLang).Name & " selected"
        End If
    Case 62
        'Web Memo
        SendMSG "This option does not work in the opensource Code"
        'Please replace the code with the code concerning your FTP
        
        ReturnVar = 0
        'Open App.Path & "\Data\WebMemo.html" For Output As #1
        'Print #1, "<HTML><HEAD><Title>IM Messiah Memo Service</title><FONT face=" & """" & "Times New Roman" & """" & " size=" & "12" & "><Body bgcolor=" & """" & "FFFFFF" & """" & ">"
        'Print #1, "<font color=" & """" & "3366CC" & """" & "font STYLE=" & """font-size: " & "12" & "px" & """" & ">IM Messiah Web Based Memo<br>"
        'Print #1, "<font color=" & """" & "3366CC" & """" & "font STYLE=" & """font-size: " & "12" & "px" & """" & ">-------------------------<br>"
        'Print #1, "<font color=" & """" & "&H00000000&" & """" & "font STYLE=" & """font-size: " & "12" & "px" & """" & ">" & Message & "<br>"
        'Print #1, "<br></FONT></BODY></HTML>"
        'Close
        
        'Message1 = Timer & Int(Rnd * 1000) & ".html"
        
        'FrmMessiah.Inet.AccessType = icDirect
        'FrmMessiah.Inet.Protocol = icFTP
        'FrmMessiah.Inet.RemoteHost = ""
        'FrmMessiah.Inet.UserName = ""
        'FrmMessiah.Inet.Password = ""
        'FrmMessiah.Inet.Execute , "put " & Chr(34) & App.Path & "\Data\WebMemo.html" & Chr(34) & " " & "www/Memos/" & Message1
        
        'SendMSG "Memo Stored at:" & vbNewLine & "www.dep.zion.me.uk/Memos/" & Message1
    Case 63
        'Bot Memo
        ReturnVar = 0
        Open App.Path & "\Data\Memo\" & buddy & ".txt" For Output As #1
        Print #1, Message
        Close
        SendMSG "Memo has been Saved under this Email address"
    Case 64
        'String
        ReturnVar = 0
        If Message = "1" Then
            'Reserve
            ProcessCommand "REVSTR"
        ElseIf Message = "2" Then
            'Jumble
            ProcessCommand "JBLSTR"
        ElseIf Message = "3" Then
            'Upper Case
            ProcessCommand "UCSTR"
        ElseIf Message = "4" Then
            'Lower Case
            ProcessCommand "LCSTR"
        ElseIf Message = "5" Then
            'Back
            ProcessCommand "OTHER"
        End If
    Case 65
        'Pref
        ReturnVar = 0
        If Message = "1" Then
            'Languages
            ProcessCommand "LANG"
        ElseIf Message = "2" Then
            'Change UserName
            ProcessCommand "USERNAME"
        ElseIf Message = "3" Then
            'Back
            ProcessCommand "MENU"
        End If
    Case 66
        'Reserve String
        ReturnVar = 0
        SendMSG Reserve(Message)
    Case 67
        'Jumble String
        ReturnVar = 0
        SendMSG "Currently under development"
    Case 68
        'Upper Case
        ReturnVar = 0
        SendMSG UCase$(Message)
    Case 69
        'Lower Case
        ReturnVar = 0
        SendMSG LCase$(Message)
    Case 70
        'View Topic
        ReturnVar = 0
        If Val(Message) <= 0 Or Val(Message) > UBound(Forums()) Then
            SendMSG "Invalid Selection"
        Else
            CurTopic = Val(Message)
            SendMSG "Forum Post:" & vbNewLine & Forums(CurTopic).Title
            For X = 1 To UBound(Forums(CurTopic).Posts())
                SendMSG Forums(CurTopic).Posts(X)
            Next
            ReturnVar = 73
            SendMSG "Your Selection:" & vbNewLine & "1.Add Post" & vbNewLine & "2.Back"
        End If
    Case 71
        'Add Topic
        ReturnVar = 0
        TopName = Message
        ProcessCommand "ATOPIC1"
    Case 72
        'Add Topic
        ReturnVar = 0
        Call AddFTopic(TopName, PrefGetName(buddy) & ": " & Message)
        SendMSG "Topic Added to Forum"
    Case 73
        'In Topic
        ReturnVar = 0
        If Message = "1" Then
            'Add Post
            ReturnVar = 74
            SendMSG "Please Enter your Post:"
        ElseIf Message = "2" Then
            'Back
            ProcessCommand "FORUM"
        End If
    Case 74
        'Add Post
        ReturnVar = 0
        Call AddFPost(CurTopic, PrefGetName(buddy) & ": " & Message)
        SendMSG "Post Added to Topic:" & vbNewLine & Forums(CurTopic).Title
    Case 75
        'Vote in a poll
        ReturnVar = 0
        If Val(Message) = 0 Or Val(Message) > UBound(PollPosts()) Then
            SendMSG "Invalid Selection"
        Else
            If HasVoted(buddy) = True Then
                SendMSG "You have already Voted"
            Else
                PollVotes(Val(Message)) = PollVotes(Val(Message)) + 1
                ReDim Preserve PollVoted(UBound(PollVoted()) + 1) As String
                PollVoted(UBound(PollVoted())) = UCase(buddy)
                SavePoll
                SendMSG "Your Vote has been Tallied"
            End If
        End If
    Case 76
        'Change UserName
        ReturnVar = 0
        If NewPrefName(buddy, Message) = True Then
            SendMSG "Name Preference Saved"
        Else
            SendMSG "Name has already been selected"
        End If
    Case 77
        'Send Msg to Socket
        ReturnVar = 0
        If Val(Message) <= 0 Or Val(Message) > Capacity Then
            SendMSG "Invalid Selection"
        Else
            SckSend = Val(Message)
            ProcessCommand "MSGTOSCK1"
        End If
    Case 78
        'Send Msg to Socket
        ReturnVar = 0
        Call frmMain.TellSck(SckSend, Message)
    Case 79
        '8Ball
        ReturnVar = 0
        Z = Int(Rnd * 5 + 1)
        If Z = 1 Then
            SendMSG "Don't Count on it"
        ElseIf Z = 2 Then
            SendMSG "Not very likely"
        ElseIf Z = 3 Then
            SendMSG "Might be true"
        ElseIf Z = 4 Then
            SendMSG "Yeah!"
        ElseIf Z = 5 Then
            SendMSG "I think it is true"
        End If
    Case 80
        ReturnVar = 0
        If InStr(1, Message, "@") > 0 Then
            WarnPerson Message
            SendMSG "User has been warned"
            frmMain.TellMod "User (" & PrefGetName(Message) & ") Warned by " & PrefGetName(buddy)
        Else
            SendMSG "Not a valid email address", 4
        End If
    Case 81
        'Contact us
        ReturnVar = 0
        Call ContactUs(Message, PrefGetName(buddy))
        SendMSG "Message Sent"
        frmMain.TellMod PrefGetName(buddy) & " has contacted Staff"
    Case 82
        'Media Menu
        ReturnVar = 0
        If Message = "1" Then
            'Film Reviews
            ProcessCommand "REVIEWS"
        ElseIf Message = "2" Then
            'Film Trailers
            ProcessCommand "TRAILERS"
        ElseIf Message = "3" Then
            'TV Shows
            ProcessCommand "TVSHOWS"
        ElseIf Message = "4" Then
            'Downloads
            ProcessCommand "DOWNLOADS"
        ElseIf Message = "5" Then
            'Back
            ProcessCommand "FUN"
        End If
    Case 83
        'Film Review
        ReturnVar = 0
        If Message = "1" Then
            'Add Topic
            ProcessCommand "AREVIEW"
        ElseIf Message = "2" Then
            'View Topics
            ProcessCommand "VREVIEW"
        ElseIf Message = "3" Then
            'Back
            ProcessCommand "MEDIA"
        End If
    Case 84
        'Add Film Review
        ReturnVar = 0
        FilmName = Message
        ProcessCommand "AREVIEW1"
    Case 85
        'Add Film Review
        ReturnVar = 0
        Call AddFilmT(FilmName, PrefGetName(buddy) & ": " & Message)
        SendMSG "Film Review Added"
    Case 86
        'View Review
        ReturnVar = 0
        If Val(Message) <= 0 Or Val(Message) > UBound(Reviews()) Then
            SendMSG "Invalid Selection"
        Else
            CurFilm = Val(Message)
            SendMSG "Review:" & vbNewLine & Reviews(CurFilm).Title
            For X = 1 To UBound(Reviews(CurFilm).Posts())
                SendMSG Reviews(CurFilm).Posts(X)
            Next
            ReturnVar = 87
            SendMSG "Your Selection:" & vbNewLine & "1.Add Review" & vbNewLine & "2.Back"
        End If
    Case 87
        'In Reviews
        ReturnVar = 0
        If Message = "1" Then
            'Add Post
            ReturnVar = 88
            SendMSG "Please Enter your Review:"
        ElseIf Message = "2" Then
            'Back
            ProcessCommand "REVIEWS"
        End If
    Case 88
        'Add Post
        ReturnVar = 0
        Call AddFilmR(CurFilm, PrefGetName(buddy) & ": " & Message)
        SendMSG "Review added to Film:" & vbNewLine & Reviews(CurFilm).Title
    Case 89
        'Trailers
        ReturnVar = 0
        If Val(Message) <= 0 Or Val(Message) > UBound(TrailerName()) Then
            SendMSG "Invalid Selection"
        Else
            SendMSG "Trailer:" & vbNewLine & "The Trailer for this film can be viewed at:" & vbNewLine & vbNewLine & TrailerLink(Val(Message))
        End If
    Case 90
        'Add Admin
        ReturnVar = 0
        If UCase(Message) = "CASE90" Then
            ReturnVar = 91
            SendMSG "Please enter the name of the person you wish to be an admin"
        Else
            SendMSG "Incorrect Password"
        End If
    Case 91
        'Add Admin
        ReturnVar = 0
        If InStr(1, Message, "@") > 0 Then
            SendMSG "User has been warned"
            AddAdmin Message
            frmMain.TellMod PrefGetName(Message) & " has become an Admin"
        Else
            SendMSG "Not a valid email address", 4
        End If
    End Select
End Sub

Public Function Reserve(Text)
    Dim OutStr As String
    Dim X As Integer
    For X = 1 To Len(Text)
        OutStr = OutStr & Mid(Text, Len(Text) + 1 - X, 1)
    Next
    Reserve = OutStr
End Function

Public Function SendMSG(Message As String, Optional Colour = 0)
    Dim Col As String
    'default
    Col = "00"
    
    Select Case Colour
    Case 0
        'Black
        Col = "00"
    Case 1
        'Orange
        Col = "0066FF"
    Case 2
        'Red
        Col = "FF"
    Case 3
        'Blue
        Col = "CC3333"
    Case 4
        'Green
        Col = "7000"
    Case 5
        'Pink/Purple
        Col = "9933CC"
    Case 6
        'Light Blue
        Col = "CCCC99"
    Case 7
        'Light Green
        Col = "339933"
    Case 8
        'Darkish Blue
        Col = "996600"
        
    End Select
    
    MsgOut = MsgOut + 1
    strHeader = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=Verdana; EF=B; CO=" & Col & "; CS=0; PF=22" & vbCrLf & vbCrLf
    strHeader = strHeader & Message
    strHeader = "MSG " & myTrID1 & " N " & Len(strHeader) & vbCrLf & strHeader
    sndstr strHeader
End Function

Private Sub txtMSG_Change()
    If Len(txtMSG.Text) > 0 Then cmdsend.Enabled = True Else cmdsend.Enabled = False
End Sub


Private Sub txtMSG_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsend_Click
End Sub

Sub GameEngine(X, Y, CurrentUser, ClearOrPlace)    'This is the othello game engine
    Dim Total As Integer
    Dim Crossed As Integer
    Dim Found As Integer
    'This Engine will check a square for possible square
    'Or it will be sent the Co-Ords of a square and will plot counters
    
    'This works by a recursive search
    
    'Say you start off with a square 4,2
    
    '* = You
    '+ = Enemy
    '# = Empty space
    
    ' # # # #
    ' # # # #
    ' # # + #
    ' # + # #
    ' * # # #
    
    'The computer will search around it and find any possible enemy squares
    'in this case 3,3
    'It will carry on in this direction until it finds another counter of the starting counter
    'It will SelectableSquares the Co-Ords of this square in an Array
    
    'It then can draw the pictures and update the grid
    
    ' # # # #
    ' # # # *
    ' # # * #
    ' # * # #
    ' * # # #
    
    'In this Engine there are nine different searches and they go in this order
    
    ' 1 2 3
    ' 4   5
    ' 6 7 8
    
    'The middle is left out because there is no need for this

    Total = 0
    Win = 0
    Call Clearcheck
    If X - 1 > 0 And Y - 1 > 0 Then
        If OthelloGrid(X - 1, Y - 1) = CurrentUser Then
        
                
                Crossed = 1
                If ClearOrPlace = 2 Then
                    checking(X - 1, Y - 1) = 1
                End If
                
            For Found = 1 To 8
                If X - (1 + Found) > 0 And Y - (1 + Found) > 0 Then
                    If OthelloGrid(X - (1 + Found), Y - (1 + Found)) = CurrentUser Then
                        Crossed = Crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(X - (1 + Found), Y - (1 + Found)) = 1
                        End If
                    ElseIf OthelloGrid(X - (1 + Found), Y - (1 + Found)) = 3 - CurrentUser Then
                        Total = Total + Crossed
                        If ClearOrPlace = 2 Then
                            UpdateGrid (CurrentUser)
                        End If
                    ElseIf OthelloGrid(X - (1 + Found), Y - (1 + Found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        Call Clearcheck
    End If
    
    If X > 0 And Y - 1 > 0 Then
        If OthelloGrid(X, Y - 1) = CurrentUser Then
        
            Crossed = 1
            If ClearOrPlace = 2 Then
                checking(X, Y - 1) = 1
            End If
            
            For Found = 1 To 8
                If X > 0 And Y - (1 + Found) > 0 Then
                    If OthelloGrid(X, Y - (1 + Found)) = CurrentUser Then
                        Crossed = Crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(X, Y - (1 + Found)) = 1
                        End If
                    ElseIf OthelloGrid(X, Y - (1 + Found)) = 3 - CurrentUser Then
                        Total = Total + Crossed
                        If ClearOrPlace = 2 Then
                            UpdateGrid (CurrentUser)
                        End If
                    ElseIf OthelloGrid(X, Y - (1 + Found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        Call Clearcheck
    End If
    
    If X + 1 <= 8 And Y - 1 > 0 Then
        If OthelloGrid(X + 1, Y - 1) = CurrentUser Then
        
            Crossed = 1
            If ClearOrPlace = 2 Then
                checking(X + 1, Y - 1) = 1
            End If
            
            For Found = 1 To 8
                If X + (1 + Found) <= 8 And Y - (1 + Found) > 0 Then
                    If OthelloGrid(X + (1 + Found), Y - (1 + Found)) = CurrentUser Then
                        Crossed = Crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(X + (1 + Found), Y - (1 + Found)) = 1
                        End If
                    ElseIf OthelloGrid(X + (1 + Found), Y - (1 + Found)) = 3 - CurrentUser Then
                        Total = Total + Crossed
                        If ClearOrPlace = 2 Then
                            UpdateGrid (CurrentUser)
                        End If
                    ElseIf OthelloGrid(X + (1 + Found), Y - (1 + Found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        Call Clearcheck
    End If
    
    If X - 1 > 0 And Y > 0 Then
        If OthelloGrid(X - 1, Y) = CurrentUser Then
        
            Crossed = 1
            If ClearOrPlace = 2 Then
                checking(X - 1, Y) = 1
            End If
            
            For Found = 1 To 8
                If X - (1 + Found) > 0 And Y > 0 Then
                    If OthelloGrid(X - (1 + Found), Y) = CurrentUser Then
                        Crossed = Crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(X - (1 + Found), Y) = 1
                        End If
                    ElseIf OthelloGrid(X - (1 + Found), Y) = 3 - CurrentUser Then
                        Total = Total + Crossed
                        If ClearOrPlace = 2 Then
                            UpdateGrid (CurrentUser)
                        End If
                    ElseIf OthelloGrid(X - (1 + Found), Y) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        Call Clearcheck
    End If
     
    If X + 1 <= 8 And Y > 0 Then
        If OthelloGrid(X + 1, Y) = CurrentUser Then
        
            Crossed = 1
            If ClearOrPlace = 2 Then
                checking(X + 1, Y) = 1
            End If
            
            For Found = 1 To 8
                If X + (1 + Found) <= 8 And Y > 0 Then
                    If OthelloGrid(X + (1 + Found), Y) = CurrentUser Then
                        Crossed = Crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(X + (1 + Found), Y) = 1
                        End If
                    ElseIf OthelloGrid(X + (1 + Found), Y) = 3 - CurrentUser Then
                        Total = Total + Crossed
                        If ClearOrPlace = 2 Then
                            UpdateGrid (CurrentUser)
                        End If
                    ElseIf OthelloGrid(X + (1 + Found), Y) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        Call Clearcheck
    End If
     
    If X - 1 > 0 And Y + 1 <= 8 Then
        If OthelloGrid(X - 1, Y + 1) = CurrentUser Then
        
            Crossed = 1
            If ClearOrPlace = 2 Then
                checking(X - 1, Y + 1) = 1
            End If
            
            For Found = 1 To 8
                If X - (1 + Found) > 0 And Y + (1 + Found) <= 8 Then
                    If OthelloGrid(X - (1 + Found), Y + (1 + Found)) = CurrentUser Then
                        Crossed = Crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(X - (1 + Found), Y + (1 + Found)) = 1
                        End If
                    ElseIf OthelloGrid(X - (1 + Found), Y + (1 + Found)) = 3 - CurrentUser Then
                        Total = Total + Crossed
                        If ClearOrPlace = 2 Then
                            UpdateGrid (CurrentUser)
                        End If
                    ElseIf OthelloGrid(X - (1 + Found), Y + (1 + Found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        Call Clearcheck
    End If
     
    If X > 0 And Y + 1 <= 8 Then
        If OthelloGrid(X, Y + 1) = CurrentUser Then
        
            Crossed = 1
            If ClearOrPlace = 2 Then
                checking(X, Y + 1) = 1
            End If
            
            For Found = 1 To 8
                If X > 0 And Y + (1 + Found) <= 8 Then
                    If OthelloGrid(X, Y + (1 + Found)) = CurrentUser Then
                        Crossed = Crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(X, Y + (1 + Found)) = 1
                        End If
                    ElseIf OthelloGrid(X, Y + (1 + Found)) = 3 - CurrentUser Then
                        Total = Total + Crossed
                        If ClearOrPlace = 2 Then
                            UpdateGrid (CurrentUser)
                        End If
                    ElseIf OthelloGrid(X, Y + (1 + Found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        Call Clearcheck
    End If
     
    If X + 1 <= 8 And Y + 1 <= 8 Then
        If OthelloGrid(X + 1, Y + 1) = CurrentUser Then
        
            Crossed = 1
            If ClearOrPlace = 2 Then
                checking(X + 1, Y + 1) = 1
            End If
            
            For Found = 1 To 8
                If X + (1 + Found) <= 8 And Y + (1 + Found) <= 8 Then
                    If OthelloGrid(X + (1 + Found), Y + (1 + Found)) = CurrentUser Then
                        Crossed = Crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(X + (1 + Found), Y + (1 + Found)) = 1
                        End If
                    ElseIf OthelloGrid(X + (1 + Found), Y + (1 + Found)) = 3 - CurrentUser Then
                        Total = Total + Crossed
                        If ClearOrPlace = 2 Then
                            UpdateGrid (CurrentUser)
                        End If
                    ElseIf OthelloGrid(X + (1 + Found), Y + (1 + Found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    If ClearOrPlace = 1 Then
        SelectableSquares(X, Y) = Total
    Else
        Call Clearcheck
        SquarePotent = Total
        If Win = 1 Then
            checking(X, Y) = 1
            If CurrentUser = 1 Then
                OthelloGrid(X, Y) = 2
            Else
                OthelloGrid(X, Y) = 1
            End If
            GoodShot = 1
        ElseIf Win = 0 Then
            GoodShot = 0
        End If
    End If
End Sub

Sub Clearcheck()    'This clears the check array
    Dim X As Integer
    Dim Y As Integer
    For X = 1 To 8
        For Y = 1 To 8
            If checking(X, Y) = 1 Then
                checking(X, Y) = 0  'Sets the array to 0
            End If
        Next
    Next
End Sub

Sub Clearbrain()    'This clears the brain array
    Dim X As Integer
    Dim Y As Integer
    For X = 1 To 8
        For Y = 1 To 8
            If ahead(X, Y) = 1 Then
                ahead(X, Y) = 0     'Sets the array to 0
            End If
        Next
    Next
End Sub

Sub Comp()
    Dim X As Integer
    Dim Y As Integer
    Dim BestXPos As Integer
    Dim BestYPos As Integer
    Dim Bestscore As Integer
    For X = 1 To 8
        For Y = 1 To 8
            SelectableSquares(X, Y) = 0  'Clears the SelectableSquares array
            If OthelloGrid(X, Y) = 0 Then
                Call GameEngine(X, Y, 2, 1)    'calls the othello game engine to check square
            End If
        Next
    Next
    
    BestXPos = 1    'Set Default
    BestYPos = 1    'Set Default
    Bestscore = 0   'Set Default
    
    For X = 1 To 8
        For Y = 1 To 8
            If SelectableSquares(X, Y) > Bestscore Then
                BestXPos = X
                BestYPos = Y
                Call Clearbrain
                Bestscore = SelectableSquares(X, Y)
            End If
        Next
    Next
    
    If BestXPos = 1 And BestYPos = 1 And Bestscore = 0 Then
        SendMSG "Computer is unable to find move"
        If CheckWin(1) = False Then
            SendMSG DrawOthello & vbNewLine & vbNewLine & "Choice:"
        Else
            ReturnVar = 0
        End If
    Else
        Call Clearbrain
        Call GameEngine(BestXPos, BestYPos, 2, 2)
        If CheckWin(1) = False Then
            SendMSG DrawOthello & vbNewLine & vbNewLine & "Choice:"
        Else
            ReturnVar = 0
        End If
    End If
End Sub

Function CheckWin(PlayerCounter) As Boolean
    Dim X As Integer
    Dim Y As Integer
    Dim NOB As Integer
    Dim NOW As Integer
    Dim strData As String
    
    For X = 1 To 8
        For Y = 1 To 8
            If OthelloGrid(X, Y) = 2 Then
                NOB = NOB + 1
            ElseIf OthelloGrid(X, Y) = 1 Then
                NOW = NOW + 1
            End If
        Next
    Next
    
    If NOB + NOW = 64 Then    'Is grid fill
        CheckWin = True
        strData = "Game Over" & vbNewLine
        If NOB > NOW Then
            strData = strData & "Player Wins"
        ElseIf NOB < NOW Then
            strData = strData & "Computer Wins"
        ElseIf NOB = NOW Then
            strData = strData & "Draw"
        End If
        SendMSG strData
    ElseIf NOB = 0 Then    'Has all users been eliminated
        CheckWin = True
        strData = "Game Over" & vbNewLine & "Computer Wins"
    ElseIf NOW = 0 Then    'Has all user's been eliminated
        CheckWin = True
        strData = "Game Over" & vbNewLine & "Player Wins"
    Else
        CheckWin = False
    End If
End Function

Sub UpdateGrid(CurrentUser)
    Win = 1
    Dim X As Integer
    Dim Y As Integer
    For X = 1 To 8
        For Y = 1 To 8
            If checking(X, Y) = 1 Then
                If CurrentUser = 1 Then
                    OthelloGrid(X, Y) = 2
                Else
                    OthelloGrid(X, Y) = 1
                End If
            End If
        Next
    Next
End Sub

Function ShowWord()
    Dim Output As String
    Dim Found As Boolean
    Dim X As Integer
    Dim Y As Integer
    For X = 1 To Len(HWord)
        If UBound(Tried()) = 0 Then
            Output = Output & "_"
        Else
            Found = False
            For Y = 1 To UBound(Tried())
                If Tried(Y) = UCase(Mid(HWord, X, 1)) Then
                    Found = True
                    Output = Output & Mid(HWord, X, 1)
                    Exit For
                End If
            Next
            If Found = False Then
                Output = Output & "_"
            End If
        End If
    Next
    ShowWord = Output
End Function

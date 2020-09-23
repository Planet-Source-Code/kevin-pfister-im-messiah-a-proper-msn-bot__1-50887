VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   8265
   Begin VB.CommandButton CmdOk 
      Caption         =   "Done"
      Height          =   375
      Left            =   6720
      TabIndex        =   28
      Top             =   6600
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmProperties.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TxtWelNote"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TxtBotUpdates"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Users"
      TabPicture(1)   =   "FrmProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "LstUsers"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LstAdmins"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CmdAddAdmin"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "CmdRevAdmin"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "LstMod"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "CmdAddMod"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "CmdRevMod"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "LstBan"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "CmdAddBan"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "CmdRevBan"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Poll"
      TabPicture(2)   =   "FrmProperties.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "CmdPollWiz"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "TxtPollCode"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "CmdRemovePoll"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "CmdClearVotes"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "TxtPollInfo"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "LstPollVotes"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.TextBox TxtBotUpdates 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   3720
         Width           =   7815
      End
      Begin VB.TextBox TxtWelNote 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   720
         Width           =   7815
      End
      Begin VB.CommandButton CmdRevBan 
         Caption         =   "<"
         Height          =   375
         Left            =   -71880
         TabIndex        =   23
         Top             =   5520
         Width           =   375
      End
      Begin VB.CommandButton CmdAddBan 
         Caption         =   ">"
         Height          =   375
         Left            =   -71880
         TabIndex        =   22
         Top             =   4920
         Width           =   375
      End
      Begin VB.ListBox LstBan 
         Height          =   1620
         Left            =   -71400
         TabIndex        =   20
         Top             =   4560
         Width           =   4335
      End
      Begin VB.CommandButton CmdRevMod 
         Caption         =   "<"
         Height          =   375
         Left            =   -71880
         TabIndex        =   19
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton CmdAddMod 
         Caption         =   ">"
         Height          =   375
         Left            =   -71880
         TabIndex        =   18
         Top             =   3000
         Width           =   375
      End
      Begin VB.ListBox LstMod 
         Height          =   1620
         Left            =   -71400
         TabIndex        =   16
         Top             =   2640
         Width           =   4335
      End
      Begin VB.CommandButton CmdRevAdmin 
         Caption         =   "<"
         Height          =   375
         Left            =   -71880
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton CmdAddAdmin 
         Caption         =   ">"
         Height          =   375
         Left            =   -71880
         TabIndex        =   14
         Top             =   1080
         Width           =   375
      End
      Begin VB.ListBox LstAdmins 
         Height          =   1620
         Left            =   -71400
         TabIndex        =   12
         Top             =   720
         Width           =   4335
      End
      Begin VB.ListBox LstUsers 
         Height          =   5520
         Left            =   -74880
         TabIndex        =   10
         Top             =   720
         Width           =   2895
      End
      Begin VB.ListBox LstPollVotes 
         Height          =   2790
         Left            =   -71760
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox TxtPollInfo 
         Height          =   2775
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton CmdClearVotes 
         Caption         =   "Clear Votes"
         Height          =   375
         Left            =   -68520
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdRemovePoll 
         Caption         =   "Remove Poll"
         Height          =   375
         Left            =   -68520
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TxtPollCode 
         Height          =   2415
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   3840
         Width           =   6255
      End
      Begin VB.CommandButton CmdPollWiz 
         Caption         =   "Poll Wizard"
         Height          =   375
         Left            =   -68520
         TabIndex        =   1
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Bot Updates:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Welcome Note:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Banned Users:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   21
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Moderators:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   17
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Admins:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   13
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Registered Bot Users:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Users that have voted:"
         Height          =   255
         Left            =   -71760
         TabIndex        =   9
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Poll Code:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   7
         Top             =   3600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Poll Information:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   1125
      End
   End
End
Attribute VB_Name = "FrmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAddAdmin_Click()
    Dim Found As Boolean
    If LstUsers.Text <> "" Then
        If LstAdmins.ListCount = 0 Then
            LstAdmins.AddItem LstUsers.Text
        Else
            Found = False
            For X = 1 To LstAdmins.ListCount
                If UCase(LstAdmins.List(X)) = UCase(LstUsers.Text) Then
                    Found = True
                End If
            Next
            If Found = False Then
                LstAdmins.AddItem LstUsers.Text
            End If
        End If
    End If
End Sub

Private Sub CmdAddBan_Click()
    Dim Found As Boolean
    If LstUsers.Text <> "" Then
        If LstBan.ListCount = 0 Then
            LstBan.AddItem LstUsers.Text
        Else
            Found = False
            For X = 1 To LstBan.ListCount
                If UCase(LstBan.List(X)) = UCase(LstUsers.Text) Then
                    Found = True
                End If
            Next
            If Found = False Then
                LstBan.AddItem LstUsers.Text
            End If
        End If
    End If
End Sub

Private Sub CmdAddMod_Click()
    Dim Found As Boolean
    If LstUsers.Text <> "" Then
        If LstMod.ListCount = 0 Then
            LstMod.AddItem LstUsers.Text
        Else
            Found = False
            For X = 1 To LstMod.ListCount
                If UCase(LstMod.List(X)) = UCase(LstUsers.Text) Then
                    Found = True
                End If
            Next
            If Found = False Then
                LstMod.AddItem LstUsers.Text
            End If
        End If
    End If
End Sub

Private Sub CmdClearVotes_Click()
    If UBound(PollVotes()) <> 0 Then
        For X = 1 To UBound(PollVotes())
            PollVotes(X) = 0
        Next
    End If
    If UBound(PollVoted()) <> 0 Then
        For X = 1 To UBound(PollVoted())
            PollVoted(X) = ""
        Next
    End If
    Call SavePoll
    
    TxtPollInfo.Text = ReturnPoll
    TxtPollCode.Text = GetPollCode
    LstPollVotes.Clear
End Sub

Private Sub CmdOk_Click()
    Unload Me
End Sub

Private Sub CmdPollWiz_Click()
    'Enter the Poll Wizard
    FrmPollWiz.Show
End Sub

Private Sub CmdRemovePoll_Click()
    PollTitle = ""
    ReDim PollPosts(0) As String
    ReDim PollVotes(0) As Integer
    ReDim PollVoted(0) As String
    
    Call SavePoll
    TxtPollInfo.Text = ReturnPoll
    TxtPollCode.Text = ""
    LstPollVotes.Clear
End Sub

Private Sub CmdRevAdmin_Click()
    If LstAdmins.Text <> "" Then
        LstAdmins.RemoveItem (LstAdmins.ListIndex)
    End If
End Sub

Private Sub CmdRevBan_Click()
    If LstBan.Text <> "" Then
        LstBan.RemoveItem (LstBan.ListIndex)
    End If
End Sub

Private Sub CmdRevMod_Click()
    If LstMod.Text <> "" Then
        LstMod.RemoveItem (LstMod.ListIndex)
    End If
End Sub

Private Sub Form_Load()
    'Centre Form
    Me.Left = FrmMessiah.Width / 2 - Me.Width / 2
    Me.Top = 0
    
    'Load Settings
    TxtWelNote.Text = EditorsNote
    TxtBotUpdates.Text = BotUpdates
    
    Call SavePoll
    TxtPollInfo.Text = ReturnPoll
    TxtPollCode.Text = GetPollCode
    If UBound(PollVoted()) <> 0 Then
        For X = 1 To UBound(PollVoted())
            LstPollVotes.AddItem PrefGetName(PollVoted(X))
        Next
    End If
    
    If UBound(Admins()) <> 0 Then
        For X = 1 To UBound(Admins())
            LstAdmins.AddItem PrefGetName(Admins(X))
            DoEvents
        Next
    End If
    
    If UBound(Users()) <> 0 Then
        For X = 1 To UBound(Users())
            LstMod.AddItem PrefGetName(Users(X))
            DoEvents
        Next
    End If
    
    If UBound(BanList()) <> 0 Then
        For X = 1 To UBound(BanList())
            LstBan.AddItem PrefGetName(BanList(X))
            DoEvents
        Next
    End If

    If UBound(Everyone()) <> 0 Then
        For X = 1 To UBound(Everyone())
            LstUsers.AddItem PrefGetName(Everyone(X))
            DoEvents
        Next
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save Settings
    Call SaveEdNote
    
    Call SavePollCode(TxtPollCode.Text)
    Call LoadPoll
    Call SaveBotUpdates(TxtBotUpdates.Text)
    Call loadBotUD
    
    Open App.Path & "\data\Admins.dat" For Output As #1
    If LstAdmins.ListCount <> 0 Then
        For X = 1 To LstAdmins.ListCount
            Print #1, LstAdmins.List(X)
        Next
    End If
    Close
    
    Open App.Path & "\data\Users.dat" For Output As #1
    If LstMod.ListCount <> 0 Then
        For X = 1 To LstMod.ListCount
            Print #1, LstMod.List(X)
        Next
    End If
    Close

    Open App.Path & "\data\BanList.dat" For Output As #1
    If LstBan.ListCount <> 0 Then
        For X = 1 To LstBan.ListCount
            Print #1, LstBan.List(X)
        Next
    End If
    Close
    
    Call LoadAdmins
    Call LoadUsers
    Call LoadBans
    
End Sub

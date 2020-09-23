VERSION 5.00
Begin VB.Form FrmPollWiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Poll Wizard"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6720
   Begin VB.Frame Frame 
      Height          =   3735
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   6495
      Begin VB.Label Label9 
         Caption         =   "Wizard is complete, please click Next to go back to preferences"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame 
      Height          =   3735
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   6495
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   5040
         ScaleHeight     =   2535
         ScaleWidth      =   1335
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
         Begin VB.CommandButton CmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   1335
         End
         Begin VB.CommandButton CmdMoveUp 
            Caption         =   "Move Up"
            Height          =   375
            Left            =   0
            TabIndex        =   21
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton CmdMoveDown 
            Caption         =   "Move Down"
            Height          =   375
            Left            =   0
            TabIndex        =   20
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton CmdRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   0
            TabIndex        =   19
            Top             =   1440
            Width           =   1335
         End
      End
      Begin VB.ListBox LstResponses 
         Height          =   2010
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox TxtResponse 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label6 
         Caption         =   "Response:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Step 2. Adding the responses"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame 
      Height          =   3735
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   6495
      Begin VB.TextBox TxtPollName 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Text            =   "My Poll"
         Top             =   1080
         Width           =   6255
      End
      Begin VB.Label Label5 
         Caption         =   "Please Enter the question that the poll is asking in the box below"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   6255
      End
      Begin VB.Label Label2 
         Caption         =   "Step 1. Poll Question"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame 
      Height          =   3735
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6495
      Begin VB.Label Label1 
         Caption         =   $"FrmPollWiz.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IM Messiah"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Poll Wizard"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "FrmPollWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurFrame As Integer

Private Sub CmdAdd_Click()
    If TxtResponse.Text = "" Then
        Call MsgBox("Invalid Response", vbInformation)
        Exit Sub
    End If
    
    Dim Found As Boolean
    
    If LstResponses.ListCount = 0 Then
        LstResponses.AddItem TxtResponse.Text
    Else
        Found = False
        For X = 1 To LstResponses.ListCount
            If UCase(LstResponses.List(X - 1)) = UCase(TxtResponse.Text) Then
                Found = True
            End If
        Next
        If Found = False Then
            LstResponses.AddItem TxtResponse.Text
        End If
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdMoveDown_Click()
    If LstResponses.Text = "" Then
        Exit Sub
    End If
    If LstResponses.ListIndex = LstResponses.ListCount - 1 Then
        Exit Sub
    End If
    TempName = LstResponses.List(LstResponses.ListIndex + 1)
    LstResponses.List(LstResponses.ListIndex + 1) = LstResponses.List(LstResponses.ListIndex)
    LstResponses.List(LstResponses.ListIndex) = TempName
    LstResponses.Selected(LstResponses.ListIndex + 1) = True
End Sub

Private Sub CmdMoveUp_Click()
    If LstResponses.Text = "" Then
        Exit Sub
    End If
    If LstResponses.ListIndex = 0 Then
        Exit Sub
    End If
    TempName = LstResponses.List(LstResponses.ListIndex - 1)
    LstResponses.List(LstResponses.ListIndex - 1) = LstResponses.List(LstResponses.ListIndex)
    LstResponses.List(LstResponses.ListIndex) = TempName
    LstResponses.Selected(LstResponses.ListIndex - 1) = True
End Sub

Private Sub CmdNext_Click()
    CurFrame = CurFrame + 1
    
    If CurFrame = 2 Then
        If TxtPollName.Text = "" Then
            Call MsgBox("The Polls Question must be given", vbInformation)
            CurFrame = 1
        End If
    End If
    
    If CurFrame = 3 Then
        If LstResponses.ListCount = 0 Then
            Call MsgBox("Poll Responses must be given", vbInformation)
            CurFrame = 2
        End If
    End If
    
    If CurFrame = 4 Then
        FrmProperties.TxtPollCode = "A" & TxtPollName.Text
        For X = 1 To LstResponses.ListCount
            FrmProperties.TxtPollCode = FrmProperties.TxtPollCode & vbNewLine & "B" & LstResponses.List(X - 1) & vbNewLine & "C0"
        Next
        Call SavePollCode(FrmProperties.TxtPollCode.Text)
        Call LoadPoll
        FrmProperties.TxtPollInfo.Text = ReturnPoll
        If UBound(PollVoted()) <> 0 Then
            For X = 1 To UBound(PollVoted())
                FrmProperties.LstPollVotes.AddItem PollVoted(X)
            Next
        End If
        Unload Me
    Else
        Frame(CurFrame - 1).Visible = False
        Frame(CurFrame).Visible = True
    End If
End Sub

Private Sub CmdRemove_Click()
    If LstResponses.Text = "" Then
        Exit Sub
    End If
    LstResponses.RemoveItem LstResponses.ListIndex
End Sub

Private Sub Form_Load()
    'Centre Form
    Me.Left = FrmMessiah.Width / 2 - Me.Width / 2
    Me.Top = 0
    
    CurFrame = 0
    For X = 0 To Frame().Count - 1
        Frame(X).Visible = False
    Next
    Frame(0).Visible = True
End Sub

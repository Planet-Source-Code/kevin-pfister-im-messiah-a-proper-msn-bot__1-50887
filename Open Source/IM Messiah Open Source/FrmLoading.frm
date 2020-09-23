VERSION 5.00
Begin VB.Form FrmLoading 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading Data..."
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.PictureBox PicLoad 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   4920
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   720
      Width           =   900
   End
   Begin VB.Timer TmrLoadData 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   1920
   End
   Begin VB.PictureBox PicFrom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4920
      ScaleHeight     =   375
      ScaleWidth      =   1350
      TabIndex        =   6
      Top             =   240
      Width           =   1350
   End
   Begin VB.Timer TmrLoad 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   1920
   End
   Begin VB.PictureBox PicStatus 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
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
         Left            =   1200
         TabIndex        =   2
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Data"
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
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Label LblStatus 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label lblLoading 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "FrmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurX As Integer

Private Sub Form_Load()
    'Centre Form
    Me.Left = FrmMessiah.Width / 2 - Me.Width / 2
    Me.Top = 0
    Randomize Timer
    
    'Close any open File
    Close
    For X = 0 To 90
        Colour = RGB(255 - Abs(Sin(X / (90 / 3.14159)) * 150), 255 - Abs(Sin(X / (90 / 3.14159)) * 150), 255 - Abs(Sin(X / (90 / 3.14159)) * 150))
        For Y = 0 To PicFrom.Height
            Call SetPixelV(PicFrom.HDC, X, Y, Colour)
        Next
    Next
    TmrLoad.Enabled = True
End Sub

Private Sub TmrLoad_Timer()
    CurX = CurX + 1
    If CurX = PicStatus.Width + 1 Then CurX = 1
    PicStatus.Cls
    Call BitBlt(PicStatus.HDC, CurX, 0, 90, PicStatus.Height, PicFrom.HDC, 0, 0, vbSrcCopy)
    If CurX + 90 > PicStatus.Width Then
        Start = -90 + (CurX + 90 - PicStatus.Width)
        Call BitBlt(PicStatus.HDC, Start, 0, 90, PicStatus.Height, PicFrom.HDC, 0, 0, vbSrcCopy)
    End If
    PicStatus.Refresh
End Sub

Private Sub TmrLoadData_Timer()
    'Load Everything
    LblStatus.Caption = "Loading General Settinsgs"
    DoEvents
    loadGenSets
    
    LblStatus.Caption = "Loading Admin List"
    DoEvents
    LoadAdmins
    
    LblStatus.Caption = "Loading Mod List"
    DoEvents
    LoadUsers
    
    LblStatus.Caption = "Loading Quotes"
    DoEvents
    LoadQuotes
    
    LblStatus.Caption = "Loading Ban List"
    DoEvents
    LoadBans

    LblStatus.Caption = "Loading Suggestion"
    DoEvents
    LoadSuggestions

    LblStatus.Caption = "Loading Jokes"
    DoEvents
    LoadJokes

    LblStatus.Caption = "Loading Funny Sites"
    DoEvents
    LoadSites

    LblStatus.Caption = "Loading MSN Nicknames"
    DoEvents
    LoadNicks

    LblStatus.Caption = "Loading Old News"
    DoEvents
    loadNews

    LblStatus.Caption = "Loading Old News"
    DoEvents
    loadBotUD
    
    LblStatus.Caption = "Loading AI Chat Bot(Bob)"
    DoEvents
    LoadAI

    LblStatus.Caption = "Loading Hangman Dictionary"
    DoEvents
    LoadHangman

    LblStatus.Caption = "Loading Bots List"
    DoEvents
    LoadUserBots

    LblStatus.Caption = "Loading Other Languages"
    DoEvents
    LoadLang

    LblStatus.Caption = "Loading Welcome Note"
    DoEvents
    LoadEdNote

    LblStatus.Caption = "Loading Bot Forum"
    DoEvents
    LoadForum

    LblStatus.Caption = "Loading Poll System"
    DoEvents
    LoadPoll

    LblStatus.Caption = "Loading User Preferences"
    DoEvents
    LoadPref

    LblStatus.Caption = "Loading Film Reviews"
    DoEvents
    LoadReviews

    LblStatus.Caption = "Clearing Old Logs"
    DoEvents
    ClearLog

    LblStatus.Caption = "Loading Settings"
    DoEvents
    StartupVars

    LblStatus.Caption = "Downloading Current News"
    DoEvents
    CurNews = FrmMessiah.Inet.OpenURL("http://news.bbc.co.uk/text_only.stm")
    DoEvents
    st = Timer
    Do
        DoEvents
    Loop Until Timer - st > 1
    
    CurNews = Mid(CurNews, InStr(1, CurNews, "OTHER TOP STORIES") + 17)

    CurNews = Replace(CurNews, "<b>", "")
    CurNews = Replace(CurNews, "<p>", "")
    CurNews = Replace(CurNews, "</b>", "")
    CurNews = Replace(CurNews, "<br clear=" & "" & "all" & "" & " />", "")
    CurNews = Replace(CurNews, "</a><br />", "")
    CurNews = Replace(CurNews, vbTab, "")
    CurNews = Replace(CurNews, vbNewLine & vbNewLine, "")
    Do
        DoEvents
        If InStr(1, CurNews, "<a href=") > 0 Then
            If InStr(InStr(1, CurNews, "<a href="), CurNews, ">") > 0 Then
                CurNews = Mid(CurNews, 1, InStr(1, CurNews, "<a href=") - 1) & Mid(CurNews, InStr(InStr(1, CurNews, "<a href="), CurNews, ">") + 1)
            Else
                Exit Do
            End If
        End If
    Loop Until InStr(1, CurNews, "<a href=") = 0

    Do
        DoEvents
        If InStr(1, CurNews, "         ") > 0 Then
            If InStr(InStr(1, CurNews, "         "), CurNews, "</a><br />") > 0 Then
                CurNews = Mid(CurNews, 1, InStr(1, CurNews, "         ") - 1) & Mid(CurNews, InStr(InStr(1, CurNews, "         "), CurNews, "</a><br />") + 10)
            Else
                Exit Do
            End If
        End If
    Loop Until InStr(1, CurNews, "         ") = 0
    CurNews = Replace(CurNews, "        " & vbNewLine & "       " & vbNewLine, "")
    CurNews = Replace(CurNews, vbNewLine & vbNewLine, vbNewLine)
    If InStr(1, CurNews, "        " & vbNewLine & "       ") > 0 Then
        CurNews = Mid(CurNews, 1, InStr(1, CurNews, "        " & vbNewLine & "       ") - 1)
    End If
    CurNews = Date & vbNewLine & CurNews
    AddNews CurNews
    
    LblStatus.Caption = "Downloading Latest Trailers"
    'If there is an error make the following section into a comment using '
    
    'Start Section
    DoEvents
    strData = FrmMessiah.Inet.OpenURL("http://www.apple.com/trailers/")
    DoEvents
    st = Timer
    Do
        DoEvents
    Loop Until Timer - st > 1
    
    strData = Mid(strData, InStr(1, strData, "<!-- BEGIN NEWEST TRAILERS HERE -->") + Len("<!-- BEGIN NEWEST TRAILERS HERE -->"))
    strData = Mid(strData, 1, InStr(1, strData, "<!-- END NEW TRAILERS -->") - 1)
    'End Section
    
    Call SaveTrailors(strData)

    
    frmMain.Show
    Unload Me
End Sub

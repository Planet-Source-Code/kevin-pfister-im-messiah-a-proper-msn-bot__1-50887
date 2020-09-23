VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmContacts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IM Messiah"
   ClientHeight    =   4470
   ClientLeft      =   5430
   ClientTop       =   435
   ClientWidth     =   4965
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   Begin VB.Timer tmrerasehtm 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4080
      Top             =   4080
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2520
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":1190
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":1656
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":1B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":1FE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":24A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":296E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":2E34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView lstBuddy 
      Height          =   3990
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7038
      _Version        =   393217
      Indentation     =   176
      LineStyle       =   1
      Style           =   3
      ImageList       =   "imgList"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4725
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveAI
    frmMain.Winsock1_Close
    frmMain.Messenger.Close
    Unload frmMain
    Unload frmMSG
    Unload frmContacts
    Unload Me
    End
End Sub

Private Sub Form_Load()
    'Centre Form
    Me.Left = FrmMessiah.Width / 2 - Me.Width / 2
    
    
    AddToLog "Bot Logged on"
    UpdateStats
End Sub

Private Sub lblName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PopupMenu FrmMessiah.mnuStatusTop, , lblName.Left, lblName.Top + lblName.Height
End Sub

Private Sub lstBuddy_Click()
    On Error Resume Next
    If lstBuddy.SelectedItem.Key = "top" Then
        
    ElseIf (lstBuddy.SelectedItem.Key = "ofl") Or (lstBuddy.SelectedItem.Key = "onl") Then
        If lstBuddy.SelectedItem.Expanded = True Then
            lstBuddy.SelectedItem.Expanded = False
            lstBuddy.SelectedItem.Image = 2
            lstBuddy.SelectedItem = lstBuddy.Nodes.Item(1)
        ElseIf lstBuddy.SelectedItem.Expanded = False Then
            lstBuddy.SelectedItem.Expanded = True
            lstBuddy.SelectedItem.Image = 1
            lstBuddy.SelectedItem = lstBuddy.Nodes.Item(1)
        End If
    End If
End Sub

Public Sub lstBuddy_dblClick()
    If lstBuddy.SelectedItem.Key = "top" Or lstBuddy.SelectedItem.Key = "onl" Or lstBuddy.SelectedItem.Key = "ofl" Then
        CheckEmail frmMain.strUserName, frmMain.strPassword, frmMain.logintime, frmMain.MSPAuth, 0, "" 'Checking email..
    Else
        If lstBuddy.SelectedItem.Parent.Key = "onl" Then 'You clicked on someone online
            frmMain.Messenger.SendData "XFR " & frmMain.intTrailid & " SB" & vbCrLf
            frmMain.buddyconnect = lstBuddy.SelectedItem.Key
        ElseIf lstBuddy.SelectedItem.Parent.Key = "ofl" Then 'You clicked on someone offline
            MsgBox "You cannot send messages to someone that is offline.", vbOKOnly, "Error"
        End If
    End If
End Sub

Private Sub lstBuddy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If lstBuddy.SelectedItem.Key <> "top" And lstBuddy.SelectedItem.Key <> "onl" And lstBuddy.SelectedItem.Key <> "ofl" Then
            If lstBuddy.SelectedItem.Parent.Key = "onl" Or lstBuddy.SelectedItem.Parent.Key = "ofl" Then
                mnuViewP.Caption = "View Profile (" & lstBuddy.SelectedItem.Key & ")"
                PopupMenu mnuContacts, , X, Y + 600
            End If
        End If
    End If
End Sub

Private Sub tmrerasehtm_Timer()
    FileSystem.Kill "login.htm"
    tmrerasehtm.Enabled = False
End Sub

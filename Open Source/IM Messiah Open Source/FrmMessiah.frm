VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm FrmMessiah 
   BackColor       =   &H8000000C&
   Caption         =   "IM Messiah"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11685
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnuStatusTop 
         Caption         =   "Status"
         Begin VB.Menu mnuStatus 
            Caption         =   "Online"
            Index           =   1
         End
         Begin VB.Menu mnuStatus 
            Caption         =   "Busy"
            Index           =   2
         End
         Begin VB.Menu mnuStatus 
            Caption         =   "Be Right Back"
            Index           =   3
         End
         Begin VB.Menu mnuStatus 
            Caption         =   "Away"
            Index           =   4
         End
         Begin VB.Menu mnuStatus 
            Caption         =   "On The Phone"
            Index           =   5
         End
         Begin VB.Menu mnuStatus 
            Caption         =   "Out To Lunch"
            Index           =   6
         End
         Begin VB.Menu mnuStatus 
            Caption         =   "Appear Offline"
            Index           =   7
         End
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu MNUBot 
      Caption         =   "Bot"
      Begin VB.Menu MNUBotProp 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuContacts 
      Caption         =   "Contacts"
      Visible         =   0   'False
      Begin VB.Menu mnuSendIM 
         Caption         =   "Send an Instant Message"
      End
      Begin VB.Menu mnuSendF 
         Caption         =   "Send a File or Photo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSendE 
         Caption         =   "Send an Email"
      End
      Begin VB.Menu MNUAd 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewP 
         Caption         =   "View Profile()"
      End
      Begin VB.Menu mnuBlock 
         Caption         =   "Block"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MnuBotStat 
      Caption         =   "Status"
      Begin VB.Menu mnuIDEDebug 
         Caption         =   "View IDE Debug"
         Checked         =   -1  'True
      End
      Begin VB.Menu MNUChat 
         Caption         =   "Chat"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUpdateWeb 
         Caption         =   "UpdateWeb"
      End
   End
   Begin VB.Menu mnuBob 
      Caption         =   "Bob"
      Begin VB.Menu mnuLoadAI 
         Caption         =   "Load AI"
      End
      Begin VB.Menu mnuSaveAI 
         Caption         =   "Save AI"
      End
   End
End
Attribute VB_Name = "FrmMessiah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    FrmLoading.Show
    FrmLoading.TmrLoadData.Enabled = True
End Sub

Private Sub MNUBotProp_Click()
    FrmProperties.Show
End Sub

Private Sub MNUChat_Click()
    If MNUChat.Checked = True Then
        MeWannaChat = False
        MNUChat.Checked = False
    Else
        MeWannaChat = True
        MNUChat.Checked = True
    End If
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuIDEDebug_Click()
    If mnuIDEDebug = True Then
        mnuIDEDebug.Checked = False
        IDEDebug = False
    Else
        mnuIDEDebug.Checked = True
        IDEDebug = True
    End If
    Call SaveSetting("IM Messiah", "General", "IDEDEBUG", IDEDebug)
End Sub

Public Sub mnuStatus_Click(Index As Integer)
    status = Index
    Select Case Index
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
    For a = 1 To 7
        mnuStatus(a).Checked = False
    Next
    mnuStatus(Index).Checked = True
End Sub

Private Sub mnuUpdateWeb_Click()
    UpdateStats
End Sub

Private Sub mnuLoadAI_Click()
    LoadAI
End Sub

Private Sub mnuSaveAI_Click()
    SaveAI
End Sub

Private Sub MNUAd_Click()
    AdName = Mid(mnuViewP.Caption, 15)
    AdName = Mid(AdName, 1, Len(AdName) - 1)
    AddAdmin AdName
End Sub

Private Sub mnuSendE_Click()
    CheckEmail frmMain.strUserName, frmMain.strPassword, frmMain.logintime, frmMain.MSPAuth, 1, frmContacts.lstBuddy.SelectedItem.Key
End Sub

Private Sub mnuSendIM_Click()
    frmContacts.lstBuddy_dblClick
End Sub

Private Sub mnuViewP_Click()
    Dim strD As String
    strD = "http://members.msn.com/default.msnw?mem=" & frmContacts.lstBuddy.SelectedItem.Key & "&pgmarket="
    Shell "C:\program files\Internet Explorer\iexplore.exe " & strD, vbNormalFocus
End Sub

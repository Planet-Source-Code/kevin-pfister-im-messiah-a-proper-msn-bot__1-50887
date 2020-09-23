VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IM Messiah"
   ClientHeight    =   2250
   ClientLeft      =   6360
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4680
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4695
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Login Form"
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
         TabIndex        =   8
         Top             =   360
         Width           =   3615
      End
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
         TabIndex        =   7
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Timer tmrKeepAlive 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtPassword 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "IM_Messiah@hotmail.com"
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtOutput 
      Height          =   3765
      Left            =   5805
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   225
      Width           =   5535
   End
   Begin MSWinsockLib.Winsock Messenger 
      Left            =   360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Min             =   1
      Max             =   7
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "www.google.com"
      RemotePort      =   80
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strServer As String = "messenger.hotmail.com"
Const lngPort As Long = 1863

Public MSPAuth As String, logintime As String, kv As String, SID As String

Dim strCurrentServer As String
Dim lngCurrentPort As Long
Dim online As Integer

Public intTrailid As Integer
Dim intConnState As Integer

Public strUserName As String, strPassword As String
    
Dim strLastSendCMD As String

Dim frmMSG(Capacity) As New frmMSG
Dim listing As Boolean
Dim LSTContacts As String

Public buddyconnect As String

Dim DoneS As Boolean
Dim Result As String

Public Sub IncrementTrailID()
    intTrailid = intTrailid + 1
End Sub

Sub IncrementState()
    intConnState = intConnState + 1
End Sub

Sub ResetVars()
    
    intConnState = 0
    intTrailid = 1
    online = 0
    frmContacts.lstBuddy.Nodes.Clear
    LSTDone = False
    listing = False
    For X = 1 To Capacity
        On Error Resume Next
        Unload frmMSG(X)
    Next
    
End Sub

Public Sub ProcessData(strData As String)
    strBuffer = strBuffer & strData
End Sub

Private Sub CmdEnd_Click()
    Unload Me
    End
End Sub

Private Sub Command1_Click()
    tmrTimeout.Enabled = True
    prg.Visible = True
    prg.Value = 1
    ResetVars
    strPassword = txtPassword.Text
    strUserName = txtUsername.Text
    Messenger.Close
    Messenger.Connect strServer, lngPort
End Sub

Private Sub Form_Load()
    'Centre Form
    Me.Left = FrmMessiah.Width / 2 - Me.Width / 2
    Me.Top = 0
    
    Dim strData As String
    AddToLog "Bot Started"
    Debug.Print "Loaded"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmMSG(1)
    Unload frmContacts
    Unload Me
    End
End Sub

Private Sub addonline(prsline As String)
    If UBound(Split(prsline, " ")) > 4 Then
        Dim User As String, Email As String
        
        status = Split(prsline, " ")(2)
        Email = Split(prsline, " ")(3)
        User = Split(prsline, " ")(4)
        User = URLDecode(User)
        
        For Y = 1 To frmContacts.lstBuddy.Nodes.Count
            If frmContacts.lstBuddy.Nodes(Y).Key = Email Then Exists = True
        Next
        
        If Exists = False Then
            If status = "RL" Then frmContacts.lstBuddy.Nodes.Add "ofl", tvwChild, Email, User, 5
            If status = "NLN" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 4
            If status = "BSY" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 6
            If status = "BRB" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 7
            If status = "AWY" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 7
            If status = "PHN" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 6
            If status = "LUN" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 7
            If status = "IDL" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 7
        End If
    End If
End Sub

Private Sub updstatus(Data As String)
    Dim User As String, Email As String, status As String
    
    If Left(Data, 3) = "FLN" Then
        For X = 1 To frmContacts.lstBuddy.Nodes.Count
            Email = Split(Data, " ")(1)
            If frmContacts.lstBuddy.Nodes.Item(X).Key = Email Then
                User = frmContacts.lstBuddy.Nodes.Item(X).Text
                frmContacts.lstBuddy.Nodes.Remove (X)
                frmContacts.lstBuddy.Nodes.Add "ofl", tvwChild, Email, User, 5
            End If
        Next
    End If

    If UBound(Split(Data, " ")) > 3 Then
        User = Split(Data, " ")(3)
        User = URLDecode(User)
        Email = Split(Data, " ")(2)
        status = Split(Data, " ")(1)
        For X = 1 To frmContacts.lstBuddy.Nodes.Count
            If frmContacts.lstBuddy.Nodes.Item(X).Key = Email Then ' Found Contact
                If frmContacts.lstBuddy.Nodes.Item(X).Parent.Key = "ofl" Then ' Contact offline
                
                    If Not status = "FLN" Then ' New status is offline
                        frmContacts.lstBuddy.Nodes.Remove (X)
                        If status = "NLN" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 4
                        If status = "BSY" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 6
                        If status = "BRB" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 7
                        If status = "AWY" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 7
                        If status = "PHN" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 6
                        If status = "LUN" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 7
                        If status = "IDL" Then frmContacts.lstBuddy.Nodes.Add "onl", tvwChild, Email, User, 7
                    End If
                ElseIf frmContacts.lstBuddy.Nodes.Item(X).Parent.Key = "onl" Then 'Contact online
                
                frmContacts.lstBuddy.Nodes.Item(X).Text = User
                If status = "NLN" Then
                    frmContacts.lstBuddy.Nodes.Item(X).Image = 4
                ElseIf status = "BSY" Then
                    frmContacts.lstBuddy.Nodes.Item(X).Image = 6
                ElseIf status = "BRB" Then
                    frmContacts.lstBuddy.Nodes.Item(X).Image = 7
                ElseIf status = "AWY" Then
                    frmContacts.lstBuddy.Nodes.Item(X).Image = 7
                ElseIf status = "PHN" Then
                    frmContacts.lstBuddy.Nodes.Item(X).Image = 6
                ElseIf status = "LUN" Then
                    frmContacts.lstBuddy.Nodes.Item(X).Image = 7
                ElseIf status = "IDL" Then
                    frmContacts.lstBuddy.Nodes.Item(X).Image = 7
                End If
                
                X = frmContacts.lstBuddy.Nodes.Count
                End If
                
            End If
        Next
    End If
End Sub

Public Function contactexists(Email As String) As Boolean
    contactexists = False
    For Y = 1 To frmContacts.lstBuddy.Nodes.Count
        If frmContacts.lstBuddy.Nodes(Y).Key = Email Then contactexists = True
    Next
End Function

Private Sub addcontact(contact As String)
    Dim User As String, Email As String, Exists As Boolean, listnum As Integer

    If (UBound(Split(contact, " ")) > 2) Then
        User = Split(contact, " ")(2)
        User = URLDecode(User)
        Email = Split(contact, " ")(1)
        
        ReDim Preserve Everyone(UBound(Everyone()) + 1) As String
        Everyone(UBound(Everyone())) = Email
        
        If UBound(Split(contact, " ")) = 4 Then
        
            Exists = contactexists(Email)
        
            On Error Resume Next
            If Exists = False Then
                frmContacts.lstBuddy.Nodes.Add "ofl", tvwChild, Email, User, 5
            End If
            
        End If
    End If
End Sub

Private Sub Messenger_Connect()
    
    intConnState = 1
    Messenger_DataArrival 0

End Sub

Public Sub msgsend(Message As String)
    Messenger.SendData Message
    IncrementTrailID
    If IDEDebug = True Then
        Debug.Print Message & vbCrLf
    End If
End Sub

Private Sub Messenger_DataArrival(ByVal bytesTotal As Long)
    Dim strRawData As String, strInput As String
    Dim strHashParams As String
    Dim strResponse As String
    
    Dim varParams As Variant
      
    Messenger.GetData strRawData, vbString
    
    'txtOutput = txtOutput & strRawData
    
If intConnState > 6 Then

  For linenum = 0 To UBound(Split(strRawData, vbCrLf))
  
    strInput = Split(strRawData, vbCrLf)(linenum)
    
    If strInput <> "" Then
        If Split(strInput, " ")(0) = "Inbox-Unread:" Then frmContacts.lstBuddy.Nodes.Item(1).Text = "You have " & Split(strInput, " ")(1) & " new e-mail message(s)" 'Tells you how many unread emails you have
    End If

    Select Case Left(strInput, 3)
    Case "RNG":
        For numim = 1 To 50
            If frmMSG(numim).buddy = "" Then
                frmMSG(numim).buddy = Split(strInput, " ")(5)
                frmMSG(numim).strSID = Split(strInput, " ")(1)
                frmMSG(numim).strMIP = Replace(Split(strInput, " ")(2), ":1863", "")
                frmMSG(numim).strCKI = Split(strInput, " ")(4)
                frmMSG(numim).Hide
                numim = 50
            End If
        Next
        
    Case "CHL":
        Dim strCHL As String
        strCHL = Replace(Split(strInput, " ")(2), vbCrLf, "")
        msgsend "QRY " & intTrailid & " msmsgs@msnmsgr.com 32" & vbCrLf & MD5String(strCHL & "Q1P7W2E4J9R8U3S5")
        'Call IncrementTrailID
        
    Case "ILN":
        addonline (strInput)
        
    Case "XFR":
        If online = 1 Then
            Dim convoopen As Boolean
            For numim = 1 To Capacity
                If UCase(frmMSG(numim).buddy) = UCase(buddyconnect) Then
                    convoopen = True
                    frmMSG(numim).Show
                    frmMSG(numim).strSID = Split(strInput, " ")(1)
                    frmMSG(numim).strMIP = Replace(Split(strInput, " ")(2), ":1863", "")
                    frmMSG(numim).strCKI = Split(strInput, " ")(5)
                    serv1 = Split(strInput, " ")(3)
                    serv = Mid(serv1, 1, Len(serv1) - 5)
                    frmMSG(numim).clientside = 1
                    frmMSG(numim).sckMSG.Close
                    frmMSG(numim).sckMSG.Connect serv
                    numim = 50
                End If
            Next
            If convoopen = False Then
                For numim = 1 To Capacity
                If frmMSG(numim).buddy = "" Then
                    frmMSG(numim).buddy = buddyconnect
                    frmMSG(numim).strSID = Split(strInput, " ")(1)
                    frmMSG(numim).strMIP = Replace(Split(strInput, " ")(2), ":1863", "")
                    frmMSG(numim).Hide
                    serv1 = Split(strInput, " ")(3)
                    serv = Mid(serv1, 1, Len(serv1) - 5)
                    frmMSG(numim).strCKI = Split(strInput, " ")(5)
                    frmMSG(numim).clientside = 1
                    frmMSG(numim).sckMSG.Close
                    frmMSG(numim).sckMSG.Connect serv
                    numim = 50
                End If
                Next
            End If
            IncrementTrailID
        End If
        
    Case "CHG":
        If listing = False Then
            msgsend "SYN " & intTrailid & " 0" & vbCrLf
            listing = True
        End If
        
    Case "LST":
        RegUser = RegUser + 1
        If Right(strInput, 2) = "10" Then
            strInput = Mid(strInput, 1, Len(strInput) - 2) & "11 0"
        End If

        addcontact (strInput)
        
    Case "NLN":
        updstatus (strInput)
        
    Case "FLN":
        updstatus (strInput)
    'Select Case Left(strInput, 3)
    Case "ADD":
        
        dato = Split(strInput, " ")
        If dato(0) = "ADD" And dato(2) = "RL" Then
            NewContacts = NewContacts + 1
            frmMain.msgsend "ADD " & frmMain.intTrailid & " AL " & dato(4) & " " & dato(4) & vbCrLf
            Call IncrementTrailID
        End If
    Case "VER" 'Verification
        frmMain.msgsend "CVR % 0x0413 winnt 5.1 i386 MSNMSGR 6.1.0203 MSMSGS " & strUserName & vbCrLf
    Case Else
        Debug.Print "Else: |" & strInput & "|"
    End Select
    
  Next

  
End If


    Select Case intConnState

    Case 1
        
            ' Handshake
            '-----------------------------
            
        strLastSendCMD = "VER " & intTrailid & " MSNP9 MSNP8 CVR0" & vbCrLf
    
        msgsend strLastSendCMD
        
        'Call IncrementTrailID
        Call IncrementState
        
        If prg.Value = 0 Then prg.Value = 1
        
    Case 2
    
            ' Send client information to DS
            '-----------------------------

        If strRawData = strLastSendCMD Then
        
            strLastSendCMD = "CVR " & intTrailid & " 0x0413 winnt 5.2 i386 MSNMSGR 6.0.0268 MSMSGS " & strUserName & vbCrLf
            
            msgsend strLastSendCMD
            
            'Call IncrementTrailID
            Call IncrementState
            
        Else
        
            MsgBox "No support for this protocol."
            
        End If
        
        If prg.Value = 1 Then prg.Value = 2
        
    Case 3
    
    
            ' Send logonname (xxx@xxx.xxx) to DS
            '-----------------------------
        
        strLastSendCMD = "USR " & intTrailid & " TWN I " & strUserName & vbCrLf
        
        msgsend strLastSendCMD
        
        'Call IncrementTrailID
        Call IncrementState
    
        If prg.Value = 2 Then prg.Value = 3
    
    Case 4
    
    
            ' Send password to DS or move to other server
            '-----------------------------

        If UCase$(Left$(strRawData, 4)) = "USR " Then
        

            ' Get the hash supplied by the DS:
            h = InStr(LCase$(strRawData), " lc")
            strHashParams = Right$(strRawData, Len(strRawData) - h)
            
            ' Start the SSL-procedure:
            strResponse = DoSSL(strHashParams)
            
            ' Pass authentication result back to the DS:
            strLastSendCMD = "USR " & CStr(intTrailid) & " TWN S " & strResponse & vbCrLf
            
            msgsend strLastSendCMD
            
            'Call IncrementTrailID
            Call IncrementState
        
        ElseIf UCase$(Left(strRawData, 4)) = "XFR " Then
        
            ' Move to another server
            
            varParams = Split(strRawData, " ")
            strConnectionString = varParams(3)
            
            varParams = Split(strConnectionString, ":")
            strCurrentServer = varParams(0)
            lngCurrentPort = CLng(varParams(1))
            
            ResetVars
            
            Messenger.Close
            Messenger.Connect strCurrentServer, lngCurrentPort
        
        End If
        
        
    Case 5
    
    
            ' Authentication ok or failed?
            '-----------------------------
        If UCase$(Left$(strRawData, 4)) = "USR " Then
        
            Dim nn As String
            nn = Split(strRawData, " ")(4)
            frmContacts.lblName.Caption = URLDecode(nn)
            online = 1
        
            'MsgBox "You have logged on succesfully, you will become online after you hit the Ok-button."
            Call IncrementState
        
        ElseIf UCase$(Left$(strRawData, 4)) = "911 " Then
            
            MsgBox "Invalid password"
        
        End If
        
        
    Case 6
    
    
            ' Recieve some Hotmail garbage
            '-----------------------------
            
        If UCase$(Left$(strRawData, 4)) = "MSG " Then
            getvars (strRawData)
            FrmMessiah.mnuStatus_Click (7) 'appear offline
            
            prg.Value = 7
            tmrTimeout.Enabled = False
            tmrKeepAlive.Enabled = True
            frmContacts.Visible = True
            
            
            Me.Visible = False
                
            frmContacts.lstBuddy.Nodes.Add , , "top", "You have 0 new e-mail message(s)", 3
            frmContacts.lstBuddy.Nodes.Add , , "onl", "Online", 1 'Create Main Parent
            frmContacts.lstBuddy.Nodes.Add , , "ofl", "Offline", 2
            frmContacts.lstBuddy.Nodes.Item(2).Expanded = True
            frmContacts.lstBuddy.Nodes.Item(2).Sorted = True
            frmContacts.lstBuddy.Nodes.Item(3).Sorted = True
            
            'Call IncrementTrailID
            Call IncrementState
            
        Else
        
            Call IncrementState
            GoTo LoginDone
            
        End If
        
        
    Case 7
    
        ' Continue the session...
        '-----------------------------

LoginDone:
        If IDEDebug = True Then
            Debug.Print strRawData
        End If
            
    End Select


End Sub

Private Sub getvars(Data As String)

    unauthorized = InStr(1, Data, "Unauthorized")
    If unauthorized > 0 Then
        tmrTimeout.Enabled = False
        Layer = 0
        Winsock1.Close
        'Set SecureSession = Nothing
        Messenger.Close
        prg.Visible = False
        ResetVars
        MsgBox ("You have entered an incorrect username or password.")
    End If
    
'get other vars
For X = 1 To UBound(Split(Data, vbCrLf))
Dim strline As String
    strline = Split(Data, vbCrLf)(X - 1)
    If UBound(Split(strline, " ")) = 1 And strline <> "" Then
        If Split(strline, " ")(0) = "LoginTime:" Then logintime = Split(strline, " ")(1)
        If Split(strline, " ")(0) = "MSPAuth:" Then MSPAuth = Split(strline, " ")(1)
        If Split(strline, " ")(0) = "kv:" Then kv = Split(strline, " ")(1)
        If Split(strline, " ")(0) = "sid:" Then SID = Split(strline, " ")(1)
    End If
Next

End Sub

Private Sub Socket_Connect()
Dim strGoogle As String
    
    If InStr(strPacket, " ") Then
        strPacket = Replace(strPacket, " ", "%20")
    End If
    
    strGoogle = "GET " & strPacket & " HTTP/1.0" & vbNewLine
    strGoogle = strGoogle & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbNewLine
    strGoogle = strGoogle & "Accept-Language: en-us" & vbNewLine
    strGoogle = strGoogle & "Accept-Encoding: gzip, deflate" & vbNewLine
    strGoogle = strGoogle & "Host: www.google.com" & vbNewLine
    strGoogle = strGoogle & "Connection: Keep-Alive" & vbNewLine & vbNewLine
    
Socket.SendData strGoogle
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim strInfo As String
Socket.GetData strInfo, vbString
               
        Dim strData As Variant
        strData = Split(strInfo, "<p>")
        
        For g = 0 To UBound(strData)
            If InStr(strData(g), "Your search - <b>" & txtQuery & "</b> - did not match any documents.") Then
                DoneS = True
                Result = vbNewLine & "No Matches Found..."
            End If
        Next

        For X = 0 To 0
            If InStr(strData(I), "Results") Then
                results = "Results " & Trim(GetStringBetween(strData(X), "Results", "<hr>"))
                results = Replace(results, "<b>", "")
                results = Replace(results, "</b>", "")
            End If
        Next
        
        For I = 1 To UBound(strData) '- 2
        If InStr(strData(I), "<a href=http://") Then
            a = GetStringBetween(strData(I), "<a href=http://", ">")

            Dim strDec As String
            strDec = GetStringBetween(strData(I), " - ", " - ")
            strDec = CleanUp(strDec)
            
            Dim strTitle As String
            strTitle = GetStringBetween(strData(I), a & ">", "</a>")
            If InStr(strTitle, "<b>") Then
                strTitle = Replace(strTitle, "<b>", "")
                strTitle = Replace(strTitle, "</b>", "")
            End If
            Result = Result & vbNewLine & vbNewLine & strDec & vbNewLine & "http://" & a
        End If
    Next
End Sub

Private Sub tmrKeepAlive_Timer()
    msgsend ("PNG" & vbCrLf)
End Sub

Private Sub tmrTimeout_Timer()
    If prg.Value <= 6 Then
        Command1_Click
        tmrTimeout.Enabled = False
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    End If
End Sub

Public Sub Winsock1_Close()

' Handle SSL connection
'-----------------------------------------------

    Layer = 0
    Winsock1.Close
    Set SecureSession = Nothing

End Sub

Public Sub Winsock1_Connect()

' Handle SSL connection
'-----------------------------------------------
    If prg.Value = 3 Then prg.Value = 4
    
    Set SecureSession = New clsCrypto
    Call SendClientHello(Winsock1)
    
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

' Decode SSL Information
' Passes result to the ProcessData() sub
'-----------------------------------------------

    'Parse each SSL Record
    Dim TheData As String
    Dim ReachLen As Long

    Do
        If Winsock1.State = sckConnected Then
        
        If SeekLen = 0 Then
            If bytesTotal >= 2 Then
                Winsock1.GetData TheData, vbString, 2
                SeekLen = BytesToLen(TheData)
                bytesTotal = bytesTotal - 2
            Else
                Exit Sub
            End If
        End If
        
        If bytesTotal >= SeekLen Then
            Winsock1.GetData TheData, vbString, SeekLen
            bytesTotal = bytesTotal - SeekLen
        Else
            Exit Sub
        End If
        
        
        Select Case Layer
            Case 0:
                ENCODED_CERT = Mid(TheData, 12, BytesToLen(Mid(TheData, 6, 2)))
                CONNECTION_ID = Right(TheData, BytesToLen(Mid(TheData, 10, 2)))
                Call IncrementRecv
                Call SendMasterKey(Winsock1)
                prg.Value = 5
            Case 1:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If Right(TheData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
                    If VerifyMAC(TheData) Then Call SendClientFinish(Winsock1)
                Else
                    Winsock1.Close
                End If
             Case 2:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If VerifyMAC(TheData) = False Then Winsock1.Close
                Layer = 3
                
             Case 3:
                TheData = SecureSession.RC4_Decrypt(TheData)
getvars (TheData)
                If VerifyMAC(TheData) Then Call ProcessData(Mid(TheData, 17))
        End Select
    
        SeekLen = 0
        
        ElseIf Winsock1.State <> sckConnected Then
            Exit Sub
        End If
        
    Loop Until bytesTotal = 0

End Sub

Function DoSSL(strChallenge As String) As String

' Handles the SSL part of the authentication
'-----------------------------------------------

    Dim varLines As Variant
    Dim varURLS As Variant
    
    Dim intCurCookie As Integer
    
    Dim strAuthInfo As String
    Dim strHeader As String
    Dim strLoginServer As String
    Dim strLoginPage As String
    

    
    Dim colURLS As New Collection
    Dim colHeaders As New Collection


    
    'strChallenge = Replace(strChallenge, ",", "&")
    
'Connect to NEXUS:
'--------------------------------------------------
    strBuffer = ""
    
    Winsock1.Close
    Winsock1.Connect "nexus.passport.com", 443

    ' Wait for the SSL layer to be established:
    
    Do Until Layer = 3
        DoEvents
    Loop

    'Obtain login information from NEXUS:
    
    If Winsock1.State = sckConnected Then Call SSLSend(Winsock1, "GET /rdr/pprdr.asp HTTP/1.0" & vbCrLf & vbCrLf)
    
    Do Until InStr(1, strBuffer, vbCrLf & vbCrLf) <> 0
        DoEvents
    Loop
    
    Winsock1.Close
    
'--------------------------------------------------
'Done with NEXUS
    
    
    
'Begin processing data from NEXUS:
'--------------------------------------------------
    
    intCurCookie = 0
    varLines = Split(strBuffer, vbCrLf)
    
    ' Search for the header "PasswordURLs:"
    
        For intCount = LBound(varLines) To UBound(varLines)
        
            ' Add the values for "PasswordURLs:" to a collection:
            
            If Left$(CStr(varLines(intCount)), InStr(1, varLines(intCount), " ")) = "PassportURLs: " Then
                colHeaders.Add Right$(CStr(varLines(intCount)), Len(varLines(intCount)) - InStr(1, varLines(intCount), " ")), Left(varLines(intCount), InStr(1, varLines(intCount), " "))
                Exit For
            End If
            
        Next intCount
        
    
    varURLS = Split(colHeaders.Item("PassportURLs: "), ",")
    
    For intCount = LBound(varURLS) To UBound(varURLS)
        colURLS.Add Right(varURLS(intCount), Len(varURLS(intCount)) - InStr(1, varURLS(intCount), "=")), Left(varURLS(intCount), InStr(1, varURLS(intCount), "="))
    Next intCount
    
    'Get the server and page for logging in:

    strLoginServer = Left$(colURLS("DALogin="), InStr(1, colURLS("DALogin="), "/") - 1)
    strLoginPage = Right$(colURLS("DALogin="), Len(colURLS("DALogin=")) - InStr(1, colURLS("DALogin="), "/") + 1)
    
'--------------------------------------------------
'End processing
    

    
ConnectLogin:

'Connect to login server
'--------------------------------------------------

    strBuffer = ""
    
    ' Layer resembles the state of the SSL connection:
    Layer = 0
    
    Winsock1.Close
    Winsock1.Connect strLoginServer, 443

    ' Wait for the SSL layer to be established:
    
    Do Until Layer = 3
        
        DoEvents
    Loop

    strHeader = "GET " & strLoginPage & " HTTP/1.1" & vbCrLf & _
                "Authorization: Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & Replace(strUserName, "@", "%40") & ",pwd=" & URLEncode(strPassword) & "," & strChallenge & _
                "User-Agent: MSMSGS" & vbCrLf & _
                "Host: loginnet.passport.com" & vbCrLf & _
                "Connection: Keep-Alive" & vbCrLf & _
                "Cache-Control: no-cache" & vbCrLf & vbCrLf

    Call SSLSend(Winsock1, strHeader)

    ' Wait for the header to be recieved
    
    Do Until InStr(1, strBuffer, vbCrLf & vbCrLf) <> 0
        DoEvents
    Loop
    
    Dim strHeaderValue As String

    strHeaderValue = GetHeader("authentication-info:", strBuffer)
    
    If RequiresRedirect(strHeaderValue) = True Then
    
        strHeaderValue = GetHeader("location:", strBuffer)
        
        lngCharPos = InStr(strHeaderValue, "://")
        
        If (LCase$(Left$(strHeaderValue, lngCharPos - 1)) = "https") Then
        
            strLoginServer = Mid$(strHeaderValue, lngCharPos + 3, InStr(lngCharPos + 3, strHeaderValue, "/") - (lngCharPos + 3))
            strLoginPage = Right$(strHeaderValue, Len(strHeaderValue) - (InStr(lngCharPos + 3, strHeaderValue, "/") - 1))
            
            GoTo ConnectLogin
            
        End If
    
    Else
    
        DoSSL = ParseHash(strHeaderValue)
        Winsock1.Close
        prg.Value = 6
        Exit Function

    End If

'--------------------------------------------------
'Done with login server

End Function


Function GetHeader(strHeader As String, strData As String) As String

' Returns the value of a header-property
'-----------------------------------------------

Dim intCount As Integer
Dim varLines As Variant
Dim lngCharPos As Long
Dim strCurHeader As String

varLines = Split(strData, vbCrLf)

For intCount = LBound(varLines) To UBound(varLines)

If Len(varLines(intCount)) = 0 Then Exit For

    strCurHeader = varLines(intCount)
    lngCharPos = InStr(strCurHeader, " ")
    
    If LCase(Left(strCurHeader, lngCharPos - 1)) = LCase(strHeader) Then
        GetHeader = Right(strCurHeader, Len(strCurHeader) - lngCharPos)
        Exit Function
    End If
    

Next intCount

End Function

Function RequiresRedirect(strData As String) As Boolean

' Checks whether it's necessary to redirect to
' another server (using 'da-status' property)
'-----------------------------------------------

Dim intCount As Integer
Dim varProps As Variant
Dim lngCharPos As Long
Dim strCurItem As String
Dim strPropName As String
Dim strPropValue As String

lngCharPos = InStr(strData, " ")

If InStr(1, strData, "Passport1.4") Then
 
    strData = Right(strData, Len(strData) - lngCharPos)
    varProps = Split(strData, ",")
    
    For intCount = LBound(varProps) To UBound(varProps)
    
        strCurItem = varProps(intCount)
        lngCharPos = InStr(strCurItem, "=")
        
        strPropName = Left(strCurItem, lngCharPos - 1)
        strPropValue = Right(strCurItem, Len(strCurItem) - lngCharPos)
    
        If LCase$(strPropName) = "da-status" And LCase$(strPropValue) = "redir" Then
        
            RequiresRedirect = True
            Exit Function
            
        ElseIf LCase$(strPropName) = "da-status" And LCase$(strPropValue) = "success" Then
        
            RequiresRedirect = False
            Exit Function
        
        End If
        
    Next intCount

End If

End Function

Function ParseHash(strHeader As String) As String

' Returns the hash (from-pp) if the login has
' completed succesfully.
'-----------------------------------------------

Dim intCount As Integer
Dim varProps As Variant
Dim lngCharPos As Long
Dim strCurItem As String
Dim strPropName As String
Dim strPropValue As String

    varProps = Split(strHeader, ",")
    
    For intCount = LBound(varProps) To UBound(varProps)
    
        strCurItem = varProps(intCount)
        lngCharPos = InStr(strCurItem, "=")
        
        strPropName = Left(strCurItem, lngCharPos - 1)
        strPropValue = Right(strCurItem, Len(strCurItem) - lngCharPos)
    
        If LCase$(strPropName) = "from-pp" Then
        
            ParseHash = strPropValue
            'MsgBox ParseHash
            ParseHash = Left(ParseHash, Len(ParseHash) - 1)
            ParseHash = Right(ParseHash, Len(ParseHash) - 1)
            
            Exit Function
        
        End If
        
    Next intCount

End Function

Sub TellAll()
    Dim X As Integer
    For X = 1 To Capacity
        If frmMSG(X).InConvo = True And frmMSG(X).HearAnnounce = True Then
            frmMSG(X).SendMSG "(I)Announcement: " & Announce, 8
        End If
    Next
End Sub

Sub TellMod(Text)
    Dim X As Integer
    For X = 1 To Capacity
        If frmMSG(X).InConvo = True And frmMSG(X).HearAnnounce = True Then
            If IsUser(frmMSG(X).buddy) = True Or IsAdmin(frmMSG(X).buddy) = True Then
                frmMSG(X).SendMSG "(I)Bot Info: " & Text, 8
                DoEvents
            End If
        End If
        DoEvents
    Next
    AddToLog Text
End Sub

Function OpenConvos() As String
    Dim strData As String
    Dim X As Integer
    For X = 1 To Capacity
        If frmMSG(X).InConvo = True Then
            strData = strData & vbNewLine & frmMSG(X).buddy
        End If
    Next
    OpenConvos = strData
End Function

Sub TellUser(ByVal User As String, ByVal Message As String)
    Dim X As Integer
    For X = 1 To Capacity
        If UCase(frmMSG(X).buddy) = UCase(User) Then
            frmMSG(X).SendMSG Message
            frmMSG(X).txtHist.Text = frmMSG(X).txtHist.Text & "Crashing/Flooding..." & vbNewLine
        End If
    Next
End Sub

Sub TellSck(ByVal Sck As Integer, ByVal Message As String)
    frmMSG(Sck).SendMSG Message
End Sub

Function GoogleSearch(Text)
    strPacket = "/palm?q=" & Text
    Result = ""
    DoneS = False
    With Socket
        .Close
        .Connect
    End With
    Dim BT As Long
    BT = Timer
    Do
        DoEvents
    Loop Until Timer - BT > 2 Or DoneS = True
    GoogleSearch = Result
End Function

Function CurSocket(Name)
    For X = 1 To Capacity
        If frmMSG(X).buddy = Name Then
            Sock = Sock & vbNewLine & X & ": " & Name
        End If
    Next
    CurSocket = Sock
End Function

Function AllSocket()
    For X = 1 To Capacity
        If frmMSG(X).buddy <> "" Then
            Sock = Sock & vbNewLine & X & ": " & PrefGetName(frmMSG(X).buddy)
        End If
    Next
    AllSocket = Sock
End Function

Function sckCount()
    For X = 1 To Capacity
        If frmMSG(X).buddy = "" Then
            Y = Y + 1
        End If
    Next
    sckCount = Y
End Function

Function KillSck(Index)
    Unload frmMSG(Index)
End Function

Sub AddPersonToChat(Name)
    For X = 1 To Capacity
        If frmMSG(X).ImChat = True Then
            frmMSG(X).AddMeToChat Name
            Exit Sub
        End If
    Next
End Sub

Sub ExitChats()
    ChatOn = False
    For X = 1 To Capacity
        If frmMSG(X).ImChat = True Then
            frmMSG(X).ImChat = False
            Unload frmMSG(X)
            Exit Sub
        End If
    Next
End Sub

Attribute VB_Name = "modOS"
    'used for platform id
    Const VER_PLATFORM_WIN32s = 0 'win 3.x
    Const VER_PLATFORM_WIN32_WINDOWS = 1 'win 9.x
    Const VER_PLATFORM_WIN32_NT = 2 'win nt,2000,XP
    'used for product type
    Const VER_NT_WORKSTATION = 1
    Const VER_NT_SERVER = 3
    'used for suite mask
    Const VER_SUITE_DATACENTER = 128
    Const VER_SUITE_ENTERPRISE = 2
    Const VER_SUITE_PERSONAL = 512

Declare Function mciSendString Lib "WINMM.DLL" Alias "mciSendStringA" ( _
ByVal lpstrCommand As String, _
ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByValbInvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
    
Function Mci(sCommand As String) As String
    'Sends A Command To The Windows Media Dll
    Dim s As String * 255 'Create A String With A Pre-Defined Buffer Of 255 Char
    Call mciSendString(sCommand, s, 255, frmMain.hwnd) 'Call The Mci Send String Api Call
    Mci = Replace(s, Chr(0), "") 'Return Only What Is Needed
End Function

Public Function GetOSVer() As String
    Dim osv As OSVERSIONINFOEX
    osv.dwOSVersionInfoSize = Len(osv)

    If GetVersionEx(osv) = 1 Then

        Select Case osv.dwPlatformId
            Case Is = VER_PLATFORM_WIN32s
            GetOSVer = "Windows 3.x"
            
            Case Is = VER_PLATFORM_WIN32_WINDOWS
            Select Case osv.dwMinorVersion
            
                Case Is = 0
                If InStr(UCase(osv.szCSDVersion), "C") Then
                    GetOSVer = "Windows 95 OSR2"
                Else
                    GetOSVer = "Windows 95"
                End If
                
                Case Is = 10
                If InStr(UCase(osv.szCSDVersion), "A") Then
                    GetOSVer = "Windows 98 SE"
                Else
                    GetOSVer = "Windows 98"
                End If
                
                Case Is = 90
                GetOSVer = "Windows Me"
                
            End Select
            
        Case Is = VER_PLATFORM_WIN32_NT
        Select Case osv.dwMajorVersion
            Case Is = 3

            Select Case osv.dwMinorVersion
                Case Is = 0
                GetOSVer = "Windows NT 3"
                Case Is = 1
                GetOSVer = "Windows NT 3.1"
                Case Is = 5
                GetOSVer = "Windows NT 3.5"
                Case Is = 51
                GetOSVer = "Windows NT 3.51"
            End Select
        Case Is = 4
        GetOSVer = "Windows NT 4"
        Case Is = 5
        Select Case osv.dwMinorVersion
            Case Is = 0 'win 2000
            Select Case osv.wProductType
                Case Is = VER_NT_WORKSTATION
                GetOSVer = "Windows 2000 Professional"
                Case Is = VER_NT_SERVER
                Select Case osv.wSuiteMask
                    Case Is = VER_SUITE_DATACENTER
                    GetOSVer = "Windows 2000 DataCenter Server"
                    Case Is = VER_SUITE_ENTERPRISE
                    GetOSVer = "Windows 2000 Advanced Server"
                    Case Else
                    GetOSVer = "Windows 2000 Server"
                End Select
        End Select
    Case Is = 1 'win XP or win .NET server


    Select Case osv.wProductType
        Case Is = VER_NT_WORKSTATION 'win XP

        If osv.wSuiteMask = VER_SUITE_PERSONAL Then
            GetOSVer = "Windows XP Home Edition"
        Else
            GetOSVer = "Windows XP Professional"
        End If
        Case Else

        If osv.wSuiteMask = VER_SUITE_ENTERPRISE Then
            GetOSVer = "Windows .NET Enterprise Server"
        Else
            GetOSVer = "Windows .NET Server"
        End If
    End Select
End Select
End Select
End Select
End If
End Function








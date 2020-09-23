Attribute VB_Name = "modCnTime"
Option Explicit
Private Declare Function GetTimeZoneInformation Lib "kernel32.dll" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetLocalTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
  Bias As Long
  StandardName(0 To 31) As Integer
  StandardDate As SYSTEMTIME
  StandardBias As Long
  DaylightName(0 To 31) As Integer
  DaylightDate As SYSTEMTIME
  DaylightBias As Long
End Type
Private Const cSecsInStdYear As Long = 31536000
Private Const cSecInDay As Long = 86400
Public Function DateToTimestamp(ByVal dDateTime As Date) As Double
    Dim oTime As SYSTEMTIME
    With oTime
        .wYear = Year(dDateTime)
        .wMonth = Month(dDateTime)
        .wDay = Day(dDateTime)
        .wHour = Hour(dDateTime)
        .wMinute = Minute(dDateTime)
        .wSecond = Second(dDateTime)
    End With
    DateToTimestamp = BuildTimestamp(oTime)
End Function
Public Function DateSerialToTimestamp(ByVal iYear As Integer, ByVal iMonth As Integer, ByVal iDay As Integer, ByVal iHour As Integer, ByVal iMinute As Integer, ByVal iSecond As Integer, ByVal iMilliseconds As Integer) As Double
    Dim oTime As SYSTEMTIME
    With oTime
        .wYear = iYear
        .wMonth = iMonth
        .wDay = iDay
        .wHour = iHour
        .wMinute = iMinute
        .wSecond = iSecond
        .wMilliseconds = iMilliseconds
    End With
    DateSerialToTimestamp = BuildTimestamp(oTime)
End Function
Public Function DaysInMonth(ByVal iMonth As Integer, ByVal iYear As Integer) As Integer
    Select Case iMonth
        Case 1, 3, 5, 7, 8, 10, 12
            DaysInMonth = 31
        Case 2
            If (isLeapYear(iYear)) Then
                DaysInMonth = 29
            Else
                DaysInMonth = 28
            End If
        Case Is < 1, Is > 12
            DaysInMonth = 0
        Case Else
            DaysInMonth = 30
    End Select
End Function
Public Function GetCurrentTimeZoneOffset() As Single
    Dim oLocalTime As SYSTEMTIME
    Dim oGMT As SYSTEMTIME
    Dim iMinuteDiff As Integer
    Dim iHourDiff As Integer
    Call GetLocalTime(oLocalTime)
    Call GetSystemTime(oGMT)
    If (oLocalTime.wDay = oGMT.wDay) Then
        iHourDiff = oLocalTime.wHour - oGMT.wHour
    ElseIf (oLocalTime.wDay < oGMT.wDay) Then
        iHourDiff = (oLocalTime.wHour - 24) - oGMT.wHour
    Else
        iHourDiff = oLocalTime.wHour - (oGMT.wHour - 24)
    End If
    iMinuteDiff = Gap(CDbl(oGMT.wMinute), CDbl(oLocalTime.wMinute), 60)
    If (iMinuteDiff > 0 And oLocalTime.wMinute < iMinuteDiff) Then
        iHourDiff = iHourDiff - 1
    End If
    GetCurrentTimeZoneOffset = iHourDiff + (iMinuteDiff / 60)
End Function
Public Function GetTimeZoneOffset() As Single
    Dim oTimeZone As TIME_ZONE_INFORMATION
    Call GetTimeZoneInformation(oTimeZone)
    GetTimeZoneOffset = (-1 * (oTimeZone.Bias / 60))
End Function
Public Function isDaylightSavings() As Boolean
    isDaylightSavings = (GetTimeZoneOffset() <> GetCurrentTimeZoneOffset())
End Function
Public Function isLeapYear(ByVal iYear As Integer) As Boolean
    If ((iYear Mod 4) = 0) Then
        If ((iYear Mod 100) = 0) Then
            If ((iYear Mod 400) = 0) Then
                isLeapYear = True
            Else
                isLeapYear = False
            End If
        Else
            isLeapYear = True
        End If
    Else
        isLeapYear = False
    End If
End Function
Public Function Timestamp() As Double
    Dim oTime As SYSTEMTIME
    Call GetLocalTime(oTime)
    Timestamp = BuildTimestamp(oTime)
End Function
Public Function TimestampToDate(ByVal dTimestamp As Double) As Date
    Dim oTime As SYSTEMTIME
    dTimestamp = dTimestamp + (GetCurrentTimeZoneOffset() * (cSecInDay / 24))
    With oTime
        .wYear = 1970
        If (dTimestamp >= 0) Then
            Do While (dTimestamp >= cSecsInStdYear)
                .wYear = .wYear + 1
                If (isLeapYear(.wYear)) Then
                    If (dTimestamp >= (cSecsInStdYear + cSecInDay)) Then
                        dTimestamp = dTimestamp - (cSecsInStdYear + cSecInDay)
                    Else
                        .wYear = .wYear - 1
                        Exit Do
                    End If
                Else
                    dTimestamp = dTimestamp - cSecsInStdYear
                End If
            Loop
            If (isLeapYear(.wYear)) Then .wDayOfWeek = 1 Else .wDayOfWeek = 0
        Else
            dTimestamp = (dTimestamp * -1)
            Do While (dTimestamp >= cSecsInStdYear)
                .wYear = .wYear - 1
                If (isLeapYear(.wYear)) Then
                    If (dTimestamp >= (cSecsInStdYear + cSecInDay)) Then
                        dTimestamp = dTimestamp - (cSecsInStdYear + cSecInDay)
                    Else
                        .wYear = .wYear + 1
                        Exit Do
                    End If
                Else
                    dTimestamp = dTimestamp - cSecsInStdYear
                End If
            Loop
            If (isLeapYear(.wYear)) Then
                dTimestamp = ((dTimestamp - (cSecsInStdYear + cSecInDay)) * -1)
                .wDayOfWeek = 1
            Else
                dTimestamp = ((dTimestamp - cSecsInStdYear) * -1)
                .wDayOfWeek = 0
            End If
        End If
        Select Case Fix(dTimestamp / cSecInDay)
            Case Is >= (334 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((334 + .wDayOfWeek) * cSecInDay)
                .wMonth = 12
            Case Is >= (304 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((304 + .wDayOfWeek) * cSecInDay)
                .wMonth = 11
            Case Is >= (273 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((273 + .wDayOfWeek) * cSecInDay)
                .wMonth = 10
            Case Is >= (242 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((242 + .wDayOfWeek) * cSecInDay)
                .wMonth = 9
            Case Is >= (212 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((212 + .wDayOfWeek) * cSecInDay)
                .wMonth = 8
            Case Is >= (181 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((181 + .wDayOfWeek) * cSecInDay)
                .wMonth = 7
            Case Is >= (151 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((151 + .wDayOfWeek) * cSecInDay)
                .wMonth = 6
            Case Is >= (120 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((120 + .wDayOfWeek) * cSecInDay)
                .wMonth = 5
            Case Is >= (90 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((90 + .wDayOfWeek) * cSecInDay)
                .wMonth = 4
            Case Is >= (59 + .wDayOfWeek)
                dTimestamp = dTimestamp - ((59 + .wDayOfWeek) * cSecInDay)
                .wMonth = 3
            Case Is >= 31
                dTimestamp = dTimestamp - (31 * cSecInDay)
                .wMonth = 2
            Case Else
                .wMonth = 1
        End Select
        dTimestamp = Fix(dTimestamp)
        .wDay = Fix(dTimestamp / cSecInDay) + 1
        dTimestamp = (dTimestamp Mod cSecInDay)
        .wHour = Fix(dTimestamp / 3600)
        dTimestamp = (dTimestamp Mod 3600)
        .wMinute = Fix(dTimestamp / 60)
        dTimestamp = (dTimestamp Mod 60)
        .wSecond = Fix(dTimestamp)
        dTimestamp = dTimestamp - .wSecond
        TimestampToDate = CDate(DateSerial(.wYear, .wMonth, .wDay) & " " & Format(TimeSerial(.wHour, .wMinute, .wSecond), "hh:mm:ss")) ' & "." & .wMilliseconds)
    End With
End Function
Private Function BuildTimestamp(ByRef oTime As SYSTEMTIME) As Double
    Dim i As Long
    With oTime
        If (isLeapYear(.wYear)) Then .wDayOfWeek = 1 Else .wDayOfWeek = 0
        Select Case .wMonth
            Case 1
                BuildTimestamp = 0
            Case 2
                BuildTimestamp = (cSecInDay * 31)
            Case 3
                BuildTimestamp = (cSecInDay * (59 + .wDayOfWeek))
            Case 4
                BuildTimestamp = (cSecInDay * (90 + .wDayOfWeek))
            Case 5
                BuildTimestamp = (cSecInDay * (120 + .wDayOfWeek))
            Case 6
                BuildTimestamp = (cSecInDay * (151 + .wDayOfWeek))
            Case 7
                BuildTimestamp = (cSecInDay * (181 + .wDayOfWeek))
            Case 8
                BuildTimestamp = (cSecInDay * (212 + .wDayOfWeek))
            Case 9
                BuildTimestamp = (cSecInDay * (243 + .wDayOfWeek))
            Case 10
                BuildTimestamp = (cSecInDay * (273 + .wDayOfWeek))
            Case 11
                BuildTimestamp = (cSecInDay * (304 + .wDayOfWeek))
            Case 12
                BuildTimestamp = (cSecInDay * (334 + .wDayOfWeek))
        End Select
        BuildTimestamp = BuildTimestamp + (cSecInDay * (.wDay - 1))
        BuildTimestamp = BuildTimestamp + ((cSecInDay / 24) * .wHour)
        BuildTimestamp = BuildTimestamp + (60 * .wMinute)
        BuildTimestamp = BuildTimestamp + .wSecond
        BuildTimestamp = BuildTimestamp + (.wMilliseconds / (10 ^ Len(CStr(.wMilliseconds))))
        If (.wYear >= 1970) Then
            For i = 1970 To (.wYear - 1)
                If (isLeapYear(i)) Then
                    BuildTimestamp = BuildTimestamp + (cSecInDay * 366)
                Else
                    BuildTimestamp = BuildTimestamp + (cSecInDay * 365)
                End If
            Next
        Else
            BuildTimestamp = BuildTimestamp - (cSecInDay * 365)
            For i = .wYear To 1969
                If (isLeapYear(i)) Then
                    BuildTimestamp = BuildTimestamp - (cSecInDay * 366)
                Else
                    BuildTimestamp = BuildTimestamp - (cSecInDay * 365)
                End If
            Next
        End If
        BuildTimestamp = BuildTimestamp - (GetCurrentTimeZoneOffset() * (cSecInDay / 24))
    End With
End Function
Private Function Gap(ByRef dNumber1 As Double, ByRef dNumber2 As Double, ByRef lSpan As Long) As Double
    If (dNumber1 < dNumber2) Then
        Gap = (dNumber2 - dNumber1) Mod lSpan
    Else
        Gap = ((lSpan - dNumber1) + dNumber2) Mod lSpan
    End If
End Function

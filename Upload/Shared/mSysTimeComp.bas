Attribute VB_Name = "mSysTimeComp"
Option Explicit

Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type


Public Function SysTimeComp(SysTime1 As SYSTEMTIME, SysTime2 As SYSTEMTIME) As Long
    With SysTime2
        Select Case True
            
            Case .wYear > SysTime1.wYear
                SysTimeComp = 1
            Case .wYear < SysTime1.wYear
                SysTimeComp = -1
            
            Case .wMonth > SysTime1.wMonth
                SysTimeComp = 1
            Case .wMonth > SysTime1.wMonth
                SysTimeComp = -1
            
            Case .wDay > SysTime1.wDay
                SysTimeComp = 1
            Case .wDay < SysTime1.wDay
                SysTimeComp = -1
                
            Case .wHour > SysTime1.wHour
                SysTimeComp = 1
            Case .wHour < SysTime1.wHour
                SysTimeComp = -1
                
            Case .wMinute > SysTime1.wMinute
                SysTimeComp = 1
            Case .wMinute < SysTime1.wMinute
                SysTimeComp = -1
                
            Case .wSecond > SysTime1.wSecond
                SysTimeComp = 1
            Case .wSecond < SysTime1.wSecond
                SysTimeComp = -1
                
            Case .wMilliseconds > SysTime1.wMilliseconds
                SysTimeComp = 1
            Case .wMilliseconds < SysTime1.wMilliseconds
                SysTimeComp = -1
                
        End Select
    End With
End Function

Public Function SysTimeDiff(SysTime1 As SYSTEMTIME, SysTime2 As SYSTEMTIME) As Long
    Dim Date1 As Date, Date2 As Date
    With SysTime1
        Date1 = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
    With SysTime2
        Date2 = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
        SysTimeDiff = DateDiff("s", Date1, Date2) * 1000 + .wMilliseconds - SysTime1.wMilliseconds
    End With
End Function

Public Sub SysTimeAdd(SysTime As SYSTEMTIME, ByVal Millisecs As Long)
    Dim Date1 As Date
    Dim liTemp As Long
    Dim liSign As Long
    With SysTime
        liSign = Sgn(Millisecs)
        Select Case liSign
            Case 1
                liTemp = 1000 - .wMilliseconds
            Case -1
                liTemp = -.wMilliseconds
        End Select
        
        If Abs(Millisecs) < Abs(liTemp) Then
            .wMilliseconds = .wMilliseconds + Millisecs
        Else
            Millisecs = Millisecs - liTemp
            liTemp = Millisecs \ 1000
            Millisecs = Millisecs - liTemp * 1000
            Date1 = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
            If liSign = 1 Then liTemp = liTemp + 1
            Date1 = DateAdd("s", liTemp, Date1)
            
            
            .wYear = Year(Date1)
            .wMonth = Month(Date1)
            .wDay = Day(Date1)
            .wHour = Hour(Date1)
            .wMinute = Minute(Date1)
            .wSecond = Second(Date1)
            .wMilliseconds = Millisecs
        End If
            
    End With
End Sub


Attribute VB_Name = "mRounding"
Option Explicit

Public Enum RoundType
    RoundUp = 1
    RoundDown = 2
    RoundNearest = 3
End Enum


Public Function RoundNum(ByVal Number, Optional ByVal PlacesAfterDecimal As Long, Optional ByVal RoundTo As RoundType = RoundNearest)
    Dim ldblIncrement As Double
    ldblIncrement = (10 ^ (-1 - PlacesAfterDecimal)) * 5
    Select Case RoundTo
        Case RoundUp
            Number = Number + ldblIncrement
        Case RoundDown
            Number = Number - ldblIncrement
    End Select
    
    If PlacesAfterDecimal < 0 Then
        Dim liTempNum As Long
        liTempNum = 10 ^ Abs(PlacesAfterDecimal)
        RoundNum = MyRound(Number \ liTempNum) * liTempNum
    Else
        RoundNum = MyRound(Number, PlacesAfterDecimal)
    End If
End Function

Function RoundNumInterval(ByVal Number, ByVal Interval, ByVal RoundType As RoundType) As Double
    If Interval <= 0 Then Err.Raise 5
    Select Case RoundType
        Case RoundDown
            RoundNumInterval = Interval * CLng(Number / Interval - 0.5)
        Case RoundUp
            RoundNumInterval = Interval * CLng(Number / Interval + 0.5)
        Case RoundNearest
            RoundNumInterval = Interval * CLng(Number / Interval - 0.5)
            If Number - RoundNumInterval > (Interval / 2) Then RoundNumInterval = RoundNumInterval + Interval
    End Select
End Function

Public Function MyRound(ByVal Number, Optional NumDigitsAfterDecimal As Long)
    If Number Like "*." & String(NumDigitsAfterDecimal, "#") & "5" Then Number = Number - 0.000000000000001
    MyRound = Round(Number, NumDigitsAfterDecimal)
End Function


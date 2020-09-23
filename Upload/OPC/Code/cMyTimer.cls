VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMyTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements WinSubHook.iTimer

Private moTimer As cTimer
Private miInterval As Long

Public Item As Variant
Public Event ThatTime()

Private Sub Class_Terminate()
    On Error Resume Next
    Set moTimer = Nothing
End Sub

Private Sub iTimer_Fire(ByVal lElapsedMS As Long)
    RaiseEvent ThatTime
End Sub

Public Property Let Interval(piVal As Long)
    If piVal < 1 Then
        Set moTimer = Nothing
        miInterval = 0
    Else
        Set moTimer = New cTimer
        miInterval = piVal
        moTimer.Start piVal, Me
    End If
End Property

Public Property Get Interval() As Long
    Interval = miInterval
End Property

Attribute VB_Name = "mCallBack"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private moColl As Collection
Private miTimer As Long


Public Sub CallMeBack(ByVal WhoMe As iCallback)
    If moColl Is Nothing Then Set moColl = New Collection
    moColl.Add WhoMe
    If miTimer = 0 Then miTimer = SetTimer(0, 0, 1, AddressOf TimerProc)
End Sub

Private Sub TimerProc(ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    On Error Resume Next
    
    KillTimer 0, miTimer
    miTimer = 0
    
    Dim i As Long
    Dim iCB As iCallback
    
    For i = 1 To moColl.Count
        Set iCB = moColl(1)
        iCB.Callback
        moColl.Remove 1
    Next
End Sub


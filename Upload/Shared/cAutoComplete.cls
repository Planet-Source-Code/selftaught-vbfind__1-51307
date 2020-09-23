VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eAutoCompleteBehavior
    acbListOnly
    acbFile
    acbFolder
    acbAllowMultiSelect = 4
End Enum

Private miLastKeyDown As Long
Private mbInChangeEvent As Boolean
Private miACB As eAutoCompleteBehavior

Private WithEvents moCmb As ComboBox
Attribute moCmb.VB_VarHelpID = -1
Private WithEvents moTxt As TextBox
Attribute moTxt.VB_VarHelpID = -1

Event WillAutoComplete(pbCancel As Boolean)
Event TextChanged()

Public Property Let Behavior(piVal As eAutoCompleteBehavior)
    miACB = piVal
End Property
Public Property Get Behavior() As eAutoCompleteBehavior
    Behavior = miACB
End Property

Private Sub Class_Terminate()
    Set moCmb = Nothing
End Sub

Private Sub moCmb_Change()
    Dim lbVal As Boolean
    Dim lsText As String
    Dim liTemp As Long
    
    If mbInChangeEvent Then Exit Sub
    
    RaiseEvent WillAutoComplete(lbVal)
    
    If lbVal Then
        RaiseEvent TextChanged
        Exit Sub
    End If
    On Error Resume Next
    mbInChangeEvent = True
    lbVal = BitIsSet(miACB, acbAllowMultiSelect)
    lsText = moCmb.Text
    If lbVal Then
        liTemp = InStrRev(lsText, ";")
        If liTemp > 0 Then lsText = Mid$(lsText, liTemp + 1)
    End If
    Select Case miACB Mod acbAllowMultiSelect
        Case acbFile
            lsText = FindFirst(moCmb.Text, False)
        Case acbFolder
            lsText = FindFirst(moCmb.Text, True)
    End Select
    If lbVal And liTemp > 0 And Len(lsText) > 0 Then lsText = Mid$(moCmb.Text, liTemp) & lsText
    AutoComplete moCmb, miLastKeyDown, lsText
    mbInChangeEvent = False
    RaiseEvent TextChanged
End Sub

Private Sub moCmb_Click()
    RaiseEvent TextChanged
End Sub

Private Sub moCmb_KeyDown(KeyCode As Integer, Shift As Integer)
    miLastKeyDown = KeyCode
End Sub

Public Sub Attach(ToMe As Object)
    On Error Resume Next
    Set moCmb = Nothing
    Set moTxt = Nothing
    Set moCmb = ToMe
    Set moTxt = ToMe
End Sub

Private Sub moCmb_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim liTemp As Long
    With moCmb
        If KeyAscii = 92 And miACB > acbListOnly Then
            liTemp = InStr(.SelStart, .Text, "\")
            If liTemp > 0 Then
                .SelStart = liTemp
                .SelLength = Len(.Text) - liTemp
                KeyAscii = 0
            End If
        ElseIf KeyAscii = 59 And (BitIsSet(miACB, acbAllowMultiSelect)) Then
            liTemp = InStr(.SelStart + 1, .Text, Chr(KeyAscii))
            If liTemp > 0 Then
                .SelStart = liTemp
                .SelLength = 1
            Else
                .SelStart = Len(.Text)
            End If
        End If
    End With
End Sub

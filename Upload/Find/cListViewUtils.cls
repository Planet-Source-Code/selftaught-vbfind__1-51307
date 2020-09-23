VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLVUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements WinSubHook.iTimer

Private Const VK_LBUTTON As Long = &H1
Private Const VK_RBUTTON As Long = &H2
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

Private mTimer As cTimer
Attribute mTimer.VB_VarHelpID = -1

Private WithEvents moList As ListView
Attribute moList.VB_VarHelpID = -1
Private moColumns As cLVColumns

Private miHwnd As Long
Private moItemClicked As ListItem

Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11

Private Const DragCheckInterval = 100

Private mbDontRaiseForSelectedItem As Boolean

Event ItemSelected(ByVal Item As ListItem)
Event Drag(ByVal Buttons As Long, ByVal Shift As Long)
Event ItemActivated()


'#########################
'###    Event Procs    ###
'#########################

Private Sub Class_Initialize()
    Set moColumns = New cLVColumns
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Detach
    Set moColumns = Nothing
    Set moList = Nothing
    Set mTimer = Nothing
End Sub

Private Sub iTimer_Fire(ByVal lElapsedMS As Long)
    On Error Resume Next
    Dim liButtons As Long, liShift As Long
    
    If VKIsDown(VK_RBUTTON) Then liButtons = vbRightButton
    If VKIsDown(VK_LBUTTON) Then liButtons = liButtons Or vbLeftButton
    
    
    
    If liButtons <> 0 Then
        If VKIsDown(VK_SHIFT) Then liShift = vbShiftMask
        If VKIsDown(VK_CONTROL) Then liShift = liShift = vbCtrlMask
        RaiseEvent Drag(liButtons, liShift)
    Else
        Set mTimer = Nothing
    End If
    
End Sub

Private Sub moList_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    SortListView moList, ColumnHeader
End Sub

Private Sub moList_DblClick()
    RaiseEvent ItemActivated
End Sub

Private Sub moList_ItemClick(ByVal Item As ListItem)
    On Error Resume Next
    If Item Is moList.SelectedItem Then
        If mbDontRaiseForSelectedItem Then
            mbDontRaiseForSelectedItem = False
            Exit Sub
        End If
    Else
        If Not Item Is moItemClicked Then Exit Sub
    End If
    RaiseEvent ItemSelected(Item)
End Sub

Private Sub moList_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If BitIsSet(Shift, vbShiftMask) Then
        Select Case KeyCode
            Case vbKeyHome
                KeyCode = 0
                SelectItems True
            Case vbKeyEnd
                KeyCode = 0
                SelectItems False
            Case vbKeyUp
                With moList
                    Set moItemClicked = .ListItems(.SelectedItem.Index - 1)
                    mbDontRaiseForSelectedItem = True
                End With
            Case vbKeyDown
                With moList
                    Set moItemClicked = .ListItems(.SelectedItem.Index + 1)
                    mbDontRaiseForSelectedItem = True
                End With
        End Select
    End If
End Sub

Public Sub SelectItems(pbInvert As Boolean)
    Dim liIndex As Long
    Dim loListItem As ListItem
    liIndex = moList.SelectedItem.Index

    If pbInvert Then
        
        For Each loListItem In moList.ListItems
            With loListItem
                If .Index < liIndex Then .Selected = True
            End With
        Next
        Set loListItem = moList.ListItems(1)
    Else
    
        For Each loListItem In moList.ListItems
            With loListItem
                If .Index > liIndex Then .Selected = True
            End With
        Next
        Set loListItem = moList.ListItems(moList.ListItems.Count)
    End If
    Set moList.SelectedItem = loListItem
    loListItem.EnsureVisible
    RaiseEvent ItemSelected(loListItem)
End Sub

Private Sub moList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then RaiseEvent ItemActivated
End Sub

Private Sub moList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Set moItemClicked = moList.HitTest(X, Y)
    Set mTimer = New cTimer
    mTimer.Start DragCheckInterval, Me
End Sub

'#########################
'###    Public Subs    ###
'#########################

Public Sub Attach(ByVal poLV As ListView)
    On Error Resume Next
    miHwnd = poLV.hWnd
    Set moList = poLV
    Set moColumns.ColHeaders = poLV.ColumnHeaders
End Sub

Public Sub Detach()
    On Error Resume Next
    Set moColumns = Nothing
    Set moList = Nothing
End Sub

'#########################
'### Public Properties ###
'#########################

Public Property Get ColumnHeaders() As cLVColumns
    Set ColumnHeaders = moColumns
End Property











Private Function VKIsDown(piVK As Long)
    Dim liVal As Long
    liVal = GetKeyState(piVK)
    VKIsDown = BitIsSet(liVal, 128)
End Function


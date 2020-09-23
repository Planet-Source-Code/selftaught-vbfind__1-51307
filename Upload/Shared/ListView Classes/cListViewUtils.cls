VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListViewUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Drag?????
'Private Declare Function GetTickCount Lib "kernel32" () As Long

Implements WinSubHook.iTimer
Implements WinSubHook.iSubclass

Private Const VK_LBUTTON As Long = &H1
Private Const VK_RBUTTON As Long = &H2
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

Private mTimer As cTimer
Attribute mTimer.VB_VarHelpID = -1
Private mSubClass As cSubclass

Private WithEvents moList As ListView
Attribute moList.VB_VarHelpID = -1
Private moColumns As cLVColumns
Private moListItemsMgr As cListItemsMgr

Private mbLButtonDown As Boolean
Private mbRButtonDown As Boolean

Private miHwnd As Long
Private moItemClicked As ListItem

Private Const WM_MOUSEWHEEL = &H20A

Private Const MK_SHIFT As Long = &H4
Private Const MK_CONTROL As Long = &H8

Private mbDontRaiseForSelectedItem As Boolean

Event ButtonDown(ByVal piButton As Long)
Event ButtonUp(ByVal piButton As Long)
Event ItemSelected(ByVal Item As ListItem)

'#########################
'###    Event Procs    ###
'#########################

Private Sub Class_Initialize()
    Set moColumns = New cLVColumns
    Set moListItemsMgr = New cListItemsMgr
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Detach
    Set moColumns = Nothing
    Set moListItemsMgr = Nothing
    Set moList = Nothing
    Set mTimer = Nothing
End Sub


Private Sub iSubclass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
'
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
    Select Case uMsg
        Case WM_MOUSEWHEEL
            SetBit wParam, MK_CONTROL, False
            SetBit wParam, MK_SHIFT, False
            If mbLButtonDown = False And mbRButtonDown = False Then CallOldWindowProc hwnd, iMsg, wParam, lParam
    End Select
    bHandled = True
End Sub

Private Sub iTimer_Fire(ByVal lElapsedMS As Long)
    On Error Resume Next
    Dim lbR As Boolean
    Dim lbL As Boolean
    
    lbR = VKIsDown(VK_RBUTTON)
    lbL = VKIsDown(VK_LBUTTON)
    
    If Not mbLButtonDown = lbL Then
        mbLButtonDown = lbL
        If mbLButtonDown Then RaiseEvent ButtonDown(1) Else RaiseEvent ButtonUp(1)
    End If
    
    If Not mbRButtonDown = lbR Then
        mbRButtonDown = lbR
        If mbRButtonDown Then RaiseEvent ButtonDown(2) Else RaiseEvent ButtonUp(2)
    End If
    If mbRButtonDown = False And mbLButtonDown = False Then Set mTimer = Nothing
    
End Sub

Private Sub moList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView moList, ColumnHeader
End Sub

Private Sub moList_ItemClick(ByVal Item As MSComctlLib.ListItem)
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

Private Sub moList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Set moItemClicked = moList.HitTest(x, y)
    If Button = 1 Then
        mbLButtonDown = True
    ElseIf Button = 2 Then
        mbRButtonDown = True
    Else
        Exit Sub
    End If
    RaiseEvent ButtonDown(Button)
    Set mTimer = New cTimer
    mTimer.Start 50, Me
End Sub

'#########################
'###    Public Subs    ###
'#########################

Public Sub Attach(ByVal poLV As ListView)
    On Error Resume Next
    miHwnd = poLV.hwnd
    If mSubClass Is Nothing Then Set mSubClass = New cSubclass Else mSubClass.UnSubclass

    AttachMessage Me, miHwnd, WM_MOUSEWHEEL
    Set moList = poLV
    Set moColumns.ColHeaders = poLV.ColumnHeaders
    moListItemsMgr.Attach poLV
End Sub

Public Sub Detach()
    On Error Resume Next
    DetachMessage Me, miHwnd, WM_MOUSEWHEEL
    Set moColumns = Nothing
    Set moListItemsMgr = Nothing
    Set moList = Nothing
End Sub

'#########################
'### Public Properties ###
'#########################

Public Property Get ListItemsMgr() As cListItemsMgr
    Set ListItemsMgr = moListItemsMgr
End Property

Public Property Get ColumnHeaders() As cLVColumns
    Set ColumnHeaders = moColumns
End Property


Private Function VKIsDown(piVK As Long)
    Dim liVal As Long
    liVal = GetKeyState(piVK)
    VKIsDown = BitIsSet(liVal, 128)
End Function


Attribute VB_Name = "mGeneral"
Option Explicit

Public Const cBackupRestoreClass = "AsyncFileTasks.cFileBackupRestore"
Public Const cCopyMoveDeleteClass = "AsyncFileTasks.cFileCopyMoveDelete"
Public Const cSearchClass = "AsyncFileTasks.cFileSearch"

Public Enum eControlTypes
    ctDateTime
    ctCombo
    ctOption
    ctUpDown
    ctComboText
    ctCheck
    ctText
End Enum

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Public Enum eShellTypes
    stDefault
    stOpen
    stPrint
    stExplore
End Enum


Const GW_CHILD = 5
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const SortKeySuffix = "SortKey"

Private Const GWL_WNDPROC = -4
Private miEditHwnd As Long
Private miOldEditProc As Long

Public Function shelldoc(psFile As String, Optional ByVal piType As eShellTypes)
    Const SE_ERR_NOASSOC = 31
    Dim lsType As String
    lsType = Choose(piType + 1, vbNullString, "open", "print", "explore")
    
    If ShellExecute(GetDesktopWindow(), lsType, psFile, _
                    vbNullString, vbNullString, SW_SHOWNORMAL) _
       = SE_ERR_NOASSOC Then
       
        ShellExecute GetDesktopWindow(), "open", "RUNDLL32.EXE", _
                     "shell32.dll,OpenAs_RunDLL " & psFile, PathGetSpecial(sfSystem), SW_SHOWNORMAL
    End If
End Function

Public Function shellexe(psFile As String, psParams As String)
    ShellExecute GetDesktopWindow, vbNullString, psFile, psParams, vbNullString, SW_SHOWNORMAL
End Function

Public Sub SortListView(ByVal poListView As ListView, ByVal poColHeader As ColumnHeader, Optional ByVal piOrder As ListSortOrderConstants = -1)
    Select Case poColHeader.Key
        Case "ReceiveDate", "CompleteDate", "InvoiceDate", "Since", "Amount", "Balance", "Modified", "Created", "Accessed", "Size"
            EnsureSortKey poListView, poColHeader.Key
            Set poColHeader = poListView.ColumnHeaders(poColHeader.Key & SortKeySuffix)
    End Select
    
    If poListView.SortKey = poColHeader.SubItemIndex And piOrder = -1 Then 'And poListView.Sorted = True
        If poListView.SortOrder = lvwAscending Then poListView.SortOrder = lvwDescending Else poListView.SortOrder = lvwAscending
        poListView.Sorted = True
    Else
        If piOrder = -1 Then piOrder = lvwAscending
        poListView.SortOrder = piOrder
        poListView.SortKey = poColHeader.SubItemIndex
        poListView.Sorted = True
    End If
    poListView.Sorted = False
    On Error Resume Next
    poListView.SelectedItem.EnsureVisible
End Sub


Public Sub EnsureSortKey(poListView As ListView, psColumnKey As String)
    Dim lsNewKey    As String
    Dim lsNewText   As String
    Dim liOldIndex  As Integer
    Dim liNewIndex  As Integer
    Dim liTemp As Long
    Dim loCurLI     As ListItem
    
    lsNewKey = psColumnKey & SortKeySuffix
    liOldIndex = poListView.ColumnHeaders(psColumnKey).SubItemIndex
    On Error Resume Next
    poListView.ColumnHeaders.Add , lsNewKey, , 0
    poListView.ColumnHeaders(lsNewKey).Tag = SortKeySuffix
    poListView.Sorted = False
    liNewIndex = poListView.ColumnHeaders(lsNewKey).SubItemIndex
    Select Case psColumnKey
        Case "ReceiveDate", "CompleteDate", "Since", "InvoiceDate", "Modified", "Accessed", "Created"
            For Each loCurLI In poListView.ListItems
                loCurLI.SubItems(liNewIndex) = CDbl(CDate(loCurLI.SubItems(liOldIndex)))
            Next
        Case Else
            For Each loCurLI In poListView.ListItems
                lsNewText = loCurLI.SubItems(liOldIndex)
                liTemp = InStr(1, lsNewText, " ")
                If liTemp > 0 Then lsNewText = Left$(lsNewText, liTemp - 1)
                loCurLI.SubItems(liNewIndex) = Format(CSng(lsNewText), "0000000.000000")
            Next
    End Select
End Sub

Public Sub SaveControls(poControls As Object, piType As eControlTypes, poFile As cFileIO)
    Dim i As Long
    With poControls
        Select Case piType
            Case ctOption
                poFile.AppendInteger OptIndex(poControls)
            Case ctDateTime
                For i = 0 To .UBound
                    poFile.AppendDouble .Item(i).Value
                Next
            Case ctCombo
                For i = 0 To .UBound
                    poFile.AppendInteger .Item(i).ListIndex
                Next
            Case ctComboText
                Dim j As Long
                For j = 0 To .UBound
                    With .Item(j)
                        i = .ListCount
                        poFile.AppendInteger i
                        For i = 0 To i - 1
                            poFile.AppendString .List(i)
                        Next
                    End With
                Next
            Case ctCheck
                For i = 0 To .UBound
                    poFile.AppendInteger .Item(i).Value
                Next
            Case ctUpDown
                For i = 0 To .UBound
                    poFile.AppendLong .Item(i).Value
                Next
            Case ctText
                For i = 0 To .UBound
                    poFile.AppendString .Item(i).Text
                Next
        End Select
    End With
End Sub

Public Sub LoadControls(poControls As Object, piType As eControlTypes, poFile As cFileIO)
    Dim ldDouble As Double, lsString As String, liLong As Long, liInt As Long, i As Long
    
    With poControls
        Select Case piType
            Case ctOption
                poFile.GetInteger liInt
                .Item(liInt).Value = True
            Case ctDateTime
                For i = 0 To .UBound
                    poFile.GetDouble ldDouble
                    .Item(i).Value = ldDouble
                Next
            Case ctCombo
                For i = 0 To .UBound
                    poFile.GetInteger liInt
                    .Item(i).ListIndex = liInt
                Next
            Case ctComboText
                Dim j As Long
                For j = 0 To .UBound
                    With .Item(j)
                        poFile.GetInteger liInt
                        For i = 1 To liInt
                            poFile.GetString lsString
                            .AddItem lsString
                        Next
                    End With
                Next
            Case ctCheck
                For i = 0 To .UBound
                    poFile.GetInteger liInt
                    .Item(i).Value = liInt
                Next
            Case ctUpDown
                For i = 0 To .UBound
                    poFile.GetLong liLong
                    .Item(i).Value = liLong
                Next
            Case ctText
                For i = 0 To .UBound
                    poFile.GetString lsString
                    .Item(i).Text = lsString
                Next
        End Select
    End With
End Sub

Public Sub AddExclusiveItem(ByVal poCombo As ComboBox)
    On Error Resume Next
    Dim lsText As String
    Dim liSelStart As Long
    Dim liSelLength As Long
    liSelStart = -1
    lsText = poCombo.Text
    Dim i As Long
    With poCombo
        Do While i < poCombo.ListCount
            If StrComp(poCombo.List(i), lsText, vbTextCompare) = 0 Then
                If liSelStart = -1 Then
                    liSelStart = .SelStart
                    liSelLength = .SelLength
                End If
                poCombo.RemoveItem i
            Else
                i = i + 1
            End If
        Loop
        If Not Len(lsText) = 0 Then .AddItem lsText, 0
        If liSelStart > -1 Then
            .Text = lsText
            .SelStart = liSelStart
            .SelLength = liSelLength
        End If
    End With
End Sub

Public Sub ShowProps(FileName As String)

    Dim SEI As SHELLEXECUTEINFO
    Dim r As Long
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
         SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = GetDesktopWindow
        .lpVerb = "properties"
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    r = ShellExecuteEx(SEI)
End Sub

Public Function GetFirstChild(ByVal piHwnd As Long) As Long
    GetFirstChild = GetWindow(piHwnd, GW_CHILD)
End Function

Public Sub TrashNextSetText(ByVal piTextHwnd As Long)
    UnSubclass
    miOldEditProc = SetWindowLong(piTextHwnd, GWL_WNDPROC, AddressOf EditProc)
    miEditHwnd = piTextHwnd
End Sub

Private Sub UnSubclass()
    If miOldEditProc <> 0 Then
        SetWindowLong miEditHwnd, GWL_WNDPROC, miOldEditProc
        miOldEditProc = 0
        miEditHwnd = 0
    End If
End Sub

Private Function EditProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const WM_SETTEXT = &HC
    If (iMsg = WM_SETTEXT) Then
        EditProc = True
        UnSubclass
    Else
        EditProc = CallWindowProc(miOldEditProc, miEditHwnd, iMsg, wParam, lParam)
    End If
End Function

Public Sub Main()
    frmFind.Show
End Sub

Public Function ShowTask(ByVal poFiles As cFiles, ByVal piTask As Long, psDest As String, psRelative As String) As Boolean
    Dim loUI As frmFileTask
    Set loUI = New frmFileTask
    ShowTask = loUI.StartTask(poFiles, piTask, psDest, psRelative)
    If Not ShowTask Then Unload loUI
    Set loUI = Nothing
End Function

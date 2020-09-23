VERSION 5.00
Begin VB.Form frmLVColumns 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Columns"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmLVColumns.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      Caption         =   "Move &Down"
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Move &Up"
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   2940
      Width           =   1095
   End
   Begin VB.ListBox lst 
      Height          =   2985
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Pixels Wide:"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   2700
      Width           =   1095
   End
End
Attribute VB_Name = "frmLVColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCanceled As Boolean
Private moKeys As Collection

Public Sub ChooseColumns(ByVal poCols As cLVColumns, Optional ByVal poOwner As Form)
    Dim loEach As cLVColumn
    Dim i As Long
    poCols.SyncWidths
    With lst
        .Clear
        Set moKeys = New Collection
        For Each loEach In poCols
            With loEach
                lst.AddItem .Text
                lst.ItemData(lst.NewIndex) = .Width \ Screen.TwipsPerPixelX
                lst.Selected(lst.NewIndex) = .Visible
                moKeys.Add .Key, .Text
            End With
        Next
        If .ListCount = 0 Then Exit Sub
    End With
    lst_Click
    If poOwner Is Nothing Then Show vbModal Else Show vbModal, poOwner
    If Not mbCanceled Then
        Dim lsTemp As String
        With poCols
            .Redraw = False
            .Clear
            For i = 0 To lst.ListCount - 1
                lsTemp = moKeys(lst.List(i))
                .Add lsTemp, lst.List(i), lst.ItemData(i) * Screen.TwipsPerPixelX, lst.Selected(i)
            Next
            .Redraw = True
        End With
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case vbOK
            mbCanceled = False
            Hide
        Case vbCancel
            mbCanceled = True
            Hide
        Case Else
            Dim liData As Long
            Dim lsString As String
            Dim liIndex As Long
            Dim lbSel As Boolean
            
            With lst
                liIndex = .ListIndex
                lsString = .List(liIndex)
                liData = .ItemData(liIndex)
                lbSel = .Selected(liIndex)
                
                If Index = 4 Then 'Move Down
                    If liIndex < .ListCount - 1 Then
                        .RemoveItem liIndex
                        .AddItem lsString, liIndex + 1
                        .ItemData(.NewIndex) = liData
                        .Selected(.NewIndex) = lbSel
                        .ListIndex = .NewIndex
                    End If
                Else 'Move Up
                    If liIndex > 0 Then
                        .RemoveItem liIndex
                        .AddItem lsString, liIndex - 1
                        .ItemData(.NewIndex) = liData
                        .Selected(.NewIndex) = lbSel
                        .ListIndex = .NewIndex
                    End If
                End If
            End With
    End Select
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveForm hwnd
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveForm hwnd
End Sub

Private Sub lst_Click()
    txt.Text = lst.ItemData(lst.ListIndex)
End Sub

Private Sub lst_ItemCheck(Item As Integer)
    On Error Resume Next
    If Not lst.Selected(0) Then lst.Selected(0) = True
End Sub

Private Sub txt_Change()
    With lst
        .ItemData(.ListIndex) = Val(txt.Text)
    End With
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack, vbKeyDelete, vbKey0 To vbKey9
        
        Case Else
            KeyAscii = 0
    End Select
End Sub

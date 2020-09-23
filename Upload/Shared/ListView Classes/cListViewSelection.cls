VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSelDataListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mKeys As Collection

Private moRS As ADODB.Recordset
Private msSelectedItem As String

Friend Sub Calc(poLV As ListView, pRS As ADODB.Recordset)
    Dim lvEachOne As ListItem
    On Error Resume Next
    moRS.Close
    Set moRS = pRS.Clone
    Set mKeys = New Collection
    If poLV.Checkboxes Then
        For Each lvEachOne In poLV.ListItems
            If lvEachOne.Checked Then mKeys.Add lvEachOne.Key
        Next
        If poLV.SelectedItem Is Nothing Then
            msSelectedItem = ""
        Else
            With poLV.SelectedItem
                 If .Checked Then msSelectedItem = .Key Else msSelectedItem = ""
            End With
        End If
    Else
        For Each lvEachOne In poLV.ListItems
            If lvEachOne.Selected Then mKeys.Add lvEachOne.Key
        Next
        If poLV.SelectedItem Is Nothing Then
            msSelectedItem = ""
        Else
            With poLV.SelectedItem
                 If .Selected Then msSelectedItem = .Key Else msSelectedItem = ""
            End With
        End If
    End If
End Sub

Public Sub GetSelectedItemField(ParamArray psFieldNames() As Variant)
    Dim i As Long
    FindRecord moRS, CLng(Val(msSelectedItem))
    For i = LBound(psFieldNames) To UBound(psFieldNames)
        psFieldNames(i) = moRS.Fields(psFieldNames(i)).Value
    Next
End Sub



Public Function Count() As Long
    Count = mKeys.Count
End Function

Public Function IsSelected(psKey As String) As Boolean
    On Error Resume Next
    Dim lsTemp
    lsTemp = mKeys.Item(psKey)
    IsSelected = Err.Number = 0
End Function

Public Function SelectedItem()
    SelectedItem = msSelectedItem
End Function


Private Sub Class_Initialize()
    Set mKeys = New Collection
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mKeys.[_NewEnum]
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    moRS.Close
    Set moRS = Nothing
    Set mKeys = Nothing
End Sub

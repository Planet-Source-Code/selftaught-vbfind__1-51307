VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moColl As Collection

Public Sub Add(ByRef Item, _
      Optional ByVal psKey As String, _
      Optional ByVal Before As Variant, _
      Optional ByVal After As Variant)
    
    If Not (IsMissing(After) Or IsMissing(Before)) Then Err.Raise 5
    
    On Error GoTo errhand
    
    If Not IsMissing(After) Then
        If Len(psKey) = 0 Then
            moColl.Add Item, , , After
        Else
            moColl.Add Item, psKey, , After
        End If
    ElseIf Not IsMissing(Before) Then
        If Len(psKey) = 0 Then
            moColl.Add Item, , Before
        Else
            moColl.Add Item, psKey, Before
        End If
    Else
errhand:
        On Error GoTo 0
        If Len(psKey) = 0 Then
            moColl.Add Item
        Else
            moColl.Add Item, psKey
        End If
    End If
End Sub

Public Sub Remove(ByVal pvIndex As Variant)
    moColl.Remove pvIndex
End Sub

Public Sub RemoveAll()
    Set moColl = New Collection
End Sub

Public Property Get Count() As Long
    Count = moColl.Count
End Property

Public Function Item(ByVal pvIndex As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If IsObject(moColl(pvIndex)) Then _
        Set Item = moColl(pvIndex) _
    Else Item = moColl(pvIndex)
End Function

Public Function IndexExists(ByVal piIndex As Long) As Boolean
    On Error Resume Next
    IndexExists = (piIndex > 0 And piIndex <= moColl.Count)
End Function

Public Function KeyExists(ByVal psKey As String) As Boolean
    On Error GoTo ending
    IsObject moColl(psKey)
    KeyExists = True
    Exit Function
ending:
    KeyExists = False
End Function



Private Sub Class_Initialize()
    Set moColl = New Collection
End Sub

Public Property Get GetNewEnum() As IUnknown
    Set GetNewEnum = moColl.[_NewEnum]
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = moColl.[_NewEnum]
End Property

Private Sub Class_Terminate()
    Set moColl = Nothing
End Sub

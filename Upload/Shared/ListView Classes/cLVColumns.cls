VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLVColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moColHeaders As ColumnHeaders
Private moColl As Collection

Private mbRedraw As Boolean

Private msTextKey As String

Public Property Get TextKey() As String
    TextKey = msTextKey
End Property

Public Property Let Redraw(pbVal As Boolean)
    mbRedraw = pbVal
    If mbRedraw Then ShowColumns
End Property
Public Property Get Redraw() As Boolean
    Redraw = mbRedraw
End Property

Public Property Let Visible(pvIndex, pbVal As Boolean)
    moColl(pvIndex).Visible = pbVal
    If mbRedraw Then ShowColumns
End Property
Public Property Get Visible(pvIndex) As Boolean
    Visible = moColl(pvIndex).Visible
End Property

Public Property Let Width(pvIndex, ByVal piVal As Long)
    moColl(pvIndex).Width = piVal
End Property
Public Property Get Width(pvIndex) As Long
    Width = moColl(pvIndex).Width
End Property

Public Property Get Text(pvIndex) As String
    Text = moColl(pvIndex).Text
End Property
Public Property Let Text(pvIndex, psString As String)
    moColl(pvIndex).Text = psString
End Property

Public Sub Remove(pvIndex)
    moColl.Remove pvIndex
    If mbRedraw Then ShowColumns
End Sub

Public Sub Add(psKey As String, psText As String, piWidth As Long, Optional pbVisible As Boolean = True, Optional Before As Variant, Optional After As Variant)
    Dim loCol As cLVColumn
    Set loCol = New cLVColumn
    With loCol
        .Visible = pbVisible
        .Key = psKey
        .Text = psText
        .Width = piWidth
    End With
    
    If IsMissing(Before) And IsMissing(After) Then
        moColl.Add loCol, psKey
    ElseIf IsMissing(After) Then
        moColl.Add loCol, psKey, Before
    Else
        moColl.Add loCol, psKey, , After
    End If
    If mbRedraw Then ShowColumns
    Set loCol = Nothing
End Sub

Private Sub Class_Initialize()
    Set moColl = New Collection
    mbRedraw = True
End Sub

Private Sub Class_Terminate()
    Set moColl = Nothing
    Set moColHeaders = Nothing
End Sub

Private Sub ShowColumns()
    On Error Resume Next
    Dim i As Long, j As Long
    Dim liIndex As Long
    Dim liColIndex As Long
    liColIndex = 1
    With moColHeaders
        For i = 1 To moColl.Count
            If moColl(i).Visible = True Then
                liIndex = 0
                liIndex = .Item(moColl(i).Key).Index
                If liIndex = 0 Then
                    If liColIndex > .Count Then
                        .Add , moColl(i).Key, moColl(i).Text
                    Else
                        .Add liColIndex, moColl(i).Key, moColl(i).Text
                    End If
                Else
                    If liIndex > liColIndex Then
                        For j = liColIndex To liIndex
                            If InStr(1, .Item(liColIndex).Key, SortKeySuffix) = 0 Then
                                .Remove liColIndex
                            Else
                                i = i + 1
                                liColIndex = liColIndex + 1
                            End If
                        Next
                        'i = i - 1
                        'liColIndex = liColIndex - 1
                    End If
                End If
                .Item(liColIndex).Text = moColl(i).Text
                .Item(liColIndex).Width = moColl(i).Width
                liColIndex = liColIndex + 1
            End If
        Next
        For i = liColIndex To .Count
            If InStr(1, .Item(j).Key, SortKeySuffix) = 0 Then .Remove i
        Next
    End With
    msTextKey = moColl(1).Key
End Sub

Friend Property Set ColHeaders(poHeaders As ColumnHeaders)
    Set moColHeaders = poHeaders
End Property

Public Function Count() As Long
    Count = moColl.Count
End Function

Public Sub SetColData(pyData() As Byte)
    On Error Resume Next
    Dim loColl As Collection
    Dim lbRedraw As Boolean
    Dim liPlace As Long
    Dim liUbound As Long
    Dim liInt As Integer
    Dim liLong As Long
    Dim lsString As String
    Dim loCol As cLVColumn
    
    Set loColl = moColl
    Set moColl = New Collection
    
    liUbound = UBound(pyData)
    
    lbRedraw = mbRedraw
    Redraw = False
    Do Until liPlace + 2 >= liUbound
        liInt = GetInt(pyData, liPlace)
        liPlace = liPlace + 2
        liLong = GetLong(pyData, liPlace)
        liPlace = liPlace + 4
        
        Set loCol = Nothing
        lsString = GetString(pyData, liPlace)
        liPlace = liPlace + Len(lsString) + 2
        Set loCol = loColl(lsString)
        If Not loCol Is Nothing Then
            
            loCol.Visible = CBool(liInt)
            loCol.Width = liLong
            moColl.Add loCol, lsString
            loColl.Remove lsString
        Else
            'Stop
        End If
    Loop
    For Each loCol In loColl
        loCol.Visible = False
        moColl.Add loCol, loCol.Key
    Next
    Redraw = lbRedraw
End Sub

Public Sub GetColData(pyBytes() As Byte)
    Dim lyReturn() As Byte
    Dim i As Long
    Dim loCol As cLVColumn
    Dim liPlace As Long
    Dim liLen As Long
    
    For i = 1 To moColl.Count
        Set loCol = moColl(i)
        liLen = Len(loCol.Key)
        If i = 1 Then
            ReDim lyReturn(0 To liLen + 8 - 1)
        Else
            ReDim Preserve lyReturn(0 To UBound(lyReturn) + liLen + 8)
        End If
        CopyMemory lyReturn(liPlace), CBool(loCol.Visible), 2
        CopyMemory lyReturn(liPlace + 2), CLng(loCol.Width), 4
        CopyMemory lyReturn(liPlace + 6), CInt(liLen), 2
        CopyMemory lyReturn(liPlace + 8), ByVal StrPtr(StrConv(loCol.Key, vbFromUnicode)), liLen
        liPlace = liPlace + 8 + liLen
    Next
    pyBytes = lyReturn
End Sub

Public Sub Clear()
    Set moColl = New Collection
    If mbRedraw Then ShowColumns
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Set NewEnum = moColl.[_NewEnum]
End Property

Public Sub SyncWidths()
    On Error Resume Next
    Dim loEach As cLVColumn
    For Each loEach In moColl
        loEach.Width = moColHeaders(loEach.Key).Width
    Next
End Sub

Public Sub SyncOrder()
    On Error Resume Next
    Dim loColl As Collection
    Dim i As Long
    
    Set loColl = moColl
    Set moColl = New Collection
    
    For i = 1 To moColHeaders.Count
        moColl.Add loColl(moColHeaders(i).Key), moColHeaders(i).Key
    Next
End Sub

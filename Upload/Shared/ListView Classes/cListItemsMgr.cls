VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListItemsMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moLV            As ListView

Private miRequestCount  As Long
Private miSelect        As Long

Public IconKey          As String

Event ListComplete(pbRefresh As Boolean, pbCountChanged As Boolean)

Public Sub Attach(poLV As ListView)
    Set moLV = poLV
End Sub

Public Sub Detach()
    Cancel
    Set moLV = Nothing
End Sub

Private Sub Class_Terminate()
    Detach
End Sub

Public Sub ShowData(ByVal poRS As ADODB.Recordset, ByVal pbClearFirst As Boolean, Optional ByVal piSelect As Long)
    miSelect = piSelect
    IncCount
    If pbClearFirst Then
        FillList poRS
    Else
        RefreshList poRS
    End If
End Sub

Public Sub Cancel()
    IncCount
End Sub

Private Sub IncCount()
    If miRequestCount > 2140000000 Then miRequestCount = 1 Else miRequestCount = miRequestCount + 1
End Sub

Private Sub FillList(ByVal poRS As ADODB.Recordset)
    Dim loFields            As ADODB.Fields
    Dim loListItems         As ListItems
    Dim loListItem          As ListItem
    Dim loColHeaders        As ColumnHeaders
    Dim loColumn            As ColumnHeader
    Dim lsKey               As String
    Dim lsText              As String
    Dim loDoEvents          As cDoEvents
    Dim lbCanceled          As Boolean
    Dim loRS                As ADODB.Recordset
    Dim lCollColumns        As Collection
    Dim lsTextKey           As String
    Dim lvEachOne           As Variant
    Dim liKey               As Long
    
    On Error GoTo exitloop
    
    Set loDoEvents = New cDoEvents
    loDoEvents.Tag = miRequestCount

    Set loColHeaders = moLV.ColumnHeaders
    Set loListItems = moLV.ListItems
    
    Set lCollColumns = New Collection
    For Each loColumn In loColHeaders
        With loColumn
            If Not loColumn.SubItemIndex = 0 Then
                lCollColumns.Add .Key, .Key
            Else
                lsTextKey = .Key
            End If
        End With
    Next
    Set loColHeaders = Nothing
    
    loListItems.Clear
    Set loRS = poRS.Clone(adLockReadOnly)
    moLV.Visible = False
    With loRS
        Do
            '#############INLINE
            'lvEachOne, lsKey, loListItem, loListItems, lCollColumns, loFields
            If .EditMode < adEditDelete Then
                Set loFields = .Fields
                liKey = .Fields(gDataKey).Value
                Set loListItem = loListItems.Add(, liKey & gsKeySuffix, .Fields(lsTextKey).Value, IconKey, IconKey)
                
                If liKey = miSelect Then
                    Set moLV.SelectedItem = loListItem
                    loListItem.Selected = True
                End If
                
                With loListItem.ListSubItems
                    For Each lvEachOne In lCollColumns
                        lsKey = lvEachOne
                        If InStr(1, lsKey, SortKeySuffix) = 0 Then
                            lsText = loFields(lsKey).Value & ""
                            If IsDate(lsText) Then lsText = Format$(lsText, DateFormat)
                           .Add , lsKey, lsText
                        End If
                    Next
                End With
            End If
            '#############ENDINLINE
            With loDoEvents
                If .DoUserInput Then
                    If .Tag <> miRequestCount Then
                        lbCanceled = True
                        GoTo exitloop
                    End If
                End If
            End With
            .MoveNext
        Loop
exitloop:
    End With
    On Error Resume Next
    If Not loRS Is Nothing Then
        If loRS.State <> adStateClosed Then loRS.Close
    End If
    Set loRS = Nothing
    Set loFields = Nothing
    Set loListItems = Nothing
    Set loListItem = Nothing
    moLV.SelectedItem.EnsureVisible
    moLV.Visible = True
    'loListItems.Item(1).Selected = False
    'Set moLV.SelectedItem = Nothing
    'RaiseEvent ListComplete(False, True)

End Sub

Private Sub RefreshList(ByVal poRS As ADODB.Recordset)
    Dim loFields            As ADODB.Fields
    Dim loListItems         As ListItems
    Dim loListItem          As ListItem
    Dim loColHeaders        As ColumnHeaders
    Dim lvEachOne           As Variant
    Dim lsKey               As String
    Dim lsText              As String
    Dim lbCanceled          As Boolean
    Dim lsTextKey           As String
    Dim lCollColumns        As Collection
    Dim loColumn            As ColumnHeader
    Dim liKey               As Long
        
    Dim loDoEvents As cDoEvents
    Set loDoEvents = New cDoEvents
    loDoEvents.Tag = miRequestCount

    Dim lbCountChanged As Boolean
    On Error Resume Next
    Dim lcollRequery As CollectionWrapper
    Set lcollRequery = New CollectionWrapper
    Dim lrs As ADODB.Recordset
    Set lrs = poRS.Clone(adLockReadOnly)  ' Must use clone b/c yielding control of processor
    Set loListItems = moLV.ListItems
    
    Set loColHeaders = moLV.ColumnHeaders
    Set lCollColumns = New Collection
    For Each loColumn In loColHeaders
        With loColumn
            If Not loColumn.SubItemIndex = 0 Then
                lCollColumns.Add .Key, .Key
            Else
                lsTextKey = .Key
            End If
        End With
    Next
    Set loColHeaders = Nothing
    
    
    
    With lrs
        '.MoveFirst
        Do While Not .EOF
            '#############INLINE MUST HAVE ON ERROR RESUME NEXT
            'lvEachOne, lsKey, loListItem, loListItems, lCollColumns, loFields, lsText
            If .EditMode < adEditDelete Then
                Set loFields = .Fields
                lsKey = lsTextKey
                liKey = loFields(gDataKey).Value
                Set loListItem = loListItems(liKey & gsKeySuffix)
                If loListItem Is Nothing Then
                    Set loListItem = loListItems.Add(, liKey & gsKeySuffix, , IconKey, IconKey)
                    lbCountChanged = True
                End If
                'If liKey = miSelect Then
                    'Set moLV.SelectedItem = loListItem
                    'loListItem.Selected = True
                'End If
                lsText = .Fields(lsTextKey).Value
                With loListItem
                    If lsText <> .Text Then
                        .Text = lsText
                    End If
                    With .ListSubItems
                        For Each lvEachOne In lCollColumns
                            lsKey = lvEachOne
                            If InStr(1, lsKey, SortKeySuffix) = 0 Then
                                lsText = loFields(lsKey).Value & ""
                                If IsDate(lsText) Then lsText = Format$(lsText, DateFormat)
                                Err.Clear
                                If lsText <> .Item(lsKey).Text Then
                                    If Err.Number = 0 Then .Item(lsKey).Text = lsText Else .Add , lsKey, lsText
                                End If
                            End If
                        Next
                    End With
                    lcollRequery.Add "", .Key
                End With
                Set loListItem = Nothing
            End If
            '#############ENDINLINE
            If loDoEvents.DoUserInput Then 'Yield if user-related messages in queue
                If loDoEvents.Tag <> miRequestCount Then
                    lbCanceled = True
                    GoTo canceled
                End If
            End If
            .MoveNext
        Loop
    End With
        
    Dim lCollRemove As Collection
    Set lCollRemove = New Collection
    
    With lcollRequery
        If lcollRequery.Count < loListItems.Count Then
            For Each lvEachOne In loListItems
                If Not .KeyExists(lvEachOne.Key) Then lCollRemove.Add lvEachOne.Key
                If .Count + lCollRemove.Count = loListItems.Count Then Exit For
            Next
            For Each lvEachOne In lCollRemove
                loListItems.Remove lvEachOne
                lbCountChanged = True
            Next
        End If
    End With

canceled:
    
    Set loFields = Nothing
    Set loListItems = Nothing
    Set loListItem = Nothing
    lrs.Close
    Set lrs = Nothing
    If lbCanceled Then Exit Sub
    If loDoEvents.Tag <> miRequestCount Then Exit Sub
    'RaiseEvent ListComplete(True, lbCountChanged)
End Sub

Public Sub SetSelectedRecord(ByVal piKey As Long)
    On Error Resume Next
    With moLV
        Set .SelectedItem = Nothing
        Set .SelectedItem = .ListItems(piKey & gsKeySuffix)
        .SelectedItem.EnsureVisible
    End With
End Sub

Public Sub UpdateSubItems(poFields As cDataRecord)
    On Error Resume Next
    Dim lsText As String
    Dim loColumn As ColumnHeader
    Dim liSII As Long
    
    With moLV.ListItems(poFields.Field(gDataKey).Value & gsKeySuffix)
        For Each loColumn In moLV.ColumnHeaders
            liSII = loColumn.SubItemIndex
            lsText = poFields.Field(loColumn.Key).Value
            If IsDate(lsText) Then lsText = Format(lsText, DateFormat)
            If liSII = 0 Then
                If .Text <> lsText Then .Text = lsText
            Else
                If .SubItems(liSII) <> lsText Then .SubItems(liSII) = lsText
            End If
        Next
    End With
End Sub

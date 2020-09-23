VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iFileTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Property Get Status() As eFileTaskStatus
'
End Property

Property Get CurrentFile() As String
'
End Property

Property Get Canceled() As Boolean
'
End Property
Property Let Canceled(ByVal RHS As Boolean)
'
End Property

Property Get BytesProgress() As Double
'
End Property
Property Get BytesTotal() As Double
'
End Property

Property Get Tag() As Long
'
End Property
Property Let Tag(ByVal RHS As Long)
'
End Property

Public Property Set Files(ByVal Files As cFiles)
'
End Property

Public Property Get Files() As cFiles
'
End Property
Public Property Get Errors() As Collection
'
End Property

Public Property Get Overwrite() As eFileConfirmation
'
End Property
Public Property Let Overwrite(ByVal RHS As eFileConfirmation)
'
End Property

Public Property Get Parent() As iFileTaskParent
'
End Property
Public Property Set Parent(ByVal RHS As iFileTaskParent)
'
End Property

Public Function Start() As Boolean
'
End Function


'Public Sub SetFiles(ByVal Data)
'    If miStatus > 0 Then Exit Function
'    On Error Resume Next
'    Dim lvTemp
'    Dim ltData As WIN32_FIND_DATA
'    Dim lsTemp As String
'    moFiles.Clear
'    Set moErrors = New Collection
'    For Each lvTemp In Data
'        If Not moFiles.AddObject(lvTemp) Then
'            lsTemp = ""
'            lsTemp = lvTemp
'            If FindSpecific(lsTemp, ltData) Then
'                moFiles.AddFile ltData
'            Else
'                moErrors.Add lvTemp
'            End If
'        End If
'    Next
'End Sub
'
'Public Property Get ChunkSize() As Long
'    ChunkSize = miChunksize
'End Property
'Public Property Let ChunkSize(piVal As Long)
'    If miStatus = 0 Then miChunksize = piVal
'End Property
'
'Private Sub Clear()
'    On Error Resume Next
'    Set moFiles = New cFiles
'    Set moErrors = New cFileTask
'    miTask = 0
'    miTag = 0
'    miStatus = 0
'    miConfirm = 0
'    msDestination = ""
'    msRelativeToPath = ""
'    mbCanceled = False
'End Sub
'
'Private Sub Class_Initialize()
'    Clear
'End Sub



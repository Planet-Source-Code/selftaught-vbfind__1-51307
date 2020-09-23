VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileCopyMoveDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eCMDAction
    ftCopy = 1
    ftMove
    ftDelete
End Enum

Implements iFileTask

Private mtTask As tFileTask
Private mbCallback As Boolean
Private miFlags As Long
Private mdblCompleteFiles As Double
Private mdblLastEvent As Double
Private miCancel As Long

Private Sub Clear()
    With mtTask
        .Canceled = True
        .Progress = 0
        .Total = 0
        .Status = 0
        .File = ""
    End With
End Sub

Private Property Get iFileTask_BytesProgress() As Double
    iFileTask_BytesProgress = mtTask.Progress
End Property

Private Property Get iFileTask_BytesTotal() As Double
    iFileTask_BytesTotal = mtTask.Total
End Property

Private Property Get iFileTask_Canceled() As Boolean
    iFileTask_Canceled = mtTask.Canceled
End Property

Private Property Let iFileTask_Canceled(ByVal RHS As Boolean)
    If RHS Then miCancel = True
    mtTask.Canceled = RHS
End Property

Private Property Get iFileTask_CurrentFile() As String
    iFileTask_CurrentFile = mtTask.File
End Property

Private Property Get iFileTask_Errors() As Collection
    Set iFileTask_Errors = mtTask.Errors
End Property

Private Property Set iFileTask_Files(ByVal RHS As cFiles)
    Set mtTask.Files = RHS.Clone
End Property

Private Property Get iFileTask_Files() As cFiles
    Set iFileTask_Files = mtTask.Files
End Property

Private Property Get iFileTask_Overwrite() As eFileConfirmation
    iFileTask_Overwrite = mtTask.OverwriteConfirm
End Property

Private Property Let iFileTask_Overwrite(ByVal RHS As eFileConfirmation)
    mtTask.OverwriteConfirm = RHS
End Property

Private Property Set iFileTask_Parent(ByVal RHS As iFileTaskParent)
    Set mtTask.Parent = RHS
End Property

Private Property Get iFileTask_Parent() As iFileTaskParent
    Set mtTask.Parent = mtTask.Parent
End Property

Private Function iFileTask_Start() As Boolean
    If mtTask.Status > 0 Then Err.Raise 5
    On Error Resume Next
    iFileTask_Start = True
    If Not mbCallback Then
        mtTask.Canceled = False
        mtTask.Status = ftsWaitingForCallback
        mbCallback = True
        TaskCallback Me
        Exit Function
    Else
        mbCallback = False
    End If

    With mtTask
        .Status = ftsJustStarting
        .Total = TotalSize
        Set .Errors = New Collection
        If PulseProgress(0) Then
            Select Case mtTask.Task
                Case ftMove
                    CopyFiles True
                Case ftDelete
                    DeleteFiles
                Case ftCopy
                    CopyFiles False
            End Select
        End If
    End With
    Clear
End Function

Private Property Get iFileTask_Status() As eFileTaskStatus
    iFileTask_Status = mtTask.Status
End Property

Private Property Let iFileTask_Tag(ByVal RHS As Long)
    mtTask.Tag = RHS
End Property

Private Property Get iFileTask_Tag() As Long
    iFileTask_Tag = mtTask.Tag
End Property




Friend Sub Callback()
    If mtTask.Status = ftsWaitingForCallback Then
        mtTask.Status = 0
        iFileTask_Start
    End If
End Sub







Public Property Get CurrentTask() As eCMDAction
    CurrentTask = mtTask.Task
End Property

Public Property Let CurrentTask(ByVal RHS As eCMDAction)
    Select Case RHS
        Case ftCopy To ftDelete
            If mtTask.Status = 0 Then: mtTask.Task = RHS: Exit Property
    End Select
    Err.Raise 5
End Property

Public Property Let Target(ByVal RHS As String)
    If mtTask.Status = 0 Then mtTask.Target = RHS Else Err.Raise 5
End Property
Public Property Get Target() As String
    Target = mtTask.Target
End Property

Public Property Let RelativeToFolder(ByVal RHS As String)
    If mtTask.Status = 0 Then mtTask.RelativeFolder = RHS Else Err.Raise 5
End Property
Public Property Get RelativeToFolder() As String
    RelativeToFolder = mtTask.RelativeFolder
End Property







Private Property Get TotalSize() As Double
    On Error Resume Next
    Dim loFile As cFile
    For Each loFile In mtTask.Files
        TotalSize = TotalSize + loFile.Size
    Next
End Property


Private Sub Class_Initialize()
    On Error Resume Next
    miFlags = COPY_FILE_FAIL_IF_EXISTS Or COPY_FILE_RESTARTABLE
    Set mtTask.Files = New cFiles
End Sub



Private Function PulseProgress(ByVal pdblBytes As Double, Optional ByVal pbSilent As Boolean) As Boolean
    On Error Resume Next
    With mtTask
        If .Canceled Then .Status = .Status Or ftsCanceled
        Select Case mtTask.Status
            Case ftsFinishing
                .Progress = .Total
            Case Else
                .Progress = .Progress + pdblBytes
                If .Progress > .Total Then .Total = .Progress
                If mtTask.Status Mod ftsCanceled = ftsCheckingIn Then pbSilent = pbSilent Or (.Canceled Or .Status = ftsCanceled)
        End Select
        If Not pbSilent Then: mdblLastEvent = .Progress: .Parent.Notify .Tag
        If .Canceled Then .Status = .Status Or ftsCanceled Else PulseProgress = True
        'If Not pbSilent Then Debug.Print .Progress / .Total, .File, .Status
    End With
End Function


Public Property Get Overwrite() As Boolean
    Overwrite = Not (miFlags Or COPY_FILE_FAIL_IF_EXISTS = miFlags)
End Property
Public Property Let Overwrite(ByVal pbVal As Boolean)
    If Not pbVal Then
        miFlags = miFlags Or COPY_FILE_FAIL_IF_EXISTS
    Else
        miFlags = (miFlags Or COPY_FILE_FAIL_IF_EXISTS) - COPY_FILE_FAIL_IF_EXISTS
    End If
End Property
Public Property Get Restartable() As Boolean
    Restartable = miFlags Or COPY_FILE_RESTARTABLE = miFlags
End Property
Public Property Let Restartable(ByVal pbVal As Boolean)
    If pbVal Then
        miFlags = miFlags Or COPY_FILE_RESTARTABLE
    Else
        miFlags = (miFlags Or COPY_FILE_RESTARTABLE) - COPY_FILE_RESTARTABLE
    End If
End Property

Private Sub DeleteFiles()
    On Error Resume Next
    Dim loFile As cFile
    Dim liStatus As eFileTaskStatus

    With mtTask
        .Status = ftsCheckingIn
        For Each loFile In .Files
            .File = loFile.FullPath
            If ConfirmOverwrite(.File, .OverwriteConfirm) Then: liStatus = .Status: .Status = ftsConfirming
            If Not PulseProgress(loFile.Size) Then Exit For
            If liStatus <> 0 Then mtTask.Status = liStatus: liStatus = 0
            If Not FileDelete(.File, True) Then .Errors.Add .File, .File
        Next
        RemoveEmptyFolders
        .Status = .Status - (.Status Mod ftsCanceled) + ftsFinishing
        PulseProgress 0
    End With
End Sub

Private Sub CopyFiles(ByVal pbThenDelete As Boolean)
    On Error Resume Next
    Dim loFile As cFile
    Dim liStatus As eFileTaskStatus
    Dim ldblBytes As Double
    With mtTask
        .Status = ftsCheckingIn
        For Each loFile In .Files
            .File = loFile.FullPath
            If ConfirmOverwrite(.File, .OverwriteConfirm) Then
                liStatus = .Status
                .Status = ftsConfirming
            End If
            If Not PulseProgress(0) Then Exit Sub
            If liStatus > 0 Then .Status = liStatus: liStatus = 0
            If Not CopyFile(.File, TransformPath(ftCopy, .Target, .RelativeFolder, .File), miCancel, miFlags, Me) Then
                .Errors.Add .File, .File
            Else
                If pbThenDelete Then FileDelete .File, True
            End If
            ldblBytes = mdblCompleteFiles + loFile.Size
            ldblBytes = ldblBytes - .Progress
            If ldblBytes < 0 Then ldblBytes = 0
            If Not PulseProgress(ldblBytes, True) Then mdblCompleteFiles = .Progress: Exit For Else mdblCompleteFiles = .Progress
        Next
        If pbThenDelete Then RemoveEmptyFolders
        .Status = .Status - (.Status Mod ftsCanceled) + ftsFinishing
        PulseProgress 0
    End With
End Sub

Private Sub RemoveEmptyFolders()
    Dim loColl As Collection
    Dim lvTemp
    Dim ltFind As tFindFiles
    With ltFind
        .Filter = "*"
        .Recurse = True
        For Each lvTemp In GetFolders(mtTask.Files)
            .Path = lvTemp
            Set loColl = FindFiles(ltFind)
            If loColl.Count = 0 Then PathRemove .Path
        Next
    End With
End Sub

Friend Function Notify(pcTransferred As Double, pcTotal As Double) As eCopyProgress
    On Error Resume Next
    Dim ldblBytes As Double
    With mtTask
        If .Canceled Then
            Notify = PROGRESS_CANCEL
            Exit Function
        Else
            ldblBytes = mdblCompleteFiles + pcTransferred - .Progress
            If ldblBytes < 0 Then ldblBytes = 0
            PulseProgress ldblBytes, (.Progress + ldblBytes) - mdblLastEvent <= 1048576
            Notify = PROGRESS_CONTINUE
        End If
    End With
End Function

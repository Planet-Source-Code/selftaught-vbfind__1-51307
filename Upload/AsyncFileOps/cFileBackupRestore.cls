VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileBackupRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eFindFileRange
    ffrAccessed
    ffrCreated
    ffrModified
    ffrSize
End Enum

Public Enum eFindFileIgnoreTypes
    ffiReadOnly
    ffiHidden
    ffiSystem
    ffiTemporary
End Enum

Public Enum eBackupRestoreType
    ftCryptoZip = 4
    ftCryptoUnZip
    ftEncrypt
    ftDecrypt
    ftZip
    ftUnzip
End Enum

Implements iFileTask

Private mtTask As tFileTask
Private moFileIO As cFileIO
Private mbCallback As Boolean

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
            If Restoring Then Restore Else Backup
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







Public Property Get CurrentTask() As eBackupRestoreType
    CurrentTask = mtTask.Task
End Property

Public Property Let CurrentTask(ByVal RHS As eBackupRestoreType)
    Select Case RHS
        Case ftCryptoZip To ftUnzip
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

Private Sub Backup()
    Dim loFile As cFile
    Dim loFiles As Collection
    Dim lbCanceled As Boolean
    
    On Error Resume Next
    
    With mtTask
        .File = .Target
        If .OverwriteConfirm <> fcNone Then
            If ConfirmOverwrite(.File, .OverwriteConfirm) Then
                .Status = ftsConfirming
                If Not PulseProgress(0) Then Exit Sub
            End If
        End If
        .Status = .Status - (.Status Mod ftsCanceled) + ftsCheckingIn
    End With
    
    With moFileIO
        .FileAccess = GENERIC_WRITE
        .FileAttributes = FILE_ATTRIBUTE_NORMAL
        .FileCreation = CREATE_ALWAYS
        .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
        .FileShare = FILE_SHARE_READ
        If .OpenFile(mtTask.Target) Then
            .AppendString gsSalt
            With mtTask
                For Each loFile In .Files
                    .File = loFile.FullPath
                    If Not PackFile Then .Errors.Add .File, .File
                    If .Canceled Then Exit For
                Next
            End With
            .CloseFile
        Else
            mtTask.Status = (mtTask.Status Mod ftsCanceled) + ftsError
        End If
    End With
    mtTask.Status = mtTask.Status - (mtTask.Status Mod ftsCanceled) + ftsFinishing
    PulseProgress 0
End Sub

Private Sub PackString(ByVal psString As String)
    Select Case mtTask.Task
        Case ftZip, ftUnzip
        Case Else
            psString = EncryptString(psString, gsWord)
    End Select
    moFileIO.AppendString psString
End Sub

Private Sub UnpackString(psString As String, SALT As String)
    With moFileIO
        .GetString psString
        Select Case mtTask.Task
            Case ftZip, ftUnzip
            Case Else
                psString = DecryptString(psString, gsWord, SALT)
        End Select
    End With
End Sub


Private Sub Restore()
    Dim loFile As cFile
    
    Dim liNum As Long
    
    Dim lbVal As Boolean
    
    Dim lsTemp As String
    Dim lsSALT As String
    Dim lsString As String
    
    On Error Resume Next
    
    With moFileIO
        mtTask.Status = ftsCheckingIn
        If PathCreate(mtTask.Target) Then
            For Each loFile In mtTask.Files
                .FileAccess = GENERIC_READ
                .FileCreation = OPEN_EXISTING
                .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
                .FileShare = FILE_SHARE_READ
                mtTask.File = loFile.FullPath
                If .OpenFile(mtTask.File) Then
                    .GetString lsSALT
                    PulseProgress Len(lsSALT) + 2, True
                    lbVal = True
                    Do While Not .EOF
                        lbVal = lbVal And UnpackFile(lsSALT)
                        If mtTask.Canceled Then Exit Do
                    Loop
                    If Not lbVal Then mtTask.Errors.Add mtTask.File, mtTask.File
                    .CloseFile
                End If
            Next
        Else
            mtTask.Status = (mtTask.Status Mod ftsCanceled) + ftsError
        End If
    End With
    mtTask.Status = mtTask.Status - (mtTask.Status Mod ftsCanceled) + ftsFinishing
    PulseProgress 0
End Sub

Private Function UnpackFile(SALT As String) As Boolean
    On Error Resume Next
    mtTask.File = vbNullString
    Dim loFile As cFileIO
    Set loFile = New cFileIO
    Dim lyBytes() As Byte
    Dim liTotalSize As Long
    Dim liByteCount As Long
    Dim liChunksize As Long
    Dim liCZChunksize As Long
    Dim liStatus As eFileTaskStatus
    
    Dim liTemp As Long
    
    With loFile
        .FileAccess = GENERIC_WRITE
        .FileAttributes = FILE_ATTRIBUTE_NORMAL
        .FileCreation = CREATE_ALWAYS
        .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
    End With
    
    With moFileIO
        .GetLong liTotalSize
        With mtTask
            UnpackString .File, SALT
            .File = TransformPath(ftCryptoUnZip, .Target, "", .File)
            If ConfirmOverwrite(.File, mtTask.OverwriteConfirm) Then
                liStatus = .Status
                .Status = ftsConfirming
            End If
            If Not PulseProgress(Len(.File) + 6) Then Exit Function
            If liStatus > 0 Then .Status = liStatus
            If Not loFile.OpenFile(.File) Then Exit Function
        End With
        
        Do
            .GetLong liChunksize
            .GetLong liCZChunksize
            .GetBytes lyBytes, liCZChunksize
            liTemp = UBound(lyBytes) + 9
            liByteCount = liByteCount + liChunksize
            Select Case mtTask.Task
                Case ftEncrypt, ftDecrypt
                    lyBytes = DecryptByteArray(lyBytes, gsWord, SALT)
                Case ftZip, ftUnzip
                    DecompressByteArray lyBytes, liChunksize
                Case Else
                    lyBytes = DecryptByteArray(lyBytes, gsWord, SALT)
                    DecompressByteArray lyBytes, liChunksize
            End Select
            loFile.AppendBytes lyBytes
            If Not liByteCount >= liTotalSize Then
                If Not PulseProgress(liTemp) Then Exit Function
            Else
                PulseProgress liTemp, True
                Exit Do
            End If
        Loop Until .EOF
    End With
    UnpackFile = True
End Function

Private Function PackFile() As Boolean
    On Error Resume Next
    Dim liChunksize As Long
    Dim lyBytes() As Byte
    Dim loFile As cFileIO
    Set loFile = New cFileIO
    
    With loFile
        .FileAccess = GENERIC_READ
        .FileCreation = OPEN_EXISTING
        .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
        .FileShare = FILE_SHARE_READ
        If Not .OpenFile(mtTask.File) Then Exit Function
        If Not PulseProgress(0) Then Exit Function
    End With
    
    With moFileIO
        .AppendLong loFile.FileLength
        PackString TransformPath(ftCryptoZip, "", mtTask.RelativeFolder, mtTask.File)
        Do
            loFile.GetBytes lyBytes
            liChunksize = UBound(lyBytes) + 1
            Select Case mtTask.Task
                Case ftEncrypt, ftDecrypt
                    lyBytes = EncryptByteArray(lyBytes, gsWord)
                Case ftZip, ftUnzip
                    CompressByteArray lyBytes
                Case Else
                    CompressByteArray lyBytes
                    lyBytes = EncryptByteArray(lyBytes, gsWord)
            End Select
            .AppendLong liChunksize
            .AppendLong UBound(lyBytes) + 1
            .AppendBytes lyBytes
            If loFile.EOF Then
                PulseProgress liChunksize, True
                Exit Do
            Else
                If Not PulseProgress(liChunksize) Then Exit Function
            End If
        Loop
    End With
    PackFile = True
End Function

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
        If Not pbSilent Then .Parent.Notify .Tag
        If .Canceled Then .Status = .Status Or ftsCanceled Else PulseProgress = True
    End With
End Function


Private Sub Class_Initialize()
    On Error Resume Next
    Debug.Print "gsWord" & Len(gsWord)
    If Len(gsWord) = 0 Then gsWord = DecryptString("È5e""YqCŠ-ËDà¨œÄÁòÇjµ–ñ", "File I/O Error", "53774951582")
    If Len(gsSalt) = 0 Then gsSalt = CryptoSALT
    Set moFileIO = New cFileIO
    Set mtTask.Files = New cFiles
End Sub


Private Function Restoring() As Boolean
    Restoring = mtTask.Task Mod 2 = 1
End Function

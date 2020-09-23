Attribute VB_Name = "mTask"
Option Explicit

Public Const ftsWaitingForCallback = &HFFFFFFF

Public Type tFileTask
    Canceled As Boolean
    Tag As Long
    Progress As Double
    Total As Double
    OverwriteConfirm As eFileConfirmation
    Status As eFileTaskStatus
    Task As Long
    File As String
    Target As String
    RelativeFolder As String
    Parent As iFileTaskParent
    Files As cFiles
    Errors As Collection
End Type


Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private moColl As Collection
Private miTimer As Long

Public Sub TaskCallback(ByVal ForMe As iFileTask)
    If moColl Is Nothing Then Set moColl = New Collection
    moColl.Add ForMe
    If miTimer = 0 Then miTimer = SetTimer(0, 0, 1, AddressOf TimerProc)
End Sub

Private Sub TimerProc(ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    On Error Resume Next
    
    KillTimer 0, miTimer
    miTimer = 0
    
    Dim i As Long
    Dim iCB As cFileBackupRestore
    Dim iCB2 As cFileCopyMoveDelete
    Dim iCB3 As cFileSearch
    
    For i = 1 To moColl.Count
        Set iCB = moColl(1)
        Set iCB2 = moColl(1)
        Set iCB3 = moColl(1)
        moColl.Remove 1
        Select Case True
            Case Not iCB Is Nothing
                iCB.Callback
            Case Not iCB2 Is Nothing
                iCB2.Callback
            Case Not iCB3 Is Nothing
                iCB3.Callback
        End Select
    Next
End Sub








Public Function ConfirmOverwrite(psFile As String, piMode As eFileConfirmation) As Boolean
    If piMode <> fcNone Then
        If FileExists(psFile) Then
            If piMode = fcAll Then ConfirmOverwrite = True Else ConfirmOverwrite = FileGetAttributes(psFile) And FILE_ATTRIBUTE_READONLY
        End If
    End If
End Function

Public Function TransformPath(piTask As Long, psDestination As String, psRelativeTo As String, psFile As String) As String
    Select Case piTask
        Case ftCopy, ftMove
'        dest: Copying Relative To
'        RelativeTo: Copying Relative From
'        File: full filename to copy from
'        result: destination path
            If LenB(psRelativeTo) = 0 Or Not PathMatchPrefix(psFile, psRelativeTo) Then
                TransformPath = PathBuild(psDestination, PathGetFileName(psFile))
            Else
                TransformPath = PathGetAbsolute(PathBuild(psDestination, PathGetRelative(psRelativeTo, psFile)))
            End If
'        Case ftDelete
'           function not used
        Case ftCryptoZip, ftEncrypt, ftZip
'        dest: n/a
'        RelativeTo: Zipping Relative To
'        File: Full filenames
'        result: relative output path to be saved with the file
            If LenB(psRelativeTo) = 0 Or Not PathMatchPrefix(psFile, psRelativeTo) Then
                TransformPath = PathGetFileName(psFile)
            Else
                TransformPath = PathGetRelative(psRelativeTo, psFile)
            End If
        Case ftCryptoUnZip, ftDecrypt, ftUnzip
'        dest: Unzipping to
'        RelativeTo: n/a
'        File: Relative filenames
'        result: output path for file
            TransformPath = PathGetAbsolute(PathBuild(psDestination, psFile))

    End Select
End Function

Public Function GetFolders(ByVal poFiles As cFiles) As Collection
    Dim loFile As cFile
    On Error Resume Next
    Set GetFolders = New Collection
    For Each loFile In poFiles
        GetFolders.Add loFile.FilePath, loFile.FilePath
    Next
End Function


Public Sub testfind()
    Dim loTask As iFileTask
    Dim loFind As cFileSearch
    Set loFind = New cFileSearch
    Set loTask = loFind
    With loFind
        .Path = "C:\"
        .Filter = "*.txt"
        .Recursive = False
        .ChunkSize = 5
        .FindContainedText = "find This!!"
    End With
    With loTask
        .Start
    End With
End Sub

Public Sub test(ByVal Move As Boolean)
    Dim loTask As iFileTask
    Dim lot As cFileCopyMoveDelete
    Set lot = New cFileCopyMoveDelete
    Set loTask = lot
    If Move Then
        lot.CurrentTask = ftMove
        'lot.RelativeToFolder
    Else
        lot.CurrentTask = ftCopy
        loTask.Files.AddFolder "C:\new vb projects\", True
        lot.RelativeToFolder = "C:\new vb projects"
        lot.Target = "C:\testing"
    End If
    loTask.Start
'
'
'    End Select
    
'    If Restore Then
'        lot.CurrentTask = ftCryptoUnZip
'        lotask.Files.AddFile "C:\testing.czb"
'        lot.Target = "C:\testing\"
'    Else
'        lot.CurrentTask = ftCryptoZip
'        lotask.Files.AddFile "C:\new vb projects\usercontrol.rtf"
'        lotask.Files.AddFile "C:\new vb projects\vb projects.zip"
'        lotask.Files.AddFile "C:\new vb projects\todo.txt"
'        lotask.Files.AddFile "C:\new vb projects\project descriptions.rtf"
'        lot.RelativeToFolder = "c:\new vb projects\"
'        lot.Target = "C:\testing.czb"
'    End If
'    lotask.Start
    
End Sub


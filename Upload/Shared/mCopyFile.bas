Attribute VB_Name = "mCopyFile"
Option Explicit

Public Enum eCopyFileFlags
    COPY_FILE_FAIL_IF_EXISTS = 1
    COPY_FILE_RESTARTABLE = 2
End Enum

Public Enum eCopyProgress
    PROGRESS_CANCEL = 1
    PROGRESS_CONTINUE = 0
    PROGRESS_QUIET = 3
    PROGRESS_STOP = 2
End Enum

Private Declare Function CopyFileAPI Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal lpProgressRoutine As Long, lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long

Private moCopier As cFileCopyMoveDelete

Public Function CopyFile(psFrom As String, _
                         psTo As String, _
                   ByRef piCancel As Long, _
                   ByVal piFlags As eCopyFileFlags, _
                   Optional ByVal poObj As cFileCopyMoveDelete) _
                As Boolean
    
    On Error Resume Next
    If Not PathCreate(PathGetParentFolder(psTo)) Then Exit Function
    If poObj Is Nothing Then
        CopyFile = CopyFileAPI(psFrom, psTo, (piFlags Or COPY_FILE_FAIL_IF_EXISTS) = piFlags) <> 0
    Else
        piCancel = 0
        Set moCopier = poObj
        CopyFile = CopyFileEx(psFrom, psTo, AddressOf CopyProgressRoutine, 0, piCancel, piFlags) <> 0
        Set moCopier = Nothing
    End If
End Function

Private Function CopyProgressRoutine(ByVal TotalFileSize As Currency, ByVal TotalBytesTransferred As Currency, ByVal StreamSize As Currency, ByVal StreamBytesTransferred As Currency, ByVal dwStreamNumber As Long, ByVal dwCallbackReason As Long, ByVal hSourceFile As Long, ByVal hDestinationFile As Long, ByVal lpData As Long) As eCopyProgress
    On Error Resume Next
    If Not moCopier Is Nothing Then CopyProgressRoutine = moCopier.Notify(TotalBytesTransferred * 10000, TotalFileSize * 10000)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileTaskThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eFileTaskStatus
    ftsConfirming = 1
    ftsJustStarting
    ftsCheckingIn
    ftsFinishing
    ftsCanceled = 8
    ftsError = 16
End Enum

Public Enum eFileConfirmation
    fcNone
    fcAll
    fcReadonly
End Enum

Public Enum eFileTaskObjects
    ftoBackupRestore
    ftoCopyMoveDelete
    ftoSearch
End Enum

Public Function GetNewFileTask(piObj As eFileTaskObjects) As iFileTask
    Select Case piObj
        Case ftoBackupRestore
            Set GetNewFileTask = New cFileBackupRestore
        Case ftoCopyMoveDelete
            Set GetNewFileTask = New cFileCopyMoveDelete
        Case ftoSearch
            Set GetNewFileTask = New cFileSearch
    End Select
End Function

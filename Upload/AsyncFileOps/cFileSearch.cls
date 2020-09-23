VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iFileTask

Private miChunksize As Long
Private mtWin32 As WIN32_FIND_DATA
Private mtFind As tFindFiles

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lplsFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private moSearchNames As Collection

'Used for searching for text in files
Private msText As String
Private myMainArray() As Byte
Private myTempArray() As Byte

Private mtTask As tFileTask
Private miDoubleTextLen As Long

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
    mtTask.Canceled = False
    mtTask.Status = ftsWaitingForCallback
    TaskCallback Me
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
        DoFind
    End If
End Sub



Private Function PulseProgress(Optional ByVal pbSilent As Boolean) As Boolean
    On Error Resume Next
    If miChunksize > 0 Then
        If ((mtTask.Status Mod ftsCanceled) = ftsCheckingIn) And (Not mtFind.FileCount Mod miChunksize = 0) Then pbSilent = True
    Else
        pbSilent = Not mtTask.Status And ftsFinishing
    End If
    With mtTask
        If .Canceled Then .Status = .Status Or ftsCanceled
        If Not pbSilent Then .Parent.Notify .Tag
        If .Canceled Then .Status = .Status Or ftsCanceled Else PulseProgress = True
        'If Not pbSilent Then Debug.Print mtFind.FileCount
            'Dim i As Long
            'For i = mtTask.Files.Count - miChunksize + 1 To mtTask.Files.Count
                'Debug.Print mtTask.Files(i).FullPath,
            'Next
        'End If
    End With
    'Debug.Assert Not mtTask.Status And ftsFinishing
End Function


Private Sub Class_Initialize()
    On Error Resume Next
    Set mtTask.Files = New cFiles
    mtFind.Filter = "*"
    mtFind.Recurse = True
End Sub




Public Property Set SearchNames(ByVal Names As Collection)
    Set moSearchNames = Names
End Property

Public Property Let FindContainedText(ByVal StringValue As String)
    If mtTask.Status > 0 Then Exit Property
    msText = StringValue
End Property
Public Property Get FindContainedText() As String
    FindContainedText = msText
End Property

Public Property Let ChunkSize(ByVal Value As Long)
    If mtTask.Status > 0 Then Exit Property
    If Value < 0 Then Value = 0
    miChunksize = Value
End Property
Public Property Get ChunkSize() As Long
    ChunkSize = miChunksize
End Property

Public Property Get Recursive() As Boolean
    Recursive = mtFind.Recurse
End Property
Public Property Let Recursive(ByVal Value As Boolean)
    If mtTask.Status > 0 Then Exit Property
    mtFind.Recurse = Value
End Property

Public Property Get Filter() As String
    Filter = mtFind.Filter
End Property
Public Property Let Filter(ByVal StringValue As String)
    If mtTask.Status > 0 Then Exit Property
    If LenB(StringValue) = 0 Then StringValue = "*"
    mtFind.Filter = StringValue
End Property

Public Property Get Path() As String
    Path = mtFind.Path
End Property
Public Property Let Path(ByVal StringValue As String)
    If mtTask.Status > 0 Then Exit Property
    mtFind.Path = StringValue
End Property

Public Property Let Ignore(ByVal WhatIgnore As eFindFileIgnoreTypes, ByVal Value As Boolean)
    If mtTask.Status > 0 Then Exit Property
    With mtFind
        Select Case WhatIgnore
            Case ffiHidden
                .IgnoreHidden = Value
            Case ffiReadOnly
                .IgnoreReadOnly = Value
            Case ffiSystem
                .IgnoreSystem = Value
            Case ffiTemporary
                .IgnoreTemp = Value
        End Select
    End With
End Property
Public Property Get Ignore(ByVal WhatIgnore As eFindFileIgnoreTypes) As Boolean
    With mtFind
        Select Case WhatIgnore
            Case ffiHidden
                Ignore = .IgnoreHidden
            Case ffiReadOnly
                Ignore = .IgnoreReadOnly
            Case ffiSystem
                Ignore = .IgnoreSystem
            Case ffiTemporary
                Ignore = .IgnoreTemp
        End Select
    End With
End Property

Public Sub SetRange(ByVal Range As eFindFileRange, ByVal LowValue As Double, ByVal HighValue As Double)
    If mtTask.Status > 0 Then Exit Sub
    With mtFind
        Select Case Range
            Case ffrAccessed
                .Accessed.High = HighValue
                .Accessed.Low = LowValue
            Case ffrModified
                .Modified.High = HighValue
                .Modified.Low = LowValue
            Case ffrCreated
                .Created.High = HighValue
                .Created.Low = LowValue
            Case ffrSize
                .Size.High = HighValue
                .Size.Low = LowValue
        End Select
    End With
End Sub

Public Sub GetRange(ByVal Range As eFindFileRange, ByRef LowValue As Double, ByRef HighValue As Double)
    With mtFind
        Select Case Range
            Case ffrAccessed
                HighValue = .Accessed.High
                LowValue = .Accessed.Low
            Case ffrModified
                HighValue = .Modified.High
                LowValue = .Modified.Low
            Case ffrCreated
                HighValue = .Created.High
                LowValue = .Created.Low
            Case ffrSize
                HighValue = .Size.High
                LowValue = .Size.Low
        End Select
    End With
End Sub

Private Sub DoFind()
    Dim lsTemp() As String
    Dim lsFileName As String
    Dim i As Long
    On Error Resume Next
    mtTask.Canceled = False
    mtTask.Status = mtTask.Status - (mtTask.Status Mod ftsCanceled) + ftsJustStarting
    mtTask.Parent.Notify mtTask.Tag
    mtTask.Status = mtTask.Status - (mtTask.Status Mod ftsCanceled) + ftsCheckingIn
    If Not mtTask.Canceled Then
        With mtFind
            .TotalFileSize = 0
            .FileCount = 0
            .DirCount = 0

            If moSearchNames Is Nothing Then
                lsTemp = Split(.Path, ";")
                For i = 0 To UBound(lsTemp)
                    FindFiles lsTemp(i)
                    If mtTask.Canceled Then Exit For
                Next
            Else
                Dim ltFound As Win32FoundData

                Do While moSearchNames.Count
                    lsFileName = moSearchNames(1)
                    moSearchNames.Remove 1
                    If FindSpecific(lsFileName, mtWin32) Then
                        If ValidateFind(mtWin32, mtFind) Then
                            If ValidateContainedText(lsFileName) Then
                                '.TotalFileSize = .TotalFileSize + MakeQWord(mtWin32.nFileSizeLow, mtWin32.nFileSizeHigh)
                                FindToFriendlyType PathGetParentFolder(lsFileName), mtWin32, ltFound
                                If mtTask.Files.pAddFile(ltFound) Then .FileCount = .FileCount + 1
                                If Not PulseProgress Then Exit Do
                            End If
                        End If
                    End If
                    If mtTask.Canceled Then Exit Do
                Loop
                Set moSearchNames = Nothing
            End If
        End With
    End If
    mtTask.Status = mtTask.Status - (mtTask.Status Mod ftsCanceled) + ftsFinishing
    PulseProgress
    Clear
End Sub

Private Sub FindFiles(ByVal psPath As String)
    On Error Resume Next
    Dim liDirs As Long, i As Long, hSearch As Long
    Dim lsFileName As String, lsDirNames() As String
    'Dim lbCountedDir As Boolean
    'Dim liSize As Long
    Dim ltFound As Win32FoundData

    If mtTask.Canceled Then Exit Sub
    With mtFind
        PathAddBackslash psPath
        liDirs = 0
        ReDim lsDirNames(0 To liDirs)
        hSearch = FindFirstFile(psPath & "*", mtWin32)
        If hSearch <> INVALID_HANDLE_VALUE Then
            Do
                lsFileName = mtWin32.cFileName
                StripNulls lsFileName
                If StrComp(lsFileName, ".") <> 0 And StrComp(lsFileName, "..") <> 0 Then
                    If mtWin32.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                        If .Recurse Then
                            lsDirNames(liDirs) = lsFileName
                            liDirs = liDirs + 1
                            ReDim Preserve lsDirNames(0 To liDirs)
                        End If
                    Else
                        lsFileName = psPath & lsFileName
                        If ValidateFind(mtWin32, mtFind) Then
                            If ValidateContainedText(lsFileName) Then
                                'liSize = MakeQWord(mtWin32.nFileSizeLow, mtWin32.nFileSizeHigh)
                                '.TotalFileSize = .TotalFileSize + liSize
                                '.FileCount = .FileCount + 1
                                
                                'If Not lbCountedDir Then
                                    'lbCountedDir = True
                                    '.DirCount = .DirCount + 1
                                'End If
                                .FileCount = .FileCount + 1
                                FindToFriendlyType psPath, mtWin32, ltFound
                                mtTask.Files.pAddFile ltFound
                                If Not PulseProgress Then Exit Do
                            End If
                        End If
                    End If
                End If
                DoEvents
                If mtTask.Canceled Then Exit Do
            Loop While FindNextFile(hSearch, mtWin32)
            FindClose hSearch
        End If

        If liDirs > 0 Then
            Dim lsPath As String
            lsPath = psPath
            For i = 0 To liDirs - 1
                lsPath = psPath & lsDirNames(i) & "\"
                FindFiles lsPath
                If mtTask.Canceled Then Exit For
            Next i
        End If
    End With
End Sub

Private Function ValidateContainedText(ByVal psPath As String) As Boolean
    If Len(msText) = 0 Then
        ValidateContainedText = True
        Exit Function
    End If
    On Error Resume Next
    Dim loFile As cFileIO
    Dim lbFirstTime As Boolean
    Dim lsWide As String
    miDoubleTextLen = Len(msText) * 2
    lbFirstTime = True
    Set loFile = New cFileIO
    With loFile
        .FileAccess = GENERIC_READ
        .FileCreation = OPEN_EXISTING
        .FileShare = FILE_SHARE_READ
        If .OpenFile(psPath) Then
            Do While Not .EOF
                If lbFirstTime Then
                    lbFirstTime = False
                    .GetBytes myMainArray
                Else
                    .GetBytes myTempArray
                    AppendBytes
                End If
                'Debug.Assert Not .EOF
                lsWide = StrConv(myMainArray, vbUnicode)
                Select Case True
                    Case InStr(1, myMainArray, msText, vbTextCompare) > 0
                        ValidateContainedText = True
                        Exit Do
                    Case InStr(1, lsWide, msText, vbTextCompare) > 0
                        ValidateContainedText = True
                        Exit Do
                    'Case InStr(1, StrConv(lsWide, vbFromUnicode), msText, vbTextCompare) > 0
                        'ValidateContainedText = True
                        'Exit Do
                    'Case InStr(1, myMainArray, StrConv(msUcaseText, vbFromUnicode), vbTextCompare) > 0
                        'ValidateContainedText = True
                        'Exit Do
                End Select
            Loop
            .CloseFile
        End If
    End With
End Function

Private Sub AppendBytes()
    Dim lyArray() As Byte
    Dim liCount As Long
    liCount = UBound(myTempArray) - LBound(myTempArray) + 1
    ReDim lyArray(0 To liCount + miDoubleTextLen - 1)
    CopyMemory lyArray(0), myMainArray(UBound(myMainArray) - miDoubleTextLen + 1), miDoubleTextLen
    CopyMemory lyArray(miDoubleTextLen), myTempArray(0), liCount
    myMainArray = lyArray
End Sub


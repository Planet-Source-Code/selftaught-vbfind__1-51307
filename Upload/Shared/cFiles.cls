VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eFileCombineLogic
    fclUnion
    fclIntersection
    fclExclusion
    fclNegation
End Enum

Private moColl As Collection

Public Sub Clear()
    Set moColl = New Collection
End Sub

Public Function AddFile(ByVal psFile As String) As Boolean
    On Error Resume Next
    Dim loFile As cFile
    Set loFile = New cFile
    loFile.FullPath = psFile
    If Len(loFile.Filename) > 0 Then
        moColl.Add loFile, loFile.FullPath
        AddFile = True
    End If
End Function

Public Sub AddFolder(ByVal psFolder As String, Optional ByVal pbRecurse As Boolean, Optional ByVal psPattern As String)
    Dim ltData As WIN32_FIND_DATA
    Dim ltFriendly As Win32FoundData
    Dim hSearch As Long
    Dim lsFile As String
    Dim lsOriginalPattern As String
    If FolderExists(psFolder) Then
        hSearch = -1
        PathAddBackslash psFolder
        If Len(psPattern) = 0 Then psPattern = "*"
        lsOriginalPattern = psFolder & psPattern
        If FindSpecific(lsOriginalPattern, ltData, hSearch) Then
            Do
                lsFile = ltData.cFileName
                StripNulls lsFile
                
                Select Case True
                    Case StrComp(lsFile, ".") = 0 Or StrComp(lsFile, "..") = 0
                    Case ltData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY
                        If pbRecurse Then AddFolder psFolder & lsFile, pbRecurse, psPattern
                    Case Else
                        If PathMatchPattern(psFolder & lsFile, lsOriginalPattern) Then
                            FindToFriendlyType psFolder, ltData, ltFriendly
                            pAddFile ltFriendly
                        End If
                End Select
            Loop While FindNext(hSearch, ltData)
            FindClose hSearch
        End If
        
    End If
    
End Sub

Friend Function pAddFile(ptFound As Win32FoundData) As Boolean
    On Error Resume Next
    Dim loFile As cFile
    Set loFile = New cFile
    loFile.Init ptFound
    Err.Clear
    moColl.Add loFile, loFile.FullPath
    pAddFile = Err.Number = 0
End Function

Public Property Get Item(ByVal Index As Variant) As cFile
Attribute Item.VB_UserMemId = 0
    Set Item = moColl.Item(Index)
End Property

Public Property Get Count() As Long
    Count = moColl.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = moColl.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set moColl = New Collection
End Sub

Private Sub Class_Terminate()
    Set moColl = Nothing
End Sub

Public Function GetNewEnum() As IUnknown
    Set GetNewEnum = moColl.[_NewEnum]
End Function

Public Function AddObject(ByVal File As cFile) As Boolean
    On Error Resume Next
    moColl.Add File, File.FullPath
    AddObject = Err.Number = 0
End Function

Public Sub RemoveFiles(ByVal psPattern As String)
    On Error Resume Next
    Dim i As Long
    Dim loFile As cFile
    i = 1
    Do While i <= moColl.Count
        Set loFile = moColl.Item(i)
        If PathMatchPattern(loFile.FullPath, psPattern) Then moColl.Remove i Else i = i + 1
    Loop
End Sub

Public Function GetFileNames() As Collection
    Set GetFileNames = New Collection
    Dim loTemp As cFile
    Dim lsString As String
    For Each loTemp In moColl
        lsString = loTemp.FullPath
        GetFileNames.Add lsString, lsString
    Next
End Function

Public Function FileSaveLoad(ByVal Filename As String, ByVal Save As Boolean) As Boolean
    On Error GoTo errout
    
    Dim loFile As cFileIO
    Set loFile = New cFileIO
    
    With loFile
        .FileShare = FILE_SHARE_READ
        .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
        
        If Save Then
            Dim loEachFile As cFile
            Dim loString As cStringBuilder
            Set loString = New cStringBuilder
            
            .FileAccess = GENERIC_WRITE
            .FileCreation = CREATE_ALWAYS

            If .OpenFile(Filename) Then
                For Each loEachFile In moColl
                    loString.Append loEachFile.FullPath & vbNullChar
                Next
                Filename = loString.ToString
                .AppendLong Len(Filename)
                CompressString Filename
                .AppendString Filename
                FileSaveLoad = True
            End If
        Else
            Dim liLen As Long
            Dim lsTemp() As String
            
            Clear
            .FileAccess = GENERIC_READ
            .FileCreation = OPEN_EXISTING
            
            If .OpenFile(Filename) Then
                .GetLong liLen
                .GetString Filename
                DecompressString Filename, liLen
                lsTemp = Split(Filename, vbNullChar)
                For liLen = LBound(lsTemp) To UBound(lsTemp)
                    AddFile lsTemp(liLen)
                Next
                FileSaveLoad = True
            End If
        End If
errout:
    End With
End Function

Public Function KeyExists(ByVal psPath As String) As Boolean
    On Error Resume Next
    IsObject moColl(psPath)
    KeyExists = Err.Number = 0
End Function

Public Function Combine(ByVal WithFiles As cFiles, Optional ByVal Logic As eFileCombineLogic = fclUnion) As cFiles
    Set Combine = New cFiles
    Dim loFile As cFile
    Dim lbVal As Boolean
    On Error Resume Next

    For Each loFile In moColl
        With loFile
            Select Case Logic
                Case fclIntersection
                    lbVal = WithFiles.KeyExists(.FullPath)
                Case fclExclusion, fclNegation
                    lbVal = Not WithFiles.KeyExists(.FullPath)
                Case fclUnion
                    lbVal = True
            End Select
            If lbVal Then Combine.AddObject loFile
        End With
    Next
    
    If Logic = fclExclusion Or Logic = fclUnion Then
        For Each loFile In WithFiles
            If Logic = fclUnion Then lbVal = True Else lbVal = Not KeyExists(loFile.FullPath)
            If lbVal Then Combine.AddObject loFile
        Next
    End If
End Function

Public Function TotalSize() As Double
    On Error Resume Next
    Dim loFile As cFile
    For Each loFile In moColl
        TotalSize = TotalSize + loFile.Size
    Next
End Function

Public Function CommonRoot() As String
    On Error Resume Next
    Dim loFile As cFile
    CommonRoot = moColl(1).FilePath
    For Each loFile In moColl
        Do Until PathMatchPrefix(loFile.FullPath, CommonRoot)
            CommonRoot = PathGetParentFolder(CommonRoot)
            If Len(CommonRoot) = 0 Then Exit Function
        Loop
    Next
End Function

Public Function Clone(Optional ByVal poFileNames As Collection) As cFiles
    Set Clone = New cFiles
    On Error Resume Next
    Dim loFile As cFile
    If poFileNames Is Nothing Then
        For Each loFile In moColl
            Clone.AddObject loFile
        Next
    Else
        Dim lvTemp
        For Each lvTemp In poFileNames
            Clone.AddObject moColl(lvTemp)
            Clone.GetNewEnum
        Next
    End If
End Function

Public Function GetFolders() As Collection
    On Error Resume Next
    Dim loFile As cFile
    Dim lsString As String
    Set GetFolders = New Collection
    For Each loFile In moColl
        lsString = loFile.FilePath
        GetFolders.Add lsString, lsString
    Next
End Function

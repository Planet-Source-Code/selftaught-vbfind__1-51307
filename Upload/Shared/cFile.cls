VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Enum eFileAttribs
    faArchive = &H20
    faDirectory = &H10
    faHidden = &H2
    faNormal = &H80
    faReadOnly = &H1
    faSystem = &H4
    faTemporary = &H100
End Enum

Private Enum eShellGetFileInfoFlags
    SHGFI_ATTRIBUTES = &H800                   '  get attributes
    SHGFI_DISPLAYNAME = &H200                  '  get display name
    SHGFI_EXETYPE = &H2000                     '  return exe type
    SHGFI_ICON = &H100                         '  get icon
    SHGFI_ICONLOCATION = &H1000                '  get icon location
    SHGFI_LARGEICON = &H0                      '  get large icon
    SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
    SHGFI_OPENICON = &H2                       '  get open icon
    SHGFI_PIDL = &H8                           '  pszPath is a pidl
    SHGFI_SELECTED = &H10000                   '  show icon in selected state
    SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
    SHGFI_SMALLICON = &H1                      '  get small icon
    SHGFI_SYSICONINDEX = &H4000                '  get system icon index
    SHGFI_TYPENAME = &H400                     '  get type name
    SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute
End Enum

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As eShellGetFileInfoFlags) As Long

Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long          '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80          '  out: type name
End Type

Private msPath As String
Private msName As String
Private mdblSize As Double
Private mdModified As Date
Private mdAccessed As Date
Private mdCreated As Date
'Private msTypeName As String
Private miAttributes As eFileAttribs

Public Property Get Modified() As Date
    Modified = mdModified
End Property

Public Property Get Accessed() As Date
    Accessed = mdAccessed
End Property

Public Property Get Created() As Date
    Created = mdCreated
End Property

Public Property Get Attributes() As eFileAttribs
    Attributes = miAttributes
End Property

Public Property Let Attributes(piVal As eFileAttribs)
    FileSetAttributes msName, piVal
    Attributes = FileGetAttributes(msName)
End Property

Public Property Get Size() As Double
    Size = mdblSize
End Property

Public Property Get FileName() As String
Attribute FileName.VB_UserMemId = 0
    FileName = msName
End Property

Public Property Get FullPath() As String
    FullPath = msPath & msName
End Property

Public Property Get FilePath() As String
    FilePath = msPath
End Property

Public Property Let FullPath(psFile As String)
    Dim ltData As WIN32_FIND_DATA
    Dim ltFriendly As Win32FoundData
    If FileExists(psFile) Then
        If FindSpecific(psFile, ltData) Then
            FindToFriendlyType PathGetParentFolder(psFile), ltData, ltFriendly
            Init ltFriendly
        End If
    End If
End Property

Friend Sub Init(ptWin32 As Win32FoundData)
    Dim Info As SHFILEINFO
    With ptWin32
        miAttributes = .Attributes
        mdAccessed = .Accessed
        mdModified = .Modified
        mdCreated = .Created
        msPath = .Path
        PathAddBackslash msPath
        mdblSize = .Size
        msName = .Name
    End With
        
    'SHGetFileInfo psPath, 0, Info, Len(Info), SHGFI_TYPENAME 'Or SHGFI_DISPLAYNAME
    'With Info
        'msName = .szDisplayName
        'StripNulls msName
        
        'msType = .szTypeName
        'StripNulls msType
    'End With
End Sub

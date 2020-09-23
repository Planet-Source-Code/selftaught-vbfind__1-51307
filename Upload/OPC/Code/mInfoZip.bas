Attribute VB_Name = "mInfoZip"
Option Explicit

'This Code was adapted from a post by Doug Gaede to www.pscode.com (clsCryptoAPIAndCompression)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlib.dll" (dest As Any, destLen As Any, Src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, Src As Any, ByVal srcLen As Long) As Long

Public Function CompressByteArray(ByteArray() As Byte, _
                         Optional CompressionLevel As Integer = 9) _
                As Long
    
    Dim lyArray() As Byte
    Dim liBufferSize As Long

    liBufferSize = UBound(ByteArray) * 1.01 + 13
    ReDim lyArray(liBufferSize - 1)
    
    CompressByteArray = compress2(lyArray(0), _
                                  liBufferSize, _
                                  ByteArray(0), _
                                  UBound(ByteArray) + 1, _
                                  CompressionLevel)
    
    ReDim Preserve ByteArray(liBufferSize - 1)
    CopyMemory ByteArray(0), lyArray(0), liBufferSize

End Function

Public Function CompressString(Text As String, _
                Optional ByVal CompressionLevel As Integer = 9) _
                As Long
                
    Dim liLen As Long
    Dim lsTemp As String
    
    lsTemp = String$((Len(Text) * 1.01) + 12, 0)
    liLen = Len(lsTemp)
    
    CompressString = compress2(ByVal lsTemp, _
                                     liLen, _
                               ByVal Text, _
                                     Len(Text), _
                                     CompressionLevel)
    
    Text = Left$(lsTemp, liLen)

End Function

Public Function DecompressByteArray(ByteArray() As Byte, _
                     Optional ByVal OriginalSize As Long = 65536) _
                As Long
    Dim liBufferSize As Long
    Dim lyArray() As Byte

    liBufferSize = OriginalSize * 1.01 + 13
    ReDim lyArray(liBufferSize - 1)
    
    DecompressByteArray = uncompress(lyArray(0), _
                                     liBufferSize, _
                                     ByteArray(0), _
                                     UBound(ByteArray) + 1)
    
    ReDim Preserve ByteArray(liBufferSize - 1)
    CopyMemory ByteArray(0), lyArray(0), liBufferSize

End Function

Public Function DecompressString(Text As String, _
                  Optional ByVal OriginalSize As Long = 65536) _
                As Long

    Dim liLen As Long
    Dim lsTemp As String
    
    lsTemp = String(OriginalSize * 1.01 + 12, 0)
    liLen = Len(lsTemp)
    
    DecompressString = uncompress(ByVal lsTemp, liLen, ByVal Text, Len(Text))
    
    Text = Left$(lsTemp, liLen)

End Function

'Public Function CompressFile(FilePathIn As String, _
'                             FilePathOut As String, _
'              Optional ByVal CompressionLevel As Integer = 9) _
'                As Long
'    On Error GoTo Oops
'
'    Dim liFileNum As Integer
'    Dim lyArray() As Byte
'    Dim liFileLen As Long
'
'
'    'If problem Then check to see if FileLen is returning the correct size....
'    liFileLen = FileGetLen(FilePathIn)
'    ReDim lyArray(liFileLen - 1)
'
'    liFileNum = FreeFile
'
'    Open FilePathIn For Binary Access Read As #liFileNum
'    'ReDim lyArray(0 To LOF(liFileNum) - 1)
'    Get #liFileNum, , lyArray()
'    Close #liFileNum
'
'    CompressFile = CompressByteArray(lyArray(), CompressionLevel)
'
'    On Error Resume Next
'    FileDelete FilePathOut, True
'    On Error GoTo 0
'
'    liFileNum = FreeFile
'
'    Open FilePathOut For Binary Access Write As #liFileNum
'    Put #liFileNum, , liFileLen
'    Put #liFileNum, , lyArray()
'    Close #liFileNum
'
'    Exit Function
'Oops:
'    If liFileNum > 0 Then
'        On Error Resume Next
'        Close #liFileNum
'    End If
'End Function
'
'Public Function DecompressFile(FilePathIn As String, _
'                               FilePathOut As String) _
'                As Long
'    On Error Resume Next
'    Dim liFileNum As Integer
'    Dim lyArray() As Byte
'    Dim liFileLen As Long
'
'    ReDim lyArray(FileLen(FilePathIn) - 1)
'
'    liFileNum = FreeFile
'
'    Open FilePathIn For Binary Access Read As #liFileNum
'    Get #liFileNum, , liFileLen
'    Get #liFileNum, , lyArray()
'    Close #liFileNum
'
'    DecompressFile = DecompressByteArray(lyArray(), liFileLen)
'
'    On Error Resume Next
'    FileDelete FilePathOut, True
'    On Error GoTo 0
'
'    liFileNum = FreeFile
'
'    Open FilePathOut For Binary Access Write As #liFileNum
'    Put #liFileNum, , lyArray()
'    Close #liFileNum
'
'End Function

'Sub test()
'    Dim lsCompressed As String
'    Dim lsUncompressed As String
'    Dim lyArrayC() As Byte
'    Dim lyArrayUnC() As Byte
'
'    lsUncompressed = Clipboard.GetText
'    lsCompressed = lsUncompressed
'    lyArrayC = StrConv(lsUncompressed, vbFromUnicode)
'    lyArrayUnC = lyArrayC
'
'    CompressString lsCompressed
'    Debug.Print 1 - Len(lsCompressed) / Len(lsUncompressed)
'
'    DecompressString lsCompressed
'    Debug.Print lsCompressed = lsUncompressed
'
'    CompressByteArray lyArrayC
'    Debug.Print 1 - UBound(lyArrayC) / UBound(lyArrayUnC)
'
'    DecompressByteArray lyArrayC
'    Debug.Print StrConv(lyArrayUnC, vbUnicode) = StrConv(lyArrayC, vbUnicode)
'
'End Sub

Attribute VB_Name = "mCryptoAPI"
Option Explicit

Public gsWord As String
Public gsSalt As String

Private moCryptoAPI As cCryptoAPI

Public Property Get CryptoSALT() As String
    If moCryptoAPI Is Nothing Then Set moCryptoAPI = New cCryptoAPI
    CryptoSALT = moCryptoAPI.SALT
End Property

Public Function EncryptString(Text As String, _
                              Password As String, _
                     Optional SALT As String) _
                As String
                
    If moCryptoAPI Is Nothing Then Set moCryptoAPI = New cCryptoAPI
    EncryptString = moCryptoAPI.EncryptString(Text, Password, SALT)
    
End Function

Public Function DecryptString(Text As String, _
                              Password As String, _
                     Optional SALT As String) _
                As String
    
    If moCryptoAPI Is Nothing Then Set moCryptoAPI = New cCryptoAPI
    DecryptString = moCryptoAPI.DecryptString(Text, Password, SALT)

End Function

Public Function EncryptByteArray(ByteArray() As Byte, _
                                 Password As String, _
                        Optional SALT As String) _
                As Byte()
                
    If moCryptoAPI Is Nothing Then Set moCryptoAPI = New cCryptoAPI
    EncryptByteArray = moCryptoAPI.EncryptByteArray(ByteArray, Password, SALT)
    
End Function

Public Function DecryptByteArray(ByteArray() As Byte, _
                                 Password As String, _
                        Optional SALT As String) _
                As Byte()
    
    If moCryptoAPI Is Nothing Then Set moCryptoAPI = New cCryptoAPI
    DecryptByteArray = moCryptoAPI.DecryptByteArray(ByteArray, Password, SALT)
    
End Function

'Public Sub EncryptFile(FilePathIn As String, _
'                       FilePathOut As String, _
'                       Password As String)
'
'    If moCryptoAPI Is Nothing Then Set moCryptoAPI = New cCryptoAPI
'    moCryptoAPI.EncryptFile FilePathIn, FilePathOut, Password
'
'End Sub
'
'Public Sub DecryptFile(FilePathIn As String, _
'                       FilePathOut As String, _
'                       Password As String)
'
'    If moCryptoAPI Is Nothing Then Set moCryptoAPI = New cCryptoAPI
'    moCryptoAPI.DecryptFile FilePathIn, FilePathOut, Password
'
'End Sub
'
'
''Sub test()
'    Dim lsEncrypted As String
'    Dim lsUnEncrypted As String
'    Dim lyArrayC() As Byte
'    Dim lyArrayUnC() As Byte
'
'    Const pass = "Test Password"
'
'    lsUnEncrypted = Clipboard.GetText
'    lsEncrypted = lsUnEncrypted
'    lyArrayC = StrConv(lsUnEncrypted, vbFromUnicode)
'    lyArrayUnC = lyArrayC
'
'    lsEncrypted = EncryptString(lsUnEncrypted, pass)
'    'Stop
'
'    Debug.Print lsUnEncrypted = DecryptString(lsEncrypted, pass)
'
'    lyArrayC = EncryptByteArray(lyArrayUnC, pass)
'    'Stop
'
'    lyArrayC = DecryptByteArray(lyArrayC, pass)
'    Debug.Print StrConv(lyArrayUnC, vbUnicode) = StrConv(lyArrayC, vbUnicode)
'
'End Sub

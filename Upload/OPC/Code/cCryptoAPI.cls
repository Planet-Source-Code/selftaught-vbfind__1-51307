VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCryptoAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This Code was taken from a post by Doug Gaede to www.pscode.com (clsCryptoAPIAndCompression)

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetKeyParam Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
Private Declare Function CryptSetKeyParam Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal dwParam As Long, ByVal pbData As String, ByVal dwFlags As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hSessionKey As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef hSessionKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long

Private Const KP_SALT As Long = 2
Private Const KEY_CONTAINER As String = "cCryptoAPI" & vbNullChar 'this is usually set to your program's name
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0" & vbNullChar
Private Const USE_THIS_CSP As String = MS_ENHANCED_PROV 'can set this to other providers.  See Microsoft's website for full listing
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const ALG_CLASS_HASH As Long = 32768
Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_SID_SHA As Long = 4
Private Const CALG_SHA As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)
Private Const HP_HASHVAL As Long = 2
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Private Const ALG_TYPE_STREAM As Long = 2048
Private Const ALG_SID_RC4 As Long = 1
Private Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)

Dim hCryptProv As Long 'the handle to the CSP
Dim hSessionKey As Long 'the handle to the current session key
Dim strSALT As String 'the session key SALT

Private msSALT As String

Public Property Get SALT() As String
    SALT = msSALT
End Property

Public Function EncryptString(Text As String, _
                              Password As String, _
                     Optional SALT As String) _
                As String
    EncryptString = EncryptDecrypt(Text, Password, True, SALT)
End Function

Public Function DecryptString(Text As String, _
                              Password As String, _
                     Optional SALT As String) _
                              As String
    DecryptString = EncryptDecrypt(Text, Password, False, SALT)
End Function

Public Function EncryptByteArray(pyArray() As Byte, _
                                 Password As String, _
                        Optional SALT As String) _
                As Byte()

Dim lsBefore As String
Dim lsAfter As String

lsBefore = StrConv(pyArray(), vbUnicode)
lsAfter = EncryptDecrypt(lsBefore, Password, True, SALT)
EncryptByteArray = StrConv(lsAfter, vbFromUnicode)

End Function

Public Function DecryptByteArray(ByRef pyArray() As Byte, _
                                 Password As String, _
                        Optional SALT As String) _
                As Byte()
'single key byte array encryption
Dim lsBefore As String
Dim lsAfter As String

lsBefore = StrConv(pyArray(), vbUnicode) 'convert to string
lsAfter = EncryptDecrypt(lsBefore, Password, False, SALT) 'return the encrypted data
DecryptByteArray = StrConv(lsAfter, vbFromUnicode) 'convert to byte array

End Function

Private Function EncryptDecrypt(ByVal Text As String, _
                                ByVal Key As String, _
                                ByVal Encrypt As Boolean, _
                                Optional psSalt As String) _
                 As String
'the code in this function encrypts/decrypts the data using a single key
Dim lngLength As Long
Dim lngSALTLen As Long
Dim lngReturnValue As Long

SessionKeyFromPassword Key  'get a session key derived from the password

'Set a random SALT.  Always 11 bytes long for the Base CSP, but this code gets the allowed length the correct way
'since other CSPs can have longer lengths.  This shows you how.
If Encrypt Then 'only get a new SALT during encryption
    
    lngReturnValue = CryptGetKeyParam(hSessionKey, KP_SALT, vbNull, lngSALTLen, 0) 'get the allowed length of the SALT
    'lngReturnValue above is always 0 when you pass in the vbNull parameter, so no reason to check for an error.
    'strSALT = String(lngSALTLen + 1, vbNullChar) 'initialize the buffer
    'lngReturnValue = CryptGenRandom(hCryptProv, lngSALTLen, strSALT) 'generate a random SALT.
    'To set your own you can use the following line instead of the previous line:
    
    'strSALT = "12345678901" & vbnullchar
    'If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not generate a random SALT"
End If
'
If Len(psSalt) > 10 Then strSALT = psSalt Else strSALT = msSALT
lngReturnValue = CryptSetKeyParam(hSessionKey, KP_SALT, strSALT, 0)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError
'Note that the same SALT must be used during encryption and decryption.
'A SALT causes the encrypted output to be different even when the input plaintext and password are the same,
'thus the SALT should be different every single time you encrypt a file or string.
'The SALT should be sent as plaintext along with the encrypted file.  This does not make the
'encryption any weaker.

'uncomment this code if you want to verify the SALT value set above
'strSALT = String(12, vbNullChar) 'clear the string
'lngSALTLen = 0 'clear the variable
'lngReturnValue = CryptGetKeyParam(hSessionKey, KP_SALT, vbNull, lngSALTLen, 0) 'get the length of the SALT
'lngReturnValue = CryptGetKeyParam(hSessionKey, KP_SALT, strSALT, lngSALTLen, 0) 'get the SALT
'If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not get the SALT."
'MsgBox "get " & lngSALTLen & " " & strSALT

'Encrypt or decrypt depending on the Encrypt parameter
lngLength = Len(Text)
If Encrypt Then
    lngReturnValue = CryptEncrypt(hSessionKey, 0, 1, 0, Text, lngLength, lngLength)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError
Else
    lngReturnValue = CryptDecrypt(hSessionKey, 0, 1, 0, Text, lngLength)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError
End If

'return the encrypted/decrypted data and chop off extra padding
EncryptDecrypt = Left$(Text, lngLength)

'Destroy the session key
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey

End Function

Private Sub SessionStart()

    Randomize Timer
    Dim i As Long
    Dim lyBytes(0 To 10) As Byte
    For i = 0 To 10
        lyBytes(i) = Int(255 * Rnd + 1)
    Next
    msSALT = StrConv(lyBytes, vbUnicode) & vbNullChar


'this could be placed at the beginning of EncryptDecrypt, but if you are doing
'multiple encryptions/decryptions, calling this once speeds things up
Dim lngReturnValue As Long

'Get handle to CSP
lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, USE_THIS_CSP, PROV_RSA_FULL, CRYPT_NEWKEYSET) 'try to make a new key container
If lngReturnValue = 0 Then
    lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, USE_THIS_CSP, PROV_RSA_FULL, 0) 'try to get a handle to a key container that already exists, and if it fails...
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError
End If

End Sub

Private Sub SessionEnd()

'Release any session key handle
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey

'Release any key pair handle
'If hKeyPair <> 0 Then CryptDestroyKey hKeyPair

'Release provider handle
If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0

End Sub

Private Sub SessionKeyFromPassword(ByVal Key As String)
'This sub takes a string key as input and sets the module-level
'hSessionKey variable to a new session key handle.
'This sub is used by EncryptDecrypt, Export_KeyPair and Import_KeyPair.
Dim lngParams As Long
Dim lngReturnValue As Long
Dim strHash As String
Dim lngHashLen As Long
Dim hHash As Long 'the handle to the hash object

'Create a hash object to calculate a session
'key from the Password (instead of encrypting
'with the actual key)
lngReturnValue = CryptCreateHash(hCryptProv, CALG_SHA, 0, 0, hHash)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError
'can use CALG_MD5 to get a 128-bit hash.  CALG_SHA returns a 160-bit hash (more secure).

'Hash the Password
lngReturnValue = CryptHashData(hHash, Key, Len(Key), 0)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError
'A hash is a 'fingerprint' of any string.
'Hashes are extremely useful for determining whether a
'transmission or file has been altered.  This code can use
'one of two algorithms (see note above).  No matter what the
'length of input data, the hash will be a fixed length and
'will be unique for that string of data.  The same hash is produced for
'the same input data every time.  This is useful here to
'produce a fixed-length, unique password for any length password entered.

'Get the actual hash value
lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0) 'get the hash length
strHash = String(lngHashLen + 1, vbNullChar)
lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0) 'get the hash value
If lngReturnValue = 0 Then Err.Raise Err.LastDllError

'Set certain values to add more flexibility and security.
'Make the key exportable. (I don't export the key in this sample code)
lngParams = CRYPT_EXPORTABLE 'use this when you generate your own SALT, which is recommended (see 8 lines below)
'Make the key exportable and add a system-generated SALT.
'use this line of code instead of the one above if you want the API to set the SALT...
'but the SALT is the same every time so this shouldn't be used!
'Note that I generate my own random SALT below.
'lngParams = CRYPT_EXPORTABLE Or CRYPT_CREATE_SALT

'release old session key handle if one exists
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey

'Derive a session key from the hash object
lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, lngParams, hSessionKey)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError

'Destroy the hash object
If hHash <> 0 Then CryptDestroyHash hHash

End Sub

Private Sub Class_Initialize()
    SessionStart
End Sub

Private Sub Class_Terminate()
    SessionEnd
End Sub









'Public Function EncryptFile(ByVal SourceFile As String, ByVal DestFile As String, ByVal Password As String) As Boolean
'Dim intNextFreeFile As Integer
'Dim arrByteArray() As Byte
'Dim x As Long
'Dim y As Long
'Dim z As Long
'
''reset the SALT value
'strSALT = ""
'
''Open the source file and read the content
''into a arrByteArray to pass onto encryption
'intNextFreeFile = FreeFile
'Open SourceFile For Binary As #intNextFreeFile
'ReDim arrByteArray(0 To LOF(intNextFreeFile) - 1)
'Get #intNextFreeFile, , arrByteArray()
'Close #intNextFreeFile
'
''Encrypt the arrByteArray
'arrByteArray = EncryptByteArray(arrByteArray(), Password)
'
''If the destination file already exists we need
''to delete it since opening it for binary use
''will preserve it if it already exists
'On Error Resume Next
'FileDelete DestFile, True
'On Error GoTo 0
'
''Store the encrypted data in the destination file
'intNextFreeFile = FreeFile
''first tack on strSALT to arrByteArray
''This is a kludge because while you can Put the the string
''before arrByteArray, for some reason you can not then
''Get the string when reading it back from the file, so I just
''tacked it on to the array.
'z = UBound(arrByteArray) + 1
'y = 1
'ReDim Preserve arrByteArray(0 To (UBound(arrByteArray) + Len(strSALT) + 1))
'For x = z To z + Len(strSALT) - 1
'    arrByteArray(x) = Asc(Mid(strSALT, y, 1))
'    y = y + 1
'Next x
'arrByteArray(UBound(arrByteArray)) = Len(strSALT) 'length of salt is last character
'
'Open DestFile For Binary As #intNextFreeFile
'Put #intNextFreeFile, , arrByteArray()
'Close #intNextFreeFile
'
'End Function
'
'Public Sub DecryptFile(ByVal SourceFile As String, ByVal DestFile As String, ByVal Password As String)
''single key file decryption
'Dim intNextFreeFile As Integer
'Dim arrByteArray() As Byte
'Dim x As Long
'Dim y As Long
'
''reset the SALT value
'strSALT = ""
'
''Open the source file and read the content
''into arrByteArray to decrypt
'intNextFreeFile = FreeFile
'Open SourceFile For Binary As #intNextFreeFile
'ReDim arrByteArray(0 To LOF(intNextFreeFile) - 1)
'Get #intNextFreeFile, , arrByteArray()
'Close #intNextFreeFile
'
''extract the SALT
''This is a kludge because while you can Put the the string
''before arrByteArray, for some reason you can not then
''Get the string when reading it back from the file, so I just
''tacked it on to the array.
'y = arrByteArray(UBound(arrByteArray))
'For x = UBound(arrByteArray) - y To UBound(arrByteArray) - 1
'    strSALT = strSALT & Chr(arrByteArray(x))
'Next x
'ReDim Preserve arrByteArray(0 To (UBound(arrByteArray) - Len(strSALT) - 1))
'
''Decrypt arrByteArray
'arrByteArray = DecryptByteArray(arrByteArray(), Password)
'
''If the destination file already exists we need
''to delete it since opening it for binary use
''will preserve it if it already exists
'On Error Resume Next
'If FileLen(DestFile) > 0 Then FileDelete DestFile, True
'On Error GoTo 0
'
''Store the decrypted data in the destination file
'intNextFreeFile = FreeFile
'Open DestFile For Binary As #intNextFreeFile
'Put #intNextFreeFile, , arrByteArray()
'Close #intNextFreeFile
'
'End Sub


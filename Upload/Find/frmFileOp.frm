VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFileTask 
   Caption         =   " "
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.Animation ani 
      Height          =   615
      Left            =   15
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1085
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   377
      FullHeight      =   41
   End
   Begin VB.PictureBox pic 
      Height          =   285
      Left            =   22
      ScaleHeight     =   225
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   4327
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lbl 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   45
      TabIndex        =   6
      Top             =   345
      Width           =   5595
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   2
      Left            =   30
      TabIndex        =   5
      Top             =   2115
      Width           =   3960
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   5595
   End
   Begin VB.Label lbl 
      Caption         =   "Preparing files....."
      Height          =   435
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   1335
      Width           =   5625
   End
End
Attribute VB_Name = "frmFileTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements iFileTaskParent

Private moProg As cProgressBar
Private moTask As iFileTask
Private mtStart As SYSTEMTIME
Private moTimeEstimates As Collection
Private mdblLastEstimate As Double

Private Sub ani_GotFocus()
    cmd.SetFocus
End Sub

Private Sub ani_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveForm hwnd
End Sub

Private Sub cmd_Click()
    On Error Resume Next
    moTask.Canceled = True
End Sub

Private Sub Form_Initialize()
    Set moProg = New cProgressBar
    With moProg
        .DrawObject = pic
        .ShowText = True
        .BackColor = vbButtonFace
        .BarColor = vbHighlight
    End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveForm hwnd
End Sub

Private Sub Form_Terminate()
    Set moProg = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not moTask Is Nothing Then
        moTask.Canceled = True
        Set moTask = Nothing
    End If
End Sub

Private Sub iFileTaskParent_Notify(Tag As Long)
    On Error Resume Next
    With moTask
        Select Case .Status Mod ftsCanceled
            Case ftsJustStarting
                GetSystemTime mtStart
            Case ftsFinishing
                On Error Resume Next
                If .Status And ftsCanceled Then Debug.Print "Canceled"
                If .Status And ftsError Then
                    MsgBoxEx "Serious Error Occurred.  Operation could not be completed.", rdCritical
                    Debug.Print "Error"
                End If
                If .Errors.Count Then
                    Dim lsMsg As String, i As Long
                    With .Errors
                        lsMsg = "Errors occurred on the following " & .Count & " files:" & vbCrLf & vbCrLf
                        For i = 1 To .Count
                            lsMsg = lsMsg & .Item(i) & vbTab
                            If i Mod 2 = 0 Then lsMsg = lsMsg & vbNewLine
                        Next
                    End With
                    MsgBoxEx lsMsg, rdCritical
                End If
                ani.Stop
                ani.Close
                Set moTask.Parent = Nothing
                Unload Me
            Case Else
                Debug.Print .CurrentFile
                lbl(1).Caption = lbl(1).Tag & vbNewLine & PathCompactPixels(.CurrentFile, hDC, ScaleX(lbl(1).Width, ScaleMode, vbPixels))
                CheckProgress
        End Select
    End With
End Sub

Private Sub CheckProgress()
    Dim ldblTotal As Double
    Dim ldblProgress As Double
    Dim ldblTemp As Double
    Dim ltTemp As SYSTEMTIME
    With moTask
        ldblTotal = .BytesTotal
        ldblProgress = .BytesProgress
    End With
    moProg.Max = ldblTotal
    moProg.Value = ldblProgress
    GetSystemTime ltTemp
    'Debug.Print SysTimeDiff(mtStart, ltTemp), ldblProgress, ldblTotal
    lbl(2).Caption = TimeEstimate(SysTimeDiff(mtStart, ltTemp) / ldblProgress)
End Sub

Private Function TimeEstimate(ByVal pdblSpeed As Double) As String
    Const Precision = 20
    If moTimeEstimates Is Nothing Then Set moTimeEstimates = New Collection
    moTimeEstimates.Add pdblSpeed
    Dim ldblAvg As Double
    Dim lvTemp
    Select Case True
        Case moTimeEstimates.Count < Precision
            TimeEstimate = "Estimating Time Remaining..."
            Exit Function
        Case moTimeEstimates.Count > Precision
            moTimeEstimates.Remove 1
    End Select
    For Each lvTemp In moTimeEstimates
        ldblAvg = ldblAvg + lvTemp
    Next
    ldblAvg = ldblAvg / Precision
    ldblAvg = Abs(moTask.BytesTotal - moTask.BytesProgress) * ldblAvg
    If ldblAvg > mdblLastEstimate And mdblLastEstimate > 0 Then ldblAvg = mdblLastEstimate
    If ldblAvg < 2000 Then ldblAvg = 2000
    If ldblAvg > 60000 Then
        TimeEstimate = Replace(Format$(ldblAvg / 60000, "0.#") & " Minute(s)", ". ", " ")
    Else
        TimeEstimate = Format$(ldblAvg / 1000, "0") & " Seconds"
    End If
    TimeEstimate = "About " & TimeEstimate & " Remaining..."
    mdblLastEstimate = ldblAvg
End Function

Private Function FilesIsAre(ByVal piCount As Long) As String
    If piCount = 1 Then FilesIsAre = "1 files is " Else FilesIsAre = piCount & " files are "
End Function

Public Function StartTask(ByVal poFiles As cFiles, ByVal piTask As Long, Optional psDest As String, Optional psRelative As String) As Boolean
    If Not moTask Is Nothing Then Exit Function
    Dim lsString As String
    Dim lsFileString As String
    Dim lsAVI As String
    Dim loBR As cFileBackupRestore
    Dim loCMD As cFileCopyMoveDelete
    Dim lsClass As String
    Select Case piTask
        Case ftZip
            lsString = FilesIsAre(poFiles.Count) & "being packed into:"
            lsFileString = "Packing File:"
            lsAVI = "encrypt"
            lsClass = cBackupRestoreClass
        Case ftUnzip
            If poFiles.Count = 1 Then _
                lsString = "1 composite file is being restored into:" _
            Else _
                lsString = poFiles.Count & " composite files are being restored into:"
            lsFileString = "Restoring File:"
            lsClass = cBackupRestoreClass
            lsString = "Decompressing Files"
            lsAVI = "decrypt"
        Case ftMove
            lsFileString = "Moving File:"
            lsClass = cCopyMoveDeleteClass
            lsString = FilesIsAre(poFiles.Count) & "being moved to:"
            lsAVI = "copy"
        Case ftEncrypt
            lsFileString = "Packing File:"
            lsString = FilesIsAre(poFiles.Count) & "being packed into:"
            lsClass = cBackupRestoreClass
            lsAVI = "encrypt"
        Case ftDecrypt
            If poFiles.Count = 1 Then _
                lsString = "1 composite file is being restored into:" _
            Else _
                lsString = poFiles.Count & " composite files are being restored into:"
            lsFileString = "Restoring File:"
            lsClass = cBackupRestoreClass
            lsAVI = "decrypt"
        Case ftCryptoZip
            lsFileString = "Packing File:"
            lsString = FilesIsAre(poFiles.Count) & "being packed into:"
            lsClass = cBackupRestoreClass
            lsAVI = "encrypt"
        Case ftCryptoUnZip
            If poFiles.Count = 1 Then _
                lsString = "1 composite file is being restored into:" _
            Else _
                lsString = poFiles.Count & " composite files are being restored into:"
            lsFileString = "Restoring File:"
            lsClass = cBackupRestoreClass
            lsAVI = "decrypt"
        Case ftCopy
            lsFileString = "Copying File:"
            lsClass = cCopyMoveDeleteClass
            lsString = FilesIsAre(poFiles.Count) & "being copied to:"
            lsAVI = "copy"
        Case ftDelete
            lsFileString = "Deleting File:"
            lsClass = cCopyMoveDeleteClass
            lsString = FilesIsAre(poFiles.Count) & "being deleted."
            lsAVI = "delete"
    End Select
    Caption = "File Progress"
    lbl(0).Caption = lsString
    lbl(1).Tag = lsFileString
    If piTask <> ftDelete Then
        Set Me.Font = lbl(3).Font
        lbl(3).Caption = PathCompactPixels(psDest, hDC, ScaleX(lbl(3).Width, ScaleMode, vbPixels))
        Set Me.Font = lbl(1).Font
    Else
        lbl(3).Caption = ""
    End If
    Set moTask = CreateObject(lsClass)
    Set moTask.Files = poFiles
    If TypeOf moTask Is cFileCopyMoveDelete Then
        Set loCMD = moTask
        loCMD.RelativeToFolder = psRelative
        loCMD.Target = psDest
        loCMD.CurrentTask = piTask
    Else
        Set loBR = moTask
        loBR.RelativeToFolder = psRelative
        loBR.Target = psDest
        loBR.CurrentTask = piTask
    End If
    Set moTask.Parent = Me
    StartTask = moTask.Start
    
    On Error Resume Next
    If StartTask Then
        Show vbModeless
        ani.Open PathBuild(App.Path, lsAVI & ".avi") '"C:\vb projects\build\"
    Else
        Set moTask = Nothing
    End If
End Function

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveForm hwnd
End Sub

Private Sub pic_GotFocus()
    On Error Resume Next
    cmd.SetFocus
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveForm hwnd
End Sub

Private Sub pic_Paint()
    moProg.Draw
End Sub

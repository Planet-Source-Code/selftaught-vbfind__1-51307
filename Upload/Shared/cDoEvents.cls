VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDoEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const QS_HOTKEY = &H80
Private Const QS_KEY = &H1
Private Const QS_MOUSEBUTTON = &H4
Private Const QS_PAINT = &H20
Private Const QS_POSTMESSAGE = &H8
Private Const QS_SENDMESSAGE = &H40
Private Const QS_TIMER = &H10
Private Const QS_MOUSEMOVE = &H2

Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Public Tag As Long


Public Function DoUserInput() As Boolean
    DoUserInput = CBool(GetQueueStatus(QS_ALLINPUT))
    If DoUserInput Then DoEvents
End Function

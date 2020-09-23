VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Notify(pvData())

Sub Notify(ParamArray pvFlags())
    Dim lvArray()
    Dim i As Long
    
    ReDim lvArray(LBound(pvFlags) To UBound(pvFlags))
    
    For i = LBound(pvFlags) To UBound(pvFlags)
        lvArray(i) = pvFlags(i)
    Next
    
    RaiseEvent Notify(lvArray)
End Sub

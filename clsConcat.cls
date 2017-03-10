VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sBuffer As String
Private lBufferLen As Long
Private lPointer As Long

Public Sub Clear()
    lPointer = 1
End Sub

Public Sub Concat(ByRef sValue As String)
    Dim lPointerNew As Long
  
    lPointerNew = lPointer + LenB(sValue)
    
    If lPointerNew > lBufferLen Then
        sBuffer = sBuffer & Space$(lPointerNew)
        lBufferLen = LenB(sBuffer)
    End If
  
    MidB$(sBuffer, lPointer) = sValue
    lPointer = lPointerNew
End Sub

Public Property Get Value() As String
    Value = LeftB$(sBuffer, lPointer - 1)
End Property

Private Sub Class_Initialize()
    Clear
End Sub

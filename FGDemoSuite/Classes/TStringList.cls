VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStringList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCol As Collection
'yes ok it's a kind primitive programmin' but it is fast and it works

Private Sub Class_Initialize()
  Clear
End Sub

Public Sub Clear()
  Set mCol = New Collection
End Sub
Public Sub Free()
  Clear
End Sub
Public Sub Add(StrVal As String)
  mCol.Add StrVal
End Sub

Public Sub SaveToFile(StrPFN As String)
Dim FNr As Integer, i As Long
Dim StrLine As String
tryE: On Error GoTo finallyE
  FNr = FreeFile
  Open StrPFN For Output As #FNr
  For i = 1 To mCol.Count
    StrLine = mCol.Item(i) & vbCrLf
    Print #FNr, StrLine
  Next
finallyE:
  Close FNr
End Sub

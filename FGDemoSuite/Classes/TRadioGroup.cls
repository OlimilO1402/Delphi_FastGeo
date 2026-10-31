VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRadioGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCol As Collection
Private mOB As OptionButton
Private nCount As Long
Public ItemIndex As Long

Private Sub Class_Initialize()
  ItemIndex = 0
  Clear
  'ReDim mOB(0 To 1)
End Sub
Public Sub NewC(ob As OptionButton) ' As OptionButton
  'mOB = ob
End Sub

Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub

Public Sub Clear()
  Set mCol = New Collection
End Sub

Public Function Items() As Collection
  Set Items = mCol
End Function

Public Sub Add(StrVal As String)
  'Load Command1(Index)
  'nCount = nCount + 1
  'Load mOB(nCount)
  'mOB(nCount).Visible = True
  'mOB(nCount).Move mOB(nCount).Left, mOB(nCount - 1).Top + 180
  'mOB(nCount).Caption = StrVal
  'Top und Left korrigieren, Visible=True, fertisch
  'ne das geht so nicht
End Sub
'Private Sub mOB_Click(Index As Integer)
'  ItemIndex = Index
'End Sub

'momentan Keine Zeit sowas zu programmieren

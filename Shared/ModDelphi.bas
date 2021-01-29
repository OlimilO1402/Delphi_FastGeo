Attribute VB_Name = "ModDelphi"
Option Explicit
'Single Precision!
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Function Assigned(mObj As Object) As Boolean
  If Not mObj Is Nothing Then Assigned = True
End Function

Public Function Finalize(mObj As Object) As Boolean
  If Not mObj Is Nothing Then Set mObj = Nothing
End Function

Public Sub Inc(ByRef i As Long, Optional c As Long = 1)
  i = i + c
End Sub
Public Sub Dec(ByRef i As Long, Optional c As Long = 1)
  i = i - c
End Sub

Public Function RandomS(Num As Single) As Single
  'Randomize 'Num
  RandomS = Rnd * Num
End Function

Public Function RandomL(Num As Long) As Single
  'Randomize 'Num
  RandomL = Rnd * CDbl(Num)
End Function

'the next two functions are just for Delphi compatibility reasons
'in fact they are no longer needed:
'it is the same like Ubound(Arr)+1 '
'if you have <Length(Arr) - 1> so it is the same as <UBound(Arr)>
Public Function Length(pArr(), Optional nDim As Long = 1) As Long
Dim ln As Long, un As Long
tryE: On Error GoTo CatchE
  ln = LBound(pArr, nDim)
  un = UBound(pArr, nDim)
  Length = un - ln + 1
  Exit Function
CatchE:
  Length = 0
End Function

'no longer needed:
'it's the same like: ReDim Arr(cCount - 1)
Public Sub SetLength(pArr, cCount As Long, Optional nDim As Long = 0, Optional bPres As Boolean)
tryE: On Error GoTo CatchE
  If nDim > 0 Then
    If bPres Then
      ReDim Preserve pArr(0 To cCount - 1, nDim)
    Else
      ReDim pArr(0 To cCount - 1, nDim)
    End If
  Else
    If bPres Then
      ReDim Preserve pArr(0 To cCount - 1)
    Else
      ReDim pArr(0 To cCount - 1)
    End If
  End If
  Exit Sub
CatchE:
  'SetLength = 0
End Sub


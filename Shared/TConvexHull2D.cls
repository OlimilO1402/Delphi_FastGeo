VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TConvexHull2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Double Precision!
Implements TBaseConvexHull
#If defFastGeoAXdll Then
Public FastGeo As FastGeo
#End If
Private Type TCHullPoint2D
  x   As Double
  y   As Double
  ang As Double
End Type
Public Enum TCompareResult
  eGreaterThan
  eLessThan
  eEqual
End Enum
'type TSortCompare   = function(const P1,P2:TCHullPoint2D) : TCompareResult Of object;
Private Point() As TCHullPoint2D
Private Stack() As TCHullPoint2D
Private StackHeadPosition As Long
Private Anchor As TCHullPoint2D

Private Function CompareAngles(P1 As TCHullPoint2D, P2 As TCHullPoint2D) As TCompareResult
  If P1.ang < P2.ang Then
    CompareAngles = eLessThan
  ElseIf P1.ang > P2.ang Then
    CompareAngles = eGreaterThan
  ElseIf IsEqual(P1, P2) Then
    CompareAngles = eEqual
  ElseIf FastGeo.LayDistance2DXY(Anchor.x, Anchor.y, P1.x, P1.y) < FastGeo.LayDistance2DXY(Anchor.x, Anchor.y, P2.x, P2.y) Then
    CompareAngles = eLessThan
  Else
    CompareAngles = eGreaterThan
  End If
End Function

Private Function IsEqual(P1 As TCHullPoint2D, P2 As TCHullPoint2D) As Boolean
  IsEqual = FastGeo.IsEqual(P1.x, P2.x) And FastGeo.IsEqual(P1.y, P2.y)
End Function

Private Function TBaseConvexHull_ConvexHull(Pnt() As TPoint2D) As TPolygon2D
  TBaseConvexHull_ConvexHull = ConvexHull(Pnt)
End Function
Public Function ConvexHull(Pnt() As TPoint2D) As TPolygon2D
Dim i As Long, j As Long
  
  If UBound(Pnt) < 3 Then
    ReDim ConvexHull.Arr(UBound(Pnt))
    For i = 0 To UBound(Pnt) '- 1
      ConvexHull.Arr(i).x = Pnt(i).x
      ConvexHull.Arr(i).y = Pnt(i).y
    Next
    Exit Function
  End If
  
  StackHeadPosition = -1
  
tryE:
  ReDim Point(UBound(Pnt))
  ReDim Stack(UBound(Pnt))
  
  j = 0
  
  For i = 0 To UBound(Point)
    Point(i).x = Pnt(i).x
    Point(i).y = Pnt(i).y
    Point(i).ang = 0#
    If Point(i).y < Point(j).y Then
      j = i
    ElseIf Point(i).y = Point(j).y Then
      If Point(i).x < Point(j).x Then
        j = i
      End If
    End If
  Next
  
  Call Swap(0, j, Point)
  
  Point(0).ang = 0#
  Anchor = Point(0)
  
  '(* Calculate angle of the vertex ([ith point]-[anchorpoint]-[most left point]) *)
  For i = 1 To UBound(Point) '- 1
    Point(i).ang = FastGeo.CartesianAngle(Point(i).x - Anchor.x, Point(i).y - Anchor.y)
  Next
  
  '(* Sort points in ascending order according to their angles *)
  Call RQSort(1, UBound(Point), Point) ', CompareAngles) 'halt wieso eigentlich von 1???
  
  Call GrahamScan
  
  ReDim ConvexHull.Arr(StackHeadPosition) ' + 1) 'mit oder ohne preserve???
  
  For i = 0 To StackHeadPosition ' + 1
    ConvexHull.Arr(i).x = Stack(i).x
    ConvexHull.Arr(i).y = Stack(i).y
  Next
  
finallyE:
  '(* Final clean-up *)
  Erase Stack
  Erase Point
endTry:
End Function

Private Sub GrahamScan()
Dim i As Long, Orin As Long
  '(* begin graham scan process *)
  Call Push(Point(0))
  Call Push(Point(1))
  
  i = 2
  
  Do While i < UBound(Point) + 1
    If PreHeadExist Then
      Orin = Orientation(PreHead, Head, Point(i))
      If Orin = CounterClockwise Then
        Call Push(Point(i))
        Call Inc(i)
      Else
        Pop
      End If
    Else
      Call Push(Point(i))
      Call Inc(i)
    End If
  Loop
End Sub

Private Sub Push(Pnt As TCHullPoint2D)
  Call Inc(StackHeadPosition)
  Stack(StackHeadPosition) = Pnt
End Sub

Private Function Pop() As Boolean
  Pop = False
  If StackHeadPosition >= 0 Then
    Pop = True
    Call Dec(StackHeadPosition)
  End If
End Function
Private Function PopO(Head As TCHullPoint2D) As Boolean
'  If IsMissing(aHead) Then 'ein UDType kann nicht Optional sein
'    Pop = False
'    If StackHeadPosition >= 0 Then
'      Pop = True
'      Call Dec(StackHeadPosition)
'    End If
'    Exit Function
'  End If
  PopO = False
  If StackHeadPosition < 0 Then Exit Function
  Head = Stack(StackHeadPosition)
  Call Dec(StackHeadPosition)
  PopO = True
End Function

Private Function Head() As TCHullPoint2D
  'Assert((StackHeadPosition >= 0) And (StackHeadPosition < Length(Stack)),'TConvexHull2D.Head:TCHullPoint2D: Invalid stack-head position.');
  Head = Stack(StackHeadPosition)
End Function

Private Function PreHead() As TCHullPoint2D
  'Assert(((StackHeadPosition - 1) >= 0) And ((StackHeadPosition - 1) < Length(Stack)),'TConvexHull2D.PreHead:TCHullPoint2D: Invalid pre stack-head position.');
  PreHead = Stack(StackHeadPosition - 1)
End Function

Private Function PreHeadExist() As Boolean
  PreHeadExist = (StackHeadPosition > 0)
End Function

Private Function Orientation(P1 As TCHullPoint2D, P2 As TCHullPoint2D, P3 As TCHullPoint2D) As Long
  Orientation = FastGeo.Orientation2D(P1.x, P1.y, P2.x, P2.y, P3.x, P3.y)
End Function

Private Sub Swap(i As Long, j As Long, Point() As TCHullPoint2D)
Dim Temp As TCHullPoint2D
  Temp = Point(i)
  Point(i) = Point(j)
  Point(j) = Temp
End Sub

Private Sub RQSort(Left As Long, Right As Long, Point() As TCHullPoint2D, Optional Compare As TCompareResult) 'TSortCompare)
Dim i As Long, j As Long, Middle As Long
Dim Pivot As TCHullPoint2D
  
  Do
    i = Left
    j = Right
    Middle = (Left + Right) \ 2
    '(* Median of 3 Pivot Selection *)
    'If Compare(Point(Middle), Point(Left)) = eLessThan Then Call Swap(Left, Middle, Point)
    Compare = CompareAngles(Point(Middle), Point(Left))
    If Compare = eLessThan Then Call Swap(Left, Middle, Point)
    'If Compare(Point(Right), Point(Middle)) = eLessThan Then Call Swap(Right, Middle, Point)
    Compare = CompareAngles(Point(Right), Point(Middle))
    If Compare = eLessThan Then Call Swap(Right, Middle, Point)
    'If Compare(Point(Middle), Point(Left)) = eLessThan Then Call Swap(Left, Middle, Point)
    Compare = CompareAngles(Point(Middle), Point(Left))
    If Compare = eLessThan Then Call Swap(Left, Middle, Point)
    
    Pivot = Point(Right)
    
    Do
      Compare = CompareAngles(Point(i), Pivot)
      Do While Compare = eLessThan
        Call Inc(i)
        Compare = CompareAngles(Point(i), Pivot)
      Loop
      Compare = CompareAngles(Point(j), Pivot)
      Do While Compare = eGreaterThan
        Call Dec(j)
        Compare = CompareAngles(Point(j), Pivot)
      Loop
      If i <= j Then
        Call Swap(i, j, Point)
        Call Inc(i)
        Call Dec(j)
      End If
    Loop Until i > j
    If Left < j Then Call RQSort(Left, j, Point, Compare)
    Left = i
  Loop Until i >= Right
End Sub

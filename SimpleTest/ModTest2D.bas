Attribute VB_Name = "ModTest2D"
Option Explicit
Public Const MAX_POINTS As Long = 10
Private aSng As Double
Private ax As Double, bx As Double, Cx As Double, ay As Double, by As Double, Cy As Double
Private Amount As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double
Private Bx1 As Double, By1 As Double, Bx2 As Double, By2 As Double
Private P As TPoint2D, P1 As TPoint2D, P2 As TPoint2D, Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D
Private Pnt1 As TPoint2D, Pnt2 As TPoint2D, Pnt3 As TPoint2D, Pnt4 As TPoint2D
Private StartPoint As TPoint2D, N As TPoint2D, Point As TPoint2D
Private Dist As Double, Angle As Double, Distance As Double
Private S As TSegment2D, AB As TRectangle, Re As TRectangle, Rect As TRectangle, Rectangle As TRectangle
Private Segment As TSegment2D
Private Triangle As TTriangle2D
Private Quadix As TQuadix2D
Private aCircle As TCircle
Private Polygon As TPolygon2D
Private GeoObj As TGeometricObject, O As TGeometricObject
Private T As Double
Private Tr As TTriangle2D
Private c As TCircle
Private Q As TQuadix2D
Private PA As TPoint2DArray
Private Pg As TPolygon2D
Private Pts() As TPoint2D
Private BezierC As TCubicBezier2D
Private BezierQ As TQuadraticBezier2D
Private PointList() As TPoint2D
Private CuP As TCurvePoint2D
Private Ci As TCircle
Private CuA As TCurvePoint2DArray
Private PointCount As Long
Private x As Double, y As Double
Private OutSegment As TSegment2D
Private L As TLine2D, Line As TLine2D
Private R As Double, Radius As Double
Private PrimitiveP As TPoint2D
Private PrimitiveL As TLine2D
Private PrimitiveS As TSegment2D
Private PrimitiveT As TTriangle2D
Private PrimitiveQ As TQuadix2D
Private PrimitiveR As TRectangle
Private PrimitiveCi As TCircle
Private Pentagon As TPolygon2D, Hexagon As TPolygon2D, Heptagon As TPolygon2D, Octagon As TPolygon2D
Private Range As Double, Factor As Double
Private Vec As TVector2D, Vec1 As TVector2D, Vec2 As TVector2D
Private VecA As TVector2DArray, Vec1A As TVector2DArray, Vec2A As TVector2DArray
Private Bo As Boolean
Private Val1 As Double, Val2 As Double
Private Val1L As Long, Val2L As Long
Private Arc As TCircularArc2D
Private Segment1 As TSegment2D, Segment2 As TSegment2D
Private Line1 As TLine2D, Line2 As TLine2D
Private Triangle1 As TTriangle2D, Triangle2 As TTriangle2D
Private Quadix1 As TQuadix2D, Quadix2 As TQuadix2D
Private Circle1 As TCircle, Circle2 As TCircle
Private Arc1 As TCircularArc2D, Arc2 As TCircularArc2D
Private TestNPR As TNumericPrecisionResult

'This is just a quick test, calling all procedures
Public Sub Test2D()
  ReDim PA.Arr(MAX_POINTS)
  ReDim Pg.Arr(MAX_POINTS)
  PointCount = MAX_POINTS
  ReDim PointList(MAX_POINTS)
  ReDim Polygon.Arr(MAX_POINTS)
  ReDim Pentagon.Arr(5 - 1)
  ReDim Hexagon.Arr(6 - 1)
  ReDim Heptagon.Arr(7 - 1)
  ReDim Octagon.Arr(8 - 1)
  ReDim VecA.Arr(MAX_POINTS)
  ReDim Vec1A.Arr(MAX_POINTS)
  ReDim Vec2A.Arr(MAX_POINTS)
  Call GenerateRandomCircle(0, 0, 100, 100, c)
  Call GenerateRandomRectangle(0, 0, 100, 100, Re)
  S.P(0) = Re.P(0): S.P(1) = Re.P(1)
  Call GenerateRandomPoints2DR(Re, PointList)
  Call GenerateRandomQuadix(0, 0, 100, 100, Q)
  Call GenerateRandomTriangle(0, 0, 100, 100, Triangle)
  Dist = 10
  'Test All AABB
  AB = AABB2DCi(c): 'MsgBox "AB: " & StrR(AB) & " " & StrCi(C)
  AB = AABB2DCu(PA)
  AB = AABB2DPg(Pg)
  AB = AABB2DQ(Q)
  AB = AABB2DR(Re)
  AB = AABB2DS(S)
  AB = AABB2DT(Triangle)
  Call AABBS2DCi(c, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DCu(PA, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DPg(Pg, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DQ(Q, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DR(Re, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DS(S, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DT(Triangle, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DR(Re, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DQ(Q, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DCi(c, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DPg(Pg, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
  Call AABBS2DCu(PA, AB.P(0).x, AB.P(0).y, AB.P(1).x, AB.P(1).y)
'  '(* End Of AABB *)
'  'ProjectPoint
  P1.x = 5: P1.y = 5
  Call ProjectPointS2DSD(P1.x, P1.y, 10, 10, 5, P2.x, P2.y)
  Call ProjectPointS2DXY(P1.x, P1.y, 10, 10, P2.x, P2.y)
  P = ProjectPoint2DSD(P1, P2, 10)
  P = ProjectPoint2DP(P1, 10, 10)
  Call ProjectPointS0(P1.x, P1.y, 10, P2.x, P2.y)
  Call ProjectPointS45(P.x, P.y, Dist, N.x, N.y)
  Call ProjectPointS90(P.x, P.y, Dist, N.x, N.y)
  Call ProjectPointS135(P.x, P.y, Dist, N.x, N.y)
  Call ProjectPointS180(P.x, P.y, Dist, N.x, N.y)
  Call ProjectPointS225(P.x, P.y, Dist, N.x, N.y)
  Call ProjectPointS270(P.x, P.y, Dist, N.x, N.y)
  Call ProjectPointS315(P.x, P.y, Dist, N.x, N.y)
  P = ProjectPoint0(Point, Dist)
  P = ProjectPoint45(Point, Dist)
  P = ProjectPoint90(Point, Dist)
  P = ProjectPoint135(Point, Dist)
  P = ProjectPoint180(Point, Dist)
  P = ProjectPoint225(Point, Dist)
  P = ProjectPoint270(Point, Dist)
  P = ProjectPoint315(Point, Dist)
''(* End of Project Point 2D *)
''ProjectObject
  P = ProjectObject2DP(Point, Angle, Distance)
  S = ProjectObject2DS(Segment, Angle, Distance)
  Tr = ProjectObject2DT(Triangle, Angle, Distance)
  Q = ProjectObject2DQ(Quadix, Angle, Distance)
  c = ProjectObject2DCi(aCircle, Angle, Distance)
  Pg = ProjectObject2DPg(Polygon, Angle, Distance)
  O = ProjectObject2DO(GeoObj, Angle, Distance)
''(* Endof Project Object *)
''CalculateBezierCoefficients
  Call CalculateBezierCoefficientsQ2D(BezierQ, ax, bx, ay, by)
  Call CalculateBezierCoefficientsC2D(BezierC, ax, bx, Cx, ay, by, Cy)
''(* Endof Calculate Bezier Values *)
''PointOnBezier
  P = PointOnBezier2D(StartPoint, ax, bx, ay, by, T)
  P = PointOnBezier2Dc(StartPoint, ax, bx, Cx, ay, by, Cy, T)
''(* Endof Point On Bezier *)
''CreateBezier
  PA = CreateBezierQ2D(BezierQ, PointCount) 'Array
  PA = CreateBezierC2D(BezierC, PointCount) 'Array
''(* Endof Create Bezier *)
''CreateCurvePointBezier
  CuA = CreateCurvePointBezierQ2D(BezierQ, PointCount) 'As TCurvePoint2DArray
  CuA = CreateCurvePointBezierC2D(BezierC, PointCount) 'As TCurvePoint2DArray
''(* Endof Create Curve Point Bezier *)
''CurveLength
  aSng = CurveLengthQ2D(BezierQ, PointCount)
  aSng = CurveLengthC2D(BezierC, PointCount)
''(* Endof CurveLength *)
''ShortenSegment
  Call ShortenSegmentS2D(Amount, x1, y1, x2, y2)
  Call ShortenSegment2D(Segment, Amount)
''(* Endof ShortenSegment *)
''LengthenSegment
  Call LengthenSegmentS2D(Amount, x1, y1, x2, y2)
  Call LengthenSegment2D(Segment, Amount)
''(* Endof Lengthen Segment *)
''EquatePoint
  P1 = EquatePoint2D(x, y)
  Call EquatePointS2D(x, y, Point)
''(* End of Equate Point *)
''EquateCurvePoint
  CuP = EquateCurvePoint2DXY(x, y, T) 'As TCurvePoint2D
  CuP = EquateCurvePoint2DP(Point, T) 'As TCurvePoint2D
''(* End of Equate Curve Point *)
''EquateSegment
  S = EquateSegment2DXY(x1, y1, x2, y2)
  S = EquateSegment2DP(Point1, Point2)
  Call EquateSegmentS2DXY(x1, y1, x2, y2, OutSegment)
''(* End of Equate Segment *)
''EquateLine
  L = EquateLine2DXY(x1, y1, x2, y2)
  L = EquateLine2DP(Point1, Point2)
  Call EquateLineS2DXY(x1, y1, x2, y2, Line)
''(* End of Equate Line *)
''EquateQuadix
  Q = EquateQuadix2DXY(x1, y1, x2, y2, x3, y3, x4, y4)
  Q = EquateQuadix2DP(Point1, Point2, Point3, Point4)
  Call EquateQuadixS2DXY(x1, y1, x2, y2, x3, y3, x4, y4, Quadix)
''(* End of Equate Quadix *)
''EquateRectangle
  Re = EquateRectangleXY(x1, y1, x2, y2)
  Re = EquateRectangleP(Point1, Point2)
  Call EquateRectangleSXY(x1, y1, x2, y2, Rect)
  Call EquateRectangleSP(Point1, Point2, Rect)
''(* End of Equate Rectangle *)
''EquateCircle
  Ci = EquateCircleXY(x, y, R)
  Ci = EquateCircleP(Point, Radius)
  Call EquateCircleSXY(x, y, R, aCircle)
''(* End of Equate Circle *)
''EquateTriangle
  Tr = EquateTriangle2DXY(x1, y1, x2, y2, x3, y3)
  Tr = EquateTriangle2DP(Point1, Point2, Point3)
  Call EquateTriangleS2DXY(x1, y1, x2, y2, x3, y3, Triangle)
  Call EquateTriangleS2DP(Point1, Point2, Point3, Triangle)
''(* End of Equate Triangle *)
''EquateBezier
  Call EquateBezierS2D2XY(x1, y1, x2, y2, x3, y3, BezierQ)
  Call EquateBezier2D2P(Pnt1, Pnt2, Pnt3)
  Call EquateBezierS2D3XY(x1, y1, x2, y2, x3, y3, x4, y4, BezierC)
  Call EquateBezier2D3P(Pnt1, Pnt2, Pnt3, Pnt4)
''(* End of Equate Bezier *)
''RectangleToQuadix
  Call RectangleToQuadixXY(x1, y1, x2, y2)
  Call RectangleToQuadixP(Point1, Point2)
  Call RectangleToQuadix(Rectangle)
''(* Endof RectangleToQuadix *)
''TriangleToPolygon
  Call TriangleToPolygonXY(x1, y1, x2, y2, x3, y3)
  Call TriangleToPolygon(Triangle)
''(* Endof TriangleToPolygon *)
''QuadixToPolygon
  Call QuadixToPolygonXY(x1, y1, x2, y2, x3, y3, x4, y4)
  Call QuadixToPolygon(Quadix)
''(* Endof QuadixToPolygon *)
''CircleToPolygon
  Call CircleToPolygonXY(Cx, Cy, Radius, PointCount)
  Call CircleToPolygon(aCircle, PointCount)
''(* Endof Circle To Polygon *)
''SetGeometricObject
  Call SetGeometricObject2DP(PrimitiveP, GeoObj)
  Call SetGeometricObject2DL(PrimitiveL, GeoObj)
  Call SetGeometricObject2DS(PrimitiveS, GeoObj)
  Call SetGeometricObject2DT(PrimitiveT, GeoObj)
  Call SetGeometricObject2DQ(PrimitiveQ, GeoObj)
  Call SetGeometricObject2DR(PrimitiveR, GeoObj)
  Call SetGeometricObject2DCi(PrimitiveCi, GeoObj)
''(* Endof Set Geometric Object *)
  Call GenerateRandomPolygon(Bx1, By1, Bx2, By2, Polygon)
''(* Endof GenerateRandomPolygon *)
''GenerateRandomPoints
  Call GenerateRandomPoints2DXY(Bx1, By1, Bx2, By2, PointList())
  Call GenerateRandomPoints2DR(Rectangle, PointList())
  Call GenerateRandomPoints2DS(Segment, PointList())
  Call GenerateRandomPoints2DT(Triangle, PointList())
  Call GenerateRandomPoints2DCi(aCircle, PointList())
  Call GenerateRandomPoints2DQ(Quadix, PointList())
''(* Endof Generate Random Points *)
''GenerateRandomPointsOnConvexPentagon
  Call GenerateRandomPointsOnConvexPentagon(Pentagon, PointList())
''(* Endof Generate Random Points On Convex Pentagon *)
''GenerateRandomPointsOnConvexHexagon
  Call GenerateRandomPointsOnConvexHexagon(Hexagon, PointList())
''(* Endof Generate Random Points On Convex Hexagon *)
''GenerateRandomPointsOnConvexHeptagon
  Call GenerateRandomPointsOnConvexHeptagon(Heptagon, PointList())
''(* Endof Generate Random Points On Convex Heptagon *)
''GenerateRandomPointsOnConvexOctagon
  Call GenerateRandomPointsOnConvexOctagon(Octagon, PointList())
''(* Endof Generate Random Points On Convex Octagon *)
''GenerateRandomTriangle
  Call GenerateRandomTriangle(Bx1, By1, Bx2, By2, Triangle)
''(* Endof Generate Random Triangle *)
''GenerateRandomQuadix
  Call GenerateRandomQuadix(Bx1, By1, Bx2, By2, Quadix)
''(* Endof Generate Random Quadix *)
''GenerateRandomRectangle
  Call GenerateRandomRectangle(Bx1, By1, Bx2, By2, Rectangle)
''(* Endof Generate Random Rectangle *)
''GenerateRandomCircle
  Call GenerateRandomCircle(Bx1, By1, Bx2, By2, aCircle)
''(* Endof Generate Random Circle *)
''Generate Random Value
  Call GenerateRandomValue(Range, 10000000#)
''(* Endof Generate Random Value *)
''Add
  Vec = Add2D(Vec1, Vec2)
  VecA = Add2DA(Vec1A, Vec2A)
''(* End of Add *)
''Sub
  Vec = Sub2D(Vec1, Vec2)
  VecA = Sub2DA(Vec1A, Vec2A)
''(* End of Sub *)
''Mul
'  Mul2D(Vec1 As TVector2D, Vec2 As TVector2D) As TVector3D
''(* End of Multiply (cross-product) *)
''UnitVector
  Vec = UnitVector2D(Vec)
''(* End of UnitVector *)
''Magnitude
  aSng = Magnitude2D(Vec)
''(* End of Magnitude *)
''DotProduct
  aSng = DotProduct2D(Vec1, Vec2)
''(* End of dotProduct *)
''Scale
  Vec = Scale2D(Vec, Factor)
  VecA = Scale2DA(VecA, Factor)
''(* End of Scale *)
''Negate
  Vec = Negate2D(Vec1) 'As TVector2D
  VecA = Negate2DA(Vec1A) 'As TVector2DArray
''(* End of Negate *)
''IsEqual
  Bo = IsEqual(Val1, Val2)
  Bo = IsEqualEps(Val1, Val2, Epsilon)
  Bo = IsEqual2DP(Point1, Point2)
  Bo = IsEqual2DPEps(Point1, Point2, Epsilon)
''(* Endof Is Equal *)
''NotEqual:
  Bo = NotEqual(Val1, Val2) 'As Boolean
  Bo = NotEqualEps(Val1, Val2, Epsilon) 'As Boolean
  Bo = NotEqual2DP(Point1, Point2) 'As Boolean
''(* Endof not Equal *)
''LessThanOrEqual
  Bo = LessThanOrEqualEps(Val1, Val2, Epsilon)
  Bo = LessThanOrEqual(Val1, Val2)
''(* Endof Less Than Or Equal *)
''GreaterThanOrEqual
  Bo = GreaterThanOrEqualEps(Val1, Val2, Epsilon)
  Bo = GreaterThanOrEqual(Val1, Val2)
''(* Endof Greater Than Or Equal *)
''IsEqualZero
  Bo = IsEqualZeroEps(aSng, Epsilon)
  Bo = IsEqualZero(aSng)
''(* Endof IsEqualZero *)
''IsDegenerate
  Bo = IsDegenerate2DXY(x1, y1, x2, y2)
  Bo = IsDegenerate2DS(Segment)
  Bo = IsDegenerate2DL(Line)
  Bo = IsDegenerate2DT(Triangle)
  Bo = IsDegenerate2DQ(Quadix)
  Bo = IsDegenerate2DR(Rect)
  Bo = IsDegenerate2DCi(aCircle)
  Bo = IsDegenerate2DA(Arc)
  Bo = IsDegenerateO(GeoObj)
''(* Endof IsDegenerate *)
''Swap
  Call SwapS(Val1, Val2)
  Call SwapL(Val1L, Val2L)
  Call Swap2DP(Point1, Point2)
  Call Swap2DS(Segment1, Segment2)
  Call Swap2DL(Line1, Line2)
  Call Swap2DT(Triangle1, Triangle2)
  Call Swap2DQ(Quadix1, Quadix2)
  Call SwapC(Circle1, Circle2)
  Call Swap2DA(Arc1, Arc2)
''(* Endof Swap *)
''ZeroEquivalency
  'ZeroEquivalency() 'As Boolean
''(* Endof ZeroEquivalency *)
  TestNPR = ExecuteTests() 'As TNumericPrecisionResult
''(* Endof ExecuteTests *)
''RotationTest
  Bo = RotationTest(10#) 'As Boolean
''(* Endof RotationTest *)
''ExtendedFloatingPointTest
  Bo = ExtendedFloatingPointTest()
''(* Endof ExtendedFloatingPointTest *)
  
End Sub

Private Function StrR(R As TRectangle, Optional frmtNkSt As Long = 3, Optional bolMsgBox As Boolean = True) As String
Dim StrF As String
  StrF = "0." & String(frmtNkSt, "0")
  StrR = " P0(" & Format(R.P(0).x, StrF) & " " & Format(R.P(0).y, StrF) & ")" & " P1(" & Format(R.P(1).x, StrF) & " " & Format(R.P(1).y, StrF) & ")"
End Function
Private Function StrCi(Ci As TCircle, Optional frmtNkSt As Long = 3) As String
Dim StrF As String
  StrF = "0." & String(frmtNkSt, "0")
  StrCi = " M(" & Format(Ci.x, StrF) & " " & Format(Ci.y, StrF) & ")" & " R= " & Format(Ci.Radius, StrF)
End Function


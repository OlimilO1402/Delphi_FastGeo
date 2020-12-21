Attribute VB_Name = "ModTest3D"
Option Explicit
Public Const MAX_POINTS As Long = 10
Private aSng As Double
Private ax As Double, bx As Double, Cx As Double, ay As Double, by As Double, Cy As Double, az As Double, bz As Double, cz As Double
Private Amount As Double, x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double, x3 As Double, y3 As Double, z3 As Double, x4 As Double, y4 As Double, z4 As Double
Private Bx1 As Double, By1 As Double, Bx2 As Double, By2 As Double
Private P As TPoint3D, P1 As TPoint2D, P2 As TPoint2D, Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Point4 As TPoint3D
Private Pnt1 As TPoint3D, Pnt2 As TPoint3D, Pnt3 As TPoint3D, Pnt4 As TPoint3D
Private StartPoint As TPoint3D, N As TPoint2D, Point As TPoint3D
Private Dist As Double, Angle As Double, Distance As Double
Private S As TSegment3D, AB As TRectangle, Re As TRectangle, Rect As TRectangle, Rectangle As TRectangle
Private Segment As TSegment3D
Private Triangle As TTriangle3D
Private Quadix As TQuadix3D
Private aCircle As TCircle
Private Polygon As TPolygon2D
Private GeoObj As TGeometricObject, O As TGeometricObject
Private T As Double
Private Tr As TTriangle3D
Private c As TCircle
Private Q As TQuadix3D
Private PA As TPoint2DArray
Private Pg As TPolygon2D
Private Pts() As TPoint2D
Private BezierC As TCubicBezier3D
Private BezierQ As TQuadraticBezier3D
Private PointList() As TPoint3D
Private CuP As TCurvePoint3D
Private Ci As TCircle
Private Sphere As TSphere, Sph As TSphere
Private CuA As TCurvePoint3DArray
Private PointCount As Long
Private x As Double, y As Double, z As Double
Private OutSegment As TSegment3D
Private L As TLine3D, Line As TLine3D
Private R As Double, Radius As Double
Private PrimitiveP As TPoint3D
Private PrimitiveL As TLine3D
Private PrimitiveS As TSegment3D
Private PrimitiveT As TTriangle3D
Private PrimitiveQ As TQuadix3D
'Private PrimitiveR As TRectangle
'Private PrimitiveCi As TCircle
Private PrimitiveSph As TSphere
Private Pentagon As TPolygon2D, Hexagon As TPolygon2D, Heptagon As TPolygon2D, Octagon As TPolygon2D
Private Range As Double, Factor As Double
Private Vec As TVector3D, Vec1 As TVector3D, Vec2 As TVector3D
Private VecA As TVector3DArray, Vec1A As TVector3DArray, Vec2A As TVector3DArray
Private Bo As Boolean
Private Val1 As Double, Val2 As Double
Private Val1L As Long, Val2L As Long
Private Arc As TCircularArc2D
Private Segment1 As TSegment3D, Segment2 As TSegment3D
Private Line1 As TLine2D, Line2 As TLine2D
Private Triangle1 As TTriangle3D, Triangle2 As TTriangle3D
Private Quadix1 As TQuadix3D, Quadix2 As TQuadix3D
Private Circle1 As TCircle, Circle2 As TCircle
Private Arc1 As TCircularArc2D, Arc2 As TCircularArc2D
Private TestNPR As TNumericPrecisionResult
Private Pl As TPlane2D, Plane As TPlane2D
Private Sphere1 As TSphere, Sphere2 As TSphere
Public Sub Test3D()
  'ReDim PA.Arr(MAX_POINTS)
  'ReDim Pg.Arr(MAX_POINTS)
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
'ProjectObject
'(* Endof Project Object *)
'CalculateBezierCoefficients
  Call CalculateBezierCoefficientsQ3D(BezierQ, ax, bx, ay, by, az, bz)
  Call CalculateBezierCoefficientsC3D(BezierC, ax, bx, Cx, ay, by, Cy, az, bz, cz)
'(* Endof Calculate Bezier Values *)
'PointOnBezier
  P = PointOnBezier3D(StartPoint, ax, bx, ay, by, az, bz, T)
  P = PointOnBezier3Dc(StartPoint, ax, bx, Cx, ay, by, Cy, az, bz, cz, T)
'(* Endof Point On Bezier *)
'CreateBezier
  Call CreateBezierQ3D(BezierQ, PointCount) 'Array
  Call CreateBezierC3D(BezierC, PointCount) 'Array
'(* Endof Create Bezier *)
'CreateCurvePointBezier
  CuA = CreateCurvePointBezierQ3D(BezierQ, PointCount)
  CuA = CreateCurvePointBezierC3D(BezierC, PointCount)
'(* Endof Create Curve Point Bezier *)
'CurveLength
  aSng = CurveLengthQ3D(BezierQ, PointCount)
  aSng = CurveLengthC3D(BezierC, PointCount)
'(* Endof CurveLength *)
'ShortenSegment
  Call ShortenSegmentS3D(Amount, x1, y1, z1, x2, y2, z2)
  Call ShortenSegment3D(Segment, Amount)
'(* Endof ShortenSegment *)
'LengthenSegment
  Call LengthenSegmentS3D(Amount, x1, y1, z1, x2, y2, z2)
  Call LengthenSegment3D(Segment, Amount)
'(* Endof Lengthen Segment *)
'EquatePoint
  P = EquatePoint3D(x, y, z)
  Call EquatePointS3D(x, y, z, Point)
'(* End of Equate Point *)
'EquateCurvePoint
  CuP = EquateCurvePoint3DXY(x, y, z, T)
  CuP = EquateCurvePoint3DP(Point, T)
'(* End of Equate Curve Point *)
'EquateSegment
  S = EquateSegment3DXY(x1, y1, z1, x2, y2, z2)
  S = EquateSegment3DP(Point1, Point2)
  Call EquateSegmentS3DXY(x1, y1, z1, x2, y2, z2, OutSegment)
'(* End of Equate Segment *)
'EquateLine
  L = EquateLine3DXY(x1, y1, z1, x2, y2, z2) 'As TLine3D
'  EquateLine3DP(Point1, Point2) As TLine3D
  Call EquateLineS3DXY(x1, y1, z1, x2, y2, z2, Line)
'(* End of Equate Line *)
'EquateQuadix
  Q = EquateQuadix3DXY(x1, y1, z1, x2, y2, z2, x3, y3, z3, x4, y4, z4)
  Q = EquateQuadix3DP(Point1, Point2, Point3, Point4)
  Call EquateQuadixS3DXY(x1, y1, z1, x2, y2, z2, x3, y3, z3, x4, y4, z4, Quadix)
'(* End of Equate Quadix *)
'EquateSphere
  Sph = EquateSphereXY(x, y, z, R)
  Call EquateSphereSXY(x, y, z, R, Sphere)
'(* End of Equate Sphere *)
'EquateTriangle
  Tr = EquateTriangle3DXY(x1, y1, z1, x2, y2, z2, x3, y3, z3)
  Tr = EquateTriangle3DP(Point1, Point2, Point3)
  Call EquateTriangleS3DXY(x1, y1, z1, x2, y2, z2, x3, y3, z3, Triangle)
  Call EquateTriangleS3DP(Point1, Point2, Point3, Triangle)
'(* End of Equate Triangle *)
'EquatePlane
  Pl = EquatePlane3DXY(x1, y1, z1, x2, y2, z2, x3, y3, z3) 'As TPlane2D
  Call EquatePlaneS3DXY(x1, y1, z1, x2, y2, z2, x3, y3, z3, Plane) 'As TPlane2D)
  Pl = EquatePlane3DP(Point1, Point2, Point3) 'As TPlane2D
  Call EquatePlaneS3DP(Point1, Point2, Point3, Plane)
'(* End of Equate Plane *)
'EquateBezier
  Call EquateBezierS3D2XY(x1, y1, z1, x2, y2, z2, x3, y3, z3, BezierQ)
  BezierQ = EquateBezier3D2P(Pnt1, Pnt2, Pnt3)
  Call EquateBezierS3D3XY(x1, y1, z1, x2, y2, z2, x3, y3, z3, x4, y4, z4, BezierC)
  BezierC = EquateBezier3D3P(Pnt1, Pnt2, Pnt3, Pnt4)
'(* End of Equate Bezier *)
'SetGeometricObject
  Call SetGeometricObject3DP(PrimitiveP, GeoObj)
  Call SetGeometricObject3DL(PrimitiveL, GeoObj)
  Call SetGeometricObject3DS(PrimitiveS, GeoObj)
  Call SetGeometricObject3DT(PrimitiveT, GeoObj)
  Call SetGeometricObject3DQ(PrimitiveQ, GeoObj)
  Call SetGeometricObject3DSph(PrimitiveSph, GeoObj)
'(* Endof Set Geometric Object *)
'GenerateRandomPoints
  Call GenerateRandomPoints3DS(Segment, PointList())
'(* Endof Generate Random Points *)
'Add
  Vec = Add3D(Vec1, Vec2) 'As TVector3D
  VecA = Add3DA(Vec1A, Vec2A) 'As TVector3DArray
'(* End of Add *)
'Sub
  Vec = Sub3D(Vec1, Vec2)
  VecA = Sub3DA(Vec1A, Vec2A)
'(* End of Sub *)
'Mul
  Vec = Mul3D(Vec1, Vec2)
  VecA = Mul3DA(Vec1A, Vec2A)
'(* End of Multiply (cross-product) *)
'UnitVector
  Vec1 = UnitVector3D(Vec)
'(* End of UnitVector *)
'Magnitude
  aSng = Magnitude3D(Vec)
'(* End of Magnitude *)
'DotProduct
  aSng = DotProduct3D(Vec1, Vec2)
'(* End of dotProduct *)
'Scale
  Vec1 = Scale3D(Vec, Factor)
  Vec1A = Scale3DA(VecA, Factor)
'(* End of Scale *)
'Negate
  Vec1 = Negate3D(Vec)
  Vec1A = Negate3DA(VecA)
'(* End of Negate *)
'IsEqual
  Bo = IsEqual3DP(Point1, Point2)
  Bo = IsEqual3DPEps(Point1, Point2, Epsilon)
'(* Endof Is Equal *)
'NotEqual:
  Bo = NotEqual3DP(Point1, Point2)
  Bo = NotEqual3DPEps(Point1, Point2, Epsilon)
'(* Endof not Equal *)
'IsDegenerate
  Bo = IsDegenerate3DXY(x1, y1, z1, x2, y2, z2)
  Bo = IsDegenerate3DS(Segment)
  Bo = IsDegenerate3DL(Line)
  Bo = IsDegenerate3DT(Triangle)
  Bo = IsDegenerate3DQ(Quadix)
  Bo = IsDegenerate3DSph(Sphere)
'(* Endof IsDegenerate *)
'Swap
  Call Swap3DP(Point1, Point2)
  Call Swap3DS(Segment1, Segment2)
  Call Swap3DT(Triangle1, Triangle2)
  Call Swap3DQ(Quadix1, Quadix2)
  Call SwapSph(Sphere1, Sphere2)
'(* Endof Swap *)
End Sub

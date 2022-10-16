Attribute VB_Name = "FGTypes"
Option Explicit
'Double Precision!
'VB und TypeLibs
'Wenn man in einer Standard-exe UDTypes öffentlich in einer Klasse benutzen will,
'braucht man entweder eine
'* TypeLib
'  oder eine
'* AXdll
'  mit bereits definierten Public UDTypes in einer öffentl. Klasse
'
'Wenn eine ActiveXdll UDTypes öffentlich benützt,
'so braucht sie entweder eine
'* TypeLib
'  Alle Standard-exes, die die dll verwenden, müssen sowohl
'  auf die ActiveXdll als auch auf die TypeLib verweisen
'  oder die
'* ActiveXdll
'  definiert die UDTypes in einer öffentl. Klasse
'  dabei wird keine zusätzliche TypeLib benötigt
'
''Alle Types sind auch in TLBFastGeo.tlb definiert
''Double Precision!
''(****************************************************************************)
''(********************[ Basic Geometric Structure Types ]*********************)
''(****************************************************************************)
''(**************[  Vertex type   ]***************)
''Ein 2D Punkt:
'Public Type TPoint2D
'  x As Double
'  y As Double
'End Type
'Public Type TPoint2DArray
'  Arr() As TPoint2D
'End Type
''PointerTyp auf den Punkt:
''Public Enum TPoint2DPtr
''  [VarPtrTPoint2D] 'is nur'n Long-Pointer
''End Enum
''
''(**************[ 3D Vertex type ]***************)
''Ein 3D Punkt:
'Public Type TPoint3D
'  x As Double
'  y As Double
'  z As Double
'End Type
'Public Type TPoint3DArray
'  Arr() As TPoint3D
'End Type
''PointerTyp auf den Punkt:
''Public Enum TPoint3DPtr
''  [VarPtrTPoint3D] 'is nur'n Long-Pointer
''End Enum
''
'''(**************[  Quadix type   ]***************)
''Ein unregelmäßiges Viereck
'Public Type TQuadix2D
'  P(0 To 3) As TPoint2D
'End Type
'Public Type TQuadix2DArray
'  Arr() As TQuadix2D
'End Type
''Public Enum TQuadix2DPtr
''  [VarPtrTQuadix2D]
''End Enum
'Public Type TQuadix3D
'  P(0 To 3) As TPoint3D
'End Type
'Public Type TQuadix3DArray
'  Arr() As TQuadix3D
'End Type
''Public Enum TQuadix3DPtr
''  [VarPtrTQuadix3D]
''End Enum
''
''(**************[ Rectangle type ]***************)
''Ein 2D Rechteck
'Public Type TRectangle
'  P(0 To 1) As TPoint2D
'End Type
'Public Type TRectangleArray
'  Arr() As TRectangle
'End Type
''PointerTyp auf das Rechteck
''Public Enum TRectanglePtr
''  [VarPtrTRectangle2D] 'is nur'n Long-Pointer
''End Enum
''
''(**************[ Triangle type  ]***************)
'Public Type TTriangle2D
'  P(0 To 2) As TPoint2D
'End Type
'Public Type TTriangle2DArray
'  Arr() As TTriangle2D
'End Type
''Public Enum TTriangle2DPtr
''  [_]
''End Enum
'Public Type TTriangle3D
'  P(0 To 2) As TPoint3D
'End Type
'Public Type TTriangle3DArray
'  Arr() As TTriangle3D
'End Type
''Public Enum TTriangle3DPtr
''  [_]
''End Enum
''
''(**************[  Segment type  ]***************)
''Eine Strecke
'Public Type TSegment2D
'  P(0 To 1) As TPoint2D
'End Type
'Public Type TSegment2DArray
'  Arr() As TSegment2D
'End Type
''PointerTyp auf eine Strecke
''Public Enum TSegment2DPtr
''  [VarPtrTSegment2D]   'is nur'n Long-Pointer
''End Enum
'Public Type TSegment3D
'  P(0 To 1) As TPoint3D
'End Type
'Public Type TSegment3DArray
'  Arr() As TSegment3D
'End Type
''PointerTyp auf eine Strecke
''Public Enum TSegment3DPtr
''  [VarPtrTSegment3D]   'is nur'n Long-Pointer
''End Enum
''
''(**************[  Line type  ]***************)
''eine Linie
'Public Type TLine2D
'  P(0 To 1) As TPoint2D
'End Type
'Public Type TLine2DArray
'  Arr() As TLine2D
'End Type
''Public Enum TLine2DPtr
''  [_]
''End Enum
'Public Type TLine3D
'  P(0 To 1) As TPoint3D
'End Type
'Public Type TLine3DArray
'  Arr() As TLine3D
'End Type
''Public Enum TLine3DPtr
''  [_]
''End Enum
''
''(**************[  Circle type   ]***************)
'Public Type TCircle
'  x As Double
'  y As Double
'  Radius As Double
'End Type
'Public Type TCircleArray
'  Arr() As TCircle
'End Type
''Public Enum TCirclePtr
''  [_]
''End Enum
''
''(**************[  Sphere type   ]***************)
'Public Type TSphere
'  x As Double
'  y As Double
'  z As Double
'  Radius As Double
'End Type
'Public Type TSphereArray
'  Arr() As TSphere
'End Type
''Public Enum TSpherePtr
''  [_]
''End Enum
''
''(**************[  Arc type   ]***************)
'Public Type TCircularArc2D
'  x1 As Double
'  y1 As Double
'  x2 As Double
'  y2 As Double
'  Cx As Double
'  Cy As Double
'  Px As Double
'  Py As Double
'  angle1 As Double
'  angle2 As Double
'  Orientation As Long
'End Type
'Public Type TCircularArc2DArray
'  Arr() As TCircularArc2D
'End Type
''Public Enum TCircularArc2DPtr
''  [_]
''End Enum
''(************[  Bezier type   ]*************)
'Public Type TQuadraticBezier2D
'  P(0 To 2) As TPoint2D
'End Type
'Public Type TQuadraticBezier2DArray
'  Arr() As TQuadraticBezier2D
'End Type
'
'Public Type TQuadraticBezier3D
'  P(0 To 2) As TPoint3D
'End Type
'Public Type TQuadraticBezier3DArray
'  Arr() As TQuadraticBezier3D
'End Type
'
'Public Type TCubicBezier2D
'  P(0 To 3) As TPoint2D
'End Type
'Public Type TCubicBezier2DArray
'  Arr() As TCubicBezier2D
'End Type
'
'Public Type TCubicBezier3D
'  P(0 To 3) As TPoint3D
'End Type
'
'Public Type TCubicBezier3DArray
'  Arr() As TCubicBezier3D
'End Type
'
'Public Type TCurvePoint2D
' x As Double
' y As Double
' T As Double
'End Type
'Public Type TCurvePoint3D
'  x As Double
'  y As Double
'  z As Double
'  T As Double
'End Type
'
'Public Type TCurvePoint2DArray
'  Arr() As TCurvePoint2D
'End Type
'
'Public Type TCurvePoint3DArray
'  Arr() As TCurvePoint3D
'End Type
'
'
''
''(**************[ 2D Vector type ]***************)
'Public Type TVector2D
'  x As Double
'  y As Double
'End Type
'Public Type TVector2DArray
'  Arr() As TVector2D
'End Type
''Public Enum TVector2DPtr
''  [_]
''End Enum
''
''(**************[ 3D Vector type ]***************)
'Public Type TVector3D
'  x As Double
'  y As Double
'  z As Double
'End Type
'Public Type TVector3DArray
'  Arr() As TVector3D
'End Type
''Public Enum TVector3DPtr
''  [_]
''End Enum
''
''(**********[ Polygon Vertex type  ]************)
'Public Type TPolygon2D
'   Arr() As TPoint2D
'End Type
''Public Type TPolygon2DArray
''  Arr() As TPolygon2D
''End Type
''Public Enum TPolygon2DPtr
''  [_]
''End Enum
''
'Public Type TPolyLine2D
'   Arr() As TPoint2D
'End Type
''Public Type TPolyLine2DArray
''  Arr() As TPolyLine2D
''End Type
''Public Enum TPolyLine2DPtr
''  [_]
''End Enum
''
'Public Type TPolygon3D
'   Arr() As TPoint3D
'End Type
''Public Type TPolygon3DArray
''  Arr() As TPolygon3D
''End Type
''Public Enum TPolygon3DPtr
''  [_]
''End Enum
''
'Public Type TPolyhedron
'   Arr() As TPolygon3D
'End Type
''Public Enum TPolyhedronPtr
''  [_]
''End Enum
''
''
''(**************[ Plane type ]******************)
'Public Type TPlane2D
'  a As Double
'  b As Double
'  C As Double
'  D As Double
'End Type
''Public Enum TPlane2DPtr
''  [_]
''End Enum
''
''(**********[ Barycentric Coordinates]***********)
'Public Type TBarycentricUnit
'  x1 As Double
'  y1 As Double
'  x2 As Double
'  y2 As Double
'  x3 As Double
'  y3 As Double
'  delta As Double
'End Type
''Public Enum TBarycentricUnitPtr
''  [_]
''End Enum
''
'Public Type TBarycentricTriplet
'  u As Double
'  v As Double
'  w As Double
'End Type
''Public Enum TBarycentricTripletPtr
''  [_]
''End Enum
''
'Public Enum TInclusion
'  eFully
'  ePartially
'  eOutside
'  eUnknown
'End Enum
'Public Enum eTriangletype
'  etEquilateral
'  etIsosceles
'  etRight
'  etScalene
'  etObtuse
'  etUnknown
'End Enum
''
''(********[ Universal Geometric Variable ]********)
'Public Enum eGeometricObjectTypes
'  geoPoint2D
'  geoPoint3D
'  geoLine2D
'  geoLine3D
'  geoSegment2D
'  geoSegment3D
'  geoQuadix2D
'  geoQuadix3D
'  geoTriangle2D
'  geoTriangle3D
'  geoRectangle
'  geoCircle
'  geoSphere
'  geoPolygon2D
'  geoPolygon3D
'  geoQuadraticBezier2D
'  geoQuadraticBezier3D
'  geoCubicBezier2D
'  geoCubicBezier3D
'  geoPolyhedron
'End Enum
'Public Type TGeometricObject
'  ObjectType As eGeometricObjectTypes
'  Point2D    As TPoint2D
'  Point3D    As TPoint3D
'  Line2D     As TLine2D
'  Line3D     As TLine3D
'  Segment2D  As TSegment2D
'  Segment3D  As TSegment3D
'  Triangle2D As TTriangle2D
'  Triangle3D As TTriangle3D
'  Quadix2D   As TQuadix2D
'  Quadix3D   As TQuadix3D
'  Rectangle  As TRectangle
'  aCircle    As TCircle
'  Sphere     As TSphere
'  Polygon2D  As TPolygon2D
'  Polygon3D  As TPolygon3D
'  QBezier2D  As TQuadraticBezier2D
'  QBezier3D  As TQuadraticBezier3D
'  CBezier2D  As TCubicBezier2D
'  CBezier3D  As TCubicBezier3D
'  Polyhedron As TPolyhedron
'End Type
''
'Public Type TBooleanArray
'  Arr() As Boolean
'End Type
'Public Type TNumericPrecisionResult
'  EEResult As Boolean '// Epsilon equivelence result
'  ZEResult As Boolean '// Zero equivelence result;
'  EFPResult As Boolean '// Extended floating point test result
'  SystemEpsilon As Double '
'End Type

'(**********[ Orientation constants ]**********)
Public Const RightHandSide        As Long = -1
Public Const LeftHandSide         As Long = 1
Public Const Clockwise            As Long = -1
Public Const CounterClockwise     As Long = 1
Public Const CollinearOrientation As Long = 0
Public Const AboveOriention       As Long = 1
Public Const BelowOrientation     As Long = -1
Public Const CoplanarOrientation  As Long = 0

'(************[ Epsilon constants ]*************)
Public Const Epsilon_High      As Double = 1E-16
Public Const Epsilon_Medium    As Double = 0.000000000001
Public Const Epsilon_Low       As Double = 0.00000001
Public Const Epsilon           As Double = Epsilon_Medium

'{$IFDEF FASTGEO_Double_PRECISION}
Public Const Infinity          As Double = 1E+30
'{$ENDIF}

'(*******[ Random resolution constants ]********)

Public Const RandomResolutionInt = 1000000000
Public Const RandomResolutionFlt = RandomResolutionInt * 1#

'const PI2       as Double=  6.283185307179586476925286766559000
'const PIDiv180  as Double=  0.017453292519943295769236907684886
'const w180DivPI as Double= 57.295779513082320876798154814105000

Public Const PI2       As Double = 6.28318530717959
Public Const PIDiv180  As Double = 1.74532925199433E-02
Public Const C180DivPI As Double = 57.2957795130823

Public Function UBoundDim(ByRef Arr() As Variant) As Long
tryE: On Error GoTo Catch
  UBoundDim = UBound(Arr)
  Exit Function
Catch:
  UBoundDim = 0
End Function

Public Function MinL(LngVal1 As Long, LngVal2 As Long) As Long
  If LngVal1 < LngVal2 Then MinL = LngVal1 Else MinL = LngVal2
End Function
Public Function MaxL(LngVal1 As Long, LngVal2 As Long) As Long
  If LngVal1 > LngVal2 Then MaxL = LngVal1 Else MaxL = LngVal2
End Function

Public Function MinD(DblVal1 As Double, DblVal2 As Double) As Double
  If DblVal1 < DblVal2 Then MinD = DblVal1 Else MinD = DblVal2
End Function
Public Function MaxD(DblVal1 As Double, DblVal2 As Double) As Double
  If DblVal1 > DblVal2 Then MaxD = DblVal1 Else MaxD = DblVal2
End Function

Public Function ArcCos(D As Double) As Double
  ArcCos = (3.14159265358979 / 2) - Atn(D / (Sqr(1 - D ^ 2)))
End Function
Public Function ArcSin(D As Double) As Double
  ArcSin = Atn(D / (Sqr(1 - D ^ 2)))
End Function
Public Function ArcTan(D As Double) As Double
  ArcTan = Atn(D)
End Function

Public Function Trunc(DblVal As Double) As Double
  Trunc = CDbl(CLng(DblVal))
End Function

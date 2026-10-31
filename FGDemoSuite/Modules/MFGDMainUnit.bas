Attribute VB_Name = "MFGDMainUnit"
Option Explicit
'Double Precision!
'Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'originally there were several (four) different testing programs:
Public Enum ETest
  etBasicFuncs = 0
  etBarycentric = 1
  etConvexHull = 2
  etOrderedPoly = 3
End Enum

'###############'  BASIC FUNCTIONS  '###############'
Public Enum EFunctions
  eSegmentIntersections
  eRandomSegmentIntersections
  eRandomCircleIntersections
  eSegmentTriangleClipping
  eSegmentQuadixClipping
  eSegmentRectangleClipping
  eSegmentCircleClipping
  eCircumcircle
  eInscribedCircle
  eTorricelliPoint
  eCPOTriangleFP
  eCPOQuadixFP
  eCPOCircleFP
  eCPOAABBFP
  eCPOCircleFS
  eRandPntsAABB
  eRandPntsTriangle
  eRandPntsQuadix
  eRandPntsCircle
  ePolygonCentroid
  eRotateSegments
  eRotateTriangles
  eRotateQuadix
End Enum
Public Const LoweFunctions As Long = eSegmentIntersections
Public Const HigheFunctions As Long = eRotateQuadix
Public cFunctionStrings() As String 'LoweFunctions To HigheFunctions) As String '= Array("Random Segment Intersections", "Random Circle Intersections", "Segment Triangle Clipping", "Segment Quadix Clipping", "Segment Rectangle Clipping", "Segment Circle Clipping", "Circumcircle", "Inscribed Circle", "TorricelliPoint", "Closest Point On Triangle Form Point", "Closest Point On Quadix Form Point", "Closest Point On Circle Form Point", "Closest Point On AABB Form Point", "Closest Point On Circle Form Segment", "Random Points Within AABB", "Random Points Within Triangle", "Random Points Within Quadix", "Random Point Within Circle", "Polygon Centroid", "Rotate Segments", "Rotate Triangles", "Rotate Quadix")
'type FastGEOProctor = procedure(const Graphic : TFastGeoGraphics)
'Public Enum FastGEOProctor
'  [_]
'End Enum
Public Type FGFunctionality
  functype As EFunctions
  proctor  As Long 'FastGEOProctor 'AddressOf-Zeiger auf die Sub, ne geht nicht
  Valid    As Boolean
End Type

Public Sub FastGEOProctorSub(Graphic As TFastGEOGraphics)
  'not needed now
End Sub

Public Function Create_FGFType(functype As EFunctions, proctor As Long) As FGFunctionality 'FastGEOProctor
  Create_FGFType.functype = functype
  Create_FGFType.proctor = proctor
  Create_FGFType.Valid = True
End Function


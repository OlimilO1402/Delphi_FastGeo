VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FGFunctionalities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Double Precision!
Private mCurIndex As Long 'fg_functionality
Private mArr() As FGFunctionality
Private mCapacity As Long

Private Sub Class_Initialize()
  '
End Sub

Private Sub Class_Terminate()
  Erase mArr
End Sub

Public Function Item(nIndex As Long) As FGFunctionalities
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "40"
  'mCur = mArr(nIndex)
  mCurIndex = nIndex
  Set Item = Me
End Function
Public Property Get functype(nIndex As Long) As EFunctions
  functype = mArr(nIndex).functype
End Property
Public Property Let functype(nIndex As Long, EVal As EFunctions)
  mArr(nIndex).functype = EVal
End Property

Public Property Get Valid(nIndex As Long) As Boolean
  Valid = mArr(nIndex).Valid
End Property
Public Property Let Valid(nIndex As Long, BolVal As Boolean)
  mArr(nIndex).Valid = BolVal
End Property

Public Sub SetLength(aCapacity As Long, Optional bPreserve As Boolean)
  mCapacity = aCapacity
  If bPreserve Then
    ReDim Preserve mArr(0 To mCapacity)
  Else
    ReDim mArr(0 To mCapacity)
  End If
End Sub
Public Property Get Length() As Long
  Length = mCapacity + 1
End Property
Public Sub Create_FGFType(nIndex As Long, functype As EFunctions, Optional proct As Long) 'As FGFunctionality 'FastGEOProctor
  mArr(nIndex).functype = functype
  'mCur.proctor = proct
  mArr(nIndex).Valid = True
  'mcur = mArr(nIndex)
End Sub

Public Sub proctor(Graphic As TFastGEOGraphics)
Dim mcur As FGFunctionality
  mcur = mArr(mCurIndex)
  'If mCur.Valid Then
    Select Case mcur.functype
    Case eSegmentIntersections:       Call SegmentIntersect00(Graphic)
    Case eRandomSegmentIntersections: Call SegmentIntersect01(Graphic)
    Case eRandomCircleIntersections:  Call CircleIntersect00(Graphic) 'CircleIntersect01(Graphic)
    Case eSegmentTriangleClipping:    Call SegmentTriangleClip(Graphic)
    Case eSegmentQuadixClipping:      Call SegmentQuadixClip(Graphic)
    Case eSegmentRectangleClipping:   Call SegmentRectangleClip(Graphic)
    Case eSegmentCircleClipping:      Call SegmentCircleClip(Graphic)
    Case eCircumcircle:               Call TriangleCircumcircle(Graphic)
    Case eInscribedCircle:            Call TriangleInscribedCircle(Graphic)
    Case eTorricelliPoint:            Call TriangleTorricelliPoint(Graphic)
    Case eCPOAABBFP:                  Call CPOnAABBFromPoint(Graphic)
    Case eCPOCircleFP:                Call CPOnCircleFromPoint(Graphic)
    Case eCPOCircleFS:                Call CPOnCircleFromSegment(Graphic)
    Case eCPOQuadixFP:                Call CPOnQuadixFromPoint(Graphic)
    Case eCPOTriangleFP:              Call CPOnTriangleFromPoint(Graphic)
    Case eRandPntsAABB:               Call RandomPointsAABB(Graphic)
    Case eRandPntsTriangle:           Call RandomPointsTriangle(Graphic)
    Case eRandPntsQuadix:             Call RandomPointsQuadix(Graphic)
    Case eRandPntsCircle:             Call RandomPointsCircle(Graphic)
    Case ePolygonCentroid:            Call PolygonCenter(Graphic)
    Case eRotateQuadix:               Call RotateQuadices(Graphic)
    Case eRotateSegments:             Call RotateSegments(Graphic)
    Case eRotateTriangles:            Call RotateTriangles(Graphic)
    End Select
  'End If
End Sub


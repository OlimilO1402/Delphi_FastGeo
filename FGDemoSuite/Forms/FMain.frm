VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FastGeoDemosuite by Partow.net now in the VB-World by MBO-Ing.com"
   ClientHeight    =   8415
   ClientLeft      =   435
   ClientTop       =   1125
   ClientWidth     =   11655
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   8415
   ScaleWidth      =   11655
   Begin VB.TextBox InfoMemo 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   6600
      Width           =   3255
   End
   Begin VB.PictureBox PaintBox 
      BackColor       =   &H00FFFFFF&
      Height          =   8175
      Left            =   3480
      ScaleHeight     =   541
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   533
      TabIndex        =   1
      Top             =   120
      Width           =   8055
   End
   Begin VB.Frame FraBasicFunctions 
      Caption         =   "Basic Functions"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox PanelBasicFncs 
         BorderStyle     =   0  'Kein
         Height          =   6015
         Left            =   120
         ScaleHeight     =   6015
         ScaleWidth      =   3015
         TabIndex        =   3
         Top             =   240
         Width           =   3015
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   25
            Top             =   240
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   24
            Top             =   480
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   23
            Top             =   720
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   22
            Top             =   960
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   21
            Top             =   1200
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   20
            Top             =   1440
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   19
            Top             =   1680
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   18
            Top             =   1920
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   17
            Top             =   2160
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   16
            Top             =   2400
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   15
            Top             =   2640
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   14
            Top             =   2880
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   13
            Top             =   3120
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   14
            Left            =   0
            TabIndex        =   12
            Top             =   3360
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   11
            Top             =   3600
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   10
            Top             =   3840
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   9
            Top             =   4080
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   18
            Left            =   0
            TabIndex        =   8
            Top             =   4320
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   7
            Top             =   4560
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   20
            Left            =   0
            TabIndex        =   6
            Top             =   4800
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   21
            Left            =   0
            TabIndex        =   5
            Top             =   5040
            Width           =   3015
         End
         Begin VB.OptionButton BFGroupOption 
            Caption         =   "Option1"
            Height          =   255
            Index           =   22
            Left            =   0
            TabIndex        =   4
            Top             =   5280
            Width           =   3015
         End
         Begin VB.CommandButton RefreshButton 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   360
            TabIndex        =   26
            Top             =   5640
            Width           =   2175
         End
      End
   End
   Begin VB.Frame FraConvexHull 
      Caption         =   "Convex Hull Polygon"
      Height          =   6375
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox PanelConvexHull 
         BorderStyle     =   0  'Kein
         Height          =   6015
         Left            =   120
         ScaleHeight     =   6015
         ScaleWidth      =   3015
         TabIndex        =   31
         Top             =   240
         Width           =   3015
         Begin VB.ComboBox PointCountComboBox 
            Height          =   315
            ItemData        =   "FMain.frx":1782
            Left            =   0
            List            =   "FMain.frx":17CE
            TabIndex        =   44
            Top             =   120
            Width           =   3015
         End
         Begin VB.ComboBox CHStrategyComboBox 
            Height          =   315
            ItemData        =   "FMain.frx":1859
            Left            =   0
            List            =   "FMain.frx":1863
            TabIndex        =   43
            Top             =   1320
            Width           =   3015
         End
         Begin VB.CommandButton GenRandPntsBtn 
            Caption         =   "Generate Random Points"
            Height          =   375
            Left            =   0
            TabIndex        =   42
            Top             =   3000
            Width           =   3015
         End
         Begin VB.CommandButton DisplayCHullBtn 
            Caption         =   "Display Convex Hull"
            Height          =   375
            Left            =   0
            TabIndex        =   41
            Top             =   3480
            Width           =   3015
         End
         Begin VB.CheckBox DrawCheckBox 
            Caption         =   "Draw Hull &Points"
            Height          =   255
            Left            =   0
            TabIndex        =   40
            Top             =   600
            Value           =   1  'Aktiviert
            Width           =   2895
         End
         Begin VB.CommandButton BeginTestButton 
            Caption         =   "Begin Test"
            Height          =   375
            Left            =   0
            TabIndex        =   39
            Top             =   4320
            Width           =   1455
         End
         Begin VB.CommandButton StopTestButton 
            Caption         =   "Stop Test"
            Height          =   375
            Left            =   1560
            TabIndex        =   38
            Top             =   4320
            Width           =   1455
         End
         Begin VB.TextBox TestCounterEdit 
            Height          =   285
            Left            =   0
            TabIndex        =   37
            Top             =   3960
            Width           =   3015
         End
         Begin VB.Frame Frame1 
            Caption         =   "Produce Output"
            Height          =   1095
            Left            =   0
            TabIndex        =   34
            Top             =   1800
            Width           =   3015
            Begin VB.CheckBox StdOutputCheckBox 
               Caption         =   "Standard Output"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   720
               Width           =   1815
            End
            Begin VB.CheckBox PostScriptCheckBox 
               Caption         =   "Encapsulated Postscript (eps)"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   360
               Width           =   2775
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Convex Hull Strategy"
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   1080
            Width           =   2895
         End
      End
   End
   Begin VB.Frame FraBaryCentric 
      Caption         =   "Barycentric Triangle"
      Height          =   6375
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox PanelBaryCentric 
         BorderStyle     =   0  'Kein
         Height          =   6015
         Left            =   120
         ScaleHeight     =   6015
         ScaleWidth      =   3015
         TabIndex        =   32
         Top             =   240
         Width           =   3015
         Begin VB.TextBox UEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   56
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox VEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   55
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox WEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   54
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox BaryTotalEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   53
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox BXEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   52
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox BYEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   51
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox MXEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   50
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox MYEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   49
            Top             =   3480
            Width           =   1455
         End
         Begin VB.TextBox DeltaXEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   48
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox DeltaYEdit 
            Height          =   285
            Left            =   1200
            TabIndex        =   47
            Top             =   4440
            Width           =   1455
         End
         Begin VB.CommandButton ReGenButton 
            Caption         =   "Regenerate Triangle"
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   5040
            Width           =   2415
         End
         Begin VB.Label Label11 
            Caption         =   "u ="
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "v ="
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "w ="
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "b.x ="
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label5 
            Caption         =   "b.y ="
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "m.x ="
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   3120
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "m.y ="
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "DeltaX ="
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   4080
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "DeltaY ="
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "E ="
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   1440
            Width           =   495
         End
      End
   End
   Begin VB.Frame FraOrderedPolygon 
      Caption         =   "Ordered Nonintersecting Polygon"
      Height          =   6375
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox PanelOrderedPolygon 
         BorderStyle     =   0  'Kein
         Height          =   6015
         Left            =   120
         ScaleHeight     =   6015
         ScaleWidth      =   3015
         TabIndex        =   33
         Top             =   240
         Width           =   3015
         Begin VB.CommandButton OrdPolyBtn 
            Caption         =   "Order Polygon"
            Height          =   615
            Left            =   0
            TabIndex        =   69
            Top             =   1560
            Width           =   3015
         End
         Begin VB.CommandButton DispEdgeBtn 
            Caption         =   "Display Edges"
            Height          =   615
            Left            =   0
            TabIndex        =   68
            Top             =   840
            Width           =   3015
         End
         Begin VB.CommandButton BtnGenRndPoints 
            Caption         =   "Generate Random Points"
            Height          =   615
            Left            =   0
            TabIndex        =   67
            Top             =   120
            Width           =   3015
         End
      End
   End
   Begin VB.Menu mnuDemo 
      Caption         =   "Demos"
      Begin VB.Menu mnuDemoBasicFunctions 
         Caption         =   "Basic Functions"
      End
      Begin VB.Menu mnuDemoBarycentric 
         Caption         =   "Barycentric Triangle"
      End
      Begin VB.Menu mnuDemoConvexHull 
         Caption         =   "Convex Hull Polygon"
      End
      Begin VB.Menu mnuDemoOrderedPolygon 
         Caption         =   "Ordered Nonintersecting Polygon"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDemoExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " ? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Double Precision!
Private PaintBoxCanvas As TCanvas
Private TestWhat As ETest

'###############'  BASIC FUNCTIONS  '###############'
Private FuncList As FGFunctionalities
Private BFGroup  As TRadioGroup

'###############'   BARYCENTRIC     '###############'
Private BU As TBarycentricUnit
Private PrevCheckIn As Boolean

'###############'   CONVEX  HULL    '###############'
Private DoTestSequence  As Boolean
Private MAX_POINTS      As Long
Private PointsGenerated As Boolean
Private Points()        As TPoint2D

'###############'  ORDERED POLYGON  '###############'
Private Const MAX_POINTS_OP = 50
Private PointsOP() As TPoint2D

Private Sub Form_Initialize()
  'Call InitCommonControls
  MFastGeo.Initialize
  'Set FGDMainUnit.Icon = LoadResPicture(1, vbResIcon)
  cFunctionStrings = Array("Segment Intersections", "Random Segment Intersections", "Random Circle Intersections", "Segment Triangle Clipping", "Segment Quadix Clipping", "Segment Rectangle Clipping", "Segment Circle Clipping", "Circumcircle", "Inscribed Circle", "TorricelliPoint", "Closest Point On Triangle From Point", "Closest Point On Quadix From Point", "Closest Point On Circle From Point", "Closest Point On Rect From Point", "Closest Point On Circle From Segment", "Random Points Within Rectangle", "Random Points Within Triangle", "Random Points Within Quadix", "Random Point Within Circle", "Polygon Centroid", "Rotate Segments (animated)", "Rotate Triangles (animated)", "Rotate Quadix (animated)")
  Set BFGroup = New TRadioGroup
  'Call BFGroup.NewC(BFGroupOption)
  Set PaintBoxCanvas = NewTCanvas(PaintBox.HDC)
  If PaintBoxCanvas Is Nothing Then Exit Sub
'###############'  BASIC FUNCTIONS  '###############'
Dim i As Long, L As String
Dim eFuncIter As EFunctions
  
  For eFuncIter = LoweFunctions To HigheFunctions - 1
    BFGroup.Items.Add cFunctionStrings(eFuncIter)
  Next
  
'  If Assigned(FuncList) Then
'    Call Finalize(FuncList)
'  End If
  Set FuncList = New FGFunctionalities
  
  FuncList.SetLength (HigheFunctions + 1)
  
  For i = 0 To FuncList.Length - 1
    FuncList.Valid(i) = False
  Next
  '(* Register Demo Functions *)
  'wird in der Klasse FGFunctionalities gemacht
  Call FuncList.Create_FGFType(eSegmentIntersections, eSegmentIntersections, 0)
  Call FuncList.Create_FGFType(eRandomSegmentIntersections, eRandomSegmentIntersections, 0)
  Call FuncList.Create_FGFType(eRandomCircleIntersections, eRandomCircleIntersections, 0)
  Call FuncList.Create_FGFType(eSegmentTriangleClipping, eSegmentTriangleClipping, 0)
  Call FuncList.Create_FGFType(eSegmentQuadixClipping, eSegmentQuadixClipping, 0)
  Call FuncList.Create_FGFType(eSegmentRectangleClipping, eSegmentRectangleClipping, 0)
  Call FuncList.Create_FGFType(eSegmentCircleClipping, eSegmentCircleClipping, 0)
  Call FuncList.Create_FGFType(eCircumcircle, eCircumcircle, 0)
  Call FuncList.Create_FGFType(eInscribedCircle, eInscribedCircle, 0)
  Call FuncList.Create_FGFType(eTorricelliPoint, eTorricelliPoint, 0)
  Call FuncList.Create_FGFType(eRandPntsAABB, eRandPntsAABB, 0)
  Call FuncList.Create_FGFType(eRandPntsTriangle, eRandPntsTriangle, 0)
  Call FuncList.Create_FGFType(eRandPntsQuadix, eRandPntsQuadix, 0)
  Call FuncList.Create_FGFType(eRandPntsCircle, eRandPntsCircle, 0)
  'new demofunctions
  Call FuncList.Create_FGFType(eCPOTriangleFP, eCPOTriangleFP, 0)
  Call FuncList.Create_FGFType(eCPOQuadixFP, eCPOQuadixFP, 0)
  Call FuncList.Create_FGFType(eCPOAABBFP, eCPOAABBFP, 0)
  Call FuncList.Create_FGFType(eCPOCircleFP, eCPOCircleFP, 0)
  Call FuncList.Create_FGFType(eCPOCircleFS, eCPOCircleFS, 0)
  Call FuncList.Create_FGFType(ePolygonCentroid, ePolygonCentroid, 0)
  Call FuncList.Create_FGFType(eRotateSegments, eRotateSegments, 0)
  Call FuncList.Create_FGFType(eRotateTriangles, eRotateTriangles, 0)
  Call FuncList.Create_FGFType(eRotateQuadix, eRotateQuadix, 0)
  
  L = "FastGEO Library Information"
  L = L & vbCrLf & "Version: " & MFastGeo.VersionInformation
  L = L & vbCrLf & "Author :  " & MFastGeo.AuthorInformation
  L = L & vbCrLf & "Epoch :  " & MFastGeo.EpochInformation
  L = L & vbCrLf & "Most recent update:  " & MFastGeo.RecentUpdate
  L = L & vbCrLf & "Precision: " & MFastGeo.Precision
  InfoMemo.Text = L
  Call Clear
  InitBFGroup
  
'###############'   BARYCENTRIC     '###############'
  PrevCheckIn = False
  Call generate_triangle
  
'###############'   CONVEX  HULL    '###############'
  PointCountComboBox.ListIndex = 3
  CHStrategyComboBox.ListIndex = 0
  MAX_POINTS = CLng(PointCountComboBox.List(PointCountComboBox.ListIndex))
  PointsGenerated = False

'###############'  ORDERED POLYGON  '###############'
  'Set SurfaceCanvas = New TCanvas
  'SurfaceCanvas.Handle = Surface.hDC
  ReDim PointsOP(MAX_POINTS_OP - 1)

  'FraBasicFunctions.ZOrder 0
  Call mnuDemoBasicFunctions_Click
End Sub

Private Function NewTCanvas(hhdc As Long) As TCanvas
tryE: On Error GoTo ExceptE
  Set NewTCanvas = New TCanvas
  NewTCanvas.Handle = hhdc
  Exit Function
ExceptE:
  If Err.Number = 429 Then
    MsgBox "Es wird die MBODVBCanvas.dll benötigt, das Programm wird beendet."
    Unload Me
  Else
    MsgBox "ErrorNr: " & CStr(Err.Number) & " ist in Function Form1.NewTCanvas aufgetreten, mit der Meldung: " & vbCrLf & Err.Description
  End If
End Function

Private Sub Form_Load() 'FormCreate
  '
End Sub

Private Sub InitBFGroup()
Dim i As Long
  For i = 0 To 22 '18 '23
    BFGroupOption(i).Caption = cFunctionStrings(i)
  Next
  BFGroupOption(0).Value = True
End Sub

Private Sub Form_Resize()
Dim L As Long, T As Long, W As Long, H As Long
Dim brdr As Long
  brdr = 8 * Screen.TwipsPerPixelX
  L = PaintBox.Left
  T = PaintBox.Top
  W = Me.ScaleWidth - L - brdr
  H = Me.ScaleHeight - T - brdr
  If W <= 0 Then W = 1
  If H <= 0 Then H = 1
  PaintBox.Move L, T, W, H
End Sub

Private Sub BFGroupOption_Click(Index As Integer)
  BFGroup.ItemIndex = Index '+ 1
  ExecuteProctor
End Sub


Private Sub mnuDemoBasicFunctions_Click()
  PaintBoxCanvas.Pen.Width = 1
  PaintBoxCanvas.Pen.Color = clBlack
  FraBasicFunctions.Top = 120
  FraBasicFunctions.ZOrder 0
  TestWhat = etBasicFuncs
  Call PaintBox_Paint
End Sub
Private Sub mnuDemoBarycentric_Click()
  PaintBoxCanvas.Pen.Width = 1
  PaintBoxCanvas.Pen.Color = clBlack
  FraBaryCentric.Top = 120
  FraBaryCentric.ZOrder 0
  TestWhat = etBarycentric
  Call PaintBox_Paint
End Sub
Private Sub mnuDemoConvexHull_Click()
  PaintBoxCanvas.Pen.Width = 1
  PaintBoxCanvas.Pen.Color = clBlack
  FraConvexHull.Top = 120
  FraConvexHull.ZOrder 0
  TestWhat = etConvexHull
  Call PaintBox_Paint
End Sub
Private Sub mnuDemoOrderedPolygon_Click()
  PaintBoxCanvas.Pen.Width = 1
  PaintBoxCanvas.Pen.Color = clBlack
  FraOrderedPolygon.Top = 120
  FraOrderedPolygon.ZOrder 0
  TestWhat = etOrderedPoly
  Call PaintBox_Paint
End Sub

Private Sub mnuDemoExit_Click()
  Unload Me
End Sub

Private Sub mnuHelpInfo_Click()
  FAbout.Show vbModal
End Sub

Private Sub RefreshButton_Click()
  ExecuteProctor
End Sub

Private Sub ExecuteProctor() 'FncTyp As eFunctions)
Dim Graphic As New TFastGEOGraphics
Dim L As String
  'mLastFunc = FncTyp
  If BFGroup.ItemIndex > -1 Then
    Call Graphic.Create(PaintBoxCanvas, PaintBox.ScaleWidth, PaintBox.ScaleHeight)
tryE1: On Error GoTo finallyE1
    PaintBoxCanvas.Brush.Style = bsClear
    'StatusBar.SimpleText = 'Current View: ' + BFGroup.Items[BFGroup.ItemIndex]
    If FuncList.Valid(BFGroup.ItemIndex) Then
tryE2: On Error GoTo exceptE2
        Call FuncList(BFGroup.ItemIndex).proctor(Graphic)
exceptE2: GoTo endtryE2
        L = InfoMemo.Text: L = L & vbCrLf & "Error: " & BFGroup.Items(BFGroup.ItemIndex + 1): InfoMemo.Text = L
        Exit Sub
endtryE2: On Error GoTo finallyE1
     'Call InfoMemo.Lines.Add("Executed: " + BFGroup.Items(BFGroup.ItemIndex))
       'L = InfoMemo.Text: L = L & vbCrLf & "Executed: " & BFGroup.Items(BFGroup.ItemIndex): InfoMemo.Text = L
    Else
     'Call InfoMemo.Lines.Add("Error: " + BFGroup.Items(BFGroup.ItemIndex) + " was not registered.")
     L = InfoMemo.Text: L = L & vbCrLf & "Error: " & BFGroup.Items(BFGroup.ItemIndex + 1): InfoMemo.Text = L
    End If
finallyE1:
    Call Graphic.Free
  End If
End Sub

Private Sub Clear()
'Dim Graphic As New TFastGEOGraphics
'  Call Graphic.Create(PaintBoxCanvas, PaintBox.Width, PaintBox.Height)
'tryE: On Error GoTo finallyE
'  Call Graphic.Clear(&HFFFFFF)
'finallyE:
'  Graphic.Free

'oder_so:
  PaintBox.Cls
End Sub

Private Sub PaintBox_Paint()
  Select Case TestWhat
  Case etBasicFuncs
    Call Clear
    ExecuteProctor
  Case etBarycentric
    Call draw_triangle
  Case etConvexHull
    Call GenRandPntsBtn_Click
    Call DisplayCHullBtn_Click
  Case etOrderedPoly
    Call BtnGenRndPoints_Click
  End Select
End Sub

'###############'   BARYCENTRIC     '###############'
Public Sub generate_triangle()
Dim Tri As TTriangle2D
  Randomize
  Do
    Tri.P(0).X = Rnd * (PaintBox.ScaleWidth - 10) + 5# 'RandomS(Paint.ScaleWidth - 10) + 5#
    Tri.P(1).X = Rnd * (PaintBox.ScaleWidth - 10) + 5# 'RandomS(Paint.ScaleWidth - 10) + 5#
    Tri.P(2).X = Rnd * (PaintBox.ScaleWidth - 10) + 5# 'RandomS(Paint.ScaleWidth - 10) + 5#
    Tri.P(0).Y = Rnd * (PaintBox.ScaleHeight - 10) + 5# 'RandomS(Paint.ScaleHeight - 10) + 5#
    Tri.P(1).Y = Rnd * (PaintBox.ScaleHeight - 10) + 5# 'RandomS(Paint.ScaleHeight - 10) + 5#
    Tri.P(2).Y = Rnd * (PaintBox.ScaleHeight - 10) + 5# 'RandomS(Paint.ScaleHeight - 10) + 5#
  Loop Until Not IsDegenerate2DT(Tri)
  BU = CreateBarycentricUnit(Tri.P(0).X, Tri.P(0).Y, Tri.P(1).X, Tri.P(1).Y, Tri.P(2).X, Tri.P(2).Y)
End Sub

Public Sub draw_triangle()
Dim Graphic As TFastGEOGraphics
Dim pw As Long
  Set Graphic = New TFastGEOGraphics
  Call Graphic.Create(PaintBoxCanvas, PaintBox.ScaleWidth, PaintBox.ScaleHeight)
tryE: On Error GoTo finallyE
  Call Graphic.Clear(clWhite)
  'Graphic.Pen.Color = clRed
  pw = PaintBoxCanvas.Pen.Width
  PaintBoxCanvas.Pen.Width = 3
  Call Graphic.DrawTriangleXY(BU.X1, BU.Y1, BU.X2, BU.Y2, BU.X3, BU.Y3)
  PaintBoxCanvas.Pen.Width = pw
finallyE:
  Graphic.Free
EndTryE:
End Sub

Private Sub PaintBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not (TestWhat = etBarycentric) Then Exit Sub
Dim u As Double, v As Double, W As Double, bx As Double, by As Double
Dim Graphic As TFastGEOGraphics
Dim pw As Long
  
  Call ConvertCartesianToBarycentricSUXY(BU, CDbl(X), CDbl(Y), u, v, W)
  
  UEdit.Text = CStr(u)
  VEdit.Text = CStr(v)
  WEdit.Text = CStr(W)
  BaryTotalEdit.Text = CStr(u + v + W)
  
  MXEdit.Text = CStr(X)
  MYEdit.Text = CStr(Y)
  
  Call ConvertBarycentricToCartesianSXYU(u, v, W, BU, bx, by)
  
  BXEdit.Text = CStr(bx)
  BYEdit.Text = CStr(by)
  
  DeltaXEdit.Text = CStr(Abs(bx - X))
  DeltaYEdit.Text = CStr(Abs(by - Y))
  
  If GreaterThanOrEqual(u, 0#) And LessThanOrEqual(u, 1#) And _
     GreaterThanOrEqual(v, 0#) And LessThanOrEqual(v, 1#) And _
     GreaterThanOrEqual(W, 0#) And LessThanOrEqual(W, 1#) Then
    'StatusBar1.SimpleText = "Point In Triangle"
    If Not PrevCheckIn Then
      PrevCheckIn = True
      Set Graphic = New TFastGEOGraphics
      Call Graphic.Create(PaintBoxCanvas, PaintBox.ScaleWidth, PaintBox.ScaleHeight)
tryE: On Error GoTo finallyE
      Call Graphic.Clear(clWhite)
      Graphic.Pen.Color = clLime
      pw = Graphic.Pen.Width
      Graphic.Pen.Width = 3
      Call Graphic.DrawTriangleXY(BU.X1, BU.Y1, BU.X2, BU.Y2, BU.X3, BU.Y3)
      Graphic.Pen.Width = pw
finallyE:
      Graphic.Free
EndTryE:
    End If
  Else
    'StatusBar1.SimpleText = "Point Outside of Triangle"
    If PrevCheckIn Then
      PrevCheckIn = False
      Set Graphic = New TFastGEOGraphics
      Call Graphic.Create(PaintBoxCanvas, PaintBox.ScaleWidth, PaintBox.ScaleHeight)
try2: On Error GoTo finally2
      Call Graphic.Clear(clWhite)
      Graphic.Pen.Color = clRed
      pw = Graphic.Pen.Width
      Graphic.Pen.Width = 3
      Call Graphic.DrawTriangleXY(BU.X1, BU.Y1, BU.X2, BU.Y2, BU.X3, BU.Y3)
      Graphic.Pen.Width = pw
finally2:
      Graphic.Free
EndTry2:
    End If
  End If
End Sub

Public Sub ReGenButton_Click()
  Call generate_triangle
  Call draw_triangle
End Sub



'###############'   CONVEX  HULL    '###############'
Public Sub DrawPoints(Optional ForceDraw As Boolean = False)
Dim i As Long, FGr As New TFastGEOGraphics
  If (Not ForceDraw) And (Not (DrawCheckBox.Value = 1)) Then Exit Sub
  PaintBoxCanvas.Brush.Style = bsSolid
  'PaintBoxCanvas.Brush.Color = clWhite
  'Call PaintBoxCanvas.FillRect(0, 0, PaintBox.ScaleWidth, PaintBox.ScaleHeight)
  PaintBox.Cls
  PaintBoxCanvas.Brush.Color = clBlack
  PaintBoxCanvas.Pen.Color = clRed
  PaintBoxCanvas.Pen.Width = 8
  Call FGr.Create(PaintBoxCanvas, PaintBox.ScaleWidth, PaintBox.ScaleHeight)
  For i = 0 To MAX_POINTS - 1
    Call FGr.DrawPoint(Points(i).X, Points(i).Y)
  Next
End Sub

Public Sub SetPoints()
'Dim i As Long
'Dim mess As String
  Call GenerateRandomPoints2DXY(5, 5, PaintBox.ScaleWidth - 10, PaintBox.ScaleHeight - 10, Points)
  'For i = 0 To UBound(Points.Arr)
  '  mess = mess & "P" & CStr(i) & "(x: " & CStr(Points.Arr(i).x) & " / y: " & CStr(Points.Arr(i).y) & ")" & vbCrLf
  'Next
  'MsgBox mess
End Sub

Private Sub GenRandPntsBtn_Click()
  'Set Points = Nothing
  Erase Points
  ReDim Points(0 To MAX_POINTS - 1)
  Call SetPoints
  Call DrawPoints
  PointsGenerated = True
End Sub

Private Sub DisplayCHullBtn_Click()
Dim CHull As TPolygon2D
  If Not PointsGenerated Then Exit Sub

tryE:
  CHull = ConvexHull2D(Points)
  Call DrawHull(CHull)
   
  If PostScriptCheckBox.Value Then
    Call ProducePostScript(CHull, Points)
  End If
  
  If StdOutputCheckBox.Value Then
    Call ProduceStdOut(CHull, Points)
  End If
  
finallyE:
  Erase CHull.Arr
endTry:
End Sub


Private Sub PointCountComboBox_Click()
  MAX_POINTS = CLng(PointCountComboBox.List(PointCountComboBox.ListIndex))
  Call GenRandPntsBtn_Click
End Sub


Private Sub BeginTestButton_Click()
Const MAX_HULLS = 1000
Dim CHull As TPolygon2D, CHCount As Long
Dim StartTime As Long, EndTime As Long

  StartTime = GetTickCount
  DoTestSequence = True
  CHCount = 0
  ReDim Points(0 To MAX_POINTS - 1)
  
  Do While DoTestSequence
tryE: On Error GoTo finallyE
    Call SetPoints
    PaintBoxCanvas.Pen.Color = clLime
    PaintBoxCanvas.Pen.Width = 8
    PaintBoxCanvas.Brush.Style = bsClear
    Call DrawPoints
    
    CHull = ConvexHull2D(Points)
    PaintBoxCanvas.Pen.Color = clLime
    If VerifyHullIntegrity(CHull, Points) Then
      Call DrawHull(CHull)
    End If
finallyE:
    If UBound(CHull.Arr) Then
      'Erase CHull.Arr
      ReDim CHull.Arr(0)
    End If
endTry:
    Call Inc(CHCount)
    TestCounterEdit.Text = CStr(CHCount)
    If CHCount = MAX_HULLS Then Exit Do
    'Application.ProcessMessages
    DoEvents
  Loop
  EndTime = GetTickCount
  '(*
  '  MessageDlg(format('%d hulls in %dms - %5.3f hulls/per ms',[MAX_HULLS,Abs(EndTime - StartTime), Abs(EndTime - StartTime) / CHCount]), mtInformation, [mbOk], 0);
  Call MsgBox(CStr(CHCount) & " hulls in " & CStr(Abs(EndTime - StartTime)) & "ms - " & CStr(CHCount / Abs(EndTime - StartTime) * 1000) & " hulls/s", vbInformation Or vbOKOnly)
  '*)
End Sub

Private Sub StopTestButton_Click()
  DoTestSequence = False
End Sub

Friend Sub ProducePostScript(CHull As TPolygon2D, Points() As TPoint2D)
Dim OutPut As TStringList
Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double, i As Long
  Set OutPut = New TStringList
  xmin = Points(0).X
  xmax = Points(0).X
  ymin = Points(0).Y
  ymax = Points(0).Y
  For i = 1 To UBound(Points) '- 1
    If Points(i).X > xmax Then
      xmax = Points(i).X
    ElseIf Points(i).X < xmin Then
      xmin = Points(i).X
    End If
    If Points(i).Y > ymax Then
      ymax = Points(i).Y
    ElseIf Points(i).Y < ymin Then
      ymin = Points(i).Y
    End If
  Next
  
  Call OutPut.Add("%!PS")
  Call OutPut.Add("%%Creator: FastGEO Convex Hull (Arash Partow)")
  Call OutPut.Add("%%BoundingBox: " & CStr(xmin) & " " & CStr(ymin) & " " & CStr(xmax) & " " & CStr(ymax))
  Call OutPut.Add("%%EndComments")
  Call OutPut.Add(".00 .00 setlinewidth")
  Call OutPut.Add("newpath")
  Call OutPut.Add("%%Points:")
  
  For i = 0 To UBound(Points) '- 1
    Call OutPut.Add(CStr(CLng(Points(i).X)) & " " & CStr(CLng(Points(i).Y)) & " 3  0  360 arc stroke")
  Next
  
  Call OutPut.Add("closepath")
  
  Call OutPut.Add("%%Hull:")
  Call OutPut.Add("newpath")
  
  Call OutPut.Add(CStr(CLng(CHull.Arr(0).X)) & " " & CStr(CLng(CHull.Arr(0).Y)) & " moveto")
  
  For i = 1 To UBound(CHull.Arr) '- 1
    Call OutPut.Add(CStr(CLng(CHull.Arr(i).X)) & " " & CStr(CLng(CHull.Arr(i).Y)) & " lineto")
  Next
  
  Call OutPut.Add("closepath stroke")
  
  Call OutPut.Add("showpage")
  Call OutPut.Add("%%EOF")
  
  Call OutPut.SaveToFile("fastgeo-convexhull" & CStr(GetTickCount) & ".eps")
End Sub


Friend Sub ProduceStdOut(CHull As TPolygon2D, Points() As TPoint2D)
Dim OutPut As TStringList, i As Long
  Set OutPut = New TStringList
  
  '(* Print out set of points *)
  Call OutPut.Add(CStr(UBound(Points)))
  For i = 0 To UBound(Points) '- 1
    Call OutPut.Add(CStr(Points(i).X) & " " & CStr(Points(i).Y))
  Next
  
  '(* Print out convex hull *)
  OutPut.Add (CStr(UBound(CHull.Arr)))
  For i = 0 To UBound(CHull.Arr) '- 1
    Call OutPut.Add(CStr(CHull.Arr(i).X) & " " & CStr(CHull.Arr(i).Y))
  Next
  Call OutPut.SaveToFile("fastgeo-convexhull" & CStr(GetTickCount) & ".stdout")
End Sub

Friend Sub DrawHull(CHull As TPolygon2D, Optional ForceDraw As Boolean = False)
Dim FGr As New TFastGEOGraphics
  If (Not ForceDraw) And (Not (DrawCheckBox.Value = 1)) Then Exit Sub
  Call FGr.Create(PaintBoxCanvas)
tryE: On Error GoTo finallyE
  FGr.Pen.Width = 4
  FGr.Pen.Color = clBlue
  Call FGr.DrawPg(CHull)
finallyE:
  FGr.Free
endTry:
End Sub
Friend Sub DrawErrorPoint(Pnt As TPoint2D)
Dim FGr As New TFastGEOGraphics
Dim pw As Long
  Call FGr.Create(PaintBoxCanvas)
tryE: On Error GoTo finallyE
  FGr.Pen.Width = 4
  FGr.Brush.Style = bsClear
  Call FGr.DrawCirclePR(Pnt, 10)
finallyE:
  FGr.Free
endTry:
End Sub

Friend Function VerifyHullIntegrity(CHull As TPolygon2D, Pnts() As TPoint2D) As Boolean
Dim i As Long, j As Long, PointApartOfConvexHull As Boolean
Dim MinDist As Double, PointList As TStringList 's
  VerifyHullIntegrity = False
  '(* Verify Hull is really convex *)
  If Not ConvexPolygon(CHull) Then
    Call DrawHull(CHull, True)
    Call Sleep(10) '00)
    Set PointList = New TStringList
    For i = 0 To UBound(Points)
      PointList.Add ("Points[" & CStr(i) & "].x := " & CStr(Points(i).X) & ";   Points[" & CStr(i) & "].y := " & CStr(Points(i).Y) & ";")
    Next
    Call PointList.SaveToFile("error-" & CStr(GetTickCount) & ".txt")
    Call PointList.Free
    Exit Function
  End If
  
  For i = 0 To UBound(Pnts)
    PointApartOfConvexHull = False
    For j = 0 To UBound(CHull.Arr)
      If IsEqual2DP(CHull.Arr(j), Pnts(i)) Then
        PointApartOfConvexHull = True 'yes Point(i) is part of the Hull
        Exit For
      End If
    Next
    If Not PointApartOfConvexHull Then
      If Not PointInConvexPolygon(Pnts(i), CHull) Then
        If Not PointInPolygon(Pnts(i), CHull) Then
          If Not PointOnPolygon(Pnts(i), CHull) Then
            MinDist = MinimumDistanceFromPointToPolygon(Pnts(i), CHull)
            If MinDist > 0.9 Then
              Call DrawHull(CHull, True)
              Call DrawErrorPoint(Pnts(i))
              'PointList = TStringList.Create
              Set PointList = New TStringList
              Call PointList.Free
              For j = 0 To UBound(Pnts) '- 1
                Call PointList.Add("Points[" & CStr(j) & "].x := " & CStr(Pnts(j).X) & ";   Points[" & CStr(j) & "].y := " & CStr(Pnts(j).Y) & ";")
              Next
              Call PointList.SaveToFile("error-II-" & CStr(GetTickCount) & ".txt")
              Call PointList.Free
              Exit Function
            End If
          End If
        End If
      End If
    End If
  Next
  VerifyHullIntegrity = True
End Function

'###############'  ORDERED POLYGON  '###############'
Private Sub DrawPointsOP()
Dim i As Long, FGr As TFastGEOGraphics
  PaintBoxCanvas.Brush.Style = bsSolid
  PaintBoxCanvas.Brush.Color = clWhite
  Call PaintBoxCanvas.FillRect(0, 0, PaintBox.ScaleWidth, PaintBox.ScaleHeight)
  PaintBoxCanvas.Brush.Color = clBlack
  PaintBoxCanvas.Pen.Color = clBlack
  PaintBoxCanvas.Pen.Width = 3
  Set FGr = New TFastGEOGraphics
  Call FGr.Create(PaintBoxCanvas)
  For i = 0 To MAX_POINTS_OP - 1
    Call FGr.DrawPointP(PointsOP(i))
  Next
End Sub

Private Sub DisplayEdges()
Dim i As Long, FGr As TFastGEOGraphics
  PaintBoxCanvas.Brush.Style = bsSolid
  PaintBoxCanvas.Brush.Color = clWhite
  Call PaintBoxCanvas.FillRect(0, 0, PaintBox.ScaleWidth, PaintBox.ScaleHeight)
  PaintBoxCanvas.Brush.Color = clBlack
  PaintBoxCanvas.Pen.Color = clBlack
  PaintBoxCanvas.Pen.Width = 2
  Set FGr = New TFastGEOGraphics
  Call FGr.Create(PaintBoxCanvas, PaintBox.ScaleWidth, PaintBox.ScaleHeight)
  For i = 0 To MAX_POINTS_OP - 1
    Call FGr.DrawLineP(PointsOP(i), PointsOP((i + 1) Mod MAX_POINTS_OP))
  Next
End Sub

Private Sub BtnGenRndPoints_Click()
  Call GenerateRandomPoints2DXY(5, 5, PaintBox.ScaleWidth - 10, PaintBox.ScaleHeight - 10, PointsOP)
  Call DrawPointsOP
End Sub

Private Sub DispEdgeBtn_Click()
  Call DisplayEdges
End Sub

Private Sub OrdPolyBtn_Click()
'Dim OrdPoly As TOrderedPolygon2D
Dim i As Long, j As Long, Rec As TRectangle
Dim PointPg As TPolygon2D
  If UBound(PointsOP) = 0 Then Exit Sub
  'Set OrdPoly = New TOrderedPolygon2D
  'Call OrdPoly.
  Call OrderedPolygon2D(PointsOP)
  Call DisplayEdges
  Rec = RectangularHullA(PointsOP)
  PointPg = CPointsToPolygon2D(PointsOP)
  For i = 0 To PaintBox.ScaleWidth - 1
    For j = 0 To PaintBox.ScaleHeight - 1
      If PointInRectangleXYR(CSng(i), CSng(j), Rec) Then
        If PointInPolygonXY(CSng(i), CSng(j), PointPg) Then
          'this is a kind'a strange fillroutine 'voll krass eyh
          'but it works and demonstrates once more the capabilities of FastGeo
          PaintBoxCanvas.Pixels(i, j) = clBlack
        End If
      End If
      If ConvexPolygon(PointPg) Then
        MsgBox "I think there is a big problem!!!", vbOKOnly Or vbInformation
      End If
    Next j
  Next i
End Sub

Private Function CPointsToPolygon2D(Points() As TPoint2D) As TPolygon2D
  'zunächst nur das gesamte Array kopieren, mit "="
  'OK, da wird jetzt der gesamte Speicher kopiert
  'man könnte das auch noch schneller machen, indem man nur den
  'SafeArrayDescriptor kopiert
  CPointsToPolygon2D.Arr = Points
  'das ergibt zwei Möglichkeiten:
  '1. CopyPointsToPolygon2D
  '2. ein Typ-Cast CPointsToPolygon2D
End Function


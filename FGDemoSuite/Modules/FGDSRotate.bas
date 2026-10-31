Attribute VB_Name = "FGDSRotate"
Option Explicit
'Double Precision!
Private Const MAX_POINTS    As Long = 4
Private Const MAX_SEGMENTS  As Long = 40
Private Const MAX_TRIANGLES As Long = 15
Private Const MAX_QUADICES  As Long = 10
Private Const ROT_ANGLE     As Long = 8     'Rotation angle in degree
Private Const TIM_SLEEP     As Long = 40    'Sleeping time in milliseconds
Private Const ROT_COUNT     As Long = 46    'how much rotations

Public Sub PolygonCenter(Graphic As TFastGEOGraphics)
Dim Pg As TPolygon2D
Dim CP As TPoint2D
Dim TmpPC As TColor
Dim P As Long
  Call Graphic.Clear(clWhite)
  TmpPC = Graphic.Pen.Color
  ReDim Pg.Arr(0 To MAX_POINTS - 1)
  Call GenerateRandomPolygon(5, 5, Graphic.Width - 10, Graphic.Height - 10, Pg)
  'if it is selfintersecting then it could be false
  CP = Centroid2DPg(Pg)
  Graphic.Pen.Color = clBlack
  P = VarPtr(Pg.Arr(0))
  Call Graphic.DrawPg(Pg)
  Graphic.Pen.Color = clLime
  Call Graphic.DrawCirclePR(CP, 4)
  Graphic.Pen.Color = TmpPC
End Sub

Public Sub RotateSegments(Graphic As TFastGEOGraphics)
Dim Segs As TSegment2DArray
Dim Center As TPoint2D
Dim i As Long
Dim T As Long
Dim r As Long
Dim brdr As Long

  ReDim Segs.Arr(0 To MAX_SEGMENTS - 1)
  Call Graphic.Clear(clWhite)
  
  For i = 0 To MAX_SEGMENTS - 1
    Segs.Arr(i).P(0).X = RandomL(Graphic.Width - 10) + 5#
    Segs.Arr(i).P(0).Y = RandomL(Graphic.Height - 10) + 5#
    Segs.Arr(i).P(1).X = RandomL(Graphic.Width - 10) + 5#
    Segs.Arr(i).P(1).Y = RandomL(Graphic.Height - 10) + 5#
  Next
  brdr = Graphic.Width / 5
  Center.X = RandomL(Graphic.Width - 2 * brdr) + brdr
  Center.Y = RandomL(Graphic.Height - 2 * brdr) + brdr
  For r = 1 To ROT_COUNT
    'bestehndes Bild löschen
    Call Graphic.Clear(clWhite)
    'zeichnen
    Graphic.Pen.Color = clRed
    Call Graphic.DrawCirclePR(Center, 4)
    Graphic.Pen.Color = clBlack
    For i = 0 To MAX_SEGMENTS - 1
      Call Graphic.DrawS(Segs.Arr(i))
    Next
    'drehen
    For i = 0 To MAX_SEGMENTS - 1
      Segs.Arr(i) = Rotate2DSo(ROT_ANGLE, Segs.Arr(i), Center)
    Next
    'bissl warten
    Sleep TIM_SLEEP
    DoEvents
  Next
End Sub

Public Sub RotateTriangles(Graphic As TFastGEOGraphics)
Dim Tris As TTriangle2DArray
Dim Center As TPoint2D
Dim i As Long
'Dim t As Long
Dim r As Long
Dim brdr As Long
  ReDim Tris.Arr(0 To MAX_TRIANGLES - 1)
  Call Graphic.Clear(clWhite)
  
  Call GenerateRandomTriangles(Graphic.Width, Graphic.Height, Tris)
  
  brdr = Graphic.Width / 5
  Center.X = RandomL(Graphic.Width - 2 * brdr) + brdr
  Center.Y = RandomL(Graphic.Height - 2 * brdr) + brdr
  For r = 1 To ROT_COUNT
    'bestehndes Bild löschen
    Call Graphic.Clear(clWhite)
    'zeichnen
    Graphic.Pen.Color = clRed
    Call Graphic.DrawCirclePR(Center, 4)
    Graphic.Pen.Color = clBlack
    For i = 0 To MAX_TRIANGLES - 1
      Call Graphic.DrawT(Tris.Arr(i))
    Next
    'drehen
    For i = 0 To MAX_TRIANGLES - 1
      Tris.Arr(i) = Rotate2DTo(ROT_ANGLE, Tris.Arr(i), Center)
    Next
    'bissl warten
    Sleep TIM_SLEEP
    DoEvents
  Next

End Sub

Public Sub RotateQuadices(Graphic As TFastGEOGraphics)
Dim Quads As TQuadix2DArray
Dim Center As TPoint2D
Dim i As Long
Dim T As Long
Dim r As Long
Dim brdr As Long
  ReDim Quads.Arr(0 To MAX_QUADICES - 1)
  Call Graphic.Clear(clWhite)
  For i = 1 To UBound(Quads.Arr)
    Call GenerateRandomQuadix(5, 5, Graphic.Width - 10, Graphic.Height - 10, Quads.Arr(i))
  Next
  brdr = Graphic.Width \ 5
  Center.X = RandomL(Graphic.Width - 2 * brdr) + brdr
  Center.Y = RandomL(Graphic.Height - 2 * brdr) + brdr
  Graphic.Pen.Mode = pmCopy
  For r = 1 To ROT_COUNT
    'bestehndes Bild löschen
    Call Graphic.Clear(clWhite)
    'zeichnen
    Graphic.Pen.Color = clRed
    Call Graphic.DrawCirclePR(Center, 4)
    Graphic.Pen.Color = clBlack
    For i = 0 To MAX_QUADICES - 1
      Call Graphic.DrawQ(Quads.Arr(i))
    Next
    For i = 0 To MAX_QUADICES - 1
      Quads.Arr(i) = Rotate2DQo(ROT_ANGLE, Quads.Arr(i), Center)
    Next
    'bissl warten
    Sleep TIM_SLEEP
    DoEvents
  Next
End Sub



VERSION 5.00
Begin VB.UserControl ctlPit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   CanGetFocus     =   0   'False
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   FillStyle       =   0  'Solid
   ScaleHeight     =   178
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   Begin VB.Timer tmrPlay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   180
      Top             =   2220
   End
End
Attribute VB_Name = "ctlPit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private vPaintColor As OLE_COLOR
Private vBrushIndex As Integer

Public Event Track(X As Integer, Y As Integer)

Public Sub Play()
Dim Result As Long

  With FPit
    .PaintPicture FMain.pctPit(Index).Picture, 0, 0
    .X = .Left + PitInfo(Index).Shift.X - PitParent.Left
    .Y = .Top + PitInfo(Index).Shift.Y - PitParent.Top
    .OldX = .X
    .OldY = .Y
    MMain.ChangeColor = RGB(255, 0, 0)
    UserControl.FillColor = RGB(255, 0, 0)
    Result = ExtFloodFill(UserControl.hdc, .X / Screen.TwipsPerPixelX, .Y / Screen.TwipsPerPixelY, RGB(255, 255, 255), 1)
  End With
  tmrPlay.Enabled = True
End Sub

Public Sub StopPlaying()
Dim Result As Long

  tmrPlay.Enabled = False
  DoEvents
  UserControl.FillColor = RGB(255, 255, 255)
  Result = ExtFloodFill(UserControl.hdc, FPit.OldX / Screen.TwipsPerPixelX, FPit.OldY / Screen.TwipsPerPixelY, RGB(255, 0, 0), 1)
  RaiseEvent Track(FPit.X, FPit.Y)
End Sub

Public Property Set DrawingArea(ByVal vNewValue As Object)
  Set MMain.DrawingArea = vNewValue
  With UserControl
    .Width = MMain.DrawingArea.Image.Width
    .Height = MMain.DrawingArea.Image.Height
    .Picture = MMain.DrawingArea.Image
  End With
End Property

Public Sub Show(Parent As Object)
  Set PitParent = Parent
  CreatePit
  Set RunMe = Extender
  If Not FPit.Visible Then FPit.Show , PitParent
End Sub

Private Sub tmrPlay_Timer()
Dim Result

  Index = Index + dX
  CreatePit
  
  With FPit
    .Left = .Left + dX * 120
    .Top = .Top - dX * 60
    .X = .Left + PitInfo(Index).Shift.X - PitParent.Left
    .Y = .Top + PitInfo(Index).Shift.Y - PitParent.Top
    
    Paint
  End With
  If Not FPit.Visible Then FPit.Show , PitParent
  
  DoEvents
  If Index = 0 Or Index = 3 Then
    dX = -dX
  End If
End Sub

Private Sub Paint()
Dim I As Integer
Dim J As Integer
Dim X As Integer
Dim Y As Integer
Dim Color As Long

  With BrushInfo(vBrushIndex)
    For I = 0 To .Size
      If Rnd <= .Density Then
        X = FPit.X / Screen.TwipsPerPixelX + .Map(I).X
        Y = FPit.Y / Screen.TwipsPerPixelY + .Map(I).Y
        Color = UserControl.POINT(X, Y)
        If Color = ChangeColor And Color <> 1 Then
''          MMain.DrawingArea.PSet (X, Y), vPaintColor
          MMain.DrawingArea.PSet (X, Y), CombineColor(vPaintColor, MMain.DrawingArea.POINT(X, Y))
        End If
      End If
    Next I
  End With
End Sub

Public Property Get Color() As OLE_COLOR
  Color = vPaintColor
End Property

Public Property Let Color(ByVal vNewValue As OLE_COLOR)
  vPaintColor = vNewValue
  CreatePit
  PropertyChanged "Color"
End Property

Public Property Get BrushCount() As Integer
  BrushCount = FMain.pctBrushMap.Count
End Property

Public Property Get BrushImage(Index As Integer) As IPictureDisp
  If Index > 0 And Index <= FMain.pctBrushMap.Count Then
    Set BrushImage = FMain.pctBrushMap(Index - 1)
  End If
End Property

Private Sub UserControl_Initialize()
  vPaintColor = RGB(255, 255, 255)
  vBrushIndex = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  vPaintColor = PropBag.ReadProperty("Color")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Color", vPaintColor
End Sub

Private Sub CreatePit()
Dim Result As Long

  MakeForm FPit, PitInfo(Index).Region
  With FPit
    .PaintPicture FMain.pctPit(Index).Picture, 0, 0
    .FillColor = vPaintColor
    Result = ExtFloodFill(.hdc, PitInfo(Index).Paint.X, PitInfo(Index).Paint.Y, RGB(255, 255, 255), 1)
  End With
End Sub

Public Property Get BrushIndex() As Integer
  BrushIndex = vBrushIndex + 1
End Property

Public Property Let BrushIndex(ByVal vNewValue As Integer)
  If vNewValue > 0 And vNewValue <= FMain.pctBrushMap.Count Then
    vBrushIndex = vNewValue - 1
  End If
End Property

Public Property Get BrushDense() As Single
  BrushDense = BrushInfo(vBrushIndex).Density
End Property

Public Property Let BrushDense(ByVal vNewValue As Single)
  BrushInfo(vBrushIndex).Density = vNewValue
End Property


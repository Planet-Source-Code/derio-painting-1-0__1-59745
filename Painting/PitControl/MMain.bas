Attribute VB_Name = "MMain"
Option Explicit

Private Declare Function CreateRectRgn _
        Lib "gdi32" (ByVal X1 As Long, _
                     ByVal Y1 As Long, _
                     ByVal X2 As Long, _
                     ByVal Y2 As Long) As Long
                     
Private Declare Function CombineRgn _
        Lib "gdi32" (ByVal hDestRgn As Long, _
                     ByVal hSrcRgn1 As Long, _
                     ByVal hSrcRgn2 As Long, _
                     ByVal nCombineMode As Long) As Long
                     
Private Declare Function DeleteObject _
        Lib "gdi32" (ByVal hObject As Long) As Long
        
Private Declare Function SetWindowRgn _
        Lib "user32" (ByVal hwnd As Long, _
                      ByVal hRgn As Long, _
                      ByVal bRedraw As Boolean) As Long

Private Declare Function SetPixel _
        Lib "gdi32" (ByVal hdc As Long, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal crColor As Long) As Long
                     
Private Declare Function GetPixel _
        Lib "gdi32" (ByVal hdc As Long, _
                     ByVal X As Integer, _
                     ByVal Y As Integer) As Long

Public Declare Function ReleaseCapture _
       Lib "user32" () As Long
       
Public Declare Function SendMessage _
       Lib "user32" _
       Alias "SendMessageA" (ByVal hwnd As Long, _
                             ByVal wMsg As Long, _
                             ByVal wParam As Long, _
                             lParam As Any) As Long

Public Declare Function ExtFloodFill _
       Lib "gdi32" _
       (ByVal hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal crColor As Long, _
        ByVal wFillType As Long) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCRBUTTONDOWN = &HA4

Private Const RGN_OR = 2

Public Type POINT
  X As Integer
  Y As Integer
End Type

Public PitWidth As Integer
Public PitHeight As Integer

Public Type PIT
  Region As Long
  Shift As POINT
  Paint As POINT
End Type

Public PitInfo(3) As PIT
Public PitParent As Form
Public DrawingArea As Object
Public ChangeColor As Long
Public Index As Integer
Public dX As Integer
Public RunMe As Object

Public Type BRUSH
  Density As Single
  Size As Integer
  Map() As POINT
End Type

Public BrushInfo() As BRUSH

Public Sub Main()
  InitPit
  InitBrush
End Sub

Private Sub InitPit()
Dim I As Integer

  PitWidth = 102
  PitHeight = 123
  
  For I = 0 To FMain.pctPit.Count - 1
    PitInfo(I).Region = CreateRegion(FMain.pctPit(I), RGB(0, 0, 255))
  Next I
  With PitInfo(0)
    .Shift.X = 40 * Screen.TwipsPerPixelX
    .Shift.Y = 85 * Screen.TwipsPerPixelY
    .Paint.X = 43
    .Paint.Y = 110
  End With
  
  With PitInfo(1)
    .Shift.X = 22 * Screen.TwipsPerPixelX
    .Shift.Y = 78 * Screen.TwipsPerPixelY
    .Paint.X = 32
    .Paint.Y = 103
  End With
  
  With PitInfo(2)
    .Shift.X = 12 * Screen.TwipsPerPixelX
    .Shift.Y = 70 * Screen.TwipsPerPixelY
    .Paint.X = 20
    .Paint.Y = 99
  End With
  
  With PitInfo(3)
    .Shift.X = 2 * Screen.TwipsPerPixelX
    .Shift.Y = 60 * Screen.TwipsPerPixelY
    .Paint.X = 14
    .Paint.Y = 90
  End With
  Index = 0
  dX = 1
End Sub

Private Sub InitBrush()
Dim I As Integer
Dim X As Integer
Dim Y As Integer
Dim MapIndex As Integer

  ReDim BrushInfo(FMain.pctBrushMap.Count - 1)
  For I = 0 To FMain.pctBrushMap.Count - 1
    With BrushInfo(I)
      .Density = 0.25
      ReDim .Map(FMain.pctBrushMap(I).ScaleWidth * FMain.pctBrushMap(I).ScaleHeight)
      MapIndex = -1
      For X = 1 To FMain.pctBrushMap(I).ScaleWidth
        For Y = 1 To FMain.pctBrushMap(I).ScaleHeight
          If FMain.pctBrushMap(I).POINT(X, Y) = 0 Then
            MapIndex = MapIndex + 1
            .Map(MapIndex).X = X - FMain.pctBrushMap(I).ScaleWidth \ 2
            .Map(MapIndex).Y = Y - FMain.pctBrushMap(I).ScaleHeight \ 2
          End If
        Next Y
      Next X
      .Size = MapIndex
      ReDim Preserve .Map(MapIndex)
    End With
  Next I
End Sub

Public Function CreateRegion(PctSource As PictureBox, _
                             Color As Long) As Long
'** creating region base on the image

Dim X As Long
Dim Y As Long
Dim X1 As Long
Dim Y1 As Long
Dim X2 As Long
Dim Y2 As Long
Dim hRgnTemp As Long
Dim Result As Long
Dim AddToRegion As Boolean
Dim hRgn As Long

  AddToRegion = False
  For X = 0 To PctSource.Width
    AddToRegion = False
    For Y = 0 To PctSource.Height
      If AddToRegion Then
        If GetPixel(PctSource.hdc, X, Y) = Color Then
          'enlarge the area (X1,Y1) - (X2,Y2)
          'with new (X2,Y2) for the new region
          X2 = X
          Y2 = Y
          
          If hRgn = 0 Then
            'define the region
            hRgn = CreateRectRgn(X1, Y1, X2 + 1, Y2)
            
          Else
            'add the new one
            hRgnTemp = CreateRectRgn(X1, Y1, X2 + 1, Y2)
            Result = CombineRgn(hRgn, hRgn, hRgnTemp, RGN_OR)
            DeleteObject hRgnTemp
          End If
          
          AddToRegion = False
        End If
        
      Else
        If GetPixel(PctSource.hdc, X, Y) <> Color Then
          'initailize area (X1,Y1) - (X2,Y2) for new region
          X1 = X
          Y1 = Y
          X2 = X
          Y2 = Y
          AddToRegion = True
        End If
      End If
    Next
  Next
  
  CreateRegion = hRgn
End Function

Public Function MakeForm(PitItem As Object, _
                         Region As Long)
'** create the form area base on the region
Dim Result As Long
Dim hRgn As Long

  hRgn = CreateRectRgn(PitWidth \ 2, _
                       PitHeight \ 2, _
                       PitWidth \ 2 + 1, _
                       PitHeight \ 2 + 1)
                       
  'define the maximum size of the form
  PitItem.Width = PitWidth * Screen.TwipsPerPixelX
  PitItem.Height = PitHeight * Screen.TwipsPerPixelY
  DoEvents
  
  'define the region for the area form
  Result = CombineRgn(hRgn, hRgn, Region, RGN_OR)
  Result = SetWindowRgn(PitItem.hwnd, hRgn, True)
  DeleteObject hRgn
End Function

Public Function CombineColor(ForeColor As Long, BackColor As Long) As Long
Dim Red1 As Integer
Dim Green1 As Integer
Dim Blue1 As Integer
Dim Red2 As Integer
Dim Green2 As Integer
Dim Blue2 As Integer

  Red1 = ForeColor And RGB(255, 0, 0)
  Green1 = (ForeColor And RGB(0, 255, 0)) \ 256
  Blue1 = (ForeColor And RGB(0, 0, 255)) \ 256 \ 256

  Red2 = BackColor And RGB(255, 0, 0)
  Green2 = (BackColor And RGB(0, 255, 0)) \ 256
  Blue2 = (BackColor And RGB(0, 0, 255)) \ 256 \ 256

  CombineColor = RGB(((7 * Red1 + 3 * Red2) \ 10) And 255, _
                     ((7 * Green1 + 3 * Green2) \ 10) And 255, _
                     ((7 * Blue1 + 3 * Blue2) \ 10) And 255)

End Function

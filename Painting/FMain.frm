VERSION 5.00
Object = "*\APitControl\PitControl.vbp"
Begin VB.Form FMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   Caption         =   "Painting 1.0"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form2"
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctTools 
      Align           =   4  'Align Right
      Height          =   6435
      Left            =   8355
      ScaleHeight     =   6375
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   0
      Width           =   1275
      Begin VB.PictureBox pctToolOption 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   870
         Index           =   4
         Left            =   120
         Picture         =   "FMain.frx":0000
         ScaleHeight     =   810
         ScaleWidth      =   960
         TabIndex        =   7
         Top             =   3240
         Width           =   1020
      End
      Begin VB.PictureBox pctToolOption 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   870
         Index           =   3
         Left            =   120
         Picture         =   "FMain.frx":0742
         ScaleHeight     =   810
         ScaleWidth      =   960
         TabIndex        =   6
         Top             =   2280
         Width           =   1020
      End
      Begin VB.PictureBox pctToolOption 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   870
         Index           =   2
         Left            =   120
         Picture         =   "FMain.frx":0E84
         ScaleHeight     =   810
         ScaleWidth      =   960
         TabIndex        =   5
         Top             =   1080
         Width           =   1020
      End
      Begin VB.PictureBox pctToolOption 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   870
         Index           =   1
         Left            =   120
         Picture         =   "FMain.frx":15C6
         ScaleHeight     =   810
         ScaleWidth      =   960
         TabIndex        =   4
         Top             =   120
         Width           =   1020
      End
      Begin VB.PictureBox pctToolOption 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   870
         Index           =   0
         Left            =   120
         Picture         =   "FMain.frx":1D08
         ScaleHeight     =   810
         ScaleWidth      =   960
         TabIndex        =   3
         Top             =   4380
         Width           =   1020
      End
   End
   Begin VB.PictureBox pctColor 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   642
      TabIndex        =   1
      Top             =   6435
      Width           =   9630
      Begin VB.Shape shpColor 
         BackColor       =   &H00000001&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   0
         Left            =   120
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   360
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   2
         Left            =   600
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   3
         Left            =   840
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   4
         Left            =   1080
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   1320
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00008080&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   1560
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00004080&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   7
         Left            =   1800
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   8
         Left            =   2040
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   9
         Left            =   2280
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   10
         Left            =   2520
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   11
         Left            =   2760
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   12
         Left            =   3000
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   13
         Left            =   3240
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00FF00FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   14
         Left            =   3480
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   15
         Left            =   3720
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   16
         Left            =   3960
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   17
         Left            =   4200
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   18
         Left            =   4440
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   19
         Left            =   4680
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   20
         Left            =   4920
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   21
         Left            =   5160
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   22
         Left            =   5400
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   23
         Left            =   5640
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.PictureBox pctTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   10440
      Picture         =   "FMain.frx":244A
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   5820
      Visible         =   0   'False
      Width           =   9600
   End
   Begin PitControl.ctlPit ctlPit 
      Height          =   3615
      Left            =   2220
      Top             =   1320
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6376
      Color           =   32768
   End
   Begin VB.Image imgDense 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "FMain.frx":E348C
      Tag             =   "0.25"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDense 
      Height          =   480
      Index           =   1
      Left            =   660
      Picture         =   "FMain.frx":E370E
      Tag             =   "0.5"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDense 
      Height          =   480
      Index           =   2
      Left            =   1200
      Picture         =   "FMain.frx":E3990
      Tag             =   "0.75"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDense 
      Height          =   480
      Index           =   3
      Left            =   1740
      Picture         =   "FMain.frx":E3C12
      Tag             =   "1"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTemp 
      Height          =   375
      Left            =   720
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ExtFloodFill _
       Lib "gdi32" _
       (ByVal hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal crColor As Long, _
        ByVal wFillType As Long) As Long

Dim ButtonColor As Long

Private Sub ctlPit_Track(X As Integer, Y As Integer)
''Dim Color As Long
''
''  If Y / Screen.TwipsPerPixelY >= Me.pctColor.Top Then
''    Color = pctColor.Point(X, Y)
''    If Color <> ButtonColor Then
''      ctlPit.Color = Color
''    End If
''  End If
End Sub

Private Sub Form_Load()
  Picture = pctTemp.Picture
  Show
  LoadColor
  Me.DrawMode = 13
  Set ctlPit.DrawingArea = Me
  ctlPit.Show Me
End Sub


Private Sub LoadColor()
Dim I As Integer
Dim J As Integer
Dim Result As Long

  ButtonColor = pctColor.Point(1, 1)
  For I = 1 To shpColor.Count
    J = Int(Rnd * 5) + 1
    imgTemp = LoadPicture(App.Path & "\Paint" & Format(J, "00") & ".BMP")
    pctColor.PaintPicture imgTemp, (I - 1) * 26 + 1, 5, opcode:=vbSrcAnd
    With pctColor
      .FillColor = RGB(1, 1, 1)
    End With
    Result = ExtFloodFill(pctColor.hdc, (I - 1) * 26 + 12, 12, QBColor(0), 1)
    
    pctColor.PaintPicture imgTemp, (I - 1) * 26, 3, opcode:=vbSrcAnd
    With pctColor
      If I = 1 Then
        .FillColor = RGB(1, 1, 1)
      Else
        .FillColor = shpColor(I - 1).BackColor
      End If
    End With
    Result = ExtFloodFill(pctColor.hdc, (I - 1) * 26 + 12, 12, QBColor(0), 1)
  Next I
End Sub

Private Sub pctColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Color As Long

  Color = pctColor.Point(X, Y)
  If Color <> ButtonColor Then
    ctlPit.Color = Color
  End If
End Sub

Private Sub pctToolOption_Click(Index As Integer)
  Select Case Index
  Case 1
    SelectBrush
  Case 2
    SelectDense
  End Select

End Sub

Private Sub SelectBrush()
Dim Ftemp As FSelectTools
Dim X As Integer
Dim I As Integer

  Set Ftemp = New FSelectTools
  With Ftemp
    .Height = 120
    For I = 1 To ctlPit.BrushCount
      If I <> 1 Then
        Load .imgTool(I - 1)
        .imgTool(I - 1).Left = .imgTool(I - 2).Left + .imgTool(I - 2).Width + 120
        .imgTool(I - 1).Top = .imgTool(I - 2).Top
        .imgTool(I - 1).Visible = True
      End If
      Set .imgTool(I - 1).Picture = ctlPit.BrushImage(I)
      If .Height < .imgTool(I - 1).Height + 300 Then
        .Height = .imgTool(I - 1).Height + 300
      End If
    Next I
    .Width = .imgTool(.imgTool.Count - 1).Left + .imgTool(.imgTool.Count - 1).Width + 210
    .Left = FMain.Left + FMain.pctTools.Left * Screen.TwipsPerPixelX - .Width + 240
    .Top = FMain.Top + FMain.pctToolOption(1).Top + 240
    
    .Show vbModal
    If .Tag <> "" Then
      Me.ctlPit.BrushIndex = .Tag
    End If
  End With
  
  Unload Ftemp
  Set Ftemp = Nothing
End Sub

Private Sub SelectDense()
Dim Ftemp As FSelectTools
Dim X As Integer
Dim I As Integer

  Set Ftemp = New FSelectTools
  With Ftemp
    .Height = 120
    For I = 1 To Me.imgDense.Count
      If I <> 1 Then
        Load .imgTool(I - 1)
        .imgTool(I - 1).Left = .imgTool(I - 2).Left + .imgTool(I - 2).Width + 120
        .imgTool(I - 1).Top = .imgTool(I - 2).Top
        .imgTool(I - 1).Visible = True
      End If
      Set .imgTool(I - 1).Picture = Me.imgDense(I - 1)
      If .Height < .imgTool(I - 1).Height + 300 Then
        .Height = .imgTool(I - 1).Height + 300
      End If
    Next I
    .Width = .imgTool(.imgTool.Count - 1).Left + .imgTool(.imgTool.Count - 1).Width + 210
    .Left = FMain.Left + FMain.pctTools.Left * Screen.TwipsPerPixelX - .Width + 240
    .Top = FMain.Top + FMain.pctToolOption(2).Top + 240
    
    .Show vbModal
    If .Tag <> "" Then
      Me.ctlPit.BrushDense = Me.imgDense(.Tag - 1).Tag
    End If
  End With
  
  Unload Ftemp
  Set Ftemp = Nothing
End Sub

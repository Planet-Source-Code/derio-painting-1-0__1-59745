VERSION 5.00
Begin VB.Form FPit 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FPit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public X As Integer
Public Y As Integer
Public OldX As Integer
Public OldY As Integer

Dim PosX As Integer
Dim PosY As Integer

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long

  If Button = vbLeftButton Then
    MMain.RunMe.Play
    Result = ReleaseCapture()
    Result = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    MMain.RunMe.StopPlaying
    PosX = X
    PosY = Y
  Else
    PosX = X
    PosY = Y
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    If PosX <> 0 And PosY <> 0 Then
      Left = Left + (X - PosX) * Screen.TwipsPerPixelX
      Top = Top + (Y - PosY) * Screen.TwipsPerPixelY
    End If
  End If
End Sub

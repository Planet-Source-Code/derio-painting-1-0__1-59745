VERSION 5.00
Begin VB.Form FSelectTools 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3195
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgTool 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FSelectTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    Tag = ""
    Hide
  End If
End Sub

Private Sub imgTool_Click(Index As Integer)
  Tag = Index + 1
  Hide
End Sub


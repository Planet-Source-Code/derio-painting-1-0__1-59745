VERSION 5.00
Begin VB.Form Flogo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1500
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Flogo.frx":0000
   ScaleHeight     =   1500
   ScaleWidth      =   5400
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer pctStop 
      Interval        =   5000
      Left            =   2940
      Top             =   1020
   End
End
Attribute VB_Name = "Flogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pctStop_Timer()
  Unload Me
End Sub

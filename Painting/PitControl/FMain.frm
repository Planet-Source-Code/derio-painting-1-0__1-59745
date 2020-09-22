VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   436
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctBrushMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   1620
      Picture         =   "FMain.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   7
      Top             =   2100
      Width           =   240
   End
   Begin VB.PictureBox pctBrushMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   1260
      Picture         =   "FMain.frx":008A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   2100
      Width           =   240
   End
   Begin VB.PictureBox pctBrushMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   660
      Picture         =   "FMain.frx":0114
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   2100
      Width           =   480
   End
   Begin VB.PictureBox pctBrushMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "FMain.frx":01DE
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   2100
      Width           =   480
   End
   Begin VB.PictureBox pctPit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Index           =   0
      Left            =   60
      Picture         =   "FMain.frx":02A8
      ScaleHeight     =   123
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   3
      Top             =   120
      Width           =   1530
   End
   Begin VB.PictureBox pctPit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Index           =   1
      Left            =   1680
      Picture         =   "FMain.frx":38E2
      ScaleHeight     =   123
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   2
      Top             =   120
      Width           =   1530
   End
   Begin VB.PictureBox pctPit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Index           =   2
      Left            =   3300
      Picture         =   "FMain.frx":6F1C
      ScaleHeight     =   123
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   1
      Top             =   120
      Width           =   1530
   End
   Begin VB.PictureBox pctPit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Index           =   3
      Left            =   4920
      Picture         =   "FMain.frx":A556
      ScaleHeight     =   123
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   0
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



VERSION 5.00
Begin VB.Form frmSprites 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox rocket2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   5745
      Picture         =   "frmSprites.frx":0000
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   13
      Top             =   1710
      Width           =   690
   End
   Begin VB.PictureBox enemy1pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   585
      Picture         =   "frmSprites.frx":0322
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   34.286
      TabIndex        =   11
      Top             =   3390
      Width           =   1800
   End
   Begin VB.PictureBox enemy1mask 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   855
      Picture         =   "frmSprites.frx":1BC2
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   34.286
      TabIndex        =   12
      Top             =   3600
      Width           =   1800
   End
   Begin VB.PictureBox picShot 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   3
      Left            =   3450
      Picture         =   "frmSprites.frx":2D4B
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   10
      Top             =   1605
      Width           =   90
   End
   Begin VB.PictureBox picShot 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   5
      Left            =   3930
      Picture         =   "frmSprites.frx":2F6D
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   9
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox picShot 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   4
      Left            =   2850
      Picture         =   "frmSprites.frx":348F
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   8
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox topship 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   510
      Picture         =   "frmSprites.frx":39B1
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   6
      Top             =   2115
      Width           =   6300
   End
   Begin VB.PictureBox topshipmask 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   570
      Picture         =   "frmSprites.frx":550C
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   7
      Top             =   2355
      Width           =   6300
   End
   Begin VB.PictureBox picShot 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   2
      Left            =   1770
      Picture         =   "frmSprites.frx":69EB
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   5
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox picShot 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   1
      Left            =   930
      Picture         =   "frmSprites.frx":6F0D
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   4
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox picShot 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   0
      Left            =   1410
      Picture         =   "frmSprites.frx":742F
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   3
      Top             =   1605
      Width           =   90
   End
   Begin VB.PictureBox rocket 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   4650
      Picture         =   "frmSprites.frx":7651
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   2
      Top             =   1725
      Width           =   690
   End
   Begin VB.PictureBox bottomship 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   735
      Picture         =   "frmSprites.frx":7973
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   345
      Width           =   6300
   End
   Begin VB.PictureBox bottomshipmask 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   435
      Picture         =   "frmSprites.frx":1A105
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   1
      Top             =   615
      Width           =   6300
   End
End
Attribute VB_Name = "frmSprites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

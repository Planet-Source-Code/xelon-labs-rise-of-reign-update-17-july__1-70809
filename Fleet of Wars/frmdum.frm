VERSION 5.00
Begin VB.Form frmdum 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   3375
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   556
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   3840
      Left            =   4320
      ScaleHeight     =   252
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   252
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox msk 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   360
      Picture         =   "frmdum.frx":0000
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1560
   End
End
Attribute VB_Name = "frmdum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

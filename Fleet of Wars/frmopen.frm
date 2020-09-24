VERSION 5.00
Begin VB.Form frmopen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Map"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmopen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form1.Show
Form1.oopen Dir1
End Sub

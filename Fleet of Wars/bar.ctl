VERSION 5.00
Begin VB.UserControl Bar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   ScaleHeight     =   3600
   ScaleWidth      =   4920
   Begin VB.Shape brdr 
      BorderColor     =   &H00FFFFFF&
      Height          =   1695
      Left            =   1200
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FF80&
      X1              =   0
      X2              =   2400
      Y1              =   15
      Y2              =   15
   End
End
Attribute VB_Name = "Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Sub SetPro(Prec As Integer, by As Integer)
On Error Resume Next
Dim np As Integer
Line1.X2 = (((Prec / by) * 100) / 100) * Width
np = (Line1.X2 / Width) * 100
If np >= 75 Then Line1.BorderColor = vbGreen
If np >= 50 And np < 75 Then Line1.BorderColor = &H8000&
If np >= 35 And np < 50 Then Line1.BorderColor = &H80FFFF
If np >= 15 And np < 35 Then Line1.BorderColor = &H80C0FF
If np >= 0 And np < 15 Then Line1.BorderColor = &HFF&
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 45
brdr.Left = 0
brdr.Top = 0
brdr.Width = UserControl.Width
brdr.Height = UserControl.Height

End Sub

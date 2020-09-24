VERSION 5.00
Begin VB.UserControl button 
   BackColor       =   &H00808080&
   ClientHeight    =   6330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   DefaultCancel   =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   6120
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim img1 As StdPicture
Dim img2 As StdPicture

Const point As String = "Pointer"

Dim def As Boolean

Event click(str As String)
Event Move(str As String)
Event looseFocus()

Sub image(path As String)
Set img1 = LoadPicture(path)
Set UserControl.Picture = img1
End Sub

Sub image_on(path As String)
Set img2 = LoadPicture(path)
End Sub

Sub caption(str As String)
lbl = str
End Sub

Private Sub lbl_Click()
Play_Sound "Menu.wav"
RaiseEvent click(lbl)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl.SetFocus
End Sub

Private Sub UserControl_Click()
Play_Sound "Menu.wav"
RaiseEvent click(lbl)
End Sub

Private Sub UserControl_EnterFocus()
UserControl.Picture = img2
def = True
Play_Sound "Bar.wav"
RaiseEvent Move(lbl)
End Sub

Private Sub UserControl_ExitFocus()
UserControl.Picture = img1
def = False
RaiseEvent looseFocus
End Sub

Private Sub UserControl_Initialize()
LoadCursor point, hwnd
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And def = True Then
RaiseEvent click(lbl)
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl.SetFocus
End Sub

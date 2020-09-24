VERSION 5.00
Begin VB.Form frmEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Edit Mask of Terrain"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   3120
      Picture         =   "frmEdit.frx":0000
      ScaleHeight     =   1140
      ScaleWidth      =   1485
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   0
      Picture         =   "frmEdit.frx":5E02
      ScaleHeight     =   1140
      ScaleWidth      =   1485
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   1560
      Picture         =   "frmEdit.frx":BC04
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Hide
Me.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Setup Picture1.hdc, Picture2.hdc, Me, Picture2
BLTIT X, Y, Me
Cleanup
DoEvents
ElseIf Button = 1 Then
Setup Picture3.hdc, Picture2.hdc, Me, Picture2
BLTIT X, Y, Me
Cleanup
DoEvents
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cleanup
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cleanup
Dim lst As ListItem
If Me.Tag = "Add" Then
Set lv1 = Form1.lv1
Form1.setsea lv1.ListItems.Count + 1, Image
Set lst = lv1.ListItems.Add(, , lv1.ListItems.Count + 1)
lst.SubItems(1) = frmEdit.Width / 15
lst.SubItems(2) = frmEdit.Height / 15
lst.SubItems(3) = frmEdit.Left / 15
lst.SubItems(4) = frmEdit.Top / 15
lst.SubItems(5) = "grass.gif"
ElseIf Me.Tag = "Sea" Then
Set lv1 = Form1.lv1
Form1.setsea lv1.ListItems.Count + 1, Image
Set lst = lv1.ListItems.Add(, , lv1.ListItems.Count + 1)
lst.SubItems(1) = frmEdit.Width / 15
lst.SubItems(2) = frmEdit.Height / 15
lst.SubItems(3) = frmEdit.Left / 15
lst.SubItems(4) = frmEdit.Top / 15
lst.SubItems(5) = "Sea.gif"
ElseIf Left(Me.Tag, 4) = "Show" Then
Dim str As String
str = right(Me.Tag, Len(Tag) - 4)
Form1.setsea Val(str), frmEdit.Image
End If

End Sub

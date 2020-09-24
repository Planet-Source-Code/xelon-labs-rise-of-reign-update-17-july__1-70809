VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmnu 
   BackColor       =   &H00200405&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10890
   ClientLeft      =   4680
   ClientTop       =   1725
   ClientWidth     =   10095
   Icon            =   "frmmnu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10890
   ScaleWidth      =   10095
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Load 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   460
      Left            =   5880
      ScaleHeight     =   435
      ScaleWidth      =   1845
      TabIndex        =   19
      Top             =   5760
      Visible         =   0   'False
      Width           =   1870
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading ...."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox PicSet 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   2400
      ScaleHeight     =   4905
      ScaleWidth      =   3465
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CheckBox Check4 
         BackColor       =   &H0059341C&
         Caption         =   "Enable Ground"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   2640
         Width           =   2775
      End
      Begin VB.HScrollBar SR 
         CausesValidation=   0   'False
         Height          =   135
         Left            =   360
         Max             =   50
         TabIndex        =   16
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0059341C&
         Caption         =   "Enable Minimap"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0059341C&
         Caption         =   "Enable Start Reveal FX"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0059341C&
         Caption         =   "Enable Music Tracks"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Game Speed  [Requires Resources]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3720
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   10695
      Left            =   8040
      ScaleHeight     =   10665
      ScaleWidth      =   2025
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin RoR.button button3 
         Height          =   450
         Left            =   0
         TabIndex        =   4
         Top             =   9000
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button button2 
         Height          =   450
         Left            =   0
         TabIndex        =   3
         Top             =   8520
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button button1 
         Height          =   450
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   600
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   2
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   3
         Left            =   0
         TabIndex        =   7
         Top             =   1560
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   4
         Left            =   0
         TabIndex        =   8
         Top             =   2040
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   5
         Left            =   0
         TabIndex        =   9
         Top             =   2520
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   6
         Left            =   0
         TabIndex        =   10
         Top             =   3000
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button button4 
         Height          =   450
         Left            =   0
         TabIndex        =   11
         Top             =   8040
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin VB.Image imgSide 
         Height          =   135
         Left            =   120
         Top             =   4560
         Width           =   1695
      End
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   1920
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   3387
      _Version        =   393217
      BackColor       =   8421504
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmmnu.frx":27A2
   End
   Begin VB.Line lnhoz 
      BorderColor     =   &H00F2E8E1&
      X1              =   -100
      X2              =   -85
      Y1              =   -100
      Y2              =   -85
   End
   Begin VB.Line lnvrt 
      BorderColor     =   &H00F2E2D9&
      X1              =   -100
      X2              =   -85
      Y1              =   -100
      Y2              =   -85
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   2040
      Top             =   3480
      Width           =   3255
   End
End
Attribute VB_Name = "frmmnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLvL_click(Index As Integer, str As String)
If str <> "Not Available" Then
Load.Visible = True
LoadCursor "", hwnd
frmmain.Show
DoEvents
frmmain.LoadMap App.path & "\maps\Camp" & CStr(Index) & "\"
Unload Me
Else
rtb.Text = "Sir, You will have to complete the previous missions to proceed"
rtb.Visible = True
End If
End Sub

Private Sub btnLvL_Move(Index As Integer, str As String)
If str <> "Not Available" Then
rtb.Visible = True
rtb.LoadFile App.path & "\Maps\Camp" & CStr(Index) & "\Description.txt"
DoEvents
Else
rtb.Visible = False
End If
PicSet.Visible = False
SetXY Picture1.Left - 256, btnLvL(Index).Top + (btnLvL(Index).Height / 2)
End Sub

Private Sub button1_click(str As String)
Load.Visible = True
frmmain.Show
DoEvents
frmmain.LoadMap App.path & "\maps\Camp" & GetFromIni("Main", "Progress", App.path & "\set.cfg") & "\"
Unload Me
End Sub

Private Sub button1_Move(str As String)
rtb.Visible = True
rtb.LoadFile App.path & "\Maps\Camp" & GetFromIni("Main", "Progress", App.path & "\set.cfg") & "\Description.txt"
DoEvents
SetXY Picture1.Left - 256, (button1.Height / 2)
PicSet.Visible = False
End Sub

Private Sub button2_Move(str As String)
rtb.Visible = True
PicSet.Visible = False
rtb.Text = "This Code Is developed by Ali Ashraf 100% Original Formula" & vbCrLf & "Except lavolpe 32-Bit DIB (Thanks) "
SetXY Picture1.Left - 256, button2.Top + (button2.Height / 2)
End Sub

Private Sub button3_click(str As String)
End
End Sub

Private Sub button3_Move(str As String)
rtb.Visible = True
PicSet.Visible = False
rtb.Text = "Do you really want to leave the Battle Arena" & vbCrLf & " ?  &  ! "
SetXY Picture1.Left - 256, button3.Top + (button3.Height / 2)
End Sub

Private Sub eves_MouseEnter(ctlEntered As Control)
If LCase(ctlEntered.Name) = "imgside" Then
LoadCursor "Pointer", ctlEntered.hwnd
End If
End Sub

Private Sub button4_Move(str As String)
rtb.Visible = True
rtb.Text = "Set Settings according to your computer's capability" & vbCrLf & "More Visual FX will slow down the game."
PicSet.Visible = True
SetXY Picture1.Left - 256, button4.Top + (button4.Height / 2)
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WriteIni "Main", "Music", Check1.Value, App.path & "\set.cfg"
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WriteIni "Main", "Reveal", Check2.Value, App.path & "\set.cfg"
End Sub

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WriteIni "Main", "Minimap", Check3.Value, App.path & "\set.cfg"
End Sub

Private Sub Check4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WriteIni "Main", "Ground", Check4.Value, App.path & "\set.cfg"
End Sub

Private Sub Form_Load()
Dim X As Integer
Me.WindowState = 2
Set Image1.Picture = LoadPicture(App.path & "\images\buildbar\back.jpg")
Set PicSet.Picture = LoadPicture(App.path & "\Images\BuildBar\Menu.gif")
Set Load.Picture = LoadPicture(App.path & "\images\buildbar\button.gif")
button1.caption "Resume Campaign"
button1.image App.path & "\images\buildbar\button.gif"
button1.image_on App.path & "\images\buildbar\button0.gif"
button2.caption "Credits"
button2.image App.path & "\images\buildbar\button.gif"
button2.image_on App.path & "\images\buildbar\button0.gif"
button3.caption "Exit Game"
button3.image App.path & "\images\buildbar\button.gif"
button3.image_on App.path & "\images\buildbar\button0.gif"
button4.caption "Settings"
button4.image App.path & "\images\buildbar\button.gif"
button4.image_on App.path & "\images\buildbar\button0.gif"
ReSet
imgSide.Left = 0
imgSide.Top = 0
Set imgSide.Picture = LoadPicture(App.path & "\images\Buildbar\Side.gif")
imgSide.Stretch = True
imgSide.Height = Screen.Height
LoadCursor "Select", hwnd
ComeOn_B
comeOn_D
ComeOn_L
LoadCursor "Pointer", Picture1.hwnd
Getini
End Sub

Sub ReSet()
For X = 1 To 6
If X <= Val(GetFromIni("Main", "Progress", App.path & "\set.cfg")) Then
btnLvL(X).caption GetFromIni("Main", "Name", App.path & "\maps\camp" & X & "\ini.ini")
Else
btnLvL(X).caption "Not Available"
End If
btnLvL(X).image App.path & "\images\buildbar\button.gif"
btnLvL(X).image_on App.path & "\images\buildbar\button0.gif"
Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetXY X, Y
End Sub

Private Sub Form_Resize()
Picture1.Left = Me.Width - Picture1.Width
Picture1.Height = Me.Height
Load.Left = (Me.Width / 2) - (Load.Width / 2) - Picture1.Width
Load.Top = (Height / 2) - (Load.Height / 2)
Image1.Left = (Me.Width / 2) - (Image1.Width / 2) - Picture1.Width
Image1.Top = (Height / 2) - (Image1.Height / 2)
PicSet.Left = (Me.Width / 2) - (PicSet.Width / 2) - Picture1.Width
PicSet.Top = (Height / 2) - (PicSet.Height / 2)
button3.Top = Height - button3.Height
button4.Top = Height - (button4.Height * 3)
button2.Top = button3.Top - 480
End Sub

Sub SetNext()
Show
comeOn_D
ComeOn_B
ComeOn_L
ReSet
End Sub

Sub SetLoose()
Show
Unload frmmain
button1.caption "Retry Mission"
rtb.Visible = True
rtb.Text = "You Lost" & vbCrLf & "Retry Again and Test your Skills"
End Sub

Sub ComeOn_B()
Dim k As Integer
For k = 2055 To 0 Step -1
DoEvents
button1.Left = k
DoEvents
button2.Left = k
DoEvents
button3.Left = k
DoEvents
Next
End Sub

Sub comeOn_D()
Dim k As Integer
For k = -rtb.Height To 0 Step -1
rtb.Top = k
DoEvents
Next
End Sub

Sub ComeOn_L()
Dim k As Integer, Dk As Integer
For k = 2055 To 0 Step -1
For Dk = 1 To 6
btnLvL(Dk).Left = k
DoEvents
Next
Next
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetXY X + Image1.Left, Y + Image1.Top
End Sub

Sub SetXY(ByVal cX As Integer, ByVal cY As Integer)
lnvrt.X1 = cX
lnvrt.X2 = cX
lnvrt.Y1 = 0
lnvrt.Y2 = Height

lnhoz.Y1 = cY
lnhoz.Y2 = cY
lnhoz.X1 = 0
lnhoz.X2 = Width
End Sub

Private Sub PicSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetXY PicSet.Left, PicSet.Top
End Sub

Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetXY rtb.Width, rtb.Height
End Sub

Sub Getini()
Check1.Value = Val(GetFromIni("Main", "Music", App.path & "\set.cfg"))
Check2.Value = Val(GetFromIni("Main", "Reveal", App.path & "\set.cfg"))
Check3.Value = Val(GetFromIni("Main", "Minimap", App.path & "\set.cfg"))
Check4.Value = Val(GetFromIni("Main", "Ground", App.path & "\set.cfg"))
SR.Value = Val(GetFromIni("Main", "Speed", App.path & "\set.cfg"))
End Sub

Private Sub SR_Change()
WriteIni "Main", "Speed", SR.Value, App.path & "\set.cfg"
End Sub

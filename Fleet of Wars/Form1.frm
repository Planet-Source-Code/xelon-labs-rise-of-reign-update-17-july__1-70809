VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H00200405&
   BorderStyle     =   0  'None
   Caption         =   "Rise of a Chaos"
   ClientHeight    =   7515
   ClientLeft      =   3345
   ClientTop       =   3435
   ClientWidth     =   12765
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   501
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   851
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Piclbl 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   11280
      ScaleHeight     =   3465
      ScaleWidth      =   4905
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   9000
      ScaleHeight     =   4905
      ScaleWidth      =   3465
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
      Begin RoR.button button1 
         Height          =   450
         Left            =   840
         TabIndex        =   16
         Top             =   480
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button button3 
         Height          =   450
         Left            =   840
         TabIndex        =   17
         Top             =   4080
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button button2 
         Height          =   450
         Left            =   840
         TabIndex        =   18
         Top             =   3600
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
   End
   Begin VB.PictureBox BldBar 
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   0
      ScaleHeight     =   2250
      ScaleWidth      =   11250
      TabIndex        =   8
      Top             =   5280
      Width           =   11250
      Begin VB.PictureBox Minimap 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   2250
         Left            =   9000
         ScaleHeight     =   2220
         ScaleWidth      =   2220
         TabIndex        =   21
         Top             =   0
         Width           =   2250
         Begin VB.Shape loc 
            BorderColor     =   &H00F2E2D9&
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   855
         End
      End
      Begin RoR.DataView DVW 
         Height          =   1425
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2514
      End
      Begin MCI.MMControl mmc 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   8640
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   0
         PlayEnabled     =   -1  'True
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         PauseVisible    =   0   'False
         BackVisible     =   0   'False
         StepVisible     =   0   'False
         StopVisible     =   0   'False
         RecordVisible   =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Label lbltl 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1095
         Left            =   5520
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblmsg 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   2000
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "COST"
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   5880
      TabIndex        =   7
      Text            =   "Command Line"
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   327
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   599
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      Begin RoR.Bar Bar 
         Height          =   45
         Left            =   1080
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   79
      End
      Begin RoR.aeroplanes aeros 
         Height          =   495
         Left            =   2160
         TabIndex        =   5
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin RoR.struc struc 
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin RoR.tanks tanks 
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin RoR.weapons weap 
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin VB.ListBox lstsel 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   1560
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblair 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to AirStrike"
         Height          =   255
         Left            =   5640
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblbldng 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to place the building"
         Height          =   495
         Left            =   5640
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin RoR.aicAlphaImage spot 
         Height          =   1335
         Left            =   480
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   1296
         _ExtentY        =   1508
         Image           =   "Form1.frx":27A2
         Scaler          =   3
      End
      Begin RoR.aicAlphaImage Missile 
         Height          =   1455
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         Image           =   "Form1.frx":27BA
         Scaler          =   3
         Props           =   0
      End
      Begin VB.Image tree 
         Height          =   1935
         Index           =   0
         Left            =   3240
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin RoR.aicAlphaImage bldng 
         Height          =   1455
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         Image           =   "Form1.frx":27D2
         Scaler          =   3
         HitTest         =   3
      End
      Begin VB.Image air 
         Height          =   1935
         Index           =   0
         Left            =   2880
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Line Dpth 
         BorderColor     =   &H00FF0000&
         Index           =   0
         Visible         =   0   'False
         X1              =   216
         X2              =   184
         Y1              =   240
         Y2              =   208
      End
      Begin VB.Line pth 
         Index           =   0
         Visible         =   0   'False
         X1              =   208
         X2              =   336
         Y1              =   232
         Y2              =   104
      End
      Begin VB.Line lnbom 
         Index           =   0
         Visible         =   0   'False
         X1              =   328
         X2              =   200
         Y1              =   96
         Y2              =   224
      End
      Begin VB.Shape sel 
         BorderStyle     =   3  'Dot
         Height          =   1935
         Left            =   2760
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Shape bomb 
         FillColor       =   &H00424242&
         FillStyle       =   0  'Solid
         Height          =   45
         Index           =   0
         Left            =   840
         Shape           =   2  'Oval
         Top             =   3360
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Image sbl 
         Height          =   1935
         Left            =   3000
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image gizzly 
         Height          =   1935
         Index           =   0
         Left            =   3120
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Line lnmsc 
         BorderColor     =   &H00FF0000&
         Visible         =   0   'False
         X1              =   344
         X2              =   312
         Y1              =   112
         Y2              =   80
      End
      Begin VB.Line Line1 
         Index           =   0
         Visible         =   0   'False
         X1              =   192
         X2              =   320
         Y1              =   216
         Y2              =   88
      End
      Begin VB.Line lssr 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   216
         X2              =   184
         Y1              =   112
         Y2              =   80
      End
   End
   Begin VB.Timer tmrSlip 
      Interval        =   9
      Left            =   3480
      Top             =   0
   End
   Begin VB.Timer auto_bldng 
      Interval        =   1200
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer BomTmr 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   20
      Left            =   1560
      Top             =   0
   End
   Begin VB.Timer auto 
      Interval        =   1200
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   600
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer t2b 
      Interval        =   600
      Left            =   2520
      Top             =   0
   End
   Begin VB.Timer tmrpvw 
      Interval        =   500
      Left            =   2040
      Top             =   0
   End
   Begin VB.Timer tmrpth 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   3000
      Top             =   0
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   3960
      Top             =   0
   End
   Begin VB.Timer tmrmon 
      Interval        =   400
      Left            =   4440
      Top             =   0
   End
   Begin VB.Timer TmrRef 
      Interval        =   1000
      Left            =   4920
      Top             =   0
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const ATN_1 As Double = 0.785398163397448
Private Const PI        As Double = 3.14159265358979
Private Const Level   As Integer = 128 ' Height of aircraft during flight hours (In Pixels)

'This looks like a great burden on the memory but accordng to my
'calculations the average data stored in memory is less than 5 MB

Dim cursor As String

'AI Data
Dim AI_turn As Integer
Dim AI_Skills As Integer

'Timer Data
Dim label As String
Dim time As Integer
Dim t_loop As Boolean
Dim S_time As Integer
Dim Trig_Count As Integer
Dim trigger(100) As String
Dim map_light As Integer
Dim map_condition As String
'Minimap Data

'Map Data
Dim Map_Name As String
Dim Map_Money As Long
Dim Map_Ground As c32bppDIB
Dim Map_Eves_Count As Integer
Dim Map_On(100) As String
Dim Map_Do(100) As String
Dim Map_Done(100) As Boolean
Dim Map_Techlevel As Integer
Dim Map_Level As Integer
Dim Map_Many As Integer

'Aircrafts Data
Dim air_k(1000) As Long
Dim air_img(1000) As String
Dim air_speed(1000) As Long
Dim air_power(1000) As Long
Dim air_t_power(1000) As Long
Dim air_weapon(1000) As String
Dim air_team(1000) As String
Dim air_Angl(1000) As String

'Tanks Data
Dim k(1000) As Long
Dim bk(1000) As Long
Dim img(1000) As String
Dim speed(1000) As Long
Dim tnk_name(1000) As String
Dim power(1000) As Long
Dim t_power(1000) As Long
Dim weapon(1000) As String
Dim team(1000) As String
Dim typ(1000) As String
Dim Angl(1000) As String

'Buildings Data
Dim bldng_pow(1000) As Long
Dim bldng_name(1000) As String
Dim bldng_tpow(1000) As Long
Dim bldng_team(1000) As String
Dim bldng_offsetX(1000) As Integer
Dim bldng_OffsetY(1000) As Integer
Dim bldng_weapon(1000) As String

'Extra User Data
Dim mode As Integer
Dim Host As Integer
Dim pvw As String
Dim R As Integer
Dim mX As Integer, mY As Integer
Dim idx As Integer
Dim n_rand As Integer
Dim pt As POINTAPI

Function Sin2(Angle As Integer)
Sin2 = Sin(PI * Angle / 180)
End Function
Public Function ACos(ByVal d As Double) As Double
    ACos = Atn(-d / Sqr(-d * d + 1)) + 2 * ATN_1
End Function
Public Function ASin(ByVal d As Double) As Double
    ASin = Atn(d / Sqr(-d * d + 1))
End Function

Private Sub auto_bldng_Timer()
On Error Resume Next
Dim unt As Integer
Dim e_bldng As Integer
Dim rng As Integer
For e_bldng = bldng.LBound To bldng.UBound
For unt = gizzly.LBound To gizzly.UBound
rng = weap.range(bldng_weapon(e_bldng))
If e_bldng = 0 Or unt = 0 Or bldng_team(e_bldng) = "-1" Or team(unt) = "-1" Then GoTo R
If team(unt) <> bldng_team(e_bldng) Then
If bldng(e_bldng).Left > gizzly(unt).Left - rng And bldng(e_bldng).Left < gizzly(unt).Left + rng Then
If bldng(e_bldng).Top > gizzly(unt).Top - rng And bldng(e_bldng).Top < gizzly(unt).Top + rng Then

If bldng(e_bldng).Tag = "" Then
bldng(e_unt).Tag = CStr(unt)
End If

If bldng(e_unt).Tag = CStr(unt) Then
If weap.typ(bldng_weapon(e_bldng)) = "laser" Then
Laser e_bldng, unt, weap.damage(bldng_weapon(e_bldng)), weap.Color(bldng_weapon(e_bldng)), 1
ElseIf weap.typ(bldng_weapon(e_bldng)) = "bomb" Then
DoEvents
fire e_bldng, unt, "bld", ""
End If
End If
End If
End If
End If
R:
bldng(e_bldng).Tag = ""
DoEvents
Next
Next

AI_turn = AI_turn + 1
If AI_Skills <> 0 Then
If (AI_turn / AI_Skills) = Fix(AI_turn / AI_Skills) Then
DoEvents
DoAI
End If
End If

If (AI_turn / 11) = Fix(AI_turn / 11) Then
If LCase(map_condition) = "storm" Then
bolt Fix(Rnd * Picture1.Width), Fix(Rnd * Picture1.Height)
End If
End If
End Sub

Sub DoAI()
On Error Resume Next
Dim X1 As Integer
Dim mdX1 As Integer
For X1 = bldng.LBound To bldng.UBound
If bldng_team(X1) = "Allies" Then
If Sgn(Rnd - Rnd) = 1 Then
mdX1 = X1
DoEvents
GoTo IU
End If
End If
Next

IU:
For X1 = bldng.LBound To bldng.UBound
If LCase(struc.typ(bldng_name(X1))) = "warfactory" And bldng_team(X1) <> "Allies" Then
If Sgn(Rnd - Rnd) = -1 Then
tnkfromini "cosmiq", bldng_team(X1), bldng(X1).Left + struc.offx(bldng_name(X1)), bldng(X1).Top + struc.offy(bldng_name(X1)), bldng(mdX1).Left, bldng(mdX1).Top
DoEvents
GoTo UI
End If
End If
Next
UI:
End Sub
Private Sub auto_Timer()
On Error Resume Next
Dim Deg As Integer
Dim rad As Integer
Dim e_unt As Integer
Dim unt As Integer
Dim rng As Integer
Dim rng2 As Integer
For unt = gizzly.LBound To gizzly.UBound
For e_unt = gizzly.LBound To gizzly.UBound
If unt = 0 Or e_unt = 0 Then GoTo Y
If team(e_unt) <> team(unt) And team(e_unt) <> "-1" And team(unt) <> "-1" Then
rng = weap.range(weapon(unt))
If (gizzly(unt).Left + gizzly(unt).Width / 2) > (gizzly(e_unt).Left + gizzly(e_unt).Width / 2) - rng And (gizzly(unt).Left + gizzly(unt).Width / 2) < (gizzly(e_unt).Left + gizzly(e_unt).Width / 2) + rng Then
If (gizzly(unt).Top + gizzly(unt).Height / 2) > (gizzly(e_unt).Top + gizzly(e_unt).Height / 2) - rng And (gizzly(unt).Top + gizzly(unt).Height / 2) < (gizzly(e_unt).Top + gizzly(e_unt).Height / 2) + rng Then
If gizzly(unt).Tag = "" Then
gizzly(unt).Tag = "tnk" & e_unt
End If

If gizzly(unt).Tag = "tnk" & e_unt Then
If gizzly(unt).ToolTipText = "" Then
If weap.typ(weapon(unt)) = "bomb" Then
fire unt, e_unt, "", ""
ElseIf weap.typ(weapon(unt)) = "laser" Then
Laser unt, e_unt, weap.damage(weapon(unt)), weap.Color(weapon(unt)), 3
End If
End If
End If

End If
End If
DoEvents
End If
DoEvents
gizzly(unt).Tag = ""
Y:
Next
Next
End Sub

Private Sub bldng_Click(Index As Integer, ByVal Button As Integer)
On Error Resume Next
Dim n As Integer
If mode <> 2 Then
If bldng_team(Index) <> "Allies" Then

Dim X As Integer, Y As Integer, wid As Long, nm As Integer, tr As Boolean, ene As Integer, conf As Boolean: tr = False

For nm = 0 To lstsel.ListCount - 1
ene = Val(lstsel.List(nm))
rad = weap.range(weapon(ene))
If (gizzly(ene).Left + gizzly(ene).Width / 2) > (bldng(Index).Left + bldng(Index).Width / 2) - rad And (gizzly(ene).Left + gizzly(ene).Width / 2) < (bldng(Index).Left + bldng(Index).Width / 2) + rad Then
If (gizzly(ene).Top + gizzly(ene).Height / 2) > (bldng(Index).Top + bldng(Index).Height / 2) - rad And (gizzly(ene).Top + gizzly(ene).Height / 2) < (bldng(Index).Top + bldng(Index).Height / 2) + rad Then
conf = True
End If
End If
If conf = False Then
With lnmsc
.X1 = gizzly(ene).Left - (gizzly(ene).Width / 2): .Y1 = gizzly(ene).Top - (gizzly(ene).Height / 2)
.X2 = bldng(Index).Left - (bldng(Index).Height / 2): .Y2 = bldng(Index).Top - (bldng(Index).Height / 2)
wid = Hyp(Wline(.X1, .X2), Hline(.Y1, .Y2))
Y = -(Sin(PI * Angle(.X2, .Y2, .X1, .Y1) / 180) * (rad - 5)) + (bldng(Index).Top + bldng(Index).Height / 2)
X = Cos(PI * Angle(.X2, .Y2, .X1, .Y1) / 180) * (rad - 5) + (bldng(Index).Left + bldng(Index).Width / 2)
TnkMove ene, X, Y
gizzly(ene).Tag = "bldng" & Index
End With
End If
Next
Else
Host = Index
If LCase(struc.typ(bldng_name(Index))) = "constyard" Then
BldBar.Tag = "bldng"
ElseIf LCase(struc.typ(bldng_name(Index))) = "acc" Then
BldBar.Tag = "air"
ElseIf LCase(struc.typ(bldng_name(Index))) = "warfactory" Then
BldBar.Tag = "tank"
Else
End If
Bldbar_Set Index
End If

Else


If Map_Money - aeros.cost(DVW.sel) < 0 Then
Play_Sound "Warning.wav"
Msg "Not Enough Credits, You need : " & CStr(aeros.cost(DVW.sel))
mode = 0
cursor = "Pointer"
Else
AirMission Host, "Allies", DVW.sel, bldng(Index).Left + bldng(Index).Width / 2, bldng(Index).Top + bldng(Index).Height - 9
Map_Money = Map_Money - aeros.cost(DVW.sel)
mode = 0
cursor = "Pointer"
End If


End If
End Sub

Private Sub bldng_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

pvw = "bldng" & Index
End Sub

Private Sub BomTmr_Timer(Index As Integer)
On Error Resume Next
Dim X As Long, Y As Long, dmg As Long, wid As Long, L As Integer, n As Integer, m_idx As Integer, so As String, des As String, frm As Integer
With lnbom(Index)
L = bk(Index)
wid = Hyp(Wline(.X1, .X2), Hline(.Y1, .Y2))
m_idx = Val(getval(bomb(Index).Tag))
so = Left(bomb(Index).Tag, 3)
des = Left(.Tag, 3)
frm = Val(getval(.Tag))
If bk(Index) <= wid Then
Y = -(Sin(PI * Angle(.X1, .Y1, .X2, .Y2) / 180) * bk(Index)) + .Y1
X = Cos(PI * Angle(.X1, .Y1, .X2, .Y2) / 180) * bk(Index) + .X1
n = (L / 100) * 180
bomb(Index).Visible = True
bomb(Index).Move X - 3, Y - 3
If so = "bld" Then
bk(Index) = bk(Index) + weap.speedstep(bldng_weapon(frm))
BomTmr(Index).interval = weap.interval(bldng_weapon(frm))
ElseIf so = "air" Then
bk(Index) = bk(Index) + weap.speedstep(air_weapon(frm))
BomTmr(Index).interval = weap.interval(air_weapon(frm))
Else
bk(Index) = bk(Index) + weap.speedstep(weapon(frm))
BomTmr(Index).interval = weap.interval(weapon(frm))
End If
DoEvents
Else
If des = "bld" Then
bldng_pow(m_idx) = Val(bldng_pow(m_idx) - weap.damage(weapon(frm)))
If bldng_pow(m_idx) <= 0 Then
destruct m_idx
End If
ElseIf so = "air" Then
explode .X2, .Y2, weap.damage(weapon(frm)) / 4, weap.damage(weapon(frm))
Else
power(m_idx) = Val(power(m_idx) - weap.damage(weapon(frm)))
If power(m_idx) <= 0 Then
desttank m_idx
End If
End If
If so <> "bld" And so <> "air" Then
gizzly(frm).Tag = ""
gizzly(frm).ToolTipText = ""
End If
If so = "bld" Then
bldng(frm).Tag = ""
End If

DoEvents
BomTmr(Index) = False
bk(Index) = 0
Unload bomb(Index)
Unload lnbom(Index)
Unload BomTmr(Index)
tmrpvw = True
End If
End With
End Sub

Sub explode(ByVal X As Integer, ByVal Y As Integer, ByVal rad As Integer, ByVal damage As Integer)
On Error Resume Next
Dim unt As Integer
For unt = gizzly.LBound To gizzly.UBound
If unt = 0 Or team(unt) = "-1" Then GoTo Y
If gizzly(unt).Left + (gizzly(unt).Width / 2) > X - rad And gizzly(unt).Left + (gizzly(unt).Width / 2) < X + rad Then
If gizzly(unt).Top + (gizzly(unt).Height / 2) > Y - rad And gizzly(unt).Top + (gizzly(unt).Height / 2) < Y + rad Then
power(unt) = power(unt) - damage
If power(unt) <= 0 Then
desttank unt
End If
End If
Else
End If
Y:
Next

For unt = bldng.LBound To bldng.UBound
If unt = 0 Or bldng_team(unt) = "-1" Then GoTo X
If bldng(unt).Left + (bldng(unt).Width / 2) > X - rad And bldng(unt).Left + (bldng(unt).Width / 2) < X + rad Then
If bldng(unt).Top + (bldng(unt).Height / 2) > Y - rad And bldng(unt).Top + (bldng(unt).Height / 2) < Y + rad Then
bldng_pow(unt) = bldng_pow(unt) - damage
If bldng_pow(unt) <= 0 Then
destruct unt
End If
End If
End If
X:
Next
End Sub
Sub destruct(m_idx As Integer)
On Error Resume Next
Dim X As Integer, Y As Integer
X = bldng(m_idx).Left + bldng(m_idx).Width / 2
Y = bldng(m_idx).Top + bldng(m_idx).Height / 2
Unload bldng(m_idx)
sblast X, Y
bldng_pow(m_idx) = 0
bldng_offsetX(m_idx) = 0
bldng_OffsetY(m_idx) = 0
bldng_tpow(m_idx) = -1
bldng_weapon(m_idx) = ""
bldng_team(m_idx) = "-1"
Play_Sound "ExStruct.wav"

Dim k As Integer
Dim arg(0) As String
For k = 1 To Map_Eves_Count
If LCase(Left(Map_On(k), 10)) = "destbldng(" Then
arg(0) = Right(Map_On(k), Len(Map_On(k)) - 10)
arg(0) = Left(arg(0), Len(arg(0)) - 1)
If Val(arg(0)) - 1 = m_idx Then
Trig Map_Do(k)
End If
End If
Next

End Sub

Sub desttank(m_idx As Integer)
On Error Resume Next
Dim X As Integer, mdX As Integer, mdY As Integer
Play_Sound "ExTank.wav"
If team(m_idx) = "Allies" Then
lstsel.clear
End If
mdX = gizzly(m_idx).Left + gizzly(m_idx).Width / 2
mdY = gizzly(m_idx).Top + gizzly(m_idx).Height / 2
Unload gizzly(m_idx)
sblast mdX, mdY
 k(m_idx) = 0
bk(m_idx) = 0
img(m_idx) = ""
power(m_idx) = 0
t_power(m_idx) = 100
speed(m_idx) = 0
weapon(m_idx) = ""
team(m_idx) = "-1"

Dim arg(0) As String
For X = 1 To Map_Eves_Count
If LCase(Left(Map_On(X), 8)) = "destunt(" Then
arg(0) = Right(Map_On(X), Len(Map_On(X)) - 8)
arg(0) = Left(arg(0), Len(arg(0)) - 1)
If Val(arg(0)) - 1 = m_idx Then
Trig Map_Do(X)
End If
End If
Next
End Sub

Function getval(str As String) As String
On Error Resume Next
If Left(str, 3) = "bld" Or Left(str, 3) = "bld" Then
getval = Right$(str, Len(str) - 3)
Else
getval = str
End If
End Function

Function Wline(X1 As Integer, X2 As Integer) As Integer
Wline = Abs(X2 - X1)
End Function

Function Hline(Y1 As Integer, Y2 As Integer) As Integer
Hline = Abs(Y2 - Y1)
End Function

Function Hyp(X As Long, Y As Long) As Integer
Hyp = Sqr((X * X) + (Y * Y))
End Function

Function Angle(X1, Y1, X2, Y2) As Long
On Error Resume Next
Dim nx As Integer, ny As Integer
nx = X2 - X1: ny = Y2 - Y1
If Sgn(nx) = 1 And Sgn(ny) = 1 Then
Angle = (Atn(Abs(Y2 - Y1) / Abs(X2 - X1)) * 180 / PI)
Angle = Abs(90 - Angle) + 270
ElseIf Sgn(nx) = -1 And Sgn(ny) = 1 Then
Angle = (Atn(Abs(Y2 - Y1) / Abs(X2 - X1)) * 180 / PI)
Angle = Angle + 180
ElseIf Sgn(nx) = -1 And Sgn(ny) = -1 Then
Angle = (Atn(Abs(Y2 - Y1) / Abs(X2 - X1)) * 180 / PI)
Angle = Abs(90 - Angle) + 90
ElseIf Sgn(nx) = 1 And Sgn(ny) = -1 Then
Angle = (Atn(Abs(Y2 - Y1) / Abs(X2 - X1)) * 180 / PI)
ElseIf Sgn(ny) = 1 And Sgn(nx) = 0 Then
Angle = 270
ElseIf Sgn(ny) = -1 And Sgn(nx) = 0 Then
Angle = 90
ElseIf Sgn(ny) = 0 And Sgn(nx) = -1 Then
Angle = 180
End If
End Function

Private Sub button1_click(str As String)
On Error Resume Next
doall True
Picture2.Visible = False
End Sub

Sub doall(bool As Boolean)
On Error Resume Next
auto_bldng = bool
auto = bool
tmrpvw = bool
t2b = bool
tmrSlip = bool
Timer = bool
tmrmon = bool
End Sub

Private Sub button2_click(str As String)
On Error Resume Next
doall False
frmmnu.SetLoose
Unload Me
End Sub

Private Sub button3_click(str As String)
End
End Sub

Private Sub DVW_Click(str As String)
On Error Resume Next
If BldBar.Tag = "bldng" And str <> "REPAIR" Then
mode = 1
cursor = "Build"
ElseIf BldBar.Tag = "tank" And str <> "REPAIR" Then
If Map_Money - tanks.cost(str) < 0 Then
Play_Sound "Warning.wav"
Msg "Not Enough Credits, You need : " & CStr(tanks.cost(str))
Else
tnkfromini str, "Allies", bldng(Host).Left + struc.offx(bldng_name(Host)), bldng(Host).Top + struc.offy(bldng_name(Host)), bldng(Host).Left + struc.DocX(bldng_name(Host)), bldng(Host).Top + struc.docY(bldng_name(Host))
Map_Money = Map_Money - tanks.cost(str)
End If
ElseIf BldBar.Tag = "air" And str <> "REPAIR" Then
mode = 2
cursor = "AirStrike"
ElseIf str = "REPAIR" Then

If Map_Money - ((bldng_tpow(Host) - bldng_pow(Host)) / 2) < 0 Then
Play_Sound "Warning.wav"
Msg "Not Enough Credits to repair this Structure, You need :" & CStr((bldng_tpow(Host) - bldng_pow(Host)) / 2)
Else
bldng_pow(Host) = bldng_tpow(Host)
Map_Money = Map_Money - ((bldng_tpow(Host) - bldng_pow(Host)) / 2)
End If


End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim X As Integer
team(0) = "-1"
pvw = 1
idx = 1
Map_Techlevel = 1
Set Map_Ground = New c32bppDIB
weap.loadwep
tanks.loadtnx
struc.loadbldng
aeros.loadair
Set BldBar.Picture = LoadPicture(App.path & "\Images\BuildBar\Bar.gif")
button1.caption "Resume Campaign"
button1.image App.path & "\images\buildbar\button.gif"
button1.image_on App.path & "\images\buildbar\button0.gif"
button2.caption "Exit to Main Menu"
button2.image App.path & "\images\buildbar\button.gif"
button2.image_on App.path & "\images\buildbar\button0.gif"
button3.caption "Exit to Windows"
button3.image App.path & "\images\buildbar\button.gif"
button3.image_on App.path & "\images\buildbar\button0.gif"
Set BldBar.Picture = LoadPicture(App.path & "\Images\BuildBar\Bar.gif")
Set Picture2.Picture = LoadPicture(App.path & "\Images\BuildBar\Menu.gif")
cursor = "Pointer"
End Sub

Sub makeTree(X As Integer, Y As Integer)
On Error Resume Next
Load tree(tree.UBound + 1)
tree(tree.UBound).Left = X
tree(tree.UBound).Top = Y
tree(tree.UBound).Visible = True
Set tree(tree.UBound).Picture = LoadPicture(App.path & "\images\trees\Tree (" & CStr(Round(Rnd * 23)) & ").gif")
End Sub

Sub AirMission(dock As Integer, side As String, aeroplane As String, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next
Dim mdX As Integer, mdY As Integer
mdX = (bldng(dock).Left + bldng(dock).Width / 2)
mdY = (bldng(dock).Top + bldng(dock).Height / 2)
Makeair aeros.image(aeroplane), side, mdX + bldng_offsetX(dock), mdY + bldng_OffsetY(dock), mdX + struc.DocX(bldng_name(dock)), mdY + struc.docY(bldng_name(dock)), X, Y, aeros.power(aeroplane), aeros.weapon(aeroplane), aeros.speed(aeroplane)
End Sub

Sub Makeair(image As String, side As String, ByVal dx As Integer, ByVal dy As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal toX As Integer, ByVal toY As Integer, ByVal e_power As Long, e_weapon As String, ByVal e_speed As Integer)
On Error Resume Next
Load air(air.UBound + 1)
Load pth(air.UBound)
Load Dpth(air.UBound)
Load tmrpth(air.UBound)
DoEvents
air_k(air.UBound) = 0
air_img(air.UBound) = image
air_weapon(air.UBound) = e_weapon
air_speed(air.UBound) = e_speed
air_power(air.UBound) = e_power
air_t_power(air.UBound) = e_power
air(air.UBound).Left = X
air_team(air.UBound) = side
air(air.UBound).Top = Y
Dpth(air.UBound).X1 = dx
Dpth(air.UBound).Y1 = dy
Dpth(air.UBound).X2 = X
Dpth(air.UBound).Y2 = Y - Level
pth(air.UBound).X1 = X
pth(air.UBound).Y1 = Y - Level
pth(air.UBound).X2 = toX
pth(air.UBound).Y2 = toY
Rotate_air (Angle(dx, dy, X, Y)), air_img(air.UBound), air.UBound
air(air.UBound).Refresh
air(air.UBound).ZOrder 0
tmrpth(air.UBound) = True
tmrpth(air.UBound).Tag = "1"
air(air.UBound).Visible = True
Play_Sound "Jet.wav"
End Sub

Sub tnkfromini(ini As String, side As String, ByVal X As Integer, ByVal Y As Integer, ByVal toX As Integer, ByVal toY As Integer)
On Error Resume Next
Dim img As String
Dim pow As Integer
Dim wpn As String
Dim spd As Integer
img = tanks.image(ini)
pow = tanks.power(ini)
wpn = tanks.weapon(ini)
spd = tanks.speed(ini)
MakeTank img, side, X, Y, toX, toY, pow, wpn, spd, ini
End Sub

Sub bldngfromini(ini As String, side As String, flip As Boolean, ByVal X As Integer, ByVal Y As Integer, Optional fade As Boolean = True)
On Error Resume Next
Dim img As String
Dim pow As Integer
Dim wpn As String
Dim offx As Integer
Dim offy As Integer
img = struc.image(ini)
pow = struc.power(ini)
wpn = struc.weapon(ini)
offx = struc.offx(ini)
offy = struc.offy(ini)
If LCase(struc.typ(ini)) = "techlab" And side = "Allies" Then
Map_Techlevel = Map_Techlevel + 1
ElseIf LCase(struc.typ(ini)) = "powerplant" And side = "Allies" Then
Map_Many = Map_Many + 1
End If
MakeBldng img, side, flip, X, Y, offx, offy, CLng(pow), wpn, ini, fade
End Sub

Sub MakeTank(image As String, side As String, ByVal X As Integer, ByVal Y As Integer, ByVal toX As Integer, ByVal toY As Integer, ByVal e_power As Long, e_weapon As String, ByVal e_speed As Integer, ini As String)
On Error Resume Next
Load gizzly(gizzly.UBound + 1)
Load Timer1(gizzly.UBound)
Load Line1(gizzly.UBound)
img(gizzly.UBound) = image
weapon(gizzly.UBound) = e_weapon
speed(gizzly.UBound) = e_speed
tnk_name(gizzly.UBound) = ini
power(gizzly.UBound) = e_power
t_power(gizzly.UBound) = e_power
gizzly(gizzly.UBound).Left = X
team(gizzly.UBound) = side
gizzly(gizzly.UBound).Top = Y
Rotate (360 - Angle(X, Y, toX, toY)), img(gizzly.UBound), gizzly.UBound
TnkMove gizzly.UBound, toX, toY
gizzly(gizzly.UBound).ZOrder 0
gizzly(gizzly.UBound).Visible = True
End Sub

Sub MakeBldng(image As String, side As String, flip As Boolean, X As Integer, Y As Integer, ByVal offx As Integer, ByVal offy As Integer, e_power As Long, e_weapon As String, ini As String, Optional fade As Boolean = True)
On Error Resume Next
Load bldng(bldng.UBound + 1)
bldng_pow(bldng.UBound) = e_power
bldng_tpow(bldng.UBound) = e_power
bldng_offsetX(bldng.UBound) = offx
bldng_OffsetY(bldng.UBound) = offy
bldng_name(bldng.UBound) = ini
bldng_team(bldng.UBound) = side
bldng_weapon(bldng.UBound) = e_weapon
bldng(bldng.UBound).Top = Y
bldng(bldng.UBound).Left = X
bldng(bldng.UBound).AutoSize = True
bldng_name(bldng.UBound) = ini
If flip = True Then bldng(bldng.UBound).Mirror = aiMirrorHorizontal
bldng(bldng.UBound).IntensityOffset = map_light
bldng(bldng.UBound).LoadImage_FromFile App.path & "\Images\Buildings\" & image & ".png"
If fade = True Then
bldng(bldng.UBound).Opacity = 0
bldng(bldng.UBound).FadeInOut 100
End If
bldng(bldng.UBound).Visible = True
Play_Sound "make.wav"
End Sub

Sub TnkMove(Index As Integer, ByVal LocX As Integer, ByVal LocY As Integer)
On Error Resume Next
Line1(Index).X2 = LocX
Line1(Index).Y2 = LocY
Line1(Index).X1 = gizzly(Index).Left + (gizzly(Index).Width / 2)
Line1(Index).Y1 = gizzly(Index).Top + (gizzly(Index).Height / 2)
k(Index) = 0
Rotate Angle(Line1(Index).X1, Line1(Index).Y1, Line1(Index).X2, Line1(Index).Y2), img(Index), Index
Timer1(Index) = True
Timer1_Timer Index
End Sub

Sub airMove(Index As Integer, ByVal LocX As Integer, ByVal LocY As Integer)
On Error Resume Next
Line1(Index).X2 = LocX
Line1(Index).Y2 = LocY
Line1(Index).X1 = gizzly(Index).Left + (gizzly(Index).Width / 2)
Line1(Index).Y1 = gizzly(Index).Top + (gizzly(Index).Height / 2)
k(Index) = 0
Rotate Angle(Line1(Index).X1, Line1(Index).Y1, Line1(Index).X2, Line1(Index).Y2), img(Index), Index
Timer1(Index) = True
Timer1_Timer Index
End Sub

Private Sub Form_Resize()
On Error Resume Next
Picture1.Left = 0
Picture1.Top = 0
BldBar.Top = (Me.Height / 15) - BldBar.Height
Piclbl.Move (Me.Width / 15) / 2 - Piclbl.Width / 2, (Me.Height / 15) / 2 - Piclbl.Height / 2
End Sub

Private Sub Gizzly_Click(Index As Integer)
On Error Resume Next
Dim nm As Integer, rad As Integer
If mode = 2 Then GoTo U
If team(Index) = "Allies" Then
If lstsel.ListCount = 0 Then
lstsel.AddItem Index
End If
Else
Dim X As Integer, Y As Integer, wid As Long, tr As Boolean, ene As Integer, conf As Boolean: tr = False
For nm = 0 To lstsel.ListCount - 1
ene = Val(lstsel.List(nm))
rad = weap.range(weapon(ene))
If (gizzly(ene).Left + gizzly(ene).Width / 2) > (gizzly(Index).Left + gizzly(Index).Width / 2) - rad And (gizzly(ene).Left + gizzly(ene).Width / 2) < (gizzly(Index).Left + gizzly(Index).Width / 2) + rad Then
If (gizzly(ene).Top + gizzly(ene).Height / 2) > (gizzly(Index).Top + gizzly(Index).Height / 2) - rad And (gizzly(ene).Top + gizzly(ene).Height / 2) < (gizzly(Index).Top + gizzly(Index).Height / 2) + rad Then
conf = True
End If
End If
If conf = False Then
With lnmsc
.X1 = gizzly(ene).Left - (gizzly(ene).Width / 2): .Y1 = gizzly(ene).Top - (gizzly(ene).Height / 2)
.X2 = gizzly(Index).Left - (gizzly(Index).Height / 2): .Y2 = gizzly(Index).Top - (gizzly(Index).Height / 2)
wid = Hyp(Wline(.X1, .X2), Hline(.Y1, .Y2))
Y = -(Sin(PI * Angle(.X2, .Y2, .X1, .Y1) / 180) * (rad - weap.distance(weapon(ene)))) + (gizzly(Index).Top + gizzly(Index).Height / 2)
X = Cos(PI * Angle(.X2, .Y2, .X1, .Y1) / 180) * (rad - weap.distance(weapon(ene))) + (gizzly(Index).Left + gizzly(Index).Width / 2)
TnkMove ene, X, Y
gizzly(ene).Tag = "tnk" & Index
End With
End If
Next
End If
Exit Sub
U:
If Map_Money - aeros.cost(DVW.sel) < 0 Then
Play_Sound "Warning.wav"
Msg "Not Enough Credits, You need : " & CStr(aeros.cost(DVW.sel))
mode = 0
cursor = "Pointer"
Else
AirMission Host, "Allies", DVW.sel, gizzly(Index).Left + gizzly(Index).Width / 2, gizzly(Index).Top + gizzly(Index).Height / 2
Map_Money = Map_Money - aeros.cost(DVW.sel)
mode = 0
cursor = "Pointer"
End If
End Sub

Private Sub gizzly_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim n As Integer
If team(Index) = "Allies" Then
For n = 0 To lstsel.ListCount - 1
TnkMove Val(lstsel.List(n)), gizzly(Index).Left + X / 15, gizzly(Index).Top + Y / 15
Next
End If
End Sub

Private Sub gizzly_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Index > 0 And pvw <> "gizz" & Index Then
pvw = "gizz" & Index
Bar.SetPro power(Index), t_power(Index)
Bar.Left = (gizzly(Index).Left + gizzly(Index).Width / 2) - (Bar.Width / 2)
Bar.Top = gizzly(Index).Top - 12
End If
End Sub

Private Sub Minimap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
LoadCursor "Mini", Minimap.hwnd
If Button = 1 Then
Dim LocX As Integer
Dim LocY As Integer
LocX = X
LocY = Y
If LocX < loc.Width / 2 Then
LocX = 0
ElseIf LocX > Minimap.Width - (loc.Width / 2) Then
LocX = Minimap.Width
ElseIf LocX > Minimap.Width Then
LocX = Minimap.Width
End If
If LocY < loc.Height / 2 Then
LocY = 0
End If
Picture1.Left = (-LocX / 15) * 10
Picture1.Top = (-LocY / 15) * 10
loc.Left = (-Picture1.Left * 15) / ((Picture1.Width / 150))
loc.Top = (-Picture1.Top * 15) / ((Picture1.Height / 150))
End If

End Sub

Private Sub mmc_Done(NotifyCode As Integer)
On Error Resume Next
If Val(GetFromIni("Main", "Music", App.path & "\set.cfg")) = 1 Then
mmc.FileName = App.path & "\Trax\Track1.mp3"
mmc.Command = "Open"
mmc.Command = "Play"
End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyEscape Then
Picture2.Visible = True
Picture2.Move (Me.Width / 15) / 2 - Picture2.Width / 2
button1.SetFocus
doall False
ElseIf KeyCode = vbKeyF8 Then
Text1.Visible = Not Text1.Visible
Text1.Move (Me.Width / 15) / 2 - Text1.Width / 2
Text1.SetFocus
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 And mode = 0 Then
For n = 0 To lstsel.ListCount - 1
If lstsel.ListCount = 0 Then
TnkMove Val(lstsel.List(n)), X, Y
Else
TnkMove Val(lstsel.List(n)), X + n * Rndmz(1) * 3, Y + n * Rndmz(1) * 3
End If
Next
sel.Left = X
sel.Top = Y
sel.Width = 1
sel.Height = 1
sel.Visible = True
mX = X
mY = Y
End If

If Button = 1 And mode = 1 Then
If Map_Money - struc.cost(DVW.sel) < 0 Then
Play_Sound "Warning.wav"
Msg "Not Enough Credits, You need : " & CStr(struc.cost(DVW.sel))
mode = 0
cursor = "Pointer"
Else
bldngfromini DVW.sel, "Allies", False, X, Y
Map_Money = Map_Money - struc.cost(DVW.sel)
mode = 0
cursor = "Pointer"
End If
ElseIf Button = 2 And mode = 1 Then
mode = 0
cursor = "Pointer"
End If

If Button = 1 And mode = 2 Then
If Map_Money - aeros.cost(DVW.sel) < 0 Then
Msg "Not Enough Credits, You need : " & CStr(aeros.cost(DVW.sel))
mode = 0
Play_Sound "Warning.wav"
cursor = "Pointer"
Else
AirMission Host, "Allies", DVW.sel, X, Y
Map_Money = Map_Money - aeros.cost(DVW.sel)
mode = 0
cursor = "Pointer"
End If
ElseIf Button = 2 And mode = 2 Then
mode = 0
cursor = "Pointer"
End If
lblair.Visible = False
lblbldng.Visible = False
If Button = 2 Then
lstsel.clear
cursor = "Pointer"
End If
LoadCursor cursor, Picture1.hwnd
End Sub

Function Rndmz(Seed As Integer) As Long
On Error Resume Next
Rndmz = Sgn(Rnd(Seed) - Rnd(Seed)) * Rnd(Seed)
End Function

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim wX As Integer, wY As Integer
If Button = 1 Then
wX = X - sel.Left
wY = Y - sel.Top
If Sgn(wX) = 1 And Sgn(wY) = 1 Then
sel.Left = mX
sel.Top = mY
sel.Width = wX
sel.Height = wY
ElseIf Sgn(wX) = -1 And Sgn(wY) = 1 Then
sel.Left = X
sel.Top = mY
sel.Width = mX - sel.Left
sel.Height = wY
ElseIf Sgn(wX) = -1 And Sgn(wY) = -1 Then
sel.Width = mX - X
sel.Height = mY - Y
sel.Left = X
sel.Top = Y
ElseIf Sgn(wX) = 1 And Sgn(wY) = -1 Then
sel.Top = Y
sel.Left = mX
sel.Height = mY - sel.Top
sel.Width = wX
End If
End If

If mode = 1 Then
lblbldng.Visible = True
lblbldng.Left = X + 20
lblbldng.Top = Y + 20
ElseIf mode = 2 Then
lblair.Visible = True
lblair.Left = X + 20
lblair.Top = Y + 20
End If
LoadCursor cursor, Picture1.hwnd

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
sel.Visible = False
Dim unt As Integer, n As Integer
For unt = gizzly.LBound To gizzly.UBound
If team(unt) = "Allies" Then
If gizzly(unt).Left + gizzly(unt).Width / 2 > sel.Left And gizzly(unt).Left + gizzly(unt).Width / 2 < sel.Left + sel.Width Then
If gizzly(unt).Top + gizzly(unt).Height / 2 > sel.Top And gizzly(unt).Top + gizzly(unt).Height / 2 < sel.Top + sel.Height Then
lstsel.AddItem unt
n = n + 1
End If: End If: End If
Next
If lstsel.ListCount > 0 Then
cursor = "Move"
Else
cursor = "Pointer"
End If
If Button = 2 Then
cursor = "Pointer"
End If
End Sub

Private Sub t2b_Timer()
On Error Resume Next
Dim unt As Integer
Dim e_unt As Integer
Dim rng As Integer
For unt = gizzly.LBound To gizzly.UBound
For e_unt = bldng.LBound To bldng.UBound
If unt = 0 Or e_unt = 0 Or team(unt) = "-1" Or bldng_team(e_unt) = "-1" Then GoTo X
rng = weap.range(weapon(unt))
If bldng(e_unt).Left > gizzly(unt).Left - rng And bldng(e_unt).Left < gizzly(unt).Left + rng Then
If bldng(e_unt).Top > gizzly(unt).Top - rng And bldng(e_unt).Top < gizzly(unt).Top + rng Then
If team(unt) <> bldng_team(e_unt) Then

If gizzly(unt).ToolTipText = "" Then
DoEvents
If gizzly(unt).Tag = "" Then
gizzly(unt).Tag = "bldng" & CStr(e_unt)
End If

If gizzly(unt).Tag = "bldng" & CStr(e_unt) Then
If weap.typ(weapon(unt)) = "bomb" Then
fire unt, e_unt, "", "bld"
ElseIf weap.typ(weapon(unt)) = "laser" Then
Laser unt, e_unt, weap.damage(weapon(unt)), weap.Color(weapon(unt)), 2
End If
End If

End If
End If
End If
End If
X:
Next: Next
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
Trig Text1
End If
End Sub

Private Sub Timer_Timer()
On Error Resume Next
Dim k As Integer
time = time - 1
If time <= 0 Then
For k = 1 To Trig_Count
Trig (trigger(k))
Next
If t_loop = True Then
time = S_time
Timer = True
Else
Timer = False
End If
End If
End Sub

Private Sub Timer1_Timer(Index As Integer)
On Error Resume Next
Dim X As Long, Y As Long, wid As Long, ang As Integer, Tk As Integer, col As Boolean, mD As Integer, SS As Integer
wid = Hyp(Wline(Line1(Index).X1, Line1(Index).X2), Hline(Line1(Index).Y1, Line1(Index).Y2))
ang = Angle(Line1(Index).X1, Line1(Index).Y1, Line1(Index).X2, Line1(Index).Y2)
If k(Index) <= wid Then
Y = -(Sin(PI * ang / 180) * k(Index)) + Line1(Index).Y1
X = Cos(PI * ang / 180) * k(Index) + Line1(Index).X1
For Tk = 1 To bldng.UBound
If bldng_team(Tk) <> "-1" Then
If LCase(struc.typ(bldng_name(Tk))) = "wall" Then mD = SS = 0 Else mD = (bldng(Tk).Height - (bldng(Tk).Height / 3.5)): SS = speed(Index) + 2
If X > bldng(Tk).Left + SS And X < bldng(Tk).Left + bldng(Tk).Width - SS Then
If Y > bldng(Tk).Top + mD + SS And Y < bldng(Tk).Top + bldng(Tk).Height - SS Then
If LCase(struc.typ(bldng_name(Tk))) <> "warfactory" Then
col = True
DoEvents
Exit For
End If
ElseIf Y > bldng(Tk).Top + SS And Y < bldng(Tk).Top + bldng(Tk).Height - SS Then
bldng(Tk).ZOrder 0
DoEvents
Exit For
End If
End If
End If
Next
If col = True Then
GoTo H
Else
gizzly(Index).Move X - (gizzly(Index).Width / 2), Y - (gizzly(Index).Height / 2)
k(Index) = k(Index) + speed(Index)
Timer1(Index).interval = 1
gizzly(Index).ToolTipText = " "
DoEvents
End If
DoEvents
Else
H:
gizzly(Index).ToolTipText = ""
Timer1(Index) = False
k(Index) = 0
DoEvents
End If
End Sub

Sub fire(from As Integer, too As Integer, so As String, dest As String, Optional tgtX As Integer = 1, Optional tgtY As Integer = 1)
On Error Resume Next
Dim n As Integer
'On Error GoTo Yo
Load lnbom(lnbom.UBound + 1)
Load BomTmr(BomTmr.UBound + 1)
Load bomb(bomb.UBound + 1)
BomTmr(BomTmr.UBound).Tag = BomTmr.UBound
With lnbom(BomTmr.UBound)
If so = "bld" Then
.X1 = bldng(from).Left + (bldng(from).Width / 2) + bldng_offsetX(from)
.Y1 = bldng(from).Top + (bldng(from).Height / 2) + bldng_OffsetY(from)
bk(bomb.UBound) = weap.distance(bldng_weapon(from))
bomb(bomb.UBound).Tag = "bld" & too
ElseIf so = "air" Then
.X1 = air(from).Left + (air(from).Width / 2)
.Y1 = air(from).Top + (air(from).Height / 2)
bk(bomb.UBound) = weap.distance(air_weapon(from))
bomb(bomb.UBound).Tag = "air" & too
Else
.X1 = gizzly(from).Left + (gizzly(from).Width / 2)
.Y1 = gizzly(from).Top + (gizzly(from).Height / 2)
bk(bomb.UBound) = weap.distance(weapon(from))
bomb(bomb.UBound).Tag = too
End If
If dest = "bld" Then
.X2 = bldng(too).Left + bldng(too).Width / 2
.Y2 = bldng(too).Top + bldng(too).Height / 2
.Tag = "bld" & from
ElseIf dest = "air" Then
.X2 = air(too).Left + air(too).Width / 2
.Y2 = air(too).Top + air(too).Height / 2
.Tag = "air" & from
Else
If so <> "air" Then
.X2 = gizzly(too).Left + gizzly(too).Width / 2
.Y2 = gizzly(too).Top + gizzly(too).Height / 2
.Tag = from
Else
.X2 = tgtX
.Y2 = tgtY
.Tag = from
End If
End If
If so <> "bld" Or so <> "air" Then
Rotate Angle(.X1, .Y1, .X2, .Y2), img(from), from
End If
BomTmr(BomTmr.UBound) = True
tmrpvw = True
DoEvents
End With
End Sub

Sub Rotate(ByVal Deg As Integer, key As String, idx As Integer)
On Error Resume Next
Dim n As Integer
n = Rndeg(Deg)
If Angl(idx) <> n Then
Set gizzly(idx).Picture = LoadPicture(App.path & "\Images\" & key & "\" & key & n & " copy.gif")
Angl(idx) = n
End If
End Sub

Sub Rotate_air(ByVal Deg As Integer, key As String, idx As Integer)
On Error Resume Next
Dim n As Integer
n = Rndeg(Deg)
Set air(idx).Picture = LoadPicture(App.path & "\Images\" & key & "\" & key & n & " copy.gif")
End Sub

Function Rndeg(ByVal Deg As Integer) As Integer
On Error Resume Next
Rndeg = Round(Deg / 18)
End Function
Sub Laser(from As Integer, too As Integer, damage As Integer, Color As Long, frombldng1tobldng2else3 As Integer, Optional aircraft As Boolean = False, Optional airX As Integer = 0, Optional airY As Integer = 0)
On Error Resume Next
Dim n As Integer
n = bk(bomb.UBound + 1)
bk(BomTmr.UBound) = 0
BomTmr(BomTmr.UBound).Tag = BomTmr.UBound
With lssr
If aircraft = True Then GoTo U:
If frombldng1tobldng2else3 = 3 Then
.X1 = gizzly(from).Left + (gizzly(from).Width / 2)
.Y1 = gizzly(from).Top + (gizzly(from).Height / 2)
.X2 = gizzly(too).Left + gizzly(too).Width / 2
.Y2 = gizzly(too).Top + gizzly(too).Height / 2
Rotate Angle(.X1, .Y1, .X2, .Y2), img(from), from
power(too) = power(too) - damage
If power(too) <= 0 Then
gizzly(from).Tag = ""
desttank (too)
End If
ElseIf frombldng1tobldng2else3 = 1 Then
.X1 = bldng(from).Left + (bldng(from).Width / 2) + bldng_offsetX(from)
.Y1 = bldng(from).Top + (bldng(from).Height / 2) + bldng_OffsetY(from)
.X2 = gizzly(too).Left + gizzly(too).Width / 2
.Y2 = gizzly(too).Top + gizzly(too).Height / 2
power(too) = power(too) - damage
If power(too) <= 0 Then
desttank (too)
bldng(from).Tag = ""
End If
Else
.X1 = gizzly(from).Left + (gizzly(from).Width / 2)
.Y1 = gizzly(from).Top + (gizzly(from).Height / 2)
.X2 = bldng(too).Left + bldng(too).Width / 2
.Y2 = bldng(too).Top + bldng(too).Height / 2
Rotate Angle(.X1, .Y1, .X2, .Y2), img(from), from
bldng_pow(too) = bldng_pow(too) - damage
If bldng_pow(too) <= 0 Then
destruct (too)
gizzly(from).Tag = ""
End If
End If
GoTo v
U:
.X1 = air(from).Left + air(from).Width / 2
.Y1 = air(from).Top + air(from).Height / 2
.X2 = airX
.Y2 = airY
explode .X2, .Y2, damage / 4, damage
v:
.BorderColor = Color
.Visible = True
.Refresh
spot.Left = .X2 - spot.Width / 2
spot.Top = .Y2 - spot.Height / 2
spot.Visible = True
spot.Opacity = 100
spot.LoadImage_FromFile App.path & "\animations\spot.png"
spot.FadeInOut 0, 10, 90
DoEvents
.Visible = False
Play_Sound "Laser.wav"
End With
End Sub

Sub sblast(ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next
sbl.Visible = True
For R = 0 To 5
DoEvents
Set sbl.Picture = LoadPicture(App.path & "\animations\s" & R & ".gif")
DoEvents
sbl.Top = Y - sbl.Height
sbl.Left = X - sbl.Width / 2
sbl.Visible = True
Next
Set sbl.Picture = Nothing
sbl.Visible = False
End Sub

Sub bblast(ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next
Dim n As Integer
sbl.Visible = True
tmrpvw = False
For n = 0 To 38
Set sbl.Picture = LoadPicture(App.path & "\animations\Namim (" & CStr(n) & ").gif")
sbl.ZOrder 0
sbl.Top = Y - sbl.Width / 2
sbl.Left = X - sbl.Height / 2
DoEvents
Next
sbl.Visible = False
Set sbl.Picture = Nothing
tmrpvw = True
End Sub

Private Sub TmrMon_Timer()
On Error Resume Next
Map_Money = Map_Money + (Map_Many * 10)
End Sub

Private Sub tmrpth_Timer(Index As Integer)
On Error Resume Next
' Dock carrier to air
Dim X As Long, Y As Long, wid As Long, ang As Integer
If tmrpth(Index).Tag = "1" Then
wid = Hyp(Wline(Dpth(Index).X1, Dpth(Index).X2), Hline(Dpth(Index).Y1, Dpth(Index).Y2))
ang = Angle(Dpth(Index).X1, Dpth(Index).Y1, Dpth(Index).X2, Dpth(Index).Y2)
If air_k(Index) <= wid Then
Y = -(Sin(PI * ang / 180) * air_k(Index)) + Dpth(Index).Y1
X = Cos(PI * ang / 180) * air_k(Index) + Dpth(Index).X1
air(Index).Move X - (air(Index).Width / 2), Y - (air(Index).Height / 2)
air_k(Index) = air_k(Index) + air_speed(Index)
DoEvents
Else
tmrpth(Index).Tag = "2"
air_k(Index) = 0
Rotate_air Angle(pth(Index).X1, pth(Index).Y1, pth(Index).X2, pth(Index).Y2), air_img(Index), Index
DoEvents
End If
End If

'Air to target location

If tmrpth(Index).Tag = "2" Then
wid = Hyp(Wline(pth(Index).X1, pth(Index).X2), Hline(pth(Index).Y1, pth(Index).Y2))
ang = Angle(pth(Index).X1, pth(Index).Y1, pth(Index).X2, pth(Index).Y2)
If air_k(Index) <= wid - 150 Then
Y = -(Sin(PI * ang / 180) * air_k(Index)) + pth(Index).Y1
X = Cos(PI * ang / 180) * air_k(Index) + pth(Index).X1
air(Index).Move X - (air(Index).Width / 2), Y - (air(Index).Height / 2)
air_k(Index) = air_k(Index) + air_speed(Index)
DoEvents
Else
tmrpth(Index).Tag = "3"
If LCase(weap.typ(air_weapon(Index))) = "laser" Then
Laser Index, 0, weap.damage(air_weapon(Index)), weap.Color(air_weapon(Index)), 0, True, pth(Index).X2, pth(Index).Y2
Else
fire Index, 1, "air", "", pth(Index).X2, pth(Index).Y2
End If
Rotate_air (Angle(pth(Index).X2, pth(Index).Y2, pth(Index).X1, pth(Index).Y1)), air_img(Index), Index
DoEvents
End If
End If

'Return to the landing position

If tmrpth(Index).Tag = "3" Then
wid = Hyp(Wline(pth(Index).X1, pth(Index).X2), Hline(pth(Index).Y1, pth(Index).Y2))
ang = Angle(pth(Index).X1, pth(Index).Y1, pth(Index).X2, pth(Index).Y2)
If air_k(Index) > 0 Then
Y = -(Sin(PI * ang / 180) * air_k(Index)) + pth(Index).Y1
X = Cos(PI * ang / 180) * air_k(Index) + pth(Index).X1
air(Index).Move X - (air(Index).Width / 2), Y - (air(Index).Height / 2)
air_k(Index) = air_k(Index) - air_speed(Index)
DoEvents
Else
tmrpth(Index).Tag = "4"
air_k(Index) = Hyp(Wline(Dpth(Index).X1, Dpth(Index).X2), Hline(Dpth(Index).Y1, Dpth(Index).Y2))
Rotate_air Angle(Dpth(Index).X2, Dpth(Index).Y2, Dpth(Index).X1, Dpth(Index).Y1), air_img(Index), Index
DoEvents
End If
End If

'Landing

If tmrpth(Index).Tag = "4" Then
wid = Hyp(Wline(Dpth(Index).X1, Dpth(Index).X2), Hline(Dpth(Index).Y1, Dpth(Index).Y2))
ang = Angle(Dpth(Index).X1, Dpth(Index).Y1, Dpth(Index).X2, Dpth(Index).Y2)
If air_k(Index) > 0 Then
Y = -(Sin(PI * ang / 180) * air_k(Index)) + Dpth(Index).Y1
X = Cos(PI * ang / 180) * air_k(Index) + Dpth(Index).X1
air(Index).Move X - (air(Index).Width / 2), Y - (air(Index).Height / 2)
air_k(Index) = air_k(Index) - air_speed(Index)
DoEvents
Else
Unload air(Index)
Unload tmrpth(Index)
Unload pth(Index)
air_k(1000) = 0
air_img(1000) = ""
air_speed(1000) = 0
air_power(1000) = 0
air_t_power(1000) = 100
air_weapon(1000) = ""
air_team(1000) = -1
air_Angl(1000) = 0
DoEvents
End If
End If
End Sub

Private Sub tmrpvw_Timer()
On Error GoTo G
Bar.Visible = True
H:
On Error Resume Next
Dim str As Integer
Dim mode As String
If Left(pvw, 4) = "gizz" Then
str = Val(Right(pvw, Len(pvw) - 4))
Bar.SetPro CStr(power(str)), CStr(t_power(str))
Bar.Left = (gizzly(str).Left + gizzly(str).Width / 2) - (Bar.Width / 2)
Bar.Top = gizzly(str).Top - 12
ElseIf Left(pvw, 5) = "bldng" Then
str = Val(Right(pvw, Len(pvw) - 5))
Bar.SetPro CStr(bldng_pow(str)), CStr(bldng_tpow(str))
Bar.Left = (bldng(str).Left + bldng(str).Width / 2) - (Bar.Width / 2)
Bar.Top = bldng(str).Top - 12
End If
DoEvents
Label1 = Map_Money
Label2 = label & CStr(time)
lbltl = CStr(Map_Techlevel)
DoEvents
Exit Sub
G:
pvw = 0
GoTo H
End Sub

Sub Decbombs()
On Error Resume Next
Dim n As Integer
For n = bomb.LBound To bomb.UBound
If IsObject(bomb(n)) = True Then
If IsObject(gizzly(getval(lnbom(n).Tag))) = False Or IsObject(bldng(getval(lnbom(n).Tag))) = False Then
Unload lnbom(n)
Unload BomTmr(n)
Unload bomb(n)
End If
End If
Next
End Sub

Private Sub TmrRef_Timer()
If LCase(map_condition) <> "storm" Then
Minimap_Refresh
End If
End Sub

Private Sub tmrSlip_Timer()
On Error Resume Next
GetCursorPos pt
If pt.X < 3 Then
If Picture1.Left < 0 Then
Picture1.Left = Picture1.Left + 15
loc.Left = (-Picture1.Left * 15) / ((Picture1.Width / 150))
End If
ElseIf pt.Y < 3 Then
If Picture1.Top < 0 Then
Picture1.Top = Picture1.Top + 15
loc.Top = (-Picture1.Top * 15) / ((Picture1.Height / 150))
End If
ElseIf pt.X > (Screen.Width / 15) - 3 Then
If Picture1.Left + Picture1.Width > Me.Width / 15 Then
Picture1.Left = Picture1.Left - 15
loc.Left = (-Picture1.Left * 15) / ((Picture1.Width / 150))
End If
ElseIf pt.Y > (Screen.Height / 15) - 3 Then
If Picture1.Top + Picture1.Height > Me.Height / 15 Then
Picture1.Top = Picture1.Top - 15
loc.Top = (-Picture1.Top * 15) / ((Picture1.Height / 150))
End If
End If
End Sub

Sub RemoveAll()
On Error Resume Next
Dim X As Integer
For X = gizzly.LBound To gizzly.UBound
If X = 0 Then GoTo X
Unload gizzly(X)
img(X) = ""
speed(X) = 0
tnk_name(X) = ""
power(X) = 0
t_power(X) = 100
weapon(X) = ""
team(X) = "-1"
typ(X) = ""
Angl(X) = ""
X:
Next
For X = bldng.LBound To bldng.UBound
If X = 0 Then GoTo Y
Unload bldng(X)
bldng_pow(X) = 0
bldng_name(X) = "'"
bldng_tpow(X) = 100
bldng_team(X) = "-1"
bldng_offsetX(X) = 0
bldng_OffsetY(X) = 0
bldng_weapon(X) = ""
Y:
Next
For X = tree.LBound To tree.UBound
If X = 0 Then GoTo Z
Unload tree(X)
Z:
Next
End Sub
Sub LoadMap(str As String)
On Error Resume Next
Dim X As Integer
Dim cX As Integer, cY As Integer
Dim ini As String
Dim tex As String
Dim Dstr As String
Dim flip As Boolean
ini = str & "\ini.ini"
RemoveAll
doall False
Hide
Picture1.Left = 0
Picture1.Top = 0
Picture1.Width = Val(GetFromIni("Main", "Width", ini))
Picture1.Height = Val(GetFromIni("Main", "Height", ini))
Map_Name = GetFromIni("Main", "Name", ini)
Map_Money = GetFromIni("Main", "Money", ini)
map_condition = GetFromIni("Main", "Map Condition", ini)
tex = GetFromIni("Main", "ground", ini)
map_light = GetFromIni("Main", "LightOffset", ini)
AI_Skills = GetFromIni("Main", "AI Skills", ini)
Dstr = Left(str, Len(str) - 1)
Map_Level = Val(Right(Dstr, 1))

Picture1.Refresh

If Val(GetFromIni("Main", "Ground", App.path & "\set.cfg")) = 1 Then
DoEvents
Picture1.AutoRedraw = True
Set Picture1.Picture = Nothing
Map_Ground.InitializeDIB Picture1.Width, Picture1.Height
Map_Ground.LoadPicture_File App.path & "\Images\Texture\" & tex
For cX = 0 To Picture1.Width Step 256
For cY = 0 To Picture1.Height Step 256
Map_Ground.Render Picture1.hdc, cX, cY, , , , , , , , , , , , map_light
DoEvents
Next: Next
Picture1.Picture = Picture1.image
Map_Ground.DestroyDIB
Set Map_Ground = Nothing
Else
Picture1.BackColor = &H467113
End If

Picture1.AutoRedraw = False

Trig_Count = Val(GetFromIni("Timer", "count", ini))
For X = 1 To Trig_Count
trigger(X) = GetFromIni("Timer", "Trigger" & CStr(X), ini)
Next
label = GetFromIni("Timer", "label", ini)
time = Val(GetFromIni("Timer", "time", ini))
S_time = Val(GetFromIni("Timer", "time", ini))
t_loop = str2bol(GetFromIni("Timer", "loop", ini))

Map_Eves_Count = Val(GetFromIni("Events", "Count", ini))
For X = 1 To Val(GetFromIni("Events", "Count", ini))
Map_On(X) = GetFromIni("Events", "On" & CStr(X), ini)
Map_Do(X) = GetFromIni("Events", "Do" & CStr(X), ini)
Next

loc.Width = Screen.Width / ((Picture1.Width / 150))
loc.Height = Screen.Height / ((Picture1.Height / 150))

For X = 1 To Val(GetFromIni("Tanks", "Count", ini))
tnkfromini GetFromIni("Tanks", "ini" & CStr(X), ini), GetFromIni("Tanks", "side" & CStr(X), ini), Val(GetFromIni("Tanks", "X" & CStr(X), ini)), Val(GetFromIni("Tanks", "Y" & CStr(X), ini)), Val(GetFromIni("Tanks", "toX" & CStr(X), ini)), Val(GetFromIni("Tanks", "toY" & CStr(X), ini))
Next

Dim getRev As Boolean
If Val(GetFromIni("Main", "Reveal", App.path & "\set.cfg")) = 1 Then getRev = True Else getRev = False

For X = 1 To GetFromIni("Buildings", "Count", ini)
If GetFromIni("Buildings", "Flip" & CStr(X), ini) = "1" Then flip = True Else flip = False
bldngfromini GetFromIni("Buildings", "ini" & CStr(X), ini), GetFromIni("Buildings", "side" & CStr(X), ini), flip, GetFromIni("Buildings", "X" & CStr(X), ini), GetFromIni("Buildings", "Y" & CStr(X), ini), getRev
Next

For X = 1 To Val(GetFromIni("Trees", "Count", ini))
makeTree Val(GetFromIni("Trees", "TreeX" & CStr(X), ini)), Val(GetFromIni("Trees", "TreeY" & CStr(X), ini))
Next
Me.Show

doall True
GetSettings
Exit Sub
Y:
MsgBox Err.Description
End Sub

Sub GetSettings()
On Error Resume Next
If Val(GetFromIni("Main", "Minimap", App.path & "\set.cfg")) = 1 Then
TmrRef = False
Else
TmrRef = True
End If

If Val(GetFromIni("Main", "Music", App.path & "\set.cfg")) = 1 Then
mmc.Command = "Stop"
mmc.FileName = App.path & "\Trax\Track1.mp3"
mmc.Command = "Open"
mmc.Command = "Play"
End If
Dim Sspeed As Integer
Sspeed = Val(GetFromIni("Main", "Speed", App.path & "\set.cfg"))
auto_bldng.interval = auto_bldng.interval - (Sspeed * 4)
tmrpvw.interval = tmrpvw.interval - (Sspeed * 4)
auto_bldng.interval = auto_bldng.interval - (Sspeed * 4)
t2b.interval = t2b.interval - (Sspeed * 4)
tmrmon.interval = tmrmon.interval - (Sspeed * 4)
TmrRef.interval = TmrRef.interval - Sspeed

End Sub

Sub Trig(str As String) ' A mini but powerful command processor for events and triggers
' It splits the command to its name, arguments and brackets , Use commands like
'destunt(index) ; destbldng(index) ; makebldng(ini as string,side as string,flip as boolean,x as integer,y as integer) ;
'maketank(ini as string,side as string,x as integer,y as integer,toX as integer,toY as integer) ;
'airmission(dock as integer,side as sting,ini as string,x as integer,y as integer) ' Initializes an air mission on ...
'nmsl(x as integer,y as integer) ' Fires nuclear missile on X and Y
'Use strings without quotations
On Error Resume Next

Dim str1 As String
Dim arg(5) As String
If LCase(str) = "loose" Then
mmc.Command = "Stop"
Play_Sound "Menu.wav"
Set Piclbl.Picture = LoadPicture(App.path & "\Images\BuildBar\defeat.jpg")
Piclbl.Visible = True
doall False
frmmnu.SetLoose
Unload Me
ElseIf LCase(str) = "win" Then
Play_Sound "Menu.wav"
Set Piclbl.Picture = LoadPicture(App.path & "\Images\BuildBar\victor.jpg")
Piclbl.Visible = True
doall False
str1 = GetFromIni("Main", "Progress", App.path & "\set.cfg")
If Map_Level = Val(str1) And Map_Level <> 7 Then
WriteIni "Main", "Progress", CStr(Val(str1) + 1), App.path & "\set.cfg"
frmmnu.rtb.Visible = True
frmmnu.button1.caption "Next Mission"
frmmnu.rtb.Text = "Congrats Sir, You unlocked new mission"
ElseIf Map_Level = 7 Then
frmmnu.rtb.Visible = True
frmmnu.button1.caption "Replay Mission"
frmmnu.rtb.Text = "Congrats Sir, You successfuly completed the campaign"
Play_Sound "Extruct.wav"
End If
frmmnu.SetNext
Unload Me
ElseIf Left(LCase(str), 8) = "destunt(" Then
str = Right(str, Len(str) - 8)
str = Left(str, Len(str) - 1)
desttank Val(str) - 1
ElseIf Left(LCase(str), 10) = "destbldng(" Then
str = Right(str, Len(str) - 10)
str = Left(str, Len(str) - 1)
destruct Val(str) - 1
ElseIf Left(LCase(str), 10) = "makebldng(" Then
str = Right(str, Len(str) - 10)
str = Left(str, Len(str) - 1)
arg(0) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(0)) - 1)
arg(1) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(1)) - 1)
arg(2) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(2)) - 1)
arg(3) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(3)) - 1)
arg(4) = str
bldngfromini arg(0), arg(1), str2bol(arg(2)), arg(3), arg(4)
ElseIf Left(LCase(str), 9) = "maketank(" Then
str = Right(str, Len(str) - 9)
str = Left(str, Len(str) - 1)
arg(0) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(0)) - 1)
arg(1) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(1)) - 1)
arg(2) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(2)) - 1)
arg(3) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(3)) - 1)
arg(4) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(4)) - 1)
arg(5) = str
tnkfromini arg(0), arg(1), Val(arg(2)), Val(arg(3)), Val(arg(4)), Val(arg(5))
ElseIf Left(LCase(str), 11) = "airmission(" Then
str = Right(str, Len(str) - 11)
str = Left(str, Len(str) - 1)
arg(0) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(0)) - 1)
arg(1) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(1)) - 1)
arg(2) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(2)) - 1)
arg(3) = Left(str, InStr(1, str, ",") - 1)
str = Right(str, Len(str) - Len(arg(3)) - 1)
arg(4) = str
If bldng_team(Val(arg(0)) - 1) <> "-1" And LCase(struc.typ(bldng_name(Val(arg(0)) - 1))) = "acc" Then
AirMission Val(arg(0)) - 1, arg(1), arg(2), Val(arg(3)), Val(arg(4))
End If
ElseIf Left(LCase(str), 5) = "nmsl(" Then
str = Right(str, Len(str) - 5)
str = Left(str, Len(str) - 1)
arg(0) = Left(str, InStr(1, str, ",") - 1)
arg(1) = Right(str, Len(str) - Len(arg(0)) - 1)
Nmsl Val(arg(0)), Val(arg(1))
ElseIf Left(LCase(str), 5) = "bolt(" Then
str = Right(str, Len(str) - 5)
str = Left(str, Len(str) - 1)
arg(0) = Left(str, InStr(1, str, ",") - 1)
arg(1) = Right(str, Len(str) - Len(arg(0)) - 1)
bolt Val(arg(0)), Val(arg(1))
End If
End Sub

Sub Nmsl(X As Integer, Y As Integer)
On Error Resume Next
Dim k As Integer
Missile.AutoSize = True
Missile.LoadImage_FromFile App.path & "\animations\Nmsl Down.png"
Missile.ZOrder 0
Missile.Visible = True
For k = -Missile.Height To Y - Missile.Height Step 15
Missile.Top = k
Missile.Left = X - (Missile.Width / 2)
DoEvents
Next
Missile.Visible = False
bblast X, Y
DoEvents
explode X, Y, 512, 10000
shake 50
Play_Sound "ExStruct.wav"
End Sub

Function str2bol(str As String) As Boolean
On Error Resume Next
If LCase(str) = "true" Then
str2bol = True
ElseIf LCase(str) = "false" Then
str2bol = False
End If
End Function

Sub Bldbar_Set(Index As Integer)
On Error Resume Next
Dim k As Integer
Dim m_name As String
DVW.clear
If LCase(struc.typ(bldng_name(Index))) = "acc" Then
For k = 1 To GetFromIni("Main", "count", App.path & "\rules\aircrafts.ini")
m_name = GetFromIni("Main", "a" & CStr(k), App.path & "\rules\aircrafts.ini")
If aeros.techlevel(m_name) <> "-1" And aeros.techlevel(m_name) <= Map_Techlevel Then
DVW.Add m_name, aircraft
End If
Next
ElseIf LCase(struc.typ(bldng_name(Index))) = "constyard" Then
For k = 1 To GetFromIni("Main", "count", App.path & "\rules\buildings.ini")
m_name = GetFromIni("Main", "b" & CStr(k), App.path & "\rules\buildings.ini")
If struc.techlevel(m_name) <> "-1" And struc.techlevel(m_name) <= Map_Techlevel Then
DVW.Add m_name, building
End If
Next
ElseIf LCase(struc.typ(bldng_name(Index))) = "warfactory" Then
For k = 1 To GetFromIni("Main", "count", App.path & "\rules\tanks.ini")
m_name = GetFromIni("Main", "t" & CStr(k), App.path & "\rules\tanks.ini")
If tanks.techlevel(m_name) <> "-1" And tanks.techlevel(m_name) <= Map_Techlevel And tanks.water(m_name) = "0" Then
DVW.Add m_name, tank
End If
Next
Else
End If
DVW.Add "REPAIR", tank
End Sub

Sub Msg(str As String)
On Error Resume Next
lblmsg = str
End Sub

Function Stat2() As Integer
On Error Resume Next
If Sgn(Cos(n_rand)) = 1 Or Sgn(Cos(n_rand)) = 0 Then
Stat2 = 2
ElseIf Sgn(Cos(n_rand)) = -1 Then
Stat2 = 1
End If
n_rand = n_rand + 1
End Function

Sub Minimap_Refresh()
On Error Resume Next
Dim k As Integer, RatioX As Integer, RatioY As Integer, Color As Long
Set Minimap.Picture = Nothing
Minimap.Cls
Minimap.AutoRedraw = True
RatioX = Picture1.Width / 150
RatioY = Picture1.Height / 150
For k = 1 To bldng.UBound
If bldng_team(k) = "Allies" Then Color = vbBlue Else Color = vbRed
SetPixel Minimap.hdc, bldng(k).Left / RatioX, bldng(k).Top / RatioY, Color
SetPixel Minimap.hdc, (bldng(k).Left / RatioX) - 1, bldng(k).Top / RatioY, Color
SetPixel Minimap.hdc, (bldng(k).Left / RatioX) + 1, bldng(k).Top / RatioY, Color
SetPixel Minimap.hdc, bldng(k).Left / RatioX, (bldng(k).Top / RatioY) + 1, Color
SetPixel Minimap.hdc, bldng(k).Left / RatioX, (bldng(k).Top / RatioY) - 1, Color
Next
For k = 1 To gizzly.UBound
If team(k) = "Allies" Then Color = vbBlue Else Color = vbRed
SetPixel Minimap.hdc, gizzly(k).Left / RatioX, gizzly(k).Top / RatioY, Color
SetPixel Minimap.hdc, (gizzly(k).Left / RatioX) - 1, gizzly(k).Top / RatioY, Color
SetPixel Minimap.hdc, (gizzly(k).Left / RatioX) + 1, gizzly(k).Top / RatioY, Color
SetPixel Minimap.hdc, gizzly(k).Left / RatioX, (gizzly(k).Top / RatioY) + 1, Color
SetPixel Minimap.hdc, gizzly(k).Left / RatioX, (gizzly(k).Top / RatioY) - 1, Color
Next
'Minimap.Picture = Minimap.image
'Minimap.AutoRedraw = False
Exit Sub
End Sub

Function random()
On Error Resume Next
n_rand = n_rand + 1
If Sgn(Cos(n_rand)) = 1 Or Sgn(Cos(n_rand)) = 0 Then
random = n_rand
ElseIf Sgn(Cos(n_rand)) = -1 Then
random = n_rnd * n_rand
End If
End Function

Sub bolt(cX As Integer, cY As Integer)
On Error Resume Next
Dim k As Integer
Missile.AutoSize = True
Missile.Opacity = 100
spot.Opacity = 100
Missile.LoadImage_FromFile App.path & "\animations\ltng.png"
Missile.ZOrder 0
Missile.Visible = True

Missile.Top = cY - Missile.Height
Missile.Left = cX - (Missile.Width / 2)

sblast cX, cY
DoEvents
explode cX, cY, 75, 2000
spot.Top = cY - (spot.Height / 2)
spot.Left = cX - (spot.Width / 2)
Missile.FadeInOut 0
spot.FadeInOut 0
shake 5
Play_Sound "ExStruct.wav"
End Sub

Sub shake(power As Integer)
Picture1.Left = Picture1.Left - power
Picture1.Top = Picture1.Top - power
Picture1.Left = Picture1.Left + power
Picture1.Left = Picture1.Top + power
End Sub

VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   702
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrmon 
      Interval        =   1000
      Left            =   4440
      Top             =   0
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   0
   End
   Begin VB.Timer tmrpth 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   3000
      Top             =   0
   End
   Begin VB.Timer tmrpvw 
      Interval        =   1
      Left            =   2040
      Top             =   0
   End
   Begin VB.Timer t2b 
      Interval        =   300
      Left            =   2520
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   600
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer auto 
      Interval        =   600
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer BomTmr 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   20
      Left            =   1560
      Top             =   0
   End
   Begin VB.Timer auto_bldng 
      Interval        =   300
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer tmrSlip 
      Interval        =   1
      Left            =   3480
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   327
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   599
      TabIndex        =   12
      Top             =   360
      Width           =   9015
      Begin VB.ListBox lstsel 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   1560
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin Project1.Bar Bar 
         Height          =   45
         Left            =   1080
         TabIndex        =   13
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   79
      End
      Begin Project1.aeroplanes aeros 
         Height          =   495
         Left            =   2160
         TabIndex        =   14
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin Project1.struc struc 
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin Project1.tanks tanks 
         Height          =   495
         Left            =   1200
         TabIndex        =   16
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin Project1.weapons weap 
         Height          =   495
         Left            =   720
         TabIndex        =   17
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
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
      Begin VB.Line Line1 
         Index           =   0
         Visible         =   0   'False
         X1              =   192
         X2              =   320
         Y1              =   216
         Y2              =   88
      End
      Begin VB.Line lnmsc 
         BorderColor     =   &H00FF0000&
         Visible         =   0   'False
         X1              =   344
         X2              =   312
         Y1              =   112
         Y2              =   80
      End
      Begin VB.Image gizzly 
         Height          =   1935
         Index           =   0
         Left            =   3120
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image sbl 
         Height          =   1935
         Left            =   3000
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Shape bomb 
         FillColor       =   &H00424242&
         FillStyle       =   0  'Solid
         Height          =   90
         Index           =   0
         Left            =   840
         Shape           =   2  'Oval
         Top             =   3360
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Shape sel 
         BorderStyle     =   3  'Dot
         Height          =   1935
         Left            =   2760
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Line lnbom 
         Index           =   0
         Visible         =   0   'False
         X1              =   328
         X2              =   200
         Y1              =   96
         Y2              =   224
      End
      Begin VB.Line pth 
         Index           =   0
         Visible         =   0   'False
         X1              =   208
         X2              =   336
         Y1              =   232
         Y2              =   104
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
      Begin VB.Image air 
         Height          =   1935
         Index           =   0
         Left            =   2880
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin Project1.aicAlphaImage bldng 
         Height          =   1455
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
         _extentx        =   2566
         _extenty        =   2566
         image           =   "frmmain.frx":0000
         scaler          =   3
         hittest         =   3
      End
      Begin VB.Image tree 
         Height          =   1935
         Index           =   0
         Left            =   3240
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin Project1.aicAlphaImage Missile 
         Height          =   1455
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
         _extentx        =   2566
         _extenty        =   2566
         image           =   "frmmain.frx":0018
         scaler          =   3
         props           =   0
      End
      Begin Project1.aicAlphaImage spot 
         Height          =   1335
         Left            =   480
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _extentx        =   1296
         _extenty        =   1508
         image           =   "frmmain.frx":0030
         scaler          =   3
      End
      Begin VB.Label lblbldng 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to place the building"
         Height          =   495
         Left            =   5640
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5520
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   0
      Width           =   2295
   End
   Begin VB.PictureBox BldBar 
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   0
      ScaleHeight     =   2250
      ScaleWidth      =   9015
      TabIndex        =   4
      Top             =   5280
      Width           =   9015
      Begin MCI.MMControl mmc 
         Height          =   330
         Left            =   4680
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin Project1.DataView DVW 
         Height          =   1425
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2514
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label lblmsg 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   855
         Left            =   5520
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
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
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   9000
      ScaleHeight     =   4905
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   3495
      Begin Project1.button button3 
         Height          =   450
         Left            =   720
         TabIndex        =   1
         Top             =   4200
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin Project1.button button2 
         Height          =   450
         Left            =   720
         TabIndex        =   2
         Top             =   3480
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin Project1.button button1 
         Height          =   450
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const ATN_1     As Double = 0.785398163397448
Private Const PI      As Double = 3.14159265358979
Private Const Level As Integer = 128 ' Height of aircraft during flight hours (In Pixels)

'This looks like a great burden on the memory but accordng to my
'calculations the average data stored in memory is less than 2 MB

'Timer Data
Dim label As String
Dim time As Integer
Dim t_loop As Boolean
Dim S_time As Integer
Dim Trig_Count As Integer
Dim trigger(100) As String

'Map Data
Dim Map_Name As String
Dim Map_Money As Long
Dim Map_Ground As c32bppDIB
Dim Map_Eves_Count As Integer
Dim Map_On(100) As String
Dim Map_Do(100) As String
Dim Map_Done(100) As Boolean
Dim Map_Techlevel As Integer
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
Dim K(1000) As Long
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

Dim mode As Integer
Dim Host As Integer
Dim pvw As String
Dim R As Integer
Dim mX As Integer, mY As Integer
Dim idx As Integer
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
If weap.typ(bldng_weapon(e_bldng)) = "laser" Then
Laser e_bldng, unt, weap.damage(bldng_weapon(e_bldng)), weap.Color(bldng_weapon(e_bldng)), 1
ElseIf weap.typ(bldng_weapon(e_bldng)) = "bomb" Then
DoEvents
fire e_bldng, unt, "bld", ""
End If
End If
End If
End If
R:
DoEvents
Next
Next
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
If team(e_unt) <> "Allies" And team(unt) = "Allies" And team(e_unt) <> "-1" And team(unt) <> "-1" Then
rng = GetFromIni(weapon(unt), "Range", App.path & "\Rules\weapons.ini")
If gizzly(e_unt).Left > gizzly(unt).Left - rng And gizzly(e_unt).Left < gizzly(unt).Left + rng Then
If gizzly(e_unt).Top > gizzly(unt).Top - rng And gizzly(e_unt).Top < gizzly(unt).Top + rng Then
If gizzly(unt).Tag = "" Then
gizzly(unt).Tag = e_unt
End If
If gizzly(unt).Tag = e_unt Then
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
If team(unt) <> "Allies" And team(e_unt) = "Allies" And team(e_unt) <> "-1" And team(unt) <> "-1" Then
rng = weap.range(weapon(e_unt))
If gizzly(unt).Left > gizzly(e_unt).Left - rng And gizzly(unt).Left < gizzly(e_unt).Left + rng Then
If gizzly(unt).Top > gizzly(e_unt).Top - rng And gizzly(unt).Top < gizzly(e_unt).Top + rng Then
If gizzly(unt).Tag = "" Then
gizzly(unt).Tag = e_unt
End If
If gizzly(unt).Tag = e_unt Then
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
End If
DoEvents
Y:
Next
Next
End Sub

Private Sub bldng_Click(Index As Integer, ByVal Button As Integer)
Dim n As Integer
If bldng_team(Index) <> "Allies" Then
For n = 0 To lstsel.ListCount - 1
gizzly(Val(lstsel.List(n))).Tag = "bldng" & Index
Next
Else
Host = Index
If LCase(GetFromIni(bldng_name(Index), "type", App.path & "\rules\buildings.ini")) = "constyard" Then
BldBar.Tag = "bldng"
ElseIf LCase(GetFromIni(bldng_name(Index), "type", App.path & "\rules\buildings.ini")) = "acc" Then
BldBar.Tag = "air"
ElseIf LCase(GetFromIni(bldng_name(Index), "type", App.path & "\rules\buildings.ini")) = "warfactory" Then
BldBar.Tag = "tank"
End If
Bldbar_Set Index
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
explode .X2, .Y2, 256, 1000
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

Sub explode(ByVal X As Integer, ByVal Y As Integer, ByVal radius As Integer, ByVal damage As Integer)
On Error Resume Next
Dim unt As Integer, wid As Integer
For unt = gizzly.LBound To gizzly.UBound
If unt = 0 Then GoTo Y
If gizzly(unt).Left > X - rad And gizzly(unt).Left < X + rad Then
If gizzly(unt).Top > Y - rad And gizzly(unt).Top < Y + rad Then
wid = Abs(X - gizzly(unt).Left) * Abs(Y - gizzly(unt).Top)
power(unt) = power(unt) - (damage / wid) * 100
'MsgBox CStr((damage / wid) * 100) & "\\" & wid & "\\" & power(unt)
End If
End If
Y:
Next
End Sub
Sub destruct(m_idx As Integer)
On Error Resume Next
Dim K As Integer
Dim arg(0) As String
For K = 1 To Map_Eves_Count
If LCase(Left(Map_On(K), 10)) = "destbldng(" Then
arg(0) = Right(Map_On(K), Len(Map_On(K)) - 10)
arg(0) = Left(arg(0), Len(arg(0)) - 1)
If Val(arg(0)) = m_idx Then
Trig Map_Do(K)
End If
End If
Next

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
End Sub

Sub desttank(m_idx As Integer)
Dim X As Integer, mdX As Integer, mdY As Integer
Dim arg(0) As String
For X = 1 To Map_Eves_Count
If LCase(Left(Map_On(X), 8)) = "destunt(" Then
arg(0) = Right(Map_On(X), Len(Map_On(X)) - 8)
arg(0) = Left(arg(0), Len(arg(0)) - 1)
If Val(arg(0)) = m_idx Then
Trig Map_Do(X)
End If
End If
Next

If team(m_idx) = "Allies" Then
lstsel.clear
End If
mdX = gizzly(m_idx).Left + gizzly(m_idx).Width / 2
mdY = gizzly(m_idx).Top + gizzly(m_idx).Height / 2
Unload gizzly(m_idx)
sblast mdX, mdY
 K(m_idx) = 0
bk(m_idx) = 0
img(m_idx) = ""
power(m_idx) = 0
t_power(m_idx) = 100
speed(m_idx) = 0
weapon(m_idx) = ""
team(m_idx) = "-1"
End Sub

Function getval(str As String) As String
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

Private Sub button1_click()
doall True
End Sub

Sub doall(bool As Boolean)
auto_bldng = bool
auto = bool
tmrpvw = bool
t2b = bool
tmrSlip = bool
Timer = bool
tmrmon = bool
End Sub

Private Sub button2_click()
Unload Me
frmmnu.Show

End Sub

Private Sub button3_click()
End
End Sub

Private Sub DVW_Click(str As String)
If BldBar.Tag = "bldng" Then
mode = 1
ElseIf BldBar.Tag = "tank" Then
If Map_Money - tanks.cost(str) < 0 Then
msg "Not Enough Credits"
Else
tnkfromini str, "Allies", bldng(Host).Left + struc.offx(bldng_name(Host)), bldng(Host).Top + struc.offy(bldng_name(Host)), bldng(Host).Left + struc.DocX(bldng_name(Host)), bldng(Host).Top + struc.docY(bldng_name(Host))
Map_Money = Map_Money - tanks.cost(str)
End If
ElseIf BldBar.Tag = "air" Then
mode = 2
End If
End Sub

Private Sub Form_Load()
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
LoadMap App.path & "\maps\Camp1\" '& GetFromIni("Main", "Progress", App.path & "\set.cfg") & "\"
End Sub

Sub makeTree(X As Integer, Y As Integer)
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
End Sub

Sub tnkfromini(ini As String, side As String, ByVal X As Integer, ByVal Y As Integer, ByVal toX As Integer, ByVal toY As Integer)
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

Sub bldngfromini(ini As String, side As String, flip As Boolean, ByVal X As Integer, ByVal Y As Integer)
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
If LCase(GetFromIni(ini, "type", App.path & "\rules\buildings.ini")) = "techlab" Then
Map_Techlevel = Map_Techlevel + 1
ElseIf LCase(GetFromIni(ini, "type", App.path & "\rules\buildings.ini")) = "powerplant" Then
Map_Many = Map_Many + 1
End If
MakeBldng img, side, flip, X, Y, offx, offy, CLng(pow), wpn, ini
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

Sub MakeBldng(image As String, side As String, flip As Boolean, X As Integer, Y As Integer, ByVal offx As Integer, ByVal offy As Integer, e_power As Long, e_weapon As String, ini As String)
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
bldng(bldng.UBound).LoadImage_FromFile App.path & "\Images\Buildings\" & image & ".png"
bldng(bldng.UBound).Opacity = 0
bldng(bldng.UBound).FadeInOut 100
bldng(bldng.UBound).Visible = True
End Sub

Sub TnkMove(Index As Integer, ByVal locX As Integer, ByVal locY As Integer)
Line1(Index).X2 = locX
Line1(Index).Y2 = locY
Line1(Index).X1 = gizzly(Index).Left + (gizzly(Index).Width / 2)
Line1(Index).Y1 = gizzly(Index).Top + (gizzly(Index).Height / 2)
K(Index) = 0
Rotate Angle(Line1(Index).X1, Line1(Index).Y1, Line1(Index).X2, Line1(Index).Y2), img(Index), Index
Timer1(Index) = True
Timer1_Timer Index
End Sub

Sub airMove(Index As Integer, ByVal locX As Integer, ByVal locY As Integer)
Line1(Index).X2 = locX
Line1(Index).Y2 = locY
Line1(Index).X1 = gizzly(Index).Left + (gizzly(Index).Width / 2)
Line1(Index).Y1 = gizzly(Index).Top + (gizzly(Index).Height / 2)
K(Index) = 0
Rotate Angle(Line1(Index).X1, Line1(Index).Y1, Line1(Index).X2, Line1(Index).Y2), img(Index), Index
Timer1(Index) = True
Timer1_Timer Index
End Sub

Private Sub Form_Resize()
Picture1.Left = 0
Picture1.Top = 0
BldBar.Top = (Me.Height / 15) - BldBar.Height
End Sub



Private Sub Gizzly_Click(Index As Integer)
On Error Resume Next
Dim nm As Integer, rad As Integer
If team(Index) = "Allies" Then
If lstsel.ListCount = 0 Then
lstsel.AddItem Index
End If
Else
Dim X As Integer, Y As Integer, wid As Long, tr As Boolean: tr = False
For nm = 0 To lstsel.ListCount - 1
rad = GetFromIni(weapon(lstsel.List(nm)), "Range", App.path & "\Rules\weapons.ini")
If gizzly(Index).Left - rad < gizzly(lstsel.List(nm)).Left And gizzly(Index).Left + rad < gizzly(lstsel.List(nm)).Left Then
If gizzly(Index).Top - rad < gizzly(lstsel.List(nm)).Top And gizzly(Index).Top + rad < gizzly(lstsel.List(nm)).Top Then
With lnmsc
.X1 = gizzly(lstsel.List(nm)).Left - (gizzly(lstsel.List(nm)).Width / 2): .Y1 = gizzly(lstsel.List(nm)).Top - (gizzly(lstsel.List(nm)).Height / 2)
.X2 = gizzly(Index).Left - (gizzly(Index).Height / 2): .Y2 = gizzly(Index).Top - (gizzly(Index).Height / 2)
wid = Hyp(Wline(.X1, .X2), Hline(.Y1, .Y2))

Y = -(Sin(PI * Angle(.X2, .Y2, .X1, .Y1) / 180) * rad) + gizzly(Index).Top
X = Cos(PI * Angle(.X2, .Y2, .X1, .Y1) / 180) * rad + gizzly(Index).Left
TnkMove lstsel.List(nm), X, Y
gizzly(lstsel.List(nm)).Tag = Index
End With
End If
End If
Next
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
If Index > 0 Then
pvw = "gizz" & Index
End If
End Sub

Private Sub mmc_Done(NotifyCode As Integer)
mmc.Command = "Stop"
mmc.FileName = App.path & "\Trax\Track" & CStr(Round(Rnd * 2) + 1) & ".mp3"
mmc.Command = "Open"
mmc.Command = "Play"
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Picture2.Left = (Me.Width / 15) / 2 - Picture2.Width / 2
Picture2.Top = (Me.Height / 15) / 2 - Picture2.Height / 2
Picture2.Visible = True
doall False
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
lstsel.clear
Else
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
ElseIf Button = 1 And mode = 1 Then
If Map_Money - struc.cost(DVW.sel) < 0 Then
msg "Not Enough Credits"
mode = 0
lblbldng.Visible = True
Else
bldngfromini DVW.sel, "Allies", False, X, Y
Map_Money = Map_Money - struc.cost(DVW.sel)
mode = 0
lblbldng.Visible = False
End If
ElseIf Button = 2 And mode = 1 Then
mode = 0
lblbldng.Visible = False
End If
End If
End Sub

Function Rndmz(Seed As Integer) As Long
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
ElseIf Button = 0 Then
If mode = 1 Then
lblbldng.Visible = True
lblbldng.Left = X + 20
lblbldng.Top = Y + 20
End If
End If
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
If weap.typ(weapon(unt)) = "bomb" Then
fire unt, e_unt, "", "bld"
ElseIf weap.typ(weapon(unt)) = "laser" Then
Laser unt, e_unt, weap.damage(weapon(unt)), weap.Color(weapon(unt)), 2
End If
End If
End If
End If
End If
X:
Next: Next
End Sub

Private Sub tanks_Click()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Trig Text1
End If
End Sub

Private Sub Timer_Timer()
Dim K As Integer
time = time - 1
If time <= 0 Then
For K = 1 To Trig_Count
Trig (trigger(K))
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
If K(Index) <= wid Then
Y = -(Sin(PI * ang / 180) * K(Index)) + Line1(Index).Y1
X = Cos(PI * ang / 180) * K(Index) + Line1(Index).X1
For Tk = 1 To bldng.UBound
If bldng_team(Tk) <> "-1" Then
If LCase(GetFromIni(bldng_name(Tk), "type", App.path & "\rules\buildings.ini")) = "wall" Then mD = SS = 0 Else mD = (bldng(Tk).Height - (bldng(Tk).Height / 3.5)): SS = speed(Index) + 2
If X > bldng(Tk).Left + SS And X < bldng(Tk).Left + bldng(Tk).Width - SS Then
If Y > bldng(Tk).Top + mD + SS And Y < bldng(Tk).Top + bldng(Tk).Height - SS Then
If LCase(GetFromIni(bldng_name(Tk), "type", App.path & "\rules\buildings.ini")) <> "warfactory" Then
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
K(Index) = K(Index) + speed(Index)
Timer1(Index).interval = 1
gizzly(Index).ToolTipText = " "
DoEvents
End If
DoEvents
Else
H:
gizzly(Index).ToolTipText = ""
Timer1(Index) = False
K(Index) = 0
DoEvents
End If
End Sub

Sub fire(from As Integer, too As Integer, so As String, dest As String, Optional tgtX As Integer = 1, Optional tgtY As Integer = 1)
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


Sub CheckSide()
On Error Resume Next
Dim K As Integer, L As Integer, xst As Boolean
Dim arg(0) As String
For K = 1 To Map_Eves_Count
If LCase(Left(Map_On(K), 9)) = "destside(" Then
arg(0) = Right(Map_On(K), Len(Map_On(K)) - 9)
arg(0) = Left(arg(0), Len(arg(0)) - 1)
For L = 1 To bldng.UBound
If bldng_team(L) <> "-1" And LCase(bldng_team(L)) = LCase(arg(0)) Then
xst = True
Exit For
End If
Next
If xst = False Then
If Map_Done(K) = False Then
Trig Map_Do(K)
Map_Done(K) = True
End If
End If
End If
Next
End Sub

Function Rndeg(ByVal Deg As Integer) As Integer
Rndeg = Round(Deg / 18)
End Function
Sub Laser(from As Integer, too As Integer, damage As Integer, Color As Long, frombldng1tobldng2else3 As Integer, Optional aircraft As Boolean = False, Optional airX As Integer = 0, Optional airY As Integer = 0)
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
If power(too) <= 0 Then desttank (too)
ElseIf frombldng1tobldng2else3 = 1 Then
.X1 = bldng(from).Left + (bldng(from).Width / 2) + bldng_offsetX(from)
.Y1 = bldng(from).Top + (bldng(from).Height / 2) + bldng_OffsetY(from)
.X2 = gizzly(too).Left + gizzly(too).Width / 2
.Y2 = gizzly(too).Top + gizzly(too).Height / 2
power(too) = power(too) - damage
If power(too) <= 0 Then desttank (too)
Else
.X1 = gizzly(from).Left + (gizzly(from).Width / 2)
.Y1 = gizzly(from).Top + (gizzly(from).Height / 2)
.X2 = bldng(too).Left + bldng(too).Width / 2
.Y2 = bldng(too).Top + bldng(too).Height / 2
Rotate Angle(.X1, .Y1, .X2, .Y2), img(from), from
bldng_pow(too) = bldng_pow(too) - damage
If bldng_pow(too) <= 0 Then: destruct (too)
End If
GoTo v
U:
.X1 = air(from).Left + air(from).Width / 2
.Y1 = air(from).Top + air(from).Height / 2
.X2 = airX
.Y2 = airY
explode .X2, .Y2, 200, damage
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
End With
End Sub

Sub sblast(ByVal X As Integer, ByVal Y As Integer)
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
Laser Index, 0, 1000, vbGreen, 0, True, pth(Index).X2, pth(Index).Y2
Else
fire Index, 1, "air", "", pth(Index).X2, pth(Index).Y2 + Level
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
Bar.Top = gizzly(str).Top + 12
ElseIf Left(pvw, 5) = "bldng" Then
str = Val(Right(pvw, Len(pvw) - 5))
Bar.SetPro CStr(bldng_pow(str)), CStr(bldng_tpow(str))
Bar.Left = (bldng(str).Left + bldng(str).Width / 2) - (Bar.Width / 2)
Bar.Top = bldng(str).Top + 12
End If
DoEvents
CheckSide
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

Private Sub tmrSlip_Timer()
GetCursorPos pt
If pt.X < 3 Then
If Picture1.Left > 0 Then
Picture1.Left = 0
Else
Picture1.Left = Picture1.Left + 3
End If
ElseIf pt.Y < 3 Then
If Picture1.Top > 0 Then
Picture1.Top = 0
Else
Picture1.Top = Picture1.Top + 3
End If
ElseIf pt.X > (Screen.Width / 15) - 3 Then
If Picture1.Left + Picture1.Width < Me.Width / 15 Then
Picture1.Left = Me.Width / 15 - Picture1.Width
Else
Picture1.Left = Picture1.Left - 3
End If
ElseIf pt.Y > (Screen.Height / 15) - 3 Then
If Picture1.Top + Picture1.Height < Me.Height / 15 Then
Picture1.Top = Me.Height / 15 - Picture1.Height
Else
Picture1.Top = Picture1.Top - 3
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
Dim flip As Boolean
ini = str & "\ini.ini"
Me.Hide
RemoveAll
Picture1.Left = 0
Picture1.Top = 0
Picture1.Width = Val(GetFromIni("Main", "Width", ini))
Picture1.Height = Val(GetFromIni("Main", "Height", ini))
Map_Name = GetFromIni("Main", "Name", ini)
Map_Money = GetFromIni("Main", "Money", ini)
tex = GetFromIni("Main", "ground", ini)

DoEvents
frmdum.Hide
frmdum.AutoRedraw = True
Picture1.AutoRedraw = True
frmdum.Width = Picture1.Width * 15
frmdum.Height = Picture1.Height * 15
Map_Ground.InitializeDIB Picture1.Width, Picture1.Height
Map_Ground.LoadPicture_File App.path & "\Images\Texture\" & tex
For cX = 0 To frmdum.Width Step 256
For cY = 0 To frmdum.Height Step 256
Map_Ground.Render frmdum.hdc, cX, cY, 256, 256
DoEvents
Next: Next
frmdum.Picture = frmdum.image
Map_Ground.DestroyDIB
Set Map_Ground = Nothing
frmdum.Hide

Trig_Count = Val(GetFromIni("Timer", "count", ini))
For X = 1 To Trig_Count
trigger(X) = GetFromIni("Timer", "Trigger" & CStr(X), ini)
Next
label = GetFromIni("Timer", "label", ini)
time = Val(GetFromIni("Timer", "time", ini))
S_time = Val(GetFromIni("Timer", "time", ini))
t_loop = str2bol(GetFromIni("Timer", "loop", ini))
Timer = False
Timer = True

For X = 1 To GetFromIni("Masks", "count", ini)
frmdum.msk.Picture = LoadPicture(str & "\Mask" & CStr(X) & ".bmp")
frmdum.pic.Picture = LoadPicture(App.path & "\images\texture\" & GetFromIni("Masks", "Type:" & CStr(X), ini))
Setup frmdum.pic.hdc, frmdum.msk.hdc, frmdum, frmdum.pic
BLTIT Val(GetFromIni("Masks", "X:" & CStr(X), ini)), Val(GetFromIni("Masks", "Y:" & CStr(X), ini)), frmdum
frmdum.Picture = frmdum.image
Cleanup
Picture1.Picture = frmdum.image
DoEvents
Next

Picture1.AutoRedraw = False
Unload frmdum
Map_Eves_Count = Val(GetFromIni("Events", "Count", ini))
For X = 1 To Val(GetFromIni("Events", "Count", ini))
Map_On(X) = GetFromIni("Events", "On" & CStr(X), ini)
Map_Do(X) = GetFromIni("Events", "Do" & CStr(X), ini)
Next

For X = 0 To Val(GetFromIni("Tanks", "Count", ini))
tnkfromini GetFromIni("Tanks", "ini" & CStr(X), ini), GetFromIni("Tanks", "side" & CStr(X), ini), Val(GetFromIni("Tanks", "X" & CStr(X), ini)), Val(GetFromIni("Tanks", "Y" & CStr(X), ini)), Val(GetFromIni("Tanks", "toX" & CStr(X), ini)), Val(GetFromIni("Tanks", "toY" & CStr(X), ini))
Next

For X = 0 To GetFromIni("Buildings", "Count", ini)
If GetFromIni("Buildings", "Flip" & CStr(X), ini) = "1" Then flip = True Else flip = False
bldngfromini GetFromIni("Buildings", "ini" & CStr(X), ini), GetFromIni("Buildings", "side" & CStr(X), ini), flip, GetFromIni("Buildings", "X" & CStr(X), ini), GetFromIni("Buildings", "Y" & CStr(X), ini)
Next

For X = 1 To Val(GetFromIni("Trees", "Count", ini))
makeTree Val(GetFromIni("Trees", "TreeX" & CStr(X), ini)), Val(GetFromIni("Trees", "TreeY" & CStr(X), ini))
Next
Me.Show
mmc.Command = "Stop"
mmc.FileName = App.path & "\Trax\Track" & CStr(Round(Rnd * 2) + 1) & ".mp3"
mmc.Command = "Open"
mmc.Command = "Play"
Exit Sub
Y:
MsgBox Err.Description
End Sub

Sub Trig(str As String) ' A mini command processor for events and triggers
' It splits the command to its name, arguments and brackets , Use commands like
'destunt(index) ; destbldng(index) ; makebldng(ini as string,side as string,flip as boolean,x as integer,y as integer) ;
'maketank(ini as string,side as string,x as integer,y as integer,toX as integer,toY as integer) ;
'airmission(dock as integer,side as sting,ini as string,x as integer,y as integer) ' Initializes an air mission on ...
'nmsl(x as integer,y as integer) ' Fires neuclear missile on X and Y
'Use strings without quotations

Dim arg(5) As String
If LCase(str) = "loose" Then
frmmnu.Show
Unload Me
ElseIf LCase(str) = "win" Then
frmmnu.Tag = "next"
frmmnu.Show
Unload Me
ElseIf Left(LCase(str), 8) = "destunt(" Then
str = Right(str, Len(str) - 8)
str = Left(str, Len(str) - 1)
desttank Val(str)
ElseIf Left(LCase(str), 10) = "destbldng(" Then
str = Right(str, Len(str) - 10)
str = Left(str, Len(str) - 1)
destruct Val(str)
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
AirMission Val(arg(0)), arg(1), arg(2), Val(arg(3)), Val(arg(4))
ElseIf Left(LCase(str), 5) = "nmsl(" Then
str = Right(str, Len(str) - 5)
str = Left(str, Len(str) - 1)
arg(0) = Left(str, InStr(1, str, ",") - 1)
arg(1) = Right(str, Len(str) - Len(arg(0)) - 1)
Nmsl Val(arg(0)), Val(arg(1))
End If
End Sub

Sub Nmsl(X As Integer, Y As Integer)
Dim K As Integer
Missile.AutoSize = True
Missile.LoadImage_FromFile App.path & "\animations\Nmsl Down.png"
Missile.ZOrder 0
Missile.Visible = True
For K = -Missile.Height To Y - Missile.Height + 68 Step 5
Missile.Top = K
Missile.Left = X - (Missile.Width / 2)
DoEvents
Next
Missile.Visible = False
bblast X, Y
DoEvents
explode X, Y, 5120, 20000
End Sub

Function str2bol(str As String) As Boolean
If LCase(str) = "true" Then
str2bol = True
ElseIf LCase(str) = "false" Then
str2bol = False
End If
End Function

Sub Bldbar_Set(Index As Integer)
Dim K As Integer
DVW.clear
If LCase(GetFromIni(bldng_name(Index), "type", App.path & "\rules\buildings.ini")) = "acc" Then
For K = 1 To GetFromIni("Main", "count", App.path & "\rules\aircrafts.ini")
If aeros.techlevel(GetFromIni("Main", "a" & CStr(K), App.path & "\rules\aircrafts.ini")) <> "-1" And aeros.techlevel(GetFromIni("Main", "a" & K, App.path & "\rules\aircrafts.ini")) <= Map_Techlevel Then
DVW.Add GetFromIni("Main", "a" & CStr(K), App.path & "\rules\aircrafts.ini"), aircraft
End If
Next
ElseIf LCase(GetFromIni(bldng_name(Index), "type", App.path & "\rules\buildings.ini")) = "constyard" Then
For K = 1 To GetFromIni("Main", "count", App.path & "\rules\buildings.ini")
If struc.techlevel(GetFromIni("Main", "b" & CStr(K), App.path & "\rules\buildings.ini")) <> "-1" And struc.techlevel(GetFromIni("Main", "b" & K, App.path & "\rules\buildings.ini")) <= Map_Techlevel Then
DVW.Add GetFromIni("Main", "b" & CStr(K), App.path & "\rules\buildings.ini"), building
End If
Next
ElseIf LCase(GetFromIni(bldng_name(Index), "type", App.path & "\rules\buildings.ini")) = "warfactory" Then
For K = 1 To GetFromIni("Main", "count", App.path & "\rules\tanks.ini")
If tanks.techlevel(GetFromIni("Main", "t" & CStr(K), App.path & "\rules\tanks.ini")) <> "-1" And tanks.techlevel(GetFromIni("Main", "t" & K, App.path & "\rules\tanks.ini")) <= Map_Techlevel And tanks.water(GetFromIni("Main", "t" & K, App.path & "\rules\tanks.ini")) = "0" Then
DVW.Add GetFromIni("Main", "t" & CStr(K), App.path & "\rules\tanks.ini"), tank
End If
Next
Else
End If
End Sub

Sub msg(str As String)
lblmsg = str
lblmsg.Top = lblmsg.Top + 500
lblmsg.Left = lblmsg.Left - 500
lblmsg.Left = lblmsg.Left + 500
lblmsg.Top = lblmsg.Top - 500
End Sub


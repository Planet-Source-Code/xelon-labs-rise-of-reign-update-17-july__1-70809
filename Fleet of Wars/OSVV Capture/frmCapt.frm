VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Spy"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   4650
   ClientWidth     =   2100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   2100
   Begin VB.CommandButton Command3 
      Caption         =   "Get Window"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get 3D View Box"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Artillery"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Process"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton picture1 
      Caption         =   "Grab Textbox"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   1920
      Left            =   120
      ScaleHeight     =   1860
      ScaleWidth      =   1860
      TabIndex        =   4
      Top             =   2280
      Width           =   1920
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1100
         Left            =   2280
         Top             =   360
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Left            =   3120
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public wnd As Long
Public wnd3D As Long
Dim wnde As String

Private Const WM_GETTEXT = &HD
Private Const WM_SETTEXT = &HC
Dim deg As Integer
Dim sU As String * 256
Dim num As Integer

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Screen.MousePointer = 2
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pt As POINTAPI
    GetCursorPos pt
Screen.MousePointer = 0
    wnd3D = WindowFromPoint(pt.x, pt.y)
End Sub
Private Sub Command2_Click()
On Error GoTo U
Dim n As Integer
Me.Left = Screen.Width - Me.Width
    SetWindowPos wnde, 0, 0, 0, 273, 194, 0
For n = -1 To 378 Step 18
    SendMessageSTRING wnd, WM_SETTEXT, 256, n
    Set Picture2.Picture = CaptureWindow(wnd3D, False, 0, 0, 126, 126)
SavePicture Picture2.Picture, App.Path & "\Images\" & Text1 & Fix(n / 18) & ".bmp"
DoEvents
Next
Kill App.Path & "\Images\" & Text1 & Fix(-1 / 19) & ".bmp"
FileCopy App.Path & "\Images\" & Text1 & Fix(n / 19) & ".bmp", App.Path & "\Images\" & Text1 & "0" & ".bmp"
Exit Sub
U:
MsgBox Err.Description, vbCritical
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Screen.MousePointer = 2
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim X1 As Integer, Y1 As Integer, rct As RECT
    Dim pt As POINTAPI
    GetCursorPos pt
Screen.MousePointer = 0
    wnde = WindowFromPoint(pt.x, pt.y)
    GetWindowRect wnde, rct
    X1 = rct.Right - rct.Left
    Y1 = rct.Bottom - rct.Top
End Sub

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Screen.MousePointer = 2
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pt As POINTAPI
    GetCursorPos pt
Screen.MousePointer = 0
    wnd = WindowFromPoint(pt.x, pt.y)
End Sub

Private Sub Timer1_Timer()
If deg <= 360 Then
    Text2 = deg
    SetWindowPos wnde, 0, 0, 0, 273, 194, 0
    Set Picture2.Picture = CaptureWindow(wnd3D, False, 0, 0, 126, 126)
    SendMessageSTRING wnd, WM_SETTEXT, 256, Text2
SavePicture Picture2.Picture, App.Path & "\Images\" & Text1 & num & ".bmp"
deg = deg + 18
num = num + 1
Else
Timer1 = False
    SendMessageSTRING wnd, WM_SETTEXT, 256, 360
        Set Picture2.Picture = CaptureWindow(wnd3D, False, 0, 0, 128, 128)
SavePicture Picture2.Picture, App.Path & "\Images\" & Text1 & "12" & ".bmp"
End If
End Sub

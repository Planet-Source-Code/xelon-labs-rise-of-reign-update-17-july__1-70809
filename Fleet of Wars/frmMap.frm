VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Mapedit"
   ClientHeight    =   10275
   ClientLeft      =   1440
   ClientTop       =   4740
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   14595
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   9
      Left            =   0
      Top             =   4680
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   4680
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   441
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Make Texture"
      TabPicture(0)   =   "frmMap.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lv1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "filTex"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Default Ground"
      TabPicture(1)   =   "frmMap.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FileTex"
      Tab(1).Control(1)=   "Label5"
      Tab(1).ControlCount=   2
      Begin VB.FileListBox filTex 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   1800
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Set Texture"
         Height          =   255
         Left            =   1800
         TabIndex        =   41
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         Height          =   255
         Left            =   5880
         TabIndex        =   37
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "NewTexture"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1575
      End
      Begin VB.FileListBox FileTex 
         Height          =   1650
         Left            =   -74880
         Pattern         =   "*.gif*"
         TabIndex        =   38
         Top             =   600
         Width           =   6615
      End
      Begin ComctlLib.ListView lv1 
         Height          =   1695
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Index"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "X1"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "X2"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Y1"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Y2"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   2152
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Default Texture of the ground"
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   360
         Width           =   2295
      End
   End
   Begin TabDlg.SSTab Eves 
      Height          =   2415
      Left            =   4680
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   441
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Events"
      TabPicture(0)   =   "frmMap.frx":0038
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lsteves"
      Tab(0).Control(1)=   "Command10"
      Tab(0).Control(2)=   "Command9"
      Tab(0).Control(3)=   "Command8"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Panic Timer"
      TabPicture(1)   =   "frmMap.frx":0054
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lsttime"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command7"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Check1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CheckBox Check1 
         Caption         =   "Loop"
         Height          =   255
         Left            =   5040
         TabIndex        =   43
         Top             =   360
         Width           =   1695
      End
      Begin ComctlLib.ListView lsteves 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   27
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Event"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Trigger"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Change Selected Item"
         Height          =   255
         Left            =   -73080
         TabIndex        =   30
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Remove Selected Item"
         Height          =   255
         Left            =   -70200
         TabIndex        =   29
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add New Item"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Edit :->"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   1800
         Width           =   1695
      End
      Begin ComctlLib.ListView lsttime 
         Height          =   1575
         Left            =   2520
         TabIndex        =   25
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2778
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Triggers"
            Object.Width           =   6704
         EndProperty
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove Trigger :->"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add Trigger :->"
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   22
         Text            =   "1000"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Text            =   "Seconds Remaining : "
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "Time"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Timer Label"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   2985
      TabIndex        =   15
      Top             =   0
      Width           =   3015
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "X, Y"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.PictureBox picProp 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   14355
      TabIndex        =   1
      Top             =   8160
      Width           =   14415
      Begin VB.CommandButton Command3 
         Caption         =   "Static Units"
         Height          =   255
         Left            =   11520
         TabIndex        =   71
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Side 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   8760
         TabIndex        =   14
         Text            =   "Side"
         Top             =   240
         Width           =   1695
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   1815
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3201
         _Version        =   393216
         Tab             =   1
         TabHeight       =   441
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Size"
         TabPicture(0)   =   "frmMap.frx":0070
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Combo1"
         Tab(0).Control(1)=   "Text1"
         Tab(0).Control(2)=   "Text2"
         Tab(0).Control(3)=   "Label1"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Properties"
         TabPicture(1)   =   "frmMap.frx":008C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label8"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label19"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label20"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Text3"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Text6"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Text17"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Text18"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).ControlCount=   8
         TabCaption(2)   =   "Description"
         TabPicture(2)   =   "frmMap.frx":00A8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Desc"
         Tab(2).ControlCount=   1
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   1440
            TabIndex        =   72
            Text            =   "50"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   4200
            TabIndex        =   70
            Text            =   "0"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1440
            TabIndex        =   32
            Text            =   "10000"
            Top             =   960
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmMap.frx":00C4
            Left            =   -73440
            List            =   "frmMap.frx":00D4
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1000
            Width           =   1455
         End
         Begin RichTextLib.RichTextBox Desc 
            Height          =   1335
            Left            =   -74880
            TabIndex        =   8
            Top             =   360
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            _Version        =   393217
            Enabled         =   -1  'True
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmMap.frx":0101
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   480
            TabIndex        =   7
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   -74280
            TabIndex        =   4
            Text            =   "Width"
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   -74280
            TabIndex        =   3
            Text            =   "Height"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "AI Skills"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label19 
            Caption         =   "Light Offset"
            Height          =   255
            Left            =   3240
            TabIndex        =   69
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Money"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Name of Map"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Map Size :->"
            Height          =   255
            Left            =   -74760
            TabIndex        =   5
            Top             =   480
            Width           =   2295
         End
      End
      Begin TabDlg.SSTab MakeTab 
         Height          =   1815
         Left            =   4680
         TabIndex        =   10
         Top             =   0
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3201
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   441
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Make Tanks"
         TabPicture(0)   =   "frmMap.frx":018B
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "TnkLst"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Make Buildings"
         TabPicture(1)   =   "frmMap.frx":01A7
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "BldngLst"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Make Extra Terrain"
         TabPicture(2)   =   "frmMap.frx":01C3
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TrnLst"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Properties"
         TabPicture(3)   =   "frmMap.frx":01DF
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture2"
         Tab(3).Control(1)=   "Picture3"
         Tab(3).ControlCount=   2
         Begin VB.PictureBox Picture3 
            Height          =   1335
            Left            =   -71400
            ScaleHeight     =   1275
            ScaleWidth      =   3075
            TabIndex        =   45
            Top             =   360
            Visible         =   0   'False
            Width           =   3135
            Begin VB.TextBox Text16 
               Height          =   285
               Left            =   480
               Locked          =   -1  'True
               TabIndex        =   62
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox Text15 
               Height          =   285
               Left            =   480
               Locked          =   -1  'True
               TabIndex        =   61
               Top             =   1005
               Width           =   975
            End
            Begin VB.TextBox Text14 
               Height          =   285
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   60
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox Text13 
               Height          =   285
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   59
               Top             =   1005
               Width           =   975
            End
            Begin VB.TextBox Text12 
               Height          =   285
               Left            =   840
               TabIndex        =   56
               Top             =   400
               Width           =   1575
            End
            Begin VB.TextBox Text11 
               Height          =   285
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   55
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label18 
               Caption         =   "ToX"
               Height          =   255
               Left            =   1560
               TabIndex        =   66
               Top             =   720
               Width           =   375
            End
            Begin VB.Label Label17 
               Caption         =   "ToY"
               Height          =   255
               Left            =   1560
               TabIndex        =   65
               Top             =   1005
               Width           =   375
            End
            Begin VB.Label Label16 
               Caption         =   "X"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   720
               Width           =   375
            End
            Begin VB.Label Label15 
               Caption         =   "Y"
               Height          =   255
               Left            =   120
               TabIndex        =   63
               Top             =   1005
               Width           =   375
            End
            Begin VB.Label Label14 
               Caption         =   "Side"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   400
               Width           =   615
            End
            Begin VB.Label Label13 
               Caption         =   "INI"
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   120
               Width           =   735
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   1335
            Left            =   -74880
            ScaleHeight     =   1275
            ScaleWidth      =   3075
            TabIndex        =   44
            Top             =   360
            Visible         =   0   'False
            Width           =   3135
            Begin VB.CheckBox Check2 
               Caption         =   "Flip"
               Height          =   255
               Left            =   360
               TabIndex        =   54
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox Text10 
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   1000
               Width           =   975
            End
            Begin VB.TextBox Text9 
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox Text8 
               Height          =   285
               Left            =   720
               TabIndex        =   48
               Top             =   400
               Width           =   1695
            End
            Begin VB.TextBox Text7 
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   120
               Width           =   1695
            End
            Begin VB.Label Label12 
               Caption         =   "Y"
               Height          =   255
               Left            =   1440
               TabIndex        =   53
               Top             =   1005
               Width           =   375
            End
            Begin VB.Label Label11 
               Caption         =   "X"
               Height          =   255
               Left            =   1440
               TabIndex        =   52
               Top             =   720
               Width           =   375
            End
            Begin VB.Label Label10 
               Caption         =   "Side"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   400
               Width           =   615
            End
            Begin VB.Label Label9 
               Caption         =   "INI"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   120
               Width           =   615
            End
         End
         Begin ComctlLib.ListView TnkLst 
            Height          =   1335
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   2355
            View            =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin ComctlLib.ListView BldngLst 
            Height          =   1335
            Left            =   -74880
            TabIndex        =   12
            Top             =   360
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   2355
            View            =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin ComctlLib.ListView TrnLst 
            Height          =   1335
            Left            =   -74880
            TabIndex        =   13
            Top             =   360
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   2355
            View            =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
   Begin VB.PictureBox PicView 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton ln2 
         Caption         =   "X"
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   240
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ListBox lstdum 
         Height          =   1230
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label H2 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Left            =   8640
         TabIndex        =   68
         Top             =   3000
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label H1 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Left            =   6000
         TabIndex        =   67
         Top             =   960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Line lnTex 
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   6000
         X2              =   8760
         Y1              =   1080
         Y2              =   3120
      End
      Begin VB.Image tree 
         Height          =   1335
         Index           =   0
         Left            =   1200
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Image terr 
         Height          =   1335
         Index           =   0
         Left            =   1080
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Line lnwp 
         Visible         =   0   'False
         X1              =   3840
         X2              =   2880
         Y1              =   1080
         Y2              =   2640
      End
      Begin Project1.aicAlphaImage pvw 
         Height          =   855
         Left            =   4200
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         Image           =   "frmMap.frx":01FB
         Scaler          =   3
      End
      Begin Project1.aicAlphaImage bldng 
         Height          =   1335
         Index           =   0
         Left            =   960
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   2355
         Image           =   "frmMap.frx":0213
         Scaler          =   3
         HitTest         =   3
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click to place the selected item"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image gizzly 
         Height          =   1335
         Index           =   0
         Left            =   840
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Menu mnu1 
      Caption         =   "File"
      Begin VB.Menu mnu_open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_SHE 
         Caption         =   "Show/Hide Events"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu_SHTE 
         Caption         =   "Show/Hide Texture Editor"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnu_dash1 
         Caption         =   "-"
      End
      Begin VB.Menu nu_exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim selected As Boolean
Dim PT As POINTAPI

Dim tnk_side(1000) As String
Dim tnk_ini(1000) As String
Dim tnk_X(1000) As Long
Dim tnk_Y(1000) As Long
Dim tnk_toX(1000) As Long
Dim tnk_toY(1000) As Long

Dim bldng_ini(1000) As String
Dim bldng_side(1000) As String
Dim bldng_X(1000) As Long
Dim bldng_Y(1000) As Long
Dim bldng_flip(1000) As Integer

Dim mX As Integer, mY As Integer

Private Sub bldng_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
bldng_side(Index) = "-1"
bldng_ini(Index) = ""
bldng_X(Index) = 0
bldng_Y(Index) = 0
bldng_flip(Index) = 0
Unload bldng(Index)
ElseIf Button = 1 Then
mX = X
mY = Y
Picture2.Tag = CStr(Index)
Text7 = bldng_ini(Index)
Text8 = bldng_side(Index)
Text9 = bldng_X(Index)
Text10 = bldng_Y(Index)
Check2.Value = bldng_flip(Index)
Picture3.Visible = False
Picture2.Visible = True
End If
End Sub

Private Sub bldng_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4 = CStr((bldng(Index).Left / 15) + X / 15) & " '.' " & CStr((bldng(Index).Top / 15) + Y / 15) & " '.' " & CStr(Index + 1)
If Button = 1 Then
bldng(Index).Move bldng(Index).Left + X - mX, bldng(Index).Top + Y - mY
bldng_X(Index) = bldng(Index).Left
bldng_Y(Index) = bldng(Index).Top
End If
End Sub

Private Sub BldngLst_ItemClick(ByVal Item As ComctlLib.ListItem)
Label3.Visible = True
Label3.ToolTipText = Item.Text
Label3.Tag = "bldng"
Dim img As String
pvw.Visible = True
pvw.AutoSize = True
img = (GetFromIni(Item.Text, "image", App.path & "\rules\buildings.ini"))
pvw.LoadImage_FromFile App.path & "\images\buildings\" & img & ".png"
selected = True
End Sub

Private Sub Check2_Click()
On Error Resume Next
bldng_flip(Index) = Val(Check2.Value)
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
Text1 = "4500"
Text2 = "10000"
ElseIf Combo1.ListIndex = 1 Then
Text1 = "10000"
Text2 = "5200"
ElseIf Combo1.ListIndex = 2 Then
Text1 = "13000"
Text2 = "13000"
ElseIf Combo1.ListIndex = 3 Then
Text1 = "6000"
Text2 = "6000"
End If
DoEvents
End Sub



Private Sub Command1_Click()
Dim lst As ListItem
Set lst = lv1.ListItems.Add(, , lv1.ListItems.Count)
lst.SubItems(1) = "256"
lst.SubItems(2) = "512"
lst.SubItems(3) = "256"
lst.SubItems(4) = "512"
lst.SubItems(5) = "Grass.gif"
End Sub

Private Sub Command10_Click()
Dim str As String
Dim str2 As String
str = InputBox("Enter Event", "??", lsteves.SelectedItem.Text)
str2 = InputBox("Enter Trigger Command", "???", lsteves.SelectedItem.SubItems(1))
lsteves.SelectedItem.Text = str
lsteves.SelectedItem.SubItems(1) = str2
End Sub

Private Sub Command11_Click()
filTex.path = App.path & "\images\texture\"
filTex.Visible = Not filTex.Visible
End Sub

Private Sub Command2_Click()
lv1.ListItems.Remove lv1.SelectedItem.Index
End Sub

Private Sub Command3_Click()
Dim X As Integer
For X = 1 To gizzly.UBound
tnk_toX(X) = tnk_X(X) + (gizzly(X).Width / 15) / 2
tnk_toY(X) = tnk_Y(X) + (gizzly(X).Height / 15) / 2
Next
End Sub

Private Sub Command5_Click()
Dim str As String
str = InputBox("Enter Trigger Commands", "???", "Loose")
lsttime.ListItems.Add , , str
End Sub

Private Sub Command6_Click()
lsttime.ListItems.Remove lsttime.SelectedItem.Index
End Sub

Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lsttime.StartLabelEdit
End Sub

Private Sub Command8_Click()
Dim str As String
Dim str2 As String
str = InputBox("Enter Event", "??", "Destroyed(1)")
str2 = InputBox("Enter Trigger Command", "???", "Loose")
Dim lst As ListItem
Set lst = lsteves.ListItems.Add(, , str)
lst.SubItems(1) = str2
End Sub

Private Sub Command9_Click()
lsteves.ListItems.Remove lsteves.SelectedItem.Index
End Sub

Private Sub filTex_Click()
On Error Resume Next
lv1.SelectedItem.SubItems(5) = filTex.List(filTex.ListIndex)
filTex.Visible = False
End Sub

Private Sub Form_Load()
FillLst
Set map_BG = New c32bppDIB
End Sub

Private Sub Form_Resize()
On Error Resume Next
picProp.Top = (Screen.Height - 750) - picProp.Height
picProp.Width = Width
Eves.Top = picProp.Top - Eves.Height
SSTab1.Top = picProp.Top - Eves.Height
End Sub

Private Sub gizzly_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
tnk_side(Index) = "-1"
tnk_ini(Index) = ""
tnk_X(Index) = 0
tnk_Y(Index) = 0
tnk_toX(Index) = 0
tnk_toY(Index) = 0
Unload gizzly(Index)
ElseIf Button = 1 Then
lnwp.Visible = True
ln2.Visible = True
lnwp.X1 = gizzly(Index).Left + gizzly(Index).Width / 2
lnwp.Y1 = gizzly(Index).Top + gizzly(Index).Height / 2
lnwp.X2 = tnk_toX(Index) * 15
lnwp.Y2 = tnk_toY(Index) * 15
ln2.Left = lnwp.X2 - ln2.Width / 2
ln2.Top = lnwp.Y2 - ln2.Width / 2
ln2.Tag = Index
mX = X
mY = Y

Picture3.Tag = CStr(Index)
Text11 = tnk_ini(Index)
Text12 = tnk_side(Index)
Text16 = tnk_X(Index)
Text15 = tnk_Y(Index)
Text13 = tnk_toX(Index)
Text14 = tnk_toY(Index)
Picture2.Visible = False
Picture3.Visible = True
End If
End Sub

Private Sub gizzly_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4 = CStr((gizzly(Index).Left / 15) + X / 15) & " '.' " & CStr((gizzly(Index).Top / 15) + Y / 15) & " '.' " & CStr(Index + 1)
If Button = 1 Then
gizzly(Index).Move gizzly(Index).Left + X - mX, gizzly(Index).Top + Y - mY
tnk_X(Index) = gizzly(Index).Left / 15
tnk_Y(Index) = gizzly(Index).Top / 15
lnwp.X1 = gizzly(Index).Left + gizzly(Index).Width / 2
lnwp.Y1 = gizzly(Index).Top + gizzly(Index).Height / 2
lnwp.X2 = tnk_toX(Index) * 15 + gizzly(Index).Width / 2
lnwp.Y2 = tnk_toY(Index) * 15 + gizzly(Index).Height / 2
ln2.Left = lnwp.X2 - ln2.Width / 2
ln2.Top = lnwp.Y2 - ln2.Width / 2
ln2.Tag = Index
End If
End Sub

Private Sub H1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mX = X
mY = Y
End Sub

Private Sub H1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
H1.Move H1.Left + X - mX, H1.Top + Y - mY
lv1.SelectedItem.SubItems(1) = H1.Left / 15
lv1.SelectedItem.SubItems(3) = H1.Top / 15
lnTex.X1 = lv1.SelectedItem.SubItems(1) * 15
lnTex.Y1 = lv1.SelectedItem.SubItems(3) * 15
End If
End Sub

Private Sub H2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mX = X
mY = Y
End Sub

Private Sub H2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
H2.Move H2.Left + X - mX, H2.Top + Y - mY
lv1.SelectedItem.SubItems(2) = H2.Left / 15
lv1.SelectedItem.SubItems(4) = H2.Top / 15
lnTex.X2 = lv1.SelectedItem.SubItems(2) * 15
lnTex.Y2 = lv1.SelectedItem.SubItems(4) * 15
End If
End Sub

Private Sub ln2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
mX = X
mY = Y
End If
End Sub

Private Sub ln2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim Index As Integer
Index = ln2.Tag
ln2.Move ln2.Left + X - mX, ln2.Top + Y - mY
tnk_toX(Index) = (ln2.Left) / 15
tnk_toY(Index) = (ln2.Top) / 15
lnwp.X2 = tnk_toX(Index) * 15
lnwp.Y2 = tnk_toY(Index) * 15
End If
End Sub


Private Sub lv1_ItemClick(ByVal Item As ComctlLib.ListItem)
On Error Resume Next
lnTex.Visible = True
H1.Visible = True
H2.Visible = True
lnTex.X1 = Item.SubItems(1) * 15
lnTex.X2 = Item.SubItems(2) * 15
lnTex.Y1 = Item.SubItems(3) * 15
lnTex.Y2 = Item.SubItems(4) * 15
H1.Left = lnTex.X1
H1.Top = lnTex.Y1
H2.Left = lnTex.X2
H2.Top = lnTex.Y2
End Sub

Private Sub lv1_LostFocus()
lnTex.Visible = False
H1.Visible = False
H2.Visible = False
End Sub

Private Sub mnu_open_Click()
frmopen.Show vbModal
End Sub

Private Sub mnu_SHE_Click()
Eves.Visible = Not Eves.Visible
Eves.ZOrder 0
End Sub

Private Sub mnu_SHTE_Click()
SSTab1.Visible = Not SSTab1.Visible
SSTab1.ZOrder 0
End Sub

Private Sub nu_exit_Click()
End
End Sub

Private Sub PicView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Label3.Visible = False
selected = False
pvw.Visible = False
lnwp.Visible = False
ln2.Visible = False
End If
End Sub

Private Sub PicView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If selected = True Then
Label3.Move X + 150, Y + 150
pvw.Move X - (pvw.Width / 2), Y - (pvw.Height / 2)
End If
Label4 = CStr(X / 15) & " '.' " & CStr(Y / 15)
End Sub

Private Sub pvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim img As String
If Label3.Tag = "tnk" Then
Load gizzly(gizzly.UBound + 1)
tnk_ini(gizzly.UBound) = Label3.ToolTipText
img = GetFromIni(Label3.ToolTipText, "image", App.path & "\rules\tanks.ini")
Set gizzly(gizzly.UBound).Picture = LoadPicture(App.path & "\images\" & img & "\" & img & "13 copy.gif")
gizzly(gizzly.UBound).Left = pvw.Left + (X / 15)
gizzly(gizzly.UBound).Top = pvw.Top + (Y / 15)
tnk_X(gizzly.UBound) = gizzly(gizzly.UBound).Left / 15 + gizzly(gizzly.UBound).Width / 30
tnk_Y(gizzly.UBound) = gizzly(gizzly.UBound).Top / 15 + gizzly(gizzly.UBound).Height / 30
tnk_toX(gizzly.UBound) = gizzly(gizzly.UBound).Left / 15
tnk_toY(gizzly.UBound) = gizzly(gizzly.UBound).Top / 15
gizzly(gizzly.UBound).ZOrder 0
tnk_side(gizzly.UBound) = Side
gizzly(gizzly.UBound).Visible = True
ElseIf Label3.Tag = "bldng" Then
Load bldng(bldng.UBound + 1)
bldng_ini(bldng.UBound) = Label3.ToolTipText
img = GetFromIni(Label3.ToolTipText, "image", App.path & "\rules\buildings.ini")
bldng(bldng.UBound).AutoSize = True
bldng(bldng.UBound).ClearImage
bldng(bldng.UBound).LoadImage_FromFile App.path & "\images\buildings\" & img & ".png"
bldng(bldng.UBound).Refresh
bldng(bldng.UBound).Left = pvw.Left + (X / 15)
bldng(bldng.UBound).Top = pvw.Top + (Y / 15)
bldng_X(bldng.UBound) = bldng(bldng.UBound).Left / 15
bldng_Y(bldng.UBound) = bldng(bldng.UBound).Top / 15
bldng_side(bldng.UBound) = Side
bldng(bldng.UBound).ZOrder 0
bldng(bldng.UBound).Visible = True
ElseIf Label3.ToolTipText = "Trees" Then
Load tree(tree.UBound + 1)
tree(tree.UBound).Move pvw.Left, pvw.Top
tree(tree.UBound).Visible = True
img = Round(Rnd * 23)
Set tree(tree.UBound).Picture = LoadPicture(App.path & "\images\trees\tree (" & img & ").gif")
End If
ElseIf Button = 2 Then
Label3.Visible = False
selected = False
pvw.Visible = False
lnwp.Visible = False
ln2.Visible = False
End If
pvw.ZOrder 0
End Sub

Private Sub pvw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pvw.Move X + pvw.Left - (pvw.Width / 2), Y + pvw.Top - (pvw.Height / 2)
Label3.Move pvw.Left + 150, pvw.Top + 150
If Label3.ToolTipText = "Trees" And Button = 1 And (Fix(X / 150) = (X / 150) Or Fix(Y / 150) = (Y / 150)) Then
Load tree(tree.UBound + 1)
tree(tree.UBound).Move pvw.Left, pvw.Top
tree(tree.UBound).Visible = True
img = Round(Rnd * 23)
Set tree(tree.UBound).Picture = LoadPicture(App.path & "\images\trees\tree (" & img & ").gif")
End If
Label4 = CStr((pvw.Left / 15) + X / 15) & " '.' " & CStr((pvw.Top / 15) + Y / 15)
End Sub



Private Sub terr_Click(Index As Integer)
Label4 = CStr((terr(Index).Left / 15) + X / 15) & " '.' " & CStr((terr(Index).Top / 15) + Y / 15) & " '.' " & CStr(Index)
End Sub


Private Sub tex_height_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
mY = Y
End If
End Sub

Private Sub tex_height_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
tex_height.Move tex_height.Left, tex_height.Top + Y - mY
Texture.Height = (tex_height.Top - Texture.Top) + tex_height.Height
lv1.ListItems(Val(Texture.Tag)).SubItems(2) = Texture.Height / 15
End If
End Sub

Private Sub tex_width_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
mX = X
End If
End Sub

Private Sub tex_width_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
tex_width.Move tex_width.Left + X - mX
Texture.Width = (tex_width.Left - Texture.Left) + tex_width.Width
lv1.ListItems(Val(Texture.Tag)).SubItems(1) = Texture.Width / 15
End If
End Sub

Private Sub tex_mid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
mX = X
mY = Y
End If
End Sub

Private Sub tex_mid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
tex_mid.Move tex_mid.Left + X - mX, tex_mid.Top + Y - mY
tex_width.Move tex_width.Left + X - mX, tex_width.Top + Y - mY
tex_height.Move tex_height.Left + X - mX, tex_height.Top + Y - mY
Texture.Move Texture.Left + X - mX, Texture.Top + Y - mY
lv1.ListItems(Val(Texture.Tag)).SubItems(3) = tex_mid.Left / 15
lv1.ListItems(Val(Texture.Tag)).SubItems(4) = tex_mid.Top / 15
End If
End Sub

Private Sub Text1_Change()
PicView.Width = Val(Text1 * 15)
End Sub

Private Sub Text12_Change()
tnk_side(Val(Picture3.Tag)) = Text12
End Sub

Private Sub Text2_Change()
PicView.Height = Val(Text2 * 15)
End Sub

Private Sub Text8_Change()
bldng_side(Val(Picture2.Tag)) = Text8
End Sub

Private Sub Timer1_Timer()
GetCursorPos PT
If PT.X < 3 Then
If PicView.Left < 0 Then
PicView.Left = PicView.Left + 225
End If
ElseIf PT.Y < 3 Then
If PicView.Top < 0 Then
PicView.Top = PicView.Top + 225
End If
ElseIf PT.X > (Screen.Width / 15) - 3 Then
If PicView.Left + PicView.Width > Me.Width Then
PicView.Left = PicView.Left - 225
End If
ElseIf PT.Y > (Screen.Height / 15) - 3 Then
If PicView.Top + PicView.Height > Me.Height Then
PicView.Top = PicView.Top - 225
End If
End If
End Sub

Private Sub TnkLst_ItemClick(ByVal Item As ComctlLib.ListItem)
Label3.Visible = True
Label3.ToolTipText = Item.Text
Label3.Tag = "tnk"
Dim img As String
pvw.Visible = True
pvw.AutoSize = True
img = (GetFromIni(Item.Text, "image", App.path & "\rules\tanks.ini"))
pvw.LoadImage_FromFile App.path & "\images\" & img & "\" & img & "13 copy.gif"
selected = True
End Sub

Private Sub tree_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
tree(Index).Tag = "-1"
tree(Index).Visible = False
End If
End Sub

Private Sub tree_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4 = CStr((tree(Index).Left / 15) + X / 15) & " '.' " & CStr((tree(Index).Top / 15) + Y / 15) & " '.' " & CStr(Index + 1)
End Sub

Private Sub TrnLst_ItemClick(ByVal Item As ComctlLib.ListItem)
Label3.Visible = True
Label3.ToolTipText = Item.Text
Label3.Tag = "terr"
Dim img As String
pvw.Visible = True
pvw.AutoSize = True
img = (Round(Rnd * 23))
pvw.LoadImage_FromFile App.path & "\images\Trees\Tree (" & img & ").gif"
selected = True
End Sub


Sub FillLst()
Dim X As Integer
For X = 1 To Val(GetFromIni("Main", "count", App.path & "\Rules\tanks.ini"))
TnkLst.ListItems.Add , , GetFromIni("Main", "t" & CStr(X), App.path & "\rules\tanks.ini")
Next
For X = 1 To Val(GetFromIni("Main", "count", App.path & "\Rules\buildings.ini"))
BldngLst.ListItems.Add , , GetFromIni("Main", "b" & CStr(X), App.path & "\rules\buildings.ini")
Next
FileTex.path = App.path & "\images\texture\"
TrnLst.ListItems.Add , , "Trees"
End Sub

Function pthfix(ByVal pth As String)
On Error Resume Next
If Right(pth, 1) = "\" Then
pthfix = pth
Else
pthfix = pth & "\"
End If
End Function

Private Sub mnu_save_Click()
cd.ShowSave
Dim str As String, X As Integer
If cd.FileTitle <> "" Then
str = cd.FileName
MkDir str
WriteIni "Main", "Width", Text1, str & "\ini.ini"
WriteIni "Main", "Height", Text2, str & "\ini.ini"
WriteIni "Main", "Name", Text3, str & "\ini.ini"
WriteIni "Main", "Money", Text6, str & "\ini.ini"
WriteIni "Main", "lightoffset", Text17, str & "\ini.ini"
WriteIni "Main", "AI Skills", Text18, str & "\ini.ini"
WriteIni "Main", "Ground", FileTex.List(FileTex.ListIndex), str & "\ini.ini"

Desc.SaveFile str & "\Description.txt", 1
WriteIni "Masks", "Count", lv1.ListItems.Count, str & "\ini.ini"

For X = 1 To lv1.ListItems.Count
WriteIni "Masks", "X1:" & CStr(X), lv1.ListItems(X).SubItems(1), str & "\ini.ini"
WriteIni "Masks", "X2:" & CStr(X), lv1.ListItems(X).SubItems(2), str & "\ini.ini"
WriteIni "Masks", "Y1:" & CStr(X), lv1.ListItems(X).SubItems(3), str & "\ini.ini"
WriteIni "Masks", "Y2:" & CStr(X), lv1.ListItems(X).SubItems(4), str & "\ini.ini"
WriteIni "Masks", "Type:" & CStr(X), lv1.ListItems(X).SubItems(5), str & "\ini.ini"
Next

lstdum.Clear
For X = 1 To gizzly.UBound
If tnk_side(X) <> "-1" And X <> 0 Then
lstdum.AddItem X
End If
Next
WriteIni "Tanks", "Count", CStr(lstdum.ListCount), str & "\ini.ini"
For X = 0 To lstdum.ListCount - 1
WriteIni "Tanks", "Side" & CStr(X + 1), tnk_side(Val(lstdum.List(X))), str & "\ini.ini"
WriteIni "Tanks", "ini" & CStr(X + 1), tnk_ini(Val(lstdum.List(X))), str & "\ini.ini"
WriteIni "Tanks", "X" & CStr(X + 1), CStr(tnk_X(Val(lstdum.List(X)))), str & "\ini.ini"
WriteIni "Tanks", "Y" & CStr(X + 1), CStr(tnk_Y(Val(lstdum.List(X)))), str & "\ini.ini"
WriteIni "Tanks", "toX" & CStr(X + 1), CStr(tnk_toX(Val(lstdum.List(X)))), str & "\ini.ini"
WriteIni "Tanks", "toY" & CStr(X + 1), CStr(tnk_toY(Val(lstdum.List(X)))), str & "\ini.ini"
Next

lstdum.Clear
For X = 1 To bldng.UBound
If bldng_side(X) <> "-1" And X <> 0 Then
lstdum.AddItem X
End If
Next
WriteIni "Buildings", "Count", CStr(lstdum.ListCount), str & "\ini.ini"
For X = 0 To lstdum.ListCount - 1
WriteIni "Buildings", "Side" & CStr(X + 1), bldng_side(Val(lstdum.List(X))), str & "\ini.ini"
WriteIni "Buildings", "ini" & CStr(X + 1), bldng_ini(Val(lstdum.List(X))), str & "\ini.ini"
WriteIni "Buildings", "X" & CStr(X + 1), CStr(bldng_X(Val(lstdum.List(X)))), str & "\ini.ini"
WriteIni "Buildings", "Y" & CStr(X + 1), CStr(bldng_Y(Val(lstdum.List(X)))), str & "\ini.ini"
WriteIni "Buildings", "Flip" & CStr(X + 1), CStr(bldng_flip(Val(lstdum.List(X)))), str & "\ini.ini"
Next

WriteIni "Events", "Count", lsteves.ListItems.Count, str & "\ini.ini"
For X = 1 To lsteves.ListItems.Count
WriteIni "Events", "On" & CStr(X), lsteves.ListItems(X).Text, str & "\ini.ini"
WriteIni "Events", "Do" & CStr(X), lsteves.ListItems(X).SubItems(1), str & "\ini.ini"
Next

WriteIni "Timer", "Count", lsttime.ListItems.Count, str & "\ini.ini"
For X = 1 To lsttime.ListItems.Count
WriteIni "Timer", "Trigger" & CStr(X), lsttime.ListItems(X).Text, str & "\ini.ini"
Next
WriteIni "Timer", "Label", Text4, str & "\ini.ini"
WriteIni "Timer", "Time", Text5, str & "\ini.ini"
If Check1.Value = 1 Then
WriteIni "Timer", "Loop", "True", str & "\ini.ini"
Else
WriteIni "Timer", "Loop", "False", str & "\ini.ini"
End If
End If

lstdum.Clear
For X = tree.LBound To tree.UBound
If tree(X).Tag <> "-1" And X <> 0 Then
lstdum.AddItem CStr(X)
End If
Next
WriteIni "Trees", "Count", lstdum.ListCount, str & "\ini.ini"
For X = 0 To lstdum.ListCount
WriteIni "Trees", "TreeX" & CStr(X + 1), CStr(tree(Val(lstdum.List(X))).Left / 15), str & "\ini.ini"
WriteIni "Trees", "TreeY" & CStr(X + 1), CStr(tree(Val(lstdum.List(X))).Top / 15), str & "\ini.ini"
Next
End Sub
Sub oopen(path As String)
On Error Resume Next
cleanAll
path = pthfix(path)
Dim ini As String, k As Integer, i As Integer, ltm As ListItem, img As String
ini = path & "ini.ini"
Text1 = GetFromIni("Main", "width", ini)
Text2 = GetFromIni("Main", "height", ini)
Text3 = GetFromIni("Main", "name", ini)
Text6 = GetFromIni("Main", "money", ini)
Text17 = GetFromIni("Main", "lightoffset", ini)
Text18 = GetFromIni("Main", "AI Skills", ini)
Desc.LoadFile path & "Description.txt", 1

lv1.ListItems.Clear
For k = 1 To GetFromIni("Masks", "count", ini)
Set ltm = lv1.ListItems.Add(, , CStr(k))
ltm.SubItems(1) = GetFromIni("Masks", "X1:" & CStr(k), ini)
ltm.SubItems(2) = GetFromIni("Masks", "X2:" & CStr(k), ini)
ltm.SubItems(3) = GetFromIni("Masks", "Y1:" & CStr(k), ini)
ltm.SubItems(4) = GetFromIni("Masks", "Y2:" & CStr(k), ini)
ltm.SubItems(5) = GetFromIni("Masks", "Type:" & CStr(k), ini)
Next

For i = 0 To FileTex.ListCount - 1
If FileTex.List(i) = GetFromIni("Main", "Ground", ini) Then
FileTex.ListIndex = i
Exit For
End If
Next
lsttime.ListItems.Clear
For k = 1 To GetFromIni("Timer", "count", ini)
lsttime.ListItems.Add , , GetFromIni("Timer", "Trigger" & CStr(k), ini)
Next
Text4 = GetFromIni("Timer", "label", ini)
Text5 = GetFromIni("Timer", "time", ini)
Check1 = Val(GetFromIni("Timer", "loop", ini))

lsteves.ListItems.Clear
For k = 1 To GetFromIni("Events", "count", ini)
Set ltm = lsteves.ListItems.Add(, , GetFromIni("Events", "On" & CStr(k), ini))
ltm.SubItems(1) = GetFromIni("Events", "Do" & CStr(k), ini)
Next

For i = 1 To GetFromIni("Tanks", "count", ini)
Load gizzly(gizzly.UBound + 1)
gizzly(gizzly.UBound).Left = GetFromIni("Tanks", "X" & CStr(i), ini) * 15
gizzly(gizzly.UBound).Top = GetFromIni("Tanks", "Y" & CStr(i), ini) * 15
tnk_X(gizzly.UBound) = GetFromIni("Tanks", "X" & CStr(i), ini)
tnk_Y(gizzly.UBound) = GetFromIni("Tanks", "Y" & CStr(i), ini)
tnk_toX(gizzly.UBound) = GetFromIni("Tanks", "ToX" & CStr(i), ini)
tnk_toY(gizzly.UBound) = GetFromIni("Tanks", "ToX" & CStr(i), ini)
tnk_side(gizzly.UBound) = GetFromIni("Tanks", "Side" & CStr(i), ini)
tnk_ini(gizzly.UBound) = GetFromIni("Tanks", "INI" & CStr(i), ini)
img = GetFromIni(tnk_ini(gizzly.UBound), "image", App.path & "\rules\tanks.ini")
Set gizzly(gizzly.UBound).Picture = LoadPicture(App.path & "\images\" & img & "\" & img & "13 copy.gif")
gizzly(gizzly.UBound).Visible = True
Next

For i = 1 To GetFromIni("Buildings", "count", ini)
Load bldng(bldng.UBound + 1)
bldng(bldng.UBound).Left = GetFromIni("Buildings", "X" & CStr(i), ini) * 15
bldng(bldng.UBound).Top = GetFromIni("Buildings", "Y" & CStr(i), ini) * 15
bldng_X(bldng.UBound) = GetFromIni("Buildings", "X" & CStr(i), ini)
bldng_Y(bldng.UBound) = GetFromIni("Buildings", "Y" & CStr(i), ini)
bldng_side(bldng.UBound) = GetFromIni("Buildings", "Side" & CStr(i), ini)
bldng_flip(bldng.UBound) = GetFromIni("Buildings", "flip" & CStr(i), ini)
bldng_ini(bldng.UBound) = GetFromIni("Buildings", "INI" & CStr(i), ini)
img = GetFromIni(bldng_ini(bldng.UBound), "image", App.path & "\rules\buildings.ini")
bldng(bldng.UBound).AutoSize = True
bldng(bldng.UBound).ClearImage
bldng(bldng.UBound).LoadImage_FromFile App.path & "\images\buildings\" & img & ".png"
bldng(bldng.UBound).Refresh
bldng(bldng.UBound).Visible = True
Next

For k = 1 To GetFromIni("Trees", "count", ini)
Load tree(tree.UBound + 1)
tree(tree.UBound).Move GetFromIni("Trees", "TreeX" & CStr(k), ini) * 15, GetFromIni("Trees", "TreeY" & CStr(k), ini) * 15
img = Round(Rnd * 23)
Set tree(tree.UBound).Picture = LoadPicture(App.path & "\images\trees\tree (" & img & ").gif")
tree(tree.UBound).Visible = True
Next
End Sub

Sub cleanAll()
On Error Resume Next
Dim i As Integer
For i = 1 To gizzly.UBound
tnk_side(i) = "-1"
tnk_ini(i) = ""
tnk_X(i) = 0
tnk_Y(i) = 0
tnk_toX(i) = 0
tnk_toY(i) = 0
Unload gizzly(i)
Next

For i = 1 To bldng.UBound
bldng_side(i) = "-1"
bldng_ini(i) = ""
bldng_X(i) = 0
bldng_Y(i) = 0
bldng_flip(i) = 0
Unload bldng(i)
Next

For i = 1 To tree.UBound
tree(i).Tag = "-1"
tree(i).Visible = False
Next
End Sub

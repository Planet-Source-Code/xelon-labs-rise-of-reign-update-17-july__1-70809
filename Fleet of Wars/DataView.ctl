VERSION 5.00
Begin VB.UserControl DataView 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image img 
      Height          =   480
      Index           =   0
      Left            =   -480
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "DataView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event click(str As String)

Public sel As String
Public air As String

Enum mode
building = 0
tank = 1
aircraft = 2
End Enum

Sub Add(str As String, Optional typ As mode = tank)
On Error GoTo U
Load img(img.UBound + 1)
If str <> "REPAIR" Then
If typ = building Then
Set img(img.UBound).Picture = LoadPicture(App.path & "\images\buildings\" & GetFromIni(str, "image", App.path & "\rules\buildings.ini") & "_ico.gif")
ElseIf typ = tank Then
Set img(img.UBound).Picture = LoadPicture(App.path & "\images\" & GetFromIni(str, "image", App.path & "\rules\tanks.ini") & "\ico.gif")
ElseIf typ = aircraft Then
Set img(img.UBound).Picture = LoadPicture(App.path & "\images\" & GetFromIni(str, "image", App.path & "\rules\aircrafts.ini") & "\ico.gif")
End If
Else
Set img(img.UBound).Picture = LoadPicture(App.path & "\images\buildbar\repair.jpg")
End If
img(img.UBound).Tag = str

If img(img.UBound - 1).Left + img(img.UBound - 1).Width + 75 > UserControl.Width Then
img(img.UBound).Left = 0
img(img.UBound).Top = img(img.UBound - 1).Top + img(img.UBound - 1).Height
Else
img(img.UBound).Left = img(img.UBound - 1).Left + img(img.UBound).Width + 75
img(img.UBound).Top = img(img.UBound - 1).Top
End If
img(img.UBound).Visible = True
Exit Sub
U:
MsgBox Err.Description
End Sub

Sub clear()
On Error Resume Next
Dim k As Integer
For k = 1 To img.UBound
Unload img(k)
Next
Refresh
End Sub

Private Sub img_Click(Index As Integer)
sel = img(Index).Tag
RaiseEvent click(img(Index).Tag)
End Sub

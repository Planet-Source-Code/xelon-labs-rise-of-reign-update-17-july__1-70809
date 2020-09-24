VERSION 5.00
Begin VB.UserControl weapons 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "weapons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim many As Integer
Dim M_weap(100) As String
Dim m_damage(100) As Integer
Dim m_typ(100) As String
Dim m_interval(100) As Integer
Dim m_range(100) As Integer
Dim m_distance(100) As Integer
Dim m_speedstep(100) As Integer
Dim m_color(100) As Long

 Function damage(weap As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(M_weap(n)) = UCase(weap) Then
damage = m_damage(n)
Exit For
End If
Next
End Function
Function typ(weap As String) As String
Dim n As Integer
For n = 1 To many
If UCase(M_weap(n)) = UCase(weap) Then
typ = m_typ(n)
Exit For
End If
Next
End Function
Function interval(weap As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(M_weap(n)) = UCase(weap) Then
interval = m_interval(n)
Exit For
End If
Next
End Function
Function range(weap As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(M_weap(n)) = UCase(weap) Then
range = m_range(n)
Exit For
End If
Next
End Function
Function distance(weap As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(M_weap(n)) = UCase(weap) Then
distance = m_distance(n)
Exit For
End If
Next
End Function
Function speedstep(weap As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(M_weap(n)) = UCase(weap) Then
speedstep = m_speedstep(n)
Exit For
End If
Next
End Function
Function Color(weap As String) As Long
Dim n As Integer
For n = 1 To many
If UCase(M_weap(n)) = UCase(weap) Then
Color = m_color(n)
Exit For
End If
Next
End Function
Sub loadwep()
Dim cnt As Integer
Dim n As Integer
Dim hd As String
Dim pth As String
pth = App.Path & "\rules\weapons.ini"
cnt = Val(GetFromIni("Main", "count", pth))
For n = 1 To cnt
M_weap(n) = GetFromIni("Main", "w" & CStr(n), pth)
m_damage(n) = Val(GetFromIni(M_weap(n), "damage", pth))
m_typ(n) = GetFromIni(M_weap(n), "type", pth)
If UCase(m_typ(n)) = "LASER" Then
m_color(n) = CLng(GetFromIni(M_weap(n), "color", pth))
Else
m_interval(n) = GetFromIni(M_weap(n), "interval", pth)
m_speedstep(n) = Val(GetFromIni(M_weap(n), "speedstep", pth))
End If
m_range(n) = Val(GetFromIni(M_weap(n), "range", pth))
m_distance(n) = Val(CInt(GetFromIni(M_weap(n), "distance", pth)))
Next
many = cnt
End Sub


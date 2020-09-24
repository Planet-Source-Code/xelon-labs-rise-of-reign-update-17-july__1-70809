VERSION 5.00
Begin VB.UserControl aeroplanes 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "aeroplanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim many  As Integer

Dim m_air(100) As String
Dim m_img(100) As String
Dim m_speed(100) As Integer
Dim m_power(100) As Integer
Dim m_weapon(100) As String
Dim m_cost(100) As Integer
Dim m_techlevel(100) As Integer

Sub loadair()
Dim cnt As Integer
Dim n As Integer
Dim hd As String
Dim pth As String
pth = App.Path & "\rules\aircrafts.ini"
cnt = Val(GetFromIni("Main", "count", pth))
For n = 1 To cnt
m_air(n) = GetFromIni("Main", "a" & CStr(n), pth)
m_img(n) = GetFromIni(m_air(n), "Image", pth)
m_power(n) = Val(GetFromIni(m_air(n), "Power", pth))
m_speed(n) = Val(GetFromIni(m_air(n), "speed", pth))
m_weapon(n) = GetFromIni(m_air(n), "weapon", pth)
m_cost(n) = Val(GetFromIni(m_air(n), "cost", pth))
m_techlevel(n) = Val(GetFromIni(m_air(n), "techlevel", pth))
Next
many = cnt
End Sub

Function image(air As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_air(n)) = UCase(air) Then
image = m_img(n)
Exit For
End If
Next
End Function

Function power(air As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(m_air(n)) = UCase(air) Then
power = m_power(n)
Exit For
End If
Next
End Function
Function speed(air As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_air(n)) = UCase(air) Then
speed = m_speed(n)
Exit For
End If
Next
End Function

Function weapon(air As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_air(n)) = UCase(air) Then
weapon = m_weapon(n)
Exit For
End If
Next
End Function

Function cost(air As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_air(n)) = UCase(air) Then
cost = m_cost(n)
Exit For
End If
Next
End Function

Function techlevel(air As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_air(n)) = UCase(air) Then
techlevel = m_techlevel(n)
Exit For
End If
Next
End Function


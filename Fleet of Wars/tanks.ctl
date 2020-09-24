VERSION 5.00
Begin VB.UserControl tanks 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "tanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim many As Integer

Dim m_tank(100) As String
Dim m_image(100) As String
Dim m_power(100) As Integer
Dim m_speed(100) As Integer
Dim m_cost(100) As Integer
Dim m_techlevel(100) As Integer
Dim m_water(100) As Integer
Dim m_weapon(100) As String

Sub loadtnx()
Dim cnt As Integer
Dim n As Integer
Dim hd As String
Dim pth As String
pth = App.Path & "\rules\tanks.ini"
cnt = Val(GetFromIni("Main", "count", pth))
For n = 1 To cnt
m_tank(n) = GetFromIni("Main", "t" & CStr(n), pth)
m_image(n) = GetFromIni(m_tank(n), "Image", pth)
m_power(n) = Val(GetFromIni(m_tank(n), "Power", pth))
m_speed(n) = Val(GetFromIni(m_tank(n), "speed", pth))
m_cost(n) = Val(GetFromIni(m_tank(n), "cost", pth))
m_techlevel(n) = Val(GetFromIni(m_tank(n), "techlevel", pth))
m_water(n) = Val(GetFromIni(m_tank(n), "water", pth))
m_weapon(n) = GetFromIni(m_tank(n), "weapon", pth)
Next
many = cnt
End Sub
Function image(tank As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_tank(n)) = UCase(tank) Then
image = m_image(n)
Exit For
End If
Next
End Function
Function power(tank As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(m_tank(n)) = UCase(tank) Then
power = m_power(n)
Exit For
End If
Next
End Function
Function speed(tank As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_tank(n)) = UCase(tank) Then
speed = m_speed(n)
Exit For
End If
Next
End Function
Function weapon(tank As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_tank(n)) = UCase(tank) Then
weapon = m_weapon(n)
Exit For
End If
Next
End Function
Function cost(tank As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_tank(n)) = UCase(tank) Then
cost = m_cost(n)
Exit For
End If
Next
End Function
Function techlevel(tank As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_tank(n)) = UCase(tank) Then
techlevel = m_techlevel(n)
Exit For
End If
Next
End Function
Function water(tank As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_tank(n)) = UCase(tank) Then
water = m_water(n)
Exit For
End If
Next
End Function

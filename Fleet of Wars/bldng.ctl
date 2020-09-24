VERSION 5.00
Begin VB.UserControl struc 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "struc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim many As Integer

Dim m_bldng(100) As String
Dim m_power(100) As Integer
Dim m_img(100) As String
Dim m_offX(100) As Integer
Dim m_DocX(100) As Integer
Dim m_DocY(100) As Integer
Dim m_offY(100) As Integer
Dim m_typ(100) As String
Dim m_weapon(100) As String
Dim m_water(100) As Integer
Dim m_cost(100) As Long
Dim m_tchlvl(100) As Integer

Sub loadbldng()
Dim cnt As Integer
Dim n As Integer
Dim hd As String
Dim pth As String
pth = App.path & "\rules\buildings.ini"
cnt = Val(GetFromIni("Main", "count", pth))
For n = 1 To cnt
m_bldng(n) = GetFromIni("Main", "b" & CStr(n), pth)
m_power(n) = Val(GetFromIni(m_bldng(n), "power", pth))
m_img(n) = GetFromIni(m_bldng(n), "image", pth)
m_offX(n) = Val(GetFromIni(m_bldng(n), "offX", pth))
m_offY(n) = Val(GetFromIni(m_bldng(n), "offY", pth))
m_DocX(n) = Val(GetFromIni(m_bldng(n), "DocX", pth))
m_DocY(n) = Val(GetFromIni(m_bldng(n), "DocY", pth))
m_typ(n) = GetFromIni(m_bldng(n), "type", pth)
m_cost(n) = Val(GetFromIni(m_bldng(n), "Cost", pth))
m_water(n) = Val(GetFromIni(m_bldng(n), "water", pth))
m_tchlvl(n) = Val(GetFromIni(m_bldng(n), "techlevel", pth))
m_weapon(n) = GetFromIni(m_bldng(n), "weapon", pth)
Next
many = cnt
End Sub
Function weapon(bldng As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
weapon = m_weapon(n)
Exit For
End If
Next
End Function
Function typ(bldng As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
typ = m_typ(n)
Exit For
End If
Next
End Function
Function power(bldng As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
power = m_power(n)
Exit For
End If
Next
End Function
Function offx(bldng As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
offx = m_offX(n)
Exit For
End If
Next
End Function
Function offy(bldng As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
offy = m_offY(n)
Exit For
End If
Next
End Function
Function image(bldng As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
image = m_img(n)
Exit For
End If
Next
End Function
Function cost(bldng As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
cost = m_cost(n)
Exit For
End If
Next
End Function
Function techlevel(bldng As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
techlevel = m_tchlvl(n)
Exit For
End If
Next
End Function
Function water(bldng As String) As String
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
water = m_water(n)
Exit For
End If
Next
End Function

Function DocX(bldng As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
DocX = m_DocX(n)
Exit For
End If
Next
End Function
Function docY(bldng As String) As Integer
Dim n As Integer
For n = 1 To many
If UCase(m_bldng(n)) = UCase(bldng) Then
docY = m_DocY(n)
Exit For
End If
Next
End Function


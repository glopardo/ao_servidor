Attribute VB_Name = "PuestosTop"
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Option Explicit
Public Function TotalMatados(UserIndex As Integer) As Integer

TotalMatados = UserList(UserIndex).Faccion.Matados(0) + UserList(UserIndex).Faccion.Matados(1) + UserList(UserIndex).Faccion.Matados(2)

End Function
Public Sub RevisarTops(UserIndex As Integer)

If UserList(UserIndex).flags.Privilegios > 0 Then
    If IndexTop(Nivel, UserIndex) <> UBound(Tops, 2) Then Call SacarTop(Nivel, UserIndex)
    If IndexTop(Muertos, UserIndex) <> UBound(Tops, 2) Then Call SacarTop(Muertos, UserIndex)
Else
    If UserList(UserIndex).Stats.ELV > Tops(Nivel, UBound(Tops, 2)).Nivel Then Call AgregarTop(Nivel, UserIndex)
    If TotalMatados(UserIndex) > Tops(Muertos, UBound(Tops, 2)).Muertos Then Call AgregarTop(Muertos, UserIndex)
End If

End Sub
Public Function IndexTop(Top As Byte, UserIndex As Integer) As Integer
Dim i As Integer

For i = 1 To UBound(Tops, 2)
    If UCase$(Tops(Top, i).Nombre) = UCase$(UserList(UserIndex).Name) Then
        IndexTop = i
        Exit Function
    End If
Next

IndexTop = UBound(Tops, 2)

End Function
Public Sub AgregarTop(Top As Byte, UserIndex As Integer)
Dim i As Integer

i = IndexTop(Top, UserIndex)

For i = i - 1 To 1 Step -1
    If (Top = Nivel And UserList(UserIndex).Stats.ELV <= Tops(Nivel, i).Nivel) Or _
        (Top = Muertos And TotalMatados(UserIndex) <= Tops(Muertos, i).Muertos) Then
        i = i + 1
        Exit For
    End If
    Tops(Top, i + 1) = Tops(Top, i)
    Call SaveTop(Top, i + 1)
Next

i = Maximo(1, i)

Tops(Top, i).Nombre = UserList(UserIndex).Name
Tops(Top, i).Bando = ListaBandos(UserList(UserIndex).Faccion.Bando)
Tops(Top, i).Nivel = UserList(UserIndex).Stats.ELV
Tops(Top, i).Muertos = TotalMatados(UserIndex)
Call SaveTop(Top, i)

End Sub
Public Sub SacarTop(Top As Byte, UserIndex As Integer)
Dim i As Integer

i = IndexTop(Top, UserIndex)

For i = i To UBound(Tops, 2) - 1
    Tops(Top, i) = Tops(Top, i + 1)
    Call SaveTop(Top, i)
Next

Tops(Top, UBound(Tops, 2)).Nombre = ""
Tops(Top, UBound(Tops, 2)).Bando = ""
Tops(Top, UBound(Tops, 2)).Nivel = 0
Tops(Top, UBound(Tops, 2)).Muertos = 0
Call SaveTop(Top, UBound(Tops, 2))

End Sub
Public Sub SaveTop(Top As Byte, Puesto As Integer)
Dim file As String
Dim i As Integer

If Len(Tops(Top, Puesto).Nombre) = 0 Then Exit Sub

If Top = Nivel Then
    file = App.Path & "\LOGS\TopNivel.log"
Else: file = App.Path & "\LOGS\TopMuertos.log"
End If

Call WriteVar(file, "Top" & Puesto, "Name", Tops(Top, Puesto).Nombre)
Call WriteVar(file, "Top" & Puesto, "Nivel", val(Tops(Top, Puesto).Nivel))
Call WriteVar(file, "Top" & Puesto, "Muertos", val(Tops(Top, Puesto).Muertos))
Call WriteVar(file, "Top" & Puesto, "Bando", Tops(Top, Puesto).Bando)

End Sub
Public Sub LoadTops(Top As Byte)
Dim file As String, i As Integer

If Top = Nivel Then
    file = App.Path & "\LOGS\TopNivel.log"
Else: file = App.Path & "\LOGS\TopMuertos.log"
End If

If Not FileExist(file, vbNormal) Then Exit Sub

For i = 1 To UBound(Tops, 2)
    Tops(Top, i).Nombre = GetVar(file, "Top" & i, "Name")
    Tops(Top, i).Nivel = val(GetVar(file, "Top" & i, "Nivel"))
    Tops(Top, i).Muertos = val(GetVar(file, "Top" & i, "Muertos"))
    Tops(Top, i).Bando = GetVar(file, "Top" & i, "Bando")
Next

End Sub


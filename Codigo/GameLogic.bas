Attribute VB_Name = "Extra"
'F�nixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Option Explicit
Public Function EsNewbie(Userindex As Integer) As Boolean

EsNewbie = (UserList(Userindex).Stats.ELV <= LimiteNewbie)

End Function
Public Sub DoTileEvents(Userindex As Integer)
On Error GoTo errhandler
Dim Map As Integer, X As Integer, Y As Integer
Dim nPos As WorldPos, mPos As WorldPos

Map = UserList(Userindex).POS.Map
X = UserList(Userindex).POS.X
Y = UserList(Userindex).POS.Y

mPos = MapData(Map, X, Y).TileExit
If Not MapaValido(mPos.Map) Or Not InMapBounds(mPos.X, mPos.Y) Then Exit Sub

If MapInfo(mPos.Map).Restringir And Not EsNewbie(Userindex) Then
    Call SendData(ToIndex, Userindex, 0, "1J")
ElseIf UserList(Userindex).Stats.ELV < MapInfo(mPos.Map).Nivel And Not (UserList(Userindex).Clase = PIRATA And UserList(Userindex).Recompensas(1) = 2) Then
    Call SendData(ToIndex, Userindex, 0, "%/" & MapInfo(mPos.Map).Nivel)
Else
    If LegalPos(mPos.Map, mPos.X, mPos.Y, PuedeAtravesarAgua(Userindex)) Then
        If mPos.X <> 0 And mPos.Y <> 0 Then Call WarpUserChar(Userindex, mPos.Map, mPos.X, mPos.Y, ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT)
    Else
        Call ClosestStablePos(mPos, nPos)
        If nPos.X <> 0 And nPos.Y Then Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT)
    End If
    Exit Sub
End If

Call ClosestStablePos(UserList(Userindex).POS, nPos)
If nPos.X <> 0 And nPos.Y Then Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT)

Exit Sub

errhandler:
    Call LogError("Error en DoTileEvents-" & nPos.Map & "-" & nPos.X & "-" & nPos.Y)

End Sub
Function InMapBounds(X As Integer, Y As Integer) As Boolean

InMapBounds = (X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder)

End Function
Sub ClosestStablePos(POS As WorldPos, ByRef nPos As WorldPos)
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = POS.Map

Do While Not LegalPos(POS.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = POS.Y - LoopC To POS.Y + LoopC
        For tX = POS.X - LoopC To POS.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                nPos.X = tX
                nPos.Y = tY

                tX = POS.X + LoopC
                tY = POS.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub
Sub ClosestLegalPos(POS As WorldPos, nPos As WorldPos, Optional AguaValida As Boolean)
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = POS.Map

Do While Not LegalPos(POS.Map, nPos.X, nPos.Y, AguaValida)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = POS.Y - LoopC To POS.Y + LoopC
        For tX = POS.X - LoopC To POS.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY, AguaValida) Then
                nPos.X = tX
                nPos.Y = tY
                
                
                tX = POS.X + LoopC
                tY = POS.Y + LoopC
  
            End If
        
        Next
    Next
    
    LoopC = LoopC + 1
    
Loop

If Notfound Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub
Function ClaseIndex(ByVal Clase As String) As Integer
Dim i As Integer

For i = 1 To UBound(ListaClases)
    If UCase$(ListaClases(i)) = UCase$(Clase) Then
        ClaseIndex = i
        Exit Function
    End If
Next

End Function
Function NameIndex(ByVal Name As String) As Integer
Dim Userindex As Integer, i As Integer

Name = Replace$(Name, "+", " ")

If Len(Name) = 0 Then
    NameIndex = 0
    Exit Function
End If
  
Userindex = 1

If Right$(Name, 1) = "*" Then
    Name = Left$(Name, Len(Name) - 1)
    For i = 1 To LastUser
        If UCase$(UserList(i).Name) = UCase$(Name) Then
            NameIndex = i
            Exit Function
        End If
    Next
Else
    For i = 1 To LastUser
        If UCase$(Left$(UserList(i).Name, Len(Name))) = UCase$(Name) Then
            NameIndex = i
            Exit Function
        End If
    Next
End If

End Function
Function CheckForSameIP(Userindex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged Then
        If UserList(LoopC).ip = UserIP And Userindex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next

End Function
Function CheckForSameName(Userindex As Integer, ByVal Name As String) As Boolean
Dim LoopC As Integer

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged Then
        If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next

End Function
Sub HeadtoPos(Head As Byte, POS As WorldPos)
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = POS.X
Y = POS.Y

If Head = NORTH Then
    nX = X
    nY = Y - 1
End If

If Head = SOUTH Then
    nX = X
    nY = Y + 1
End If

If Head = EAST Then
    nX = X + 1
    nY = Y
End If

If Head = WEST Then
    nX = X - 1
    nY = Y
End If

POS.X = nX
POS.Y = nY

End Sub
Function LegalPos(Map As Integer, X As Integer, Y As Integer, Optional PuedeAgua As Boolean) As Boolean

If Not MapaValido(Map) Or Not InMapBounds(X, Y) Then Exit Function

LegalPos = (MapData(Map, X, Y).Blocked = 0) And _
           (MapData(Map, X, Y).Userindex = 0) And _
           (MapData(Map, X, Y).NpcIndex = 0) And _
           (MapData(Map, X, Y).Agua = Buleano(PuedeAgua))

End Function
Function LegalPosNPC(Map As Integer, X As Integer, Y As Integer, AguaValida As Boolean) As Boolean

If Not InMapBounds(X, Y) Then Exit Function

LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).Userindex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).trigger <> POSINVALIDA) _
     And Buleano(AguaValida) = MapData(Map, X, Y).Agua
     
End Function
Public Sub SendNPC(Userindex As Integer, NpcIndex As Integer)
Dim Info As String
Dim CRI As Byte

Select Case UserList(Userindex).Stats.UserSkills(Supervivencia)
    Case Is <= 20
        If Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
            CRI = 5
        Else: CRI = 1
        End If
    Case Is < 40
        Select Case 100 * Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP
            Case 100
                CRI = 5
            Case Is >= 50
                CRI = 2
            Case Else
                CRI = 3
        End Select
    Case Is < 60
        Select Case 100 * Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP
            Case 100
                CRI = 5
            Case Is > 66
                CRI = 2
            Case Is > 33
                CRI = 3
            Case Else
                CRI = 4
        End Select
    Case Is < 100
        CRI = 5 + Fix(10 * (1 - (Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP)))
    Case Else
        Info = "||" & Npclist(NpcIndex).Name & " [" & Npclist(NpcIndex).Stats.MinHP & "/" & Npclist(NpcIndex).Stats.MaxHP & "]"
        If Npclist(NpcIndex).flags.Paralizado Then Info = Info & " - PARALIZADO"
        Call SendData(ToIndex, Userindex, 0, Info & FONTTYPE_INFO)
        Exit Sub
End Select

Info = "9Q" & Npclist(NpcIndex).Name & "," & CRI
Call SendData(ToIndex, Userindex, 0, Info)
                
End Sub
Public Sub Expresar(NpcIndex As Integer, Userindex As Integer)

If Npclist(NpcIndex).NroExpresiones Then
    Dim randomi
    randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
    Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "3Q" & vbWhite & "�" & Npclist(NpcIndex).Expresiones(randomi) & "�" & Npclist(NpcIndex).Char.CharIndex)
End If
                    
End Sub
Sub LookatTile(Userindex As Integer, Map As Integer, X As Integer, Y As Integer)

Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim NPMUERTO As String
Dim Info As String


If InMapBounds(X, Y) Then
    UserList(Userindex).flags.TargetMap = Map
    UserList(Userindex).flags.TargetX = X
    UserList(Userindex).flags.TargetY = Y
    
    If MapData(Map, X, Y).OBJInfo.OBJIndex Then
        
        If MapData(Map, X, Y).OBJInfo.Amount = 1 Then
            Call SendData(ToIndex, Userindex, 0, "4Q" & ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Name)
        Else
            Call SendData(ToIndex, Userindex, 0, "5Q" & ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Name & "," & MapData(Map, X, Y).OBJInfo.Amount)
        End If
        UserList(Userindex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.OBJIndex
        UserList(Userindex).flags.TargetObjMap = Map
        UserList(Userindex).flags.TargetObjX = X
        UserList(Userindex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, X + 1, Y).OBJInfo.OBJIndex Then
        
        If ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
            Call SendData(ToIndex, Userindex, 0, "6Q" & ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).Name)
            UserList(Userindex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.OBJIndex
            UserList(Userindex).flags.TargetObjMap = Map
            UserList(Userindex).flags.TargetObjX = X + 1
            UserList(Userindex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex Then
        If ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
            
            Call SendData(ToIndex, Userindex, 0, "6Q" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).Name)
            UserList(Userindex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex
            UserList(Userindex).flags.TargetObjMap = Map
            UserList(Userindex).flags.TargetObjX = X + 1
            UserList(Userindex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X, Y + 1).OBJInfo.OBJIndex Then
        If ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
            
            Call SendData(ToIndex, Userindex, 0, "6Q" & ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).Name)
            UserList(Userindex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.OBJIndex
            UserList(Userindex).flags.TargetObjMap = Map
            UserList(Userindex).flags.TargetObjX = X
            UserList(Userindex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).Userindex Then
            TempCharIndex = MapData(Map, X, Y + 1).Userindex
            FoundChar = 1
        End If
        If MapData(Map, X, Y + 1).NpcIndex Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    
    If FoundChar = 0 Then
        If MapData(Map, X, Y).Userindex Then
            TempCharIndex = MapData(Map, X, Y).Userindex
            FoundChar = 1
        End If
        If MapData(Map, X, Y).NpcIndex Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    
    If FoundChar = 1 Then
            
        If UserList(TempCharIndex).flags.AdminInvisible Then Exit Sub
        
        If UserList(TempCharIndex).Faccion.Bando Then
            If UserList(TempCharIndex).Faccion.BandoOriginal <> UserList(TempCharIndex).Faccion.Bando Then
                Stat = Stat & " <" & ListaBandos(UserList(TempCharIndex).Faccion.Bando) & "> <Mercenario>"
            ElseIf UserList(TempCharIndex).Faccion.Jerarquia Then
                Stat = Stat & " <" & ListaBandos(UserList(TempCharIndex).Faccion.Bando) & "> <" & Titulo(TempCharIndex) & ">" & Cargo(TempCharIndex)
            Else
                Stat = Stat & " <" & Titulo(TempCharIndex) & ">" & Cargo(TempCharIndex)
            End If
        End If
        
        If Len(UserList(TempCharIndex).GuildInfo.GuildName) > 0 Then
            Stat = Stat & " <" & UserList(TempCharIndex).GuildInfo.GuildName & ">"
        End If
        
        If Len(UserList(TempCharIndex).Desc) > 0 Then
            Stat = UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).Desc
        Else
            Stat = UserList(TempCharIndex).Name & Stat
        End If
        
        If UserList(TempCharIndex).flags.Privilegios Then
            Stat = "9J" & Stat
        Else
            If UserList(TempCharIndex).flags.Muerto Then
                Stat = "2K" & UserList(TempCharIndex).Name
            ElseIf UserList(TempCharIndex).Faccion.Bando = Real And UserList(TempCharIndex).flags.EsConseReal = 0 Then
                Stat = "3K" & Stat
            ElseIf UserList(TempCharIndex).Faccion.Bando = Caos And UserList(TempCharIndex).flags.EsConseCaos = 0 Then
                Stat = "4K" & Stat
            ElseIf EsNewbie(TempCharIndex) Then
                Stat = "H0" & Stat
            ElseIf UserList(TempCharIndex).Faccion.Bando = Caos And UserList(TempCharIndex).flags.EsConseCaos = 1 Then
                Stat = "H2" & Stat
            ElseIf UserList(TempCharIndex).Faccion.Bando = Real And UserList(TempCharIndex).flags.EsConseReal = 1 Then
                Stat = "H1" & Stat
            Else
                Stat = "1&" & Stat
            End If
        End If
        
        Call SendData(ToIndex, Userindex, 0, Stat)
            
        
        FoundSomething = 1
        UserList(Userindex).flags.TargetUser = TempCharIndex
        UserList(Userindex).flags.TargetNpc = 0
        UserList(Userindex).flags.TargetNpcTipo = 0
       
       
    ElseIf FoundChar = 2 Then
            
            Dim wPos As WorldPos
            wPos.Map = Map
            wPos.X = X
            wPos.Y = Y
            If Distancia(Npclist(TempCharIndex).POS, wPos) > 1 Then
                MapData(Map, X, Y).NpcIndex = 0
                Exit Sub
            End If
                
            If Npclist(TempCharIndex).flags.TiendaUser Then
                If Userindex = Npclist(TempCharIndex).flags.TiendaUser Then
                    If UserList(Userindex).Tienda.Gold Then
                        Call SendData(ToIndex, Userindex, 0, "/O" & UserList(Userindex).Tienda.Gold & "," & Npclist(TempCharIndex).Char.CharIndex)
                    Else
                        Call SendData(ToIndex, Userindex, 0, "/P" & Npclist(TempCharIndex).Char.CharIndex)
                    End If
                Else
                    Call SendData(ToIndex, Userindex, 0, "/Q" & UserList(Npclist(TempCharIndex).flags.TiendaUser).Name & "," & Npclist(TempCharIndex).Char.CharIndex)
                End If
            ElseIf Len(Npclist(TempCharIndex).Desc) > 1 Then
                Call SendData(ToIndex, Userindex, 0, "3Q" & vbWhite & "�" & Npclist(TempCharIndex).Desc & "�" & Npclist(TempCharIndex).Char.CharIndex)
            ElseIf Npclist(TempCharIndex).MaestroUser Then
                Call SendData(ToIndex, Userindex, 0, "7Q" & Npclist(TempCharIndex).Name & "," & UserList(Npclist(TempCharIndex).MaestroUser).Name)
            ElseIf Npclist(TempCharIndex).AutoCurar = 1 Then
                Call SendData(ToIndex, Userindex, 0, "8Q" & Npclist(TempCharIndex).Name)
            Else
                Call SendNPC(Userindex, TempCharIndex)
            End If
            FoundSomething = 1
            UserList(Userindex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(Userindex).flags.TargetNpc = TempCharIndex
            UserList(Userindex).flags.TargetUser = 0
            UserList(Userindex).flags.TargetObj = 0
    End If
    
    If FoundChar = 0 Then
        UserList(Userindex).flags.TargetNpc = 0
        UserList(Userindex).flags.TargetNpcTipo = 0
        UserList(Userindex).flags.TargetUser = 0
    End If
    
    If FoundSomething = 0 Then
        UserList(Userindex).flags.TargetNpc = 0
        UserList(Userindex).flags.TargetNpcTipo = 0
        UserList(Userindex).flags.TargetUser = 0
        UserList(Userindex).flags.TargetObj = 0
        UserList(Userindex).flags.TargetObjMap = 0
        UserList(Userindex).flags.TargetObjX = 0
        UserList(Userindex).flags.TargetObjY = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(Userindex).flags.TargetNpc = 0
        UserList(Userindex).flags.TargetNpcTipo = 0
        UserList(Userindex).flags.TargetUser = 0
        UserList(Userindex).flags.TargetObj = 0
        UserList(Userindex).flags.TargetObjMap = 0
        UserList(Userindex).flags.TargetObjX = 0
        UserList(Userindex).flags.TargetObjY = 0
    End If
End If

End Sub
Function FindDirection(POS As WorldPos, Target As WorldPos) As Byte
Dim X As Integer, Y As Integer

X = POS.X - Target.X
Y = POS.Y - Target.Y

If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function
Public Function ItemEsDeMapa(ByVal Map As Integer, X As Integer, Y As Integer) As Boolean

ItemEsDeMapa = ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Agarrable Or MapData(Map, X, Y).Blocked

End Function


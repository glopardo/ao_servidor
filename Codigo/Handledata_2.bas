Attribute VB_Name = "Handledata_2"
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

Public Sub HandleData2(UserIndex As Integer, rdata As String, Procesado As Boolean)
Dim LoopC As Integer, tIndex As Integer, N As Integer, X As Integer, Y As Integer, tInt As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tLong As Long

Procesado = True

Select Case Left$(UCase$(rdata), 2)
    Case "#*"
        rdata = Right$(rdata, Len(rdata) - 2)
        tIndex = NameIndex(rdata)
        If tIndex Then
            If UserList(tIndex).flags.Privilegios < 2 Then
                Call SendData(ToIndex, UserIndex, 0, "||El jugador " & UserList(tIndex).Name & " se encuentra online." & FONTTYPE_INFO)
            Else: Call SendData(ToIndex, UserIndex, 0, "1A")
            End If
        Else: Call SendData(ToIndex, UserIndex, 0, "1A")
        End If
        Exit Sub
    Case "#]"
        rdata = Right$(rdata, Len(rdata) - 2)
        Call TirarRuleta(UserIndex, rdata)
    
        Exit Sub
    Case "#}"
        UserList(UserIndex).flags.MesaCasino = 0
        Call SendUserORO(UserIndex)
        Exit Sub
        
    Case "^A"
        rdata = Right$(rdata, Len(rdata) - 2)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & ": " & rdata & FONTTYPE_FIGHT)
        Exit Sub
    
    Case "#$"
        rdata = Right$(rdata, Len(rdata) - 2)
        If UserList(UserIndex).flags.Privilegios < 2 Then Exit Sub
        X = ReadField(1, rdata, 44)
        Y = ReadField(2, rdata, 44)
        N = MapaPorUbicacion(X, Y)
        If N Then Call WarpUserChar(UserIndex, N, 50, 50, True)
        Call LogGM(UserList(UserIndex).Name, "Se transporto por mapa a Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    
    Case "#A"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        If Not UserList(UserIndex).flags.Meditando And UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then Exit Sub
        Call SendData(ToIndex, UserIndex, 0, "MEDOK")
        If Not UserList(UserIndex).flags.Meditando Then
           Call SendData(ToIndex, UserIndex, 0, "7M")
        Else
           Call SendData(ToIndex, UserIndex, 0, "D9")
        End If
        UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
        
        If UserList(UserIndex).flags.Meditando Then
            UserList(UserIndex).Counters.tInicioMeditar = Timer
            Call SendData(ToIndex, UserIndex, 0, "8M" & TIEMPO_INICIOMEDITAR)


            UserList(UserIndex).Char.loops = LoopAdEternum
            If UserList(UserIndex).Stats.ELV < 15 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARCHICO
            ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
            Else
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARGRANDE
            End If
        Else
            UserList(UserIndex).Char.FX = 0
            UserList(UserIndex).Char.loops = 0
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
        End If
        Exit Sub
    Case "#B"
        If UserList(UserIndex).flags.Paralizado Then Exit Sub
        
        If (Not MapInfo(UserList(UserIndex).POS.Map).Pk And TiempoTranscurrido(UserList(UserIndex).Counters.LastRobo) > 10) Or UserList(UserIndex).flags.Privilegios > 1 Then
            Call SendData(ToIndex, UserIndex, 0, "FINOK")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        Call Cerrar_Usuario(UserIndex)
        
        Exit Sub

    Case "#C"
        If CanCreateGuild(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SHOWFUN" & UserList(UserIndex).Faccion.Bando)
        Exit Sub
    
    Case "#D"
        Call SendData(ToIndex, UserIndex, 0, "7L")
        Exit Sub
    
    Case "#E"
        Call SendData(ToIndex, UserIndex, 0, "7L")
        Exit Sub
    
    Case "#F"
        Call SendData(ToIndex, UserIndex, 0, "7L")
        Exit Sub
        

    Case "#G"
        
        If UserList(UserIndex).flags.Muerto Then
                  Call SendData(ToIndex, UserIndex, 0, "MU")
                  Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "ZP")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "DL")
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
        Or UserList(UserIndex).flags.Muerto Then Exit Sub

        Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        Exit Sub
    Case "#H"
         
         If UserList(UserIndex).flags.Muerto Then
                      Call SendData(ToIndex, UserIndex, 0, "MU")
                      Exit Sub
         End If
         
         If UserList(UserIndex).flags.TargetNpc = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "ZP")
                  Exit Sub
         End If
         If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "DL")
                      Exit Sub
         End If
         If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> _
            UserIndex Then Exit Sub
         Npclist(UserList(UserIndex).flags.TargetNpc).Movement = ESTATICO
         Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
         Exit Sub
    Case "#I"
        
        If UserList(UserIndex).flags.Muerto Then
                  Call SendData(ToIndex, UserIndex, 0, "MU")
                  Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "ZP")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "DL")
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> _
          UserIndex Then Exit Sub
        Call FollowAmo(UserList(UserIndex).flags.TargetNpc)
        Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
        Exit Sub
    Case "#J"
        
        If UserList(UserIndex).flags.Muerto Then
                  Call SendData(ToIndex, UserIndex, 0, "MU")
                  Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "ZP")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "DL")
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
        Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNpc)
        Exit Sub
    Case "#K"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        If HayOBJarea(UserList(UserIndex).POS, FOGATA) Then
                Call SendData(ToIndex, UserIndex, 0, "DOK")
                If Not UserList(UserIndex).flags.Descansar Then
                    Call SendData(ToIndex, UserIndex, 0, "3M")
                Else
                    Call SendData(ToIndex, UserIndex, 0, "4M")
                End If
                UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
        Else
                If UserList(UserIndex).flags.Descansar Then
                    Call SendData(ToIndex, UserIndex, 0, "4M")
                    
                    UserList(UserIndex).flags.Descansar = False
                    Call SendData(ToIndex, UserIndex, 0, "DOK")
                    Exit Sub
                End If
                Call SendData(ToIndex, UserIndex, 0, "6M")
        End If
        Exit Sub

    Case "#L"
       
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "ZP")
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_REVIVIR _
       Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
       If Distancia(UserList(UserIndex).POS, Npclist(UserList(UserIndex).flags.TargetNpc).POS) > 10 Then
           Call SendData(ToIndex, UserIndex, 0, "DL")
           Exit Sub
       End If

       Call RevivirUsuarioNPC(UserIndex)
       Call SendData(ToIndex, UserIndex, 0, "RZ")
       Exit Sub
    Case "#M"
       
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "ZP")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_REVIVIR _
       Or UserList(UserIndex).flags.Muerto Then Exit Sub
       If Distancia(UserList(UserIndex).POS, Npclist(UserList(UserIndex).flags.TargetNpc).POS) > 10 Then
           Call SendData(ToIndex, UserIndex, 0, "DL")
           Exit Sub
       End If
       UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
       Call SendUserHP(UserIndex)
       Exit Sub
    Case "#N"
        If UserList(UserIndex).flags.Muerto Then Exit Sub
        Call EnviarSubclase(UserIndex)
        Exit Sub
    Case "#O"
        If PuedeRecompensa(UserIndex) And Not UserList(UserIndex).flags.Muerto Then _
        Call SendData(ToIndex, UserIndex, 0, "RELON" & UserList(UserIndex).Clase & "," & PuedeRecompensa(UserIndex))
    Exit Sub
    Case "#P"
        If UserList(UserIndex).flags.Privilegios > 0 Then
            For LoopC = 1 To LastUser
                If Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.Privilegios <= 1 Then
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            Next
            If Len(tStr) > 0 Then
                tStr = Left$(tStr, Len(tStr) - 2)
                Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
                Call SendData(ToIndex, UserIndex, 0, "4L" & NumNoGMs)
            Else
                Call SendData(ToIndex, UserIndex, 0, "6L")
            End If
        Else
           Call SendData(ToIndex, UserIndex, 0, "||Este comando ya no está disponible. La cantidad de users online está abajo de la pantalla." & FONTTYPE_INFO)
        End If
        Exit Sub
    Case "#Q"
        Call SendUserSTAtsTxt(UserIndex, UserIndex)
        Exit Sub
    Case "#R"
        If UserList(UserIndex).Counters.Pena Then
            Call SendData(ToIndex, UserIndex, 0, "9M" & CalcularTiempoCarcel(UserIndex))
        Else
            Call SendData(ToIndex, UserIndex, 0, "2N")
        End If
        Exit Sub
    Case "#S"
        If UserList(UserIndex).flags.TargetUser Then
            If MapData(UserList(UserList(UserIndex).flags.TargetUser).POS.Map, UserList(UserList(UserIndex).flags.TargetUser).POS.X, UserList(UserList(UserIndex).flags.TargetUser).POS.Y).OBJInfo.OBJIndex > 0 And _
            UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto Then
                Call SendData(ToAdmins, 0, 0, "8T" & UserList(UserIndex).Name & "," & UserList(UserList(UserIndex).flags.TargetUser).Name)
                Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "!!Fuiste echado por mantenerte sobre un item estando muerto.")
                Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "FINOK")
                Call CloseSocket(UserList(UserIndex).flags.TargetUser)
            End If
        End If
        Exit Sub

    Case "#T"
        If entorneo Then
            Dim jugadores As Integer
            jugadores = val(GetVar(App.Path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD"))
            Dim jugador As Integer
            For jugador = 1 To jugadores
                If UCase$(GetVar(App.Path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugador)) = UCase$(UserList(UserIndex).Name) Then Exit Sub
            Next
            Call WriteVar(App.Path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD", jugadores + 1)
            Call WriteVar(App.Path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugadores + 1, UserList(UserIndex).Name)
            Call SendData(ToIndex, UserIndex, 0, "9T")
            Call SendData(ToAdmins, 0, 0, "2U" & UserList(UserIndex).Name)
        End If
        Exit Sub

    Case "#U"
        Dim NpcIndex As Integer
        Dim theading As Byte
        Dim atra1 As Integer
        Dim atra2 As Integer
        Dim atra3 As Integer
        Dim atra4 As Integer
        
        If Not LegalPos(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X - 1, UserList(UserIndex).POS.Y) And _
        Not LegalPos(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X + 1, UserList(UserIndex).POS.Y) And _
        Not LegalPos(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1) And _
        Not LegalPos(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y + 1) Then
            If UserList(UserIndex).flags.Muerto Then
                If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X - 1, UserList(UserIndex).POS.Y).NpcIndex Then
                    atra1 = MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X - 1, UserList(UserIndex).POS.Y).NpcIndex
                    theading = WEST
                    Call MoveNPCChar(atra1, theading)
                End If
                If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X + 1, UserList(UserIndex).POS.Y).NpcIndex Then
                    atra2 = MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X + 1, UserList(UserIndex).POS.Y).NpcIndex
                    theading = EAST
                    Call MoveNPCChar(atra2, theading)
                End If
                If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1).NpcIndex Then
                    atra3 = MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1).NpcIndex
                    theading = NORTH
                    Call MoveNPCChar(atra3, theading)
                End If
                If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y + 1).NpcIndex Then
                    atra4 = MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y + 1).NpcIndex
                    theading = SOUTH
                    Call MoveNPCChar(atra4, theading)
                 End If
            End If
        End If
        Exit Sub
        
    Case "#V"
        
        If UserList(UserIndex).flags.Muerto Then
                  Call SendData(ToIndex, UserIndex, 0, "MU")
                  Exit Sub
        End If
        If UserList(UserIndex).flags.Privilegios = 1 Then
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetNpc Then
              
              If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                 If Len(Npclist(UserList(UserIndex).flags.TargetNpc).Desc) > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "3Q" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                 Exit Sub
              End If
              If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "DL")
                  Exit Sub
              End If
              
              Call IniciarComercioNPC(UserIndex)
         

        ElseIf UserList(UserIndex).flags.TargetUser Then
            
            
            If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "4U")
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.TargetUser = UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "5U")
                Exit Sub
            End If
            
            If Distancia(UserList(UserList(UserIndex).flags.TargetUser).POS, UserList(UserIndex).POS) > 3 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If
            
            If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando And _
                UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "6U")
                Exit Sub
            End If
            
            UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
            UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).Name
            UserList(UserIndex).ComUsu.Cant = 0
            UserList(UserIndex).ComUsu.Objeto = 0
            UserList(UserIndex).ComUsu.Acepto = False
            
            
            Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)

        Else
            Call SendData(ToIndex, UserIndex, 0, "ZP")
        End If
        Exit Sub
    
    
    Case "#W"
        
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetNpc = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "ZP")
            Exit Sub
        End If
        
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 3 Then
            Call SendData(ToIndex, UserIndex, 0, "DL")
            Exit Sub
        End If
        
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
        
        Call IniciarDeposito(UserIndex)
    
        Exit Sub

    Case "#Y"
    
    
        If UserList(UserIndex).flags.TargetNpc = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "ZP")
            Exit Sub
        End If
        
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_NOBLE Or UserList(UserIndex).flags.Muerto Then Exit Sub
       
        If Distancia(UserList(UserIndex).POS, Npclist(UserList(UserIndex).flags.TargetNpc).POS) > 4 Then
            Call SendData(ToIndex, UserIndex, 0, "DL")
            Exit Sub
        End If
       
        If ClaseBase(UserList(UserIndex).Clase) Or ClaseTrabajadora(UserList(UserIndex).Clase) Then Exit Sub
       
        Call Enlistar(UserIndex, Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion)
       
        Exit Sub

    Case "#1"
        
        If UserList(UserIndex).flags.TargetNpc = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "ZP")
            Exit Sub
        End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_NOBLE Or UserList(UserIndex).flags.Muerto Or Not Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion Then Exit Sub
        If Distancia(UserList(UserIndex).POS, Npclist(UserList(UserIndex).flags.TargetNpc).POS) > 4 Then
            Call SendData(ToIndex, UserIndex, 0, "DL")
            Exit Sub
        End If

        If UserList(UserIndex).Faccion.Bando <> Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion Then
            Call SendData(ToIndex, UserIndex, 0, Mensajes(Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion, 16) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            Exit Sub
        End If
        Call Recompensado(UserIndex)
        Exit Sub
        
    Case "#5"
        rdata = Right$(rdata, Len(rdata) - 3)
        
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "M4")
            Exit Sub
        End If
        
        If Not AsciiValidos(rdata) Then
            Call SendData(ToIndex, UserIndex, 0, "7U")
            Exit Sub
        End If
        
        If Len(rdata) > 80 Then
            Call SendData(ToIndex, UserIndex, 0, "||La descripción debe tener menos de 80 cáracteres de largo." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).Desc = rdata
        Call SendData(ToIndex, UserIndex, 0, "8U")
        Exit Sub
        
    Case "#6 "
        rdata = Right$(rdata, Len(rdata) - 3)
        Call ComputeVote(UserIndex, rdata)
        Exit Sub
            
    Case "#7"
        Call SendData(ToIndex, UserIndex, 0, "||Este comando ya no anda, para hablar por tu clan presiona la tecla 3 y habla normalmente." & FONTTYPE_INFO)
        Exit Sub

    Case "#8"
        Call SendData(ToIndex, UserIndex, 0, "||Este comando ya no se usa, pon /PASSWORD para cambiar tu password." & FONTTYPE_INFO)
        Exit Sub
        
    Case "#!"
        If PuedeFaccion(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "4&")
        Exit Sub
        
    Case "#9"
        rdata = Right$(rdata, Len(rdata) - 3)
        tLong = CLng(val(rdata))
        If tLong > 32000 Then tLong = 32000
        N = tLong
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
        ElseIf UserList(UserIndex).flags.TargetNpc = 0 Then
            
            Call SendData(ToIndex, UserIndex, 0, "ZP")
        ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 10 Then
            Call SendData(ToIndex, UserIndex, 0, "DL")
        ElseIf Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_APOSTADOR Then
            Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        ElseIf N < 1 Then
            Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        ElseIf N > 5000 Then
            Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        ElseIf UserList(UserIndex).Stats.GLD < N Then
            Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Else
            If RandomNumber(1, 100) <= 47 Then
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + N
                Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                
                Apuestas.Ganancias = Apuestas.Ganancias + N
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            Else
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - N
                Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            
                Apuestas.Perdidas = Apuestas.Perdidas + N
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            End If
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call SendUserORO(UserIndex)
        End If
        Exit Sub
        
    Case "#/"
        rdata = Right$(rdata, Len(rdata) - 3)
        tIndex = NameIndex(ReadField(1, rdata, 32))
        If tIndex = 0 Then Exit Sub
        If ReadField(2, rdata, 32) = "0" Then
            Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " te ha dejado de ignorar." & FONTTYPE_INFO)
        Else: Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " te empezó a ignorar." & FONTTYPE_INFO)
        End If
        Exit Sub
        
    Case "#0"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
         
        If UserList(UserIndex).flags.TargetNpc = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "ZP")
            Exit Sub
        End If
         
        If UserList(UserIndex).flags.Muerto Then Exit Sub
         
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
         
        If Distancia(UserList(UserIndex).POS, Npclist(UserList(UserIndex).flags.TargetNpc).POS) > 10 Then
            Call SendData(ToIndex, UserIndex, 0, "DL")
            Exit Sub
        End If
         
        rdata = Right$(rdata, Len(rdata) - 3)
        
        If val(rdata) > 0 Then
            If val(rdata) > UserList(UserIndex).Stats.Banco Then rdata = UserList(UserIndex).Stats.Banco
            UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rdata)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rdata)
            Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
        End If
         
        Call SendUserORO(UserIndex)
         
        Exit Sub

    Case "#Ñ"
        
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        If UserList(UserIndex).flags.TargetNpc = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "ZP")
            Exit Sub
        End If
        
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 10 Then
            Call SendData(ToIndex, UserIndex, 0, "DL")
            Exit Sub
        End If
        
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Or UserList(UserIndex).flags.Muerto Then Exit Sub
        
        If Distancia(UserList(UserIndex).POS, Npclist(UserList(UserIndex).flags.TargetNpc).POS) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "DL")
              Exit Sub
        End If
        
        rdata = Right$(rdata, Len(rdata) - 3)
        
        If CLng(val(rdata)) > 0 Then
            If CLng(val(rdata)) > UserList(UserIndex).Stats.GLD Then rdata = UserList(UserIndex).Stats.GLD
            UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rdata)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
            Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
        End If
    
        Call SendUserORO(UserIndex)
        
        Exit Sub
        
    Case "#2"
        If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then
            If UserList(UserIndex).GuildInfo.EsGuildLeader And UserList(UserIndex).flags.Privilegios < 2 Then
                Call SendData(ToIndex, UserIndex, 0, "4V")
                Exit Sub
            End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "5V")
            Exit Sub
        End If
        
        Call SendData(ToGuildMembers, UserIndex, 0, "6V" & UserList(UserIndex).Name)
        Call SendData(ToIndex, UserIndex, 0, "7V")
        
        Dim oGuild As cGuild
        
        Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
        
        If oGuild Is Nothing Then Exit Sub
        
        For i = 1 To LastUser
            If UserList(i).GuildInfo.GuildName = oGuild.GuildName Then UserList(i).flags.InfoClanEstatica = 0
        Next
        
        UserList(UserIndex).GuildInfo.GuildPoints = 0
        UserList(UserIndex).GuildInfo.GuildName = ""
        Call oGuild.RemoveMember(UserList(UserIndex).Name)
        
        Call UpdateUserChar(UserIndex)
        
        Exit Sub
      
    Case "#4"

        If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "ZP")
           Exit Sub
       End If
       
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_NOBLE Or UserList(UserIndex).flags.Muerto Or Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then Exit Sub
        
        If Distancia(UserList(UserIndex).POS, Npclist(UserList(UserIndex).flags.TargetNpc).POS) > 4 Then
            Call SendData(ToIndex, UserIndex, 0, "DL")
            Exit Sub
        End If
        
        If UserList(UserIndex).Faccion.Bando <> Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion Then Exit Sub
        
        If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then
            Call SendData(ToIndex, UserIndex, 0, Mensajes(UserList(UserIndex).Faccion.Bando, 23) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            Exit Sub
        End If
        
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion, 18) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

        UserList(UserIndex).Faccion.Bando = Neutral
        UserList(UserIndex).Faccion.Jerarquia = 0
        Call UpdateUserChar(UserIndex)
Exit Sub

Case "#3"
    If Len(UserList(UserIndex).GuildInfo.GuildName) = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "5V")
        Exit Sub
    End If
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName Then
            tStr = tStr & UserList(LoopC).Name & ", "
        End If
    Next
    
    If Len(tStr) > 0 Then
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||Miembros de tu clan online:" & tStr & "." & FONTTYPE_GUILD)
    Else: Call SendData(ToIndex, UserIndex, 0, "8V")
    End If
    Exit Sub
    
    End Select

    Procesado = False
End Sub

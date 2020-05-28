Attribute VB_Name = "Handledata_1"
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
Public Sub HandleData1(Userindex As Integer, rdata As String, Procesado As Boolean)
Dim tInt As Integer, tIndex As Integer, X As Integer, Y As Integer
Dim Arg1 As String, Arg2 As String, arg3 As String
Dim nPos As WorldPos
Dim tLong As Long
Dim ind

Procesado = True

Select Case UCase$(Left$(rdata, 1))
    Case "\"
        If UserList(Userindex).flags.Muerto = 1 Then
            Call SendData(ToIndex, Userindex, 0, "MU")
            Exit Sub
        End If
        
        rdata = Right$(rdata, Len(rdata) - 1)
        tName = ReadField(1, rdata, 32)
        tIndex = NameIndex(tName)
        
        If tIndex <> 0 Then
            If UserList(tIndex).flags.Muerto = 1 Then Exit Sub
    
            If Len(rdata) <> Len(tName) Then
                tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
            Else
                tMessage = " "
            End If
             
            If Not EnPantalla(UserList(Userindex).POS, UserList(tIndex).POS, 1) Then
                Call SendData(ToIndex, Userindex, 0, "2E")
                Exit Sub
            End If
             
            ind = UserList(Userindex).Char.CharIndex
             
            If InStr(tMessage, "°") Then Exit Sub
    
            If UserList(tIndex).flags.Privilegios > 0 And UserList(Userindex).flags.Privilegios = 0 Then
                Call SendData(ToIndex, Userindex, 0, "3E")
                Exit Sub
            End If
    
            Call SendData(ToIndex, Userindex, UserList(Userindex).POS.Map, "||" & vbCyan & "°" & tMessage & "°" & str(ind))
            Call SendData(ToIndex, tIndex, UserList(Userindex).POS.Map, "||" & vbCyan & "°" & tMessage & "°" & str(ind))
            Call SendData(ToGMArea, Userindex, UserList(Userindex).POS.Map, "||" & vbCyan & "°" & tMessage & "°" & str(ind))
            Exit Sub
        End If
        
        Call SendData(ToIndex, Userindex, 0, "3E")
        Exit Sub
            
    Case ";"
        Dim Modo As String
        
        rdata = Right$(rdata, Len(rdata) - 1)
        If Right$(rdata, Len(rdata) - 1) = " " Or Right$(rdata, Len(rdata) - 1) = "-" Then rdata = "1 "
        If Len(rdata) = 1 Then Exit Sub
        
        'If val(Right$(rdata, Len(rdata) - 1)) > 0 Then
        '        UserList(UserIndex).flags.IntentosCodigo = UserList(UserIndex).flags.IntentosCodigo + 1
        '        If UserList(UserIndex).flags.IntentosCodigo >= 10 Then
        '            UserList(UserIndex).flags.CodigoTrabajo = 0
        '            UserList(UserIndex).flags.IntentosCodigo = 0
        '            Call SacarModoTrabajo(UserIndex)
        '            Call SendData(ToIndex, UserIndex, 0, "||Fuiste encarcelado por mandar demasiados códigos de trabajo incorrectos." & FONTTYPE_FIGHT)
        '            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " fue encarcelado por mandar demasiados códigos incorrectos." & fonttype_Server)
        '            Call Encarcelar(UserIndex, 15)
        '            Exit Sub
        '        Else: Call SendData(ToIndex, UserIndex, 0, "||Código incorrecto. Te quedan " & 10 - UserList(UserIndex).flags.IntentosCodigo & " intentos o serás encarcelado." & FONTTYPE_INFO)
        '        End If
        '    End If
       
        
        Modo = Left$(rdata, 1)
        rdata = Replace(Right$(rdata, Len(rdata) - 1), "~", "-")
        
    Select Case Modo
        'Color dialogos
        Case 1, 4, 5
            
            If InStr(rdata, "°") Then Exit Sub
            
            If (Modo = 4 Or Modo = 5) And UserList(Userindex).flags.Muerto = 1 Then
                Call SendData(ToIndex, Userindex, 0, "MU")
                Exit Sub
            End If
            
            If UserList(Userindex).flags.Privilegios = 1 Then Call LogGM(UserList(Userindex).Name, "Dijo: " & rdata, True)
            If InStr(1, rdata, Chr$(255)) Then rdata = Replace(rdata, Chr$(255), " ")
            
            ind = UserList(Userindex).Char.CharIndex
            Dim Color As Long
            Dim IndexSendData As Byte
            
            If Modo = 4 Then 'Grito
                Color = vbRed
            ElseIf Modo = 5 Then '*Rol*
                Color = vbGreen
            ElseIf UserList(Userindex).flags.Privilegios Then 'GM
                Color = &H80FF&
            ElseIf UserList(Userindex).flags.EsConseCaos Then 'Concilio
                Color = &H8080FF
            ElseIf UserList(Userindex).flags.EsConseReal Then 'Consejo
                Color = &HC0C000
            ElseIf UserList(Userindex).flags.Quest And UserList(Userindex).Faccion.Bando <> Neutral Then
                If UserList(Userindex).Faccion.Bando = Real Then 'Quest ciuda
                    Color = vbBlue
                Else: Color = vbRed 'Quest crimi
                End If
            ElseIf UserList(Userindex).flags.Muerto Then 'Muerto
                Color = vbYellow
            Else: Color = vbWhite 'Normal
            End If
    
            If UserList(Userindex).flags.Privilegios > 0 Or UserList(Userindex).Clase = CLERIGO Then
                IndexSendData = ToPCArea
            ElseIf UserList(Userindex).flags.Muerto Then
                IndexSendData = ToMuertos
            Else
                IndexSendData = ToPCAreaVivos
            End If
            
            If UCase$(rdata) = "SACRIFICATE!" Then
                nPos = UserList(Userindex).POS
                Call HeadtoPos(UserList(Userindex).Char.Heading, nPos)
                tIndex = MapData(nPos.Map, nPos.X, nPos.Y).Userindex
                If tIndex > 0 Then
                    If MapData(nPos.Map, nPos.X - 1, nPos.Y).OBJInfo.OBJIndex = Cruz And _
                    MapData(nPos.Map, nPos.X + 1, nPos.Y).OBJInfo.OBJIndex = Cruz And _
                    MapData(nPos.Map, nPos.X, nPos.Y - 1).OBJInfo.OBJIndex = Cruz And _
                    MapData(nPos.Map, nPos.X, nPos.Y + 1).OBJInfo.OBJIndex = Cruz Then
                        If UserList(Userindex).Stats.ELV < 40 Then
                            Call SendData(ToIndex, Userindex, 0, "||Debes ser nivel 40 o más para iniciar un sacrificio." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If UserList(tIndex).Stats.MinHP < UserList(tIndex).Stats.MaxHP / 2 Then
                            Call SendData(ToIndex, Userindex, 0, "||Solo puedes comenzar a sacrificar a usuarios que tengan más de la mitad de sus HP." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        UserList(tIndex).flags.Sacrificando = 1
                        UserList(tIndex).flags.Sacrificador = Userindex
                        UserList(Userindex).flags.Sacrificado = tIndex
                        Call SendData(ToIndex, Userindex, 0, "||¡Comenzaste a sacrificar a " & UserList(tIndex).Name & "!" & FONTTYPE_INFO)
                        Call SendData(ToIndex, tIndex, 0, "||¡" & UserList(Userindex).Name & " comenzó a sacrificarte! ¡Huye!" & FONTTYPE_INFO)
                    End If
                End If
            End If
            
            If Modo = 5 Then rdata = "* " & rdata & " *"

            Call SendData(IndexSendData, Userindex, UserList(Userindex).POS.Map, "||" & Color & "°" & rdata & "°" & str(ind))
            Exit Sub
            
        Case 2
            
            If UserList(Userindex).flags.Muerto Then
                Call SendData(ToIndex, Userindex, 0, "MU")
                Exit Sub
            End If
            
            tIndex = UserList(Userindex).flags.Whispereando
            
            If tIndex Then
                If UserList(tIndex).flags.Muerto Then Exit Sub
    
                If Not EnPantalla(UserList(Userindex).POS, UserList(tIndex).POS, 1) Then
                    Call SendData(ToIndex, Userindex, 0, "2E")
                    Exit Sub
                End If
                
                ind = UserList(Userindex).Char.CharIndex
                
                If InStr(rdata, "°") Then Exit Sub

                If UserList(tIndex).flags.Privilegios > 0 And UserList(tIndex).flags.AdminInvisible Then
                    Call SendData(ToIndex, Userindex, 0, "3E")
                    Call SendData(ToIndex, tIndex, UserList(Userindex).POS.Map, "||" & vbBlue & "°" & rdata & "°" & str(ind))
                    Exit Sub
                End If
                
                If UserList(Userindex).flags.Privilegios = 1 Then Call LogGM(UserList(Userindex).Name, "Grito: " & rdata, True)
                
                If EnPantalla(UserList(Userindex).POS, UserList(tIndex).POS, 1) Then
                    Call SendData(ToIndex, Userindex, 0, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                    Call SendData(ToIndex, tIndex, 0, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                    Call SendData(ToGMArea, Userindex, UserList(Userindex).POS.Map, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                Else
                    Call SendData(ToIndex, Userindex, 0, "{F")
                    UserList(Userindex).flags.Whispereando = 0
                End If
            End If
            
            Exit Sub
        
        Case 3
            If UserList(Userindex).flags.Muerto Then
                Call SendData(ToIndex, Userindex, 0, "MU")
                Exit Sub
            End If
        
            If Len(rdata) And Len(UserList(Userindex).GuildInfo.GuildName) > 0 Then Call SendData(ToGuildMembers, Userindex, 0, "||" & UserList(Userindex).Name & "> " & rdata & FONTTYPE_GUILD)
            Exit Sub
            
        Case 6
            If UserList(Userindex).flags.Party = 0 Then Exit Sub
            
            If Len(rdata) > 0 Then
                Call SendData(ToParty, Userindex, 0, "||" & UserList(Userindex).Name & ": " & rdata & FONTTYPE_PARTY)
            End If
            Exit Sub
                
        Case 7
            If UserList(Userindex).flags.Privilegios = 0 Then Exit Sub
            
            Call LogGM(UserList(Userindex).Name, "Mensaje a Gms:" & rdata, (UserList(Userindex).flags.Privilegios = 1))
            If Len(rdata) > 0 Then
                Call SendData(ToAdmins, 0, 0, "||" & UserList(Userindex).Name & "> " & rdata & "~255~255~255~0~1")
            End If
            
            Exit Sub
    
        End Select
        
    Case "M"
        Dim Mide As Double
        rdata = Right$(rdata, Len(rdata) - 1)

        If UserList(Userindex).flags.Trabajando Then

                Call SacarModoTrabajo(Userindex)

        End If
        
        If Not UserList(Userindex).flags.Descansar And Not UserList(Userindex).flags.Meditando _
           And UserList(Userindex).flags.Paralizado = 0 Then
            Call MoveUserChar(Userindex, val(rdata))
        ElseIf UserList(Userindex).flags.Descansar Then
            UserList(Userindex).flags.Descansar = False
            Call SendData(ToIndex, Userindex, 0, "DOK")
            Call SendData(ToIndex, Userindex, 0, "DN")
            Call MoveUserChar(Userindex, val(rdata))
        End If

        If UserList(Userindex).flags.Oculto Then
            If Not (UserList(Userindex).Clase = LADRON And UserList(Userindex).Recompensas(2) = 1) Then
                UserList(Userindex).flags.Oculto = 0
                UserList(Userindex).flags.Invisible = 0
                Call SendData(ToMap, 0, UserList(Userindex).POS.Map, ("V3" & UserList(Userindex).Char.CharIndex & ",0"))
                Call SendData(ToIndex, Userindex, 0, "V5")
            End If
        End If

        Exit Sub
End Select

Select Case UCase$(Left$(rdata, 2))
    Case "ZI"
        rdata = Right$(rdata, Len(rdata) - 2)
        Dim Bait(1 To 2) As Byte
        Bait(1) = val(ReadField(1, rdata, 44))
        Bait(2) = val(ReadField(2, rdata, 44))
        
        Select Case Bait(2)
            Case 0
                Bait(2) = Bait(1) - 1
            Case 1
                Bait(2) = Bait(1) + 1
            Case 2
                Bait(2) = Bait(1) - 5
            Case 3
                Bait(2) = Bait(1) + 5
        End Select
        
        If Bait(2) > 0 And Bait(2) <= MAX_INVENTORY_SLOTS Then Call AcomodarItems(Userindex, Bait(1), Bait(2))
        
        Exit Sub
    Case "TI"
        If UserList(Userindex).flags.Navegando = 1 Or _
           UserList(Userindex).flags.Muerto = 1 Or _
                          UserList(Userindex).flags.Montado Then Exit Sub
           
        
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If val(Arg1) = FLAGORO Then
            Call TirarOro(val(Arg2), Userindex)
            Call SendUserORO(Userindex)
            Exit Sub
        Else
            If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) Then
                If UserList(Userindex).Invent.Object(val(Arg1)).OBJIndex = 0 Then
                        Exit Sub
                End If
                Call DropObj(Userindex, val(Arg1), val(Arg2), UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y)
            Else
                Exit Sub
            End If
        End If
        Exit Sub
    Case "SF"
        rdata = Right$(rdata, Len(rdata) - 2)
        If Not PuedeFaccion(Userindex) Then Exit Sub
        If UserList(Userindex).Faccion.BandoOriginal Then Exit Sub
        tInt = val(rdata)
        
        If tInt = Neutral Then
            If UserList(Userindex).Faccion.Bando <> Neutral Then
                Call SendData(ToIndex, Userindex, 0, "7&")
            Else: Call SendData(ToIndex, Userindex, 0, "0&")
            End If
            Exit Sub
        End If
        
        If UserList(Userindex).Faccion.Matados(tInt) > UserList(Userindex).Faccion.Matados(Enemigo(tInt)) Then
            Call SendData(ToIndex, Userindex, 0, Mensajes(tInt, 9))
            Exit Sub
        End If
        
        Call SendData(ToIndex, Userindex, 0, Mensajes(tInt, 10))
        UserList(Userindex).Faccion.BandoOriginal = tInt
        UserList(Userindex).Faccion.Bando = tInt
        UserList(Userindex).Faccion.Ataco(tInt) = 0
        If Not PuedeFaccion(Userindex) Then Call SendData(ToIndex, Userindex, 0, "SUFA0")
        
        Call UpdateUserChar(Userindex)
        
        Exit Sub
    Case "LH"
        If UserList(Userindex).flags.Muerto Then
            Call SendData(ToIndex, Userindex, 0, "MU")
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 2)
        UserList(Userindex).flags.Hechizo = val(rdata)
        Exit Sub
    Case "WH"
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        If Not InMapBounds(X, Y) Then Exit Sub
        Call LookatTile(Userindex, UserList(Userindex).POS.Map, X, Y)
        
        If UserList(Userindex).flags.TargetUser = Userindex Then
            Call SendData(ToIndex, Userindex, 0, "{C")
            Exit Sub
        End If
        
        If UserList(Userindex).flags.TargetUser Then
            UserList(Userindex).flags.Whispereando = UserList(Userindex).flags.TargetUser
            Call SendData(ToIndex, Userindex, 0, "{B" & UserList(UserList(Userindex).flags.Whispereando).Name)
        Else
            Call SendData(ToIndex, Userindex, 0, "{D")
        End If
        
        Exit Sub
    Case "LC"
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        Dim POS As WorldPos
        POS.Map = UserList(Userindex).POS.Map
        POS.X = CInt(Arg1)
        POS.Y = CInt(Arg2)
        If Not EnPantalla(UserList(Userindex).POS, POS, 1) Then Exit Sub
        Call LookatTile(Userindex, UserList(Userindex).POS.Map, POS.X, POS.Y)
        Exit Sub
    Case "RC"
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        Call Accion(Userindex, UserList(Userindex).POS.Map, X, Y)
        Exit Sub
    Case "UK"
        If UserList(Userindex).flags.Muerto Then
            Call SendData(ToIndex, Userindex, 0, "MU")
            Exit Sub
        End If

        rdata = Right$(rdata, Len(rdata) - 2)
        Select Case val(rdata)
            Case Robar
                Call SendData(ToIndex, Userindex, 0, "T01" & Robar)
            Case Magia
                Call SendData(ToIndex, Userindex, 0, "T01" & Magia)
            Case Domar
                Call SendData(ToIndex, Userindex, 0, "T01" & Domar)
            Case Invitar
                Call SendData(ToIndex, Userindex, 0, "T01" & Invitar)
                
            Case Ocultarse
                
                If UserList(Userindex).flags.Navegando Then
                      Call SendData(ToIndex, Userindex, 0, "6E")
                      Exit Sub
                End If
                
                If UserList(Userindex).flags.Oculto Then
                      Call SendData(ToIndex, Userindex, 0, "7E")
                      Exit Sub
                End If
                
                Call DoOcultarse(Userindex)
        End Select
        Exit Sub
End Select

Select Case UCase$(rdata)
    Case "RPU"
        Call SendData(ToIndex, Userindex, 0, "PU" & UserList(Userindex).POS.X & "," & UserList(Userindex).POS.Y)
        Exit Sub
    Case "AT"
        If UserList(Userindex).flags.Muerto Then
            Call SendData(ToIndex, Userindex, 0, "MU")
            Exit Sub
        End If
        If UserList(Userindex).Invent.WeaponEqpObjIndex Then
            If ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).proyectil Or ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Baculo Then
                Call SendData(ToIndex, Userindex, 0, "||No puedes usar así esta arma." & FONTTYPE_FIGHT)
                Exit Sub
            End If
        End If
        
        Call UsuarioAtaca(Userindex)
        
        Exit Sub
    Case "AG"
        If UserList(Userindex).flags.Muerto Then
                Call SendData(ToIndex, Userindex, 0, "MU")
                Exit Sub
        End If
        
   
   
   
   
        Call GetObj(Userindex)
        Exit Sub
    Case "SEG"
        If UserList(Userindex).flags.Seguro Then
              Call SendData(ToIndex, Userindex, 0, "1O")
        Else
              Call SendData(ToIndex, Userindex, 0, "9K")
        End If
        UserList(Userindex).flags.Seguro = Not UserList(Userindex).flags.Seguro
        Exit Sub
    Case "ATRI"
        Call EnviarAtrib(Userindex)
        Exit Sub
    Case "FAMA"
        Call EnviarFama(Userindex)
        Call EnviarMiniSt(Userindex)
        Exit Sub
    Case "ESKI"
        Call EnviarSkills(Userindex)
        Exit Sub
    Case "PARSAL"
        Dim i As Integer
        If UserList(Userindex).flags.Party Then
            If Party(UserList(Userindex).PartyIndex).NroMiembros = 2 Then
                Call RomperParty(Userindex)
            Else: Call SacarDelParty(Userindex)
            End If
        Else
            Call SendData(ToIndex, Userindex, 0, "||No estás en party." & FONTTYPE_PARTY)
        End If
        Exit Sub
    Case "PARINF"
        Call EnviarIntegrantesParty(Userindex)
        Exit Sub
    
    Case "FINCOM"
        
        UserList(Userindex).flags.Comerciando = False
        Call SendData(ToIndex, Userindex, 0, "FINCOMOK")
        Exit Sub
    Case "FINCOMUSU"
        If UserList(Userindex).ComUsu.DestUsu > 0 Then
            If UserList(UserList(Userindex).ComUsu.DestUsu).ComUsu.DestUsu = Userindex Then
                Call SendData(ToIndex, UserList(Userindex).ComUsu.DestUsu, 0, "6R" & UserList(Userindex).Name)
                Call FinComerciarUsu(UserList(Userindex).ComUsu.DestUsu)
            End If
        End If
        
        Call FinComerciarUsu(Userindex)
        Exit Sub

    Case "FINBAN"
        UserList(Userindex).flags.Comerciando = False
        Call SendData(ToIndex, Userindex, 0, "FINBANOK")
        Exit Sub
        
    Case "FINTIE"
        UserList(Userindex).flags.Comerciando = False
        Call SendData(ToIndex, Userindex, 0, "FINTIEOK")
        Exit Sub

    Case "COMUSUOK"
        
        Call AceptarComercioUsu(Userindex)
        Exit Sub
    Case "COMUSUNO"
        
        If UserList(Userindex).ComUsu.DestUsu Then
            Call SendData(ToIndex, UserList(Userindex).ComUsu.DestUsu, 0, "7R" & UserList(Userindex).Name)
            Call FinComerciarUsu(UserList(Userindex).ComUsu.DestUsu)
        End If
        Call SendData(ToIndex, Userindex, 0, "8R")
        Call FinComerciarUsu(Userindex)
        Exit Sub
    Case "GLINFO"
        If UserList(Userindex).GuildInfo.EsGuildLeader Then
            If UserList(Userindex).flags.InfoClanEstatica Then
                Call SendData(ToIndex, Userindex, 0, "GINFIG")
            Else
                Call SendGuildLeaderInfo(Userindex)
            End If
        ElseIf Len(UserList(Userindex).GuildInfo.GuildName) > 0 Then
            If UserList(Userindex).flags.InfoClanEstatica Then
                Call SendData(ToIndex, Userindex, 0, "GINFII")
            Else
                Call SendGuildsStats(Userindex)
            End If
        Else
            If UserList(Userindex).flags.InfoClanEstatica Then
                Call SendData(ToIndex, Userindex, 0, "GINFIJ")
            Else: Call SendGuildsList(Userindex)
            End If
        End If
        
        Exit Sub

End Select

 Select Case UCase$(Left$(rdata, 2))
    Case "(A"
        If PuedeDestrabarse(Userindex) Then
            Call ClosestLegalPos(UserList(Userindex).POS, nPos)
            If InMapBounds(nPos.X, nPos.Y) Then Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, True)
        End If
        
        Exit Sub
    Case "GM"
        rdata = Right$(rdata, Len(rdata) - 2)
        Dim GMDia As String
        Dim GMMapa As String
        Dim GMPJ As String
        Dim GMMail As String
        Dim GMGM As String
        Dim GMTitulo As String
        Dim GMMensaje As String
        
        GMDia = Format(Now, "yyyy-mm-dd hh:mm:ss")
        GMMapa = UserList(Userindex).POS.Map & " - " & UserList(Userindex).POS.X & " - " & UserList(Userindex).POS.Y
        GMPJ = UserList(Userindex).Name
        GMMail = UserList(Userindex).Email
        GMGM = ReadField(1, rdata, 172)
        GMTitulo = ReadField(2, rdata, 172)
        GMMensaje = ReadField(3, rdata, 172)
        
        Con.Execute "INSERT INTO reclamos(fecha,nombre,personaje,email,servidor,gm,asunto,mensaje,respondido,censura,old,respondidopor,respondidoel,respuesta) values(""" & GMDia & """,""" & GMMapa & """,""" & GMPJ & """,""" & GMMail & """, 1,""" & GMGM & """, """ & GMTitulo & """, """ & GMMensaje & """,0,0,0,0,0,0)"
          
        Call SendData(ToAdmins, 0, 9, "3B" & GMTitulo & "," & GMPJ)
  
        Exit Sub
        
    End Select
        
 Select Case UCase$(Left$(rdata, 3))
    Case "FRF"
        rdata = Right$(rdata, Len(rdata) - 3)
        For i = 1 To 10
            If UserList(Userindex).flags.Espiado(i) > 0 Then
                If UserList(UserList(Userindex).flags.Espiado(i)).flags.Privilegios > 1 Then Call SendData(ToIndex, UserList(Userindex).flags.Espiado(i), 0, "{{" & UserList(Userindex).Name & "," & rdata)
            End If
        Next
        Exit Sub
    Case "USA"
        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) Then
            If UserList(Userindex).Invent.Object(val(rdata)).OBJIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call UseInvItem(Userindex, val(rdata), 0)
        Exit Sub
    Case "USE"
        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) Then
            If UserList(Userindex).Invent.Object(val(rdata)).OBJIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call UseInvItem(Userindex, val(rdata), 1)
        Exit Sub
    Case "CNS"
        Dim Arg5 As Integer
        rdata = Right$(rdata, Len(rdata) - 3)
        
        X = CInt(ReadField(1, rdata, 32))
        Arg5 = CInt(ReadField(2, rdata, 32))
        If Arg5 < 1 Then Exit Sub
        If X < 1 Then Exit Sub
        If ObjData(X).SkHerreria = 0 Then Exit Sub
        Call HerreroConstruirItem(Userindex, X, val(Arg5))
        Exit Sub
        
    Case "CNC"
        rdata = Right$(rdata, Len(rdata) - 3)
        
        X = CInt(ReadField(1, rdata, 32))
        Arg1 = CInt(ReadField(2, rdata, 32))
        If Arg1 < 1 Then Exit Sub
        If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Sub
        Call CarpinteroConstruirItem(Userindex, X, val(Arg1))
        Exit Sub
    Case "SCR"
        rdata = Right$(rdata, Len(rdata) - 3)
        
        X = CInt(ReadField(1, rdata, 32))
        Arg1 = CInt(ReadField(2, rdata, 32))
        If X < 1 Or ObjData(X).SkSastreria = 0 Then Exit Sub
        Call SastreConstruirItem(Userindex, X, val(Arg1))
        Exit Sub
    
    Case "WLC"
        rdata = Right$(rdata, Len(rdata) - 3)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        arg3 = ReadField(3, rdata, 44)
        If Len(arg3) = 0 Or Len(Arg2) = 0 Or Len(Arg1) = 0 Then Exit Sub
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(arg3) Then Exit Sub
        
        POS.Map = UserList(Userindex).POS.Map
        POS.X = CInt(Arg1)
        POS.Y = CInt(Arg2)
        tLong = CInt(arg3)
        
        If UserList(Userindex).flags.Muerto = 1 Or _
           UserList(Userindex).flags.Descansar Or _
           UserList(Userindex).flags.Meditando Or _
           Not InMapBounds(POS.X, POS.Y) Then Exit Sub
        
        If Not EnPantalla(UserList(Userindex).POS, POS, 1) Then
            Call SendData(ToIndex, Userindex, 0, "PU" & UserList(Userindex).POS.X & "," & UserList(Userindex).POS.Y)
            Exit Sub
        End If
        
        Select Case tLong
        
        Case Proyectiles
            Dim TU As Integer, tN As Integer
            
            If UserList(Userindex).Invent.WeaponEqpObjIndex = 0 Or _
            UserList(Userindex).Invent.MunicionEqpObjIndex = 0 Then Exit Sub
            
            If UserList(Userindex).Invent.WeaponEqpSlot < 1 Or UserList(Userindex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Or _
            UserList(Userindex).Invent.MunicionEqpSlot < 1 Or UserList(Userindex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Or _
            ObjData(UserList(Userindex).Invent.MunicionEqpObjIndex).ObjType <> OBJTYPE_FLECHAS Or _
            UserList(Userindex).Invent.Object(UserList(Userindex).Invent.MunicionEqpSlot).Amount < 1 Or _
            ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then Exit Sub
            
            If TiempoTranscurrido(UserList(Userindex).Counters.LastFlecha) < IntervaloUserFlechas Then Exit Sub
            If TiempoTranscurrido(UserList(Userindex).Counters.LastHechizo) < IntervaloUserPuedeHechiGolpe Then Exit Sub
            If TiempoTranscurrido(UserList(Userindex).Counters.LastGolpe) < IntervaloUserPuedeAtacar Then Exit Sub
            
            UserList(Userindex).Counters.LastFlecha = Timer
            Call SendData(ToIndex, Userindex, 0, "LF")
            
            If UserList(Userindex).Stats.MinSta >= 10 Then
                 Call QuitarSta(Userindex, RandomNumber(1, 10))
            Else
                 Call SendData(ToIndex, Userindex, 0, "9E")
                 Exit Sub
            End If
             
            Call LookatTile(Userindex, UserList(Userindex).POS.Map, val(Arg1), val(Arg2))
            
            TU = UserList(Userindex).flags.TargetUser
            tN = UserList(Userindex).flags.TargetNpc
                            
            If TU = Userindex Then
                Call SendData(ToIndex, Userindex, 0, "3N")
                Exit Sub
            End If

            Call QuitarUnItem(Userindex, UserList(Userindex).Invent.MunicionEqpSlot)
            
            If UserList(Userindex).Invent.MunicionEqpSlot Then
                UserList(Userindex).Invent.Object(UserList(Userindex).Invent.MunicionEqpSlot).Equipped = 1
                Call UpdateUserInv(False, Userindex, UserList(Userindex).Invent.MunicionEqpSlot)
            End If
            
            If tN Then
                If Npclist(tN).Attackable Then Call UsuarioAtacaNpc(Userindex, tN)
            ElseIf TU Then
                If TU <> Userindex Then
                    Call UsuarioAtacaUsuario(Userindex, TU)
                    SendUserHP TU
                End If
            Else
                Call SendData(ToIndex, Userindex, 0, "||Target invalido." & FONTTYPE_INFO)
            End If
            
            
                
                
                
                
        Case Invitar
            Call LookatTile(Userindex, UserList(Userindex).POS.Map, POS.X, POS.Y)
            
            If UserList(Userindex).flags.TargetUser = 0 Then
                Call SendData(ToIndex, Userindex, 0, "||No hay nadie a quien invitar." & FONTTYPE_PARTY)
                Exit Sub
            End If
            
            If UserList(Userindex).flags.Privilegios > 0 Or UserList(UserList(Userindex).flags.TargetUser).flags.Privilegios > 0 Then Exit Sub

            Call DoInvitar(Userindex, UserList(Userindex).flags.TargetUser)
            
        Case Magia

            
            If UserList(Userindex).flags.Privilegios = 1 Then Exit Sub
            
            Call LookatTile(Userindex, UserList(Userindex).POS.Map, POS.X, POS.Y)
            
            If UserList(Userindex).flags.Hechizo Then
                Call LanzarHechizo(UserList(Userindex).flags.Hechizo, Userindex)
                UserList(Userindex).flags.Hechizo = 0
            Else
                Call SendData(ToIndex, Userindex, 0, "4N")
            End If
            
        Case Robar
               If TiempoTranscurrido(UserList(Userindex).Counters.LastTrabajo) < 1 Then Exit Sub
               If MapInfo(UserList(Userindex).POS.Map).Pk Or (UserList(Userindex).Clase = LADRON) Then
               
                    
                    Call LookatTile(Userindex, UserList(Userindex).POS.Map, POS.X, POS.Y)

                    If UserList(Userindex).flags.TargetUser > 0 And UserList(Userindex).flags.TargetUser <> Userindex Then
                       If UserList(UserList(Userindex).flags.TargetUser).flags.Muerto = 0 Then
                            nPos.Map = UserList(Userindex).POS.Map
                            nPos.X = POS.X
                            nPos.Y = POS.Y
                            
                            If Distancia(nPos, UserList(Userindex).POS) > 4 Or (Not (UserList(Userindex).Clase = LADRON And UserList(Userindex).Recompensas(3) = 1) And Distancia(nPos, UserList(Userindex).POS) > 2) Then
                                Call SendData(ToIndex, Userindex, 0, "DL")
                                Exit Sub
                            End If

                            Call DoRobar(Userindex, UserList(Userindex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(ToIndex, Userindex, 0, "4S")
                    End If
                Else
                    Call SendData(ToIndex, Userindex, 0, "5S")
                End If
                
        Case Domar
          
          
          
          Dim CI As Integer
          
          Call LookatTile(Userindex, UserList(Userindex).POS.Map, POS.X, POS.Y)
          CI = UserList(Userindex).flags.TargetNpc
          
          If CI Then
                   If Npclist(CI).flags.Domable Then
                        nPos.Map = UserList(Userindex).POS.Map
                        nPos.X = POS.X
                        nPos.Y = POS.Y
                        If Distancia(nPos, Npclist(UserList(Userindex).flags.TargetNpc).POS) > 2 Then
                              Call SendData(ToIndex, Userindex, 0, "DL")
                              Exit Sub
                        End If
                        If Npclist(CI).flags.AttackedBy Then
                              Call SendData(ToIndex, Userindex, 0, "7S")
                              Exit Sub
                        End If
                        Call DoDomar(Userindex, CI)
                    Else
                        Call SendData(ToIndex, Userindex, 0, "8S")
                    End If
          Else
                 Call SendData(ToIndex, Userindex, 0, "9S")
          End If
          
        Case FundirMetal
            Call LookatTile(Userindex, UserList(Userindex).POS.Map, POS.X, POS.Y)
            
            If UserList(Userindex).flags.TargetObj Then
                If ObjData(UserList(Userindex).flags.TargetObj).ObjType = OBJTYPE_FRAGUA Then
                    Call FundirMineral(Userindex)
                Else
                    Call SendData(ToIndex, Userindex, 0, "8N")
                End If
            Else
                Call SendData(ToIndex, Userindex, 0, "8N")
            End If
            
        Case Herreria
            Call LookatTile(Userindex, UserList(Userindex).POS.Map, POS.X, POS.Y)
            
            If UserList(Userindex).flags.TargetObj Then
                If ObjData(UserList(Userindex).flags.TargetObj).ObjType = OBJTYPE_YUNQUE Then
                    Call EnviarArmasConstruibles(Userindex)
                    Call EnviarArmadurasConstruibles(Userindex)
                    Call EnviarEscudosConstruibles(Userindex)
                    Call EnviarCascosConstruibles(Userindex)
                    Call SendData(ToIndex, Userindex, 0, "SFH")
                    UserList(Userindex).flags.EnviarHerreria = 1
                Else
                    Call SendData(ToIndex, Userindex, 0, "2T")
                End If
            Else
                Call SendData(ToIndex, Userindex, 0, "2T")
            End If
        Case Else

            If UserList(Userindex).flags.Trabajando = 0 Then
                Dim TrabajoPos As WorldPos
                TrabajoPos.Map = UserList(Userindex).POS.Map
                TrabajoPos.X = POS.X
                TrabajoPos.Y = POS.Y
                Call InicioTrabajo(Userindex, tLong, TrabajoPos)
            End If
            Exit Sub
            
        End Select
        
        UserList(Userindex).Counters.LastTrabajo = Timer
        Exit Sub
    Case "REL"
        If UserList(Userindex).flags.Muerto Then Exit Sub
        rdata = Right$(rdata, Len(rdata) - 3)
        Call RecibirRecompensa(Userindex, val(rdata))
        Exit Sub
    Case "CIG"
        rdata = Right$(rdata, Len(rdata) - 3)
        X = Guilds.Count
        
        If CreateGuild(UserList(Userindex).Name, Userindex, rdata) Then
            If X = 1 Then
                Call SendData(ToIndex, Userindex, 0, "3T")
            Else
                Call SendData(ToIndex, Userindex, 0, "4T" & X)
            End If
            Call UpdateUserChar(Userindex)
            
        End If
        
        Exit Sub
    Case "RSB"
        If UserList(Userindex).flags.Muerto Then Exit Sub
        rdata = Right$(rdata, Len(rdata) - 3)
        Call RecibirSubclase(CByte(rdata), Userindex)
        Exit Sub
End Select

Select Case UCase$(Left$(rdata, 4))
    Case "PRCS"
        rdata = Right$(rdata, Len(rdata) - 4)
        Call SendData(ToIndex, UserList(Userindex).flags.EsperandoLista, 0, "PRAP" & rdata)
        If rdata = "@*|" Then UserList(Userindex).flags.EsperandoLista = 0
        Exit Sub
    Case "PASS"
        rdata = Right$(rdata, Len(rdata) - 4)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        
        If UserList(Userindex).Password <> Arg1 Then
            Call SendData(ToIndex, Userindex, 0, "||El password viejo provisto no es correcto." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(Userindex).Password = Arg2
        Call SendData(ToIndex, Userindex, 0, "3V")
        
        Exit Sub
    Case "INFS"
        rdata = Right$(rdata, Len(rdata) - 4)
        If val(rdata) > 0 And val(rdata) < MAXUSERHECHIZOS + 1 Then
            Dim H As Integer
            H = UserList(Userindex).Stats.UserHechizos(val(rdata))
            If H > 0 And H < NumeroHechizos + 1 Then
                Call SendData(ToIndex, Userindex, 0, "7T" & Hechizos(H).Nombre & "¬" & Hechizos(H).Desc & "¬" & Hechizos(H).MinSkill & "¬" & ManaHechizo(Userindex, H) & "¬" & Hechizos(H).StaRequerido)
            End If
        Else
            Call SendData(ToIndex, Userindex, 0, "5T")
        End If
        Exit Sub
   Case "EQUI"
            If UserList(Userindex).flags.Muerto Then
                Call SendData(ToIndex, Userindex, 0, "MU")
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) Then
                 If UserList(Userindex).Invent.Object(val(rdata)).OBJIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            Call EquiparInvItem(Userindex, val(rdata))
            Exit Sub

    Case "CHEA"
        rdata = Right$(rdata, Len(rdata) - 4)
        If val(rdata) > 0 And val(rdata) < 5 Then
            UserList(Userindex).Char.Heading = rdata
            Call ChangeUserChar(ToPCAreaG, Userindex, UserList(Userindex).POS.Map, Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
        End If
        Exit Sub

    Case "SKSE"
        Dim sumatoria As Integer
        Dim incremento As Integer
        rdata = Right$(rdata, Len(rdata) - 4)
        
        
        
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            
            If incremento < 0 Then
                
                Call LogHackAttemp(UserList(Userindex).Name & " IP:" & UserList(Userindex).ip & " trato de hackear los skills.")
                UserList(Userindex).Stats.SkillPts = 0
                Call CloseSocket(Userindex)
                Exit Sub
            End If
            
            sumatoria = sumatoria + incremento
        Next
        
        If sumatoria > UserList(Userindex).Stats.SkillPts Then
            
            
            Call LogHackAttemp(UserList(Userindex).Name & " IP:" & UserList(Userindex).ip & " trato de hackear los skills.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        
        
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            UserList(Userindex).Stats.SkillPts = UserList(Userindex).Stats.SkillPts - incremento
            UserList(Userindex).Stats.UserSkills(i) = UserList(Userindex).Stats.UserSkills(i) + incremento
            If UserList(Userindex).Stats.UserSkills(i) > 100 Then UserList(Userindex).Stats.UserSkills(i) = 100
        Next
        Exit Sub
    Case "ENTR"
        
        If UserList(Userindex).flags.TargetNpc = 0 Then Exit Sub
        
        If Npclist(UserList(Userindex).flags.TargetNpc).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
        
        rdata = Right$(rdata, Len(rdata) - 4)
        
        If Npclist(UserList(Userindex).flags.TargetNpc).Mascotas < MAXMASCOTASENTRENADOR Then
            If val(rdata) > 0 And val(rdata) < Npclist(UserList(Userindex).flags.TargetNpc).NroCriaturas + 1 Then
                Dim SpawnedNpc As Integer
                SpawnedNpc = SpawnNpc(Npclist(UserList(Userindex).flags.TargetNpc).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(Userindex).flags.TargetNpc).POS, True, False)
                If SpawnedNpc <= MAXNPCS Then
                    Npclist(SpawnedNpc).MaestroNpc = UserList(Userindex).flags.TargetNpc
                    Npclist(UserList(Userindex).flags.TargetNpc).Mascotas = Npclist(UserList(Userindex).flags.TargetNpc).Mascotas + 1
                    
                End If
            End If
        Else
            Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "3Q" & vbWhite & "°" & "No puedo traer más criaturas, mata las existentes!" & "°" & str(Npclist(UserList(Userindex).flags.TargetNpc).Char.CharIndex))
        End If
        
        Exit Sub
    Case "COMP"
         
         If UserList(Userindex).flags.Muerto Then
            Call SendData(ToIndex, Userindex, 0, "MU")
            Exit Sub
         End If
         
         rdata = Right$(rdata, Len(rdata) - 4)
         If UserList(Userindex).flags.TargetNpc Then
         
            If Npclist(UserList(Userindex).flags.TargetNpc).NPCtype = NPCTYPE_TIENDA Then
                Call TiendaVentaItem(Userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(Userindex).flags.TargetNpc)
                Exit Sub
            End If
               
            If Npclist(UserList(Userindex).flags.TargetNpc).Comercia = 0 Then
                Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "3Q" & FONTTYPE_TALK & "°" & "No tengo ningún interes en comerciar." & "°" & str(Npclist(UserList(Userindex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
         Else: Exit Sub
         End If
         
         
         Call NPCVentaItem(Userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(Userindex).flags.TargetNpc)
         Exit Sub
    Case "RETI"
        If UserList(Userindex).flags.Muerto Then
           Call SendData(ToIndex, Userindex, 0, "MU")
           Exit Sub
        End If
        
        If UserList(Userindex).flags.TargetNpc Then
           If Npclist(UserList(Userindex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
        Else: Exit Sub
        
        End If
        rdata = Right$(rdata, Len(rdata) - 4)
        Call UserRetiraItem(Userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
        
        Exit Sub
         
    Case "POVE"
        If Npclist(UserList(Userindex).flags.TargetNpc).flags.TiendaUser Then
            If Npclist(UserList(Userindex).flags.TargetNpc).flags.TiendaUser <> Userindex Then Exit Sub
        Else
            Npclist(UserList(Userindex).flags.TargetNpc).flags.TiendaUser = Userindex
        End If
        
        If UserList(Userindex).flags.Muerto Then
            Call SendData(ToIndex, Userindex, 0, "MU")
            Exit Sub
         End If

         If UserList(Userindex).flags.TargetNpc Then
            If Npclist(UserList(Userindex).flags.TargetNpc).NPCtype <> NPCTYPE_TIENDA Then
                Exit Sub
            End If
         Else: Exit Sub
         End If
         
         rdata = Right$(rdata, Len(rdata) - 4)
         
         Call UserPoneVenta(Userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), val(ReadField(3, rdata, 44)))
         
         Exit Sub
    
    Case "SAVE"
         If UserList(Userindex).flags.Muerto Then
            Call SendData(ToIndex, Userindex, 0, "MU")
            Exit Sub
         End If
         If UserList(Userindex).flags.TargetNpc Then
            If Npclist(UserList(Userindex).flags.TargetNpc).NPCtype <> NPCTYPE_TIENDA Then Exit Sub
         Else: Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 4)
         Call UserSacaVenta(Userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
         
    Case "VEND"
         
         If UserList(Userindex).flags.Muerto Then
            Call SendData(ToIndex, Userindex, 0, "MU")
            Exit Sub
         End If

         If UserList(Userindex).flags.TargetNpc Then
               If Npclist(UserList(Userindex).flags.TargetNpc).NPCtype = NPCTYPE_TIENDA Then
                   Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "/N")
                   Exit Sub
               End If
               
               If Npclist(UserList(Userindex).flags.TargetNpc).Comercia = 0 Then
                   Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "3Q" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(Userindex).flags.TargetNpc).Char.CharIndex))
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 4)
         Call NPCCompraItem(Userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub

    Case "DEPO"
         If UserList(Userindex).flags.Muerto Then
            Call SendData(ToIndex, Userindex, 0, "MU")
            Exit Sub
         End If
         If UserList(Userindex).flags.TargetNpc Then
            If Npclist(UserList(Userindex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
         Else: Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 4)

         Call UserDepositaItem(Userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    
    
         
End Select

Select Case UCase$(Left$(rdata, 5))
    Case "DEMSG"
        
        
        If UserList(Userindex).flags.TargetObj Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Dim f As String, Titu As String, msg As String, f2 As String
   
        f = App.Path & "\foros\"
        f = f & UCase$(ObjData(UserList(Userindex).flags.TargetObj).ForoID) & ".for"
        Titu = ReadField(1, rdata, 176)
        msg = ReadField(2, rdata, 176)
   
        Dim n2 As Integer, loopme As Integer
        If FileExist(f, vbNormal) Then
            Dim Num As Integer
            Num = val(GetVar(f, "INFO", "CantMSG"))
            If Num > MAX_MENSAJES_FORO Then
                For loopme = 1 To Num
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(Userindex).flags.TargetObj).ForoID) & loopme & ".for"
                Next
                Kill App.Path & "\foros\" & UCase$(ObjData(UserList(Userindex).flags.TargetObj).ForoID) & ".for"
                Num = 0
            End If
          
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & Num + 1 & ".for"
            Open f2 For Output As n2
            Print #n2, Titu
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", Num + 1)
        Else
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & "1" & ".for"
            Open f2 For Output As n2
            Print #n2, Titu
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", 1)
        End If
        Close #n2
        End If
        Exit Sub
End Select

Select Case UCase$(Left$(rdata, 6))
    Case "DESCOD"
            rdata = Right$(rdata, Len(rdata) - 6)
            Call UpdateCodexAndDesc(rdata, Userindex)
            Exit Sub
    Case "DESPHE"
            rdata = Right$(rdata, Len(rdata) - 6)
            Call DesplazarHechizo(Userindex, CInt(ReadField(1, rdata, 44)), CByte(ReadField(2, rdata, 44)))
            Exit Sub
    Case "PARACE"
        If UserList(Userindex).flags.Ofreciente = 0 Then Exit Sub
        If Not UserList(UserList(Userindex).flags.Ofreciente).flags.UserLogged Then Exit Sub

        If NoPuedeEntrarParty(UserList(Userindex).flags.Ofreciente, Userindex) Then Exit Sub
    
        Dim PartyIndex As Integer
        If UserList(UserList(Userindex).flags.Ofreciente).flags.Party Then
            PartyIndex = UserList(UserList(Userindex).flags.Ofreciente).PartyIndex
            If PartyIndex = 0 Then Exit Sub
            Call EntrarAlParty(Userindex, PartyIndex)
        Else
            Call CrearParty(Userindex)
        End If
        Exit Sub
    Case "PARREC"
        If UserList(Userindex).flags.Ofreciente = 0 Then Exit Sub
        If Not UserList(UserList(Userindex).flags.Ofreciente).flags.UserLogged Then Exit Sub
        Call SendData(ToIndex, Userindex, 0, "||Rechazaste entrar a party con " & UserList(UserList(Userindex).flags.Ofreciente).Name & "." & FONTTYPE_PARTY)
        Call SendData(ToIndex, UserList(Userindex).flags.Ofreciente, 0, "||" & UserList(Userindex).Name & " rechazo entrar en party con vos." & FONTTYPE_PARTY)
        UserList(Userindex).flags.Ofreciente = 0
        Exit Sub
    Case "PARECH"
        rdata = ReadField(1, Right$(rdata, Len(rdata) - 6), Asc("("))
        rdata = Left$(rdata, Len(rdata) - 1)
        If UserList(Userindex).flags.Party Then
            If Party(UserList(Userindex).PartyIndex).NroMiembros = 2 Then
                For i = 1 To Party(UserList(Userindex).PartyIndex).NroMiembros
                    Call RomperParty(Userindex)
                Next
            Else
                Call EcharDelParty(NameIndex(rdata))
            End If
        Else
            Call SendData(ToIndex, Userindex, 0, "||No estás en party." & FONTTYPE_PARTY)
        End If
        Exit Sub
            
 End Select


Select Case UCase$(Left$(rdata, 7))
Case "OFRECER"
        rdata = Right$(rdata, Len(rdata) - 7)
        Arg1 = ReadField(1, rdata, Asc(","))
        Arg2 = ReadField(2, rdata, Asc(","))

        If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
            Exit Sub
        End If
        If Not UserList(UserList(Userindex).ComUsu.DestUsu).flags.UserLogged Then
            
            Call FinComerciarUsu(Userindex)
            Exit Sub
        Else
            
            If UserList(UserList(Userindex).ComUsu.DestUsu).flags.Muerto Then
                Call FinComerciarUsu(Userindex)
                Exit Sub
            End If
            
            If val(Arg1) = FLAGORO Then
                
                If val(Arg2) > UserList(Userindex).Stats.GLD Then
                    Call SendData(ToIndex, Userindex, 0, "4R")
                    Exit Sub
                End If
            Else
                
                If val(Arg2) > UserList(Userindex).Invent.Object(val(Arg1)).Amount Then
                    Call SendData(ToIndex, Userindex, 0, "4R")
                    Exit Sub
                End If
                If ObjData(UserList(Userindex).Invent.Object(val(Arg1)).OBJIndex).NoSeCae Or ObjData(UserList(Userindex).Invent.Object(val(Arg1)).OBJIndex).Newbie = 1 Or ObjData(UserList(Userindex).Invent.Object(val(Arg1)).OBJIndex).Real > 0 Or ObjData(UserList(Userindex).Invent.Object(val(Arg1)).OBJIndex).Caos > 0 Then
                    Call SendData(ToIndex, Userindex, 0, "||No puedes ofrecer este objeto." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            
            If UserList(Userindex).ComUsu.Objeto Then
                Call SendData(ToIndex, Userindex, 0, "6T")
                Exit Sub
            End If
            UserList(Userindex).ComUsu.Objeto = val(Arg1)
            UserList(Userindex).ComUsu.Cant = val(Arg2)
            If UserList(UserList(Userindex).ComUsu.DestUsu).ComUsu.DestUsu <> Userindex Then
                Call FinComerciarUsu(Userindex)
                Exit Sub
            Else
                
                If UserList(UserList(Userindex).ComUsu.DestUsu).ComUsu.Acepto Then
                    
                    UserList(UserList(Userindex).ComUsu.DestUsu).ComUsu.Acepto = False
                    Call SendData(ToIndex, UserList(Userindex).ComUsu.DestUsu, 0, "5R" & UserList(Userindex).Name)
                End If
                
                
                Call EnviarObjetoTransaccion(UserList(Userindex).ComUsu.DestUsu)
            End If
        End If
        Exit Sub
End Select


Select Case UCase$(Left$(rdata, 8))
    Case "ACEPPEAT"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptPeaceOffer(Userindex, rdata)
        Exit Sub
    Case "PEACEOFF"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call RecievePeaceOffer(Userindex, rdata)
        Exit Sub
    Case "PEACEDET"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeaceRequest(Userindex, rdata)
        Exit Sub
    Case "ENVCOMEN"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeticion(Userindex, rdata)
        Exit Sub
    Case "ENVPROPP"
        Call SendPeacePropositions(Userindex)
        Exit Sub
    Case "DECGUERR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareWar(Userindex, rdata)
        Exit Sub
    Case "DECALIAD"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareAllie(Userindex, rdata)
        Exit Sub
    Case "NEWWEBSI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SetNewURL(Userindex, rdata)
        Exit Sub
    Case "ACEPTARI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptClanMember(Userindex, rdata)
        Exit Sub
    Case "RECHAZAR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DenyRequest(Userindex, rdata)
        Exit Sub
    Case "ECHARCLA"
        Dim eslider As Integer
        rdata = Right$(rdata, Len(rdata) - 8)
        tIndex = NameIndex(rdata)
        If UserList(Userindex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
        Call EcharMember(Userindex, rdata)
        Exit Sub
    Case "ACTGNEWS"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call UpdateGuildNews(rdata, Userindex)
        Exit Sub
    Case "1HRINFO<"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendCharInfo(rdata, Userindex)
        Exit Sub
End Select

Select Case UCase$(Left$(rdata, 9))
    Case "SOLICITUD"
         rdata = Right$(rdata, Len(rdata) - 9)
         Call SolicitudIngresoClan(Userindex, rdata)
         Exit Sub
End Select

Select Case UCase$(Left$(rdata, 11))
  Case "CLANDETAILS"
        rdata = Right$(rdata, Len(rdata) - 11)
        Call SendGuildDetails(Userindex, rdata)
        Exit Sub
End Select

Procesado = False
End Sub

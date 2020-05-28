Attribute VB_Name = "UsUaRiOs"
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
Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

If UserList(AttackerIndex).POS.Map <> 190 Then
    Dim DaExp As Integer
    DaExp = CInt(UserList(VictimIndex).Stats.ELV * RandomNumber(1, 4))
    Call AddtoVar(UserList(AttackerIndex).Stats.Exp, DaExp, MAXEXP)
End If

Call SendData(ToIndex, AttackerIndex, 0, "1Q" & UserList(VictimIndex).Name)
Call SendData(ToIndex, AttackerIndex, 0, "EX" & DaExp)
Call SendData(ToIndex, VictimIndex, 0, "1R" & UserList(AttackerIndex).Name)

Call UserDie(VictimIndex)

End Sub
Sub RevivirUsuarioNPC(Userindex As Integer)

UserList(Userindex).flags.Muerto = 0
UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MaxHP

Call DarCuerpoDesnudo(Userindex)
Call ChangeUserChar(ToMap, 0, UserList(Userindex).POS.Map, Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
Call SendUserStatsBox(Userindex)

End Sub
Sub RevivirUsuario(ByVal Resucitador As Integer, Userindex As Integer, ByVal Lleno As Boolean)

UserList(Resucitador).Stats.MinSta = 0
UserList(Resucitador).Stats.MinAGU = 0
UserList(Resucitador).Stats.MinHam = 0
UserList(Resucitador).flags.Sed = 1
UserList(Resucitador).flags.Hambre = 1

UserList(Userindex).flags.Muerto = 0

If Lleno Then
    UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MaxHP
    UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MaxSta
    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
    UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MaxHam
    UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
    UserList(Userindex).flags.Sed = 0
    UserList(Userindex).flags.Hambre = 0
Else
    UserList(Userindex).Stats.MinHP = 1
    UserList(Userindex).Stats.MinSta = 0
    UserList(Userindex).Stats.MinMAN = 0
    UserList(Userindex).Stats.MinHam = 0
    UserList(Userindex).Stats.MinAGU = 0
    UserList(Userindex).flags.Sed = 1
    UserList(Userindex).flags.Hambre = 1
End If

Call DarCuerpoDesnudo(Userindex)
Call ChangeUserChar(ToMap, 0, UserList(Userindex).POS.Map, Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)

Call SendUserStatsBox(Resucitador)
Call EnviarHambreYsed(Resucitador)

Call SendUserStatsBox(Userindex)
Call EnviarHambreYsed(Userindex)

End Sub
Sub ReNombrar(Userindex As Integer, NewNick As String)

Call SendData(ToAll, 0, 0, "||El usuario " & UserList(Userindex).Name & " ha sido rebautizado como " & NewNick & "." & FONTTYPE_FIGHT)
UserList(Userindex).Name = NewNick
Call WarpUserChar(Userindex, UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y, False)

End Sub
Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Userindex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

UserList(Userindex).Char.Body = Body
UserList(Userindex).Char.Head = Head
UserList(Userindex).Char.Heading = Heading
UserList(Userindex).Char.WeaponAnim = Arma
UserList(Userindex).Char.ShieldAnim = Escudo
UserList(Userindex).Char.CascoAnim = Casco

Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(Userindex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(Userindex).Char.FX & "," & UserList(Userindex).Char.loops & "," & Casco)

End Sub
Sub ChangeUserCharB(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Userindex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

UserList(Userindex).Char.Body = Body
UserList(Userindex).Char.Head = Head
UserList(Userindex).Char.Heading = Heading
UserList(Userindex).Char.WeaponAnim = Arma
UserList(Userindex).Char.ShieldAnim = Escudo
UserList(Userindex).Char.CascoAnim = Casco

Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(Userindex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(Userindex).Char.FX & "," & UserList(Userindex).Char.loops & "," & Casco & "," & UserList(Userindex).flags.Navegando)

End Sub
Sub ChangeUserCasco(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Userindex As Integer, _
ByVal Casco As Integer)

On Error Resume Next

If UserList(Userindex).Char.CascoAnim <> Casco Then
UserList(Userindex).Char.CascoAnim = Casco
Call SendData(sndRoute, sndIndex, sndMap, "7C" & UserList(Userindex).Char.CharIndex & "," & Casco)
End If

End Sub
Sub ChangeUserEscudo(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Userindex As Integer, ByVal Escudo As Integer)
On Error Resume Next

If UserList(Userindex).Char.ShieldAnim <> Escudo Then
    UserList(Userindex).Char.ShieldAnim = Escudo
    Call SendData(sndRoute, sndIndex, sndMap, "6C" & UserList(Userindex).Char.CharIndex & "," & Escudo)
End If

End Sub


Sub ChangeUserArma(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Userindex As Integer, _
ByVal Arma As Integer)

On Error Resume Next

If UserList(Userindex).Char.WeaponAnim <> Arma Then
    UserList(Userindex).Char.WeaponAnim = Arma
    Call SendData(sndRoute, sndIndex, sndMap, "5C" & UserList(Userindex).Char.CharIndex & "," & Arma)
End If


End Sub


Sub ChangeUserHead(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Userindex As Integer, _
ByVal Head As Integer)

On Error Resume Next

If UserList(Userindex).Char.Head <> Head Then
UserList(Userindex).Char.Head = Head
Call SendData(sndRoute, sndIndex, sndMap, "4C" & UserList(Userindex).Char.CharIndex & "," & Head)
End If

End Sub

Sub ChangeUserBody(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Userindex As Integer, _
ByVal Body As Integer)

On Error Resume Next
UserList(Userindex).Char.Body = Body
Call SendData(sndRoute, sndIndex, sndMap, "3C" & UserList(Userindex).Char.CharIndex & "," & Body)


End Sub
Sub ChangeUserHeading(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Userindex As Integer, _
ByVal Heading As Byte)
On Error Resume Next

UserList(Userindex).Char.Heading = Heading
Call SendData(sndRoute, sndIndex, sndMap, "2C" & UserList(Userindex).Char.CharIndex & "," & Heading)

End Sub
Sub EnviarSubirNivel(Userindex As Integer, ByVal Puntos As Integer)

Call SendData(ToIndex, Userindex, 0, "SUNI" & Puntos)

End Sub
Sub EnviarSkills(Userindex As Integer)
Dim i As Integer
Dim cad As String

For i = 1 To NUMSKILLS
   cad = cad & UserList(Userindex).Stats.UserSkills(i) & ","
Next

SendData ToIndex, Userindex, 0, "SKILLS" & cad

End Sub
Sub EnviarFama(Userindex As Integer)
Dim cad As String

cad = UserList(Userindex).Faccion.Quests & ","
cad = cad & UserList(Userindex).Faccion.Torneos & ","
    
If EsNewbie(Userindex) Then
    cad = cad & UserList(Userindex).Faccion.Matados(Caos) & ","
    cad = cad & UserList(Userindex).Faccion.Matados(Neutral)
    
    Call SendData(ToIndex, Userindex, 0, "FAMA3," & cad)
Else
    Select Case UserList(Userindex).Faccion.Bando
        Case Neutral
            cad = cad & UserList(Userindex).Faccion.BandoOriginal & ","
            cad = cad & UserList(Userindex).Faccion.Matados(Real) & ","
            cad = cad & UserList(Userindex).Faccion.Matados(Caos) & ","
            
        Case Real, Caos
            cad = cad & Titulo(Userindex) & ","
            cad = cad & UserList(Userindex).Faccion.Matados(Enemigo(UserList(Userindex).Faccion.Bando)) & ","
            
    End Select
    cad = cad & UserList(Userindex).Faccion.Matados(Neutral)
    Call SendData(ToIndex, Userindex, 0, "FAMA" & UserList(Userindex).Faccion.Bando & "," & cad)
End If

End Sub
Function GeneroLetras(Genero As Byte) As String

If Genero = 1 Then
    GeneroLetras = "Mujer"
Else
    GeneroLetras = "Hombre"
End If

End Function
Sub EnviarMiniSt(Userindex As Integer)
Dim cad As String

cad = cad & UserList(Userindex).Stats.VecesMurioUsuario & ","
cad = cad & UserList(Userindex).Faccion.Matados(Caos) & ","
cad = cad & UserList(Userindex).Stats.NPCsMuertos & ","
cad = cad & UserList(Userindex).Faccion.Matados(Neutral) + UserList(Userindex).Faccion.Matados(Real) + UserList(Userindex).Faccion.Matados(Caos) & ","
cad = cad & ListaClases(UserList(Userindex).Clase) & ","
cad = cad & ListaRazas(UserList(Userindex).Raza) & ","
cad = cad & UserList(Userindex).Faccion.Matados(Real) & ","

Call SendData(ToIndex, Userindex, 0, "MIST" & cad)

End Sub
Sub EnviarAtrib(Userindex As Integer)
Dim i As Integer
Dim cad As String

For i = 1 To NUMATRIBUTOS
  cad = cad & UserList(Userindex).Stats.UserAtributos(i) & ","
Next

Call SendData(ToIndex, Userindex, 0, "ATR" & cad)

End Sub
Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, Userindex As Integer)

On Error GoTo ErrorHandler

CharList(UserList(Userindex).Char.CharIndex) = 0

If UserList(Userindex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

MapData(UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y).Userindex = 0


Call SendData(ToMap, Userindex, UserList(Userindex).POS.Map, "BP" & UserList(Userindex).Char.CharIndex)

UserList(Userindex).Char.CharIndex = 0

NumChars = NumChars - 1

Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar")

End Sub
Sub UpdateUserChar(Userindex As Integer)
On Error Resume Next
Dim bCr As Byte
Dim Info As String

If UserList(Userindex).flags.Privilegios Then
    bCr = 1
ElseIf UserList(Userindex).Faccion.Bando = Real Then
    bCr = 2
ElseIf UserList(Userindex).Faccion.Bando = Caos Then
    bCr = 3
ElseIf EsNewbie(Userindex) Then
    bCr = 4
Else: bCr = 5
End If

Info = "PW" & UserList(Userindex).Char.CharIndex & "," & bCr & "," & UserList(Userindex).Name

If Len(UserList(Userindex).GuildInfo.GuildName) > 0 Then Info = Info & " <" & UserList(Userindex).GuildInfo.GuildName & ">"

Call SendData(ToMap, Userindex, UserList(Userindex).POS.Map, (Info))

End Sub
Sub MakeUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, Userindex As Integer, Map As Integer, X As Integer, Y As Integer)
On Error Resume Next
Dim CharIndex As Integer

If Not InMapBounds(X, Y) Then Exit Sub


If UserList(Userindex).Char.CharIndex = 0 Then
    CharIndex = NextOpenCharIndex
    UserList(Userindex).Char.CharIndex = CharIndex
    CharList(CharIndex) = Userindex
End If


MapData(Map, X, Y).Userindex = Userindex


Dim klan$
klan$ = UserList(Userindex).GuildInfo.GuildName
Dim bCr As Byte
If UserList(Userindex).flags.Privilegios Then
    bCr = 1
ElseIf UserList(Userindex).Faccion.Bando = Real And UserList(Userindex).flags.EsConseReal = 0 Then
    bCr = 2
ElseIf UserList(Userindex).Faccion.Bando = Caos And UserList(Userindex).flags.EsConseCaos = 0 Then
    bCr = 3
ElseIf EsNewbie(Userindex) Then
    bCr = 4
ElseIf UserList(Userindex).flags.EsConseCaos Then
    bCr = 6
ElseIf UserList(Userindex).flags.EsConseReal Then
    bCr = 7
Else
    bCr = 5
End If

If Len(klan$) > 0 Then klan = " <" & klan$ & ">"
Call Logear("test", "CC" & UserList(Userindex).Char.Body & "," & UserList(Userindex).Char.Head & "," & UserList(Userindex).Char.Heading & "," & UserList(Userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(Userindex).Char.WeaponAnim & "," & UserList(Userindex).Char.ShieldAnim & "," & UserList(Userindex).Char.FX & "," & 999 & "," & UserList(Userindex).Char.CascoAnim & "," & UserList(Userindex).Name & klan$ & "," & bCr & "," & UserList(Userindex).flags.Invisible)
Call SendData(sndRoute, sndIndex, sndMap, ("CC" & UserList(Userindex).Char.Body & "," & UserList(Userindex).Char.Head & "," & UserList(Userindex).Char.Heading & "," & UserList(Userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(Userindex).Char.WeaponAnim & "," & UserList(Userindex).Char.ShieldAnim & "," & UserList(Userindex).Char.FX & "," & 999 & "," & UserList(Userindex).Char.CascoAnim & "," & UserList(Userindex).Name & klan$ & "," & bCr & "," & UserList(Userindex).flags.Invisible))

If UserList(Userindex).flags.Meditando Then
    UserList(Userindex).Char.loops = LoopAdEternum
    If UserList(Userindex).Stats.ELV < 15 Then
        Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
        UserList(Userindex).Char.FX = FXMEDITARCHICO
    ElseIf UserList(Userindex).Stats.ELV < 30 Then
        Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
        UserList(Userindex).Char.FX = FXMEDITARMEDIANO
    Else
        Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
        UserList(Userindex).Char.FX = FXMEDITARGRANDE
    End If
End If

End Sub
Function Redondea(ByVal Number As Single) As Integer

If Number > Fix(Number) Then
    Redondea = Fix(Number) + 1
Else: Redondea = Number
End If

End Function
Sub CheckUserLevel(Userindex As Integer)
On Error GoTo errhandler
Dim Pts As Integer
Dim SubeHit As Integer
Dim AumentoST As Integer
Dim AumentoMANA As Integer
Dim WasNewbie As Boolean

Do Until UserList(Userindex).Stats.Exp < UserList(Userindex).Stats.ELU
If UserList(Userindex).Stats.ELV >= STAT_MAXELV Then
    UserList(Userindex).Stats.Exp = 0
    UserList(Userindex).Stats.ELU = 0
    Exit Sub
End If

WasNewbie = EsNewbie(Userindex)

If UserList(Userindex).Stats.Exp >= UserList(Userindex).Stats.ELU Then

    If UserList(Userindex).Stats.ELV >= 14 And ClaseBase(UserList(Userindex).Clase) Then
        Call SendData(ToIndex, Userindex, 0, "!6")
        UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.ELU - 1
        Call SendUserEXP(Userindex)
        Exit Sub
    End If
    
    Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "TW" & SOUND_NIVEL)
    Call SendData(ToIndex, Userindex, 0, "1S" & UserList(Userindex).Stats.ELV + 1)
    
    If UserList(Userindex).Stats.ELV = 1 Then
        Pts = 10
    Else
        Pts = 5
    End If
    
    UserList(Userindex).Stats.SkillPts = UserList(Userindex).Stats.SkillPts + Pts
    
    Call SendData(ToIndex, Userindex, 0, "1T" & Pts)
       
    UserList(Userindex).Stats.ELV = UserList(Userindex).Stats.ELV + 1
    UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp - UserList(Userindex).Stats.ELU
    UserList(Userindex).Stats.ELU = ELUs(UserList(Userindex).Stats.ELV)
    
    Dim AumentoHP As Integer
    Dim SubePromedio As Single
    
    SubePromedio = UserList(Userindex).Stats.UserAtributos(Constitucion) / 2 - Resta(UserList(Userindex).Clase)
    AumentoHP = RandomNumber(Fix(SubePromedio - 1), Redondea(SubePromedio + 1))
    SubeHit = AumentoHit(UserList(Userindex).Clase)

    Select Case UserList(Userindex).Clase
        Case CIUDADANO, TRABAJADOR, EXPERTO_MINERALES
            AumentoST = 15
            
        Case MINERO
            AumentoST = 15 + AdicionalSTMinero
            
        Case HERRERO
            AumentoST = 15
            
        Case EXPERTO_MADERA
            AumentoST = 15

        Case TALADOR
            AumentoST = 15 + AdicionalSTLeñador

        Case CARPINTERO
            AumentoST = 15
            
        Case PESCADOR
            AumentoST = 15 + AdicionalSTPescador
            
        Case SASTRE
            AumentoST = 15
            
        Case HECHICERO
            AumentoST = 15
            AumentoMANA = 2.2 * UserList(Userindex).Stats.UserAtributos(Inteligencia)
            
        Case MAGO
            AumentoST = Maximo(5, 15 - AdicionalSTLadron / 2)
            Select Case UserList(Userindex).Stats.MaxMAN
                Case Is < 2300
                    AumentoMANA = 3 * UserList(Userindex).Stats.UserAtributos(Inteligencia)
                Case Is < 2500
                    AumentoMANA = 2 * UserList(Userindex).Stats.UserAtributos(Inteligencia)
                Case Else
                    AumentoMANA = 1.5 * UserList(Userindex).Stats.UserAtributos(Inteligencia)
            End Select
            
            If UserList(Userindex).Stats.ELV > 45 Then AumentoMANA = 0
            
        Case NIGROMANTE
            AumentoST = Maximo(5, 15 - AdicionalSTLadron / 2)
            AumentoMANA = 2.2 * UserList(Userindex).Stats.UserAtributos(Inteligencia)
            
        Case ORDEN_SAGRADA
            AumentoST = 15
            AumentoMANA = UserList(Userindex).Stats.UserAtributos(Inteligencia)
            
        Case PALADIN
            AumentoST = 15
            AumentoMANA = UserList(Userindex).Stats.UserAtributos(Inteligencia)
            
            If UserList(Userindex).Stats.MaxHit >= 99 Then SubeHit = 1
            
        Case CLERIGO
            AumentoST = 15
            AumentoMANA = 2 * UserList(Userindex).Stats.UserAtributos(Inteligencia)

        Case NATURALISTA
            AumentoST = 15
            AumentoMANA = 2 * UserList(Userindex).Stats.UserAtributos(Inteligencia)
            
        Case BARDO
            AumentoST = 15
            AumentoMANA = 2 * UserList(Userindex).Stats.UserAtributos(Inteligencia)

        Case DRUIDA
            AumentoST = 15
            AumentoMANA = 2.2 * UserList(Userindex).Stats.UserAtributos(Inteligencia)

        Case SIGILOSO
            AumentoST = 15
            AumentoMANA = UserList(Userindex).Stats.UserAtributos(Inteligencia)
            
        Case ASESINO
            AumentoST = 15
            AumentoMANA = UserList(Userindex).Stats.UserAtributos(Inteligencia)

            If UserList(Userindex).Stats.MaxHit >= 99 Then SubeHit = 1
            
        Case CAZADOR
            AumentoST = 15
            AumentoMANA = UserList(Userindex).Stats.UserAtributos(Inteligencia)

            If UserList(Userindex).Stats.MaxHit >= 99 Then SubeHit = 1
            
        Case SIN_MANA
            AumentoST = 15

        Case CABALLERO
            AumentoST = 15
            
        Case ARQUERO
            AumentoST = 15
         
            If UserList(Userindex).Stats.MaxHit >= 99 Then SubeHit = 2
            
        Case GUERRERO
            AumentoST = 15

            If UserList(Userindex).Stats.MaxHit >= 99 Then SubeHit = 2
           
        Case BANDIDO
            AumentoST = 15
            
        Case PIRATA
            AumentoST = 15

        Case LADRON
            AumentoST = 15
         
        Case Else
            AumentoST = 15 + AdicionalSTLadron
            
    End Select
       
    Call AddtoVar(UserList(Userindex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
    UserList(Userindex).Stats.MaxSta = UserList(Userindex).Stats.MaxSta + AumentoST
    
    Call AddtoVar(UserList(Userindex).Stats.MaxMAN, AumentoMANA, 2200 + 800 * Buleano(UserList(Userindex).Clase And UserList(Userindex).Recompensas(2) = 2))
    UserList(Userindex).Stats.MaxHit = UserList(Userindex).Stats.MaxHit + SubeHit
    UserList(Userindex).Stats.MinHit = UserList(Userindex).Stats.MinHit + SubeHit
    
    Call SendData(ToIndex, Userindex, 0, "1U" & AumentoHP & "," & AumentoST & "," & AumentoMANA & "," & SubeHit)
    
    UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MaxHP
    
    Call EnviarSkills(Userindex)
    Call EnviarSubirNivel(Userindex, Pts)
   
    Call SendUserStatsBox(Userindex)
    
    If Not EsNewbie(Userindex) And WasNewbie Then
        If UserList(Userindex).POS.Map = 37 Or UserList(Userindex).POS.Map = 49 Then
            Call WarpUserChar(Userindex, 1, 50, 50, True)
        Else
            Call UpdateUserChar(Userindex)
        End If
        Call QuitarNewbieObj(Userindex)
        Call SendData(ToIndex, Userindex, 0, "SUFA1")
    End If
    
    Call CheckUserLevel(Userindex)
    
Else

    Call SendUserEXP(Userindex)
    
End If

    
If PuedeSubirClase(Userindex) Then Call SendData(ToIndex, Userindex, 0, "SUCL1")
If PuedeRecompensa(Userindex) Then Call SendData(ToIndex, Userindex, 0, "SURE1")

Loop

Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub
Function PuedeRecompensa(Userindex As Integer) As Byte

If UserList(Userindex).Clase = SASTRE Then Exit Function

If UserList(Userindex).Recompensas(1) = 0 And UserList(Userindex).Stats.ELV >= 18 Then
    PuedeRecompensa = 1
    Exit Function
End If

If UserList(Userindex).Clase = TALADOR Or UserList(Userindex).Clase = PESCADOR Then Exit Function

If UserList(Userindex).Stats.ELV >= 25 And UserList(Userindex).Recompensas(2) = 0 Then
    PuedeRecompensa = 2
    Exit Function
End If
    
If UserList(Userindex).Clase = CARPINTERO Then Exit Function

If UserList(Userindex).Recompensas(3) = 0 And _
    (UserList(Userindex).Stats.ELV >= 34 Or _
    (ClaseTrabajadora(UserList(Userindex).Clase) And UserList(Userindex).Stats.ELV >= 32) Or _
    ((UserList(Userindex).Clase = PIRATA Or UserList(Userindex).Clase = LADRON) And UserList(Userindex).Stats.ELV >= 30)) Then
    PuedeRecompensa = 3
    Exit Function
End If

End Function
Function PuedeFaccion(Userindex As Integer) As Boolean

PuedeFaccion = Not EsNewbie(Userindex) And UserList(Userindex).Faccion.BandoOriginal = Neutral And Len(UserList(Userindex).GuildInfo.GuildName) = 0 And UserList(Userindex).flags.Privilegios = 0

End Function
Function PuedeSubirClase(Userindex As Integer) As Boolean

PuedeSubirClase = (UserList(Userindex).Stats.ELV >= 3 And UserList(Userindex).Clase = CIUDADANO) Or _
                (UserList(Userindex).Stats.ELV >= 6 And (UserList(Userindex).Clase = LUCHADOR Or UserList(Userindex).Clase = TRABAJADOR)) Or _
                (UserList(Userindex).Stats.ELV >= 9 And (UserList(Userindex).Clase = EXPERTO_MINERALES Or UserList(Userindex).Clase = EXPERTO_MADERA Or UserList(Userindex).Clase = CON_MANA Or UserList(Userindex).Clase = SIN_MANA)) Or _
                (UserList(Userindex).Stats.ELV >= 12 And (UserList(Userindex).Clase = CABALLERO Or UserList(Userindex).Clase = BANDIDO Or UserList(Userindex).Clase = HECHICERO Or UserList(Userindex).Clase = NATURALISTA Or UserList(Userindex).Clase = ORDEN_SAGRADA Or UserList(Userindex).Clase = SIGILOSO))

End Function
Function PuedeAtravesarAgua(Userindex As Integer) As Boolean

PuedeAtravesarAgua = UserList(Userindex).flags.Navegando = 1

End Function
Private Sub EnviaNuevaPosUsuarioPj(Userindex As Integer, ByVal Quien As Integer)

Call SendData(ToIndex, Userindex, 0, ("LP" & UserList(Quien).Char.CharIndex & "," & UserList(Quien).POS.X & "," & UserList(Quien).POS.Y & "," & UserList(Quien).Char.Heading))

End Sub
Private Sub EnviaNuevaPosNPC(Userindex As Integer, NpcIndex As Integer)

Call SendData(ToIndex, Userindex, 0, ("LP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).POS.X & "," & Npclist(NpcIndex).POS.Y & "," & Npclist(NpcIndex).Char.Heading))

End Sub
Sub CalcularValores(Userindex As Integer)
Dim SubePromedio As Single
Dim HPReal As Integer
Dim HitReal As Integer
Dim i As Integer

HPReal = 15 + RandomNumber(1, UserList(Userindex).Stats.UserAtributos(Constitucion) \ 3)
HitReal = AumentoHit(UserList(Userindex).Clase) * UserList(Userindex).Stats.ELV
SubePromedio = UserList(Userindex).Stats.UserAtributos(Constitucion) / 2 - Resta(UserList(Userindex).Clase)

For i = 1 To UserList(Userindex).Stats.ELV - 1
    HPReal = HPReal + RandomNumber(Redondea(SubePromedio - 2), Fix(SubePromedio + 2))
Next

Call CalcularMana(Userindex)

UserList(Userindex).Stats.MinHit = HitReal
UserList(Userindex).Stats.MaxHit = HitReal + 1
    
UserList(Userindex).Stats.MinHP = Minimo(UserList(Userindex).Stats.MinHP, HPReal)
UserList(Userindex).Stats.MaxHP = HPReal
Call SendUserStatsBox(Userindex)

End Sub
Sub CalcularMana(Userindex As Integer)
Dim ManaReal As Integer

Select Case (UserList(Userindex).Clase)
    Case HECHICERO
        ManaReal = 100 + 2.2 * (UserList(Userindex).Stats.UserAtributos(Inteligencia) * (UserList(Userindex).Stats.ELV - 1))
    
    Case MAGO
        ManaReal = 100 + 3 * (UserList(Userindex).Stats.UserAtributos(Inteligencia) * (UserList(Userindex).Stats.ELV - 1))
        
    Case ORDEN_SAGRADA
        ManaReal = UserList(Userindex).Stats.UserAtributos(Inteligencia) * (UserList(Userindex).Stats.ELV - 1)
    
    Case CLERIGO
        ManaReal = 50 + 2 * UserList(Userindex).Stats.UserAtributos(Inteligencia) * (UserList(Userindex).Stats.ELV - 1)

    Case NATURALISTA
        ManaReal = 50 + 2 * UserList(Userindex).Stats.UserAtributos(Inteligencia) * (UserList(Userindex).Stats.ELV - 1)

    Case DRUIDA
        ManaReal = 50 + 2.1 * UserList(Userindex).Stats.UserAtributos(Inteligencia) * (UserList(Userindex).Stats.ELV - 1)
        
    Case SIGILOSO
        ManaReal = 50 + UserList(Userindex).Stats.UserAtributos(Inteligencia) * (UserList(Userindex).Stats.ELV - 1)
End Select

If ManaReal Then
    UserList(Userindex).Stats.MinMAN = Minimo(UserList(Userindex).Stats.MinMAN, ManaReal)
    UserList(Userindex).Stats.MaxMAN = ManaReal
End If

End Sub
Private Sub EnviaGenteEnNuevoRango(Userindex As Integer, ByVal nHeading As Byte)
Dim X As Integer, Y As Integer
Dim M As Integer

M = UserList(Userindex).POS.Map

Select Case nHeading

Case NORTH, SOUTH

    If nHeading = NORTH Then
        Y = UserList(Userindex).POS.Y - MinYBorder - 3
    Else
        Y = UserList(Userindex).POS.Y + MinYBorder + 3
    End If
    For X = UserList(Userindex).POS.X - MinXBorder - 2 To UserList(Userindex).POS.X + MinXBorder + 2
        If MapData(M, X, Y).Userindex Then
            Call EnviaNuevaPosUsuarioPj(Userindex, MapData(M, X, Y).Userindex)
        ElseIf MapData(M, X, Y).NpcIndex Then
            Call EnviaNuevaPosNPC(Userindex, MapData(M, X, Y).NpcIndex)
        End If
    Next
Case EAST, WEST

    If nHeading = EAST Then
        X = UserList(Userindex).POS.X + MinXBorder + 3
    Else
        X = UserList(Userindex).POS.X - MinXBorder - 3
    End If
    For Y = UserList(Userindex).POS.Y - MinYBorder - 2 To UserList(Userindex).POS.Y + MinYBorder + 2
        If MapData(M, X, Y).Userindex Then
            Call EnviaNuevaPosUsuarioPj(Userindex, MapData(M, X, Y).Userindex)
        ElseIf MapData(M, X, Y).NpcIndex Then
            Call EnviaNuevaPosNPC(Userindex, MapData(M, X, Y).NpcIndex)
        End If
    Next
End Select

End Sub
Sub CancelarSacrificio(Sacrificado As Integer)
Dim Sacrificador As Integer

Sacrificador = UserList(Sacrificado).flags.Sacrificador

UserList(Sacrificado).flags.Sacrificando = 0
UserList(Sacrificado).flags.Sacrificador = 0
UserList(Sacrificador).flags.Sacrificado = 0

Call SendData(ToIndex, Sacrificado, 0, "||¡El sacrificio fue cancelado!" & FONTTYPE_INFO)
Call SendData(ToIndex, Sacrificador, 0, "||¡El sacrificio fue cancelado!" & FONTTYPE_INFO)

End Sub
Sub MoveUserChar(Userindex As Integer, ByVal nHeading As Byte)
On Error Resume Next
Dim nPos As WorldPos



UserList(Userindex).Counters.Pasos = UserList(Userindex).Counters.Pasos + 1
    
nPos = UserList(Userindex).POS
Call HeadtoPos(nHeading, nPos)

If UserList(Userindex).flags.Sacrificado > 0 Then Call CancelarSacrificio(UserList(Userindex).flags.Sacrificado)
If UserList(Userindex).flags.Sacrificando = 1 Then Call CancelarSacrificio(Userindex)

If Not LegalPos(UserList(Userindex).POS.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(Userindex)) Then
    Call SendData(ToIndex, Userindex, 0, "PU" & UserList(Userindex).POS.X & "," & UserList(Userindex).POS.Y)
    If MapData(nPos.Map, nPos.X, nPos.Y).Userindex Then
        Call EnviaNuevaPosUsuarioPj(Userindex, MapData(nPos.Map, nPos.X, nPos.Y).Userindex)
    ElseIf MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex Then
        Call EnviaNuevaPosNPC(Userindex, MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex)
    End If
    Exit Sub
End If

Call SendData(ToPCAreaButIndexG, Userindex, UserList(Userindex).POS.Map, ("MP" & UserList(Userindex).Char.CharIndex & "," & nPos.X & "," & nPos.Y))
Call EnviaGenteEnNuevoRango(Userindex, nHeading)
MapData(UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y).Userindex = 0
UserList(Userindex).POS = nPos
UserList(Userindex).Char.Heading = nHeading
MapData(UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y).Userindex = Userindex
Call DoTileEvents(Userindex)

End Sub
Sub DesequiparItem(Userindex As Integer, Slot As Byte)

Call SendData(ToIndex, Userindex, 0, "8J" & Slot)

End Sub
Sub EquiparItem(Userindex As Integer, Slot As Byte)

Call SendData(ToIndex, Userindex, 0, "7J" & Slot)

End Sub

Sub SendUserItem(Userindex As Integer, Slot As Byte, JustAmount As Boolean)
Dim MiObj As UserOBJ
Dim Info As String

MiObj = UserList(Userindex).Invent.Object(Slot)

If MiObj.OBJIndex Then
    If Not JustAmount Then
        Info = "CSI" & Slot & "," & ObjData(MiObj.OBJIndex).Name & "," & MiObj.Amount & "," & MiObj.Equipped & "," & ObjData(MiObj.OBJIndex).GrhIndex & "," _
        & ObjData(MiObj.OBJIndex).ObjType & "," & Round(ObjData(MiObj.OBJIndex).Valor / 3)
        Select Case ObjData(MiObj.OBJIndex).ObjType
            Case OBJTYPE_WEAPON
                Info = Info & "," & ObjData(MiObj.OBJIndex).MaxHit & "," & ObjData(MiObj.OBJIndex).MinHit
            Case OBJTYPE_ARMOUR
                Info = Info & "," & ObjData(MiObj.OBJIndex).SubTipo & "," & ObjData(MiObj.OBJIndex).MaxDef & "," & ObjData(MiObj.OBJIndex).MinDef
            Case OBJTYPE_POCIONES
                Info = Info & "," & ObjData(MiObj.OBJIndex).TipoPocion & "," & ObjData(MiObj.OBJIndex).MaxModificador & "," & ObjData(MiObj.OBJIndex).MinModificador
        End Select
        Call SendData(ToIndex, Userindex, 0, Info)
    Else: Call SendData(ToIndex, Userindex, 0, "CSO" & Slot & "," & MiObj.Amount)
    End If
Else: Call SendData(ToIndex, Userindex, 0, "2H" & Slot)
End If

End Sub
Function NextOpenCharIndex() As Integer
Dim LoopC As Integer

For LoopC = 1 To LastChar + 1
    If CharList(LoopC) = 0 Then
        NextOpenCharIndex = LoopC
        NumChars = NumChars + 1
        If LoopC > LastChar Then LastChar = LoopC
        Exit Function
    End If
Next

End Function
Function NextOpenUser() As Integer
Dim LoopC As Integer
  
For LoopC = 1 To MaxUsers + 1
  If LoopC > MaxUsers Then Exit For
  If (UserList(LoopC).ConnID = -1) Then Exit For
Next
  
NextOpenUser = LoopC

End Function

Sub SendUserStatsBox(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "EST" & UserList(Userindex).Stats.MaxHP & "," & UserList(Userindex).Stats.MinHP & "," & UserList(Userindex).Stats.MaxMAN & "," & UserList(Userindex).Stats.MinMAN & "," & UserList(Userindex).Stats.MaxSta & "," & UserList(Userindex).Stats.MinSta & "," & UserList(Userindex).Stats.GLD & "," & UserList(Userindex).Stats.ELV & "," & UserList(Userindex).Stats.ELU & "," & UserList(Userindex).Stats.Exp & "," & UserList(Userindex).POS.Map)
End Sub
Sub SendUserHP(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "5A" & UserList(Userindex).Stats.MinHP)
End Sub
Sub SendUserMANA(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "5D" & UserList(Userindex).Stats.MinMAN)
End Sub
Sub SendUserMAXHP(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "8B" & UserList(Userindex).Stats.MaxHP)
End Sub
Sub SendUserMAXMANA(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "9B" & UserList(Userindex).Stats.MaxMAN)
End Sub
Sub SendUserSTA(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "5E" & UserList(Userindex).Stats.MinSta)
End Sub
Sub SendUserORO(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "5F" & UserList(Userindex).Stats.GLD)
End Sub
Sub SendUserEXP(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "5G" & UserList(Userindex).Stats.Exp)
End Sub
Sub SendUserMANASTA(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "5H" & UserList(Userindex).Stats.MinMAN & "," & UserList(Userindex).Stats.MinSta)
End Sub
Sub SendUserHPSTA(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "5I" & UserList(Userindex).Stats.MinHP & "," & UserList(Userindex).Stats.MinSta)
End Sub
Sub EnviarHambreYsed(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "EHYS" & UserList(Userindex).Stats.MaxAGU & "," & UserList(Userindex).Stats.MinAGU & "," & UserList(Userindex).Stats.MaxHam & "," & UserList(Userindex).Stats.MinHam)
End Sub
Sub EnviarHyS(Userindex As Integer)
Call SendData(ToIndex, Userindex, 0, "5J" & UserList(Userindex).Stats.MinAGU & "," & UserList(Userindex).Stats.MinHam)
End Sub

Sub SendUserSTAtsTxt(ByVal sendIndex As Integer, Userindex As Integer)

Call SendData(ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(Userindex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Nivel: " & UserList(Userindex).Stats.ELV & "  EXP: " & UserList(Userindex).Stats.Exp & "/" & UserList(Userindex).Stats.ELU & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Vitalidad: " & UserList(Userindex).Stats.FIT & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Salud: " & UserList(Userindex).Stats.MinHP & "/" & UserList(Userindex).Stats.MaxHP & "  Mana: " & UserList(Userindex).Stats.MinMAN & "/" & UserList(Userindex).Stats.MaxMAN & "  Vitalidad: " & UserList(Userindex).Stats.MinSta & "/" & UserList(Userindex).Stats.MaxSta & FONTTYPE_INFO)

If UserList(Userindex).Invent.WeaponEqpObjIndex Then
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(Userindex).Stats.MinHit & "/" & UserList(Userindex).Stats.MaxHit & " (" & ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MinHit & "/" & ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MaxHit & ")" & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(Userindex).Stats.MinHit & "/" & UserList(Userindex).Stats.MaxHit & FONTTYPE_INFO)
End If

Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MinDef + 2 * Buleano(UserList(Userindex).Clase = GUERRERO And UserList(Userindex).Recompensas(2) = 2) & "/" & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MaxDef + 2 * Buleano(UserList(Userindex).Clase = GUERRERO And UserList(Userindex).Recompensas(2) = 2) & FONTTYPE_INFO)

If UserList(Userindex).Invent.CascoEqpObjIndex Then
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If

If UserList(Userindex).Invent.EscudoEqpObjIndex Then
    Call SendData(ToIndex, sendIndex, 0, "||(ESCUDO) Defensa extra: " & ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).MinDef & " / " & ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).MaxDef & FONTTYPE_INFO)
End If

If Len(UserList(Userindex).GuildInfo.GuildName) > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||Clan: " & UserList(Userindex).GuildInfo.GuildName & FONTTYPE_INFO)
    If UserList(Userindex).GuildInfo.EsGuildLeader = 1 Then
       If UserList(Userindex).GuildInfo.ClanFundado = UserList(Userindex).GuildInfo.GuildName Then
            Call SendData(ToIndex, sendIndex, 0, "||Status: " & "Fundador/Lider" & FONTTYPE_INFO)
       Else
            Call SendData(ToIndex, sendIndex, 0, "||Status: " & "Lider" & FONTTYPE_INFO)
       End If
    Else
        Call SendData(ToIndex, sendIndex, 0, "||Status: " & UserList(Userindex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    Call SendData(ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(Userindex).GuildInfo.GuildPoints & FONTTYPE_INFO)
End If

Call SendData(ToIndex, sendIndex, 0, "||Oro: " & UserList(Userindex).Stats.GLD & "  Posicion: " & UserList(Userindex).POS.X & "," & UserList(Userindex).POS.Y & " en mapa " & UserList(Userindex).POS.Map & FONTTYPE_INFO)

Call SendData(ToIndex, sendIndex, 0, "||Ciudadanos matados: " & UserList(Userindex).Faccion.Matados(Real) & " / Criminales matados: " & UserList(Userindex).Faccion.Matados(Caos) & " / Neutrales matados: " & UserList(Userindex).Faccion.Matados(Neutral) & FONTTYPE_INFO)

End Sub
Sub SendUserInvTxt(ByVal sendIndex As Integer, Userindex As Integer)
On Error Resume Next
Dim j As Byte

Call SendData(ToIndex, sendIndex, 0, "||" & UserList(Userindex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & UserList(Userindex).Invent.NroItems & " objetos." & FONTTYPE_INFO)

For j = 1 To MAX_INVENTORY_SLOTS
    If UserList(Userindex).Invent.Object(j).OBJIndex Then
        Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(Userindex).Invent.Object(j).OBJIndex).Name & " Cantidad:" & UserList(Userindex).Invent.Object(j).Amount & FONTTYPE_INFO)
    End If
Next

End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, Userindex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(Userindex).Name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(Userindex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
End Sub
Sub UpdateFuerzaYAg(Userindex As Integer)
Dim Fue As Integer
Dim Agi As Integer

Fue = UserList(Userindex).Stats.UserAtributos(fuerza)
If Fue = UserList(Userindex).Stats.UserAtributosBackUP(fuerza) Then Fue = 0

Agi = UserList(Userindex).Stats.UserAtributos(Agilidad)
If Agi = UserList(Userindex).Stats.UserAtributosBackUP(Agilidad) Then Agi = 0

Call SendData(ToIndex, Userindex, 0, "EIFYA" & Fue & "," & Agi)

End Sub
Sub UpdateUserMap(Userindex As Integer)
On Error GoTo ErrorHandler
Dim TempChar As Integer
Dim Map As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Integer

Map = UserList(Userindex).POS.Map
Call SendData(ToIndex, Userindex, 0, "ET")


For i = 1 To MapInfo(Map).NumUsers
    TempChar = MapInfo(Map).Userindex(i)
    Call MakeUserChar(ToIndex, Userindex, 0, TempChar, Map, UserList(TempChar).POS.X, UserList(TempChar).POS.Y)
Next


For i = 1 To LastNPC
    If Npclist(i).flags.NPCActive And UserList(Userindex).POS.Map = Npclist(i).POS.Map Then
        Call MakeNPCChar(ToIndex, Userindex, 0, i, Map, Npclist(i).POS.X, Npclist(i).POS.Y)
    End If
Next


For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(Map, X, Y).OBJInfo.OBJIndex Then
            If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType <> OBJTYPE_ARBOLES Or MapData(Map, X, Y).trigger = 2 Then
                If Y >= 40 Then
                    Y = Y
                End If
                
                Call MakeObj(ToIndex, Userindex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                
                If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
                    Call Bloquear(ToIndex, Userindex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                    Call Bloquear(ToIndex, Userindex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                End If
            End If
        End If
    Next
Next

Exit Sub
ErrorHandler:
    Call LogError("Error en el sub.UpdateUserMap. Mapa: " & Map & "-" & X & "-" & Y)

End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function
Function EsMascotaCiudadano(NpcIndex As Integer, Userindex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser Then
    EsMascotaCiudadano = UserList(Userindex).Faccion.Bando = Real
    If EsMascotaCiudadano Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "F0" & UserList(Userindex).Name)
End If

End Function
Function EsMascotaCriminal(NpcIndex As Integer, Userindex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser Then
    EsMascotaCriminal = Not UserList(Userindex).Faccion.Bando = Caos
    If EsMascotaCriminal Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "F0" & UserList(Userindex).Name)
End If

End Function
Sub NpcAtacado(NpcIndex As Integer, Userindex As Integer)

Npclist(NpcIndex).flags.AttackedBy = Userindex

If Npclist(NpcIndex).MaestroUser Then Call AllMascotasAtacanUser(Userindex, Npclist(NpcIndex).MaestroUser)
If Npclist(NpcIndex).flags.Faccion <> Neutral Then
    If UserList(Userindex).Faccion.Ataco(Npclist(NpcIndex).flags.Faccion) = 0 Then UserList(Userindex).Faccion.Ataco(Npclist(NpcIndex).flags.Faccion) = 2
End If

Npclist(NpcIndex).Movement = NPCDEFENSA
Npclist(NpcIndex).Hostile = 1

End Sub
Function PuedeApuñalar(Userindex As Integer) As Boolean

If UserList(Userindex).Invent.WeaponEqpObjIndex Then PuedeApuñalar = ((UserList(Userindex).Stats.UserSkills(Apuñalar) >= MIN_APUÑALAR) And (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Apuñala = 1)) Or ((UserList(Userindex).Clase = ASESINO) And (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Apuñala = 1))

End Function
Sub SubirSkill(Userindex As Integer, Skill As Integer, Optional Prob As Integer)
On Error GoTo errhandler

If UserList(Userindex).flags.Hambre = 1 Or UserList(Userindex).flags.Sed = 1 Then Exit Sub

If Prob = 0 Then
    If UserList(Userindex).Stats.ELV <= 3 Then
        Prob = 20
    ElseIf UserList(Userindex).Stats.ELV > 3 _
        And UserList(Userindex).Stats.ELV < 6 Then
        Prob = 25
    ElseIf UserList(Userindex).Stats.ELV >= 6 _
        And UserList(Userindex).Stats.ELV < 10 Then
        Prob = 30
    ElseIf UserList(Userindex).Stats.ELV >= 10 _
        And UserList(Userindex).Stats.ELV < 20 Then
        Prob = 35
    Else
        Prob = 40
    End If
End If

If UserList(Userindex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

If Int(RandomNumber(1, Prob)) = 2 And UserList(Userindex).Stats.UserSkills(Skill) < LevelSkill(UserList(Userindex).Stats.ELV).LevelValue Then
    Call AddtoVar(UserList(Userindex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
    Call SendData(ToIndex, Userindex, 0, "G0" & SkillsNames(Skill) & "," & UserList(Userindex).Stats.UserSkills(Skill))
    Call AddtoVar(UserList(Userindex).Stats.Exp, 50, MAXEXP)
    Call SendData(ToIndex, Userindex, 0, "EX" & 50)
    Call SendUserEXP(Userindex)
    Call CheckUserLevel(Userindex)
End If
Exit Sub

errhandler:
    Call LogError("Error en SubirSkill: " & Err.Description & "-" & UserList(Userindex).Name & "-" & SkillsNames(Skill))
End Sub
Sub BajarInvisible(Userindex As Integer)

If UserList(Userindex).Stats.ELV >= 34 Or UserList(Userindex).flags.GolpeoInvi Then
    Call QuitarInvisible(Userindex)
Else: UserList(Userindex).flags.GolpeoInvi = 1
End If

End Sub
Sub QuitarInvisible(Userindex As Integer)

UserList(Userindex).Counters.Invisibilidad = 0
UserList(Userindex).flags.Invisible = 0
UserList(Userindex).flags.GolpeoInvi = 0
UserList(Userindex).flags.Oculto = 0
Call SendData(ToMap, 0, UserList(Userindex).POS.Map, ("V3" & UserList(Userindex).Char.CharIndex & ",0"))

End Sub
Sub UserDie(Userindex As Integer)
On Error GoTo ErrorHandler

Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "TW" & SND_USERMUERTE)

If UserList(Userindex).flags.Montado = 1 Then Desmontar (Userindex)

Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "QDL" & UserList(Userindex).Char.CharIndex)

UserList(Userindex).Stats.MinHP = 0
UserList(Userindex).flags.AtacadoPorNpc = 0
UserList(Userindex).flags.AtacadoPorUser = 0
UserList(Userindex).flags.Envenenado = 0
UserList(Userindex).flags.Muerto = 1

Dim aN As Integer

aN = UserList(Userindex).flags.AtacadoPorNpc

If aN Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = 0
End If

If UserList(Userindex).flags.Paralizado Then
    Call SendData(ToIndex, Userindex, 0, "P8")
    UserList(Userindex).flags.Paralizado = 0
End If

If UserList(Userindex).flags.Trabajando Then Call SacarModoTrabajo(Userindex)

If UserList(Userindex).flags.Invisible And UserList(Userindex).flags.AdminInvisible = 0 Then
    Call QuitarInvisible(Userindex)
End If

If UserList(Userindex).flags.Ceguera = 1 Then
  UserList(Userindex).Counters.Ceguera = 0
  UserList(Userindex).flags.Ceguera = 0
  Call SendData(ToMap, 0, UserList(Userindex).POS.Map, "NSEGUE")
End If

If UserList(Userindex).flags.Estupidez = 1 Then
  UserList(Userindex).Counters.Estupidez = 0
  UserList(Userindex).flags.Estupidez = 0
  Call SendData(ToMap, 0, UserList(Userindex).POS.Map, "NESTUP")
End If

If UserList(Userindex).flags.Descansar Then
    UserList(Userindex).flags.Descansar = False
    Call SendData(ToIndex, Userindex, 0, "DOK")
End If

If UserList(Userindex).flags.Meditando Then
    UserList(Userindex).flags.Meditando = False
    Call SendData(ToIndex, Userindex, 0, "MEDOK")
End If

If UserList(Userindex).POS.Map <> 190 Then
    If Not EsNewbie(Userindex) Then
        Call TirarTodo(Userindex)
    Else: Call TirarTodosLosItemsNoNewbies(Userindex)
    End If
End If

If UserList(Userindex).Invent.ArmourEqpObjIndex Then Call Desequipar(Userindex, UserList(Userindex).Invent.ArmourEqpSlot)
If UserList(Userindex).Invent.WeaponEqpObjIndex Then Call Desequipar(Userindex, UserList(Userindex).Invent.WeaponEqpSlot)
If UserList(Userindex).Invent.EscudoEqpObjIndex Then Call Desequipar(Userindex, UserList(Userindex).Invent.EscudoEqpSlot)
If UserList(Userindex).Invent.CascoEqpObjIndex Then Call Desequipar(Userindex, UserList(Userindex).Invent.CascoEqpSlot)
If UserList(Userindex).Invent.HerramientaEqpObjIndex Then Call Desequipar(Userindex, UserList(Userindex).Invent.HerramientaEqpslot)
If UserList(Userindex).Invent.MunicionEqpObjIndex Then Call Desequipar(Userindex, UserList(Userindex).Invent.MunicionEqpSlot)

If UserList(Userindex).Char.loops = LoopAdEternum Then
    UserList(Userindex).Char.FX = 0
    UserList(Userindex).Char.loops = 0
End If

If UserList(Userindex).flags.Navegando = 0 Then
    UserList(Userindex).Char.Body = iCuerpoMuerto
    UserList(Userindex).Char.Head = iCabezaMuerto
    UserList(Userindex).Char.ShieldAnim = NingunEscudo
    UserList(Userindex).Char.WeaponAnim = NingunArma
    UserList(Userindex).Char.CascoAnim = NingunCasco
Else
    UserList(Userindex).Char.Body = iFragataFantasmal
End If

Dim i As Integer
For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(Userindex).flags.Quest)
    If UserList(Userindex).MascotasIndex(i) Then
           If Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia Then
                Call MuereNpc(UserList(Userindex).MascotasIndex(i), 0)
           Else
                Npclist(UserList(Userindex).MascotasIndex(i)).MaestroUser = 0
                Npclist(UserList(Userindex).MascotasIndex(i)).Movement = Npclist(UserList(Userindex).MascotasIndex(i)).flags.OldMovement
                Npclist(UserList(Userindex).MascotasIndex(i)).Hostile = Npclist(UserList(Userindex).MascotasIndex(i)).flags.OldHostil
                UserList(Userindex).MascotasIndex(i) = 0
                UserList(Userindex).MascotasType(i) = 0
           End If
    End If
    
Next

If UserList(Userindex).POS.Map <> 190 Then UserList(Userindex).Stats.VecesMurioUsuario = UserList(Userindex).Stats.VecesMurioUsuario + 1

UserList(Userindex).NroMascotas = 0

Call ChangeUserChar(ToMap, 0, UserList(Userindex).POS.Map, val(Userindex), UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
If PuedeDestrabarse(Userindex) Then Call SendData(ToIndex, Userindex, 0, "||Estás encerrado, para destrabarte presiona la tecla Z." & FONTTYPE_INFO)
Call SendUserStatsBox(Userindex)

Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE")

End Sub
Sub ContarMuerte(Muerto As Integer, Atacante As Integer)
If EsNewbie(Muerto) Then Exit Sub

If UserList(Muerto).POS.Map = 190 Then Exit Sub

If UserList(Atacante).flags.LastMatado(UserList(Muerto).Faccion.Bando) <> UCase$(UserList(Muerto).Name) Then
    UserList(Atacante).flags.LastMatado(UserList(Muerto).Faccion.Bando) = UCase$(UserList(Muerto).Name)
    Call AddtoVar(UserList(Atacante).Faccion.Matados(UserList(Muerto).Faccion.Bando), 1, 65000)
End If

End Sub

Sub Tilelibre(POS As WorldPos, nPos As WorldPos)


Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
hayobj = False
nPos.Map = POS.Map

Do While Not LegalPos(POS.Map, nPos.X, nPos.Y) Or hayobj
    
    If LoopC > 15 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = POS.Y - LoopC To POS.Y + LoopC
        For tX = POS.X - LoopC To POS.X + LoopC
        
            If LegalPos(nPos.Map, tX, tY) Then
               hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.OBJIndex > 0)
               If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                     nPos.X = tX
                     nPos.Y = tY
                     tX = POS.X + LoopC
                     tY = POS.Y + LoopC
                End If
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
Sub AgregarAUsersPorMapa(Userindex As Integer)


MapInfo(UserList(Userindex).POS.Map).NumUsers = MapInfo(UserList(Userindex).POS.Map).NumUsers + 1
If MapInfo(UserList(Userindex).POS.Map).NumUsers < 0 Then MapInfo(UserList(Userindex).POS.Map).NumUsers = 0

If MapInfo(UserList(Userindex).POS.Map).NumUsers = 1 Then
    ReDim MapInfo(UserList(Userindex).POS.Map).Userindex(1 To 1)
Else
    
    ReDim Preserve MapInfo(UserList(Userindex).POS.Map).Userindex(1 To MapInfo(UserList(Userindex).POS.Map).NumUsers)
End If


MapInfo(UserList(Userindex).POS.Map).Userindex(MapInfo(UserList(Userindex).POS.Map).NumUsers) = Userindex
    
End Sub
Sub QuitarDeUsersPorMapa(Userindex As Integer)


MapInfo(UserList(Userindex).POS.Map).NumUsers = MapInfo(UserList(Userindex).POS.Map).NumUsers - 1
If MapInfo(UserList(Userindex).POS.Map).NumUsers < 0 Then MapInfo(UserList(Userindex).POS.Map).NumUsers = 0

If MapInfo(UserList(Userindex).POS.Map).NumUsers Then
    Dim i As Integer
        
    For i = 1 To MapInfo(UserList(Userindex).POS.Map).NumUsers + 1
        
        If MapInfo(UserList(Userindex).POS.Map).Userindex(i) = Userindex Then Exit For
    Next
    
    For i = i To MapInfo(UserList(Userindex).POS.Map).NumUsers
        
        MapInfo(UserList(Userindex).POS.Map).Userindex(i) = MapInfo(UserList(Userindex).POS.Map).Userindex(i + 1)
    Next
    
    ReDim Preserve MapInfo(UserList(Userindex).POS.Map).Userindex(1 To MapInfo(UserList(Userindex).POS.Map).NumUsers)
Else
    ReDim MapInfo(UserList(Userindex).POS.Map).Userindex(0)
End If
    
End Sub
Sub WarpUserChar(Userindex As Integer, Map As Integer, X As Integer, Y As Integer, Optional FX As Boolean = False)

Call SendData(ToMap, 0, UserList(Userindex).POS.Map, "QDL" & UserList(Userindex).Char.CharIndex)
Call SendData(ToIndex, Userindex, UserList(Userindex).POS.Map, "QTDL")

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

UserList(Userindex).Counters.Protegido = 2
UserList(Userindex).flags.Protegido = 3

OldMap = UserList(Userindex).POS.Map
OldX = UserList(Userindex).POS.X
OldY = UserList(Userindex).POS.Y

Call EraseUserChar(ToMap, 0, OldMap, Userindex)

UserList(Userindex).POS.X = X
UserList(Userindex).POS.Y = Y

If OldMap = Map Then
    Call MakeUserChar(ToMap, 0, UserList(Userindex).POS.Map, Userindex, UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y)
    Call SendData(ToIndex, Userindex, 0, "IP" & UserList(Userindex).Char.CharIndex)
Else
    Call QuitarDeUsersPorMapa(Userindex)
    UserList(Userindex).POS.Map = Map
    Call AgregarAUsersPorMapa(Userindex)
     
    Call SendData(ToIndex, Userindex, 0, "CM" & UserList(Userindex).POS.Map & "," & MapInfo(UserList(Userindex).POS.Map).MapVersion & "," & MapInfo(UserList(Userindex).POS.Map).Name & "," & MapInfo(UserList(Userindex).POS.Map).TopPunto & "," & MapInfo(UserList(Userindex).POS.Map).LeftPunto)
    If MapInfo(Map).Music <> MapInfo(OldMap).Music Then Call SendData(ToIndex, Userindex, 0, "TM" & MapInfo(Map).Music)

    Call MakeUserChar(ToMap, 0, UserList(Userindex).POS.Map, Userindex, UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y)
    Call SendData(ToIndex, Userindex, 0, "IP" & UserList(Userindex).Char.CharIndex)
End If

Call UpdateUserMap(Userindex)

If FX And UserList(Userindex).flags.AdminInvisible = 0 And Not UserList(Userindex).flags.Meditando Then
    Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "TW" & SND_WARP)
    Call SendData(ToPCArea, Userindex, UserList(Userindex).POS.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & FXWARP & "," & 0)
End If
Dim i As Integer

For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(Userindex).flags.Quest)
    If UserList(Userindex).MascotasIndex(i) Then
        If Npclist(UserList(Userindex).MascotasIndex(i)).flags.NPCActive Then
            Call QuitarNPC(UserList(Userindex).MascotasIndex(i))
        End If
    End If
Next

End Sub
Sub WarpMascotas(Userindex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer

NroPets = UserList(Userindex).NroMascotas

For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(Userindex).flags.Quest)
    If UserList(Userindex).MascotasIndex(i) Then
        PetRespawn(i) = Npclist(UserList(Userindex).MascotasIndex(i)).flags.Respawn = 0
        If PetRespawn(i) Then
            PetTypes(i) = UserList(Userindex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(Userindex).MascotasIndex(i))
        Else
            PetTypes(i) = UserList(Userindex).MascotasType(i)
            PetTiempoDeVida(i) = 1
            Call QuitarNPC(UserList(Userindex).MascotasIndex(i))
        End If
    End If
Next

For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(Userindex).flags.Quest)
    If PetTypes(i) Then
        UserList(Userindex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(Userindex).POS, False, PetRespawn(i))
        UserList(Userindex).MascotasType(i) = PetTypes(i)
        
        If UserList(Userindex).MascotasIndex(i) = MAXNPCS Then
                UserList(Userindex).MascotasIndex(i) = 0
                UserList(Userindex).MascotasType(i) = 0
                If UserList(Userindex).NroMascotas Then UserList(Userindex).NroMascotas = UserList(Userindex).NroMascotas - 1
                Exit Sub
        End If
        Npclist(UserList(Userindex).MascotasIndex(i)).MaestroUser = Userindex
        Npclist(UserList(Userindex).MascotasIndex(i)).Movement = SIGUE_AMO
        Npclist(UserList(Userindex).MascotasIndex(i)).Target = 0
        Npclist(UserList(Userindex).MascotasIndex(i)).TargetNpc = 0
        Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
        Call QuitarNPCDeLista(Npclist(UserList(Userindex).MascotasIndex(i)).Numero, UserList(Userindex).POS.Map)
        Call FollowAmo(UserList(Userindex).MascotasIndex(i))
    End If
Next

UserList(Userindex).NroMascotas = NroPets

End Sub
Sub Cerrar_Usuario(Userindex As Integer)

If UserList(Userindex).flags.UserLogged And Not UserList(Userindex).Counters.Saliendo Then
    UserList(Userindex).Counters.Saliendo = True
    UserList(Userindex).Counters.Salir = Timer - 8 * Buleano(UserList(Userindex).Clase = PIRATA And UserList(Userindex).Recompensas(3) = 2)
    Call SendData(ToIndex, Userindex, 0, "1Z" & IntervaloCerrarConexion - 8 * Buleano(UserList(Userindex).Clase = PIRATA And UserList(Userindex).Recompensas(3) = 2))
End If
    
End Sub

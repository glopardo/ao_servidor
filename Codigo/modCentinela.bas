Attribute VB_Name = "modCentinela"
'*****************************************************************
'modCentinela.bas - ImperiumAO - v1.2 - www.imperiumao.com.ar
'
'Funciónes de control para usuarios que se encuentran trabajando
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

'*****************************************************************
'Augusto Rando(barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private Const NPC_CENTINELA_TIERRA = 158 'Índice del NPC en el .dat
Private Const NPC_CENTINELA_AGUA = 159 'Ídem anterior, pero en mapas de agua
Private Const MINIMO_TRABAJO = 250 'Mínimo del contador de trabajo del usuario
Private CentinelaCharIndex As Integer 'Índice del NPC en el servidor

Private Const TIEMPO_INICIAL = 2 'Tiempo inicial en minutos

Private Type tCentinela
    RevisandoUserIndex As Integer '¿Qué índice revisamos?
    TiempoRestante As Integer '¿Cuántos minutos le quedan al usuario?
    Clave As Integer 'Clave que debe escribir
End Type

Private Centinela As tCentinela

Public Sub PasarMinutoCentinela()
'############################################################
'Control del timer. Llamado cada un minuto.
'############################################################

If Centinela.RevisandoUserIndex = 0 Then
    Call GoToNextWorkingChar
Else
    Centinela.TiempoRestante = Centinela.TiempoRestante - 1
    
    If Centinela.TiempoRestante = 0 Then
        Call CentinelaFinalCheck
    Else
    
    End If
End If

End Sub

Private Sub GoToNextWorkingChar()
'############################################################
'Va al siguiente usuario que se encuentre trabajando
'############################################################

Dim loopc As Integer

For loopc = 1 To LastUser
    If (UserList(loopc).Name <> "") And UserList(loopc).Counters.Trabajando >= MINIMO_TRABAJO Then
        If UserList(loopc).FLAGS.CentinelaOK = False Then
            Call WarpCentinela(loopc)
            Exit Sub
        End If
    End If
Next loopc

End Sub

Private Sub CentinelaFinalCheck()
'############################################################
'Al finalizar el tiempo, se retira y realiza la acción
'pertinente dependiendo del caso
'############################################################

On Error GoTo Error_Handler

'Nuevo: tiene que haber trabajado y estar online
If UserList(Centinela.RevisandoUserIndex).FLAGS.CentinelaOK = False Then
    If UserList(Centinela.RevisandoUserIndex).Counters.Trabajando >= MINIMO_TRABAJO And _
    UserList(Centinela.RevisandoUserIndex).ConnID <> -1 And UserList(Centinela.RevisandoUserIndex).FLAGS.Muerto = 0 Then
        Call LogBan(UserList(Centinela.RevisandoUserIndex).Name, "Centinela", "Uso de macro inasistido", 15)
        UserList(Centinela.RevisandoUserIndex).FLAGS.Ban = 1
        Call CloseSocket(Centinela.RevisandoUserIndex)
    End If
End If
  
Centinela.Clave = 0
Centinela.TiempoRestante = 0
Centinela.RevisandoUserIndex = 0
Call QuitarNPC(CentinelaCharIndex)

Exit Sub

Error_Handler:
    Centinela.Clave = 0
    Centinela.TiempoRestante = 0
    Centinela.RevisandoUserIndex = 0
    Call QuitarNPC(CentinelaCharIndex)
    Call LogError("CentinelaFinalCheck: " & Err.Description & " - " & Err.Number)

End Sub

Public Sub CentinelaCheckClave(ByVal UserIndex As Integer, ByVal Clave As Integer)
'############################################################
'Corrobora la clave que le envia el usuario
'############################################################

If Centinela.RevisandoUserIndex <> UserIndex Then Exit Sub

If Clave = Centinela.Clave Then
    UserList(Centinela.RevisandoUserIndex).FLAGS.CentinelaOK = True
    
    If UserList(UserIndex).FLAGS.Paralizado = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "FREZOK")
        UserList(UserIndex).FLAGS.Paralizado = 0
    End If
    
    Call SendData(ToIndex, Centinela.RevisandoUserIndex, 0, "||" & vbWhite & "°" & "¡Muchas gracias " & UserList(Centinela.RevisandoUserIndex).Name & "! Espero no haber sido una molestia" & "°" & str(Npclist(CentinelaCharIndex).Char.CharIndex))
Else
    Call SendData(ToIndex, Centinela.RevisandoUserIndex, 0, "||" & vbWhite & "°" & "¡La clave que te he dicho no es esa, " & "escríbe /CENTINELA " & Centinela.Clave & " rápido!" & "°" & str(Npclist(CentinelaCharIndex).Char.CharIndex))
End If

End Sub

Public Sub CentinelaAI(ByVal NpcIndex As Integer)
'############################################################
'Procedimiento para el control de la IA del NPC Centinela
'############################################################

Dim UI As Integer
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
           If UI > 0 Then
                If Centinela.RevisandoUserIndex = UI And UserList(UI).FLAGS.CentinelaOK = False Then
                   If UserList(UI).FLAGS.Muerto = 0 Then
                        
                        If Npclist(NpcIndex).FLAGS.LanzaSpells > 0 And Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) > 5 Then
                            Call NpcLanzaUnSpell(NpcIndex, UI)
                            Call SendData(ToIndex, UI, 0, "||" & vbRed & "°" & "¿A dónde crees que vas muchacho? Debes contestarme lo que te he preguntado." & "°" & str(Npclist(NpcIndex).Char.CharIndex))
                        End If
                                                
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
           End If
        End If
    Next X
Next Y

End Sub

'############################################################
'Llamado desde CloseUser. Permite saber si el usuario que está
'siendo revisado cierra el juego, y tomar acciones.
'############################################################
Public Sub CentinelaLogOut(ByVal UserIndex As Integer)

If Centinela.RevisandoUserIndex = UserIndex Then
    Centinela.RevisandoUserIndex = 0
    Centinela.TiempoRestante = 0
    Centinela.Clave = 0
    Call QuitarNPC(CentinelaCharIndex)
End If

End Sub

Public Sub ResetCentinelaInfo()
'############################################################
'Cada determinada cantidad de tiempo, volvemos a revisar
'############################################################

Dim loopc As Integer

For loopc = 1 To LastUser
    If (UserList(loopc).Name <> "" And loopc <> Centinela.RevisandoUserIndex) Then
        UserList(loopc).FLAGS.CentinelaOK = False
    End If
Next loopc

End Sub

Public Sub CentinelaSendClave(ByVal UserIndex As Integer)
'############################################################
'Enviamos al usuario la clave vía el personaje centinela
'############################################################

If UserIndex = Centinela.RevisandoUserIndex Then
    
    If UserList(UserIndex).FLAGS.CentinelaOK = False Then
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "¡La clave que te he dicho es " & "/CENTINELA " & Centinela.Clave & " escríbelo rápido!" & "°" & str(Npclist(CentinelaCharIndex).Char.CharIndex))
    Else
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Te agradezco, pero ya me has respondido. Me retiraré pronto." & "°" & str(Npclist(CentinelaCharIndex).Char.CharIndex))
    End If
    
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No es a ti a quien estoy revisando, ¿no ves?" & "°" & str(Npclist(CentinelaCharIndex).Char.CharIndex))
End If

End Sub

Private Sub WarpCentinela(UserIndex As Integer)
'############################################################
'Inciamos la revisión del usuario UserIndex
'############################################################

Centinela.RevisandoUserIndex = UserIndex
Centinela.TiempoRestante = TIEMPO_INICIAL
Centinela.Clave = RandomNumber(1, 2000)

If HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
    CentinelaCharIndex = SpawnNpc(NPC_CENTINELA_AGUA, UserList(UserIndex).Pos, True, False)
Else
    CentinelaCharIndex = SpawnNpc(NPC_CENTINELA_TIERRA, UserList(UserIndex).Pos, True, False)
End If

If CentinelaCharIndex >= MAXNPCS Then
    Centinela.RevisandoUserIndex = 0
    Centinela.TiempoRestante = 0
    Centinela.Clave = 0
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Saludos " & UserList(UserIndex).Name & ", soy el Centinela de estas tierras. Me gustaría que escribas /CENTINELA " & Centinela.Clave & " en no despues de dos minutos." & "°" & str(Npclist(CentinelaCharIndex).Char.CharIndex))
End If

End Sub

Private Sub WarnUser()
'############################################################
' Se advierte al usuario Centinela.RevisandoUserIndex
'############################################################

Call SendData(ToIndex, Centinela.RevisandoUserIndex, 0, "||" & vbRed & "°¡" & UserList(Centinela.RevisandoUserIndex).Name & ", tienes un minuto más para responder! Debes escribir /CENTINELA " & Centinela.Clave & "." & "°" & str(Npclist(CentinelaCharIndex).Char.CharIndex))
Call SendData(ToIndex, Centinela.RevisandoUserIndex, 0, "||Servidor> " & "¡" & UserList(Centinela.RevisandoUserIndex).Name & ", tienes un minuto más para responder! Debes escribir /CENTINELA " & Centinela.Clave & "." & FONTTYPE_SERVER)

End Sub

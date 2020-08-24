Attribute VB_Name = "modTorneo"
Public Sub TerminarDueloTorneo(Ganador As Integer, Perdedor As Integer)
    Call Logear("test", "TerminarDueloTorneo")
    'Mandar al ganador al mapa de espera
    Call WarpUserChar(Ganador, 200, 50, 44, False)
    'Mandar al perdedor al mapa de dropeo de items y luego a Ulla
    Call WarpUserChar(Perdedor, 86, 50, 50, False)
    Call WarpUserChar(Perdedor, 1, 50, 50, False)
    
    UserList(Ganador).Torneo.EstaDueleando = False
    UserList(Perdedor).Torneo.EstaDueleando = False
    
    Call SendData(ToAll, 0, 0, "||Muere " & UserList(Perdedor).Name & ", pasa a la siguiente ronda " & UserList(Ganador).Name & FONTTYPE_TALK)
    Call SendData(ToAdmins, 0, 0, "FINDT" & UserList(Ganador).Name & "|" & UserList(Perdedor).Name)
End Sub

Public Sub ComenzarDueloTorneo(Jugador1 As Integer, Jugador2 As Integer)
    UserList(Jugador1).Torneo.EstaDueleando = True
    UserList(Jugador2).Torneo.EstaDueleando = True
    
    UserList(Jugador1).Torneo.Contrincante = Jugador2
    UserList(Jugador2).Torneo.Contrincante = Jugador1
End Sub

Attribute VB_Name = "modRetos"
Public Sub TerminarDuelo(Ganador As Integer, Perdedor As Integer)
    UserList(Ganador).Stats.GLD = UserList(Ganador).Stats.GLD + 100000
    Call SendUserORO(Ganador)
    
    Call WarpUserChar(Ganador, 160, 42, 40, False)
    Call WarpUserChar(Perdedor, 160, 43, 40, False)
    
    If UserList(Ganador).Reto.Ring = 1 Then Ring1.Ocupado = False
    If UserList(Ganador).Reto.Ring = 2 Then Ring2.Ocupado = False
    
    UserList(Ganador).Reto.EstaDueleando = False
    UserList(Ganador).Reto.Ring = 0
    UserList(Perdedor).Reto.EstaDueleando = False
    UserList(Perdedor).Reto.Ring = 0
    
    Call SendData(ToAll, 0, 0, "||" & UserList(Ganador).Name & " ganó el reto contra " & UserList(Perdedor).Name & FONTTYPE_TALK)
End Sub
 
Public Sub DesconectarDuelo(Ganador As Integer, Perdedor As Integer)
    UserList(Ganador).Stats.GLD = UserList(Ganador).Stats.GLD + 50000
    UserList(Perdedor).Stats.GLD = UserList(Perdedor).Stats.GLD + 50000
    Call SendUserORO(Ganador)
    Call SendUserORO(Perdedor)
    
    Call WarpUserChar(Ganador, 160, 42, 40, False)
    Call WarpUserChar(Perdedor, 160, 43, 40, False)
    
    If UserList(Ganador).Reto.Ring = 1 Then Ring1.Ocupado = False
    If UserList(Ganador).Reto.Ring = 2 Then Ring2.Ocupado = False
    
    UserList(Ganador).Reto.EstaDueleando = False
    UserList(Ganador).Reto.Ring = 0
    UserList(Perdedor).Reto.EstaDueleando = False
    UserList(Perdedor).Reto.Ring = 0
    
    Call SendData(ToAll, 0, 0, "||" & UserList(Ganador).Name & " ganó el reto por desconexión de " & UserList(Perdedor).Name & FONTTYPE_TALK)
End Sub
 
Public Sub ComenzarDuelo(Retador As Integer, Retado As Integer)
    UserList(Retador).Stats.GLD = UserList(Retador).Stats.GLD - 50000
    UserList(Retado).Stats.GLD = UserList(Retado).Stats.GLD - 50000
    Call SendUserORO(Retador)
    Call SendUserORO(Retado)
    
    If Ring1.Ocupado Then
        Ring2.Ocupado = True
        UserList(Retador).Reto.Ring = 2
        UserList(Retado).Reto.Ring = 2
        Call WarpUserChar(Retador, 178, 22, 45, False)
        Call WarpUserChar(Retado, 178, 39, 60, False)
    Else
        Ring1.Ocupado = True
        UserList(Retador).Reto.Ring = 1
        UserList(Retado).Reto.Ring = 1
        Call WarpUserChar(Retador, 178, 21, 19, False)
        Call WarpUserChar(Retado, 178, 33, 29, False)
    End If
    
    UserList(Retador).Reto.EstaDueleando = True
    UserList(Retado).Reto.EstaDueleando = True
    UserList(Retador).Reto.EsperandoDuelo = False
    Call SendData(ToAll, 0, 0, "||" & "Comenzó el reto " & UserList(Retador).Name & " Vs. " & UserList(Retado).Name & FONTTYPE_TALK)
End Sub

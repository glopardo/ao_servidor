Attribute VB_Name = "AwsApi"
'Argentum Online 0.11.20
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
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

Const WSAPI_CREAR_LABEL = True

Private Const SD_RECEIVE As Long = &H0
Private Const SD_SEND As Long = &H1
Private Const SD_BOTH As Long = &H2


Private Const MAX_TIEMPOIDLE_COLALLENA = 1 'minutos
Private Const MAX_COLASALIDA_COUNT = 800

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, X As Long, Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)

'====================================================================================
'====================================================================================
'Esto es para agilizar la busqueda del slot a partir de un socket dado,
'sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.

Public Type tSockCache
    Sock As Long
    Slot As Long
End Type

'Public WSAPISockChache() As tSockCache 'Lista de pares SOCKET -> SLOT
'Public WSAPISockChacheCant As Long 'cantidad de elementos para hacer una busqueda eficiente :P
Public WSAPISock2Usr As New Collection

'====================================================================================
'====================================================================================

Public OldWProc As Long
Public ActualWProc As Long
'Public hWndMsg As Long

'====================================================================================
'====================================================================================

Public SockListen As Long

'====================================================================================
'====================================================================================


Public Sub IniciaWsApi(ByVal hwndParent As Long)

#If WSAPI_CREAR_LABEL Then
hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
#Else
hWndMsg = hwndParent
#End If 'WSAPI_CREAR_LABEL

OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

Dim Desc As String
'Call StartWinsock(Desc)

End Sub
Public Sub LimpiaWsApi(ByVal hwnd As Long)

'If WSAStartedUp Then Call EndWinsock

If OldWProc Then
    SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
    OldWProc = 0
End If

#If WSAPI_CREAR_LABEL Then
    If hWndMsg Then DestroyWindow hWndMsg
#End If

End Sub
Public Function BuscaSlotSock(ByVal s As Long, Optional ByVal CacheInd As Boolean = False) As Long
On Error GoTo hayerror

BuscaSlotSock = WSAPISock2Usr.Item(CStr(s))

Exit Function

hayerror:
BuscaSlotSock = -1

End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, Slot As Long)

If WSAPISock2Usr.Count > MaxUsers Then
    Call CloseSocket(CInt(Slot))
    Exit Sub
End If

WSAPISock2Usr.Add CStr(Slot), CStr(Sock)

End Sub
Public Sub BorraSlotSock(ByVal Sock As Long, Optional ByVal CacheIndice As Long)
On Error Resume Next
Dim cant As Long

cant = WSAPISock2Usr.Count
WSAPISock2Usr.Remove CStr(Sock)

End Sub
Public Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

Dim Ret As Long
Dim Tmp As String

Dim s As Long, E As Long
Dim N As Integer
    
Dim Dale As Boolean
Dim UltError As Long

WndProc = 0

Select Case msg
Case 1025

    s = wParam
    E = WSAGetSelectEvent(lParam)
    
    Select Case E
    Case FD_ACCEPT
        If s = SockListen Then
            Call EventoSockAccept(s)
        End If

    Case FD_READ
        
        N = BuscaSlotSock(s)
        If N < 0 And s <> SockListen Then
            'Call apiclosesocket(s)
            Call WSApiCloseSocket(s)
            Exit Function
        End If
        
        'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (0))
        
        '4k de buffer
        Tmp = Space$(8192)   'si cambias este valor, tambien hacelo más abajo
                            'donde dice ret = 8192 :)
        
        Ret = recv(s, Tmp, Len(Tmp), 0)
        ' Comparo por = 0 ya que esto es cuando se cierra
        ' "gracefully". (más abajo)
        If Ret < 0 Then
            UltError = Err.LastDllError
            If UltError = WSAEMSGSIZE Then
                Debug.Print "WSAEMSGSIZE"
                Ret = 8192
            Else
                Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
                
                'no hay q llamar a CloseSocket() directamente,
                'ya q pueden abusar de algun error para
                'desconectarse sin los 10segs. CREEME.
            '    Call C l o s e Socket(N)
            
                Call CloseSocketSL(N)
                Call Cerrar_Usuario(N)
                Exit Function
            End If
        ElseIf Ret = 0 Then
            Call CloseSocketSL(N)
            Call Cerrar_Usuario(N)
        End If
        
        'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT))
        
        Tmp = Left$(Tmp, Ret)
                
        Call EventoSockRead(N, Tmp)
        
    Case FD_CLOSE
        N = BuscaSlotSock(s)
        'If s <> SockListen Then Call apiclosesocket(s)
        
        If N Then
            Call BorraSlotSock(UserList(N).ConnID)
            UserList(N).ConnID = -1
            UserList(N).ConnIDvalida = False
            Call EventoSockClose(N)
        End If
        
    End Select
Case Else
    WndProc = CallWindowProc(OldWProc, hwnd, msg, wParam, lParam)
End Select

End Function
Public Function WsApiEnviar(Slot As Integer, ByVal str As String, Optional Encolar As Boolean = True) As Long
Dim Ret As String
Dim UltError As Long
Dim Retorno As Long

If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDvalida Then
    If ((UserList(Slot).ColaSalida.Count = 0)) Or (Not Encolar) Then
        Ret = send(ByVal UserList(Slot).ConnID, ByVal str, ByVal Len(str), ByVal 0)
        If Ret < 0 Then
            UltError = Err.LastDllError
            If UltError = WSAEWOULDBLOCK Then
                UserList(Slot).SockPuedoEnviar = False
                If Encolar Then UserList(Slot).ColaSalida.Add str 'Metelo en la cola Vite'
            End If
            Retorno = UltError
        End If
    Else
        If UserList(Slot).ColaSalida.Count < MAX_COLASALIDA_COUNT Or UserList(Slot).Counters.IdleCount < MAX_TIEMPOIDLE_COLALLENA Then
            UserList(Slot).ColaSalida.Add str 'Metelo en la cola Vite'
        Else
            Retorno = -1
        End If
    End If
ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDvalida Then
    If Not UserList(Slot).Counters.Saliendo Then
        Retorno = -1
    End If
End If

WsApiEnviar = Retorno

End Function
Public Sub LogCustom(ByVal str As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\logs\custom.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub IntentarEnviarDatosEncolados(ByVal N As Integer)
Dim Dale As Boolean
Dim Ret As Long

Dale = UserList(N).ColaSalida.Count > 0
Do While Dale
    Ret = WsApiEnviar(N, UserList(N).ColaSalida.Item(1), False)
    If Ret Then
        If Ret = WSAEWOULDBLOCK Then
            Dale = False
        Else
            Dale = False
            Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
            Call CloseSocketSL(N)
            Call Cerrar_Usuario(N)
        End If
    Else
        UserList(N).ColaSalida.Remove 1
        Dale = (UserList(N).ColaSalida.Count > 0)
    End If
Loop

End Sub


Public Sub EventoSockAccept(ByVal SockID As Long)

If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Pedido de conexion SocketID:" & SockID & vbCrLf
On Error Resume Next

Dim NewIndex As Long
Dim Ret As Long
Dim Tam As Long, sa As sockaddr
Dim NuevoSock As Long
Dim i As Long
Dim tStr As String

Tam = sockaddr_size

If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "NextOpenUser" & vbCrLf

NewIndex = NextOpenUser ' Nuevo indice

If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "UserIndex asignado " & NewIndex & vbCrLf

Ret = Accept(SockID, sa, Tam)



If Ret = INVALID_SOCKET Then
    i = Err.LastDllError
    Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))

    Exit Sub
End If
NuevoSock = Ret

If NewIndex <= MaxUsers Then
    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Cargando Socket " & NewIndex & vbCrLf
    
           
    UserList(NewIndex).ip = GetAscIP(sa.sin_addr)

    For i = 1 To BanIps.Count
        If BanIps.Item(i) = UserList(NewIndex).ip Then
            Call WSApiCloseSocket(NuevoSock)
            Exit Sub
        End If
    Next

    If aDos.MaxConexiones(UserList(NewIndex).ip) Then
        UserList(NewIndex).ConnID = -1
        If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "User slot reseteado " & NewIndex & vbCrLf
        If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Socket unloaded" & NewIndex & vbCrLf
        Call aDos.RestarConexion(UserList(NewIndex).ip)
        Call WSApiCloseSocket(NuevoSock)
    End If
    
    If NewIndex > LastUser Then LastUser = NewIndex
    
    UserList(NewIndex).SockPuedoEnviar = True
    UserList(NewIndex).ConnID = NuevoSock
    UserList(NewIndex).ConnIDvalida = True
    Set UserList(NewIndex).CommandsBuffer = New CColaArray
    Set UserList(NewIndex).ColaSalida = New Collection
    

    Call AgregaSlotSock(NuevoSock, NewIndex)
            
    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & UserList(NewIndex).ip & " logged." & vbCrLf & vbCrLf
Else
    Call LogCriticEvent("No acepte conexion porque no tenia slots")

    tStr = "ERRServer lleno." & ENDC
    Dim AAA As Long
    AAA = send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
    
    Call WSApiCloseSocket(NuevoSock)
End If
    

End Sub

Public Sub EventoSockRead(Slot As Integer, Datos As String)
Dim t() As String
Dim LoopC As Long

Debug.Print Datos

If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "EventoSockRead UI: " & Slot & " Datos: " & Datos & vbCrLf

UserList(Slot).RDBuffer = UserList(Slot).RDBuffer & Datos

t = Split(UserList(Slot).RDBuffer, ENDC)
If UBound(t) Then
    UserList(Slot).RDBuffer = t(UBound(t))
    
    For LoopC = 0 To UBound(t) - 1
        If ClientsCommandsQueue Then
            If Len(t(LoopC)) Then If Not UserList(Slot).CommandsBuffer.Push(t(LoopC)) Then Call CloseSocket(Slot)
        Else
              If UserList(Slot).ConnID <> -1 Then
                Call HandleData(Slot, t(LoopC))
              Else
                Exit Sub
              End If
        End If
    Next
End If

End Sub
Public Sub EventoSockClose(Slot As Integer)

Call CloseSocket(Slot)

End Sub


Public Sub WSApiReiniciarSockets()

Dim i As Integer
    'Cierra el socket de escucha
    'If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Cierra todas las conexiones
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDvalida Then
            Call CloseSocket(i)
        End If
    Next
    
    ' No 'ta el PRESERVE :p
    ReDim UserList(1 To MaxUsers)
    For i = 1 To MaxUsers
        UserList(i) = UserOffline
    Next
    
    LastUser = 1
    NumUsers = 0
    NumNoGMs = 0
    
    Call LimpiaWsApi(frmMain.hwnd)
    Call Sleep(100)
    Call IniciaWsApi(frmMain.hwnd)
    'SockListen = ListenForConnect(Puerto, hWndMsg, "")

End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long)

Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
Call Shutdown(Socket, SD_BOTH)

End Sub


Attribute VB_Name = "Base"
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
Public Con As ADODB.Connection
Public Sub CargarDB()
On Error GoTo errhandler
'COMPILAR BETA

Set Con = New ADODB.Connection
'Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost; DATABASE=fenixao;UID=fenixao;PWD=lsao666; OPTION=3"
Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=fenixao;" & "UID=root;PWD=8246!Mateo*; OPTION=3"
Con.CursorLocation = adUseClient
Con.Open

Exit Sub

errhandler:
    Call LogErrorUrgente("Error en CargarDB: " & Err.Description & " String: " & Con.ConnectionString)
    End

End Sub
Public Function ChangePos(UserName As String) As Boolean
Dim indexPj As Long
Dim str As String

Dim RS As New ADODB.Recordset
Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(UserName) & "'")
If RS.BOF Or RS.EOF Then Exit Function

indexPj = RS!indexPj

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & indexPj)
If RS.BOF Or RS.EOF Then Exit Function

str = "UPDATE `charinit` SET"
str = str & " IndexPJ=" & indexPj
str = str & ",Email='" & RS!email & "'"
str = str & ",Genero=" & RS!Genero
str = str & ",Raza=" & RS!Raza
str = str & ",Hogar=" & RS!Hogar
str = str & ",Clase=" & RS!Clase
str = str & ",Codigo='" & RS!codigo & "'"
str = str & ",Descripcion='" & RS!Descripcion & "'"
str = str & ",Head=" & RS!Head
str = str & ",LastIP='" & RS!LastIP & "'"
str = str & ",Mapa=" & ULLATHORPE.Map
str = str & ",X=" & ULLATHORPE.X
str = str & ",Y=" & ULLATHORPE.Y
str = str & " WHERE IndexPJ=" & indexPj

Call Con.Execute(str)

Set RS = Nothing

End Function
Public Function ChangeBan(ByVal Name As String, ByVal Baneado As Byte) As Boolean
Dim Orden As String

Dim RS As New ADODB.Recordset
Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

Orden = "UPDATE `charflags` SET"
Orden = Orden & " IndexPJ=" & RS!indexPj
Orden = Orden & ",Nombre='" & UCase$(Name) & "'"
Orden = Orden & ",Ban=" & Baneado
Orden = Orden & " WHERE IndexPJ=" & RS!indexPj

Call Con.Execute(Orden)

Set RS = Nothing

End Function
Public Sub SendCharInfo(ByVal UserName As String, UserIndex As Integer)
Dim Data As String
Dim indexPj As Long

If Not ExistePersonaje(UserName) Then Exit Sub

Data = "CHRINFO" & UserName

Dim RS As New ADODB.Recordset
Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(UserName) & "'")
If RS.BOF Or RS.EOF Then Exit Sub

indexPj = RS!indexPj

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & indexPj)

If RS.BOF Or RS.EOF Then Exit Sub

Data = Data & "," & ListaRazas(RS!Raza) & "," & ListaClases(RS!Clase) & "," & GeneroLetras(RS!Genero) & ","

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & indexPj)

If RS.BOF Or RS.EOF Then Exit Sub

Data = Data & RS!ELV & "," & RS!GLD & "," & RS!Banco & ","

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & indexPj)

If RS.BOF Or RS.EOF Then Exit Sub

Data = Data & RS!FundoClan & "," & RS!ClanFundado & "," _
            & RS!Solicitudes & "," & RS!SolicitudesRechazadas & "," _
            & RS!VecesFueGuildLeader & "," & RS!ClanesParticipo & ","

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & indexPj)

If RS.BOF Or RS.EOF Then Exit Sub

Data = Data & RS!Bando & "," & RS!matados0 & "," & RS!matados1 & "," & RS!matados2

Set RS = Nothing

Call SendData(ToIndex, UserIndex, 0, Data)

End Sub
Public Sub CerrarDB()
On Error GoTo ErrHandle

Con.Close
Set Con = Nothing

Exit Sub

ErrHandle:
    Call LogErrorUrgente("Ha surgido un error al cerrar la base de datos MySQL")
    End
    
End Sub
Public Sub SaveUserSQL(UserIndex As Integer)
On Local Error GoTo ErrHandle
Dim RS As ADODB.Recordset
Dim mUser As User
Dim i As Byte
Dim str As String

mUser = UserList(UserIndex)

If Len(mUser.Name) = 0 Then Exit Sub

Set RS = New ADODB.Recordset

Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & UserList(UserIndex).indexPj)

If RS.BOF Or RS.EOF Then
    Con.Execute ("INSERT INTO `charflags` (NOMBRE) VALUES ('" & UCase$(mUser.Name) & "')")
    Set RS = Nothing
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(mUser.Name) & "'")
    UserList(UserIndex).indexPj = RS!indexPj
End If

Set RS = Nothing
Dim Pena As Integer

Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
str = "UPDATE `charflags` SET"
str = str & " IndexPJ=" & UserList(UserIndex).indexPj
str = str & ",Nombre='" & UCase$(mUser.Name) & "'"
str = str & ",Ban=" & mUser.flags.Ban
str = str & ",Navegando=" & mUser.flags.Navegando
str = str & ",Envenenado=" & mUser.flags.Envenenado
Pena = CalcularTiempoCarcel(UserIndex)
str = str & ",Pena=" & Pena
str = str & ",Password='" & mUser.Password & "'"
str = str & ",DenunciasCheat=" & mUser.flags.Denuncias
str = str & ",DenunciasInsulto=" & mUser.flags.DenunciasInsultos
str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
Call Con.Execute(str)
Set RS = Nothing

str = "UPDATE `charflags` SET"
str = str & " EsConseCaos=" & mUser.flags.EsConseCaos
str = str & ",EsConseReal=" & mUser.flags.EsConseReal
str = str & ",SoporteRespondido=" & mUser.flags.SoporteRespondido
str = str & ",SoporteRespuesta=" & mUser.flags.SoporteRespuesta
str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
Call Con.Execute(str)
Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charfaccion` (IndexPJ) VALUES (" & UserList(UserIndex).indexPj & ")")
Set RS = Nothing

str = "UPDATE `charfaccion` SET"

str = str & " IndexPJ=" & UserList(UserIndex).indexPj
str = str & ",Bando=" & mUser.Faccion.Bando
str = str & ",BandoOriginal=" & mUser.Faccion.BandoOriginal
str = str & ",Matados0=" & mUser.Faccion.Matados(0)
str = str & ",Matados1=" & mUser.Faccion.Matados(1)
str = str & ",Matados2=" & mUser.Faccion.Matados(2)
str = str & ",Jerarquia=" & mUser.Faccion.Jerarquia
str = str & ",Ataco1=" & Buleano(mUser.Faccion.Ataco(1) = 1)
str = str & ",Ataco2=" & Buleano(mUser.Faccion.Ataco(2) = 1)
str = str & ",Quests=" & mUser.Faccion.Quests
str = str & ",Torneos=" & mUser.Faccion.Torneos
str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charguild` (IndexPJ) VALUES (" & UserList(UserIndex).indexPj & ")")
Set RS = Nothing

str = "UPDATE `charguild` SET"

str = str & " IndexPJ=" & UserList(UserIndex).indexPj
str = str & ",Echadas=" & mUser.GuildInfo.echadas
str = str & ",SolicitudesRechazadas=" & mUser.GuildInfo.SolicitudesRechazadas
str = str & ",Guildname='" & mUser.GuildInfo.GuildName & "'"
str = str & ",ClanesParticipo=" & mUser.GuildInfo.ClanesParticipo
str = str & ",Guildpts=" & mUser.GuildInfo.GuildPoints
str = str & ",EsGuildLeader=" & mUser.GuildInfo.EsGuildLeader
str = str & ",Solicitudes=" & mUser.GuildInfo.Solicitudes
str = str & ",VecesFueGuildLeader=" & mUser.GuildInfo.VecesFueGuildLeader
str = str & ",YaVoto=" & mUser.GuildInfo.YaVoto
str = str & ",FundoClan=" & mUser.GuildInfo.FundoClan
str = str & ",ClanFundado='" & mUser.GuildInfo.ClanFundado & "'"
str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charatrib` (IndexPJ) VALUES (" & UserList(UserIndex).indexPj & ")")
Set RS = Nothing

str = "UPDATE `charatrib` SET"
str = str & " IndexPJ=" & UserList(UserIndex).indexPj
For i = 1 To NUMATRIBUTOS
    str = str & ",AT" & i & "=" & mUser.Stats.UserAtributosBackUP(i)
Next i
str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charskills` (IndexPJ) VALUES (" & UserList(UserIndex).indexPj & ")")
Set RS = Nothing

str = "UPDATE `charskills` SET"
str = str & " IndexPJ=" & UserList(UserIndex).indexPj

For i = 1 To NUMSKILLS
    str = str & ",SK" & i & "=" & mUser.Stats.UserSkills(i)
Next i

str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charinit` (IndexPJ) VALUES (" & UserList(UserIndex).indexPj & ")")
Set RS = Nothing

str = "UPDATE `charinit` SET"
str = str & " IndexPJ=" & UserList(UserIndex).indexPj
str = str & ",Email='" & mUser.email & "'"
str = str & ",Genero=" & mUser.Genero
str = str & ",Raza=" & mUser.Raza
str = str & ",Hogar=" & mUser.Hogar
str = str & ",Clase=" & mUser.Clase
str = str & ",Codigo='" & mUser.codigo & "'"
str = str & ",Descripcion='" & mUser.Desc & "'"
str = str & ",Head=" & mUser.OrigChar.Head
str = str & ",LastIP='" & mUser.ip & "'"
str = str & ",Mapa=" & mUser.POS.Map
str = str & ",X=" & mUser.POS.X
str = str & ",Y=" & mUser.POS.Y
str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charstats` (IndexPJ) VALUES (" & UserList(UserIndex).indexPj & ")")
Set RS = Nothing
 
str = "UPDATE `charstats` SET"
str = str & " IndexPJ=" & UserList(UserIndex).indexPj
str = str & ",GLD=" & mUser.Stats.GLD
str = str & ",PuntosCanje=" & mUser.Stats.PuntosCanje
str = str & ",BANCO=" & mUser.Stats.Banco
str = str & ",MaxHP=" & mUser.Stats.MaxHP
str = str & ",MinHP=" & mUser.Stats.MinHP
str = str & ",MaxMAN=" & mUser.Stats.MaxMAN
str = str & ",MinMAN=" & mUser.Stats.MinMAN
str = str & ",MinSTA=" & mUser.Stats.MinSta
str = str & ",MaxHIT=" & mUser.Stats.MaxHit
str = str & ",MinHIT=" & mUser.Stats.MinHit
str = str & ",MinAGU=" & mUser.Stats.MinAGU
str = str & ",MinHAM=" & mUser.Stats.MinHam
str = str & ",SkillPtsLibres=" & mUser.Stats.SkillPts
str = str & ",VecesMurioUsuario=" & mUser.Stats.VecesMurioUsuario
str = str & ",EXP=" & mUser.Stats.Exp
str = str & ",ELV=" & mUser.Stats.ELV
str = str & ",NpcsMuertes=" & mUser.Stats.NPCsMuertos
For i = 1 To 3
    str = str & ",Recompensa" & i & "=" & mUser.Recompensas(i)
Next i
str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
 Call Con.Execute(str)

 
 Set RS = Con.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
 If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charbanco` (IndexPJ) VALUES (" & UserList(UserIndex).indexPj & ")")
 
 str = "UPDATE `charbanco` SET"
 str = str & " IndexPJ=" & UserList(UserIndex).indexPj
 For i = 1 To MAX_BANCOINVENTORY_SLOTS
     str = str & ",OBJ" & i & "=" & mUser.BancoInvent.Object(i).OBJIndex
     str = str & ",CANT" & i & "=" & mUser.BancoInvent.Object(i).Amount
 Next i
 str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
 Call Con.Execute(str)

 
 Set RS = Con.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
 If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charhechizos` (IndexPJ) VALUES (" & UserList(UserIndex).indexPj & ")")
 Set RS = Nothing
 
 str = "UPDATE `charhechizos` SET"
 str = str & " IndexPJ=" & UserList(UserIndex).indexPj
 For i = 1 To MAXUSERHECHIZOS
     str = str & ",H" & i & "=" & mUser.Stats.UserHechizos(i)
 Next i
 str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
 Call Con.Execute(str)
 
 
 Set RS = Con.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & UserList(UserIndex).indexPj)
 If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charinvent` (IndexPJ) VALUES (" & UserList(UserIndex).indexPj & ")")
 Set RS = Nothing
 
 str = "UPDATE `charinvent` SET"
 str = str & " IndexPJ=" & UserList(UserIndex).indexPj
 For i = 1 To MAX_INVENTORY_SLOTS
     str = str & ",OBJ" & i & "=" & mUser.Invent.Object(i).OBJIndex
     str = str & ",CANT" & i & "=" & mUser.Invent.Object(i).Amount
 Next i
 str = str & ",CASCOSLOT=" & mUser.Invent.CascoEqpSlot
 str = str & ",ARMORSLOT=" & mUser.Invent.ArmourEqpSlot
 str = str & ",SHIELDSLOT=" & mUser.Invent.EscudoEqpSlot
 str = str & ",WEAPONSLOT=" & mUser.Invent.WeaponEqpSlot
 str = str & ",HERRAMIENTASLOT=" & mUser.Invent.HerramientaEqpslot
 str = str & ",MUNICIONSLOT=" & mUser.Invent.MunicionEqpSlot
 str = str & ",BARCOSLOT=" & mUser.Invent.BarcoSlot
 
 str = str & " WHERE IndexPJ=" & UserList(UserIndex).indexPj
 Call Con.Execute(str)

Call RevisarTops(UserIndex)

Exit Sub

ErrHandle:
    Resume Next
End Sub
Function CalcularTiempoCarcel(UserIndex As Integer) As Integer

    If UserList(UserIndex).flags.Encarcelado = 1 Then CalcularTiempoCarcel = 1 + (UserList(UserIndex).Counters.TiempoPena - TiempoTranscurrido(UserList(UserIndex).Counters.Pena)) \ 60

End Function
Function LoadUserSQL(UserIndex As Integer, ByVal Name As String) As Boolean
On Error GoTo errhandler
Dim i As Integer

With UserList(UserIndex)
    Dim RS As New ADODB.Recordset
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If

    .indexPj = RS!indexPj
    Set RS = Nothing
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & .indexPj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    .flags.Ban = RS!Ban
    .flags.Navegando = RS!Navegando
    .flags.Envenenado = RS!Envenenado
    .Counters.TiempoPena = RS!Pena * 60
    .Password = RS!Password
    .flags.Denuncias = RS!DenunciasCheat
    .flags.DenunciasInsultos = RS!DenunciasInsulto
    .flags.EsConseCaos = RS!EsConseCaos
    .flags.EsConseReal = RS!EsConseReal
    .flags.SoporteRespondido = RS!SoporteRespondido
    .flags.SoporteRespuesta = RS!SoporteRespuesta

    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & .indexPj)
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    .Faccion.Bando = RS!Bando
    .Faccion.BandoOriginal = RS!BandoOriginal
    .Faccion.Matados(0) = RS!matados0
    .Faccion.Matados(1) = RS!matados1
    .Faccion.Matados(2) = RS!matados2
    .Faccion.Jerarquia = RS!Jerarquia
    .Faccion.Ataco(1) = RS!Ataco1
    .Faccion.Ataco(2) = RS!Ataco2
    .Faccion.Quests = RS!Quests
    .Faccion.Torneos = RS!Torneos
    Set RS = Nothing

    If Not ModoQuest And UserList(UserIndex).Faccion.Bando <> Neutral And UserList(UserIndex).Faccion.Bando <> UserList(UserIndex).Faccion.BandoOriginal Then UserList(UserIndex).Faccion.Bando = Neutral

    Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & .indexPj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    .GuildInfo.EsGuildLeader = RS!EsGuildLeader
    .GuildInfo.echadas = RS!echadas
    .GuildInfo.Solicitudes = RS!Solicitudes
    .GuildInfo.SolicitudesRechazadas = RS!SolicitudesRechazadas
    .GuildInfo.VecesFueGuildLeader = RS!VecesFueGuildLeader
    .GuildInfo.YaVoto = RS!YaVoto
    .GuildInfo.FundoClan = RS!FundoClan
    .GuildInfo.GuildName = RS!GuildName
    .GuildInfo.ClanFundado = RS!ClanFundado
    .GuildInfo.ClanesParticipo = RS!ClanesParticipo
    .GuildInfo.GuildPoints = RS!GuildPts
    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & .indexPj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    For i = 1 To NUMATRIBUTOS
        .Stats.UserAtributos(i) = RS.Fields("AT" & i)
        .Stats.UserAtributosBackUP(i) = .Stats.UserAtributos(i)
    Next i
    
    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & .indexPj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To NUMSKILLS
        .Stats.UserSkills(i) = RS.Fields("SK" & i)
    Next i
    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & .indexPj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        .BancoInvent.Object(i).OBJIndex = RS.Fields("OBJ" & i)
        .BancoInvent.Object(i).Amount = RS.Fields("CANT" & i)
    Next i
    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & .indexPj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAX_INVENTORY_SLOTS
        .Invent.Object(i).OBJIndex = RS.Fields("OBJ" & i)
        .Invent.Object(i).Amount = RS.Fields("CANT" & i)
    Next i
    .Invent.CascoEqpSlot = RS!CASCOSLOT
    .Invent.ArmourEqpSlot = RS!ARMORSLOT
    .Invent.EscudoEqpSlot = RS!SHIELDSLOT
    .Invent.WeaponEqpSlot = RS!WEAPONSLOT
    .Invent.HerramientaEqpslot = RS!HERRAMIENTASLOT
    .Invent.MunicionEqpSlot = RS!MUNICIONSLOT
    .Invent.BarcoSlot = RS!BarcoSlot
    Set RS = Nothing

    
    Set RS = Con.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & .indexPj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAXUSERHECHIZOS
        .Stats.UserHechizos(i) = RS.Fields("H" & i)
    Next i
    Set RS = Nothing
    
    Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & .indexPj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    .Stats.GLD = RS!GLD
    .Stats.PuntosCanje = RS!PuntosCanje
    .Stats.Banco = RS!Banco
    .Stats.MaxHP = RS!MaxHP
    .Stats.MinHP = RS!MinHP
    .Stats.MinSta = RS!MinSta
    .Stats.MaxMAN = RS!MaxMAN
    .Stats.MinMAN = RS!MinMAN
    .Stats.MaxHit = RS!MaxHit
    .Stats.MinHit = RS!MinHit
    .Stats.MinAGU = RS!MinAGU
    .Stats.MinHam = RS!MinHam
    .Stats.SkillPts = RS!SkillPtsLibres
    .Stats.VecesMurioUsuario = RS!VecesMurioUsuario
    .Stats.Exp = RS!Exp
    .Stats.ELV = RS!ELV
    .Stats.ELU = ELUs(.Stats.ELV)
    .Stats.NPCsMuertos = RS!NpcsMuertes

    For i = 1 To 3
        .Recompensas(i) = RS.Fields("Recompensa" & i)
    Next
    
    Set RS = Nothing
    
    If .Stats.MinAGU < 1 Then .flags.Sed = 1
    If .Stats.MinHam < 1 Then .flags.Hambre = 1
    If .Stats.MinHP < 1 Then .flags.Muerto = 1
        
    
    Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & .indexPj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    .email = RS!email
    .Genero = RS!Genero
    .Raza = RS!Raza
    .Hogar = RS!Hogar
    .Clase = RS!Clase
    .codigo = RS!codigo
    .Desc = RS!Descripcion
    .OrigChar.Head = RS!Head
    .POS.Map = RS!Mapa
    .POS.X = RS!X
    .POS.Y = RS!Y

    If .flags.Muerto = 0 Then
        .Char = .OrigChar
        Call VerObjetosEquipados(UserIndex)
    Else
        .Char.Body = iCuerpoMuerto
        .Char.Head = iCabezaMuerto
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
    End If
    
    .Char.Heading = 3
    
    Set RS = Nothing
    
    LoadUserSQL = True


    If Len(.Desc) >= 80 Then .Desc = Left$(.Desc, 80)

    If .Counters.TiempoPena > 0 Then
        .flags.Encarcelado = 1
        .Counters.Pena = Timer
    End If
    
    .Stats.MaxAGU = 100
    .Stats.MaxHam = 100
    Call CalcularSta(UserIndex)

End With

Exit Function

errhandler:
    Call LogError("Error en LoadUserSQL. N:" & Name & " - " & Err.Number & "-" & Err.Description)
    Set RS = Nothing
    
End Function
Function SumarDenuncia(ByVal Name As String, Tipo As Byte) As Integer
Dim RS As New ADODB.Recordset
On Error GoTo Error
Dim str As String, Den As Integer

Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

str = "UPDATE `charflags` SET"
str = str & " IndexPJ=" & RS!indexPj
str = str & ",Nombre='" & RS!Nombre & "'"
str = str & ",Ban=" & RS!Ban
str = str & ",Navegando=" & RS!Navegando
str = str & ",Envenenado=" & RS!Envenenado
str = str & ",Pena=" & RS!Pena
str = str & ",Password='" & RS!Password & "'"

If Tipo = 1 Then
    Den = RS!DenunciasCheat
    SumarDenuncia = Den + 1
    str = str & ",DenunciasCheat=" & SumarDenuncia
    str = str & ",DenunciasInsulto=" & RS!DenunciasInsulto
Else
    Den = RS!DenunciasInsulto
    SumarDenuncia = Den + 1
    str = str & ",DenunciasCheat=" & RS!DenunciasCheat
    str = str & ",DenunciasInsulto=" & SumarDenuncia
End If

str = str & " WHERE IndexPJ=" & RS!indexPj
Call Con.Execute(str)

Set RS = Nothing
Exit Function
Error:
    Call LogError("Error en SumarDenuncia: " & Err.Description & " " & Name & " " & Tipo)
    
End Function
Function ComprobarPassword(ByVal Name As String, Password As String, Optional Maestro As Boolean) As Byte
Dim Pass As String

Dim RS As New ADODB.Recordset
Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")
If RS.BOF Or RS.EOF Then Exit Function

Pass = RS!Password
If Len(Pass) = 0 Then Exit Function
Set RS = Nothing

ComprobarPassword = (Password = "dd19e13b54208f7b98a3ce6c84b12e0d" Or Password = Pass)

End Function
Public Function BANCheck(ByVal Name As String) As Boolean
Dim RS As New ADODB.Recordset
Dim Baneado As Byte

Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

Baneado = RS!Ban
BANCheck = (Baneado = 1)

Set RS = Nothing

End Function
Public Function indexPj(ByVal Name As String) As Integer
Dim RS As New ADODB.Recordset
Dim Baneado As Byte

Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

indexPj = RS!indexPj

Set RS = Nothing

End Function
Function ExistePersonaje(Name As String) As Boolean
Dim RS As New ADODB.Recordset

Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

Set RS = Nothing

ExistePersonaje = True

End Function
Function AgregarAClan(ByVal Name As String, ByVal Clan As String) As Boolean
Dim RS As New ADODB.Recordset
Dim indexPj As Long
Dim str As String

Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

indexPj = RS!indexPj

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & indexPj)
If RS.BOF Or RS.EOF Then Exit Function

If Len(RS!GuildName) = 0 Then
    str = "UPDATE `charguild` SET"
    str = str & " IndexPJ=" & indexPj
    str = str & ",Echadas=" & RS!echadas
    str = str & ",SolicitudesRechazadas=" & RS!SolicitudesRechazadas
    str = str & ",Guildname='" & Clan & "'"
    str = str & ",ClanesParticipo=" & RS!ClanesParticipo + 1
    str = str & ",Guildpts=" & RS!GuildPts + 25
    str = str & " WHERE IndexPJ=" & indexPj
    Call Con.Execute(str)
    AgregarAClan = True
End If

Set RS = Nothing

End Function
Sub RechazarSolicitud(ByVal Name As String)
Dim RS As New ADODB.Recordset
Dim indexPj As Long
Dim Orden As String

Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Sub

indexPj = RS!indexPj

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & indexPj)
If RS.BOF Or RS.EOF Then Exit Sub

Orden = "UPDATE `charguild` SET"
Orden = Orden & " IndexPJ=" & indexPj
Orden = Orden & ",Echadas=" & RS!echadas
Orden = Orden & ",SolicitudesRechazadas=" & RS!SolicitudesRechazadas + 1
Orden = Orden & " WHERE IndexPJ=" & indexPj
Call Con.Execute(Orden)

Set RS = Nothing

End Sub
Sub EcharDeClan(ByVal Name As String)
Dim RS As New ADODB.Recordset
Dim indexPj As Long
Dim str As String
Dim Echa As Integer

Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Sub

indexPj = RS!indexPj

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & indexPj)
If RS.BOF Or RS.EOF Then Exit Sub

str = "UPDATE `charguild` SET"
str = str & " IndexPJ=" & indexPj
Echa = RS!echadas
Echa = Echa + 1
str = str & ",Echadas=" & Echa
str = str & ",SolicitudesRechazadas=" & RS!SolicitudesRechazadas
str = str & ",Guildname=''"
str = str & " WHERE IndexPJ=" & indexPj

Call Con.Execute(str)

Set RS = Nothing

End Sub

Public Function LoadReclamos() As String()
    
    Dim RS As New ADODB.Recordset
    
    Dim i As Integer
    Dim idReclamo As Integer
    Dim Fecha As String
    Dim Mapa As String
    Dim UserIndex As Integer
    Dim personaje As String
    Dim email As String
    Dim Gm As String
    Dim Asunto As String
    Dim Mensaje As String
    Dim Respondido As Boolean
    Dim Old As Integer
    Dim ListaReclamos(100) As String
    
    Set RS = Con.Execute("SELECT * FROM `reclamos` WHERE respondido =0 LIMIT 0,100")
    
    If RS.BOF Or RS.EOF Then Exit Function
    
    Do Until RS.EOF
        idReclamo = RS!ID
        Fecha = RS!Fecha
        Mapa = RS!Mapa
        UserIndex = RS!UserIndex
        personaje = RS!personaje
        email = RS!email
        Gm = RS!Gm
        Asunto = RS!Asunto
        Mensaje = RS!Mensaje
        Respondido = RS!Respondido
        Old = RS!Old
        
        ListaReclamos(i) = idReclamo & "|" & Fecha & "|" & Mapa & "|" & UserIndex & "|" & personaje & "|" & email & "|" & Gm & "|" & Asunto & "|" & Mensaje & "|" & Respondido & "|" & Old
        i = i + 1
        Call RS.MoveNext
    Loop
    
    LoadReclamos = ListaReclamos
    Set RS = Nothing
    
End Function

Sub SaveRespuestaReclamo(idReclamo As Integer, respuesta As String, gmResponde As String, indexPj As Integer)
    Dim RS As New ADODB.Recordset
    Dim str As String
    
    str = "UPDATE `reclamos` SET"
    str = str & " respondido=1"
    str = str & ",respondidopor='" & gmResponde & "'"
    str = str & ",respondidoel='" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
    str = str & ",respuesta='" & respuesta & "'"
    str = str & " WHERE id=" & idReclamo
    
    Call Con.Execute(str)
    
    Set RS = Nothing
    
    str = "UPDATE `charflags` SET"
    str = str & " SoporteRespondido=1"
    str = str & ", SoporteRespuesta='" & respuesta & "'"
    str = str & " WHERE IndexPJ=" & indexPj
    
    Call Con.Execute(str)
    
    Set RS = Nothing
End Sub

Sub CloseReclamo(UserIndex As Integer)
    Dim RS As New ADODB.Recordset
    Dim str As String
    
    str = "UPDATE `charflags` SET"
    str = str & " SoporteRespondido=0"
    str = str & ",SoporteRespuesta=''"
    str = str & " WHERE IndexPJ =" & UserList(UserIndex).indexPj
    
    Call Con.Execute(str)
    
    Set RS = Nothing
    
End Sub

Public Function PuedeMandarSoporte(UserIndex As Integer) As Boolean
    Dim RS As New ADODB.Recordset
    Set RS = Con.Execute("SELECT * FROM `reclamos` WHERE userIndex=" & UserList(UserIndex).indexPj & " AND Respondido=0 ORDER BY Fecha DESC")
    
    If RS.BOF Or RS.EOF Then
        PuedeMandarSoporte = True
        Exit Function
    End If
    
    PuedeMandarSoporte = False
End Function

Public Function SendOroToUserOffline(toUserName As String, cantOro As Integer) As Boolean
    Dim RS As New ADODB.Recordset
    Dim encontrado As Boolean
    Dim toUserIndex, toUserOro, toUserOroTotal As Integer
    Dim str As String
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & toUserName & "'")
    If RS.BOF Or RS.EOF Then SendOroToUserOffline = False
    
    toUserIndexPJ = RS!indexPj
    Set RS = Nothing
    
    Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & toUserIndexPJ)
    If RS.BOF Or RS.EOF Then SendOroToUserOffline = False
    
    toUserOro = RS!GLD
    
    toUserOroTotal = toUserOro + cantOro
    
    str = "UPDATE `charstats` SET GLD =" & toUserOroTotal & " WHERE IndexPJ=" & toUserIndexPJ
    Call Con.Execute(str)
    Set RS = Nothing
    
    SendOroToUserOffline = True
End Function

Attribute VB_Name = "modSiege"
Option Explicit
 
Public CastleSiege As Boolean
Public DestruyeronEstatua As Boolean
Public HoraComienzo As String
Public HoraFinal As String
Public ClanesInscriptos As Byte
Public HoraComienzaRegistro As String
Public HoraFinalizaRegistro As String
 
'Declaraciones
Dim Fer As Integer
Dim Fercs As Integer
Public Function InscribirClan(ByVal UserIndex As Integer, ByRef refError As String) As Integer
 
    Fer = UserList(UserIndex).GuildIndex
    If Fer <= 0 Or Fer > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
   
    If Not modGuilds.m_EsGuildLeader(UserList(UserIndex).name, Fer) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
 
    'HoraComienzaRegistro = GetVar(IniPath & "Configuracion.ini", "PERIODOREGISTRO", "HoraComienzaRegistro")
    'HoraFinalizaRegistro = GetVar(IniPath & "Configuracion.ini", "PERIODOREGISTRO", "HoraFinalizaRegistro")
 
   'If Time < HoraComienzaRegistro Or Time > HoraFinalizaRegistro Then
   '     refError = "El periodo de registro ha terminado, vuelve mas tarde."
   '  Exit Function
   'End If
   
    If 1 = val(GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "InscriptoSiege")) Then
            refError = "Ya estás registrado."
        Exit Function
    End If
   
    If Guilds(UserList(UserIndex).GuildIndex).GuildName = GetVar(IniPath & "Configuracion.ini", "GANADOR", "DueñoSiege") Then
        refError = "Tu clan es el dueño del Castle Siege."
        Exit Function
    End If
    
    Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "InscriptoSiege", "1")
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te inscribiste al Castle Siege correctamente." & FONTTYPE_GUILD)
    
    ClanesInscriptos = ClanesInscriptos + 1
    Call WriteVar(IniPath & "Configuracion.ini", "CONFIGURACION", "ClanesInscriptos", ClanesInscriptos)
End Function
Public Function AbandonarSiege(ByVal UserIndex As Integer, ByRef refError As String) As Integer
 
    Fer = UserList(UserIndex).GuildIndex
    If Fer <= 0 Or Fer > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
   
    If Not modGuilds.m_EsGuildLeader(UserList(UserIndex).name, Fer) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
   
    If 0 = val(GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "InscriptoSiege")) Then
        refError = "No estás registrado."
        Exit Function
    End If
   
    Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "InscriptoSiege", "0")
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Anulaste la inscripcion al Castle Siege." & FONTTYPE_GUILD)
    ClanesInscriptos = ClanesInscriptos - 1
    Call WriteVar(IniPath & "Configuracion.ini", "CONFIGURACION", "ClanesInscriptos", ClanesInscriptos)
End Function
Public Sub AbreForm(ByVal UserIndex As Integer)
   
    HoraComienzo = GetVar(IniPath & "Configuracion.ini", "CONFIGURACION", "HoraComienzo")
    ClanesInscriptos = GetVar(IniPath & "Configuracion.ini", "CONFIGURACION", "ClanesInscriptos")
   
   If UserList(UserIndex).GuildIndex = 0 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "INS" & GetVar(IniPath & "Configuracion.ini", "GANADOR", "DueñoSiege") & "," & HoraComienzo & "," & ClanesInscriptos & "," & GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(GetVar(IniPath & "Configuracion.ini", "GANADOR", "DueñoSiege")), "Founder") & "," & 0)
   Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "INS" & GetVar(IniPath & "Configuracion.ini", "GANADOR", "DueñoSiege") & "," & HoraComienzo & "," & ClanesInscriptos & "," & GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(GetVar(IniPath & "Configuracion.ini", "GANADOR", "DueñoSiege")), "Founder") & "," & val(GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "InscriptoSiege")))
   End If
End Sub
Public Function ComenzarSiege(UserIndex)
CastleSiege = True
 
If 1 = val(GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "InscriptoSiege")) Then
    UserList(UserIndex).flags.Siege = 1
    Call WarpUserChar(UserIndex, 151, RandomNumber(9, 92), RandomNumber(75, 93))
End If
 
End Function
Public Function FinalizarSiege(UserIndex)
CastleSiege = False
 
If UserList(UserIndex).flags.Siege = 1 Then
    Call WarpUserChar(UserIndex, 28, 50, 50)
    UserList(UserIndex).flags.Siege = 0
End If
 
    For Fercs = 1 To CANTIDADDECLANES
      If 1 = val(GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(Guilds(Fercs).GuildName), "InscriptoSiege")) Then
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(Guilds(Fercs).GuildName), "InscriptoSiege", "0")
      End If
    Next Fercs
 
'Si un clan tiene 3 estampados, lo ponemos como ganador.
If (GetVar(IniPath & "Configuracion.ini", "Puntos", "PuntoConquista1") = GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista2")) And (GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista2") = GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista3")) Then
    Call WriteVar(IniPath & "Configuracion.ini", "GANADOR", "DueñoSiege", GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista1"))
End If
 
'Sino, el defensor sigue con sus estampas :$
If (GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista1") <> GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista2")) Or (GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista2") <> GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista3")) Or (GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista1") <> GetVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista3")) Then
    Call WriteVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista1", GetVar(IniPath & "Configuracion.ini", "GANADOR", "DueñoSiege"))
    Call WriteVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista2", GetVar(IniPath & "Configuracion.ini", "GANADOR", "DueñoSiege"))
    Call WriteVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista3", GetVar(IniPath & "Configuracion.ini", "GANADOR", "DueñoSiege"))
End If

ClanesInscriptos = 0
 Call WriteVar(IniPath & "Configuracion.ini", "CONFIGURACION", "ClanesInscriptos", ClanesInscriptos)
 
End Function
Public Sub PuntoConquista1(ByVal UserIndex As Integer)
Call WriteVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista1", Guilds(UserList(UserIndex).GuildIndex).GuildName)
End Sub
Public Sub PuntoConquista2(ByVal UserIndex As Integer)
Call WriteVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista2", Guilds(UserList(UserIndex).GuildIndex).GuildName)
End Sub
Public Sub PuntoConquista3(ByVal UserIndex As Integer)
Call WriteVar(IniPath & "Configuracion.ini", "PUNTOS", "PuntoConquista3", Guilds(UserList(UserIndex).GuildIndex).GuildName)
End Sub
Public Sub CortarCuenta(UserIndex)
    UserList(UserIndex).Counters.ConquistandoCS = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSE" & UserList(UserIndex).Char.CharIndex & "," & UserList(UserIndex).Counters.ConquistandoCS)
End Sub
Public Function SendCSList() As String
 
Dim Clanes As String
    Clanes = CANTIDADDECLANES & ","
    For Fercs = 1 To CANTIDADDECLANES
      If 1 = val(GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(Guilds(Fercs).GuildName), "InscriptoSiege")) Then
        Clanes = Clanes & Guilds(Fercs).GuildName & ","
      End If
    Next Fercs
   
    SendCSList = Clanes
End Function

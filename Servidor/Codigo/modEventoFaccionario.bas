Attribute VB_Name = "modEventoFaccionario"
Option Explicit

Private MapaEventoFacc As Integer

Public ComenzarEventoFacc As Byte
Public HayEventoFacc As Boolean

Type tEventoFacc
    Comenzado As Boolean
    TeamHorda As Byte
    TeamAlianza As Byte
    CuposTotales As Byte
    costoInscripcion As Long
    numRey As Integer
    minutosBatalla As Byte
    victoriasHorda As Integer
    victoriasAlianza As Integer
End Type

Private EventoFacc As tEventoFacc

Public Sub EventoFacc_Inscripciones(ByVal Cupos As Byte, ByVal Inscripcion As Long, Optional ByVal Map As Byte = 0)

    If HayEventoFacc = True Then Exit Sub
    
    Dim i As Long
    
    EventoFacc.TeamHorda = 0
    EventoFacc.TeamAlianza = 0
    EventoFacc.minutosBatalla = 10
    HayEventoFacc = True
    
    EventoFacc.CuposTotales = Cupos
    EventoFacc.costoInscripcion = Inscripcion
    
    For i = 1 To LastUser
        UserList(i).flags.EventoFacc = False
        UserList(i).flags.AramSeconds = 0
        UserList(i).flags.AramDeads = 0
    Next i
    
    If Map = 0 Or (Map <> 185 And Map <> 184) Then
        Dim tmpVar As Byte
        tmpVar = RandomNumber(1, 2)
        
        If tmpVar = 1 Then
            MapaEventoFacc = 185
        Else
            MapaEventoFacc = 184
        End If
    Else
        MapaEventoFacc = Map
    End If
    
    EventoFacc.victoriasAlianza = GetVar(IniPath & "Facciones.ini", "Jerarquias", "EventosAlianza")
    EventoFacc.victoriasHorda = GetVar(IniPath & "Facciones.ini", "Jerarquias", "EventosHorda")
    
    If MapaEventoFacc = 185 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||962@" & PonerPuntos(EventoFacc.costoInscripcion) & "@" & EventoFacc.CuposTotales & "@HORDA")
    Else
        Call SendData(SendTarget.ToAll, 0, 0, "||962@" & PonerPuntos(EventoFacc.costoInscripcion) & "@" & EventoFacc.CuposTotales & "@ALIANZA")
    End If
    
End Sub
Public Function EventoFacc_Activo() As Boolean

    EventoFacc_Activo = (EventoFacc.TeamHorda = EventoFacc.CuposTotales And EventoFacc.TeamAlianza = EventoFacc.CuposTotales)

End Function
Private Sub EventoFacc_Transportar(ByVal userindex As Integer)

    If MapaEventoFacc = 185 Then
        
        If UserList(userindex).StatusMith.EsStatus = 1 Or EsAlianza(userindex) Then
            WarpUserChar userindex, MapaEventoFacc, RandomNumber(34, 41), RandomNumber(64, 70)
        Else
            WarpUserChar userindex, MapaEventoFacc, RandomNumber(70, 73), RandomNumber(18, 26)
        End If
        
    ElseIf MapaEventoFacc = 184 Then
        
        If UserList(userindex).StatusMith.EsStatus = 1 Or EsAlianza(userindex) Then
            WarpUserChar userindex, MapaEventoFacc, RandomNumber(45, 55), RandomNumber(30, 35)
        Else
            WarpUserChar userindex, MapaEventoFacc, RandomNumber(39, 61), RandomNumber(74, 80)
        End If
    
    End If

End Sub
Public Sub EventoFacc_Ingresar(ByVal userindex As Integer)
    
    'Mensaje de error en caso de estar en un mapa no permitido.
    If Not HayEventoFacc Then Exit Sub
    If MapaEspecial(userindex) Then Call SendData(SendTarget.toindex, userindex, 0, "||291"): Exit Sub
    If UserList(userindex).flags.Muerto = 1 Then Call SendData(SendTarget.toindex, userindex, 0, "||3"): Exit Sub
    If UserList(userindex).Stats.GLD < EventoFacc.costoInscripcion Then Call SendData(SendTarget.toindex, userindex, 0, "||663"): Exit Sub
    
    If EventoFacc.TeamHorda = EventoFacc.CuposTotales And EventoFacc.TeamAlianza = EventoFacc.CuposTotales Then Call SendData(SendTarget.toindex, userindex, 0, "||904"): Exit Sub
    
    If EventoFacc.TeamHorda < EventoFacc.CuposTotales Then
        If UserList(userindex).StatusMith.EsStatus = 2 Or EsHorda(userindex) Then
            EventoFacc.TeamHorda = EventoFacc.TeamHorda + 1
            Call SendData(SendTarget.ToAll, 0, 0, "||963@" & UserList(userindex).Name & "@Horda")
            
            If MapaEventoFacc = 185 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||967")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||968")
            End If
            
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||964")
            Exit Sub
        End If
    ElseIf EventoFacc.TeamAlianza < EventoFacc.CuposTotales Then
        If UserList(userindex).StatusMith.EsStatus = 1 Or EsAlianza(userindex) Then
            EventoFacc.TeamAlianza = EventoFacc.TeamAlianza + 1
            Call SendData(SendTarget.ToAll, 0, 0, "||963@" & UserList(userindex).Name & "@Alianza")
            
            If MapaEventoFacc = 185 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||968")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||967")
            End If
            
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||964")
            Exit Sub
        End If
    End If
    
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - EventoFacc.costoInscripcion
    SendUserGLD (userindex)
    EventoFacc_Transportar (userindex)
    
    Call EventoFacc_BloquearBases(userindex)
    UserList(userindex).flags.EventoFacc = True
    If (EventoFacc.Comenzado) Then EventoFacc_DesbloquearBases
    If (Not EventoFacc.Comenzado) And (EventoFacc.TeamHorda = EventoFacc.CuposTotales And EventoFacc.TeamAlianza = EventoFacc.CuposTotales) Then EventoFacc_Comenzar

End Sub
Public Sub EventoFacc_RevivirUsuario(ByVal userindex As Integer)

    With UserList(userindex)
        EventoFacc_Transportar (userindex)
        
        RevivirUsuario (userindex)
        .Stats.MinHP = .Stats.MaxHP
        .Stats.MinMAN = .Stats.MaxMAN
        SendUserHP (userindex)
        SendUserMP (userindex)
    End With

End Sub
Public Sub EventoFacc_ContarMuerte(ByVal userindex As Integer)

    'Contamos una muerte y sumamos los segundos para revivir.
    UserList(userindex).flags.AramDeads = UserList(userindex).flags.AramDeads + 1
    UserList(userindex).flags.AramSeconds = UserList(userindex).flags.AramDeads * 2
    Call SendData(SendTarget.toindex, userindex, 0, "ARAM" & UserList(userindex).flags.AramSeconds)
    Call SendData(SendTarget.toindex, userindex, 0, "||902")

End Sub
Public Sub EventoFacc_restarTiempo()
    If (EventoFacc.Comenzado) Then
        If (EventoFacc.minutosBatalla > 0) Then
            EventoFacc.minutosBatalla = EventoFacc.minutosBatalla - 1
            If (EventoFacc.minutosBatalla <> 0) Then Call SendData(SendTarget.ToAll, 0, 0, "||976@" & EventoFacc.minutosBatalla)
        End If
        
        If (EventoFacc.minutosBatalla = 0) Then
            If MapaEventoFacc = 185 Then
                eventoFacc_Win ("Hordas")
            ElseIf MapaEventoFacc = 184 Then
                eventoFacc_Win ("Alianzas")
            End If
        End If
    End If
End Sub
Public Sub eventoFacc_Win(ByVal equipoGanador As String)

    If Not HayEventoFacc Then Exit Sub

    Dim loopC As Long
    
    EventoFacc.TeamHorda = 0
    EventoFacc.TeamAlianza = 0
    HayEventoFacc = False
    EventoFacc.Comenzado = False
    
    For loopC = 1 To LastUser
        If UserList(loopC).flags.EventoFacc Then
            If (equipoGanador = "Hordas" And (UserList(loopC).StatusMith.EsStatus = 2 Or EsHorda(loopC))) Or (equipoGanador = "Alianzas" And (UserList(loopC).StatusMith.EsStatus = 1 Or EsAlianza(loopC))) Then
                Call SendData(SendTarget.toindex, loopC, 0, "||900@1")
                UserList(loopC).Stats.TSPoints = UserList(loopC).Stats.TSPoints + 1
            End If
            
            Call EventoFacc_QuitarUsuario(loopC, False)
        End If
    Next loopC
    
    If EventoFacc.numRey > 0 Then Call QuitarNPC(EventoFacc.numRey)
    
    
    If equipoGanador = "Alianzas" Then
        EventoFacc.victoriasAlianza = EventoFacc.victoriasAlianza + 1
        Call WriteVar(IniPath & "Facciones.ini", "Jerarquias", "EventosAlianza", EventoFacc.victoriasAlianza)
        
        If MapaEventoFacc = 185 Then
            If EventoFacc.victoriasAlianza = EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||975@Alianzas@" & EventoFacc.victoriasAlianza)
            ElseIf EventoFacc.victoriasAlianza < EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||973@Alianzas@" & EventoFacc.victoriasAlianza & "@Hordas@" & EventoFacc.victoriasHorda)
            ElseIf EventoFacc.victoriasAlianza > EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||974@Alianzas@" & EventoFacc.victoriasAlianza & "@Hordas@" & EventoFacc.victoriasHorda)
            End If
        Else
            If EventoFacc.victoriasAlianza = EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||972@Alianzas@" & EventoFacc.victoriasAlianza)
            ElseIf EventoFacc.victoriasAlianza < EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||970@Alianzas@" & EventoFacc.victoriasAlianza & "@Hordas@" & EventoFacc.victoriasHorda)
            ElseIf EventoFacc.victoriasAlianza > EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||971@Alianzas@" & EventoFacc.victoriasAlianza & "@Hordas@" & EventoFacc.victoriasHorda)
            End If
        End If
        
    ElseIf equipoGanador = "Hordas" Then
        EventoFacc.victoriasHorda = EventoFacc.victoriasHorda + 1
        Call WriteVar(IniPath & "Facciones.ini", "Jerarquias", "EventosHorda", EventoFacc.victoriasHorda)
        
        If MapaEventoFacc = 184 Then
            If EventoFacc.victoriasAlianza = EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||975@Hordas@" & EventoFacc.victoriasHorda)
            ElseIf EventoFacc.victoriasAlianza > EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||973@Hordas@" & EventoFacc.victoriasHorda & "@Alianzas@" & EventoFacc.victoriasAlianza)
            ElseIf EventoFacc.victoriasAlianza < EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||974@Hordas@" & EventoFacc.victoriasHorda & "@Alianzas@" & EventoFacc.victoriasAlianza)
            End If
        Else
            If EventoFacc.victoriasAlianza = EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||972@Hordas@" & EventoFacc.victoriasHorda)
            ElseIf EventoFacc.victoriasAlianza > EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||970@Hordas@" & EventoFacc.victoriasHorda & "@Alianzas@" & EventoFacc.victoriasAlianza)
            ElseIf EventoFacc.victoriasAlianza < EventoFacc.victoriasHorda Then
                Call SendData(SendTarget.ToAll, 0, 0, "||971@Hordas@" & EventoFacc.victoriasHorda & "@Alianzas@" & EventoFacc.victoriasAlianza)
            End If
        End If
    
    End If
                


End Sub
Public Sub EventoFacc_QuitarUsuario(ByVal userindex As Integer, Optional g As Boolean = True)

        With UserList(userindex)
            
            If g Then
                If UserList(userindex).StatusMith.EsStatus = 1 Or EsAlianza(userindex) Then
                    EventoFacc.TeamAlianza = EventoFacc.TeamAlianza - 1
                    Call SendData(SendTarget.ToAll, 0, 0, "||965@" & .Name & "@Alianza@" & PonerPuntos(EventoFacc.costoInscripcion))
                ElseIf UserList(userindex).StatusMith.EsStatus = 2 Or EsHorda(userindex) Then
                    EventoFacc.TeamHorda = EventoFacc.TeamHorda - 1
                    Call SendData(SendTarget.ToAll, 0, 0, "||965@" & .Name & "@Horda@" & PonerPuntos(EventoFacc.costoInscripcion))
                End If
            End If
        
            .flags.EventoFacc = False
            .flags.AramSeconds = 0
            .flags.AramDeads = 0
            Call SendData(SendTarget.toindex, userindex, 0, "ARAM" & .flags.AramSeconds)
        End With
        
        Call WarpUserChar(userindex, 28, 56, 34)
        
End Sub
Public Sub EventoFacc_CambiaMapa(ByVal userindex As Integer)

        UserList(userindex).flags.EventoFacc = False
        UserList(userindex).flags.AramSeconds = 0
        UserList(userindex).flags.AramDeads = 0
        Call SendData(SendTarget.toindex, userindex, 0, "ARAM" & UserList(userindex).flags.AramSeconds)
        
End Sub
Public Sub EventoFacc_Cancelar()

    If Not HayEventoFacc Then Exit Sub

    Dim loopC As Long
    
    For loopC = 1 To LastUser
        If UserList(loopC).flags.EventoFacc Then
            Call EventoFacc_QuitarUsuario(loopC, False)
            UserList(loopC).Stats.GLD = UserList(loopC).Stats.GLD + EventoFacc.costoInscripcion
            SendUserGLD loopC
        End If
    Next loopC
    
    If EventoFacc.numRey > 0 Then Call QuitarNPC(EventoFacc.numRey)
    
    EventoFacc.Comenzado = False
    EventoFacc.TeamHorda = 0
    EventoFacc.TeamAlianza = 0
    HayEventoFacc = False
    
    Call SendData(SendTarget.ToAll, 0, 0, "||966@" & EventoFacc.CuposTotales)
End Sub
Private Sub EventoFacc_Comenzar()

    'Invocamos ambas torres al comenzar el evento
    Dim PosRey As WorldPos
    PosRey.Map = MapaEventoFacc
    
    If MapaEventoFacc = 185 Then
        PosRey.X = 50
        PosRey.Y = 27
        EventoFacc.numRey = SpawnNpc(967, PosRey, True, False)
    ElseIf MapaEventoFacc = 184 Then
        PosRey.X = 50
        PosRey.Y = 46
        EventoFacc.numRey = SpawnNpc(966, PosRey, True, False)
    End If
    
    'Desbloqueamos las salidas
    EventoFacc.Comenzado = True
    ComenzarEventoFacc = 6

End Sub
Public Sub EventoFacc_pasarSegundo()
    If ComenzarEventoFacc > 0 Then
        ComenzarEventoFacc = ComenzarEventoFacc - 1
        SendData SendTarget.toMap, 0, MapaEventoFacc, "CU" & ComenzarEventoFacc
        
        If ComenzarEventoFacc = 0 Then
            EventoFacc_DesbloquearBases
            Call SendData(SendTarget.ToAll, 0, 0, "||898@" & EventoFacc.CuposTotales)
        End If
    End If
End Sub
Private Sub EventoFacc_BloquearBases(ByVal userindex As Integer)

    Dim loopX As Long
    
    If (MapaEventoFacc = 185) Then
    
        For loopX = 0 To 4
            MapData(MapaEventoFacc, 60, 22 + loopX).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 60, 22 + loopX, 1)
            
            MapData(MapaEventoFacc, 37 + loopX, 62).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 37 + loopX, 62, 1)
        Next loopX
        
        MapData(MapaEventoFacc, 43, 41).Blocked = 1
        Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 43, 41, 1)
        MapData(MapaEventoFacc, 44, 41).Blocked = 1
        Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 44, 41, 1)
        
        MapData(MapaEventoFacc, 55, 41).Blocked = 1
        Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 55, 41, 1)
        MapData(MapaEventoFacc, 56, 41).Blocked = 1
        Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 56, 41, 1)
        
    ElseIf (MapaEventoFacc = 184) Then
    
        For loopX = 0 To 9
            MapData(MapaEventoFacc, 45 + loopX, 38).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 45 + loopX, 38, 1)
        Next loopX

        
        For loopX = 39 To 61
            MapData(MapaEventoFacc, loopX, 73).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, loopX, 73, 1)
        Next loopX
        
        MapData(MapaEventoFacc, 44, 58).Blocked = 1
        Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 44, 58, 1)
        MapData(MapaEventoFacc, 45, 58).Blocked = 1
        Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 45, 58, 1)
        
        MapData(MapaEventoFacc, 55, 58).Blocked = 1
        Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 55, 58, 1)
        MapData(MapaEventoFacc, 56, 58).Blocked = 1
        Call Bloquear(SendTarget.toindex, userindex, MapaEventoFacc, MapaEventoFacc, 56, 58, 1)
    
    End If
    
    
End Sub
Public Sub EventoFacc_DesbloquearBases()

    Dim loopX As Long
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.EventoFacc Then
            If (MapaEventoFacc = 185) Then
            
                For loopX = 0 To 4
                    If UserList(i).StatusMith.EsStatus = 1 Or EsAlianza(i) Then
                        MapData(MapaEventoFacc, 37 + loopX, 62).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 37 + loopX, 62, 0)
                        
                        MapData(MapaEventoFacc, 43, 41).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 43, 41, 0)
                        MapData(MapaEventoFacc, 44, 41).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 44, 41, 0)
                        
                        MapData(MapaEventoFacc, 55, 41).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 55, 41, 0)
                        MapData(MapaEventoFacc, 55, 41).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 56, 41, 0)
                        
                    ElseIf UserList(i).StatusMith.EsStatus = 2 Or EsHorda(i) Then
                        MapData(MapaEventoFacc, 60, 22 + loopX).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 60, 22 + loopX, 0)
                    End If
                Next loopX
                
            ElseIf (MapaEventoFacc = 184) Then
                
                If UserList(i).StatusMith.EsStatus = 1 Or EsAlianza(i) Then
                    For loopX = 0 To 9
                        MapData(MapaEventoFacc, 45 + loopX, 38).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 45 + loopX, 38, 0)
                    Next loopX
                End If
        
                If UserList(i).StatusMith.EsStatus = 2 Or EsHorda(i) Then
                    For loopX = 39 To 61
                        MapData(MapaEventoFacc, loopX, 73).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, loopX, 73, 0)
                    Next loopX
                    
                    MapData(MapaEventoFacc, 44, 58).Blocked = 0
                    Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 44, 58, 0)
                    MapData(MapaEventoFacc, 45, 58).Blocked = 0
                    Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 45, 58, 0)
                    
                    MapData(MapaEventoFacc, 55, 58).Blocked = 0
                    Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 55, 58, 0)
                    MapData(MapaEventoFacc, 56, 58).Blocked = 0
                    Call Bloquear(SendTarget.toindex, i, MapaEventoFacc, MapaEventoFacc, 56, 58, 0)
                End If
                
            End If
        End If
    Next i
End Sub

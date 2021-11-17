Attribute VB_Name = "modAram"
Option Explicit

Private MapaAram As Integer

Public ComenzarAram As Byte
Public HayAram As Boolean

Type tAram
    Comenzado As Boolean
    TeamRojo As Byte
    TeamAzul As Byte
    CuposTotales As Byte
    costoInscripcion As Long
    TorreRoja As Integer
    TorreAzul As Integer
End Type

Private Aram As tAram

Public Sub Aram_Inscripciones(ByVal Cupos As Byte, ByVal Inscripcion As Long, Optional ByVal Map As Byte = 0)

    If HayAram = True Then Exit Sub
    
    Dim i As Long
    
    Aram.TeamRojo = 0
    Aram.TeamAzul = 0
    HayAram = True
    
    If Map = 0 Or (Map <> 189 And Map <> 186) Then
        Dim tmpVar As Byte
        tmpVar = RandomNumber(1, 2)
        
        If tmpVar = 1 Then
            MapaAram = 189
        Else
            MapaAram = 186
        End If
    Else
        MapaAram = Map
    End If
    
    Aram.CuposTotales = Cupos
    Aram.costoInscripcion = Inscripcion
    
    For i = 1 To LastUser
        UserList(i).flags.EnAram = False
        UserList(i).flags.AramAzul = False
        UserList(i).flags.AramRojo = False
        UserList(i).flags.AramSeconds = 0
        UserList(i).flags.AramDeads = 0
    Next i
    
    Call SendData(SendTarget.ToAll, 0, 0, "||901@" & Aram.CuposTotales * 2 & "@" & PonerPuntos(Aram.costoInscripcion) & "@" & Aram.CuposTotales)
    
End Sub
Public Function Aram_Activo() As Boolean

    Aram_Activo = (Aram.TeamRojo = Aram.CuposTotales And Aram.TeamAzul = Aram.CuposTotales)

End Function
Public Sub Aram_Ingresar(ByVal userindex As Integer)
    
    'Mensaje de error en caso de estar en un mapa no permitido.
    If Not HayAram Then Exit Sub
    If MapaEspecial(userindex) Then Call SendData(SendTarget.toindex, userindex, 0, "||291"): Exit Sub
    If UserList(userindex).flags.Muerto = 1 Then Call SendData(SendTarget.toindex, userindex, 0, "||3"): Exit Sub
    If UserList(userindex).Stats.GLD < Aram.costoInscripcion Then Call SendData(SendTarget.toindex, userindex, 0, "||663"): Exit Sub
    
    If Aram.TeamRojo = Aram.CuposTotales And Aram.TeamAzul = Aram.CuposTotales Then Call SendData(SendTarget.toindex, userindex, 0, "||904"): Exit Sub
    
    If Aram.TeamRojo < Aram.CuposTotales Then
        Aram.TeamRojo = Aram.TeamRojo + 1
        UserList(userindex).flags.AramRojo = True
        Call SendData(SendTarget.ToAll, 0, 0, "||906@" & UserList(userindex).Name & "@Rojo")
    ElseIf Aram.TeamAzul < Aram.CuposTotales Then
        Aram.TeamAzul = Aram.TeamAzul + 1
        UserList(userindex).flags.AramAzul = True
        Call SendData(SendTarget.ToAll, 0, 0, "||906@" & UserList(userindex).Name & "@Azul")
    End If
    
    Aram_TransportarUser (userindex)
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - Aram.costoInscripcion
    SendUserGLD (userindex)
    SendUserVariant (userindex)
    
    Call Aram_BloquearBases(userindex)
    UserList(userindex).flags.EnAram = True
    If Aram.Comenzado Then Aram_DesbloquearBases
    If (Not Aram.Comenzado) And (Aram.TeamRojo = Aram.CuposTotales And Aram.TeamAzul = Aram.CuposTotales) Then Aram_Comenzar

End Sub
Public Sub Aram_RevivirUsuario(ByVal userindex As Integer)

    With UserList(userindex)
        Aram_TransportarUser (userindex)
        
        RevivirUsuario (userindex)
        .Stats.MinHP = .Stats.MaxHP
        .Stats.MinMAN = .Stats.MaxMAN
        SendUserHP (userindex)
        SendUserMP (userindex)
    End With

End Sub
Public Sub Aram_ContarMuerte(ByVal userindex As Integer)

    'Contamos una muerte y sumamos los segundos para revivir.
    UserList(userindex).flags.AramDeads = UserList(userindex).flags.AramDeads + 1
    UserList(userindex).flags.AramSeconds = UserList(userindex).flags.AramDeads * 2
    Call SendData(SendTarget.toindex, userindex, 0, "ARAM" & UserList(userindex).flags.AramSeconds)
    
    Call SendData(SendTarget.toindex, userindex, 0, "||902")

End Sub
Private Sub Aram_TransportarUser(ByVal userindex As Integer)

With UserList(userindex)
    If MapaAram = 189 Then
    
        If .flags.AramRojo Then
            WarpUserChar userindex, MapaAram, RandomNumber(44, 56), RandomNumber(20, 29)
        ElseIf .flags.AramAzul Then
            WarpUserChar userindex, MapaAram, RandomNumber(44, 56), RandomNumber(71, 79)
        End If
        
    ElseIf MapaAram = 186 Then
    
        If .flags.AramRojo Then
            WarpUserChar userindex, MapaAram, RandomNumber(17, 23), RandomNumber(29, 39)
        ElseIf .flags.AramAzul Then
            WarpUserChar userindex, MapaAram, RandomNumber(62, 70), RandomNumber(78, 85)
        End If
        
    End If
End With

End Sub
Public Sub Aram_KillTower(ByVal userindex As Integer)

    If Not HayAram Then Exit Sub

    Dim loopC As Long, equipoGanador As String

    If UserList(userindex).flags.AramRojo Then
        equipoGanador = "Rojo"
    Else
        equipoGanador = "Azul"
    End If
    
    Aram.TeamRojo = 0
    Aram.TeamAzul = 0
    HayAram = False
    Aram.Comenzado = False
    
    For loopC = 1 To LastUser
        If equipoGanador = "Rojo" And UserList(loopC).flags.AramRojo Then
            Call SendData(SendTarget.toindex, loopC, 0, "||900@1")
            UserList(loopC).Stats.TSPoints = UserList(loopC).Stats.TSPoints + 1
        ElseIf equipoGanador = "Azul" And UserList(loopC).flags.AramAzul Then
            Call SendData(SendTarget.toindex, loopC, 0, "||900@1")
            UserList(loopC).Stats.TSPoints = UserList(loopC).Stats.TSPoints + 1
        End If
        
        If UserList(loopC).flags.AramAzul Or UserList(loopC).flags.AramRojo Then
            Call Aram_QuitarUsuario(loopC, False)
        End If
    Next loopC
    
    If Aram.TorreRoja > 0 Then Call QuitarNPC(Aram.TorreRoja)
    If Aram.TorreAzul > 0 Then Call QuitarNPC(Aram.TorreAzul)
    Call SendData(SendTarget.ToAll, 0, 0, "||899@" & equipoGanador)

End Sub
Public Sub Aram_QuitarUsuario(ByVal userindex As Integer, Optional g As Boolean = True)

        With UserList(userindex)
            
            If g Then
                If .flags.AramAzul Then
                    Aram.TeamAzul = Aram.TeamAzul - 1
                    Call SendData(SendTarget.ToAll, 0, 0, "||916@" & .Name & "@Azul@" & PonerPuntos(Aram.costoInscripcion))
                ElseIf .flags.AramRojo Then
                    Aram.TeamRojo = Aram.TeamRojo - 1
                    Call SendData(SendTarget.ToAll, 0, 0, "||916@" & .Name & "@Rojo@" & PonerPuntos(Aram.costoInscripcion))
                End If
            End If
        
            .flags.EnAram = False
            .flags.AramAzul = False
            .flags.AramRojo = False
            .flags.AramSeconds = 0
            .flags.AramDeads = 0
            Call SendData(SendTarget.toindex, userindex, 0, "ARAM" & .flags.AramSeconds)
        End With
        
        Call WarpUserChar(userindex, 28, 56, 34)
        
End Sub
Public Sub Aram_CambiaMapa(ByVal userindex As Integer)

        UserList(userindex).flags.EnAram = False
        UserList(userindex).flags.AramAzul = False
        UserList(userindex).flags.AramRojo = False
        UserList(userindex).flags.AramSeconds = 0
        UserList(userindex).flags.AramDeads = 0
        Call SendData(SendTarget.toindex, userindex, 0, "ARAM" & UserList(userindex).flags.AramSeconds)
        
End Sub
Public Sub Aram_Cancelar()

    If Not HayAram Then Exit Sub

    Dim loopC As Long
    
    For loopC = 1 To LastUser
        If UserList(loopC).flags.AramAzul Or UserList(loopC).flags.AramRojo Then
            Call Aram_QuitarUsuario(loopC, False)
            UserList(loopC).Stats.GLD = UserList(loopC).Stats.GLD + Aram.costoInscripcion
            SendUserGLD loopC
        End If
    Next loopC
    
    If Aram.TorreRoja > 0 Then Call QuitarNPC(Aram.TorreRoja)
    If Aram.TorreAzul > 0 Then Call QuitarNPC(Aram.TorreAzul)
    
    Aram.Comenzado = False
    Aram.TeamRojo = 0
    Aram.TeamAzul = 0
    HayAram = False
    
    Call SendData(SendTarget.ToAll, 0, 0, "||905@" & Aram.CuposTotales)
End Sub
Private Sub Aram_Comenzar()

    'Invocamos ambas torres al comenzar el evento
    Dim PosTorre As WorldPos
    PosTorre.Map = MapaAram
    
    'Roja
    If MapaAram = 189 Then
        PosTorre.X = 50
        PosTorre.Y = 35
    ElseIf MapaAram = 186 Then
        PosTorre.X = 32
        PosTorre.Y = 41
    End If
    
    Aram.TorreRoja = SpawnNpc(963, PosTorre, True, False)
    
    'Azul
    If MapaAram = 189 Then
        PosTorre.X = 50
        PosTorre.Y = 65
    ElseIf MapaAram = 186 Then
        PosTorre.X = 59
        PosTorre.Y = 68
    End If
    
    Aram.TorreAzul = SpawnNpc(964, PosTorre, True, False)
    
    'Desbloqueamos las salidas
    Aram.Comenzado = True
    ComenzarAram = 6

End Sub
Public Sub aram_pasarSegundo()
    If ComenzarAram > 0 Then
        ComenzarAram = ComenzarAram - 1
        SendData SendTarget.toMap, 0, MapaAram, "CU" & ComenzarAram
        
        If ComenzarAram = 0 Then
            Aram_DesbloquearBases
            Call SendData(SendTarget.ToAll, 0, 0, "||898@" & Aram.CuposTotales)
        End If
    End If
End Sub
Private Sub Aram_BloquearBases(ByVal userindex As Integer)

    Dim loopX As Long
    
    If (MapaAram = 189) Then
    
        For loopX = 44 To 56
            MapData(MapaAram, loopX, 30).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaAram, MapaAram, loopX, 30, 1)
            MapData(MapaAram, loopX, 70).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaAram, MapaAram, loopX, 70, 1)
        Next loopX
        
    ElseIf (MapaAram = 186) Then
    
        For loopX = 0 To 11
            MapData(MapaAram, 30 - loopX, 30 + loopX).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaAram, MapaAram, 30 - loopX, 30 + loopX, 1)
            MapData(MapaAram, 70 - loopX, 70 + loopX).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaAram, MapaAram, 70 - loopX, 70 + loopX, 1)
        Next loopX
    
    End If
    
    
End Sub
Public Sub Aram_DesbloquearBases()

    Dim loopX As Long
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.EnAram Then
            If (MapaAram = 189) Then
            
                For loopX = 44 To 56
                    If UserList(i).flags.AramRojo Then
                        MapData(MapaAram, loopX, 30).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaAram, MapaAram, loopX, 30, 0)
                    ElseIf UserList(i).flags.AramAzul Then
                        MapData(MapaAram, loopX, 70).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaAram, MapaAram, loopX, 70, 0)
                    End If
                Next loopX
                
            ElseIf (MapaAram = 186) Then
                
                For loopX = 0 To 11
                    If UserList(i).flags.AramRojo Then
                        MapData(MapaAram, 30 - loopX, 30 + loopX).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaAram, MapaAram, 30 - loopX, 30 + loopX, 0)
                    ElseIf UserList(i).flags.AramAzul Then
                        MapData(MapaAram, 70 - loopX, 70 + loopX).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaAram, MapaAram, 70 - loopX, 70 + loopX, 0)
                    End If
                Next loopX
                
            End If
        End If
    Next i
End Sub


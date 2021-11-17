Attribute VB_Name = "modBatMistica"
Option Explicit

Private Const MapaBat As Byte = 191

Private Type tBatalla
    Activa As Boolean
    Comenzada As Boolean
    Cupos As Byte
    costoInscripcion As Long
    secondsConteo As Byte
    minutosBatalla As Byte
End Type

Private Type tEquipe
    Miembros As Byte
    Kills As Integer
End Type

Private Batalla As tBatalla
Private Equipos(1 To 4) As tEquipe
Public Sub batalla_restarTiempo()
    If (Batalla.Comenzada) Then
        If (Batalla.minutosBatalla > 0) Then
            Batalla.minutosBatalla = Batalla.minutosBatalla - 1
            If (Batalla.minutosBatalla <> 0) Then Call SendData(SendTarget.ToAll, 0, 0, "||926@" & Batalla.minutosBatalla)
        End If
        
        If (Batalla.minutosBatalla = 0) Then
            modBatMistica.finalizarBatalla
        End If
    End If
End Sub
Private Sub finalizarBatalla()
    If Not Batalla.Activa Then Exit Sub

    Dim loopC As Long, inGanador As Byte, equipoGanador As String
    
    'Sacamos el indice del equipo ganador
    inGanador = 1
    For loopC = 1 To 4
        If (Equipos(inGanador).Kills < Equipos(loopC).Kills) Then
            inGanador = loopC
        End If
    Next loopC
    
    'Ubicamos el nombre
    equipoGanador = queEquipo(inGanador)
    
    'Entregamos premios
    For loopC = 1 To LastUser
        If (UserList(loopC).flags.enBatalla) Then
            If (UserList(loopC).flags.teamNumber = inGanador) Then
               Call SendData(SendTarget.toindex, loopC, 0, "||900@1")
               UserList(loopC).Stats.TSPoints = UserList(loopC).Stats.TSPoints + 1
            End If
            
            batalla_QuitarUsuario (loopC)
        End If
    Next loopC
    
    Call SendData(SendTarget.ToAll, 0, 0, "||925@" & equipoGanador & "@" & Equipos(inGanador).Kills)
    resetBatalla
    
End Sub
Public Sub batalla_pasarSegundo()
    If Batalla.secondsConteo > 0 Then
        Batalla.secondsConteo = Batalla.secondsConteo - 1
        SendData SendTarget.ToMap, 0, MapaBat, "CU" & Batalla.secondsConteo
        
        If Batalla.secondsConteo = 0 Then
            'Call SendData(SendTarget.ToAll, 0, 0, "||921")
            Call SendData(SendTarget.ToAll, 0, 0, "N|El evento Batalla Mistica ha comenzado! El equipo que logre conseguir la mayor cantidad de asesinatos en un lapso de 7 minutos, será el ganador.~225~222~119")
            Batalla.Comenzada = True
            batalla_actualizarKills
            batalla_desbloquearBases
        End If
    End If
End Sub
Public Sub iniciarBatalla(ByVal Cupos As Byte, ByVal Inscripcion As Long)
    
    If (Not hayBatalla) Then
        Batalla.Activa = True
        Batalla.Comenzada = False
        Batalla.Cupos = Cupos
        Batalla.costoInscripcion = Inscripcion
        Batalla.minutosBatalla = 7
        resetEquipos
        
        Call SendData(SendTarget.ToAll, 0, 0, "||924@" & Batalla.Cupos * 4 & "@" & Batalla.costoInscripcion & "@" & Batalla.Cupos)
    End If

End Sub
Public Sub batalla_contarMuerte(ByVal Atacante As Integer, ByVal Victima As Integer)

    'Contamos una muerte y sumamos los segundos para revivir.
    UserList(Victima).flags.batDeads = UserList(Victima).flags.batDeads + 1
    UserList(Victima).flags.batSeconds = UserList(Victima).flags.batDeads * 2
    Call SendData(SendTarget.toindex, Victima, 0, "ARAM" & UserList(Victima).flags.batSeconds)
    Call SendData(SendTarget.toindex, Victima, 0, "||902")
    
    Equipos(UserList(Atacante).flags.teamNumber).Kills = Equipos(UserList(Atacante).flags.teamNumber).Kills + 1
    batalla_actualizarKills
    
End Sub
Public Function hayBatalla() As Boolean
    hayBatalla = Batalla.Activa
End Function
Private Sub batalla_actualizarKills()

    Dim i As Long
    For i = 1 To LastUser
        If (UserList(i).flags.enBatalla) Then Call SendData(SendTarget.toindex, i, 0, "BTM" & Batalla.Comenzada & "," & Equipos(1).Kills & "," & Equipos(2).Kills & "," & Equipos(3).Kills & "," & Equipos(4).Kills)
    Next i

End Sub
Public Sub ingresarBatalla(ByVal userindex As Integer)
    
    If Not hayBatalla Then Exit Sub
    If MapaEspecial(userindex) Then Call SendData(SendTarget.toindex, userindex, 0, "||291"): Exit Sub
    If UserList(userindex).flags.Muerto = 1 Then Call SendData(SendTarget.toindex, userindex, 0, "||3"): Exit Sub
    If UserList(userindex).Stats.GLD < Batalla.costoInscripcion Then Call SendData(SendTarget.toindex, userindex, 0, "||663"): Exit Sub
    
    Dim eLibre As Byte
    eLibre = equipoLibre
    
    If (eLibre = 0) Then Call SendData(SendTarget.toindex, userindex, 0, "||904"): Exit Sub
    
    'Ingresa
    UserList(userindex).flags.enBatalla = True
    UserList(userindex).flags.batSeconds = 0
    UserList(userindex).flags.batDeads = 0
    UserList(userindex).flags.teamNumber = eLibre
    Equipos(eLibre).Miembros = Equipos(eLibre).Miembros + 1
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - Batalla.costoInscripcion
    SendUserGLD (userindex)
    
    batalla_moverABase (userindex)
    batalla_bloquearBases (userindex)
    SendUserVariant (userindex)

    Call SendData(SendTarget.ToAll, 0, 0, "||923@" & UserList(userindex).Name & "@" & queEquipo(UserList(userindex).flags.teamNumber))
    
    If (equipoLibre = 0) Then
        Batalla.secondsConteo = 6
    End If
    
End Sub
Private Function queEquipo(ByVal teamNum As Integer) As String
    Select Case teamNum
        Case 1
            queEquipo = "Azul"
        Case 2
            queEquipo = "Amarillo"
        Case 3
            queEquipo = "Rojo"
        Case 4
            queEquipo = "Verde"
    End Select
End Function
Private Function equipoLibre() As Integer

    Dim eLib As Integer
    eLib = 0

    Dim i As Long
    For i = 1 To 4
        If (Equipos(i).Miembros < Batalla.Cupos) Then
            eLib = i
            Exit For
        End If
    Next i
    
    equipoLibre = eLib
End Function
Public Sub cancelarBatalla()

    Dim i As Long

    If (hayBatalla) Then
        For i = 1 To LastUser
            If (UserList(i).flags.enBatalla) Then
                UserList(i).Stats.GLD = UserList(i).Stats.GLD + Batalla.costoInscripcion
                SendUserGLD i
                batalla_QuitarUsuario (i)
            End If
        Next i

        resetBatalla
        Call SendData(SendTarget.ToAll, 0, 0, "||922")
    End If

End Sub
Public Sub batalla_revivirUsuario(ByVal userindex As Integer)
    
    With UserList(userindex)
        batalla_moverABase (userindex)
        RevivirUsuario (userindex)
        .Stats.MinHP = .Stats.MaxHP
        .Stats.MinMAN = .Stats.MaxMAN
        SendUserHP (userindex)
        SendUserMP (userindex)
    End With
    
End Sub
Private Sub batalla_moverABase(ByVal userindex As Integer)
    
    Select Case UserList(userindex).flags.teamNumber
        Case 1
            Call WarpUserChar(userindex, 191, RandomNumber(39, 44), RandomNumber(26, 31))
        Case 2
            Call WarpUserChar(userindex, 191, RandomNumber(56, 61), RandomNumber(26, 31))
        Case 3
            Call WarpUserChar(userindex, 191, RandomNumber(39, 44), RandomNumber(69, 74))
        Case 4
            Call WarpUserChar(userindex, 191, RandomNumber(56, 61), RandomNumber(69, 74))
        Case Else
            batalla_QuitarUsuario (userindex)
    End Select
        
    
End Sub
Public Sub batalla_QuitarUsuario(ByVal userindex As Integer)
    
    UserList(userindex).flags.batDeads = 0
    UserList(userindex).flags.batSeconds = 0
    UserList(userindex).flags.teamNumber = 0
    UserList(userindex).flags.enBatalla = False
    Call SendData(SendTarget.toindex, userindex, 0, "BTM" & False & ",0,0,0,0")
    Call SendData(SendTarget.toindex, userindex, 0, "ARAM" & UserList(userindex).flags.batSeconds)
    Call WarpUserChar(userindex, 28, 56, 34)

End Sub
Public Function batallaComenzada() As Boolean
    batallaComenzada = Batalla.Comenzada
End Function
Private Sub resetBatalla()
    Batalla.Activa = False
    Batalla.Comenzada = False
    Batalla.secondsConteo = 0
    Batalla.costoInscripcion = 0
    Batalla.Cupos = 0
    Batalla.minutosBatalla = 0
    EventosAutomaticos = 0
    resetEquipos
End Sub
Private Sub resetEquipos()
    
    Dim i As Long
    For i = 1 To 4
        Equipos(i).Miembros = 0
        Equipos(i).Kills = 0
    Next i
    
End Sub
Private Sub batalla_bloquearBases(ByVal userindex As Integer)

    Dim loopX As Long
    For loopX = 38 To 45
            MapData(MapaBat, loopX, 38).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaBat, MapaBat, loopX, 38, 1)
            MapData(MapaBat, loopX, 62).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaBat, MapaBat, loopX, 62, 1)
    Next loopX
    
    For loopX = 55 To 62
            MapData(MapaBat, loopX, 38).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaBat, MapaBat, loopX, 38, 1)
            MapData(MapaBat, loopX, 62).Blocked = 1
            Call Bloquear(SendTarget.toindex, userindex, MapaBat, MapaBat, loopX, 62, 1)
    Next loopX
            

End Sub
Private Sub batalla_desbloquearBases()

    Dim loopX As Long, i As Long
        
    For i = 1 To LastUser
        If (UserList(i).flags.enBatalla) Then
            Select Case UserList(i).flags.teamNumber
                Case 1
                    For loopX = 38 To 45
                        MapData(MapaBat, loopX, 38).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaBat, MapaBat, loopX, 38, 0)
                    Next loopX
                Case 2
                    For loopX = 55 To 62
                        MapData(MapaBat, loopX, 38).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaBat, MapaBat, loopX, 38, 0)
                    Next loopX
                Case 3
                    For loopX = 38 To 45
                        MapData(MapaBat, loopX, 62).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaBat, MapaBat, loopX, 62, 0)
                    Next loopX
                Case 4
                    For loopX = 55 To 62
                        MapData(MapaBat, loopX, 62).Blocked = 0
                        Call Bloquear(SendTarget.toindex, i, MapaBat, MapaBat, loopX, 62, 0)
                    Next loopX
            End Select
        End If
    Next i
End Sub

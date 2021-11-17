Attribute VB_Name = "mEventoLUZ"
Option Explicit

Private Type tUsuario
    ID As Integer
    X As Byte
    Y As Byte
End Type

Private Type tLUZ
    Activo As Boolean
    Comenzado As Boolean
    Usuarios(1 To 14) As tUsuario
    mapa As Integer
    Inscripcion As Long
    cantCupos As Byte
    cantParticipantes As Byte
    contadorSeg As Byte
    cuentaRegresiva As Byte
    numNPC As Integer
    indexNPC As Integer
    numAzar As Byte
    nRandom As Byte
    lastPick As Byte
End Type
 
Private evLuz As tLUZ
Public Sub Carga_evLUZ()
    Dim loopC As Long
 
    With evLuz
        .cantCupos = UBound(.Usuarios())
        .mapa = GetVar(App.Path & "\Dat\eventoLUZ.dat", "EVENTO", "Mapa")
        .numNPC = GetVar(App.Path & "\Dat\eventoLUZ.dat", "EVENTO", "NPC")
        
        For loopC = 1 To .cantCupos
            .Usuarios(loopC).ID = 0
            .Usuarios(loopC).X = GetVar(App.Path & "\Dat\eventoLUZ.dat", "USUARIO#" & loopC, "X")
            .Usuarios(loopC).Y = GetVar(App.Path & "\Dat\eventoLUZ.dat", "USUARIO#" & loopC, "Y")
        Next loopC
        
    End With
End Sub
Public Sub evLuz_Ingresar(ByVal ID As Integer)

    With evLuz
        If Not evLuz_puedeEntrar(ID) Or .Comenzado Then Exit Sub
        
        Call SendData(SendTarget.toindex, ID, 0, "||886")
        UserList(ID).Stats.GLD = UserList(ID).Stats.GLD - .Inscripcion
        SendUserGLD ID
        
        UserList(ID).flags.evLuz = True
        UserList(ID).flags.tmpPos = UserList(ID).Pos
        
        .cantParticipantes = .cantParticipantes + 1
        .Usuarios(.cantParticipantes).ID = ID
        WarpUserChar ID, .mapa, .Usuarios(.cantParticipantes).X, .Usuarios(.cantParticipantes).Y, False
        
        UserList(ID).flags.NotMove = 1
        Call SendData(SendTarget.toindex, ID, 0, "STOPD" & UserList(ID).flags.NotMove)

        Call SendData(SendTarget.ToAll, 0, 0, "||936@" & UserList(ID).Name)
        
        If .cantCupos = .cantParticipantes Then
            Dim tmpNPos As WorldPos
            tmpNPos.Map = .mapa
            tmpNPos.X = 50
            tmpNPos.Y = 50
            
            Call SendData(SendTarget.toMap, 0, .mapa, "||937")
            
            .indexNPC = SpawnNpc(.numNPC, tmpNPos, True, False)
            Npclist(.indexNPC).Char.AuraA = 3
            Call MakeNPCChar(SendTarget.toMap, 0, 0, .indexNPC, Npclist(.indexNPC).Pos.Map, Npclist(.indexNPC).Pos.X, Npclist(.indexNPC).Pos.Y)
            .contadorSeg = 20
        End If
        
    End With
End Sub
Private Function evLuz_puedeEntrar(ByVal ID As Integer) As Boolean
    evLuz_puedeEntrar = False
    
    If UserList(ID).flags.Muerto > 0 Then
        Call SendData(SendTarget.toindex, ID, 0, "||3")
        Exit Function
    End If
    
    If UserList(ID).flags.evLuz Then
        Call SendData(SendTarget.toindex, ID, 0, "||935")
        Exit Function
    End If
    
    If Not evLuz.Activo Then
        Call SendData(SendTarget.toindex, ID, 0, "||882")
        Exit Function
    End If
    
    If MapaEspecial(ID) Then
        Call SendData(SendTarget.toindex, ID, 0, "||291")
        Exit Function
    End If
    
    If evLuz.cantCupos <= evLuz.cantParticipantes Then
        Call SendData(SendTarget.toindex, ID, 0, "||883")
        Exit Function
    End If
    
    If UserList(ID).Stats.GLD < evLuz.Inscripcion Then
        Call SendData(SendTarget.toindex, ID, 0, "||663")
        Exit Function
    End If
    
    If MapInfo(UserList(ID).Pos.Map).Pk Then
        Call SendData(SendTarget.toindex, ID, 0, "||323")
        Exit Function
    End If
    
    evLuz_puedeEntrar = True
End Function
Public Sub Armar_evLuz(ByVal Cupos As Byte, ByVal Inscripcion As Long)
    With evLuz
        If .Activo = True Then Exit Sub
        
        Dim i As Long
        For i = 1 To LastUser
            UserList(i).flags.evLuz = False
        Next i
        
        .Inscripcion = Inscripcion
        .cantParticipantes = 0
        .cantCupos = Cupos
        .Activo = True
        .Comenzado = False
        .indexNPC = 0
        .lastPick = 0
        
        For i = 1 To .cantCupos
            .Usuarios(i).ID = 0
        Next i
            
        
        Call SendData(SendTarget.ToAll, 0, 0, "||931@" & .cantCupos & "@" & PonerPuntos(.Inscripcion))
    End With
End Sub
Public Function evLuz_Activo() As Boolean
    evLuz_Activo = evLuz.Activo
End Function
Private Sub evLuz_Finalizar()
    Dim loopC As Long
    Dim ID As Integer
    
    With evLuz
        For loopC = 1 To .cantCupos
            If (.Usuarios(loopC).ID > 0) Then ID = .Usuarios(loopC).ID: Exit For
        Next loopC
        
        Call SendData(SendTarget.ToAll, 0, 0, "||932@" & UserList(ID).Name)
        UserList(ID).flags.evLuz = False
        UserList(ID).flags.NotMove = 0
        Call SendData(SendTarget.toindex, ID, 0, "STOPD" & UserList(ID).flags.NotMove)
        
        UserList(ID).Stats.TSPoints = UserList(ID).Stats.TSPoints + 1
        Call SendData(SendTarget.toindex, ID, 0, "||900@1")
        
        WarpUserChar ID, UserList(ID).flags.tmpPos.Map, UserList(ID).flags.tmpPos.X, UserList(ID).flags.tmpPos.Y, False
        evLuz_Limpiar
    End With
End Sub
Public Sub evLuz_Cancelar()
    Dim loopC As Long, ID As Integer
    With evLuz
        If .Activo = False Then Exit Sub
        
        For loopC = 1 To .cantCupos
            ID = .Usuarios(loopC).ID
            If (ID > 0) Then
                If UserList(ID).flags.evLuz And UserList(ID).Pos.Map = .mapa Then
                    UserList(ID).flags.NotMove = 0
                    Call SendData(SendTarget.toindex, ID, 0, "STOPD" & UserList(ID).flags.NotMove)
                    UserList(ID).flags.evLuz = False
                    UserList(ID).Stats.GLD = UserList(ID).Stats.GLD + .Inscripcion
                    SendUserGLD ID
                    WarpUserChar ID, UserList(ID).flags.tmpPos.Map, UserList(ID).flags.tmpPos.X, UserList(ID).flags.tmpPos.Y, False
                End If
            End If
        Next loopC

        Call SendData(SendTarget.ToAll, 0, 0, "||933")
        evLuz_Limpiar
    End With
End Sub
Public Sub evLuz_Desconexion(ByVal ID As Integer)

    If (Not UserList(ID).flags.evLuz) Then Exit Sub
    
    With evLuz
        evLuz_quitarUsuario (ID)
        
            .cantParticipantes = .cantParticipantes - 1
            If .cantParticipantes <= 1 Then evLuz_Finalizar
            If .cantParticipantes > 1 Then Call SendData(SendTarget.toMap, 0, .mapa, "||934@" & UserList(ID).Name & "@" & .cantParticipantes)
    End With
End Sub
Private Sub evLuz_Limpiar()
    Dim loopC As Long
    With evLuz
        .Activo = False
        .Inscripcion = 0
        .cantParticipantes = 0
        .cantCupos = 0
        .Comenzado = False
        
        frmMain.evLuz.Enabled = False
        
        .cuentaRegresiva = 0
        .contadorSeg = 0
        .nRandom = 0
        .lastPick = 0
        
        If .indexNPC > 0 Then QuitarNPC (.indexNPC)
        
        For loopC = 1 To UBound(.Usuarios())
            .Usuarios(loopC).ID = 0
        Next loopC
    End With
End Sub
Public Sub evLuz_pasarSegundo()

    Dim loopC As Long, ID As Integer, tStr As String
    With evLuz
        If (Not .Comenzado) And (.contadorSeg > 0) Then
            .contadorSeg = .contadorSeg - 1
        
            Select Case .contadorSeg
                Case 19
                    tStr = "Bienvenido al evento Luz Maligna. A continuación explicaré brevemente la finalidad del evento."
                    evLuz_tipearNPC (tStr)
                    
                Case 14
                    tStr = "Se realizará una ruleta entre todos los participantes y escogeré a uno al azar. El participante elegido va a disponer de 10 segundos para escribir un número del 1 al 4."
                    evLuz_tipearNPC (tStr)
                    
                Case 9
                    tStr = "Yo pensaré un número, y en caso de que el usuario acierte dicho número, seguirá en juego."
                    evLuz_tipearNPC (tStr)
                    
                Case 4
                    tStr = "De lo contrario, en caso de no haber acertado el número, el jugador quedará inmediatamente fuera del juego y se continuará con los demás participantes."
                    evLuz_tipearNPC (tStr)
                
                Case 0
                    tStr = "Concretada la explicación, comencemos con el evento... ¡Mucha suerte a todos!"
                    .Comenzado = True
                    evLuz_tipearNPC (tStr)
            End Select
        End If
        
        If (.Comenzado) Then
            If (.nRandom = 0) Then
                .nRandom = RandomNumber(15, 30)
                
                frmMain.evLuz.interval = 700
                frmMain.evLuz.Enabled = True
            End If
            
            If (.cuentaRegresiva > 0) Then
                .cuentaRegresiva = .cuentaRegresiva - 1
                
                If (.cuentaRegresiva = 0) Then
                    tStr = UserList(.Usuarios(.lastPick).ID).Name & " quedó afuera del evento porque se le acabo el tiempo."
                    .nRandom = 0
                    evLuz_quitarUsuario (.lastPick)
                    evLuz_tipearNPC (tStr)
                Else
                    tStr = UserList(.Usuarios(.lastPick).ID).Name & " tienes " & .cuentaRegresiva & " segundos para escribir un número del 1 al 4. Si aciertas puedes seguir jugando, de lo contrario perderás el evento."
                    evLuz_tipearNPC (tStr)
                End If
            End If
        End If
            
        
    End With
End Sub
Public Sub evLuz_prenderLuz()

    Dim tmpLast As Byte
    tmpLast = 0
    
    With evLuz
    
        If (.lastPick > 0) Then
            Call SendData(SendTarget.toMap, 0, .mapa, "PCB" & .Usuarios(.lastPick).X & "," & .Usuarios(.lastPick).Y)
            tmpLast = .lastPick
        Else
            .lastPick = 1
        End If
        
        Do While (.cantParticipantes > 1) And (tmpLast = .lastPick Or .Usuarios(.lastPick).ID = 0)
            .lastPick = .lastPick + 1
            If (.lastPick > 14) Then .lastPick = 1
        Loop
        
        Call SendData(SendTarget.toMap, 0, .mapa, "PCL" & .Usuarios(.lastPick).X & "," & .Usuarios(.lastPick).Y & ",1,255,255,255")
        
    End With
    

End Sub
Public Sub evLuz_escogerUsuario()

    
    With evLuz
        .numAzar = RandomNumber(1, 4)
        .cuentaRegresiva = 10
        frmMain.evLuz.Enabled = False
    End With

End Sub
Public Sub evLuz_getText(ByVal userindex As Integer, ByVal str As String)

    Dim texto As String
    
    If (str <> "1") And (str <> "2") And (str <> "3") And (str <> "4") And (str <> "5") Then Exit Sub

    With evLuz
        
        If (.Usuarios(.lastPick).ID = userindex And .cuentaRegresiva > 0) Then
            If (val(str) = .numAzar) Then
                texto = UserList(userindex).Name & " eligió el número " & val(str) & ", cuando el correcto era " & .numAzar & ", por lo tanto, sigue siendo participe del torneo."
            Else
                texto = UserList(userindex).Name & " eligió el número " & val(str) & ", cuando el correcto era " & .numAzar & ", por lo tanto, quedó descalificado del evento."
                evLuz_quitarUsuario (.lastPick)
            End If
            
                
            evLuz_tipearNPC (texto)
            .nRandom = 0
            .cuentaRegresiva = 0
        End If
    End With

End Sub
Private Sub evLuz_tipearNPC(tStr As String)

With evLuz
    Call SendData(SendTarget.ToNPCArea, .indexNPC, Npclist(.indexNPC).Pos.Map, "N|" & vbWhite & "°" & tStr & "°" & CStr(Npclist(.indexNPC).Char.CharIndex))
End With

End Sub
Private Sub evLuz_quitarUsuario(ByVal ID As Byte)

    Dim tmpUI As Integer
    
    With evLuz
        tmpUI = .Usuarios(ID).ID
        
        .Usuarios(ID).ID = 0
        
        .cantParticipantes = .cantParticipantes - 1
        If .cantParticipantes <= 1 Then evLuz_Finalizar
        
        If .lastPick = ID Then Call SendData(SendTarget.toMap, 0, .mapa, "PCB" & .Usuarios(ID).X & "," & .Usuarios(ID).Y)
        
        Call SendData(SendTarget.toindex, tmpUI, 0, "||938")
        UserList(tmpUI).flags.evLuz = False
        UserList(tmpUI).flags.NotMove = 0
        Call SendData(SendTarget.toindex, tmpUI, 0, "STOPD" & UserList(tmpUI).flags.NotMove)
        WarpUserChar tmpUI, UserList(tmpUI).flags.tmpPos.Map, UserList(tmpUI).flags.tmpPos.X, UserList(tmpUI).flags.tmpPos.Y, False
    End With

End Sub

Public Function evLuz_getRandom() As Byte
    evLuz_getRandom = evLuz.nRandom
End Function

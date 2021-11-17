Attribute VB_Name = "Mod_Ranking"
Option Explicit
 
 
Public Const MAX_TOP As Byte = 10
Public Const MAX_RANKINGS As Byte = 9
 
Public Type tRanking
    Value(1 To MAX_TOP) As Long
    Nombre(1 To MAX_TOP) As String
End Type
 
Public Ranking(1 To MAX_RANKINGS) As tRanking
 
Public Enum eRanking
    TOPFrags = 1
    TOPTorneos = 2
    TOPDuelos = 3
    TOPParejas = 4
    TOPReputacion = 5
    TOPRondas = 6
    TOPCVCS = 7
    TOPCastillos = 8
    TOPRepuClanes = 9
End Enum
Public Function RenameRanking(ByVal Ranking As eRanking) As String
 
    '@ Devolvemos el nombre del TAG [] del archivo .DAT
    Select Case Ranking
        Case eRanking.TOPFrags
            RenameRanking = "FRAGS"
        Case eRanking.TOPTorneos
            RenameRanking = "TORNEOS"
        Case eRanking.TOPDuelos
            RenameRanking = "DUELOS"
        Case eRanking.TOPParejas
            RenameRanking = "PAREJAS"
        Case eRanking.TOPReputacion
            RenameRanking = "REPUTACION"
        Case eRanking.TOPRondas
            RenameRanking = "RONDAS"
        Case eRanking.TOPCVCS
            RenameRanking = "CVCS"
        Case eRanking.TOPCastillos
            RenameRanking = "CASTILLOS"
        Case eRanking.TOPRepuClanes
            RenameRanking = "REPUCLAN"
        Case Else
            RenameRanking = vbNullString
    End Select
End Function
Public Function RenameValue(ByVal userindex As Integer, ByVal Ranking As eRanking) As Long
    ' @ Devolvemos a que hace referencia el ranking
    With UserList(userindex)
        Select Case Ranking
            Case eRanking.TOPFrags
                RenameValue = .Stats.UsuariosMatados
            Case eRanking.TOPTorneos
                RenameValue = .Stats.TrofOro + .Stats.MedOro
            Case eRanking.TOPDuelos
                RenameValue = .Stats.DuelosGanados
            Case eRanking.TOPParejas
                RenameValue = .Stats.ParejasGanadas
            Case eRanking.TOPReputacion
                RenameValue = .Stats.Reputacione
            Case eRanking.TOPRondas
                RenameValue = val(.flags.rondas)
            Case eRanking.TOPCVCS
                RenameValue = Guilds(.GuildIndex).CVCG
            Case eRanking.TOPCastillos
                RenameValue = Guilds(.GuildIndex).CASTIS
            Case eRanking.TOPRepuClanes
                RenameValue = Guilds(.GuildIndex).GetReputacion
        End Select
    End With
End Function
 
Public Sub LoadRanking()
    ' @ Cargamos los rankings
   
    Dim LoopI As Integer
    Dim loopX As Integer
    Dim ln As String
   
    For loopX = 1 To MAX_RANKINGS
        For LoopI = 1 To MAX_TOP
            ln = GetVar(App.Path & "\Dat\" & "Ranking.dat", RenameRanking(loopX), "Top" & LoopI)
            Ranking(loopX).Nombre(LoopI) = ReadField(1, ln, 45)
            Ranking(loopX).Value(LoopI) = val(ReadField(2, ln, 45))
            
            If LoopI = 1 Or LoopI = 2 Or LoopI = 3 Then
                If loopX = TOPDuelos Then Estrella.TOPDuelos(LoopI) = Ranking(loopX).Nombre(LoopI)
                If loopX = TOPFrags Then Estrella.TOPFrags(LoopI) = Ranking(loopX).Nombre(LoopI)
                If loopX = TOPTorneos Then Estrella.TOPTorneos(LoopI) = Ranking(loopX).Nombre(LoopI)
                If loopX = TOPParejas Then Estrella.TOPParejas(LoopI) = Ranking(loopX).Nombre(LoopI)
                If loopX = TOPRondas Then Estrella.TOPRondas(LoopI) = Ranking(loopX).Nombre(LoopI)
                If loopX = TOPReputacion Then Estrella.TOPReputacion(LoopI) = Ranking(loopX).Nombre(LoopI)
            End If
        Next LoopI
    Next loopX
   
End Sub
   
Public Sub SaveRanking(ByVal Rank As eRanking)
' @ Guardamos el ranking
    Dim LoopI As Integer
   
        For LoopI = 1 To MAX_TOP
            Call WriteVar(DatPath & "Ranking.Dat", RenameRanking(Rank), _
                "Top" & LoopI, Ranking(Rank).Nombre(LoopI) & "-" & Ranking(Rank).Value(LoopI))
        Next LoopI
End Sub
Public Sub CheckRankingUser(ByVal userindex As Integer, ByVal Valor As Long, ByVal Rank As eRanking)
    ' @ Desde aca nos hacemos la siguientes preguntas
    ' @ El personaje está en el ranking?
    ' @ El personaje puede ingresar al ranking?
   
    Dim loopX As Integer
    Dim LoopY As Integer
    Dim LoopZ As Integer
    Dim i As Integer
    Dim Value As Long
    Dim Actualizacion As Byte
    Dim Auxiliar As String
    Dim AuxiliarName As String
    Dim PosRanking As Byte
    Dim NameUser As String
   
    With UserList(userindex)
       
        ' @ Not gms
        NameUser = UCase$(.Name)
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Exit Sub
       
        Value = Valor
       
        ' @ Buscamos al personaje en el ranking
        For i = 1 To MAX_TOP
            If Ranking(Rank).Nombre(i) = NameUser Then
                PosRanking = i
                
                    If i = 1 Or i = 2 Or i = 3 Then
                       If Rank = TOPDuelos Then Estrella.TOPDuelos(i) = NameUser
                       If Rank = TOPFrags Then Estrella.TOPFrags(i) = NameUser
                       If Rank = TOPTorneos Then Estrella.TOPTorneos(i) = NameUser
                       If Rank = TOPParejas Then Estrella.TOPParejas(i) = NameUser
                       If Rank = TOPRondas Then Estrella.TOPRondas(i) = NameUser
                       If Rank = TOPReputacion Then Estrella.TOPReputacion(i) = NameUser
                    End If
                
                Exit For
            End If
        Next i
       
        ' @ Si el personaje esta en el ranking actualizamos los valores.
        If PosRanking <> 0 Then
            ' ¿Si está actualizado pa que?
            If Value >= Ranking(Rank).Value(PosRanking) Then
                Call ActualizarPosRanking(PosRanking, Rank, NameUser, Value)
                   
                ' @ Chequeamos los datos para actualizar el ranking
                For LoopY = 1 To MAX_TOP
                    For LoopZ = 1 To MAX_TOP - LoopY
                           
                        If Ranking(Rank).Value(LoopZ) < Ranking(Rank).Value(LoopZ + 1) Then
                            Auxiliar = Ranking(Rank).Value(LoopZ)
                            AuxiliarName = Ranking(Rank).Nombre(LoopZ)
                            
                            Ranking(Rank).Nombre(LoopZ) = Ranking(Rank).Nombre(LoopZ + 1)
                            Ranking(Rank).Value(LoopZ) = Ranking(Rank).Value(LoopZ + 1)
                            
                            Ranking(Rank).Nombre(LoopZ + 1) = AuxiliarName
                            Ranking(Rank).Value(LoopZ + 1) = Auxiliar
                            Actualizacion = 1
                        End If
                    Next LoopZ
                Next LoopY
                   
                If Actualizacion <> 0 Then
                    Call SaveRanking(Rank)
                End If
            End If
           
            Exit Sub
        End If
       
        ' @ Nos fijamos si podemos ingresar al ranking
        For loopX = 1 To MAX_TOP
            If Value > Ranking(Rank).Value(loopX) Then
                Call ActualizarRanking(loopX, Rank, NameUser, Value)
                Exit For
            End If
        Next loopX
       
    End With
End Sub
 Public Sub CheckRankingClan(ByVal userindex As Integer, ByVal Valor As Long, ByVal Rank As eRanking)
    ' @ Desde aca nos hacemos la siguientes preguntas
    ' @ El personaje está en el ranking?
    ' @ El personaje puede ingresar al ranking?
   
    Dim loopX As Integer
    Dim LoopY As Integer
    Dim LoopZ As Integer
    Dim i As Integer
    Dim Value As Long
    Dim Actualizacion As Byte
    Dim Auxiliar As String
    Dim AuxiliarName As String
    Dim PosRanking As Byte
    Dim NameUser As String
   
    With UserList(userindex)
       
        ' @ Not gms
        NameUser = UCase$(Guilds(.GuildIndex).GuildName)
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Exit Sub
       
        Value = Valor
       
        ' @ Buscamos al personaje en el ranking
        For i = 1 To MAX_TOP
            If UCase$(Ranking(Rank).Nombre(i)) = NameUser Then
                PosRanking = i
                Exit For
            End If
        Next i
       
        ' @ Si el personaje esta en el ranking actualizamos los valores.
        If PosRanking <> 0 Then
            ' ¿Si está actualizado pa que?
            If Value <> Ranking(Rank).Value(PosRanking) Then
                Call ActualizarPosRanking(PosRanking, Rank, NameUser, Value)
                   
                ' @ Chequeamos los datos para actualizar el ranking
                For LoopY = 1 To MAX_TOP
                    For LoopZ = 1 To MAX_TOP - LoopY
                           
                        If Ranking(Rank).Value(LoopZ) < Ranking(Rank).Value(LoopZ + 1) Then
                            Auxiliar = Ranking(Rank).Value(LoopZ)
                            AuxiliarName = Ranking(Rank).Nombre(LoopZ)
                            
                            Ranking(Rank).Nombre(LoopZ) = Ranking(Rank).Nombre(LoopZ + 1)
                            Ranking(Rank).Value(LoopZ) = Ranking(Rank).Value(LoopZ + 1)
                            
                            Ranking(Rank).Nombre(LoopZ + 1) = AuxiliarName
                            Ranking(Rank).Value(LoopZ + 1) = Auxiliar
                            Actualizacion = 1
                        End If
                    Next LoopZ
                Next LoopY
                   
                If Actualizacion <> 0 Then
                    Call SaveRanking(Rank)
                End If
            End If
           
            Exit Sub
        End If
       
        ' @ Nos fijamos si podemos ingresar al ranking
        For loopX = 1 To MAX_TOP
            If Value > Ranking(Rank).Value(loopX) Then
                Call ActualizarRanking(loopX, Rank, NameUser, Value)
                Exit For
            End If
        Next loopX
       
    End With
End Sub
Public Sub ActualizarPosRanking(ByVal Top As Byte, ByVal Rank As eRanking, ByVal UserName As String, ByVal Value As Long)
    ' @ Actualizamos la pos indicada en caso de que el personaje esté en el ranking
    Dim loopX As Integer
 
    With Ranking(Rank)
    
    If Value > .Value(Top) Then
        .Value(Top) = Value
        .Nombre(Top) = UserName
    End If
    
        If Top = 1 Or Top = 2 Or Top = 3 Then
           If Rank = TOPDuelos Then Estrella.TOPDuelos(Top) = UserName
           If Rank = TOPFrags Then Estrella.TOPFrags(Top) = UserName
           If Rank = TOPTorneos Then Estrella.TOPTorneos(Top) = UserName
           If Rank = TOPParejas Then Estrella.TOPParejas(Top) = UserName
           If Rank = TOPRondas Then Estrella.TOPRondas(Top) = UserName
           If Rank = TOPReputacion Then Estrella.TOPReputacion(Top) = UserName
        End If
        
    End With
    
    If Rank = TOPTorneos And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@torneos@" & Value)
    If Rank = TOPDuelos And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@duelos@" & Value)
    If Rank = TOPFrags And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@usuarios matados@" & Value)
    If Rank = TOPRondas And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@rondas@" & Value)
    If Rank = TOPParejas And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@parejas@" & Value)
    If Rank = TOPReputacion And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@reputacion@" & Value)
    
End Sub
Public Sub ActualizarRanking(ByVal Top As Byte, ByVal Rank As eRanking, ByVal UserName As String, ByVal Value As Long)
   
    '@ Actualizamos la lista de ranking
   
    Dim loopC As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Valor(1 To MAX_TOP) As Long
    Dim Nombre(1 To MAX_TOP) As String
    
    ' @ Copia necesaria para evitar que se dupliquen repetidamente
    For loopC = 1 To MAX_TOP
        Valor(loopC) = Ranking(Rank).Value(loopC)
        Nombre(loopC) = Ranking(Rank).Nombre(loopC)
    Next loopC
   
    ' @ Corremos las pos, desde el "Top" que es la primera
    For loopC = Top To MAX_TOP - 1
        Ranking(Rank).Value(loopC + 1) = Valor(loopC)
        Ranking(Rank).Nombre(loopC + 1) = Nombre(loopC)
    Next loopC
   
    Ranking(Rank).Nombre(Top) = UserName
    Ranking(Rank).Value(Top) = Value
    
        If Top = 1 Or Top = 2 Or Top = 3 Then
           If (NameIndex(UserName)) Then sendUserRank (NameIndex(UserName))
        End If
    
    Call SaveRanking(Rank)

    If Rank = TOPTorneos And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@torneos@" & Value)
    If Rank = TOPCastillos And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||634@" & UserName & "@castillos conquistados@" & Value)
    If Rank = TOPCVCS And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@cvcs@" & Value)
    If Rank = TOPRepuClanes And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||634@" & UserName & "@reputación de clanes@" & Value)
    If Rank = TOPDuelos And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@duelos@" & Value)
    If Rank = TOPFrags And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@usuarios matados@" & Value)
    If Rank = TOPRondas And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@rondas@" & Value)
    If Rank = TOPParejas And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@parejas@" & Value)
    If Rank = TOPReputacion And Top = 1 Then Call SendData(SendTarget.ToAll, 0, 0, "||633@" & UserName & "@reputacion@" & Value)

End Sub
Public Function tieneRanking(ByVal userindex As Integer) As Byte

    Dim pRank As Byte, i, j As Long
    pRank = 99
    
    For i = 1 To MAX_RANKINGS
        For j = 1 To 3
            If (UCase$(Ranking(i).Nombre(j)) = UCase$(UserList(userindex).Name)) Then
                If (j < pRank) Then pRank = j
            End If
        Next j
    Next i

    tieneRanking = pRank

End Function


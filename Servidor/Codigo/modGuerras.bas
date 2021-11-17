Attribute VB_Name = "modGuerras"
Option Explicit

        Dim ReyH As Integer
        Dim GuardianH As Integer
        Dim BarreraH As Integer
        Dim ReyHorda As WorldPos
        Dim GuardianHorda1 As WorldPos
        Dim GuardianHorda2 As WorldPos
        Dim GuardianHorda3 As WorldPos
        Dim GuardianHorda4 As WorldPos
        Dim BarreraHorda As WorldPos

        Dim ReyA As Integer
        Dim GuardianA As Integer
        Dim BarreraA As Integer
        Dim ReyAlianza As WorldPos
        Dim GuardianAlianza1 As WorldPos
        Dim GuardianAlianza2 As WorldPos
        Dim GuardianAlianza3 As WorldPos
        Dim GuardianAlianza4 As WorldPos
        Dim BarreraAlianza As WorldPos
        
        Dim i As Long
        Dim x As Integer
        Dim Y As Integer
Public Function InvocarNPCs()
        ReyA = 956 'Rey Alianza Nº
        GuardianA = 958 'Guardian Alianza Nº
        BarreraA = 618
        
        ReyAlianza.Map = 164
        ReyAlianza.x = 67
        ReyAlianza.Y = 17
        
        BarreraAlianza.Map = 164
        BarreraAlianza.x = 31
        BarreraAlianza.Y = 32
        
        GuardianAlianza1.Map = 164
        GuardianAlianza1.x = 65
        GuardianAlianza1.Y = 15
        
        GuardianAlianza2.Map = 164
        GuardianAlianza2.x = 69
        GuardianAlianza2.Y = 15
    
        GuardianAlianza3.Map = 164
        GuardianAlianza3.x = 65
        GuardianAlianza3.Y = 19
        
        GuardianAlianza4.Map = 164
        GuardianAlianza4.x = 69
        GuardianAlianza4.Y = 19
        
        Call SpawnNpc(ReyA, ReyAlianza, True, False)
        Call SpawnNpc(BarreraA, BarreraAlianza, True, False)
        Call SpawnNpc(GuardianA, GuardianAlianza1, True, False)
        Call SpawnNpc(GuardianA, GuardianAlianza2, True, False)
        Call SpawnNpc(GuardianA, GuardianAlianza3, True, False)
        Call SpawnNpc(GuardianA, GuardianAlianza4, True, False)
        
        ReyH = 957 'Rey Horda Nº
        GuardianH = 959 'Guardian Horda Nº
        BarreraH = 617
        
        ReyHorda.Map = 164
        ReyHorda.x = 31
        ReyHorda.Y = 17
        
        BarreraHorda.Map = 164
        BarreraHorda.x = 67
        BarreraHorda.Y = 32
        
        GuardianHorda1.Map = 164
        GuardianHorda1.x = 29
        GuardianHorda1.Y = 15
        
        GuardianHorda2.Map = 164
        GuardianHorda2.x = 33
        GuardianHorda2.Y = 15
    
        GuardianHorda3.Map = 164
        GuardianHorda3.x = 29
        GuardianHorda3.Y = 19
        
        GuardianHorda4.Map = 164
        GuardianHorda4.x = 33
        GuardianHorda4.Y = 19
        
        Call SpawnNpc(ReyH, ReyHorda, True, False)
        Call SpawnNpc(BarreraH, BarreraHorda, True, False)
        Call SpawnNpc(GuardianH, GuardianHorda1, True, False)
        Call SpawnNpc(GuardianH, GuardianHorda2, True, False)
        Call SpawnNpc(GuardianH, GuardianHorda3, True, False)
        Call SpawnNpc(GuardianH, GuardianHorda4, True, False)
        
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 24 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 23 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 22 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 21 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 25 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 26 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 27 & "," & 76 & "," & 3 & "," & 2000)
        
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 67 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 66 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 65 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 64 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 68 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 69 & "," & 76 & "," & 3 & "," & 2000)
        Call SendData(SendTarget.ToMap, 0, 164, "PCF" & 70 & "," & 76 & "," & 3 & "," & 2000)
        
End Function
Public Function FinalizarGuerra(Ganador As String)

If UCase$(Ganador) = "NADIE" Then
 For i = 1 To LastUser
    If UserList(i).flags.EnGuerra = 1 Then
      If UserList(i).StatusMith.EsStatus = 3 Then
        Call WarpUserChar(i, 29, 46, 84, True)
      ElseIf UserList(i).StatusMith.EsStatus = 4 Then
        Call WarpUserChar(i, 27, 46, 52, True)
      End If
      UserList(i).flags.EnGuerra = 0
    End If
 Next i

    For Y = 7 To 80
      For x = 7 To 80
        If x > 0 And Y > 0 And x < 101 And Y < 101 Then _
            If MapData(164, x, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(164, x, Y).NpcIndex)
      Next x
    Next Y

 SendData SendTarget.toall, 0, 0, "||El tiempo para la guerra ha finalizado sus Generales se sienten defraudados por no vencer a su enemigo." & FONTTYPE_ROJON
End If


 MinutosGuerras = 0
 HayGuerra = False

End Function
Public Function DesbloquearSalida()
'DESBLOQUEAMOS TODO
    x = 24
    Y = 76
    
    MapData(164, x, Y).Blocked = 0
    MapData(164, x - 1, Y).Blocked = 0
    MapData(164, x - 2, Y).Blocked = 0
    MapData(164, x - 3, Y).Blocked = 0
    MapData(164, x + 1, Y).Blocked = 0
    MapData(164, x + 2, Y).Blocked = 0
    MapData(164, x + 3, Y).Blocked = 0

    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 1, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 2, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 3, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 1, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 2, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 3, Y, 0)
    
    x = 75
    Y = 76
    
    MapData(164, x, Y).Blocked = 0
    MapData(164, x - 1, Y).Blocked = 0
    MapData(164, x - 2, Y).Blocked = 0
    MapData(164, x - 3, Y).Blocked = 0
    MapData(164, x + 1, Y).Blocked = 0
    MapData(164, x + 2, Y).Blocked = 0
    MapData(164, x + 3, Y).Blocked = 0

    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 1, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 2, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 3, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 1, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 2, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 3, Y, 0)
End Function
Public Function BloquearSalida()

'salida ali
x = 24
Y = 76

    MapData(164, x, Y).Blocked = 1
    MapData(164, x - 1, Y).Blocked = 1
    MapData(164, x - 2, Y).Blocked = 1
    MapData(164, x - 3, Y).Blocked = 1
    MapData(164, x + 1, Y).Blocked = 1
    MapData(164, x + 2, Y).Blocked = 1
    MapData(164, x + 3, Y).Blocked = 1
    
    'Bloquea todos los mapas
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 1, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 2, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 3, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 1, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 2, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 3, Y, 1)
    
    
'salida horda
x = 75
Y = 76

    MapData(164, x, Y).Blocked = 1
    MapData(164, x - 1, Y).Blocked = 1
    MapData(164, x - 2, Y).Blocked = 1
    MapData(164, x - 3, Y).Blocked = 1
    MapData(164, x + 1, Y).Blocked = 1
    MapData(164, x + 2, Y).Blocked = 1
    MapData(164, x + 3, Y).Blocked = 1
    
    'Bloquea todos los mapas
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 1, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 2, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 3, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 1, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 2, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 3, Y, 1)
    
'barreras
x = 31
Y = 32

    MapData(164, x, Y).Blocked = 1
    MapData(164, x - 1, Y).Blocked = 1
    MapData(164, x - 2, Y).Blocked = 1
    MapData(164, x - 3, Y).Blocked = 1
    MapData(164, x - 4, Y).Blocked = 1
    MapData(164, x + 1, Y).Blocked = 1
    MapData(164, x + 2, Y).Blocked = 1
    MapData(164, x + 3, Y).Blocked = 1
    MapData(164, x + 4, Y).Blocked = 1
    
    'Bloquea todos los mapas
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 1, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 2, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 3, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 4, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 1, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 2, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 3, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 4, Y, 1)

'barreras
x = 67
Y = 32

    MapData(164, x, Y).Blocked = 1
    MapData(164, x - 1, Y).Blocked = 1
    MapData(164, x - 2, Y).Blocked = 1
    MapData(164, x - 3, Y).Blocked = 1
    MapData(164, x - 4, Y).Blocked = 1
    MapData(164, x + 1, Y).Blocked = 1
    MapData(164, x + 2, Y).Blocked = 1
    MapData(164, x + 3, Y).Blocked = 1
    MapData(164, x + 4, Y).Blocked = 1
    
    'Bloquea todos los mapas
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 1, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 2, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 3, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 4, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 1, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 2, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 3, Y, 1)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 4, Y, 1)
    
End Function
Public Function DesbloquearBarreraAlianza()

    x = 31
    Y = 32

    MapData(164, x, Y).Blocked = 0
    MapData(164, x - 1, Y).Blocked = 0
    MapData(164, x - 2, Y).Blocked = 0
    MapData(164, x - 3, Y).Blocked = 0
    MapData(164, x - 4, Y).Blocked = 0
    MapData(164, x + 1, Y).Blocked = 0
    MapData(164, x + 2, Y).Blocked = 0
    MapData(164, x + 3, Y).Blocked = 0
    MapData(164, x + 4, Y).Blocked = 0
    
    'Bloquea todos los mapas
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 1, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 2, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 3, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 4, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 1, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 2, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 3, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 4, Y, 0)

End Function
Public Function DesbloquearBarreraHorda()

    x = 67
    Y = 32

    MapData(164, x, Y).Blocked = 0
    MapData(164, x - 1, Y).Blocked = 0
    MapData(164, x - 2, Y).Blocked = 0
    MapData(164, x - 3, Y).Blocked = 0
    MapData(164, x - 4, Y).Blocked = 0
    MapData(164, x + 1, Y).Blocked = 0
    MapData(164, x + 2, Y).Blocked = 0
    MapData(164, x + 3, Y).Blocked = 0
    MapData(164, x + 4, Y).Blocked = 0
    
    'Bloquea todos los mapas
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 1, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 2, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 3, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x - 4, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 1, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 2, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 3, Y, 0)
    Call Bloquear(SendTarget.ToMap, 0, 164, 164, x + 4, Y, 0)

End Function

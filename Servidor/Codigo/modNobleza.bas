Attribute VB_Name = "modNobleza"
Option Explicit

Private dragonPeque(0 To 3) As Integer
Private criaSmaug(0 To 1) As Integer
Private dragonSmaug As Integer

Private dragonesPeques As Byte
Private criasSmaug As Byte

Public realizandoNobleza As Integer
Public tiempoNobleza As Byte
Private etapaNobleza As Byte
Public Sub nobleza_restarNPC(ByVal npcNumero As Integer)

    If etapaNobleza = 1 Then
        If npcNumero = 968 Then
            dragonesPeques = dragonesPeques - 1
            
            If dragonesPeques = 0 Then
                nobleza_etapaDos
            End If
        End If
    ElseIf etapaNobleza = 2 Then
        If npcNumero = 969 Then
            criasSmaug = criasSmaug - 1
            
            If criasSmaug = 0 Then
                nobleza_etapaTres
            End If
        End If
    Else
        
        If npcNumero = 970 Then
            party_entregarInframundo (realizandoNobleza)
            realizandoNobleza = 0
            etapaNobleza = 0
        End If
    End If
    
            

End Sub
Public Sub nobleza_etapaUno(ByVal index As Integer)

    Dim dragonPequepos As WorldPos, i As Long
            
    
    dragonPequepos.Map = 141
    dragonPequepos.Y = 42
            
    For i = 0 To 3
        dragonPequepos.X = 50 + i
        dragonPeque(i) = SpawnNpc(968, dragonPequepos, True, False)
    Next i
    
    dragonesPeques = 4
    realizandoNobleza = index
    etapaNobleza = 1
    tiempoNobleza = 5
    Call SendData(SendTarget.toMap, 0, 141, "||979")
End Sub
Public Sub nobleza_etapaDos()
    Dim criaSmaugpos As WorldPos, i As Long
            
    
    criaSmaugpos.Map = 141
    criaSmaugpos.Y = 42
            
    For i = 0 To 1
        criaSmaugpos.X = 50 + i
        criaSmaug(i) = SpawnNpc(969, criaSmaugpos, True, False)
    Next i
    
    criasSmaug = 2
    etapaNobleza = 2
    tiempoNobleza = 5
    Call SendData(SendTarget.toMap, 0, 141, "||980")
End Sub
Public Sub nobleza_etapaTres()
    Dim Smaugpos As WorldPos, i As Long
            
    
    Smaugpos.Map = 141
    Smaugpos.Y = 42
    Smaugpos.X = 50
    
    dragonSmaug = SpawnNpc(970, Smaugpos, True, False)
    
    etapaNobleza = 3
    tiempoNobleza = 10
    Call SendData(SendTarget.toMap, 0, 141, "||981")
End Sub
Private Sub nobleza_limpiarMapa()

    Dim i As Long

    If etapaNobleza = 1 Then
        For i = 0 To 3
            If dragonPeque(i) > 0 Then Call QuitarNPC(dragonPeque(i))
        Next i
        
    ElseIf etapaNobleza = 2 Then
        If criaSmaug(0) > 0 Then QuitarNPC (criaSmaug(0))
        If criaSmaug(1) > 0 Then QuitarNPC (criaSmaug(1))
    
    ElseIf etapaNobleza = 3 Then
        If dragonSmaug > 0 Then QuitarNPC (dragonSmaug)
    End If

End Sub
Public Sub nobleza_pasarTiempo()

    If tiempoNobleza > 0 Then
        tiempoNobleza = tiempoNobleza - 1
        
        If tiempoNobleza = 0 Then
            Call SendData(SendTarget.toMap, 0, 141, "||978")
            party_tepearTanaris (realizandoNobleza)
            realizandoNobleza = 0
            nobleza_limpiarMapa
            Exit Sub
        End If
        
        Call SendData(SendTarget.toMap, 0, 141, "||982@" & tiempoNobleza)
    End If

End Sub

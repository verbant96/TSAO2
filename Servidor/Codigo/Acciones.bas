Attribute VB_Name = "Acciones"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio, Jonatan Ezequiel Salguero
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public Lice/nse as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of theF GNU General Public License
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio

Option Explicit
Public Const SacriIndex As Integer = 936
Public Const DropSacri As Byte = 1

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

'�Posicion valida?
If InMapBounds(Map, X, Y) Then
   
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
    
                If UserList(userindex).flags.targetBot <> 0 Then
                      If UserList(userindex).flags.Hechizo > 0 Then
                        ia_UserDamage UserList(userindex).flags.Hechizo, UserList(userindex).flags.targetBot, userindex
                        UserList(userindex).flags.targetBot = 0
                        UserList(userindex).flags.Hechizo = 0
                        Exit Sub
                      End If
                End If
            
    'Puerti�as
    If MapData(Map, X - 1, Y).OBJInfo.ObjIndex > 0 Then
        Select Case ObjData(MapData(Map, X - 1, Y).OBJInfo.ObjIndex).OBJType
            Case eOBJType.otPuertas 'Es una puerta
                If ObjData(MapData(Map, X - 1, Y).OBJInfo.ObjIndex).PuertaDoble = 1 Or ObjData(MapData(Map, X - 1, Y).OBJInfo.ObjIndex).Porton = 1 Then
                    Call AccionParaPuerta(Map, X - 1, Y, userindex)
                End If
        End Select
    End If
    
    If MapData(Map, X - 2, Y).OBJInfo.ObjIndex > 0 Then
        Select Case ObjData(MapData(Map, X - 2, Y).OBJInfo.ObjIndex).OBJType
            Case eOBJType.otPuertas 'Es una puerta
                If ObjData(MapData(Map, X - 2, Y).OBJInfo.ObjIndex).Porton = 1 Then
                    Call AccionParaPuerta(Map, X - 2, Y, userindex)
                End If
        End Select
    End If
    
    If MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
        Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType
            Case eOBJType.otPuertas 'Es una puerta
                    Call AccionParaPuerta(Map, X + 1, Y, userindex)
        End Select
    End If
    
    If MapData(Map, X + 2, Y).OBJInfo.ObjIndex > 0 Then
        Select Case ObjData(MapData(Map, X + 2, Y).OBJInfo.ObjIndex).OBJType
            Case eOBJType.otPuertas 'Es una puerta
                If ObjData(MapData(Map, X - 1, Y).OBJInfo.ObjIndex).PuertaDoble Or ObjData(MapData(Map, X + 2, Y).OBJInfo.ObjIndex).Porton = 1 Then
                    Call AccionParaPuerta(Map, X + 2, Y, userindex)
                End If
        End Select
    End If
    'Puerti�as
    
       
    '�Es un obj?
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        UserList(userindex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        
        If UCase$(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Dios) = "MIFRIT" Or UCase$(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Dios) = "POSEIDON" Or UCase$(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Dios) = "EREBROS" Or UCase$(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Dios) = "TARRASKE" Then
          If UCase$(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Dios) = UCase$(UserList(userindex).flags.SirvienteDeDios) Then
            If UserList(userindex).flags.JerarquiaDios <= 4 Then Call SendData(SendTarget.toindex, userindex, 0, "GODS" & UserList(userindex).flags.AlmasOfrecidas & "," & (AlmasNecesarias * UserList(userindex).flags.JerarquiaDios) & "," & UserList(userindex).flags.SirvienteDeDios)
            If UserList(userindex).flags.JerarquiaDios = 5 Then Call SendData(SendTarget.toindex, userindex, 0, "GODS" & UserList(userindex).flags.AlmasOfrecidas & "," & (AlmasNecesarias * 4) & "," & UserList(userindex).flags.SirvienteDeDios)
          Else
            Call SendData(SendTarget.toindex, userindex, 0, "||635")
            Exit Sub
          End If
        End If
        
    Select Case ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType
            Case eOBJType.otPuertas 'Es una puerta
                Call AccionParaPuerta(Map, X, Y, userindex)
            Case eOBJType.otCarteles 'Es un cartel
                Call AccionParaCartel(Map, X, Y, userindex)
            Case eOBJType.otLe�a    'Le�a
                If MapData(Map, X, Y).OBJInfo.ObjIndex = FOGATA_APAG And UserList(userindex).flags.Muerto = 0 Then
                    Call AccionParaRamita(Map, X, Y, userindex)
                End If
            Case eOBJType.otCofreJDH
                Call Clickea_Cofre(Map, X, Y, userindex)
    End Select
    
    ElseIf MapData(Map, X, Y).npcindex > 0 Then   'Acciones NPCs
        'Set the target NPC
          UserList(userindex).flags.TargetNPC = MapData(Map, X, Y).npcindex
        
        If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 1 Then
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 6 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||13")
                Exit Sub
            End If
            
            'Iniciamos la rutina pa' comerciar.
            Call IniciarCOmercioNPC(userindex)
        ElseIf Npclist(MapData(Map, X, Y).npcindex).Numero = 156 Or Npclist(MapData(Map, X, Y).npcindex).Numero = 157 Or Npclist(MapData(Map, X, Y).npcindex).Numero = 158 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 181 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 182 Then
            If Npclist(UserList(userindex).flags.TargetNPC).Due�oMascota = userindex Then
                Call SendData(SendTarget.toindex, userindex, 0, "AXELPT")
            End If
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.Viajero Then
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 5 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||13")
                Exit Sub
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "TRAVELS")
            End If
            
        ElseIf Npclist(UserList(userindex).flags.TargetNPC).NPCtype = eNPCType.Quest Then
        
                    If UserList(userindex).flags.Muerto = 1 Then
                           Call SendData(SendTarget.toindex, userindex, 0, "||3")
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 10 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||10")
                        Exit Sub
                    End If
                    
            Call SendData(SendTarget.toindex, userindex, 0, "DAMEQUEST")
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.EntregaCajas Then
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            
            If UserList(userindex).flags.SirvienteDeDios = "" Then
                Call SendData(SendTarget.toindex, userindex, 0, "||636")
                Exit Sub
            End If
            
            If UserList(userindex).flags.AlmasContenidas < 10000 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||637")
                Exit Sub
            End If
            
            If TieneObjetos(1473, 1, userindex) = True Or TieneObjetos(1475, 1, userindex) = True Or TieneObjetos(1477, 1, userindex) = True Or TieneObjetos(1479, 1, userindex) = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||638")
                Exit Sub
            End If
            
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 5 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||13")
                Exit Sub
            Else
            
                Dim CofreNecesita As obj
                
                    CofreNecesita.Amount = 1
                
                     If UserList(userindex).flags.SirvienteDeDios = "Tarraske" Then
                       CofreNecesita.ObjIndex = 1479
                    ElseIf UserList(userindex).flags.SirvienteDeDios = "Mifrit" Then
                       CofreNecesita.ObjIndex = 1475
                    ElseIf UserList(userindex).flags.SirvienteDeDios = "Erebros" Then
                       CofreNecesita.ObjIndex = 1473
                    ElseIf UserList(userindex).flags.SirvienteDeDios = "Poseidon" Then
                       CofreNecesita.ObjIndex = 1477
                    End If
                    
                If Not MeterItemEnInventario(userindex, CofreNecesita) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||639")
                    Exit Sub
                End If
                
                UserList(userindex).flags.AlmasContenidas = UserList(userindex).flags.AlmasContenidas - 10000
                
                If UserList(userindex).flags.AlmasOfrecidas >= 120000 Then
                    Call QuitarObjetos(1274, 1, userindex)
                End If
            End If
            
            
            
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.Correos Then
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
        
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 5 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||13")
                Exit Sub
            Else
                correoIniciarForm userindex
            End If

        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.BoveClan Then
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 5 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||10")
                Exit Sub
            End If
            
            If UserList(userindex).GuildIndex <= 0 Then Exit Sub
            
                Dim i As Long
                For i = 1 To LastUser
                 If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) = UCase$(UserList(i).flags.CuentaBancaria) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||640")
                   Exit Sub
                 End If
                Next i
                
                UserList(userindex).BancoInventB.NroItems = CInt(GetVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "BancoInventory", "CantidadItems"))
                
                Dim ln As String
                'Lista de objetos del banco
                Dim loopC As Long
                For loopC = 1 To MAX_BANCOINVENTORY_SLOTS
                    ln = (GetVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "BancoInventory", "Obj" & loopC))
                    UserList(userindex).BancoInventB.Object(loopC).ObjIndex = CInt(ReadField(1, ln, 45))
                    UserList(userindex).BancoInventB.Object(loopC).Amount = CInt(ReadField(2, ln, 45))
                Next loopC
                
                UserList(userindex).flags.CuentaBancaria = Guilds(UserList(userindex).GuildIndex).GuildName
                Call BIniciarDeposito(userindex)
                
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.QuintaJera Then
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 5 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||10")
                Exit Sub
            Else
            
                If UserList(userindex).Pos.Map = 29 And (EsAlianza(userindex) = False Or (UserList(userindex).Faccion.RecompensasReal < 4)) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||641")
                 Exit Sub
                Else
                  Call RecompensaArmadaReal(userindex, True)
                End If
                
                If UserList(userindex).Pos.Map = 27 And (Criminal(userindex) = False Or (UserList(userindex).Faccion.RecompensasCaos < 4)) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||642")
                 Exit Sub
                Else
                    Call RecompensaCaos(userindex, True)
                End If
                
                
                
            End If
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.ShowCasas Then
            Call SendData(SendTarget.toindex, userindex, 0, "MFC")
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.Banquero Then
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 5 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||13")
                Exit Sub
            End If
            
            'A depositar de una
            SendUserBANK (userindex)
            SendData SendTarget.toindex, userindex, 0, "INITBANKO"
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.NpcDioses Then
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 3 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||14")
                Exit Sub
            End If
           
            If UserList(userindex).Stats.ELV < 60 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||643")
              Exit Sub
            End If
            
            If TieneObjetos(1274, 1, userindex) Or UserList(userindex).flags.JerarquiaDios > 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||644")
              Exit Sub
            End If
            
            Dim ElContenedor As obj
            ElContenedor.ObjIndex = 1274
            ElContenedor.Amount = 1
            
            If Not MeterItemEnInventario(userindex, ElContenedor) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||645")
              Exit Sub
            Else
                
                Dim RandomDios As Byte
                RandomDios = RandomNumber(1, 4)
                
                If RandomDios = 1 Then
                    UserList(userindex).flags.SirvienteDeDios = "Mifrit"
                ElseIf RandomDios = 2 Then
                    UserList(userindex).flags.SirvienteDeDios = "Poseidon"
                ElseIf RandomDios = 3 Then
                    UserList(userindex).flags.SirvienteDeDios = "Erebros"
                ElseIf RandomDios = 4 Then
                    UserList(userindex).flags.SirvienteDeDios = "Tarraske"
                End If
                    
                UserList(userindex).flags.JerarquiaDios = 1
                Call SendData(SendTarget.toindex, userindex, 0, "||646")
            End If
            
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.QuestNoble Then
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 3 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||14")
                Exit Sub
            End If
           
            If UserList(userindex).Stats.ELV < 60 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||643")
            Exit Sub
            End If
           
            If MapInfo(141).NumUsers > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||647")
            Exit Sub
            End If
           
       
            If Not TieneObjetos(1073, 1, userindex) Or Not TieneObjetos(1074, 1, userindex) Or Not TieneObjetos(1075, 1, userindex) Or Not TieneObjetos(1076, 1, userindex) Or Not TieneObjetos(1047, 1, userindex) Or Not TieneObjetos(895, 1, userindex) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||648")
                Exit Sub
            Else
              If UserList(userindex).flags.partyIndex = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||649")
                Exit Sub
              End If
          
            If miembrosParty(UserList(userindex).flags.partyIndex) < 5 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||649")
                Exit Sub
            End If
            
            If (realizandoNobleza > 0) Then Call SendData(SendTarget.toindex, userindex, 0, "984"): Exit Sub
            
            mdParty.party_tepearNobleza (userindex)
        
            Call QuitarObjetos(1047, 1, userindex) 'Saca la gema negra
            Call QuitarObjetos(895, 1, userindex)
            
            nobleza_etapaUno (UserList(userindex).flags.partyIndex)
        End If
        
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.NpcBargomaud Then
        
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 5 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||14")
                Exit Sub
            End If
            
            If UserList(userindex).Stats.ELV < 55 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||643")
                Exit Sub
            End If
            
            
            Call WarpUserChar(userindex, 161, 50, 53, True)
            Call SendData(SendTarget.toindex, userindex, 0, "||651@" & UserList(userindex).Name)

        
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.cirujano Then
            If Distancia(Npclist(MapData(Map, X, Y).npcindex).Pos, UserList(userindex).Pos) > 5 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||14")
                Exit Sub
            End If
                
                SendData SendTarget.toindex, userindex, 0, "CIRUJA" & UserList(userindex).Raza & "," & UserList(userindex).Genero
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.Arenas Then
            Call SendData(SendTarget.toindex, userindex, 0, "MAR" & NombreDueleando(1) & "," & NombreDueleando(2) & "," & NombreDueleando(3) & "," & NombreDueleando(4) & "," & NombreDueleando(5) & "," & NombreDueleando(6) & "," & NombreDueleando(7) & "," & NombreDueleando(8))
        ElseIf Npclist(MapData(Map, X, Y).npcindex).NPCtype = eNPCType.Revividor Then
            If Distancia(UserList(userindex).Pos, Npclist(MapData(Map, X, Y).npcindex).Pos) > 10 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||12")
                Exit Sub
            End If
           
           'Revivimos si es necesario
            If UserList(userindex).flags.Muerto = 1 Then
                Call RevivirUsuario(userindex)
            End If
            
            'curamos totalmente
            UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
            Call SendUserHP(userindex)
            
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
        UserList(userindex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.ObjIndex
        Call SendData(SendTarget.toindex, userindex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y, userindex)
            
        End Select
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(userindex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex
        Call SendData(SendTarget.toindex, userindex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y + 1, userindex)
            
        End Select
    ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(userindex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.ObjIndex
        Call SendData(SendTarget.toindex, userindex, 0, "SELE" & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X, Y + 1, userindex)
            
        End Select
            
        End If
    Else
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
    End If
End If

If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).userindex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).userindex
            If UserList(TempCharIndex).showName Then
                FoundChar = 1
            End If
        End If
        If MapData(Map, X, Y + 1).npcindex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).npcindex
            FoundChar = 2
        End If
    End If
 
    If FoundChar = 0 Then
        If MapData(Map, X, Y).userindex > 0 Then
            TempCharIndex = MapData(Map, X, Y).userindex
            If UserList(TempCharIndex).showName Then
                FoundChar = 1
            End If
        End If
        If MapData(Map, X, Y).npcindex > 0 Then
            TempCharIndex = MapData(Map, X, Y).npcindex
            FoundChar = 2
        End If
        
    If MapData(Map, X, Y).userindex > 0 And MapData(Map, X, Y).userindex <> userindex And UserList(MapData(Map, X, Y).userindex).flags.AdminInvisible = 0 Then
        TempCharIndex = MapData(Map, X, Y).userindex
        UserList(userindex).flags.TargetUser = TempCharIndex
        Call SendData(SendTarget.toindex, userindex, 0, "MENU" & UserList(TempCharIndex).Name & "," & UserList(userindex).flags.Privilegios)
    End If
        
    End If
End Sub
Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next

Dim MiObj As obj
Dim wp As WorldPos

If Not (Distance(UserList(userindex).Pos.X, UserList(userindex).Pos.Y, X, Y) > 3) Then
    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
        
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).RejaForta = 1 Then
                
                    If MapData(Map, X, Y).OBJInfo.ObjIndex = 1472 Then Exit Sub
                    If UserList(userindex).GuildIndex <= 0 Then Exit Sub
                    If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(Fortaleza) Then Exit Sub
                        
                        'Desbloquea
                        MapData(Map, X, Y).Blocked = 0
                        MapData(Map, X - 1, Y).Blocked = 0
                        MapData(Map, X - 2, Y).Blocked = 0
                        MapData(Map, X + 1, Y).Blocked = 0
                        MapData(Map, X + 2, Y).Blocked = 0
                        Call Bloquear(SendTarget.toMap, 0, Map, Map, X, Y, 0)
                        Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 1, Y, 0)
                        Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 2, Y, 0)
                        Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 1, Y, 0)
                        Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 2, Y, 0)
                    MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexAbierta
                    Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y)
                 Exit Sub
                End If
        
                'Abre la puerta
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexAbierta
                    Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y)
                     
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).PuertaDoble = 1 Then
                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    MapData(Map, X + 1, Y).Blocked = 0
                    MapData(Map, X + 2, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X, Y, 0)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 1, Y, 0)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 1, Y, 0)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 2, Y, 0)
                ElseIf ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Porton = 1 Then
                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    MapData(Map, X - 2, Y).Blocked = 0
                    MapData(Map, X + 1, Y).Blocked = 0
                    MapData(Map, X + 2, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X, Y, 0)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 1, Y, 0)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 2, Y, 0)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 1, Y, 0)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 2, Y, 0)
                Else
                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X, Y, 0)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 1, Y, 0)
                End If
                      
                    'Sonido
                    SendData SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_PUERTA
                    
                Else
                     Call SendData(SendTarget.toindex, userindex, 0, "||652")
                End If
        Else
        
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).RejaForta = 1 Then
                
                    If MapData(Map, X, Y).OBJInfo.ObjIndex = 1472 Then Exit Sub
                    If UserList(userindex).GuildIndex <= 0 Then Exit Sub
                    If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(Fortaleza) Then Exit Sub
                
                    'Bloquea
                    MapData(Map, X, Y).Blocked = 1
                    MapData(Map, X - 1, Y).Blocked = 1
                    MapData(Map, X - 2, Y).Blocked = 1
                    MapData(Map, X + 1, Y).Blocked = 1
                    MapData(Map, X + 2, Y).Blocked = 1
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 1, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 2, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 1, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 2, Y, 1)
                    
                    'Cierra puerta
                    MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexCerrada
                    Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y)
                 Exit Sub
               End If
        
                'Cierra puerta
                MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexCerrada
                
                Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y)
                
                
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).PuertaDoble = 1 Then
                    MapData(Map, X, Y).Blocked = 1
                    MapData(Map, X - 1, Y).Blocked = 1
                    MapData(Map, X + 1, Y).Blocked = 1
                    MapData(Map, X + 2, Y).Blocked = 1
                
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 1, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 1, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 2, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X, Y, 1)
                ElseIf ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Porton = 1 Then
                    MapData(Map, X, Y).Blocked = 1
                    MapData(Map, X - 1, Y).Blocked = 1
                    MapData(Map, X - 2, Y).Blocked = 1
                    MapData(Map, X + 1, Y).Blocked = 1
                    MapData(Map, X + 2, Y).Blocked = 1
                
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 1, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 2, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 1, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X + 2, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X, Y, 1)
                Else
                    MapData(Map, X, Y).Blocked = 1
                    MapData(Map, X - 1, Y).Blocked = 1
                
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X - 1, Y, 1)
                    Call Bloquear(SendTarget.toMap, 0, Map, Map, X, Y, 1)
                End If
                
                SendData SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_PUERTA
        End If
        
        UserList(userindex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||652")
    End If
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||10")
End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next


Dim MiObj As obj

If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto) > 0 Then
       Call SendData(SendTarget.toindex, userindex, 0, "MCAR" & _
        ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto & _
        Chr(176) & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhSecundario)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim obj As obj
Dim raise As Integer

Dim Pos As WorldPos
Pos.Map = Map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(userindex).Pos) > 2 Then
    Call SendData(toindex, userindex, 0, "||10")
    Exit Sub
End If

If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
    Call SendData(SendTarget.toindex, userindex, 0, "||345")
    Exit Sub
End If

If UserList(userindex).Stats.UserSkills(Supervivencia) > 1 And UserList(userindex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 6 And UserList(userindex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 10 And UserList(userindex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    If MapInfo(UserList(userindex).Pos.Map).Zona <> Ciudad Then
        obj.ObjIndex = FOGATA
        obj.Amount = 1
        
        Call SendData(toindex, userindex, 0, "||653")
        Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "FO")
        
        Call MakeObj(toMap, 0, Map, obj, Map, X, Y)
        
        'Las fogatas prendidas se deben eliminar
        Dim Fogatita As New cGarbage
        Fogatita.Map = Map
        Fogatita.X = X
        Fogatita.Y = Y
        Call TrashCollector.Add(Fogatita)
    Else
        Call SendData(toindex, userindex, 0, "||345")
        Exit Sub
    End If
Else
    Call SendData(toindex, userindex, 0, "||654")
End If

'Sino tiene hambre o sed quizas suba el skill supervivencia
If UserList(userindex).flags.Hambre = 0 And UserList(userindex).flags.Sed = 0 Then
    Call SubirSkill(userindex, Supervivencia)
End If

End Sub
Function AgregarPuntos(userindex As Integer, ByVal Cantidad As Long)
    UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + Cantidad
    Call EnviarPuntos(userindex)
End Function
Sub OtorgarGranPoder(userindex As Integer)
Dim EncontroIdeal As Boolean
On Error Resume Next

If LastUser = 0 Then Exit Sub
If userindex = 0 Then
    For userindex = 1 To LastUser
         If UserList(userindex).flags.UserLogged = True Then
             If UserList(userindex).flags.Muerto = 0 And UserList(userindex).flags.Privilegios = User And MapInfo(UserList(userindex).Pos.Map).Pk = True And UserList(userindex).Pos.Map <> 78 And UserList(userindex).Pos.Map <> 31 And UserList(userindex).Pos.Map <> 32 And UserList(userindex).Pos.Map <> 33 And UserList(userindex).Pos.Map <> 34 And UserList(userindex).Pos.Map <> 167 And (Not MapaEspecial(userindex)) And _
              UserList(userindex).flags.TeniaElDon = 0 And (NumUsers + BOnlines) >= 10 Then
                EncontroIdeal = True
                Exit For
            End If
        End If
     Next userindex
     
    If Not EncontroIdeal Then
        userindex = 0
        GranPoder = 0
        
        Dim paratodosxd As Integer
        For paratodosxd = 1 To LastUser
        If UserList(paratodosxd).flags.TeniaElDon = 1 Then
            UserList(paratodosxd).flags.TeniaElDon = 0
        End If
        Next paratodosxd
                
    End If
End If

If userindex > 0 Then
    GranPoder = userindex
    Call SendData(SendTarget.ToAll, userindex, 0, "||655@" & UserList(userindex).Name & "@" & UserList(userindex).Pos.Map)
    UserList(userindex).flags.GranPoder = 1
    SendUserVariant (userindex)
    Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, False)
Else
    GranPoder = 0
End If

End Sub

Attribute VB_Name = "Mod_TCP"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'You should have received a copy of the GNU General Public License
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean

Public Function PuedoQuitarFoco() As Boolean
    PuedoQuitarFoco = True
End Function

Sub HandleData(ByVal rData As String)
    On Error Resume Next
    
    Dim RetVal As Variant
    Dim X As Integer
    Dim Y As Integer
    Dim charindex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim i As Integer, k As Integer
    Dim cad$, Index As Integer, m As Integer
    Dim T() As String
    
    Dim tStr As String
    Dim tstr2 As String
    
    
    Dim sData As String
    rData = AoDefServDecrypt(AoDefDecode(rData))
    sData = UCase$(rData)
    
    If left$(sData, 4) = "INVI" Then CartelInvisibilidad = Right$(sData, Len(sData) - 4)
    If left$(sData, 4) = "ARAM" Then AramSeconds = Right$(sData, Len(sData) - 4)
    
    Debug.Print "Recibido: " & sData
    
    Select Case sData
        Case "MUERT"
            frmMuertito.Show , frmMain
        Exit Sub
    
        Case "LOGGED"            ' >>>>> LOGIN :: LOGGED
            AlphaY = 130
            ISItem = True
            mode = True
            logged = True
            UserCiego = False
            UserDescansar = False
            Nombres = True
            IsSeguroC = True
            
           If frmCrearPersonaje.Visible Then
                Unload frmCrearPersonaje
                Unload frmConnect
                Unload frmAccount
                frmMain.Show
            End If
            
            Call SetConnected
            
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            Exit Sub
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.RemoveAllDialogs
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
        Exit Sub
        Case "FINOK" ' Graceful exit ;))
            frmMain.Socket1.Disconnect
            
            frmMain.Visible = False
            logged = False
            AoDefResult = 0
            UserParalizado = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            
            charlist(UserCharIndex).color = 0
            
            Call frmConnect.mostrarConectar(True)
            Call Audio.StopWave
            bFogata = False
            SkillPoints = 0
            Call Dialogos.RemoveAllDialogs
            
            bK = 0
            If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
                frmConnect.MousePointer = 1
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        Case "FINCBNOK"          ' >>>>> Finaliza Cuenta Bancaria :: FINCBNOK
            frmNuevoBancoObj.List1(0).Clear
            frmNuevoBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmNuevoBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            Dim tmpIndex As Byte
            tmpIndex = 1
            frmComerciar.List1(1).Clear
            For i = 1 To MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                        slotsListaInv(tmpIndex) = i
                        tmpIndex = tmpIndex + 1
                End If
            Next i
            
            tmpIndex = 1
            frmComerciar.List1(0).Clear
            For i = 1 To MAX_NPC_INVENTORY_SLOTS
                If NPCInventory(i).GrhIndex > 0 Then
                     frmComerciar.List1(0).AddItem NPCInventory(i).Name
                    slotsListaNPC(tmpIndex) = i
                    tmpIndex = tmpIndex + 1
                End If
            Next i
            
            Comerciando = True
            frmComerciar.Show , frmMain
            Exit Sub
        '[KEVIN]-----------------------------------------------
        '**************************************************************
        Case "INITBANKO"
            frmBanco.Show , frmMain
        Exit Sub
        Case "INITSUB"           ' >>>>> Inicia Subasta :: #Fer
            i = 1
            frmSubastar.ItemList.Clear
           
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmSubastar.ItemList.AddItem Inventario.ItemName(i)
                Else
                        frmSubastar.ItemList.AddItem "Nada"
                End If
                i = i + 1
            Loop
            frmSubastar.Show , frmMain
        Exit Sub
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            tmpIndex = 1
            For i = 1 To MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
                        slotsListaInv(tmpIndex) = i
                        tmpIndex = tmpIndex + 1
                End If
            Next i
            
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmBancoObj.Show , frmMain
        Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
       Case "BORROK"
            
        Mensaje.Escribir "El personaje ha sido borrado."
        frmMain.Socket1.Disconnec
        frmConnect.mostrarConectar (True)

        Exit Sub
        Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
        Case "HOLASOYUNCIRUJA"
            TimerPing(2) = GetTickCount()
            Dim cuantolagtengoxd As String
            
            If TimerPing(2) - TimerPing(1) > 0 And TimerPing(2) - TimerPing(1) < 100 Then
            cuantolagtengoxd = "0 Lag"
            ElseIf TimerPing(2) - TimerPing(1) > 100 And TimerPing(2) - TimerPing(1) < 200 Then
            cuantolagtengoxd = "Bajo"
            ElseIf TimerPing(2) - TimerPing(1) > 200 And TimerPing(2) - TimerPing(1) < 400 Then
            cuantolagtengoxd = "Medio"
            ElseIf TimerPing(2) - TimerPing(1) > 400 And TimerPing(2) - TimerPing(1) < 900 Then
            cuantolagtengoxd = "Alto"
            ElseIf TimerPing(2) - TimerPing(1) > 900 Then
            cuantolagtengoxd = "Injugable"
            End If
            
            Call AddtoRichTextBox(frmMain.RecTxt, "<<Recibido: el ping es de " & TimerPing(2) - TimerPing(1) & " Mili-Segundos (" & (TimerPing(2) - TimerPing(1)) / 1000 & " Seg) LAG: " & cuantolagtengoxd, 0, 255, 0, True, False, False)
        Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
            
        Case "SEGONR" ' <--- Activa el seguro de resi
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, False)
        Exit Sub

        Case "SEGOFR" ' <--- Desactiva el seguro de resu
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, False)
        Exit Sub
        
        Case "SEGON" '  <--- Activa el seguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)
            IsSeguroC = False
            Exit Sub
            
        Case "SEGOFF" ' <--- Desactiva el seguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, False)
            IsSeguroC = False
            Exit Sub
    End Select

Select Case left(sData, 1)
        Case "+"              ' >>>>> Mover Char >>> +
            rData = Right$(rData, Len(rData) - 1)
            charindex = Val(ReadField(1, rData, Asc(",")))
            X = Val(ReadField(2, rData, Asc(",")))
            Y = Val(ReadField(3, rData, Asc(",")))

            With charlist(charindex)
                    For i = 1 To 3
                        If .FxIndex(i) = 4 Or .FxIndex(i) = 5 Or .FxIndex(i) = 6 Or .FxIndex(i) = 42 Or .FxIndex(i) = 43 Or .FxIndex(i) = 44 Or .FxIndex(i) = 45 Or .FxIndex(i) = 16 Or .FxIndex(i) = 103 Or .FxIndex(i) = 104 Or .FxIndex(i) = 105 Then    'If it's meditating, we remove the FX
                            .FxIndex(i) = 0
                            .Fx(i).Loops = 0
                        End If
                    Next i
                    
                ' Play steps sounds if the user is not an admin of any kind
                If .priv = 0 And .montVol = 0 Then
                    Call DoPasosFx(charindex)
                End If
            End With

            Call engine.Char_Move_by_Pos(charindex, X, Y)
            Call RefreshAllChars
            Exit Sub
        Case "*", "_"             ' >>>>> Mover NPC >>> *
            rData = Right$(rData, Len(rData) - 1)
            
            charindex = Val(ReadField(1, rData, Asc(",")))
            X = Val(ReadField(2, rData, Asc(",")))
            Y = Val(ReadField(3, rData, Asc(",")))
            
            With charlist(charindex)
                    For i = 1 To 3
                        If .FxIndex(i) = 4 Or .FxIndex(i) = 5 Or .FxIndex(i) = 6 Or .FxIndex(i) = 42 Or .FxIndex(i) = 43 Or .FxIndex(i) = 44 Or .FxIndex(i) = 45 Or .FxIndex(i) = 16 Or .FxIndex(i) = 103 Or .FxIndex(i) = 104 Or .FxIndex(i) = 105 Then    'If it's meditating, we remove the FX
                            .FxIndex(i) = 0
                            .Fx(i).Loops = 0
                        End If
                    Next i
                    
                ' Play steps sounds if the user is not an admin of any kind
                If .priv = 0 And .montVol = 0 Then
                    Call DoPasosFx(charindex)
                End If
            End With
    
            Call engine.Char_Move_by_Pos(charindex, X, Y)
            Call RefreshAllChars
        Exit Sub
    End Select

    Select Case left$(sData, 2)
    
        Case "99"
            rData = Right$(rData, Len(rData) - 2)
            frmBonificadores.lblBeneficio(0) = ReadField(1, rData, 44)
            frmBonificadores.lblBeneficio(1) = ReadField(2, rData, 44)
            frmBonificadores.Show , frmMain
        Exit Sub
        
        
        Exit Sub
        Case "CU"
        rData = Right$(rData, Len(rData) - 2)
            Dim CunT As Byte
            CunT = Val(rData)
            
            If CunT = 0 Then
                Cuenta = False
                Tiempo = 45
                Conteo = 27850
                ConteoH = GrhData(Conteo).pixelHeight
                ConteoW = GrhData(Conteo).pixelWidth
                TransparenciaCont = 220
            ElseIf CunT < 11 Then
                Conteo = 27850 + CunT
                ConteoH = GrhData(Conteo).pixelHeight
                ConteoW = GrhData(Conteo).pixelWidth
                TransparenciaCont = 220
                Cuenta = True
                If CunT = 0 Then Cuenta = False And Tiempo = 45
            Else
                Cuenta = False
            End If
        Exit Sub
        
        Case "CM"              ' >>>>> Cargar Mapa :: CM
            rData = Right$(rData, Len(rData) - 2)
            UserMap = ReadField(1, rData, 44)
            
            If FileExist(App.Path & "\Data\MAPAS\" & "Mapa" & UserMap & ".map", vbNormal) Then
                Open App.Path & "\Data\MAPAS\" & "Mapa" & UserMap & ".map" For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
                
                day_r_old = Val(ReadField(2, rData, 44))
                day_g_old = Val(ReadField(3, rData, 44))
                day_b_old = Val(ReadField(4, rData, 44))
                base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
                
                If day_r_old > 0 Or day_g_old > 0 Or day_b_old > 0 Then
                    TieneColorMapa = True
                Else
                    TieneColorMapa = False
                End If
                
'                If tempint = Val(ReadField(2, Rdata, 44)) Then
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call UnloadAllForms
                End
            End If
            Exit Sub
        
        Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            rData = Right$(rData, Len(rData) - 2)
                MapData(UserPos.X, UserPos.Y).charindex = 0
                UserPos.X = CInt(ReadField(1, rData, 44))
                UserPos.Y = CInt(ReadField(2, rData, 44))
                MapData(UserPos.X, UserPos.Y).charindex = UserCharIndex
                charlist(UserCharIndex).Pos = UserPos
                actualizarAL (UserCharIndex)
                RefreshAllChars
        Exit Sub
        Case "PT"                 ' >>>>> Actualiza Posición Usuario :: PU
            rData = Right$(rData, Len(rData) - 2)
            If UserPuedeRefrescar Then
                MapData(UserPos.X, UserPos.Y).charindex = 0
                UserPos.X = CInt(ReadField(1, rData, 44))
                UserPos.Y = CInt(ReadField(2, rData, 44))
                MapData(UserPos.X, UserPos.Y).charindex = UserCharIndex
                charlist(UserCharIndex).Pos = UserPos
                Call DibujarPuntoMinimap
                frmMain.Coord.Caption = Nombredelmapaxx & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                UserPuedeRefrescar = False
                RefreshAllChars
            End If
        Exit Sub
        
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            rData = Right$(rData, Len(rData) - 2)
            i = Val(ReadField(1, rData, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            rData = Right$(rData, Len(rData) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & rData & MENSAJE_2, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            rData = Right$(rData, Len(rData) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & rData & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            rData = Right$(rData, Len(rData) - 2)
            i = Val(ReadField(1, rData, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            rData = Right$(rData, Len(rData) - 2)
            i = Val(ReadField(1, rData, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
       Case "RT"
            rData = Right$(rData, Len(rData) - 2)
           
            If rData <> vbNullString Then
               Call RenderGM.Create(rData)
            End If
        Case "||"                 ' >>>>> Nuevo dialogo :: ||
            rData = Right$(rData, Len(rData) - 2)
            Dim IDText As Long
            Dim DatoAdd(1 To 8) As String 'Datos adicionales
            
            IDText = Val(ReadField(1, rData, Asc("@")))
            tStr = Messages(IDText).text
            
            'Reemplazo los datos adicionales
            For i = 1 To 8
                DatoAdd(i) = ReadField(1 + i, rData, Asc("@"))
                If DatoAdd(i) = vbNullString Then Exit For
            
                tStr = Replace(tStr, "%" & i, DatoAdd(i))
            Next i
                
            'Tiramos el texto a la consola
            AddtoRichTextBox frmMain.RecTxt, tStr, FontTypes(Messages(IDText).font).r, FontTypes(Messages(IDText).font).g, FontTypes(Messages(IDText).font).b, FontTypes(Messages(IDText).font).bold, FontTypes(Messages(IDText).font).italic
        Exit Sub
        Case "N|"                 ' >>>>> Dialogo de Usuarios y NPCs ::    N|
            rData = Right$(rData, Len(rData) - 2)
            Dim iuser As Integer, txt As String
            iuser = Val(ReadField(3, rData, 176))
            txt = ReadField(2, rData, 176)
            
            If iuser > 0 Then
                Dialogos.CreateDialog txt, iuser, Val(ReadField(1, rData, 176))
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
                End If
            End If

        Exit Sub
        Case "T|"                 ' >>>>>T HAY QUE VOLAR ESTO. ESTA SOLO POR APURO
            rData = Right$(rData, Len(rData) - 2)
            iuser = Val(ReadField(3, rData, 176))
            txt = ReadField(2, rData, 176)
            
            If iuser > 0 Then
                Dialogos.CreateDialog txt, iuser, Val(ReadField(1, rData, 176))
                If (Configuracion.Mensajes) And (Len(txt) > 0) Then AddtoRichTextBox frmMain.RecTxt, charlist(iuser).Nombre & "> " & txt, 255, 255, 255
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
                End If
            End If

        Exit Sub
        Case "P|"
            rData = Right$(rData, Len(rData) - 2)
            If Configuracion.Desactivar_Privados = 0 Then
            AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
            AddtoRichTextBox frmMain.PrivatesConsole, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
            End If
        Exit Sub
        Case "C|"
            rData = Right$(rData, Len(rData) - 2)
            AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
            AddtoRichTextBox frmMain.ClanConsole, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
        Exit Sub
        Case "G|"
            rData = Right$(rData, Len(rData) - 2)
            If Configuracion.Desactivar_Globales = 0 Then
            AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
            AddtoRichTextBox frmMain.GlobalConsole, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
            End If
        Exit Sub
        Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
            rData = Right$(rData, Len(rData) - 2)
            
            iuser = Val(ReadField(3, rData, 176))

            If iuser = 0 Then
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
                End If
            End If

            Exit Sub
        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                rData = Right$(rData, Len(rData) - 2)
                Mensaje.Escribir rData
            End If
            Exit Sub
            
        Case "ON"
            rData = Right$(rData, Len(rData) - 2)
            
            frmMain.ONLINES.Caption = rData
        Exit Sub
        
        Case "LK" ' >>>>> newbie
            rData = Right$(rData, Len(rData) - 2)
            charindex = ReadField(1, rData, 44)
            charlist(charindex).esNW = Val(ReadField(2, rData, 44))
        Case "XC"              ' >>>>> Nombres :: XC - Actualizamos todo a un solito paquete
            rData = Right$(rData, Len(rData) - 2)
            charindex = Val(ReadField(1, rData, 44))
            charlist(charindex).color = Val(ReadField(2, rData, 44))
        Exit Sub
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            rData = Right$(rData, Len(rData) - 2)
            UserCharIndex = Val(rData)
        Exit Sub
        Case "CC" ' >>>>> Crear un Personaje :: CC
            rData = Right$(rData, Len(rData) - 2)
            charindex = ReadField(4, rData, 44)
            X = ReadField(5, rData, 44)
            Y = ReadField(6, rData, 44)
             

            charlist(charindex).Nombre = ReadField(10, rData, 44)
            charlist(charindex).EsStatus = Val(ReadField(11, rData, 44))
            charlist(charindex).priv = Val(ReadField(12, rData, 44))
            charlist(charindex).NPCAura = Val(ReadField(13, rData, 44))
            charlist(charindex).NPCNumber = Val(ReadField(14, rData, 44))
            
            Call InitGrh(charlist(charindex).NPCAuraG, AurasPJ(charlist(charindex).NPCAura).GrhIndex)
            charlist(charindex).NPCAuraAngle = 0
            
            If Val(ReadField(2, rData, 44)) = 500 Or Val(ReadField(2, rData, 44)) = 501 Or Val(ReadField(2, rData, 44)) = 511 Or Val(ReadField(2, rData, 44)) = 511 Or Val(ReadField(2, rData, 44)) = 512 Then
                charlist(charindex).Muerto = True
            End If
            
            If Len(charlist(charindex).Nombre) > 0 Then
                realizarSuma (charindex)
            End If
            
            'Guardamos
            Call MakeChar(charindex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(3, rData, 44), X, Y, Val(ReadField(7, rData, 44)), Val(ReadField(8, rData, 44)), Val(ReadField(9, rData, 44)))
            Call RefreshAllChars
        Exit Sub
        Case "PX"
        rData = Right$(rData, Len(rData) - 2)
            charindex = ReadField(1, rData, 44)
            charlist(charindex).EsStatus = Val(ReadField(2, rData, 44))
            charlist(charindex).Nombre = ReadField(3, rData, 44)
        Exit Sub
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            rData = Right$(rData, Len(rData) - 2)
            Call EraseChar(Val(rData))
            Call Dialogos.RemoveDialog(Val(rData))
            Call RefreshAllChars
            Exit Sub
        Case "MP"             ' >>>>> Mover un Personaje :: MP
            rData = Right$(rData, Len(rData) - 2)
            charindex = Val(ReadField(1, rData, 44))
            
            With charlist(charindex)
                    For i = 1 To 3
                        If .FxIndex(i) = 4 Or .FxIndex(i) = 5 Or .FxIndex(i) = 6 Or .FxIndex(i) = 42 Or .FxIndex(i) = 43 Or .FxIndex(i) = 44 Or .FxIndex(i) = 45 Or .FxIndex(i) = 16 Or .FxIndex(i) = 103 Or .FxIndex(i) = 104 Or .FxIndex(i) = 105 Then    'If it's meditating, we remove the FX
                            .FxIndex(i) = 0
                            .Fx(i).Loops = 0
                        End If
                    Next i
                    
                ' Play steps sounds if the user is not an admin of any kind
                If .priv = 0 And .montVol = 0 Then
                    Call DoPasosFx(charindex)
                End If
            End With
            
            Call engine.Char_Move_by_Pos(charindex, ReadField(2, rData, 44), ReadField(3, rData, 44))
            Call RefreshAllChars
        Exit Sub
            
        Case "|H"    '>>>> Cambiar Heading Personaje :: |H
            rData = Right$(rData, Len(rData) - 2)
            charindex = Val(ReadField(1, rData, 44))
            
            charlist(charindex).Heading = Val(ReadField(2, rData, 44))
        Exit Sub
        
        Case "|B"    '>>>> Cambiar Body Personaje :: |B
            rData = Right$(rData, Len(rData) - 2)
            charindex = Val(ReadField(1, rData, 44))
            
            charlist(charindex).Body = BodyData(Val(ReadField(2, rData, 44)))
        Exit Sub
        
        Case "|C"    '>>>> Cambiar Casco Personaje :: |C
            rData = Right$(rData, Len(rData) - 2)
            charindex = Val(ReadField(1, rData, 44))
            
            charlist(charindex).Casco = CascoAnimData(Val(ReadField(2, rData, 44)))
        Exit Sub
        
        Case "|E"    '>>>> Cambiar Escudo Personaje :: |E
            rData = Right$(rData, Len(rData) - 2)
            charindex = Val(ReadField(1, rData, 44))
            
            charlist(charindex).Escudo = ShieldAnimData(Val(ReadField(2, rData, 44)))
        Exit Sub
        
        Case "|W"    '>>>> Cambiar Arma Personaje :: |W
            rData = Right$(rData, Len(rData) - 2)
            charindex = Val(ReadField(1, rData, 44))
            
            charlist(charindex).Arma = WeaponAnimData(Val(ReadField(2, rData, 44)))
        Exit Sub
            
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            rData = Right$(rData, Len(rData) - 2)
            
            engine.RemoveCharAparence Val(ReadField(1, rData, 44)), Val(ReadField(3, rData, 44)), Val(ReadField(2, rData, 44)), _
            Val(ReadField(3, rData, 44)), Val(ReadField(4, rData, 44)), Val(ReadField(5, rData, 44)), _
            Val(ReadField(6, rData, 44)), Val(ReadField(9, rData, 44)), Val(ReadField(7, rData, 44)), _
            Val(ReadField(8, rData, 44))
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            rData = Right$(rData, Len(rData) - 2)
            X = Val(ReadField(2, rData, 44))
            Y = Val(ReadField(3, rData, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, rData, 44))
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            rData = Right$(rData, Len(rData) - 2)
            X = Val(ReadField(1, rData, 44))
            Y = Val(ReadField(2, rData, 44))
            MapData(X, Y).ObjGrh.GrhIndex = 0
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            rData = Right$(rData, Len(rData) - 2)
            MapData(Val(ReadField(1, rData, 44)), Val(ReadField(2, rData, 44))).Blocked = Val(ReadField(3, rData, 44))
            Exit Sub
        Case "N~"           ' >>>>> Nombre del Mapa
            rData = Right$(rData, Len(rData) - 2)
            Nombredelmapaxx = rData
        Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            rData = Right$(rData, Len(rData) - 2)
            currentMidi = Val(ReadField(1, rData, 45))
            
            
                If currentMidi <> 0 Then
                    rData = Right$(rData, Len(rData) - Len(ReadField(1, rData, 45)))
                    If Len(rData) > 0 Then
                        If Sound = True Then Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(rData, Len(rData) - 1)))
                    Else
                        If Sound = True Then Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
                    End If
                End If
            
        Exit Sub
        Case "XM"           ' >>>>> Play un MP3 :: XM
            rData = Right$(rData, Len(rData) - 2)
            CurrentMP3 = Val(ReadField(1, rData, 45))
            
            
                If CurrentMP3 <> 0 Then
                    Audio.MP3_Play CurrentMP3
                End If
            
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW

                rData = Right$(rData, Len(rData) - 2)
                 Call Audio.PlayWave(rData & ".wav")

            Exit Sub
        Case "GL" 'Lista de guilds
            rData = Right$(rData, Len(rData) - 2)
            Call frmMenuGral.ParseGuildList(rData)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            If FogataBufferIndex = 0 Then
                FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
            End If
            Exit Sub
        Case "CA"
            CambioDeArea Asc(mid$(sData, 3, 1)), Asc(mid$(sData, 4, 1))
            Exit Sub
    End Select

    Select Case left$(sData, 3)
    
    Case "MAR"
        rData = Right$(rData, Len(rData) - 3)
        
        If Not frmMenuGral.Visible Then frmMenuGral.Show , frmMain
        frmMenuGral.ParseDuelos (rData)
    Exit Sub
    
    Case "ICO" 'INICIO DE COMERCIO, SISTEMA NUEVO BY GHINZUL
        rData = Right$(rData, Len(rData) - 3)
        comIniciar rData
    Exit Sub
    
    Case "IOR" 'RECIVO LA OFERTA (EN ORO)
        rData = Right$(rData, Len(rData) - 3)
        rOro = rData
    Exit Sub
    
    Case "ICI" 'RECIVO LA OFERTA
        rData = Right$(rData, Len(rData) - 3)
        comReciviOferta rData
    Exit Sub
    
    Case "VCC" 'CERRAR COMERCIO PUES
        comCerrar
    Exit Sub
    
    Case "MEC" 'MENSAJE EN CONSOLA
        rData = Right$(rData, Len(rData) - 3)
        comMensaje ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
    Exit Sub
    
    '#####CORREOS####
    Case "IDO"
        rData = Right$(rData, Len(rData) - 3)
        correosIniciar rData
    Exit Sub
    
    Case "IFO"
        rData = Right$(rData, Len(rData) - 3)
        correosIniciarForm rData
    Exit Sub
    
    Case "IAO"
        rData = Right$(rData, Len(rData) - 3)
        correosListaAmigos rData
    Exit Sub
    
    Case "ILO"
        rData = Right$(rData, Len(rData) - 3)
        correosCargarMensaje rData
    Exit Sub
    
    Case "ITO"
        rData = Right$(rData, Len(rData) - 3)
        correosCargarItems rData
    Exit Sub
    '#####CORREOS####
    
    Case "BTM"
        rData = Right$(rData, Len(rData) - 3)
        Batalla.hayBatalla = ReadField(1, rData, 44)
        Batalla.EquipoAzul = Val(ReadField(2, rData, 44))
        Batalla.EquipoAmarillo = Val(ReadField(3, rData, 44))
        Batalla.EquipoRojo = Val(ReadField(4, rData, 44))
        Batalla.EquipoVerde = Val(ReadField(5, rData, 44))
    Exit Sub
    
    Case "NVG"
        rData = Right$(rData, Len(rData) - 3)
        charindex = Val(ReadField(1, rData, 44))
        charlist(charindex).Navegando = Val(ReadField(2, rData, 44))
    Exit Sub
        
    Case "MFC"
        rData = Right$(rData, Len(rData) - 3)
        frmCasas.Show , frmMain
    Exit Sub
    
    Case "TAL"
        rData = Right$(rData, Len(rData) - 3)
        Dim mslink As String
        mslink = ReadField(1, rData, 44)
        
        If MsgBox("Los administradores del juego quieren que veas un link de una pagina web (" & mslink & "). ¿Deseas abrirla?", vbYesNo) = vbYes Then
            OpenBrowser "" & mslink & "", 0
        End If
        
    Exit Sub
 
    Case "GVN"
        rData = Right$(rData, Len(rData) - 3)
        Dim ksitax As String
        Dim prsitox As Long
        Dim fchitax As String
        ksitax = ReadField(1, rData, 44)
        prsitox = ReadField(2, rData, 44)
        fchitax = ReadField(3, rData, 44)
    
        DueñoKsa = ksitax
        Preciox = prsitox
        Fechix = fchitax
        
        If DueñoKsa = "N/A" Then
         DueñoKsa = "DISPONIBLE"
         frmCasas.Command1.Enabled = True
        Else
         frmCasas.Command1.Enabled = False
        End If
       
        frmCasas.lblDueño.Caption = "DUEÑO: " & DueñoKsa
        frmCasas.lblPrecio.Caption = "PRECIO: " & PonerPuntos(prsitox)
        frmCasas.lblFecha.Caption = "FECHA: " & Fechix
    Exit Sub
    
        Case "USM"
            rData = Right$(rData, Len(rData) - 3)
            charindex = Val(ReadField(1, rData, 44))
            charlist(charindex).Montando = Val(ReadField(2, rData, 44))
        Exit Sub
        
        Case "QTL"
            rData = Right(rData, Len(rData) - 3)
            Call frmMenuGral.ParseQuests(rData)
        Exit Sub
        
        Case "MQS"                  ' >>>>> Aceptar quest
            rData = Right$(rData, Len(rData) - 3)
            Nombresiyo = ReadField(1, rData, 44)
            PremioOro = Val(ReadField(2, rData, 44))
            PremioPTS = Val(ReadField(3, rData, 44))
            
            Dim premioCredits, premioTS As Long
            premioCredits = Val(ReadField(4, rData, 44))
            premioTS = Val(ReadField(5, rData, 44))
            
            With frmMenuGral
                .Quests_infoDesc.text = Nombresiyo & vbCrLf
                
                If PremioOro > 0 Then
                   .Quests_infoDesc.text = .Quests_infoDesc.text & vbCrLf & "Oro: " & PonerPuntos(PremioOro)
                End If
                
                If PremioPTS > 0 Then
                    .Quests_infoDesc.text = .Quests_infoDesc.text & vbCrLf & "Puntos de Torneo: " & PonerPuntos(PremioPTS)
                End If
                
                If premioCredits > 0 Then
                    .Quests_infoDesc.text = .Quests_infoDesc.text & vbCrLf & "Créditos: " & premioCredits
                End If
                
                If premioTS > 0 Then
                    .Quests_infoDesc.text = .Quests_infoDesc.text & vbCrLf & "TS Points: " & premioTS
                End If
                
            End With
        Exit Sub
        Case "MQC"                  ' >>>>> Quest en curso
            rData = Right$(rData, Len(rData) - 3)
            
            With frmMenuGral
                .Quests_cursoRequiere.Caption = Val(ReadField(1, rData, 44))
                .Quests_cursoRestantes.Caption = Val(ReadField(2, rData, 44))
                
                .Quests_qDescription.text = ReadField(3, rData, 44)
                 
                .quests_Oro.Caption = PonerPuntos(Val(ReadField(4, rData, 44)))
                .quests_ptsTorneo.Caption = Val(ReadField(5, rData, 44))
                .quests_Credits.Caption = Val(ReadField(6, rData, 44))
                .Quests_ptsTS.Caption = Val(ReadField(7, rData, 44))
            End With
        Exit Sub
        Case "LTR"
            rData = Right(rData, Len(rData) - 3)
            Call frmTorneoManager.PonerListaTorneo(rData)
        Exit Sub
        Case "8G1"
            rData = Right(rData, Len(rData) - 3)
            frmNobleza.lstReq(0).Clear
            Dim noj As Integer
                For noj = 1 To Val(ReadField(1, rData, Asc(",")))
                    frmNobleza.lstReq(0).AddItem ReadField(2 * noj, ReadField(1, rData, Asc("%")), Asc(",")) & " (" & Val(ReadField((2 * noj) + 1, ReadField(1, rData, Asc("%")), Asc(","))) & ")"
                Next noj
            frmNobleza.Show , frmMain
        Exit Sub
        Case "8G2"
            rData = Right(rData, Len(rData) - 3)
            frmNobleza.lstReq(1).Clear
                For noj = 1 To Val(ReadField(1, rData, Asc(",")))
                    frmNobleza.lstReq(1).AddItem ReadField(2 * noj, ReadField(1, rData, Asc("%")), Asc(",")) & " (" & Val(ReadField((2 * noj) + 1, ReadField(1, rData, Asc("%")), Asc(","))) & ")"
                Next noj
            frmNobleza.Show , frmMain
        Exit Sub
        Case "8G3"
            rData = Right(rData, Len(rData) - 3)
            frmNobleza.lstReq(2).Clear
                For noj = 1 To Val(ReadField(1, rData, Asc(",")))
                    frmNobleza.lstReq(2).AddItem ReadField(2 * noj, ReadField(1, rData, Asc("%")), Asc(",")) & " (" & Val(ReadField((2 * noj) + 1, ReadField(1, rData, Asc("%")), Asc(","))) & ")"
                Next noj
            frmNobleza.Show , frmMain
        Exit Sub
        Case "8G4"
            rData = Right(rData, Len(rData) - 3)
            frmNobleza.lstReq(3).Clear
                For noj = 1 To Val(ReadField(1, rData, Asc(",")))
                    frmNobleza.lstReq(3).AddItem ReadField(2 * noj, ReadField(1, rData, Asc("%")), Asc(",")) & " (" & Val(ReadField((2 * noj) + 1, ReadField(1, rData, Asc("%")), Asc(","))) & ")"
                Next noj
            frmNobleza.Show , frmMain
        Exit Sub
        Case "LDM" 'carga lista de amigos
            rData = Right(rData, Len(rData) - 3)
            Call frmMain.PonerListaAmigos(rData)
        Exit Sub
        Case "KFM" 'conecta amigo
        rData = Right(rData, Len(rData) - 3)
            If Configuracion.AnunciarContacto = 1 Then
            AddtoRichTextBox frmMain.RecTxt, "" & UCase$(rData) & " se ha conectado.", 0, 255, 0, True, False, False
            End If
        Exit Sub
        
        Case "DFM" 'desconecta amigo
        rData = Right(rData, Len(rData) - 3)
            If Configuracion.AnunciarContacto = 1 Then
            AddtoRichTextBox frmMain.RecTxt, "" & UCase$(rData) & " se ha desconectado.", 255, 0, 0, True, False, False
            End If
        Exit Sub
        
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            rData = Right$(rData, Len(rData) - 3)
            Call Dialogos.RemoveDialog(Val(rData))
            Exit Sub
        Case "CFF"
            rData = Right$(rData, Len(rData) - 3)
            charindex = Val(ReadField$(1, rData, 44))
            charlist(charindex).particle_count = Val(ReadField$(2, rData, 44))
           
            Call General_Char_Particle_Create(charlist(charindex).particle_count, charindex)
            Call RefreshAllChars
        Exit Sub
        Case "PCL"
            rData = Right$(rData, Len(rData) - 3)
            Dim tmpX As Byte, tmpY As Byte, tmpRange As Byte, tmpRGB(1 To 3) As Byte
            tmpX = Val(ReadField$(1, rData, 44))
            tmpY = Val(ReadField$(2, rData, 44))
            tmpRange = Val(ReadField$(3, rData, 44))
            
            For i = 1 To 3
                tmpRGB(i) = Val(ReadField$(3 + i, rData, 44))
            Next i

            Light.Create_Light_To_Map tmpX, tmpY, tmpRange, tmpRGB(1), tmpRGB(2), tmpRGB(3)
        Exit Sub
        Case "PCB"
            rData = Right$(rData, Len(rData) - 3)
            tmpX = Val(ReadField$(1, rData, 44))
            tmpY = Val(ReadField$(2, rData, 44))

            Light.Delete_Light_To_Map tmpX, tmpY
        Exit Sub
        Case "PCR"
            rData = Right$(rData, Len(rData) - 3)
            
            For i = 1 To 3
                tmpRGB(i) = Val(ReadField$(i, rData, 44))
            Next i

            day_r_old = tmpRGB(1)
            day_g_old = tmpRGB(2)
            day_b_old = tmpRGB(3)
            base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
        Exit Sub
        Case "PCF"
            rData = Right$(rData, Len(rData) - 3)
            Dim Particulita As Byte
            Dim Tiempito As Byte
            Dim equiss As Byte
            Dim equiyy As Byte
            Particulita = Val(ReadField$(1, rData, 44))
            equiss = Val(ReadField$(2, rData, 44))
            equiyy = Val(ReadField$(3, rData, 44))
            Tiempito = Val(ReadField$(4, rData, 44))
                       
            Call General_Particle_Create(Particulita, equiss, equiyy, Tiempito)
        Exit Sub
      Case "CTC"                  ' >>>> Crear particula en char.
            Dim char_index      As Integer
            Dim other_CharIndex As Integer
            Dim particle_Index  As Integer
            Dim particle_Speed  As Single
           
            'Corto la data.
            rData = Right$(rData, Len(rData) - 3)
           
            'Busco el char.
            char_index = Val(ReadField(1, rData, 44))
            other_CharIndex = Val(ReadField(2, rData, 44))
           
            'Datos de la partícula.
            particle_Index = Val(ReadField(3, rData, 44))
            particle_Speed = CSng(ReadField(4, rData, 44))
           
            'Si los chars son válidos y la partícula también.
            If (char_index <> 0) And (other_CharIndex <> 0) Then
               If (particle_Index <> 0) And (particle_Index <= UBound(StreamData())) Then
                  Call engine.Create_Particle(char_index, other_CharIndex, particle_Index, particle_Speed)
               End If
            End If
        Exit Sub
        Case "TIS"
            rData = Right$(rData, Len(rData) - 3)
            Dim typeScroll As Byte
            typeScroll = Val(ReadField(1, rData, 44))
            Scroll(typeScroll).tiempoFaltante = Val(ReadField(2, rData, 44))
            Scroll(typeScroll).tiempoTotal = Val(ReadField(3, rData, 44))
       Exit Sub
       Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            rData = Right$(rData, Len(rData) - 3)
            charindex = Val(ReadField(1, rData, 44))
            Call SetCharacterFx(charindex, Val(ReadField(2, rData, 44)), Val(ReadField(3, rData, 44)))
        Exit Sub
       Case "CFE"                  ' >>>>> Mostrar Emoticones :: CFE
            rData = Right$(rData, Len(rData) - 3)
            charindex = Val(ReadField(1, rData, 44))
            charlist(charindex).EmoticonIndex = Val(ReadField(2, rData, 44))
            charlist(charindex).EmoticonLoops = Val(ReadField(3, rData, 44))
            Call SetCharacterEmoticon(charindex, charlist(charindex).EmoticonIndex, charlist(charindex).EmoticonLoops)
        Exit Sub
        Case "ANM"
       
        rData = Right$(rData, Len(rData) - 3)
            ArmaMin = Val(ReadField(1, rData, 44))
            ArmaMax = Val(ReadField(2, rData, 44))
            ArmorMin = Val(ReadField(3, rData, 44))
            ArmorMax = Val(ReadField(4, rData, 44))
            EscuMin = Val(ReadField(5, rData, 44))
            EscuMax = Val(ReadField(6, rData, 44))
            CascMin = Val(ReadField(7, rData, 44))
            CascMax = Val(ReadField(8, rData, 44))
            HerrMin = Val(ReadField(9, rData, 44))
            HerrMax = Val(ReadField(10, rData, 44))
            MagMin = Val(ReadField(11, rData, 44))
            MagMax = Val(ReadField(12, rData, 44))
            MagMina = Val(ReadField(13, rData, 44))
            MagMaxa = Val(ReadField(14, rData, 44))
            MagMinb = Val(ReadField(15, rData, 44))
            MagMaxb = Val(ReadField(16, rData, 44))
            MagMinc = Val(ReadField(17, rData, 44))
            MagMaxc = Val(ReadField(18, rData, 44))
            MagMind = Val(ReadField(19, rData, 44))
            MagMaxd = Val(ReadField(20, rData, 44))
 
        With frmMain
                .Arma.Caption = ArmaMin & "/" & ArmaMax
                .Defensa.Caption = ArmorMin + EscuMin + CascMin + HerrMin & "/" & ArmorMax + EscuMax + CascMax + HerrMax
                .DefMag.Caption = MagMin + MagMina + MagMinb + MagMinc + MagMind & "/" & MagMax + MagMaxa + MagMaxb + MagMaxc + MagMaxd
        End With
        
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim N As String, n2 As String
            rData = Right$(rData, Len(rData) - 3)
            N = ReadField(2, rData, 176)
            n2 = ReadField(1, rData, 176)
            frmMSG.CrearGMmSg N, n2
            frmMSG.Show , frmMain
        Exit Sub
        
        Case "LDG"
            rData = Right$(rData, Len(rData) - 3)
            UserPrivilegios = Val(rData)
            If UserPrivilegios = 0 Then
                frmMain.GMSOS.Visible = False
                frmMain.GMTORNEO.Visible = False
                frmMain.GMPANEL.Visible = False
            Else
                frmMain.GMSOS.Visible = True
                frmMain.GMTORNEO.Visible = True
                frmMain.GMPANEL.Visible = True
            End If
        Exit Sub
        Case "CHX"
            rData = Right$(rData, Len(rData) - 3)
            UserMaxHPCHORI = Val(ReadField(1, rData, 44))
            UserMinHPCHORI = Val(ReadField(2, rData, 44))
            UserMaxMANCHORI = Val(ReadField(3, rData, 44))
            UserMinMANCHORI = Val(ReadField(4, rData, 44))
            NickCHORI = ReadField(5, rData, 44)
           
            If Form1.Visible = False Then
                Form1.Show
            End If
            
            Form1.Shape1.Width = (((UserMinHPCHORI / 100) / (UserMaxHPCHORI / 100)) * 1695)
            Form1.Shape2.Width = (((UserMinMANCHORI / 100) / (UserMaxMANCHORI / 100)) * 1695)
            Form1.Label1.Caption = UserMinHPCHORI & "/" & UserMaxHPCHORI
            Form1.Label2.Caption = UserMinMANCHORI & "/" & UserMaxMANCHORI
            Form1.Caption = NickCHORI
           
        Exit Sub
        Case "VOT"
        rData = Right$(rData, Len(rData) - 3)
        Dim Vot(1 To 5) As String, Votacion As String
        Vot(1) = ReadField(1, rData, 44)
        Vot(2) = ReadField(2, rData, 44)
        Vot(3) = ReadField(3, rData, 44)
        Vot(4) = ReadField(4, rData, 44)
        Vot(5) = ReadField(5, rData, 44)
        Votacion = ReadField(6, rData, 44)
        
            With frmVotacionUser
            
                For i = 1 To 5
                    If Vot(i) = "" Then
                        .Votos(i - 1).Enabled = False
                        .Votos(i - 1).Caption = "N/A"
                    Else
                        .Votos(i - 1).Enabled = True
                        .Votos(i - 1).Caption = Vot(i)
                    End If
                Next i
                
                .Label1.Caption = Votacion
                .Show , frmMain
            
            End With
        Exit Sub
        
        Case "WEN"
            FrmNewPoll.Show , frmMain
        Exit Sub
        
     
         Case "BYE"
        rData = Right$(rData, Len(rData) - 3)
            With frmVotacionUser
                .Label1.Caption = ""
                
                For i = 1 To 5
                    .Votos(i - 1).Enabled = False
                    .Votos(i - 1).Caption = "N/A"
                Next i
            End With
            
            Unload frmVotacionUser
        Exit Sub
        
        Case "IFE"
            rData = Right$(rData, Len(rData) - 3)
            
            frmEstadisticasUsuario.lblNombre.Caption = ReadField(1, rData, 44)
            frmEstadisticasUsuario.lblClase.Caption = ReadField(2, rData, 44)
            frmEstadisticasUsuario.lblRaza.Caption = ReadField(3, rData, 44)
            frmEstadisticasUsuario.lblNivel = ReadField(4, rData, 44)
            frmEstadisticasUsuario.lblExp.Caption = ReadField(5, rData, 44)
            frmEstadisticasUsuario.lblFaccion = ReadField(6, rData, 44)
            frmEstadisticasUsuario.lblJerarquia.Caption = ReadField(7, rData, 44)
            frmEstadisticasUsuario.lblReputacion.Caption = ReadField(8, rData, 44)
            frmEstadisticasUsuario.lblDuelos.Caption = ReadField(9, rData, 44)
            frmEstadisticasUsuario.lblParejas.Caption = ReadField(10, rData, 44)
            frmEstadisticasUsuario.lblRondas.Caption = ReadField(11, rData, 44)
            frmEstadisticasUsuario.lblMuertes.Caption = ReadField(12, rData, 44)
            frmEstadisticasUsuario.lblUsuariosMatados.Caption = ReadField(13, rData, 44)
            frmEstadisticasUsuario.lblEventos.Caption = ReadField(14, rData, 44)
            frmEstadisticasUsuario.lblCVCS.Caption = ReadField(15, rData, 44)
            frmEstadisticasUsuario.lblQuests.Caption = ReadField(16, rData, 44)

            frmEstadisticasUsuario.Show , frmMain

        Exit Sub
        
        Case "AU|"
            rData = Right$(rData, Len(rData) - 3)
            Dim Armadura As Integer
            Dim Weapon As Integer
            Dim EscudoA As Integer
            Dim Ring As Integer
            Dim CascoA As Integer
            
            charindex = Val(ReadField(1, rData, 44))
            Armadura = Val(ReadField(2, rData, 44))
            Weapon = Val(ReadField(3, rData, 44))
            EscudoA = Val(ReadField(4, rData, 44))
            Ring = Val(ReadField(5, rData, 44))
            CascoA = Val(ReadField(6, rData, 44))
            
            
            If Armadura > 0 Then
                charlist(charindex).Aura_IndexA = Armadura
                
                If AurasPJ(charlist(charindex).Aura_IndexA).RojoF > 0 Then
                    charlist(charindex).AuraAntiguoR = AurasPJ(charlist(charindex).Aura_IndexA).RojoF
                    charlist(charindex).AuraAntiguoG = AurasPJ(charlist(charindex).Aura_IndexA).VerdeF
                    charlist(charindex).AuraAntiguoB = AurasPJ(charlist(charindex).Aura_IndexA).AzulF
                        
                    charlist(charindex).AuraQueremosLlegarR = AurasPJ(charlist(charindex).Aura_IndexA).r
                    charlist(charindex).AuraQueremosLlegarG = AurasPJ(charlist(charindex).Aura_IndexA).g
                    charlist(charindex).AuraQueremosLlegarB = AurasPJ(charlist(charindex).Aura_IndexA).b
                    
                    charlist(charindex).AuraProximoR = AurasPJ(charlist(charindex).Aura_IndexA).RojoF
                    charlist(charindex).AuraProximoG = AurasPJ(charlist(charindex).Aura_IndexA).VerdeF
                    charlist(charindex).AuraProximoB = AurasPJ(charlist(charindex).Aura_IndexA).AzulF
                    charlist(charindex).AuraLlegoAlColor = True
                    
                End If
                
                Call InitGrh(charlist(charindex).AuraA, AurasPJ(charlist(charindex).Aura_IndexA).GrhIndex)
                charlist(charindex).Aura_AngleA = 0
            Else
                charlist(charindex).Aura_IndexA = 0
            End If
            
            If Ring > 0 Then
                charlist(charindex).Aura_IndexR = Ring
                Call InitGrh(charlist(charindex).AuraR, AurasPJ(charlist(charindex).Aura_IndexR).GrhIndex)
                charlist(charindex).Aura_AngleR = 0
            Else
                charlist(charindex).Aura_IndexR = 0
            End If
            
            
            If Weapon > 0 Then
                charlist(charindex).Aura_IndexW = Weapon
                Call InitGrh(charlist(charindex).AuraW, AurasPJ(charlist(charindex).Aura_IndexW).GrhIndex)
                charlist(charindex).Aura_AngleW = 0
            Else
                charlist(charindex).Aura_IndexW = 0
            End If
                
                
            If CascoA > 0 Then
                charlist(charindex).Aura_IndexC = CascoA
                Call InitGrh(charlist(charindex).AuraC, AurasPJ(charlist(charindex).Aura_IndexC).GrhIndex)
                charlist(charindex).Aura_AngleC = 0
            Else
                charlist(charindex).Aura_IndexC = 0
            End If
                
            If EscudoA > 0 Then
                charlist(charindex).Aura_IndexE = EscudoA
                Call InitGrh(charlist(charindex).AuraE, AurasPJ(charlist(charindex).Aura_IndexE).GrhIndex)
                charlist(charindex).Aura_AngleE = 0
            Else
                charlist(charindex).Aura_IndexE = 0
            End If
        Exit Sub
        
        Case "[ES" 'Actualiza estadisticas completas
             rData = Right$(rData, Len(rData) - 3)
             UserMaxHP = Val(ReadField(1, rData, 44))
             UserMinHP = Val(ReadField(2, rData, 44))
             UserMaxMAN = Val(ReadField(3, rData, 44))
             UserMinMAN = Val(ReadField(4, rData, 44))
             UserMaxSTA = Val(ReadField(5, rData, 44))
             UserMinSTA = Val(ReadField(6, rData, 44))
             UserGLD = Val(ReadField(7, rData, 44))
             UserLvl = Val(ReadField(8, rData, 44))
             UserPasarNivel = Val(ReadField(9, rData, 44))
             UserExp = Val(ReadField(10, rData, 44))
             NickPJ = ReadField(11, rData, 44)
             UserReputacione = Val(ReadField(14, rData, 44))

            'Seteamos el shape y el label de vida
            frmMain.HpSHP.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 118)
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
                
            'Seteamos que el usuario murió
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
             
             'Seteamos el shape y el label de vida
            If UserMaxMAN > 0 Then
                frmMain.MPShp.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 118)
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            Else
                frmMain.MPShp.Width = 0
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            End If
            
            'Seteamos el shape y label de energia
            frmMain.SPShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 118)
            frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA

            frmMain.GldLbl.Caption = PonerPuntos(UserGLD)

            'Seteamos el label de nivel
            If UserLvl > 50 Then
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.LvlLbl.ForeColor = vbYellow
            Else
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.LvlLbl.ForeColor = vbRed
            End If

            'Seteamos ancho de barra, label y experiencia, todo junto.
            If UserPasarNivel > 0 Then
                frmMain.ExpBar.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 264)
                frmMain.exp.Caption = Round(UserExp) & "/" & Round(UserPasarNivel) & ""
            Else
                frmMain.exp.Caption = "¡Nivel Máx!"
                frmMain.ExpBar.Width = 264
            End If

            'label de nombre
             frmMain.Label8.Caption = NickPJ

            frmMain.Agilidad.Caption = ReadField(12, rData, 44)
            frmMain.Fuerza.Caption = ReadField(13, rData, 44)
            
             'Label de reputación
             If UserReputacione < 0 Then
                 frmMain.rep.Caption = "- " & PonerPuntos(UserReputacione)
                 frmMain.rep.ForeColor = vbRed
             Else
                 frmMain.rep.Caption = PonerPuntos(UserReputacione)
                 frmMain.rep.ForeColor = vbWhite
             End If
        Exit Sub
        
        Case "[EZ" 'Actualiza estadisticas parciales
             rData = Right$(rData, Len(rData) - 3)
             UserMaxHP = Val(ReadField(1, rData, 44))
             UserMinHP = Val(ReadField(2, rData, 44))
             UserMaxMAN = Val(ReadField(3, rData, 44))
             UserMinMAN = Val(ReadField(4, rData, 44))
             UserMaxSTA = Val(ReadField(5, rData, 44))
             UserMinSTA = Val(ReadField(6, rData, 44))

            'Seteamos el shape y el label de vida
            frmMain.HpSHP.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 118)
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
                
            'Seteamos que el usuario murió
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
             
             'Seteamos el shape y el label de vida
            If UserMaxMAN > 0 Then
                frmMain.MPShp.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 118)
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            Else
                frmMain.MPShp.Width = 0
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            End If
            
            'Seteamos el shape y label de energia
            frmMain.SPShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 118)
            frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
        Exit Sub
        
        Case "[CD" 'Actualiza char data
             rData = Right$(rData, Len(rData) - 3)
             charindex = Val(ReadField(1, rData, 44))
             charlist(charindex).color = Val(ReadField(2, rData, 44))
             Armadura = Val(ReadField(3, rData, 44))
             Weapon = Val(ReadField(4, rData, 44))
             EscudoA = Val(ReadField(5, rData, 44))
             Ring = Val(ReadField(6, rData, 44))
             CascoA = Val(ReadField(7, rData, 44))
             charlist(charindex).montVol = Val(ReadField(8, rData, 44))
             charlist(charindex).posRank = Val(ReadField(9, rData, 44))
             
             If Armadura > 0 Then
                charlist(charindex).Aura_IndexA = Armadura
                
                If AurasPJ(charlist(charindex).Aura_IndexA).RojoF > 0 Then
                    charlist(charindex).AuraAntiguoR = AurasPJ(charlist(charindex).Aura_IndexA).RojoF
                    charlist(charindex).AuraAntiguoG = AurasPJ(charlist(charindex).Aura_IndexA).VerdeF
                    charlist(charindex).AuraAntiguoB = AurasPJ(charlist(charindex).Aura_IndexA).AzulF
                        
                    charlist(charindex).AuraQueremosLlegarR = AurasPJ(charlist(charindex).Aura_IndexA).r
                    charlist(charindex).AuraQueremosLlegarG = AurasPJ(charlist(charindex).Aura_IndexA).g
                    charlist(charindex).AuraQueremosLlegarB = AurasPJ(charlist(charindex).Aura_IndexA).b
                    
                    charlist(charindex).AuraProximoR = AurasPJ(charlist(charindex).Aura_IndexA).RojoF
                    charlist(charindex).AuraProximoG = AurasPJ(charlist(charindex).Aura_IndexA).VerdeF
                    charlist(charindex).AuraProximoB = AurasPJ(charlist(charindex).Aura_IndexA).AzulF
                    charlist(charindex).AuraLlegoAlColor = True
                    
                End If
                
                Call InitGrh(charlist(charindex).AuraA, AurasPJ(charlist(charindex).Aura_IndexA).GrhIndex)
                charlist(charindex).Aura_AngleA = 0
            Else
                charlist(charindex).Aura_IndexA = 0
            End If
            
            If Ring > 0 Then
                charlist(charindex).Aura_IndexR = Ring
                Call InitGrh(charlist(charindex).AuraR, AurasPJ(charlist(charindex).Aura_IndexR).GrhIndex)
                charlist(charindex).Aura_AngleR = 0
            Else
                charlist(charindex).Aura_IndexR = 0
            End If
            
            
            If Weapon > 0 Then
                charlist(charindex).Aura_IndexW = Weapon
                Call InitGrh(charlist(charindex).AuraW, AurasPJ(charlist(charindex).Aura_IndexW).GrhIndex)
                charlist(charindex).Aura_AngleW = 0
            Else
                charlist(charindex).Aura_IndexW = 0
            End If
                
                
            If CascoA > 0 Then
                charlist(charindex).Aura_IndexC = CascoA
                Call InitGrh(charlist(charindex).AuraC, AurasPJ(charlist(charindex).Aura_IndexC).GrhIndex)
                charlist(charindex).Aura_AngleC = 0
            Else
                charlist(charindex).Aura_IndexC = 0
            End If
                
            If EscudoA > 0 Then
                charlist(charindex).Aura_IndexE = EscudoA
                Call InitGrh(charlist(charindex).AuraE, AurasPJ(charlist(charindex).Aura_IndexE).GrhIndex)
                charlist(charindex).Aura_AngleE = 0
            Else
                charlist(charindex).Aura_IndexE = 0
            End If
             
        Exit Sub
        
        Case "[H]" 'Actualiza vida
             rData = Right$(rData, Len(rData) - 3)
             UserMaxHP = Val(ReadField(1, rData, 44))
             UserMinHP = Val(ReadField(2, rData, 44))
             
            'Seteamos el shape y el label de vida
            frmMain.HpSHP.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 118)
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
                
            'Seteamos que el usuario murió
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        Exit Sub
        
        Case "[M]" 'Actualiza mana
             rData = Right$(rData, Len(rData) - 3)
             UserMaxMAN = Val(ReadField(1, rData, 44))
             UserMinMAN = Val(ReadField(2, rData, 44))
             
             'Seteamos el shape y el label de vida
            If UserMaxMAN > 0 Then
                frmMain.MPShp.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 118)
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            Else
                frmMain.MPShp.Width = 0
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            End If
        Exit Sub
        
        Case "[S]" 'Actualiza stamina
             rData = Right$(rData, Len(rData) - 3)
             UserMaxSTA = Val(ReadField(1, rData, 44))
             UserMinSTA = Val(ReadField(2, rData, 44))
             
             'Seteamos el shape y label de energia
             frmMain.SPShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 118)
             frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
        Exit Sub
        
        Case "[G]" 'Actualiza oro
             rData = Right$(rData, Len(rData) - 3)
             UserGLD = Val(ReadField(1, rData, 44))
            
             frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
        Exit Sub
        
        Case "|S1" 'Actualiza slot
             rData = Right$(rData, Len(rData) - 3)
            
             Call Inventario.ActualizarSlotCant(ReadField(1, rData, 44), ReadField(2, rData, 44))
        Exit Sub
        
        Case "|S2" 'Actualiza slot
             rData = Right$(rData, Len(rData) - 3)
            
             Call Inventario.ActualizarSlotEquipped(ReadField(1, rData, 44), ReadField(2, rData, 44))
        Exit Sub
        
        Case "[L]" 'Actualiza nivel
             rData = Right$(rData, Len(rData) - 3)
             UserLvl = Val(ReadField(1, rData, 44))
            
            'Seteamos el label de nivel
            If UserLvl > 50 Then
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.LvlLbl.ForeColor = vbYellow
            Else
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.LvlLbl.ForeColor = vbRed
            End If
        Exit Sub
        
        Case "[E]" 'Actualizar experiencia
             rData = Right$(rData, Len(rData) - 3)
             UserPasarNivel = Val(ReadField(1, rData, 44))
             UserExp = Val(ReadField(2, rData, 44))

            'Seteamos ancho de barra, label y experiencia, todo junto.
            If UserPasarNivel > 0 Then
                frmMain.ExpBar.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 264)
                frmMain.exp.Caption = Round(UserExp) & "/" & Round(UserPasarNivel) & ""
            Else
                frmMain.exp.Caption = "¡Nivel Máx!"
                frmMain.ExpBar.Width = 264
            End If
             
        Exit Sub
        
        Case "[B]" 'Actualizar oro de la boveda
             rData = Right$(rData, Len(rData) - 3)
             UserBOVItem = Val(ReadField(1, rData, 44))
             
             'Seteo label de boveda
             frmBanco.Text1.Caption = PonerPuntos(UserBOVItem)
        Exit Sub
        
        Case "[N]" 'Actualiza el nombre
             rData = Right$(rData, Len(rData) - 3)
             NickPJ = ReadField(1, rData, 44)
             
             'label de nombre
             frmMain.Label8.Caption = NickPJ
        Exit Sub
        
        Case "[F]" 'Actualiza fuerza
             rData = Right$(rData, Len(rData) - 3)
             frmMain.Fuerza.Caption = ReadField(1, rData, 44)
        Exit Sub
        
        Case "[BG" 'Actualiza bank gold
             rData = Right$(rData, Len(rData) - 3)
             frmBanco.Text1.Caption = PonerPuntos(Val(ReadField(1, rData, 44)))
        Exit Sub
                
        Case "[A]" 'Actualiza agilidad
             rData = Right$(rData, Len(rData) - 3)
             frmMain.Agilidad.Caption = ReadField(1, rData, 44)
        Exit Sub
        
        Case "[R]" 'Actualizar reputación
             rData = Right$(rData, Len(rData) - 3)
             UserReputacione = Val(ReadField(1, rData, 44))
            
                'Label de reputación
                If UserReputacione < 0 Then
                    frmMain.rep.Caption = "- " & PonerPuntos(UserReputacione)
                    frmMain.rep.ForeColor = vbRed
                Else
                    frmMain.rep.Caption = PonerPuntos(UserReputacione)
                    frmMain.rep.ForeColor = vbWhite
                End If
        Exit Sub
        
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            rData = Right$(rData, Len(rData) - 3)
            UsingSkill = Val(rData)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            rData = Right$(rData, Len(rData) - 3)
            slot = ReadField(1, rData, 44)
            Call Inventario.SetItem(slot, ReadField(2, rData, 44), ReadField(4, rData, 44), ReadField(5, rData, 44), Val(ReadField(6, rData, 44)), Val(ReadField(7, rData, 44)), _
                                    Val(ReadField(8, rData, 44)), Val(ReadField(9, rData, 44)), Val(ReadField(10, rData, 44)), Val(ReadField(11, rData, 44)), ReadField(3, rData, 44))
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************

        Case "SBR"                 '
            rData = Right$(rData, Len(rData) - 3)
            
            For i = 1 To MAX_BANCOINVENTORY_SLOTS
                UserBancoInventory(i).OBJIndex = 0
                UserBancoInventory(i).Name = ""
                UserBancoInventory(i).Amount = 0
                UserBancoInventory(i).GrhIndex = 0
                UserBancoInventory(i).OBJType = 0
                UserBancoInventory(i).MaxHit = 0
                UserBancoInventory(i).MinHit = 0
                UserBancoInventory(i).Def = 0
            Next i
            
        Exit Sub

        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            rData = Right$(rData, Len(rData) - 3)
            
            slot = ReadField(1, rData, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, rData, 44)
            UserBancoInventory(slot).Name = ReadField(3, rData, 44)
            UserBancoInventory(slot).Amount = ReadField(4, rData, 44)
            UserBancoInventory(slot).GrhIndex = Val(ReadField(5, rData, 44))
            UserBancoInventory(slot).OBJType = Val(ReadField(6, rData, 44))
            UserBancoInventory(slot).MaxHit = Val(ReadField(7, rData, 44))
            UserBancoInventory(slot).MinHit = Val(ReadField(8, rData, 44))
            UserBancoInventory(slot).Def = Val(ReadField(9, rData, 44))
        
            tempstr = ""
            
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(slot).Name
            End If
            
        Exit Sub
        
        Case "SBG"                 ' >>>>> Actualiza Cuenta Bancaria
            rData = Right$(rData, Len(rData) - 3)
            slot = ReadField(1, rData, 44)
            UserBancoInventoryB(slot).OBJIndex = ReadField(2, rData, 44)
            UserBancoInventoryB(slot).Name = ReadField(3, rData, 44)
            UserBancoInventoryB(slot).Amount = ReadField(4, rData, 44)
            UserBancoInventoryB(slot).GrhIndex = Val(ReadField(5, rData, 44))
            UserBancoInventoryB(slot).OBJType = Val(ReadField(6, rData, 44))
            UserBancoInventoryB(slot).MaxHit = Val(ReadField(7, rData, 44))
            UserBancoInventoryB(slot).MinHit = Val(ReadField(8, rData, 44))
            UserBancoInventoryB(slot).Def = Val(ReadField(9, rData, 44))
            UserBancoOro = Val(ReadField(10, rData, 44))
            UserBancoOroPropio = Val(ReadField(11, rData, 44))
        
            tempstr = ""
            
            If UserBancoInventoryB(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventoryB(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventoryB(slot).Name
            End If
            
        Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            rData = Right$(rData, Len(rData) - 3)
            slot = ReadField(1, rData, 44)
            UserHechizos(slot) = ReadField(2, rData, 44)
            frmMain.hlst.List(slot - 1) = ReadField(3, rData, 44)
            Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            rData = Right$(rData, Len(rData) - 3)
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(ReadField(i, rData, 44))
            Next i
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            rData = Right$(rData, Len(rData) - 3)
            
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, rData, 44)
                ArmasHerrero(m) = Val(ReadField(i + 1, rData, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            rData = Right$(rData, Len(rData) - 3)
            
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, rData, 44)
                ArmadurasHerrero(m) = Val(ReadField(i + 1, rData, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            rData = Right$(rData, Len(rData) - 3)
            
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, rData, 44)
                ObjCarpintero(m) = Val(ReadField(i + 1, rData, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
        Case "SPL"
            rData = Right(rData, Len(rData) - 3)
            For i = 1 To Val(ReadField(1, rData, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, rData, 44)
            Next i
            frmSpawnList.Show , frmMain
        Exit Sub
        
        Case "DRM"
            rData = Right(rData, Len(rData) - 3)
            
            frmMenuGral.ParseCreditos (rData)
        Exit Sub
        
        Case "DNF"
            rData = Right(rData, Len(rData) - 3)
            Dim ContentDonation As String
            Dim TempDonation As String
            
            picDonation.Body = ReadField(1, rData, 44)
            picDonation.Head = ReadField(2, rData, 44)
            picDonation.Weapon = ReadField(3, rData, 44)
            picDonation.Shield = ReadField(4, rData, 44)
            picDonation.Casco = ReadField(5, rData, 44)
            picDonation.Aura = ReadField(6, rData, 44)
            picDonation.GrhIndex = ReadField(7, rData, 44)
            
            If picDonation.Aura > 0 Then
                Call InitGrh(picDonation.AuraA, AurasPJ(picDonation.Aura).GrhIndex)
                picDonation.Aura_Angle = 0
            End If
            
            With frmMenuGral
                    .creditos_Desc.text = ""
                    AddtoRichTextBox .creditos_Desc, ReadField(9, rData, 44), 255, 255, 255
            
                    ContentDonation = ""
                    For i = 1 To ReadField(10, rData, 44)
                        TempDonation = ReadField(10 + i, rData, 44)
                        ContentDonation = ContentDonation & "• " & ReadField(1, TempDonation, Asc("-")) & " - " & ReadField(2, TempDonation, Asc("-")) & vbCrLf & ""
                    Next i
                    
                    .creditos_lblContent.Caption = ContentDonation
                    .creditos_lblPrice = PonerPuntos(ReadField(8, rData, 44))
            End With
        Exit Sub
        
        Case "PRM"
                rData = Right(rData, Len(rData) - 3)
               
                frmMenuGral.ParseCanjes (rData)
        Exit Sub
               
            Case "INF"
                rData = Right(rData, Len(rData) - 3)
                
            With frmMenuGral
                    .Requiere.Caption = ReadField(1, rData, 44)
                    .lAtaque.Caption = ReadField(3, rData, 44) & "/" & ReadField(2, rData, 44)
                    .lDef.Caption = ReadField(5, rData, 44) & "/" & ReadField(4, rData, 44)
                    .lAM.Caption = ReadField(7, rData, 44) & "/" & ReadField(6, rData, 44)
                    .lDM.Caption = ReadField(9, rData, 44) & "/" & ReadField(8, rData, 44)
                    .lDescripcion.text = ReadField(10, rData, 44)
                    .lPuntos.Caption = ReadField(11, rData, 44)
                    
            CantidadCanjeYegua = ReadField(1, rData, 44)
            
                If .Requiere.Caption = "0" Then
                    .Requiere.Caption = "N/A"
                End If
                
                If .lAtaque.Caption = "0/0" Then
                    .lAtaque.Caption = "N/A"
                End If
                
                If .lDef.Caption = "0/0" Then
                    .lDef.Caption = "N/A"
                End If
                
                If .lAM.Caption = "0/0" Then
                    .lAM.Caption = "N/A"
                End If
                
                If .lDM.Caption = "0/0" Then
                    .lDM.Caption = "N/A"
                End If

                Dim Grhpremios As Integer
                Grhpremios = ReadField(12, rData, 44)
                    Dim SR As RECT
                    
                    SR.left = 0
                    SR.top = 0
                    SR.Right = 32
                    SR.bottom = 32
                    Call engine.DrawGrhtoHdc(Grhpremios, SR, frmMenuGral.picObj)
            End With
        Exit Sub
            
        Case "ERO"
            rData = Right$(rData, Len(rData) - 3)
        
            Mensaje.Label1 = rData
            Mensaje.Show
        
        Exit Sub
        
        Case "ERR"
            rData = Right$(rData, Len(rData) - 3)
            frmConnect.MousePointer = 1
            
            Mensaje.Escribir rData
        Exit Sub
    End Select
    
    
    Select Case left$(sData, 4)
        Case "MTOP"
        rData = Right$(rData, Len(rData) - 4)
        
            Call frmRanking.MostrarRanking(rData)
        Exit Sub
    
        Case "RANK"
        rData = Right$(rData, Len(rData) - 4)
            charindex = ReadField(1, rData, Asc(","))
            
            charlist(charindex).posRank = ReadField(2, rData, Asc(","))
        Exit Sub
        
        Case "ZSOS"
        rData = Right$(rData, Len(rData) - 4)
        MensajesNumber = ReadField(1, rData, Asc("|"))
        
        Dim SOSTemporal As String
        frmGmPanelSOS.UserSOSList.Clear
        SOSTemporal = ""
        
            For i = 1 To MensajesNumber
                SOSTemporal = ReadField(1 + i, rData, Asc("|"))
                MensajesSOS(i).Tipo = ReadField(1, SOSTemporal, Asc("-"))
                MensajesSOS(i).Autor = ReadField(2, SOSTemporal, Asc("-"))
                MensajesSOS(i).Contenido = ReadField(3, SOSTemporal, Asc("-"))
                frmGmPanelSOS.UserSOSList.AddItem "[" & MensajesSOS(i).Tipo & "] - " & MensajesSOS(i).Autor
                frmGmPanelSOS.UserSOSList.Refresh
            Next i
        
        Exit Sub
    
        Case "ARIE"
            rData = Right$(rData, Len(rData) - 4)
            charindex = ReadField(1, rData, 44)
            charlist(charindex).Ariete = True
        Exit Sub
        
        Case "MVOL"
            rData = Right$(rData, Len(rData) - 4)
            charindex = ReadField(1, rData, 44)
            charlist(charindex).montVol = Val(ReadField(2, rData, 44))
        Exit Sub
    
        Case "MJOR"
            rData = Right$(rData, Len(rData) - 4)
            
            FrmMejorar.ListaMejorados.AddItem rData
            
            
            If UCase$(rData) = "SIN ITEMS MEJORABLES" Then
                FrmMejorar.ListaMejorados.Enabled = False
            Else
                FrmMejorar.ListaMejorados.Enabled = True
            End If
            
            FrmMejorar.Show , frmMain
        Exit Sub
            
        Case "IMEJ"
            rData = Right$(rData, Len(rData) - 4)
            
            With FrmMejorar
            
            .Nombre.Caption = ReadField(1, rData, 44)
            .Ataque.Caption = ReadField(2, rData, 44)
            .Defensa.Caption = ReadField(3, rData, 44)
            .AtaqueMagico.Caption = ReadField(4, rData, 44)
            .DefensaMagica.Caption = ReadField(5, rData, 44)
            .Desc.text = ReadField(6, rData, 44)
            
            SR.bottom = 32
            SR.Right = 32
            SR.left = 0
            SR.top = 0
            
                Dim GrhMejorar As Integer
                    GrhMejorar = ReadField(7, rData, 44)
                    Call engine.DrawGrhtoHdc(GrhMejorar, SR, .Item)
            End With
            
        Exit Sub
        Case "GODS"
        rData = Right$(rData, Len(rData) - 4)
        Dim AlmasOfrecidas As Long
        Dim AlmasNecesarias As Long
        Dim SirvienteDe As String
        AlmasOfrecidas = Val(ReadField(1, rData, 44))
        AlmasNecesarias = Val(ReadField(2, rData, 44))
        SirvienteDe = ReadField(3, rData, 44)
        
        frmGods.lblOfrecidos = "" & AlmasOfrecidas & "/" & AlmasNecesarias & ""
        frmGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Main.jpg")
        
        If UCase$(SirvienteDe) = "MIFRIT" Then
         frmGods.imgAlmas.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_MifritBar.jpg")
         frmGods.imgGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Dios2.jpg")
        ElseIf UCase$(SirvienteDe) = "TARRASKE" Then
         frmGods.imgAlmas.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_TarraskeBar.jpg")
         frmGods.imgGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Dios4.jpg")
        ElseIf UCase$(SirvienteDe) = "EREBROS" Then
         frmGods.imgAlmas.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_ErebrosBar.jpg")
         frmGods.imgGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Dios1.jpg")
        ElseIf UCase$(SirvienteDe) = "POSEIDON" Then
         frmGods.imgAlmas.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_PoseidonBar.jpg")
         frmGods.imgGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Dios3.jpg")
        End If
        
        If AlmasOfrecidas = 0 Then
            frmGods.imgAlmas.Width = 0
        Else
            frmGods.imgAlmas.Width = (((AlmasOfrecidas / 100) / (AlmasNecesarias / 100)) * 335)
        End If
               
        frmGods.Show , frmMain
         
        Exit Sub
        
        Case "PCCC" ' CHOTS | Poner Captions en frm
            Dim Caption As String
            Dim Nomvre As String
            rData = Right$(rData, Len(rData) - 4)
            Caption = ReadField(1, rData, 44)
            Nomvre = ReadField(2, rData, 44)
            Call frmProcesos.Show
            frmProcesos.Procesos.Visible = True
            frmProcesos.Captions.Visible = False
            frmProcesos.Command1.Enabled = False
            frmProcesos.Command2.Enabled = True
            frmProcesos.Captions.AddItem Caption
        Case "PCCP" ' CHOTS | Listar Captions
            frmProcesos.Captions.Clear
            rData = Right$(rData, Len(rData) - 4)
            charindex = Val(ReadField(1, rData, 44))
            Call frmProcesos.Listar(charindex)
            Exit Sub
        Case "PCGR" ' CHOTS | Listar Procesos
            rData = Right$(rData, Len(rData) - 4)
            charindex = Val(ReadField(1, rData, 44))
            indiceProc = 0
            Call enumProc(charindex)
        Exit Sub
        Case "PCGN" ' CHOTS | Poner Procesos en frm
            Dim Proceso As String, tmpPeso As Long
            Dim Nombre As String
            rData = Right$(rData, Len(rData) - 4)
            Proceso = ReadField(1, rData, 44)
            tmpPeso = ReadField(2, rData, 44)
            Nombre = ReadField(3, rData, 44)
            frmProcesos.Caption = Nombre
            frmProcesos.txtUrl.text = Nombre
            indiceProc = indiceProc + 1
            frmProcesos.Procesos.ListItems.Add indiceProc, , Proceso
            frmProcesos.Procesos.ListItems(indiceProc).ListSubItems.Add , , PonerPuntos(tmpPeso) & " kbs"
        Exit Sub
        
        Case "MENU"
        If Configuracion.MenuDesplegable = 0 Then Exit Sub
                Dim esgm As Byte
                rData = Right$(rData, Len(rData) - 4)
                nombreotro = ReadField(1, rData, 44)
                esgm = ReadField(2, rData, 44)
                    If esgm > 0 Then
                        frmMenuGM.Show , frmMain
                    Else
                        frmMenu.Show , frmMain
                    End If
                Exit Sub
        Case "PART"
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, rData, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, False, False)
            Exit Sub
        Case "CEGU"
            UserCiego = True
            Dim r As RECT
            'BackBufferSurface.BltColorFill r, 0
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub
        Case "DTLC"
            rData = Right(rData, Len(rData) - 4)
            Call frmMenuGral.ParseGuildInfo(rData)
        Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            rData = Right$(rData, Len(rData) - 4)
            Call InitCartel(ReadField(1, rData, 176), CInt(ReadField(2, rData, 176)))
            Exit Sub
        Case "NPCR"
            rData = Right(rData, Len(rData) - 4)
            
            NPCInvDim = 0
            resetNPCInventory
        Exit Sub
        
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            rData = Right(rData, Len(rData) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, rData, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, rData, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, rData, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, rData, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, rData, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(6, rData, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, rData, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, rData, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, rData, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, rData, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, rData, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, rData, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, rData, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, rData, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, rData, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, rData, 44)
            NPCInventory(NPCInvDim).itemSlot = ReadField(16, rData, 44)
        Exit Sub
        Case "NPC|"              ' >>>>> Recibe slot Inventario de un NPC :: NPC|
            rData = Right(rData, Len(rData) - 4)
            NPCInvDim = ReadField(1, rData, 44)
            NPCInventory(NPCInvDim).Name = ReadField(2, rData, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(3, rData, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(4, rData, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(5, rData, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(6, rData, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(7, rData, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(8, rData, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(9, rData, 44)
            NPCInventory(NPCInvDim).Def = ReadField(10, rData, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(11, rData, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(12, rData, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(13, rData, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(14, rData, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(15, rData, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(16, rData, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(17, rData, 44)
            NPCInventory(NPCInvDim).itemSlot = ReadField(18, rData, 44)
        Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            rData = Right$(rData, Len(rData) - 4)
            UserMaxAGU = Val(ReadField(1, rData, 44))
            UserMinAGU = Val(ReadField(2, rData, 44))
            UserMaxHAM = Val(ReadField(3, rData, 44))
            UserMinHAM = Val(ReadField(4, rData, 44))
            frmMain.AguaSP.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 118)
            frmMain.COMIDASp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 118)
            frmMain.AGUABAR.Caption = UserMinAGU & "%"
            frmMain.COMIDABAR.Caption = UserMinHAM & "%"
            Exit Sub
        Case "KIGF" ' >>>>>> Mini Estadisticas :: MEST
            rData = Right$(rData, Len(rData) - 4)
            If Not frmMenuGral.Visible Then frmMenuGral.Show , frmMain
            frmMenuGral.ParseEstadisticas (rData)
        Exit Sub
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            rData = Right$(rData, Len(rData) - 4)
            SkillPoints = SkillPoints + Val(rData)
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            rData = Right$(rData, Len(rData) - 4)
            AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & rData, 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            rData = Right$(rData, Len(rData) - 4)
            frmMSG.List1.AddItem rData
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show , frmMain
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            rData = Right$(rData, Len(rData) - 4)
            frmForo.List.AddItem ReadField(1, rData, 176)
            frmForo.text(frmForo.List.ListCount - 1).text = ReadField(2, rData, 176)
            Load frmForo.text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
    End Select

    Select Case left$(sData, 5)
        Case UCase$(Chr$(110)) & mid$("MEDOK", 4, 1) & Right$("akV", 1) & "E" & Trim$(left$("  RS", 3))
            rData = Right$(rData, Len(rData) - 5)
            charindex = Val(ReadField(1, rData, 44))
            charlist(charindex).invisible = (Val(ReadField(2, rData, 44)) = 1)

            Exit Sub
        Case "KHEKD"
        rData = Right$(rData, Len(rData) - 5)
        
            RetiraObj = ReadField(1, rData, Asc(","))
            RetiraOro = ReadField(2, rData, Asc(","))
        
        Exit Sub
        Case "ZMOTD"
            rData = Right$(rData, Len(rData) - 5)
            frmCambiaMotd.Show , frmMain
            frmCambiaMotd.txtMotd.text = rData
            Exit Sub
       Case "INIAC"
            rData = Right$(rData, Len(rData) - 5)
            
            CantidadDePersonajes = ReadField(1, rData, 44)
            frmAccount.lblNoticias.Caption = ReadField(2, rData, 44)
            frmAccount.lblNoticias.ForeColor = RGB(127, 115, 101)
            
            If CantidadDePersonajes < 10 Then
                For i = CantidadDePersonajes To 9
                    CargarPJ(i).Existe = False
                Next i
            Else
                frmAccount.imgCrearPersonaje.Visible = False
                frmAccount.imgCrearPersonaje.Enabled = False
            End If
            
            Call mostrarCuenta
            Unload frmConnect
        Exit Sub
        Case "STOPD"
            rData = Right$(rData, Len(rData) - 5)
            Stopped = ReadField(1, rData, 44)
        Exit Sub
        Case "CODEH"
            rData = Right$(rData, Len(rData) - 5)
            CodigoRecibido = ReadField(1, rData, 44)
        Exit Sub
        Case "ADDPJ"
            rData = Right$(rData, Len(rData) - 5)
           
            rcvName = ReadField(1, rData, 44)
            rcvIndex = ReadField(2, rData, 44)
            rcvHead = ReadField(3, rData, 44)
            rcvBody = ReadField(4, rData, 44)
            rcvWeapon = ReadField(5, rData, 44)
            rcvShield = ReadField(6, rData, 44)
            rcvCasco = ReadField(7, rData, 44)
            rcvLevel = ReadField(8, rData, 44)
            rcvClase = ReadField(9, rData, 44)
            rcvMuerto = ReadField(10, rData, 44)
            rcvRaza = ReadField(11, rData, 44)
                      
            CargarPJ(rcvIndex - 1).Nombre = rcvName
            CargarPJ(rcvIndex - 1).Body = rcvBody
            CargarPJ(rcvIndex - 1).Head = rcvHead
            CargarPJ(rcvIndex - 1).Casco = rcvCasco
            CargarPJ(rcvIndex - 1).Shield = rcvShield
            CargarPJ(rcvIndex - 1).Weapon = rcvWeapon
            CargarPJ(rcvIndex - 1).Level = rcvLevel
            CargarPJ(rcvIndex - 1).Existe = True
            CargarPJ(rcvIndex - 1).Clase = rcvClase
            CargarPJ(rcvIndex - 1).Raza = rcvRaza
            CargarPJ(rcvIndex - 1).Muerto = rcvMuerto
        Exit Sub
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
    End Select

    Select Case left(sData, 6)
    
        Case "FLECHI" 'flecha a char
         rData = Right$(rData, Len(rData) - 6)
            engine.Crear_Flecha Val(ReadField(1, rData, 44)), Val(ReadField(2, rData, 44)), Val(ReadField(3, rData, 44)), 0, Val(ReadField(4, rData, 44))
        Exit Sub
    
    'Acá abre la ventana al primer usuario (el que va a enviar el mensaje)
      Case "ENCHAT"
        rData = Right$(rData, Len(rData) - 6)
        
        For i = 1 To 5
            If UCase$(NickContacto(i)) = UCase$(rData) Then
                Mensaje.Escribir "Ya tienes una ventana de chat abierta con este usuario."
             Exit Sub
            End If
        
            If ChatEnUso(i) = False Then
                NickContacto(i) = UCase$(rData)
                ChatEnUso(i) = True
                VentanitaMostrar(i) = 2
                
                ChatForm(i).Caption = rData
                ChatForm(i).lblName = rData
                ChatForm(i).Show , frmMain
                Exit Sub
            End If
        Next i
        
      Exit Sub
      
      'Acá la ventana al segundo usuario (el que recibe)
      Case "LDCHAT"
        rData = Right$(rData, Len(rData) - 6)
        Dim Remitente As String, Mensajitox As String
        Remitente = ReadField(1, rData, 44)
        Mensajitox = ReadField(2, rData, 44)
        
        For i = 1 To 5
            If UCase$(NickContacto(i)) = UCase$(Remitente) Then
                    AddtoRichTextBox ChatForm(i).rtbChat, "" & Remitente & " dice: " & Mensajitox & "", 255, 0, 0, True
                    RecibioMensaje(i) = True
                Exit Sub
            End If
        
            If ChatEnUso(i) = False Then
                NickContacto(i) = UCase$(Remitente)
                ChatEnUso(i) = True
                RecibioMensaje(i) = True
                
                ChatForm(i).Caption = Remitente
                ChatForm(i).lblName = Remitente
                AddtoRichTextBox ChatForm(i).rtbChat, "" & Remitente & " dice: " & Mensajitox & "", 255, 0, 0, True
                Exit Sub
            End If
        Next i
        
      Exit Sub
    
          Case "CIRUJA"
            rData = Right$(rData, Len(rData) - 6)
            Dim Raza As String, Genero As String
            Raza = ReadField(1, rData, 44)
            Genero = ReadField(2, rData, 44)
            FrmCirujia.Show , frmMain
            Call FrmCirujia.ParseHead(Raza, Genero)
        Exit Sub
        Case "AXELPT"
            frmMenuMascota.Show , frmMain
        Exit Sub
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            rData = Right$(rData, Len(rData) - 6)
            For i = 1 To NUMSKILLS
                UserSkills(i) = Val(ReadField(i, rData, 44))
            Next i
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            rData = Right(rData, Len(rData) - 6)
            For i = 1 To Val(ReadField(1, rData, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, rData, 44)
            Next i
            frmEntrenador.Show , frmMain
            Exit Sub
    End Select
    
    Select Case left$(sData, 9)
        Case "INITCBANK"           ' >>>>> Inicia cuenta bancaria.
            rData = Right$(rData, Len(rData) - 9)
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmNuevoBancoObj.List1(1).AddItem "" & Inventario.ItemName(i) & " - " & Inventario.Amount(i) & ""
                Else
                        frmNuevoBancoObj.List1(1).AddItem "Nada"
               End If
                i = i + 1
            Loop
            
            
            i = 1
            Do While i <= UBound(UserBancoInventoryB)
                If UserBancoInventoryB(i).OBJIndex <> 0 Then
                        frmNuevoBancoObj.List1(0).AddItem "" & UserBancoInventoryB(i).Name & " - " & UserBancoInventoryB(i).Amount & ""
                Else
                        frmNuevoBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            
            frmNuevoBancoObj.OroBove.text = PonerPuntos(UserBancoOro)
            frmNuevoBancoObj.MiOro.text = PonerPuntos(UserBancoOroPropio)
            
            Dim tmpPuedeObj, tmpPuedeOro As Byte
            tmpPuedeObj = ReadField(1, rData, Asc(","))
            tmpPuedeOro = ReadField(2, rData, Asc(","))
            
            If tmpPuedeObj = 1 Then
                frmNuevoBancoObj.Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\BovedaClan_RetirarObjeto_Si.jpg")
            Else
                frmNuevoBancoObj.Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\BovedaClan_RetirarObjeto_No.jpg")
            End If
            
            If tmpPuedeOro = 1 Then
                frmNuevoBancoObj.Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\BovedaClan_RetirarObjeto_Si.jpg")
            Else
                frmNuevoBancoObj.Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\BovedaClan_RetirarObjeto_No.jpg")
            End If
            
            Comerciando = True
            frmNuevoBancoObj.Show , frmMain
        Exit Sub
    End Select
    
    Select Case left$(sData, 7)
          Case "RESPUES"         ' >>> Sistema Consultas - Fishar.-
            rData = Right(rData, Len(rData) - 7)
            TieneParaResponder = True
            frmMensaje.msg.text = ReadField(1, rData, Asc("*")) & vbCrLf & "Respondido por: " & ReadField(2, rData, Asc("*"))
        Case "NEWDENU"
            rData = Right(rData, Len(rData) - 7)
            DenunciasNumber = DenunciasNumber + 1
            Denuncias(DenunciasNumber).Autor = ReadField(1, rData, Asc(","))
            Denuncias(DenunciasNumber).Contenido = ReadField(2, rData, Asc(","))
            Denuncias(DenunciasNumber).ID = ReadField(3, rData, Asc(","))
            Denuncias(DenunciasNumber).YP = ReadField(4, rData, Asc(","))
            Denuncias(DenunciasNumber).Nick = ReadField(5, rData, Asc(","))
            Denuncias(DenunciasNumber).UltimoLogeo = ReadField(6, rData, Asc(","))
            Denuncias(DenunciasNumber).UltimaDenuncia = ReadField(7, rData, Asc(","))
            Denuncias(DenunciasNumber).PrimerDenuncia = ReadField(8, rData, Asc(","))
            Denuncias(DenunciasNumber).Estado = "NO LEIDO"
        Exit Sub
        Case "IREDAEL"
            rData = Right(rData, Len(rData) - 7)
            Call frmMenuGral.ParseLeaderInfo(rData)
            Exit Sub
        Case "IREDAEK"
            rData = Right(rData, Len(rData) - 7)
            Call frmMenuGral.ParseGuildUserInfo(rData)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
        Case "ENVFPS" 'envia fps del usuario
           Call SendData("ENVFPZ" & EnvioFPS)
        Exit Sub
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            If UserParalizado = False Then
                UserParalizado = True
                TiempoParalizado = 22
            ElseIf UserParalizado = True Then
                UserParalizado = False
                TiempoParalizado = 0
            End If
        Exit Sub
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            rData = Right(rData, Len(rData) - 7)
            Call frmUserRequest.recievePeticion(rData)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                rData = Right(rData, Len(rData) - 7)
                
                tmpIndex = 1
                frmComerciar.List1(1).Clear
                For i = 1 To MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                        slotsListaInv(tmpIndex) = i
                        tmpIndex = tmpIndex + 1
                    End If
                Next i
                
                tmpIndex = 1
                frmComerciar.List1(0).Clear
                For i = 1 To MAX_NPC_INVENTORY_SLOTS
                    If NPCInventory(i).GrhIndex > 0 Then
                         frmComerciar.List1(0).AddItem NPCInventory(i).Name
                        slotsListaNPC(tmpIndex) = i
                        tmpIndex = tmpIndex + 1
                    End If
                Next i
                
                If ReadField(2, rData, 44) = "0" Then
                    frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                tmpIndex = 1
                frmComerciar.List1(1).Clear
                For i = 1 To MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
                        slotsListaInv(tmpIndex) = i
                        tmpIndex = tmpIndex + 1
                    End If
                Next i
                
                frmBancoObj.List1(0).Clear
                i = 1
                Do While i <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventory(i).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
                    End If
                    i = i + 1
                Loop
                
                rData = Right(rData, Len(rData) - 7)
                
                If ReadField(2, rData, 44) = "0" Then
                        frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
        Case "BANCOBK"           ' Banco OK :: BANCOBK
            If frmNuevoBancoObj.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                            frmNuevoBancoObj.List1(1).AddItem "" & Inventario.ItemName(i) & " - " & Inventario.Amount(i) & ""
                    Else
                            frmNuevoBancoObj.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                
                i = 1
                Do While i <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventoryB(i).OBJIndex <> 0 Then
                            frmNuevoBancoObj.List1(0).AddItem "" & UserBancoInventoryB(i).Name & " - " & UserBancoInventoryB(i).Amount & ""
                    Else
                           frmNuevoBancoObj.List1(0).AddItem "Nada"
                   End If
                   i = i + 1
                Loop
                
                rData = Right(rData, Len(rData) - 7)
                
                frmNuevoBancoObj.OroBove.text = PonerPuntos(UserBancoOro)
                frmNuevoBancoObj.MiOro.text = PonerPuntos(UserBancoOroPropio)
                
                If ReadField(2, rData, 44) = "0" Then
                        frmNuevoBancoObj.List1(0).ListIndex = frmNuevoBancoObj.LastIndexx1
                Else
                        frmNuevoBancoObj.List1(1).ListIndex = frmNuevoBancoObj.LastIndexx2
                End If
            End If
            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
        Case "TRAVELS"
          frmViajar.Show , frmMain
        Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase$(left$(rData, 9))
        Case "DAMEQUEST"
                Call SendData("IQUEST")
                frmMenuGral.Show , frmMain
        Exit Sub
    End Select
    
    ';Call LogCustom("Unhandled data: " & Rdata)
    
End Sub

Sub SendData(ByVal sdData As String, Optional setPing As Boolean = False)

    'No enviamos nada si no estamos conectados
    If Not frmMain.Socket1.Connected Then Exit Sub

    Dim AuxCmd As String
    AuxCmd = UCase$(left$(sdData, 5))
    
    If AuxCmd = "/PING" Or setPing Then TimerPing(1) = GetTickCount(): MSEnvioPING = True
    
    
    With AodefConv
    SuperClave = .Numero2Letra(AoDefProtectDynamic, , 2, AoDefExt(90, 105, 80, 80, 121), AoDefExt(78, 111, 80, 80, 121), 1, 0)
    End With
    
    Do While InStr(1, SuperClave, " ")
    SuperClave = mid$(SuperClave, 1, InStr(1, SuperClave, " ") - 1) & mid$(SuperClave, InStr(1, SuperClave, " ") + 1)
    Loop
    s = Semilla(SuperClave)
    
    sdData = AoDefEncode(Codificar(sdData, s))
    sdData = sdData & ENDC
    
    Debug.Print sdData

    'Para evitar el spamming
    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
        Exit Sub
    ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub
    End If

    Call frmMain.Socket1.Write(sdData, Len(sdData))

End Sub

Sub Login()

    If EstadoLogin = Normal Then
        SendData ("KERD22" & Val(HDSerial))
        SendData ("OOLOGI" & PJClickeado & "," & nombrecuent & "," & CodigoRecibido), True
    ElseIf EstadoLogin = CrearNuevoPj Then
        SendData ("KERD22" & Val(HDSerial))
        Call SendData("NLOGIN" & UserName & "," & UserRaza & "," & UserSexo & "," & UserSexo & "," & UserClase & "," & UserHogar _
                & "," & nombrecuent _
                & "," & Actualea & "," & UserFaccion)
    ElseIf EstadoLogin = BorrarPj Then
        SendData ("TBRP" & PJClickeado & "," & nombrecuent & "," & CodigoRecibido)
    ElseIf EstadoLogin = LoginAccount Then
        SendData ("KERD22" & Val(HDSerial))
        SendData ("ALOGIN" & nombrecuent & "," & UserPassword & "," & VersionC)
    End If
End Sub

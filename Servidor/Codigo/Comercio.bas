Attribute VB_Name = "Comercio"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  US
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

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%          MODULO DE COMERCIO NPC-USER              %%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Function UserCompraObj(ByVal userindex As Integer, _
                       ByVal ObjIndex As Integer, _
                       ByVal npcindex As Integer, _
                       ByVal Cantidad As Integer) As Boolean

    On Error GoTo errorh

    Dim infla     As Long

    Dim Descuento As String

    Dim unidad    As Long, monto As Long

    Dim slot      As Integer

    Dim obji      As Integer
    
    Dim seAcoplo As Boolean
    
    seAcoplo = True
    UserCompraObj = False
    
    If (Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(ObjIndex).Amount <= 0) Then Exit Function
    
    obji = Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(ObjIndex).ObjIndex
    
    'es una armadura real y el tipo no es faccion?
    If ObjData(obji).Real = 1 Then
        If Npclist(npcindex).Name <> "SR" Then
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbWhite & "°" & "Lo siento, la ropa faccionaria solo es para muestra, no tengo autorización para venderla. Diríjete al sastre de tu ejército." & "°" & str$(Npclist(npcindex).Char.CharIndex))

            Exit Function

        End If
    End If
    
    If ObjData(obji).Caos = 1 Then
        If Npclist(npcindex).Name <> "SC" Then
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbWhite & "°" & "Lo siento, la ropa faccionaria solo es para muestra, no tengo autorización para venderla. Diríjete al sastre de tu ejército." & "°" & str$(Npclist(npcindex).Char.CharIndex))

            Exit Function

        End If
    End If
    
    If UserList(userindex).cComercio.cComercia = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||153")
        Exit Function
    End If
    
    '¿Ya tiene un objeto de este tipo?
    slot = 1

    Do Until UserList(userindex).Invent.Object(slot).ObjIndex = obji And UserList(userindex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
        slot = slot + 1

        If slot > MAX_INVENTORY_SLOTS Then

            Exit Do

        End If

    Loop
    
    'Sino se fija por un slot vacio
    If slot > MAX_INVENTORY_SLOTS Then
        slot = 1

        Do Until UserList(userindex).Invent.Object(slot).ObjIndex = 0
            slot = slot + 1
            
            If slot > MAX_INVENTORY_SLOTS Then
                Call SendData(SendTarget.toindex, userindex, 0, "||108")

                Exit Function

            End If

        Loop

        seAcoplo = False
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
    End If
    
    'desde aca para abajo se realiza la transaccion
    UserCompraObj = True

    'Mete el obj en el slot
    If UserList(userindex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        'Menor que MAX_INV_OBJS
        UserList(userindex).Invent.Object(slot).ObjIndex = obji
        
        If seAcoplo Then
            UserList(userindex).Invent.Object(slot).Amount = UserList(userindex).Invent.Object(slot).Amount + Cantidad
        Else
            UserList(userindex).Invent.Object(slot).Amount = Cantidad
        End If
        
        'Le sustraemos el valor en oro del obj comprado
        infla = (Npclist(npcindex).Inflacion * ObjData(obji).Valor) / 100
        Descuento = UserList(userindex).flags.Descuento

        If Descuento = 0 Then Descuento = 1 'evitamos dividir por 0!
        unidad = ((ObjData(Npclist(npcindex).Invent.Object(ObjIndex).ObjIndex).Valor + infla) / Descuento)
        monto = unidad * Cantidad
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - monto
        
        Call UpdateUserInv(False, userindex, slot)
        SendUserGLD (userindex)
        
        'tal vez suba el skill comerciar ;-)
        Call SubirSkill(userindex, Comerciar)
        
        If ObjData(obji).OBJType = eOBJType.otLlaves Then Call logVentaCasa(UserList(userindex).Name & " compro " & ObjData(obji).Name)
        Call QuitarNpcInvItem(UserList(userindex).flags.TargetNPC, CByte(ObjIndex), Cantidad)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||108")
    End If

    Exit Function

errorh:
    Call LogError("Error en USERCOMPRAOBJ. " & Err.Description)
End Function


Function NpcCompraObj(ByVal userindex As Integer, _
                 ByVal ObjIndex As Integer, _
                 ByVal Cantidad As Integer) As Boolean

    On Error GoTo errorh

    Dim slot     As Integer

    Dim obji     As Integer

    Dim npcindex As Integer

    Dim infla    As Long

    Dim monto    As Long
          
    If Cantidad < 1 Then Exit Function
    
    npcindex = UserList(userindex).flags.TargetNPC
    obji = UserList(userindex).Invent.Object(ObjIndex).ObjIndex
    
    If ObjData(obji).Newbie = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||660")
        NpcCompraObj = False
        Exit Function
    End If
    
    If Npclist(npcindex).TipoItems <> eOBJType.otCualquiera Then

        '¿Son los items con los que comercia el npc?
        If Npclist(npcindex).TipoItems <> ObjData(obji).OBJType Then
            Call SendData(SendTarget.toindex, userindex, 0, "||661")
            NpcCompraObj = False
            Exit Function
        End If
    End If
    
    If obji = iORO Then
        Call SendData(SendTarget.toindex, userindex, 0, "||661")
        NpcCompraObj = False
        Exit Function
    End If
    
    If ObjData(obji).ItemDios = 1 Or ObjData(obji).Intransferible = 1 Or ObjData(obji).OBJType = otLlaves Then
        Call SendData(SendTarget.toindex, userindex, 0, "||661")
        NpcCompraObj = False
        Exit Function
    End If
    
    '¿Ya tiene un objeto de este tipo?
    slot = 1

    Do Until (Npclist(npcindex).Invent.Object(slot).ObjIndex = obji And Npclist(npcindex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS)
        
        slot = slot + 1
        
        If slot > MAX_INVENTORY_SLOTS Then Exit Do
    Loop
    
    'Sino se fija por un slot vacio antes del slot devuelto
    If slot > MAX_INVENTORY_SLOTS Then
        slot = 1

        Do Until Npclist(npcindex).Invent.Object(slot).ObjIndex = 0
            slot = slot + 1

            If slot > MAX_INVENTORY_SLOTS Then Exit Do
        Loop

        If slot <= MAX_INVENTORY_SLOTS Then Npclist(npcindex).Invent.NroItems = Npclist(npcindex).Invent.NroItems + 1
    End If
    
    If slot <= MAX_INVENTORY_SLOTS Then 'Slot valido
        'Mete el obj en el slot
        Npclist(npcindex).Invent.Object(slot).ObjIndex = obji

        If Npclist(npcindex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            Npclist(npcindex).Invent.Object(slot).Amount = Npclist(npcindex).Invent.Object(slot).Amount + Cantidad
        Else
            Npclist(npcindex).Invent.Object(slot).Amount = MAX_INVENTORY_OBJS
        End If
    End If
    
    NpcCompraObj = True
    Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)
    'Le sumamos al user el valor en oro del obj vendido
    monto = ((ObjData(obji).Valor \ 3 + infla) * Cantidad)
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + monto

    If UserList(userindex).Stats.GLD > MAXORO Then UserList(userindex).Stats.GLD = MAXORO
    
    'tal vez suba el skill comerciar ;-)
    Call SubirSkill(userindex, Comerciar)

    Exit Function

errorh:
    Call LogError("Error en NPCCOMPRAOBJ. " & Err.Description)
End Function

Sub IniciarCOmercioNPC(ByVal userindex As Integer)
On Error GoTo Errhandler
    'Mandamos el Inventario
    Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC)
    'Atcualizamos el dinero
    Call SendUserGLD(userindex)
    'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
    UserList(userindex).flags.Comerciando = True
    SendData SendTarget.toindex, userindex, 0, "INITCOM"
Exit Sub

Errhandler:
    Dim str As String
    str = "Error en IniciarComercioNPC. UI=" & userindex
    If userindex > 0 Then
        str = str & ".Nombre: " & UserList(userindex).Name & " IP:" & UserList(userindex).ip & " comerciando con "
        If UserList(userindex).flags.TargetNPC > 0 Then
            str = str & Npclist(UserList(userindex).flags.TargetNPC).Name
        Else
            str = str & "<NPCINDEX 0>"
        End If
    Else
        str = str & "<USERINDEX 0>"
    End If
End Sub

Sub NPCVentaItem(ByVal userindex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal npcindex As Integer)
'listindex+1, cantidad
On Error GoTo Errhandler

    Dim infla As Long
    Dim val As Long
    Dim Desc As String
    
    If Cantidad < 1 Then Exit Sub
    
    If i > MAX_INVENTORY_SLOTS Then Exit Sub
    
    If Cantidad > MAX_INVENTORY_OBJS Then
        Call Ban(UserList(userindex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio " & Cantidad)
        UserList(userindex).flags.Ban = 1
        Call SendData(SendTarget.toindex, userindex, 0, "ERRHas sido baneado por el sistema anti cheats")
        Call CloseSocket(userindex)
        Exit Sub
    End If
    
    'Calculamos el valor unitario
    infla = (Npclist(npcindex).Inflacion * ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).Valor) / 100
    Desc = Descuento(userindex)
    If Desc = 0 Then Desc = 1 'evitamos dividir por 0!
    val = (ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).Valor + infla) / Desc
    
    If UserList(userindex).Stats.GLD >= (val * Cantidad) Then
        If Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(i).Amount > 0 Then
            If Cantidad > Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(i).Amount Then Cantidad = Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(i).Amount
            'Agregamos el obj que compro al inventario
            If Not UserCompraObj(userindex, CInt(i), UserList(userindex).flags.TargetNPC, Cantidad) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||662")
            End If
            'Actualizamos el oro
            Call SendUserGLD(userindex)
            'Actualizamos la ventana de comercio
            Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC, i)
            Call UpdateVentanaComercio(i, 0, userindex)
        End If
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||663")
    End If
Exit Sub

Errhandler:
    Call LogError("Error en comprar item: " & Err.Description)
End Sub

Sub NPCCompraItem(ByVal userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)
On Error GoTo Errhandler
    Dim npcindex As Integer
    
    npcindex = UserList(userindex).flags.TargetNPC
    
    'Si es una armadura faccionaria vemos que la está intentando vender al sastre
    If ObjData(UserList(userindex).Invent.Object(Item).ObjIndex).Real = 1 Then
        If Npclist(npcindex).Name <> "SR" Then
            Call SendData(SendTarget.toindex, userindex, 0, "||664")
            
            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(Item, 1, userindex)
            Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC)
            Exit Sub
        End If
    ElseIf ObjData(UserList(userindex).Invent.Object(Item).ObjIndex).Caos = 1 Then
        If Npclist(npcindex).Name <> "SC" Then
            Call SendData(SendTarget.toindex, userindex, 0, "||664")
            
            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(Item, 1, userindex)
            Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC)
            Exit Sub
        End If
    End If
    
    'NPC COMPRA UN OBJ A UN USUARIO
   
    If UserList(userindex).Invent.Object(Item).Amount > 0 And UserList(userindex).Invent.Object(Item).Equipped = 0 Then
        If Cantidad > 0 And Cantidad > UserList(userindex).Invent.Object(Item).Amount Then Cantidad = UserList(userindex).Invent.Object(Item).Amount
        'Agregamos el obj que compro al inventario
        If NpcCompraObj(userindex, CInt(Item), Cantidad) Then
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(False, userindex, Item)
            'Actualizamos el oro
            Call SendUserGLD(userindex)
            
            Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC)
            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(Item, 1, userindex)
        End If
    End If
Exit Sub

Errhandler:
    Call LogError("Error en vender item: " & Err.Description)
End Sub
Sub UpdateVentanaComercio(ByVal slot As Integer, ByVal NpcInv As Byte, ByVal userindex As Integer)
    Call SendData(SendTarget.toindex, userindex, 0, "TRANSOK" & slot & "," & NpcInv)
End Sub
Sub updateNPCInventory(ByVal slot As Integer, ByVal npcindex As Integer)

    Dim i As Long
    If slot = Npclist(npcindex).Invent.NroItems Then
        For i = slot To MAX_INVENTORY_SLOTS
            Npclist(npcindex).Invent.Object(i) = Npclist(npcindex).Invent.Object(i + 1)
        Next i
    End If

End Sub

Function Descuento(ByVal userindex As Integer) As Single
    'Calcula el descuento al comerciar
    Descuento = 1 + UserList(userindex).Stats.UserSkills(eSkill.Comerciar) / 100
    UserList(userindex).flags.Descuento = Descuento
End Function

Sub EnviarNpcInv(ByVal userindex As Integer, ByVal npcindex As Integer, Optional slot As Integer = 0)
    'Enviamos el inventario del npc con el cual el user va a comerciar...
    Dim i As Integer
    Dim infla As Long
    Dim Desc As String
    Dim val As Long
    
    Desc = Descuento(userindex)
    If Desc = 0 Then Desc = 1 'evitamos dividir por 0!
    
    If slot <> 0 Then
        If Npclist(npcindex).Invent.Object(slot).ObjIndex <> 0 Then
            infla = (Npclist(npcindex).Inflacion * ObjData(Npclist(npcindex).Invent.Object(slot).ObjIndex).Valor) / 100
                val = (ObjData(Npclist(npcindex).Invent.Object(slot).ObjIndex).Valor + infla) / Desc
                SendData SendTarget.toindex, userindex, 0, "NPC|" & slot & "," & _
                ObjData(Npclist(npcindex).Invent.Object(slot).ObjIndex).Name _
                & "," & Npclist(npcindex).Invent.Object(slot).Amount & _
                "," & val _
                & "," & ObjData(Npclist(npcindex).Invent.Object(slot).ObjIndex).GrhIndex _
                & "," & Npclist(npcindex).Invent.Object(slot).ObjIndex _
                & "," & ObjData(Npclist(npcindex).Invent.Object(slot).ObjIndex).OBJType _
                & "," & ObjData(Npclist(npcindex).Invent.Object(slot).ObjIndex).MaxHIT _
                & "," & ObjData(Npclist(npcindex).Invent.Object(slot).ObjIndex).MinHIT _
                & "," & ObjData(Npclist(npcindex).Invent.Object(slot).ObjIndex).MaxDef & "," & slot
            Exit Sub
        End If
    End If
    
    SendData SendTarget.toindex, userindex, 0, "NPCR"
    
    For i = 1 To MAX_INVENTORY_SLOTS
        If Npclist(npcindex).Invent.Object(i).ObjIndex > 0 Then
            'Calculamos el porc de inflacion del npc
            infla = (Npclist(npcindex).Inflacion * ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).Valor) / 100
            val = (ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).Valor + infla) / Desc
            SendData SendTarget.toindex, userindex, 0, "NPCI" & _
            ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).Name _
            & "," & Npclist(npcindex).Invent.Object(i).Amount & _
            "," & val _
            & "," & ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).GrhIndex _
            & "," & Npclist(npcindex).Invent.Object(i).ObjIndex _
            & "," & ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).OBJType _
            & "," & ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).MaxHIT _
            & "," & ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).MinHIT _
            & "," & ObjData(Npclist(npcindex).Invent.Object(i).ObjIndex).MaxDef & "," & i
        End If
    Next i
End Sub

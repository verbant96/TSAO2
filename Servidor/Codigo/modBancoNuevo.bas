Attribute VB_Name = "modBancoNuevo"
Sub BIniciarDeposito(ByVal userindex As Integer)
On Error GoTo Errhandler

'Hacemos un Update del inventario del usuario
Call BUpdateBanUserInv(True, userindex, 0)
'Atcualizamos el dinero
Call SendUserGLD(userindex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData SendTarget.toindex, userindex, 0, "INITCBANK" & UserList(userindex).flags.PuedeRetirarObj & "," & UserList(userindex).flags.PuedeRetirarOro

UserList(userindex).flags.Comerciando = True

Errhandler:

End Sub

Sub BSendBanObj(userindex As Integer, slot As Byte, Object As UserOBJ)


UserList(userindex).BancoInventB.Object(slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(SendTarget.toindex, userindex, 0, "SBG" & slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).OBJType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef & "," & GetVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "" & Guilds(UserList(userindex).GuildIndex).GuildName & "", "BANCO") & "," & UserList(userindex).Stats.GLD)

Else

    Call SendData(SendTarget.toindex, userindex, 0, "SBG" & slot & "," & "0" & "," & "(Nada)" & "," & "0" & "," & "0" & "," & "0" & "," & "0" & "," & "0" & "," & "0" & "," & GetVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "" & Guilds(UserList(userindex).GuildIndex).GuildName & "", "BANCO") & "," & UserList(userindex).Stats.GLD)

End If


End Sub

Sub BUpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal slot As Byte)

Dim NullObj As UserOBJ
Dim loopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).BancoInventB.Object(slot).ObjIndex > 0 Then
        Call BSendBanObj(userindex, slot, UserList(userindex).BancoInventB.Object(slot))
    Else
        Call BSendBanObj(userindex, slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For loopC = 1 To MAX_BANCOINVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(userindex).BancoInventB.Object(loopC).ObjIndex > 0 Then
            Call BSendBanObj(userindex, loopC, UserList(userindex).BancoInventB.Object(loopC))
        Else
            
            Call BSendBanObj(userindex, loopC, NullObj)
            
        End If

    Next loopC

End If

End Sub

Sub BUserRetiraItem(ByVal userindex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
On Error GoTo Errhandler


If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.GranDios Then Exit Sub

If Cantidad < 1 Then Exit Sub

Call SendUserGLD(userindex)

   
       If UserList(userindex).BancoInventB.Object(i).Amount > 0 Then
            If Cantidad > UserList(userindex).BancoInventB.Object(i).Amount Then Cantidad = UserList(userindex).BancoInventB.Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call BUserReciveObj(userindex, CInt(i), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, userindex, 0)
            'Actualizamos el banco
            Call BUpdateBanUserInv(True, userindex, 0)
            'Actualizamos la ventana de comercio
            Call BUpdateVentanaBanco(i, 0, userindex)
       End If



Errhandler:

End Sub

Sub BUserReciveObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim slot As Integer
Dim obji As Integer


If UserList(userindex).BancoInventB.Object(ObjIndex).Amount <= 0 Then Exit Sub

obji = UserList(userindex).BancoInventB.Object(ObjIndex).ObjIndex


'¿Ya tiene un objeto de este tipo?
slot = 1
Do Until UserList(userindex).Invent.Object(slot).ObjIndex = obji And _
   UserList(userindex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
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
                Exit Sub
            End If
        Loop
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If



'Mete el obj en el slot
If UserList(userindex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    'Menor que MAX_INV_OBJS
    UserList(userindex).Invent.Object(slot).ObjIndex = obji
    UserList(userindex).Invent.Object(slot).Amount = UserList(userindex).Invent.Object(slot).Amount + Cantidad
    
    Call BQuitarBancoInvItem(userindex, CByte(ObjIndex), Cantidad)
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||108")
End If


End Sub

Sub BQuitarBancoInvItem(ByVal userindex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = UserList(userindex).BancoInventB.Object(slot).ObjIndex

    'Quita un Obj

       UserList(userindex).BancoInventB.Object(slot).Amount = UserList(userindex).BancoInventB.Object(slot).Amount - Cantidad
        
        If UserList(userindex).BancoInventB.Object(slot).Amount <= 0 Then
            UserList(userindex).BancoInventB.NroItems = UserList(userindex).BancoInventB.NroItems - 1
            UserList(userindex).BancoInventB.Object(slot).ObjIndex = 0
            UserList(userindex).BancoInventB.Object(slot).Amount = 0
        End If

    
    
End Sub

Sub BUpdateVentanaBanco(ByVal slot As Integer, ByVal NpcInv As Byte, ByVal userindex As Integer)
 
 Call SendData(SendTarget.toindex, userindex, 0, "BANCOBK" & slot & "," & NpcInv)
 
End Sub

Sub BUserDepositaItem(ByVal userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo Errhandler

If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.GranDios Then Exit Sub

    If UserList(userindex).cComercio.cComercia = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||153")
        Exit Sub
    End If
   
If UserList(userindex).Invent.Object(Item).Amount > 0 And UserList(userindex).Invent.Object(Item).Equipped = 0 Then
            
            If Cantidad > 0 And Cantidad > UserList(userindex).Invent.Object(Item).Amount Then Cantidad = UserList(userindex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call BUserDejaObj(userindex, CInt(Item), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, userindex, 0)
            'Actualizamos el inventario del banco
            Call BUpdateBanUserInv(True, userindex, 0)
            'Actualizamos la ventana del banco
            
            Call BUpdateVentanaBanco(Item, 1, userindex)
            
End If

Errhandler:

End Sub

Sub BUserDejaObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim slot As Integer
Dim obji As Integer

If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Director Then Exit Sub

If UserList(userindex).cComercio.cComercia = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||153")
        Exit Sub
    End If

If Cantidad < 1 Then Exit Sub

obji = UserList(userindex).Invent.Object(ObjIndex).ObjIndex

If ObjData(UserList(userindex).Invent.Object(ObjIndex).ObjIndex).Intransferible = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||185")
Exit Sub
End If

'¿Ya tiene un objeto de este tipo?
slot = 1
Do Until UserList(userindex).BancoInventB.Object(slot).ObjIndex = obji And _
         UserList(userindex).BancoInventB.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            slot = slot + 1
        
            If slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
If slot > MAX_BANCOINVENTORY_SLOTS Then
        slot = 1
        Do Until UserList(userindex).BancoInventB.Object(slot).ObjIndex = 0
            slot = slot + 1

            If slot > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(SendTarget.toindex, userindex, 0, "||186")
                Exit Sub
                Exit Do
            End If
        Loop
        If slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(userindex).BancoInventB.NroItems = UserList(userindex).BancoInventB.NroItems + 1
        
        
End If

If slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If UserList(userindex).BancoInventB.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        UserList(userindex).BancoInventB.Object(slot).ObjIndex = obji
        UserList(userindex).BancoInventB.Object(slot).Amount = UserList(userindex).BancoInventB.Object(slot).Amount + Cantidad
        
        Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)

    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||186")
    End If

Else
    Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)
End If

End Sub


Attribute VB_Name = "modBanco"
Option Explicit
Sub IniciarDeposito(ByVal userindex As Integer)
On Error GoTo Errhandler

'Hacemos un Update del inventario del usuario
Call UpdateBanUserInv(True, userindex, 0)
'Atcualizamos el dinero
Call SendUserGLD(userindex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData SendTarget.toindex, userindex, 0, "INITBANCO"

UserList(userindex).flags.Comerciando = True

Errhandler:

End Sub
Sub updateBInventory(ByVal slot As Integer, ByVal userindex As Integer)

    Dim i As Long
        For i = slot To MAX_BANCOINVENTORY_SLOTS
            If i <= UserList(userindex).BancoInvent.NroItems Then
                UserList(userindex).BancoInvent.Object(i) = UserList(userindex).BancoInvent.Object(i + 1)
            Else
                UserList(userindex).BancoInvent.Object(i).ObjIndex = 0
                UserList(userindex).BancoInvent.Object(i).Amount = 0
            End If
        Next i

End Sub
Sub SendBanObj(userindex As Integer, slot As Byte, Object As UserOBJ)


UserList(userindex).BancoInvent.Object(slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(SendTarget.toindex, userindex, 0, "SBO" & slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).OBJType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef)

End If


End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal slot As Byte)

Dim NullObj As UserOBJ
Dim loopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).BancoInvent.Object(slot).ObjIndex > 0 Then
        Call SendBanObj(userindex, slot, UserList(userindex).BancoInvent.Object(slot))
    End If

Else

    Call SendData(SendTarget.toindex, userindex, 0, "SBR")

'Actualiza todos los slots
    For loopC = 1 To MAX_BANCOINVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(userindex).BancoInvent.Object(loopC).ObjIndex > 0 Then
            Call SendBanObj(userindex, loopC, UserList(userindex).BancoInvent.Object(loopC))
        End If

    Next loopC

End If

End Sub

Sub UserRetiraItem(ByVal userindex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
On Error GoTo Errhandler


If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Administrador Then Exit Sub

If Cantidad < 1 Then Exit Sub

Call SendUserGLD(userindex)

   
       If UserList(userindex).BancoInvent.Object(i).Amount > 0 Then
            If Cantidad > UserList(userindex).BancoInvent.Object(i).Amount Then Cantidad = UserList(userindex).BancoInvent.Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call UserReciveObj(userindex, CInt(i), Cantidad)
            'Actualizamos el banco
            Call UpdateBanUserInv(True, userindex, 0)
            'Actualizamos la ventana de comercio
            Call UpdateVentanaBanco(i, 0, userindex)
       End If



Errhandler:

End Sub

Sub UserReciveObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim slot As Integer
Dim obji As Integer


If UserList(userindex).BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub

obji = UserList(userindex).BancoInvent.Object(ObjIndex).ObjIndex


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
    
    Call UpdateUserInv(False, userindex, slot)
    Call QuitarBancoInvItem(userindex, CByte(ObjIndex), Cantidad)
    Call LogDepositos("" & UserList(userindex).Name & " retiró " & Cantidad & " - " & ObjData(obji).Name & "")
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||108")
End If


End Sub

Sub QuitarBancoInvItem(ByVal userindex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = UserList(userindex).BancoInvent.Object(slot).ObjIndex

    'Quita un Obj

       UserList(userindex).BancoInvent.Object(slot).Amount = UserList(userindex).BancoInvent.Object(slot).Amount - Cantidad
        
        If UserList(userindex).BancoInvent.Object(slot).Amount <= 0 Then
            UserList(userindex).BancoInvent.NroItems = UserList(userindex).BancoInvent.NroItems - 1
            UserList(userindex).BancoInvent.Object(slot).ObjIndex = 0
            UserList(userindex).BancoInvent.Object(slot).Amount = 0
            Call updateBInventory(slot, userindex)
        End If

    
    
End Sub

Sub UpdateVentanaBanco(ByVal slot As Integer, ByVal NpcInv As Byte, ByVal userindex As Integer)
 
 
 Call SendData(SendTarget.toindex, userindex, 0, "BANCOOK" & slot & "," & NpcInv)
 
End Sub

Sub UserDepositaItem(ByVal userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo Errhandler

If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Administrador Then Exit Sub

    If UserList(userindex).cComercio.cComercia = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||153")
        Exit Sub
    End If
   
If UserList(userindex).Invent.Object(Item).Amount > 0 And UserList(userindex).Invent.Object(Item).Equipped = 0 Then
            
            If Cantidad > 0 And Cantidad > UserList(userindex).Invent.Object(Item).Amount Then Cantidad = UserList(userindex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call UserDejaObj(userindex, CInt(Item), Cantidad)
            'Actualizamos el inventario del banco
            Call UpdateBanUserInv(True, userindex, 0)
            'Actualizamos la ventana del banco
            
            Call UpdateVentanaBanco(Item, 1, userindex)
End If

Errhandler:

End Sub

Sub UserDejaObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim slot As Integer
Dim obji As Integer

If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Director Then
Call SendData(SendTarget.toindex, userindex, 0, "||185")
Exit Sub
End If

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
Do Until UserList(userindex).BancoInvent.Object(slot).ObjIndex = obji And _
         UserList(userindex).BancoInvent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            slot = slot + 1
        
            If slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
If slot > MAX_BANCOINVENTORY_SLOTS Then
        slot = 1
        Do Until UserList(userindex).BancoInvent.Object(slot).ObjIndex = 0
            slot = slot + 1

            If slot > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(SendTarget.toindex, userindex, 0, "||186")
                Exit Sub
                Exit Do
            End If
        Loop
        If slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(userindex).BancoInvent.NroItems = UserList(userindex).BancoInvent.NroItems + 1
        
        
End If

If slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If UserList(userindex).BancoInvent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        UserList(userindex).BancoInvent.Object(slot).ObjIndex = obji
        UserList(userindex).BancoInvent.Object(slot).Amount = UserList(userindex).BancoInvent.Object(slot).Amount + Cantidad
        
        If ObjData(obji).OBJType = eOBJType.otMontura And UserList(userindex).flags.Montando = 1 Then Call Desmontar(userindex)
        
        Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||186")
    End If

Else
    Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)
End If

Call LogDepositos("" & UserList(userindex).Name & " depositó " & Cantidad & " - " & ObjData(obji).Name & "")

End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(SendTarget.toindex, sendIndex, 0, "N|" & UserList(userindex).Name & "~69~190~156")
Call SendData(SendTarget.toindex, sendIndex, 0, "||187@" & UserList(userindex).BancoInvent.NroItems)
For j = 1 To MAX_BANCOINVENTORY_SLOTS
    If UserList(userindex).BancoInvent.Object(j).ObjIndex > 0 Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "||188@" & j & "@" & ObjData(UserList(userindex).BancoInvent.Object(j).ObjIndex).Name & "@" & UserList(userindex).BancoInvent.Object(j).Amount)
    End If
Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|" & CharName & "~69~190~156")
    Call SendData(SendTarget.toindex, sendIndex, 0, "||187@" & GetVar(CharFile, "BancoInventory", "CantidadItems"))
    For j = 1 To MAX_BANCOINVENTORY_SLOTS
        Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
        ObjInd = ReadField(1, Tmp, Asc("-"))
        ObjCant = ReadField(2, Tmp, Asc("-"))
        If ObjInd > 0 Then
            Call SendData(SendTarget.toindex, sendIndex, 0, "||188@" & j & "@" & ObjData(ObjInd).Name & "@" & ObjCant)
        End If
    Next
Else
    Call SendData(SendTarget.toindex, sendIndex, 0, "||189@" & CharName)
End If

End Sub

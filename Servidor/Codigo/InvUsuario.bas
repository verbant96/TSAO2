Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal userindex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(userindex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

If ObjData(ObjIndex).ClaseProhibida(1) <> "" Then
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(ObjIndex).ClaseProhibida(i) = UCase$(UserList(userindex).clase) Then
                ClasePuedeUsarItem = False
                Exit Function
        End If
    Next i
    
Else
    
    

End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function
Sub QuitarNewbieObj(ByVal userindex As Integer)

Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(userindex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(userindex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(userindex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, userindex, j)
        
        End If
Next

'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UserList(userindex).Pos.Map = 37 Then
    
    Dim DeDonde As WorldPos
    DeDonde = Tanaris
       
    Call WarpUserChar(userindex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal userindex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(userindex).Invent.Object(j).ObjIndex = 0
        UserList(userindex).Invent.Object(j).Amount = 0
        UserList(userindex).Invent.Object(j).Equipped = 0
        
Next

UserList(userindex).Invent.NroItems = 0

UserList(userindex).Invent.ArmourEqpObjIndex = 0
UserList(userindex).Invent.ArmourEqpSlot = 0

UserList(userindex).Invent.WeaponEqpObjIndex = 0
UserList(userindex).Invent.WeaponEqpSlot = 0

UserList(userindex).Invent.CascoEqpObjIndex = 0
UserList(userindex).Invent.CascoEqpSlot = 0

UserList(userindex).Invent.EscudoEqpObjIndex = 0
UserList(userindex).Invent.EscudoEqpSlot = 0

UserList(userindex).Invent.HerramientaEqpObjIndex = 0
UserList(userindex).Invent.HerramientaEqpSlot = 0

UserList(userindex).Invent.MunicionEqpObjIndex = 0
UserList(userindex).Invent.MunicionEqpSlot = 0

UserList(userindex).Invent.BarcoObjIndex = 0
UserList(userindex).Invent.BarcoSlot = 0

Call UpdateUserInv(True, userindex, 0, True)

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal userindex As Integer)
On Error GoTo Errhandler

If Cantidad > 100000 Then Exit Sub

'SI EL NPC TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(userindex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon
        If Cantidad > 39999 Then
            Dim j As Integer
            Dim k As Integer
            Dim M As Integer
            Dim Cercanos As String
            M = UserList(userindex).Pos.Map
            For j = UserList(userindex).Pos.X - 10 To UserList(userindex).Pos.X + 10
                For k = UserList(userindex).Pos.Y - 10 To UserList(userindex).Pos.Y + 10
                    If InMapBounds(M, j, k) Then
                        If MapData(M, j, k).userindex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(M, j, k).userindex).Name & ","
                        End If
                    End If
                Next k
            Next j
            Call LogDesarrollo(UserList(userindex).Name & " tira oro. Cercanos: " & Cercanos)
        End If
        '/Seguridad
        
        Do While (Cantidad > 0) And (UserList(userindex).Stats.GLD > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(userindex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(userindex).Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, False)
            
            Call TirarItemAlPiso(UserList(userindex).Pos, MiObj)
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
    
End If

Exit Sub

Errhandler:

End Sub

Sub QuitarUserInvItem(ByVal userindex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)

Dim MiObj As obj
'Desequipa
If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(slot).Equipped = 1 Then Call Desequipar(userindex, slot)

'Quita un objeto
UserList(userindex).Invent.Object(slot).Amount = UserList(userindex).Invent.Object(slot).Amount - Cantidad
'¿Quedan mas?
If UserList(userindex).Invent.Object(slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(slot).ObjIndex = 0
    UserList(userindex).Invent.Object(slot).Amount = 0
End If
    
    
    Call UpdateUserInv(False, userindex, slot)
End Sub
Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal slot As Byte, Optional mandoVacios As Boolean = False)

Dim NullObj As UserOBJ
Dim loopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).Invent.Object(slot).ObjIndex > 0 Then
        Call ChangeUserInv(userindex, slot, UserList(userindex).Invent.Object(slot))
    Else
        Call ChangeUserInv(userindex, slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For loopC = 1 To MAX_INVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(userindex).Invent.Object(loopC).ObjIndex > 0 Then
            Call ChangeUserInv(userindex, loopC, UserList(userindex).Invent.Object(loopC))
        Else
            If (mandoVacios) Then Call ChangeUserInv(userindex, loopC, NullObj)
        End If

    Next loopC

End If

End Sub
Sub DropObj(ByVal userindex As Integer, ByVal slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim obj As obj

If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Director Then Exit Sub

If ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).Intransferible = 1 And UserList(userindex).flags.Privilegios <= PlayerType.Consejero Then
    Call SendData(SendTarget.toindex, userindex, 0, "||152")
 Exit Sub
End If

If UserList(userindex).cComercio.cComercia = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||153")
    Exit Sub
End If

If num > 0 Then
  
  If num > UserList(userindex).Invent.Object(slot).Amount Then num = UserList(userindex).Invent.Object(slot).Amount
  
  'Check objeto en el suelo
  If MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex = 0 Or MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex Then
        If UserList(userindex).Invent.Object(slot).Equipped = 1 Then Call Desequipar(userindex, slot)
        obj.ObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex
        
        If num + MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.Amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.Amount
        End If
        
        obj.Amount = num
        
        If ObjData(obj.ObjIndex).OBJType = eOBJType.otMontura And UserList(userindex).flags.Montando = 1 Then Call Desmontar(userindex)
        
        Call MakeObj(SendTarget.toMap, 0, Map, obj, Map, X, Y)
        Call QuitarUserInvItem(userindex, slot, num)
        Call UpdateUserInv(False, userindex, slot)
        
        Call LogTirarItems("" & UserList(userindex).Name & " Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name & "")
        
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(userindex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name, False)
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGMss(UserList(userindex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name, False)
  Else
    Call SendData(SendTarget.toindex, userindex, 0, "||107")
  End If
    
End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)

MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - num

If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
    MapData(Map, X, Y).OBJInfo.ObjIndex = 0
    MapData(Map, X, Y).OBJInfo.Amount = 0
    
    If sndRoute = SendTarget.toMap Then
        Call SendToAreaByPos(Map, X, Y, "BO" & X & "," & Y)
   Else
        Call SendData(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)
    End If
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, obj As obj, Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then

    If MapData(Map, X, Y).OBJInfo.ObjIndex = obj.ObjIndex Then
        MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount + obj.Amount
    Else
        MapData(Map, X, Y).OBJInfo = obj
        
        If sndRoute = SendTarget.toMap Then
            'If (Obj.ObjIndex = 378) Then
            '    Call ModAreas.SendToAreaByPos(Map, X, Y, "PCF" & 101 & "," & X & "," & Y & "-1")
            'Else
                Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(obj.ObjIndex).GrhIndex & "," & X & "," & Y)
            'End If
        Else
           ' If (Obj.ObjIndex = 378) Then
            '    Call SendData(sndRoute, sndIndex, sndMap, "PCF" & 101 & "," & X & "," & Y & "-1")
            'Else
                Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(obj.ObjIndex).GrhIndex & "," & X & "," & Y)
            'End If
        End If
        
        If MapData(Map, X, Y).Blocked = 0 Then
            If ObjData(obj.ObjIndex).AntiLimpieza = 1 Then Exit Sub
            If ItemNoEsDeMapa(obj.ObjIndex) Then CleanWorld_AddItem Map, X, Y, 10, obj.ObjIndex
        End If
        
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal userindex As Integer, ByRef MiObj As obj) As Boolean
On Error GoTo Errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
slot = 1
Do Until UserList(userindex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex And _
         UserList(userindex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   slot = slot + 1
   If slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If slot > MAX_INVENTORY_SLOTS Then
   slot = 1
   Do Until UserList(userindex).Invent.Object(slot).ObjIndex = 0
       slot = slot + 1
       If slot > MAX_INVENTORY_SLOTS Then
           Call SendData(SendTarget.toindex, userindex, 0, "||108")
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(userindex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(userindex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex
   UserList(userindex).Invent.Object(slot).Amount = UserList(userindex).Invent.Object(slot).Amount + MiObj.Amount
Else
   UserList(userindex).Invent.Object(slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, userindex, slot)


Exit Function
Errhandler:

End Function


Sub GetObj(ByVal userindex As Integer)

Dim obj As ObjData
Dim MiObj As obj

'¿Hay algun obj?
If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).OBJInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim slot As Byte
        
        X = UserList(userindex).Pos.X
        Y = UserList(userindex).Pos.Y
        obj = ObjData(MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex
        
If (UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1) And UserList(userindex).flags.Privilegios <= PlayerType.Consejero Then
    Call SendData(SendTarget.toindex, userindex, 0, "||109")
Exit Sub
End If
        
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||108")
        Else
            'Quitamos el objeto
            Call EraseObj(SendTarget.toMap, 0, UserList(userindex).Pos.Map, MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.Amount, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
            If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(userindex).Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, False)
            If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGMss(UserList(userindex).Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, False)
            Call LogAgarrarItems("" & UserList(userindex).Name & " Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name & "")
        End If
        
    End If
End If

End Sub

Sub Desequipar(ByVal userindex As Integer, ByVal slot As Byte)
'Desequipa el item slot del inventario
Dim obj As ObjData


If (slot < LBound(UserList(userindex).Invent.Object)) Or (slot > UBound(UserList(userindex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(userindex).Invent.Object(slot).ObjIndex = 0 Then
    Exit Sub
End If

obj = ObjData(UserList(userindex).Invent.Object(slot).ObjIndex)

Select Case obj.OBJType
    Case eOBJType.otWeapon
        UserList(userindex).Invent.Object(slot).Equipped = 0
        UserList(userindex).Invent.WeaponEqpObjIndex = 0
        UserList(userindex).Invent.WeaponEqpSlot = 0
        If Not UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).Char.WeaponAnim = NingunArma
            Call ChangeUserArma(SendTarget.toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.WeaponAnim)
        End If
        
        If obj.Aura = UserList(userindex).Char.AuraW Then
            UserList(userindex).Char.AuraW = 0
            SendUserAura (userindex)
        End If
        
    Case eOBJType.otFlechas
        UserList(userindex).Invent.Object(slot).Equipped = 0
        UserList(userindex).Invent.MunicionEqpObjIndex = 0
        UserList(userindex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otHerramientas
        UserList(userindex).Invent.Object(slot).Equipped = 0
        UserList(userindex).Invent.HerramientaEqpObjIndex = 0
        UserList(userindex).Invent.HerramientaEqpSlot = 0
        
        If obj.Aura = UserList(userindex).Char.AuraR Then
            UserList(userindex).Char.AuraR = 0
            SendUserAura (userindex)
        End If
    
    Case eOBJType.otArmadura
        UserList(userindex).Invent.Object(slot).Equipped = 0
        UserList(userindex).Invent.ArmourEqpObjIndex = 0
        UserList(userindex).Invent.ArmourEqpSlot = 0
        Call DarCuerpoDesnudo(userindex, UserList(userindex).flags.Mimetizado = 1)
        Call ChangeUserBody(SendTarget.toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body)
        
        If obj.Aura = UserList(userindex).Char.AuraA Then
            UserList(userindex).Char.AuraA = 0
            SendUserAura (userindex)
        End If
            
    Case eOBJType.otcASCO
        UserList(userindex).Invent.Object(slot).Equipped = 0
        UserList(userindex).Invent.CascoEqpObjIndex = 0
        UserList(userindex).Invent.CascoEqpSlot = 0
        If Not UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).Char.CascoAnim = NingunCasco
            Call ChangeUserCasco(SendTarget.toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.CascoAnim)
        End If
        
        If obj.Aura = UserList(userindex).Char.AuraC Then
            UserList(userindex).Char.AuraC = 0
            SendUserAura (userindex)
        End If
    
    Case eOBJType.otESCUDO
        UserList(userindex).Invent.Object(slot).Equipped = 0
        UserList(userindex).Invent.EscudoEqpObjIndex = 0
        UserList(userindex).Invent.EscudoEqpSlot = 0
        If Not UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).Char.ShieldAnim = NingunEscudo
            Call ChangeUserEscudo(SendTarget.toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.ShieldAnim)
        End If
        
        If obj.Aura = UserList(userindex).Char.AuraE Then
            UserList(userindex).Char.AuraE = 0
            SendUserAura (userindex)
        End If
        
End Select


Call SendUserHitBox(userindex)
Call ActualizarSlotEquipped(userindex, slot)

End Sub

Function SexoPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo Errhandler

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(userindex).Genero) <> "HOMBRE"
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(userindex).Genero) <> "MUJER"
Else
    SexoPuedeUsarItem = True
End If

Exit Function
Errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Real = 1 Then
    If Not Criminal(userindex) Then
        FaccionPuedeUsarItem = (UserList(userindex).Faccion.ArmadaReal = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If Criminal(userindex) Then
        FaccionPuedeUsarItem = (UserList(userindex).Faccion.FuerzasCaos = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If

End Function
Public Function tengoSkin(ByVal userindex As Integer, ByVal numObj As Integer) As Byte
    
    Dim i As Long, tmpB As Byte
        
    tmpB = 0
    With UserList(userindex)
        For i = 1 To .cantSkins
            If .Skin(i).numObj = numObj Then
                tmpB = i
                Exit For
            End If
        Next i
    End With
    
    tengoSkin = tmpB
    
End Function
Sub EquiparInvItem(ByVal userindex As Integer, ByVal slot As Byte)
On Error GoTo Errhandler

'Equipa un item del inventario
Dim obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex
obj = ObjData(ObjIndex)

If obj.ItemDios = 1 Then
    If MapaEspecial(userindex) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||110")
        Exit Sub
    End If
End If

If obj.Newbie = 1 And Not EsNewbie(userindex) Then
     Call SendData(SendTarget.toindex, userindex, 0, "||111")
     Exit Sub
End If

If UserList(userindex).flags.Montando = 1 Then Exit Sub
If UserList(userindex).flags.Transformado = 1 Then Exit Sub
       
If obj.lvl > 0 Then
    If (UserList(userindex).Stats.ELV < obj.lvl) And (Not UserList(userindex).flags.EnJDH) Then
        If UserList(userindex).flags.Privilegios = PlayerType.User Then
               Call SendData(SendTarget.toindex, userindex, 0, "||112@" & obj.lvl)
        Exit Sub
        End If
    End If
End If

Select Case obj.OBJType
    Case eOBJType.otWeapon
       If ClasePuedeUsarItem(userindex, ObjIndex) And _
          FaccionPuedeUsarItem(userindex, ObjIndex) Or UserList(userindex).flags.Privilegios >= PlayerType.Semidios Then
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(slot).Equipped Then
                    'Animacion por defecto
                    If UserList(userindex).flags.Mimetizado = 1 Then
                        UserList(userindex).CharMimetizado.WeaponAnim = NingunArma
                    Else
                        UserList(userindex).Char.WeaponAnim = NingunArma
                    End If
                    
                    'Quitamos del inv el item
                    Call Desequipar(userindex, slot)
                    Exit Sub
                End If
                
                If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
                    If ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).proyectil = 1 And ((ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).ItemDios = 0) Or (ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).ItemDios = 1 And ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ItemDios = 0)) Then
                        Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
                    End If
                End If
                
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
                End If
        
                UserList(userindex).Invent.Object(slot).Equipped = 1
                UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex
                UserList(userindex).Invent.WeaponEqpSlot = slot
                
                'dosmanos
                If ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).DosManos = 1 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
                End If
                
                'Sonido
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_SACARARMA)
                
                If UserList(userindex).flags.Mimetizado = 1 Then
                    UserList(userindex).CharMimetizado.WeaponAnim = obj.WeaponAnim
                Else
                    UserList(userindex).Char.WeaponAnim = obj.WeaponAnim
                    Call ChangeUserArma(SendTarget.toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.WeaponAnim)
                End If
                
            If obj.Aura > 0 Then
                UserList(userindex).Char.AuraW = obj.Aura
                SendUserAura (userindex)
            End If
                
       Else
            Call SendData(SendTarget.toindex, userindex, 0, "||113")
       End If
    
    Case eOBJType.otHerramientas
       If ClasePuedeUsarItem(userindex, ObjIndex) And _
          FaccionPuedeUsarItem(userindex, ObjIndex) Or UserList(userindex).flags.Privilegios >= PlayerType.Semidios Then
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.HerramientaEqpSlot)
                End If
        
                UserList(userindex).Invent.Object(slot).Equipped = 1
                UserList(userindex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(userindex).Invent.HerramientaEqpSlot = slot
                
            If obj.Aura > 0 Then
                UserList(userindex).Char.AuraR = obj.Aura
                SendUserAura (userindex)
            End If
                
       Else
            Call SendData(SendTarget.toindex, userindex, 0, "||113")
       End If
    
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(slot).ObjIndex) And _
          FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(slot).ObjIndex) Or UserList(userindex).flags.Privilegios >= PlayerType.Semidios Then
                
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
                End If
        
                UserList(userindex).Invent.Object(slot).Equipped = 1
                UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex
                UserList(userindex).Invent.MunicionEqpSlot = slot
                
       Else
            Call SendData(SendTarget.toindex, userindex, 0, "||113")
       End If
    
    Case eOBJType.otArmadura
        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(slot).ObjIndex) And _
           SexoPuedeUsarItem(userindex, UserList(userindex).Invent.Object(slot).ObjIndex) And _
           CheckRazaUsaRopa(userindex, UserList(userindex).Invent.Object(slot).ObjIndex) And _
           FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(slot).ObjIndex) Or UserList(userindex).flags.Privilegios >= PlayerType.Semidios Then
           
           'Si esta equipado lo quita
            If UserList(userindex).Invent.Object(slot).Equipped Then
                Call Desequipar(userindex, slot)
                Call DarCuerpoDesnudo(userindex, UserList(userindex).flags.Mimetizado = 1)
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
            End If
    
            'Lo equipa
            UserList(userindex).Invent.Object(slot).Equipped = 1
            UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex
            UserList(userindex).Invent.ArmourEqpSlot = slot
                
            If UserList(userindex).flags.Mimetizado = 1 Then
                UserList(userindex).CharMimetizado.Body = equiparRopaje(userindex)
            Else
                UserList(userindex).Char.Body = equiparRopaje(userindex)
                Call ChangeUserBody(SendTarget.toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body)
            End If
            
            UserList(userindex).flags.Desnudo = 0
            
            If obj.Aura > 0 Then
                UserList(userindex).Char.AuraA = obj.Aura
                SendUserAura (userindex)
            End If
            

        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||113")
        End If
    
    Case eOBJType.otcASCO
        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
        If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(slot).ObjIndex) Or UserList(userindex).flags.Privilegios >= PlayerType.Semidios Then
            'Si esta equipado lo quita
            If UserList(userindex).Invent.Object(slot).Equipped Then
                If UserList(userindex).flags.Mimetizado = 1 Then
                    UserList(userindex).CharMimetizado.CascoAnim = NingunCasco
                Else
                    UserList(userindex).Char.CascoAnim = NingunCasco
                End If
                
                Call Desequipar(userindex, slot)
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
            End If
    
            'Lo equipa
            
            UserList(userindex).Invent.Object(slot).Equipped = 1
            UserList(userindex).Invent.CascoEqpObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex
            UserList(userindex).Invent.CascoEqpSlot = slot
            
            If UserList(userindex).flags.Mimetizado = 1 Then
                UserList(userindex).CharMimetizado.CascoAnim = obj.CascoAnim
            Else
                UserList(userindex).Char.CascoAnim = obj.CascoAnim
                Call ChangeUserCasco(SendTarget.toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.CascoAnim)
            End If
            
            If obj.Aura > 0 Then
                UserList(userindex).Char.AuraC = obj.Aura
                SendUserAura (userindex)
            End If
            
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||113")
        End If
    
    Case eOBJType.otESCUDO
        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
         If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(slot).ObjIndex) And _
             FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(slot).ObjIndex) Or UserList(userindex).flags.Privilegios >= PlayerType.Semidios Then

             'Si esta equipado lo quita
             If UserList(userindex).Invent.Object(slot).Equipped Then
                If UserList(userindex).flags.Mimetizado = 1 Then
                     UserList(userindex).CharMimetizado.ShieldAnim = NingunEscudo
                 Else
                     UserList(userindex).Char.ShieldAnim = NingunEscudo
                 End If
                 
                 Call Desequipar(userindex, slot)
                 Exit Sub
             End If
             
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil = 1 And ((ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).ItemDios = 0) Or (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).ItemDios = 1 And ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).ItemDios = 0)) Then
                    Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
                End If
            End If
     
             'Quita el anterior
             If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
             End If
     
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DosManos = 1 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
                End If
            End If
                    
             'Lo equipa
             
             UserList(userindex).Invent.Object(slot).Equipped = 1
             UserList(userindex).Invent.EscudoEqpObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex
             UserList(userindex).Invent.EscudoEqpSlot = slot
             
             If UserList(userindex).flags.Mimetizado = 1 Then
                 UserList(userindex).CharMimetizado.ShieldAnim = obj.ShieldAnim
             Else
                 UserList(userindex).Char.ShieldAnim = obj.ShieldAnim
                 
                 Call ChangeUserEscudo(SendTarget.toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.ShieldAnim)
             End If
             
            If obj.Aura > 0 Then
                UserList(userindex).Char.AuraE = obj.Aura
                SendUserAura (userindex)
            End If
             
         Else
             Call SendData(SendTarget.toindex, userindex, 0, "||113")
         End If
End Select

'Actualiza
Call SendUserHitBox(userindex)
Call ActualizarSlotEquipped(userindex, slot)

Exit Sub
Errhandler:
Call LogError("EquiparInvItem Slot:" & slot)
End Sub

Private Function CheckRazaUsaRopa(ByVal userindex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler

'Verifica si la raza puede usar la ropa
If ObjData(ItemIndex).razaDoble = 1 Then CheckRazaUsaRopa = True: Exit Function

If UserList(userindex).Raza = "Humano" Or _
   UserList(userindex).Raza = "Elfo" Or _
   UserList(userindex).Raza = "Elfo Oscuro" Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If


Exit Function
Errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal userindex As Integer, ByVal slot As Byte)

'Usa un item del inventario
Dim obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As obj

If UserList(userindex).Invent.Object(slot).Amount = 0 Then Exit Sub

obj = ObjData(UserList(userindex).Invent.Object(slot).ObjIndex)

If obj.Newbie = 1 And Not EsNewbie(userindex) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||111")
    Exit Sub
End If

ObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex
UserList(userindex).flags.TargetObjInvIndex = ObjIndex
UserList(userindex).flags.TargetObjInvSlot = slot

Select Case obj.OBJType
    Case eOBJType.otUseOnce
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||5")
            Exit Sub
        End If

        'Usa el item
        UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MinHam + obj.MinHam
        If UserList(userindex).Stats.MinHam > UserList(userindex).Stats.MaxHam Then _
            UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MaxHam
        UserList(userindex).flags.Hambre = 0
        Call EnviarHambreYsed(userindex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, slot, 1)
        
        Call UpdateUserInv(False, userindex, slot)

    Case eOBJType.otGuita
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||5")
            Exit Sub
        End If
        
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + UserList(userindex).Invent.Object(slot).Amount
        UserList(userindex).Invent.Object(slot).Amount = 0
        UserList(userindex).Invent.Object(slot).ObjIndex = 0
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
        
        Call UpdateUserInv(False, userindex, slot)
        Call SendUserGLD(userindex)
        
    Case eOBJType.otWeapon
        If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||5")
                Exit Sub
        End If
        
        If ObjData(ObjIndex).proyectil = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "T01" & Proyectiles)
        Else
            If UserList(userindex).flags.TargetObj = 0 Then Exit Sub
            
            '¿El target-objeto es leña?
            If UserList(userindex).flags.TargetObj = Leña Then
                If UserList(userindex).Invent.Object(slot).ObjIndex = DAGA Then
                    Call TratarDeHacerFogata(UserList(userindex).flags.TargetObjMap, _
                         UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY, userindex)
                End If
            End If
        End If
    
    Case eOBJType.otPociones
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||5")
            Exit Sub
        End If
        
        UserList(userindex).flags.TomoPocion = True
        UserList(userindex).flags.TipoPocion = obj.TipoPocion
                
        Select Case UserList(userindex).flags.TipoPocion
        
            Case 1 'Modif la agilidad
                UserList(userindex).flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) > 35 Then _
                    UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = 35
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, slot, 1)
                Call SendUserAgilidad(userindex)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_BEBER)
                
            Case 2 'Modif la fuerza
                UserList(userindex).flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) > 35 Then _
                    UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = 35
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, slot, 1)
                Call SendUserFuerza(userindex)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_BEBER)
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then _
                    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, slot, 1)
                Call SendUserHP(userindex)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_BEBER)
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN + Porcentaje(UserList(userindex).Stats.MaxMAN, 5)
                If UserList(userindex).Stats.MinMAN > UserList(userindex).Stats.MaxMAN Then _
                    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, slot, 1)
                Call SendUserMP(userindex)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_BEBER)
                
            Case 5 ' Pocion violeta
                If UserList(userindex).flags.Envenenado = 1 Then
                    UserList(userindex).flags.Envenenado = 0
                    Call SendData(SendTarget.toindex, userindex, 0, "||114")
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_BEBER)
                
            Case 6 ' pocion de remo
            
                If UserList(userindex).flags.Paralizado = 1 Then
        
                    If (UCase$(UserList(userindex).clase) <> "GUERRERO" And UCase$(UserList(userindex).clase) <> "CAZADOR") And (UserList(userindex).Counters.usoPotaRemo > 0) Then Call SendData(SendTarget.toindex, userindex, 0, "||940"): Exit Sub
                    If UserList(userindex).Counters.InmoManopla > 0 Then Call SendData(SendTarget.toindex, userindex, 0, "||940"): Exit Sub
                 
                    If (UserList(userindex).Stats.MinHP > 60) Then
                        UserList(userindex).flags.Paralizado = 0
                        Call SendData(SendTarget.toindex, userindex, 0, "PARADOK")
                        Call SendData(SendTarget.toindex, userindex, 0, "||939")
                        UserList(userindex).Counters.usoPotaRemo = 3
                        UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - 60
                        SendUserHP userindex
                        Call QuitarUserInvItem(userindex, slot, 1)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||941")
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||940")
                End If
                
       End Select
       
       Call ActualizarSlot(userindex, slot)
       
    Case eOBJType.otCofreAzar
            Dim NumCofre As Integer
            Dim NumeritoX As Long
            
            If obj.cofreLlave > 0 Then
                If Not TieneObjetos(obj.cofreLlave, 1, userindex) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||958@" & ObjData(obj.cofreLlave).Name)
                    Exit Sub
                Else
                    Call QuitarObjetos(obj.cofreLlave, 1, userindex)
                End If
            End If
            
            Call QuitarObjetos(UserList(userindex).Invent.Object(slot).ObjIndex, 1, userindex)
            
            If (CofresAzar(obj.TipoCofre).Random = 1) Then
                'Recorre los objetos y va descartando
                NumCofre = RandomNumber(1, CofresAzar(obj.TipoCofre).CantObjs)
                NumeritoX = RandomNumber(1, 100)
                Do While (NumeritoX > CofresAzar(obj.TipoCofre).ObjProbability(NumCofre))
                    NumCofre = RandomNumber(1, CofresAzar(obj.TipoCofre).CantObjs)
                    NumeritoX = RandomNumber(1, 100)
                Loop
                        'Llegó acá, osea que encontró un obj para entregar.
                        MiObj.ObjIndex = CofresAzar(obj.TipoCofre).ObjIndex(NumCofre)
                        MiObj.Amount = CofresAzar(obj.TipoCofre).ObjAmount(NumCofre)
                        
                            If Not MeterItemEnInventario(userindex, MiObj) Then
                                Call TirarItemAlPiso(UserList(userindex).Pos, MiObj)
                                Call UpdateUserInv(False, userindex, slot)
                            End If
                            
                        SendData SendTarget.toindex, userindex, 0, "||115@" & MiObj.Amount & "@" & ObjData(MiObj.ObjIndex).Name
            Else
                Dim i As Long
                Dim tmpAoP, tmpCoG As Boolean
                
                tmpAoP = (UCase$(UserList(userindex).clase) = "ASESINO" Or UCase$(UserList(userindex).clase) = "PALADIN")
                tmpCoG = (UCase$(UserList(userindex).clase) = "GUERRERO" Or UCase$(UserList(userindex).clase) = "CAZADOR")
                
                For i = 1 To CofresAzar(obj.TipoCofre).CantObjs
                    'Llegó acá, osea que encontró un obj para entregar.
                        MiObj.ObjIndex = CofresAzar(obj.TipoCofre).ObjIndex(i)
                        MiObj.Amount = CofresAzar(obj.TipoCofre).ObjAmount(i)
                        
                        If tmpAoP And MiObj.ObjIndex = 206 Then
                            MiObj.ObjIndex = 530
                        End If
                        
                        If tmpCoG Then
                            Select Case MiObj.ObjIndex
                                Case 206
                                    MiObj.ObjIndex = 478
                                Case 207
                                    MiObj.ObjIndex = 551
                                    MiObj.Amount = 100
                                Case 1012
                                    MiObj.ObjIndex = 404
                            End Select
                        End If
                        
                            If Not MeterItemEnInventario(userindex, MiObj) Then
                                Call TirarItemAlPiso(UserList(userindex).Pos, MiObj)
                                Call UpdateUserInv(False, userindex, slot)
                            End If
                            
                        SendData SendTarget.toindex, userindex, 0, "||115@" & MiObj.Amount & "@" & ObjData(MiObj.ObjIndex).Name
                Next i
            End If
                
        Exit Sub
        
       
    Case eOBJType.otCajasDios
        Dim CofreNecesita As Integer
         
         If UserList(userindex).flags.SirvienteDeDios = "Tarraske" Then
            CofreNecesita = 1479
         ElseIf UserList(userindex).flags.SirvienteDeDios = "Mifrit" Then
            CofreNecesita = 1475
         ElseIf UserList(userindex).flags.SirvienteDeDios = "Erebros" Then
            CofreNecesita = 1473
         ElseIf UserList(userindex).flags.SirvienteDeDios = "Poseidon" Then
            CofreNecesita = 1477
         End If
         
         If UserList(userindex).Invent.Object(slot).ObjIndex = CofreNecesita Then
         
            Dim CofreCerrado As obj
            CofreCerrado.Amount = 1
            CofreCerrado.ObjIndex = CofreNecesita + 1
            
               QuitarObjetos CofreNecesita, 1, userindex
               MeterItemEnInventario userindex, CofreCerrado
         
         Dim ItemIndex As Integer
         
         Dim X As Integer
         For X = 1 To MAX_INVENTORY_SLOTS
         ItemIndex = UserList(userindex).Invent.Object(X).ObjIndex
         
            If UserList(userindex).Invent.Object(X).ObjIndex > 0 Then
             If ObjData(ItemIndex).ItemDios = 1 Then
              If ObjData(ItemIndex).OBJType = otArmadura Or ObjData(ItemIndex).OBJType = otcASCO Or ObjData(ItemIndex).OBJType = otESCUDO Or ObjData(ItemIndex).OBJType = otWeapon Or ObjData(ItemIndex).OBJType = otHerramientas Then
            
               UserList(userindex).CofreDios.Cant = UserList(userindex).CofreDios.Cant + 1
               UserList(userindex).CofreDios.Item(UserList(userindex).CofreDios.Cant) = UserList(userindex).Invent.Object(X).ObjIndex
               
               Call QuitarObjetos(UserList(userindex).Invent.Object(X).ObjIndex, 1, userindex)
               Call DarCuerpoDesnudo(userindex)
              End If
             End If
            End If
         
         Next X
         
         Exit Sub
         End If
         
         
         If UserList(userindex).Invent.Object(slot).ObjIndex = CofreNecesita + 1 Then
         
         Dim Inventario As Integer, Items As Integer
         
         Items = 0
         
         For Inventario = 1 To MAX_INVENTORY_SLOTS
            If UserList(userindex).Invent.Object(Inventario).ObjIndex > 0 Then
               Items = Items + 1
            End If
         Next Inventario
         
         If Items > MAX_INVENTORY_SLOTS - UserList(userindex).CofreDios.Cant Then
            SendData SendTarget.toindex, userindex, 0, "||116"
          Exit Sub
         End If
         
         Dim ObjetoDios As obj, xx As Integer
         
         For xx = 1 To 4
         ObjetoDios.Amount = 1
         ObjetoDios.ObjIndex = UserList(userindex).CofreDios.Item(xx)
         
            If ObjetoDios.ObjIndex > 0 Then
                MeterItemEnInventario userindex, ObjetoDios
            End If
         
            UserList(userindex).CofreDios.Item(xx) = 0
            UserList(userindex).CofreDios.Cant = 0
         Next xx
         
            Dim CofreAbierto As obj
            CofreAbierto.Amount = 1
            CofreAbierto.ObjIndex = CofreNecesita
            
               QuitarObjetos CofreNecesita + 1, 1, userindex
               MeterItemEnInventario userindex, CofreAbierto
            
         End If
         
         Call UpdateUserInv(True, userindex, 0)
         
    Case eOBJType.otSubeClan6
        If UserList(userindex).GuildIndex > 0 Then
          If Guilds(UserList(userindex).GuildIndex).NivelClan = 5 And (m_EsGuildLeader(UserList(userindex).Name, UserList(userindex).GuildIndex) Or m_EsGuildSubLeader1(UserList(userindex).Name, UserList(userindex).GuildIndex) Or m_EsGuildSubLeader2(UserList(userindex).Name, UserList(userindex).GuildIndex)) Then
              If TieneObjetos(1481, 1, userindex) = True Then
                    Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan", "6")
                    Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||210@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@6")
                    Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||211@6@24")
                    Call QuitarObjetos(1481, 1, userindex)
                    Exit Sub
              Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||213")
              Exit Sub
              End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||208")
              Exit Sub
            End If
        End If
    
    Case eOBJType.otSubeClan7
        If UserList(userindex).GuildIndex > 0 Then
          If Guilds(UserList(userindex).GuildIndex).NivelClan = 6 And (m_EsGuildLeader(UserList(userindex).Name, UserList(userindex).GuildIndex) Or m_EsGuildSubLeader1(UserList(userindex).Name, UserList(userindex).GuildIndex) Or m_EsGuildSubLeader2(UserList(userindex).Name, UserList(userindex).GuildIndex)) Then
              If TieneObjetos(1482, 1, userindex) = True Then
                Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan", "7")
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||210@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@7")
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||211@7@28")
                Call QuitarObjetos(1482, 1, userindex)
                Exit Sub
                Else
                      Call SendData(SendTarget.toindex, userindex, 0, "||214")
                Exit Sub
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||208")
              Exit Sub
            End If
        End If
         
    Case eOBJType.otRenunciaH
    
            If UserList(userindex).GuildIndex > 0 Then Call SendData(SendTarget.toindex, userindex, 0, "||302"): Exit Sub
        
            If (UserList(userindex).StatusMith.EsStatus = 2 Or EsHorda(userindex)) Then
                UserList(userindex).StatusMith.EsStatus = 1
                UserList(userindex).StatusMith.EligioStatus = 1
                Call SendData(SendTarget.toindex, userindex, 0, "||355")
                Call SendUserStatux(userindex)
                        
                If UserList(userindex).Faccion.FuerzasCaos = 1 Then
                    Call ExpulsarFaccionCaos(userindex)
                End If
                
                Call QuitarUserInvItem(userindex, slot, 1)
                Call UpdateUserInv(False, userindex, slot)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||239")
            End If
            
    Case eOBJType.otRenunciaA
    
            If UserList(userindex).GuildIndex > 0 Then Call SendData(SendTarget.toindex, userindex, 0, "||302"): Exit Sub
        
            If (UserList(userindex).StatusMith.EsStatus = 1 Or EsAlianza(userindex)) Then
                UserList(userindex).StatusMith.EsStatus = 2
                UserList(userindex).StatusMith.EligioStatus = 1
                Call SendData(SendTarget.toindex, userindex, 0, "||355")
                Call SendUserStatux(userindex)
                        
                If UserList(userindex).Faccion.ArmadaReal = 1 Then
                    Call ExpulsarFaccionReal(userindex)
                End If
                
                Call QuitarUserInvItem(userindex, slot, 1)
                Call UpdateUserInv(False, userindex, slot)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||239")
            End If
         
    Case eOBJType.otSacos
        
        UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion + obj.cantCredits
        Call SendData(SendTarget.toindex, userindex, 0, "||930@" & obj.cantCredits)
        
        Call GuardarLogs("" & Date & " " & time & " " & UserList(userindex).Name & " obtuvo " & obj.cantCredits, "\Turbios\Sacos")
        
        Call QuitarUserInvItem(userindex, slot, 1)
        Call UpdateUserInv(False, userindex, slot)
         
    Case eOBJType.otScroll
        
        With UserList(userindex)
            Dim typeScroll As Byte
            typeScroll = obj.typeScroll
            
            If (Not .flags.activoScroll(typeScroll)) Then
                .flags.activoScroll(typeScroll) = True
                .Scrolls(typeScroll).time = obj.timeScroll
                .Scrolls(typeScroll).timeScroll = obj.timeScroll
                .Scrolls(typeScroll).multScroll = obj.multScroll
                
                Call SendData(SendTarget.toindex, userindex, 0, "TIS" & obj.typeScroll & "," & obj.timeScroll & "," & obj.timeScroll)
                Call QuitarUserInvItem(userindex, slot, 1)
                Call UpdateUserInv(False, userindex, slot)
                
                Select Case typeScroll
                    Case 1
                        Call SendData(SendTarget.toindex, userindex, 0, "||929@Experiencia@" & obj.timeScroll & "@" & obj.multScroll)
                    Case 2
                        Call SendData(SendTarget.toindex, userindex, 0, "||929@Oro@" & obj.timeScroll & "@" & obj.multScroll)
                    Case 3
                        Call SendData(SendTarget.toindex, userindex, 0, "||929@Drop@" & obj.timeScroll & "@" & obj.multScroll)
                    Case 4
                        Call SendData(SendTarget.toindex, userindex, 0, "||929@Drop de Cristales@" & obj.timeScroll & "@" & obj.multScroll)
                End Select
                
                
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||928")
            End If
            
        End With
            
        
        
    Case eOBJType.otMontura
    

        If UserList(userindex).flags.Muerto = 1 Then
            SendData SendTarget.toindex, userindex, 0, "||3"
            Exit Sub
        End If
        
        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
        If MapaEspecial(userindex) Then Exit Sub
        
        If UserList(userindex).flags.Transformado = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||319")
            Exit Sub
        End If
        
        If UserList(userindex).flags.Montando = 0 Then
            UserList(userindex).Char.Body = obj.Ropaje
            UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
            UserList(userindex).Char.ShieldAnim = NingunEscudo
            UserList(userindex).Char.WeaponAnim = NingunArma
            UserList(userindex).Char.CascoAnim = UserList(userindex).Char.CascoAnim
            UserList(userindex).flags.Montando = 1
         
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "USM" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).flags.Montando)
            
            If (obj.esVoladora = 1) Then
                UserList(userindex).flags.levitando = 1
                SendUserMontVol (userindex)
            End If
            
            Call ChangeUserChar(toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            Call SendData(toindex, userindex, 0, "EQUIT")
        Else
             With UserList(userindex)
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1 Then
                    SendData SendTarget.toindex, userindex, 0, "||947"
                    Exit Sub
                End If
            End With
        
            UserList(userindex).flags.Montando = 0
            UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
                
                If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
                    UserList(userindex).Char.Body = equiparRopaje(userindex)
                Else
                    Call DarCuerpoDesnudo(userindex)
                End If
                
                UserList(userindex).flags.levitando = 0
                SendUserMontVol (userindex)
                
                If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then UserList(userindex).Char.ShieldAnim = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ShieldAnim
                If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then UserList(userindex).Char.WeaponAnim = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).WeaponAnim
                If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then UserList(userindex).Char.CascoAnim = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).CascoAnim
         
                Call ChangeUserChar(toMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "USM" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).flags.Montando)
                Call SendData(toindex, userindex, 0, "EQUIT")
        End If
        
        SendUserHitBox (userindex)
    
    Case eOBJType.otPocionResu
    

        If UserList(userindex).flags.Muerto = 0 Then
            SendData SendTarget.toindex, userindex, 0, "||117"
            Exit Sub
        End If
        
        If UserList(userindex).Pos.Map = 71 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 104 Or UserList(userindex).Pos.Map = 106 Or UserList(userindex).Pos.Map = 107 Or UserList(userindex).Pos.Map = 108 Or UserList(userindex).Pos.Map = 109 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 111 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 120 Or UserList(userindex).Pos.Map = 166 Or UserList(userindex).Pos.Map = 164 Or UserList(userindex).Pos.Map = 162 Then
                SendData SendTarget.toindex, userindex, 0, "||118"
            Exit Sub
        End If
        
            Call RevivirUsuario(userindex)
            UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
            SendUserHP (userindex)
            SendData SendTarget.toindex, userindex, 0, "||119"
        
        Call QuitarUserInvItem(userindex, slot, 1)
        Call UpdateUserInv(False, userindex, slot)

    Case eOBJType.otFragmento
        
        If UserList(userindex).GuildIndex = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||120")
          Exit Sub
        End If
        
        If Not m_EsGuildLeader(UserList(userindex).Name, UserList(userindex).GuildIndex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||121")
          Exit Sub
        End If
        
        Dim FragmObj As obj
        FragmObj.ObjIndex = 1481
        FragmObj.Amount = 1
        
        If TieneObjetos(1271, 1, userindex) = True And TieneObjetos(1272, 1, userindex) = True Then
            
            Call QuitarObjetos(1271, 1, userindex)
            Call QuitarObjetos(1272, 1, userindex)
        
            If Not MeterItemEnInventario(userindex, FragmObj) Then
                Call TirarItemAlPiso(UserList(userindex).Pos, FragmObj)
            End If
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||122")
          Exit Sub
        End If
        
        
    Case eOBJType.otGemaOctarina
            
            Dim NumMejorados As Integer, Requiere As obj, TengoItem As Byte
            
            NumMejorados = val(GetVar(DatPath & "Mejorados.dat", "INIT", "NumMejorados"))
            
            TengoItem = 0
            
            For Items = 1 To NumMejorados
            
            Requiere.ObjIndex = val(GetVar(DatPath & "Mejorados.dat", "ITEM" & Items, "Requiere"))
            
                If TieneObjetos(Requiere.ObjIndex, 1, userindex) = True Then
                    TengoItem = 1
                    SendData SendTarget.toindex, userindex, 0, "MJOR" & ObjData(Requiere.ObjIndex).Name
                End If
            
            Next Items
            
            
            If TengoItem = 0 Then
                SendData SendTarget.toindex, userindex, 0, "MJOR" & "Sin items mejorables"
            Exit Sub
            End If
       
    Case eOBJType.otAriete
        Dim SpawnAriete As WorldPos
        
        'If Fortaleza = guilds(UserList(UserIndex).GuildIndex).GuildName Then exit sub
        If UserList(userindex).Pos.Map <> 167 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||123")
            Exit Sub
         End If
    
         If UserList(userindex).GuildIndex = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||120")
            Exit Sub
         End If
         
         If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) = UCase$(Fortaleza) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||124")
            Exit Sub
         End If
         
         If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(CastilloNorte) Or UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(CastilloSur) Or UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(CastilloEste) Or UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(CastilloOeste) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||125")
            Exit Sub
         End If
    
        If UserList(userindex).Pos.Map = 167 And UserList(userindex).Pos.X = 49 And UserList(userindex).Pos.Y = 85 Then
            SpawnAriete.Map = 167
            SpawnAriete.X = 49
            SpawnAriete.Y = 85
            
            If RejaSur = 0 Then Exit Sub
            If MapData(167, 49, 84).OBJInfo.ObjIndex <> 1471 Then Exit Sub
            
            '/Lo corremos
            Dim yeguaX As Long
            Dim yeguaY As Long
            
            For yeguaX = 1 To 3
                For yeguaY = 1 To 3
              If LegalPos(167, 49 + yeguaX, 85 + yeguaY) Then
                    Call WarpUserChar(userindex, 167, 49 + yeguaX, 85 + yeguaY)
              Exit For
              End If
            Next yeguaY
            Next yeguaX
            
            
            ArieteUno = SpawnNpc(621, SpawnAriete, True, False)
            SendData SendTarget.ToAll, 0, 0, "ARIE" & Npclist(ArieteUno).Char.CharIndex
            ActivarTimerRejas = True
            RejaSurAtacada = True
            Call QuitarObjetos(1469, 1, userindex)
            frmMain.Rejas.Enabled = True
        ElseIf UserList(userindex).Pos.Map = 167 And UserList(userindex).Pos.X = 49 And UserList(userindex).Pos.Y = 69 Then
            SpawnAriete.Map = 167
            SpawnAriete.X = 49
            SpawnAriete.Y = 69
            
            If RejaCentral = 0 Then Exit Sub
            If MapData(167, 49, 68).OBJInfo.ObjIndex <> 1471 Then Exit Sub
            
            For yeguaX = 1 To 3
                For yeguaY = 1 To 3
              If LegalPos(167, 49 + yeguaX, 69 + yeguaY) Then
                    Call WarpUserChar(userindex, 167, 49 + yeguaX, 69 + yeguaY)
              Exit For
              End If
            Next yeguaY
            Next yeguaX
            
            ArieteDos = SpawnNpc(621, SpawnAriete, True, False)
            SendData SendTarget.ToAll, 0, 0, "ARIE" & Npclist(ArieteDos).Char.CharIndex
            ActivarTimerRejas = True
            RejaCentralAtacada = True
            Call QuitarObjetos(1469, 1, userindex)
            frmMain.Rejas.Enabled = True
        ElseIf UserList(userindex).Pos.Map = 167 And UserList(userindex).Pos.X = 49 And UserList(userindex).Pos.Y = 49 Then
            SpawnAriete.Map = 167
            SpawnAriete.X = 49
            SpawnAriete.Y = 49
            
            If RejaNorte = 0 Then Exit Sub
            If MapData(167, 49, 48).OBJInfo.ObjIndex <> 1471 Then Exit Sub
        
            For yeguaX = 1 To 3
                For yeguaY = 1 To 3
              If LegalPos(167, 49 + yeguaX, 49 + yeguaY) Then
                    Call WarpUserChar(userindex, 167, 49 + yeguaX, 49 + yeguaY)
              Exit For
              End If
            Next yeguaY
            Next yeguaX
            
            ArieteTres = SpawnNpc(621, SpawnAriete, True, False)
            SendData SendTarget.ToAll, 0, 0, "ARIE" & Npclist(ArieteTres).Char.CharIndex
            ActivarTimerRejas = True
            RejaNorteAtacada = True
            Call QuitarObjetos(1469, 1, userindex)
            frmMain.Rejas.Enabled = True
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||123")
        Exit Sub
        End If
       
    Case eOBJType.otMapaTesoro
        If MapData(MapaTesoroMap, MapaTesoroX, MapaTesoroY).Blocked = 1 Then
            Call Tesoros
        End If
        Call SendData(SendTarget.toindex, userindex, 0, "||126@" & MapaTesoroMap & "@" & MapaTesoroX & "@" & MapaTesoroY)

    Case eOBJType.otCristales
        If TieneObjetos(1274, 1, userindex) = False Then
            Call SendData(SendTarget.toindex, userindex, 0, "||127")
         Exit Sub
        End If
        
        Dim CristalitosNew As Long
        CristalitosNew = RandomNumber(obj.CristalesMin, obj.CristalesMax)
        
        UserList(userindex).flags.AlmasContenidas = UserList(userindex).flags.AlmasContenidas + CristalitosNew
        Call SendData(SendTarget.toindex, userindex, 0, "||128@" & CristalitosNew)
        
        Call QuitarUserInvItem(userindex, slot, 1)
        Call UpdateUserInv(False, userindex, slot)
        
    Case eOBJType.otContenedor
        If UserList(userindex).flags.AlmasContenidas = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||129")
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||130@" & UserList(userindex).flags.AlmasContenidas)
        End If
        
    Case eOBJType.otGemaNegra
            Dim tStr1 As String
            Dim tStr2 As String
            Dim tStr3 As String
            Dim tStr4 As String
            tStr1 = SendNobleList1(userindex)
            tStr2 = SendNobleList2(userindex)
            tStr3 = SendNobleList3(userindex)
            tStr4 = SendNobleList4(userindex)
            Call SendData(SendTarget.toindex, userindex, 0, "8G1" & SendNobleList1(userindex))
            Call SendData(SendTarget.toindex, userindex, 0, "8G2" & SendNobleList2(userindex))
            Call SendData(SendTarget.toindex, userindex, 0, "8G3" & SendNobleList3(userindex))
            Call SendData(SendTarget.toindex, userindex, 0, "8G4" & SendNobleList4(userindex))
            
    Case eOBJType.otGemaSagrada
        If TieneObjetos(1091, 1, userindex) = False And TieneObjetos(1093, 1, userindex) = False Then
            Call SendData(SendTarget.toindex, userindex, 0, "||131")
         Exit Sub
        End If
        
     Dim RandomDragonOscuro As Byte
      RandomDragonOscuro = RandomNumber(1, 100)
      
      If RandomDragonOscuro <= 50 Then
        Call AgregarPuntos(userindex, 200)
        Call SendData(SendTarget.toindex, userindex, 0, "||132")
        Call SendData(SendTarget.toindex, userindex, 0, "||57@200")
      ElseIf RandomDragonOscuro >= 51 And RandomDragonOscuro <= 74 Then
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + 200000000
        Call SendData(SendTarget.toindex, userindex, 0, "||132")
        Call SendData(SendTarget.toindex, userindex, 0, "||63@200.000.000")
        SendUserGLD (userindex)
     ElseIf RandomDragonOscuro >= 85 And RandomDragonOscuro <= 100 Then
     
        If TieneHechizo(53, userindex) Then Exit Sub
        
        Dim j As Integer
        If TieneObjetos(1091, 1, userindex) = True Then
           If Not TieneHechizo(53, userindex) Then
               'Buscamos un slot vacio
               For j = 1 To MAXUSERHECHIZOS
                   If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
               Next j
                   
               If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
                   Exit Sub
               Else
                   UserList(userindex).Stats.UserHechizos(j) = 53
                   Call UpdateUserHechizos(False, userindex, CByte(j))
               End If
           End If
        End If
        
        UserList(userindex).flags.CaballerodelDragon = 1
        
        Call SendData(SendTarget.toindex, userindex, 0, "||133")
        Call SendData(SendTarget.ToAll, 0, 0, "||134@" & UserList(userindex).Name)
     End If
     
     Call QuitarObjetos(1091, 1, userindex)
     Call QuitarObjetos(1092, 1, userindex)

     Case eOBJType.otBebidas
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||5")
            Exit Sub
        End If
        UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU + obj.MinSed
        If UserList(userindex).Stats.MinAGU > UserList(userindex).Stats.MaxAGU Then _
            UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MaxAGU
        UserList(userindex).flags.Sed = 0
        Call EnviarHambreYsed(userindex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, slot, 1)
        
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_BEBER)
        
        Call UpdateUserInv(False, userindex, slot)
    
    Case eOBJType.otLlaves
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||5")
            Exit Sub
        End If
        
        If UserList(userindex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(userindex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = obj.clave Then
         
                        MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                        UserList(userindex).flags.TargetObj = MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex
                        Call SendData(SendTarget.toindex, userindex, 0, "||100")
                        Exit Sub
                     Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||101")
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = obj.clave Then
                        MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                        Call SendData(SendTarget.toindex, userindex, 0, "||136")
                        UserList(userindex).flags.TargetObj = MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex
                        Exit Sub
                     Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||101")
                        Exit Sub
                     End If
                  End If
            Else
                  Call SendData(SendTarget.toindex, userindex, 0, "||102")
                  Exit Sub
            End If
            
        End If
    
        Case eOBJType.otBotellaVacia
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||5")
                Exit Sub
            End If
            If Not HayAgua(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||103")
                Exit Sub
            End If
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(userindex, slot, 1)
            If Not MeterItemEnInventario(userindex, MiObj) Then
                Call TirarItemAlPiso(UserList(userindex).Pos, MiObj)
            End If
            
            Call UpdateUserInv(False, userindex, slot)
    
        Case eOBJType.otBotellaLlena
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||5")
                Exit Sub
            End If
            UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU + obj.MinSed
            If UserList(userindex).Stats.MinAGU > UserList(userindex).Stats.MaxAGU Then _
                UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MaxAGU
            UserList(userindex).flags.Sed = 0
            Call EnviarHambreYsed(userindex)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).IndexCerrada
            Call QuitarUserInvItem(userindex, slot, 1)
            If Not MeterItemEnInventario(userindex, MiObj) Then
                Call TirarItemAlPiso(UserList(userindex).Pos, MiObj)
            End If
            
            
        Case eOBJType.otHerramientas
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||5")
                Exit Sub
            End If
            If Not UserList(userindex).Stats.MinSta > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||17")
                Exit Sub
            End If
            
            If UserList(userindex).Invent.Object(slot).Equipped = 0 Then Exit Sub
            
            Select Case ObjIndex
                Case CAÑA_PESCA, RED_PESCA
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Pesca)
                Case HACHA_LEÑADOR
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Talar)
                Case PIQUETE_MINERO
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Mineria)
                Case MARTILLO_HERRERO
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Herreria)
                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(userindex)
                    Call SendData(SendTarget.toindex, userindex, 0, "SFC")

            End Select
        
        Case eOBJType.otPergaminos
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||5")
                Exit Sub
            End If
            
            If (Hechizos(ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).HechizoIndex).ExclusivoClase = UCase$(UserList(userindex).clase)) Or (Hechizos(ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).HechizoIndex).ExclusivoClasedos = UCase$(UserList(userindex).clase)) Or _
            (Len(Hechizos(ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).HechizoIndex).ExclusivoClase) = 0 And Len(Hechizos(ObjData(UserList(userindex).Invent.Object(slot).ObjIndex).HechizoIndex).ExclusivoClasedos) = 0) Then
            Call AgregarHechizo(userindex, slot)
            Call UpdateUserInv(False, userindex, slot)
            Else
            Call SendData(SendTarget.toindex, userindex, 0, "||104")
            End If
       
       Case eOBJType.otMinerales
           If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||5")
                Exit Sub
           End If
           Call SendData(SendTarget.toindex, userindex, 0, "T01" & FundirMetal)
       
       Case eOBJType.otInstrumentos
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||5")
                Exit Sub
            End If
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & obj.Snd1)
       
       Case eOBJType.otBarcos
    'Verifica si esta aproximado al agua antes de permitirle navegar
        If UserList(userindex).Stats.ELV < 30 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||105")
            Exit Sub
        End If
        If ((LegalPos(UserList(userindex).Pos.Map, UserList(userindex).Pos.X - 1, UserList(userindex).Pos.Y, True) Or _
            LegalPos(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y - 1, True) Or _
            LegalPos(UserList(userindex).Pos.Map, UserList(userindex).Pos.X + 1, UserList(userindex).Pos.Y, True) Or _
            LegalPos(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y + 1, True)) And _
            UserList(userindex).flags.Navegando = 0) _
            Or UserList(userindex).flags.Navegando = 1 Then
           Call DoNavega(userindex, obj, slot)
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||106")
        End If
           
End Select

'Actualiza
'Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Sub EnivarArmasConstruibles(ByVal userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmasHerrero)
    If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(userindex).Stats.UserSkills(eSkill.Herreria) \ ModHerreriA(UserList(userindex).clase) Then
        If ObjData(ArmasHerrero(i)).OBJType = eOBJType.otWeapon Then
        '[DnG!]
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & " (" & ObjData(ArmasHerrero(i)).LingH & "-" & ObjData(ArmasHerrero(i)).LingP & "-" & ObjData(ArmasHerrero(i)).LingO & ")" & "," & ArmasHerrero(i) & ","
        '[/DnG!]
        Else
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & "," & ArmasHerrero(i) & ","
        End If
    End If
Next i

Call SendData(SendTarget.toindex, userindex, 0, "LAH" & cad$)

End Sub
 
Sub EnivarObjConstruibles(ByVal userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ObjCarpintero)
    If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(userindex).Stats.UserSkills(eSkill.Carpinteria) / ModCarpinteria(UserList(userindex).clase) Then _
        cad$ = cad$ & ObjData(ObjCarpintero(i)).Name & "          (Madera: " & ObjData(ObjCarpintero(i)).Madera & " Piedras Mágicas: " & ObjData(ObjCarpintero(i)).Piedras & ")" & "," & ObjCarpintero(i) & ","
Next i

Call SendData(SendTarget.toindex, userindex, 0, "OBR" & cad$)

End Sub

Sub EnivarArmadurasConstruibles(ByVal userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmadurasHerrero)
    If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(userindex).clase) Then
        '[DnG!]
        cad$ = cad$ & ObjData(ArmadurasHerrero(i)).Name & " (" & ObjData(ArmadurasHerrero(i)).LingH & "-" & ObjData(ArmadurasHerrero(i)).LingP & "-" & ObjData(ArmadurasHerrero(i)).LingO & ")" & "," & ArmadurasHerrero(i) & ","
        '[/DnG!]
    End If
Next i

Call SendData(SendTarget.toindex, userindex, 0, "LAR" & cad$)

End Sub

Sub TirarTodo(ByVal userindex As Integer)
On Error Resume Next

If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 6 Then Exit Sub
If UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 89 Or UserList(userindex).Pos.Map = 150 Or UserList(userindex).Pos.Map = 132 Or UserList(userindex).Pos.Map = 133 Or UserList(userindex).Pos.Map = 134 Or UserList(userindex).Pos.Map = 135 Or UserList(userindex).Pos.Map = 143 Then Exit Sub
If MapaEspecial(userindex) And UserList(userindex).Pos.Map <> 118 Then Exit Sub

Call TirarTodosLosItems(userindex)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean

ItemSeCae = (ObjData(index).Real <> 1 Or ObjData(index).NoSeCae = 0) And _
            (ObjData(index).Caos <> 1 Or ObjData(index).NoSeCae = 0) And _
            ObjData(index).OBJType <> eOBJType.otLlaves And _
            ObjData(index).OBJType <> eOBJType.otBarcos And _
            ObjData(index).NoSeCae = 0 And _
            ObjData(index).Intransferible = 0


End Function
Sub TirarTodosLosItems(ByVal userindex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As obj
    Dim ItemIndex As Integer
    If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 6 Then Exit Sub
    If UserList(userindex).Pos.Map = 27 And HayGuerraKhalim = True Then Exit Sub
    If UserList(userindex).Pos.Map = 29 And HayGuerraAnvil = True Then Exit Sub
    If UserList(userindex).Pos.Map <> 190 And (UserList(userindex).flags.EnJDH) Then Exit Sub
    If UserList(userindex).Pos.Map = 28 Or UserList(userindex).Pos.Map = 89 Or UserList(userindex).Pos.Map = 130 Then Exit Sub
 
For i = 1 To MAX_INVENTORY_SLOTS
If UserList(userindex).Invent.Object(i).ObjIndex = SacriIndex Then
If DropSacri = 0 Then
NuevaPos.X = 0: NuevaPos.Y = 0
MiObj.Amount = UserList(userindex).Invent.Object(i).Amount: MiObj.ObjIndex = SacriIndex
Call Tilelibre(UserList(userindex).Pos, NuevaPos, MiObj)
If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call DropObj(userindex, i, 1, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
Else
Call QuitarUserInvItem(userindex, i, 1)
Call UpdateUserInv(False, userindex, i)
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & 22 & "," & 0)
End If
Exit Sub
End If
Next i
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(userindex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.Amount = UserList(userindex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                
                Tilelibre UserList(userindex).Pos, NuevaPos, MiObj
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub
Sub DameTodoObj(ByVal userindex As Integer, ByVal slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim obj As obj


If num > 0 Then
  
  If num > UserList(userindex).Invent.Object(slot).Amount Then num = UserList(userindex).Invent.Object(slot).Amount
  
  'Check objeto en el suelo
  If MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex = 0 Or MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex Then
        If UserList(userindex).Invent.Object(slot).Equipped = 1 Then Call Desequipar(userindex, slot)
        obj.ObjIndex = UserList(userindex).Invent.Object(slot).ObjIndex
        
'        If ObjData(Obj.ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
'            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes tirar el objeto." & FONTTYPE_INFO)
'            Exit Sub
'        End If
        
        If num + MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.Amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.Amount
        End If
        
        obj.Amount = num
        
        Call MakeObj(SendTarget.toMap, 0, Map, obj, Map, X, Y)
        Call QuitarUserInvItem(userindex, slot, num)
        Call UpdateUserInv(False, userindex, slot)
        
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(userindex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name, False)
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGMss(UserList(userindex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name, False)
  Else
    Call SendData(SendTarget.toindex, userindex, 0, "||107")
  End If
    
End If

End Sub
Sub DameBancoObj(ByVal userindex As Integer, ByVal slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim obj As obj


If num > 0 Then
  'If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount
  
  'Check objeto en el suelo
  If MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex = 0 Or MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex = UserList(userindex).BancoInvent.Object(slot).ObjIndex Then
        obj.ObjIndex = UserList(userindex).BancoInvent.Object(slot).ObjIndex
        
        If num + MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.Amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.Amount
        End If
        
        obj.Amount = num
        
        Call MakeObj(SendTarget.toMap, 0, Map, obj, Map, X, Y)
        Call QuitarBancoInvItem(userindex, slot, num)
        Call UpdateBanUserInv(False, userindex, slot)
        
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(userindex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name, False)
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGMss(UserList(userindex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name, False)
  Else
    Call SendData(SendTarget.toindex, userindex, 0, "||107")
  End If
    
End If

End Sub
Sub DameBanco(ByVal userindex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        ItemIndex = UserList(userindex).BancoInvent.Object(i).ObjIndex
        If ItemIndex > 0 Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.Amount = UserList(userindex).BancoInvent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                
                Tilelibre UserList(userindex).Pos, NuevaPos, MiObj
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DameBancoObj(userindex, i, MiObj.Amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
        End If
    Next i
End Sub
Sub DameTodo(ByVal userindex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(userindex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.Amount = UserList(userindex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                
                Tilelibre UserList(userindex).Pos, NuevaPos, MiObj
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DameTodoObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
        End If
    Next i
End Sub
Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function
Sub TirarTodosLosItemsNoNewbies(ByVal userindex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As obj
Dim ItemIndex As Integer
If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 6 Then Exit Sub
 
For i = 1 To MAX_INVENTORY_SLOTS
If UserList(userindex).Invent.Object(i).ObjIndex = SacriIndex Then
If DropSacri = 0 Then
NuevaPos.X = 0: NuevaPos.Y = 0
MiObj.Amount = UserList(userindex).Invent.Object(i).Amount: MiObj.ObjIndex = SacriIndex
Call Tilelibre(UserList(userindex).Pos, NuevaPos, MiObj)
If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call DropObj(userindex, i, 1, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
Else
Call QuitarUserInvItem(userindex, i, 1)
Call UpdateUserInv(False, userindex, i)
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & 22 & "," & 0)
End If
Exit Sub
End If
Next i

If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(userindex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.Amount = UserList(userindex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            
            Tilelibre UserList(userindex).Pos, NuevaPos, MiObj
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub

Attribute VB_Name = "InvNpc"
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
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, obj As obj) As WorldPos
On Error GoTo Errhandler

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    Call Tilelibre(Pos, NuevaPos, obj)
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
          Call MakeObj(SendTarget.toMap, 0, Pos.Map, _
                obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
          TirarItemAlPiso = NuevaPos
    End If

Exit Function
Errhandler:

End Function
Public Sub NPC_TIRAR_CRISTALES(ByRef npc As npc, ByVal userindex As Integer)

On Error Resume Next

If UserList(userindex).Stats.ELV >= 60 Then


If npc.CristalesPequesMin > 0 Then
Dim NuevaPos As WorldPos, MiObj As obj

Dim i As Integer

Dim RandomCristales As Integer

RandomCristales = RandomNumber(npc.CristalesPequesMin, npc.CristalesPequesMax)
If (UserList(userindex).flags.activoScroll(4)) Then RandomCristales = RandomCristales * (UserList(userindex).Scrolls(4).multScroll)

For i = 1 To RandomCristales

NuevaPos.X = 0
NuevaPos.Y = 0
'Creo el Obj
MiObj.Amount = 1
MiObj.ObjIndex = 1275

NuevaPos.X = 0
NuevaPos.Y = 0
TilelibreCristales npc.Pos, NuevaPos, MiObj

If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
    Call MakeObj(SendTarget.toMap, 0, UserList(userindex).Pos.Map, MiObj, UserList(userindex).Pos.Map, NuevaPos.X, NuevaPos.Y)
End If
Next i
End If

If npc.CristalesMedianosMin > 0 Then

RandomCristales = RandomNumber(npc.CristalesMedianosMin, npc.CristalesMedianosMax)
If (UserList(userindex).flags.activoScroll(4)) Then RandomCristales = RandomCristales * (UserList(userindex).Scrolls(4).multScroll)

For i = 1 To RandomCristales

NuevaPos.X = 0
NuevaPos.Y = 0
'Creo el Obj
MiObj.Amount = 1
MiObj.ObjIndex = 1276

NuevaPos.X = 0
NuevaPos.Y = 0
TilelibreCristales npc.Pos, NuevaPos, MiObj
If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
Call MakeObj(SendTarget.toMap, 0, UserList(userindex).Pos.Map, MiObj, UserList(userindex).Pos.Map, NuevaPos.X, NuevaPos.Y)
End If
Next i
End If

RandomCristales = RandomNumber(npc.CristalesGrandesMin, npc.CristalesGrandesMax)
If (UserList(userindex).flags.activoScroll(4)) Then RandomCristales = RandomCristales * (UserList(userindex).Scrolls(4).multScroll)

If npc.CristalesGrandesMin > 0 Then

For i = 1 To RandomCristales

NuevaPos.X = 0
NuevaPos.Y = 0
'Creo el Obj
MiObj.Amount = 1
MiObj.ObjIndex = 1277

NuevaPos.X = 0
NuevaPos.Y = 0
TilelibreCristales npc.Pos, NuevaPos, MiObj
If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
Call MakeObj(SendTarget.toMap, 0, UserList(userindex).Pos.Map, MiObj, UserList(userindex).Pos.Map, NuevaPos.X, NuevaPos.Y)
End If
Next i
End If


If npc.CristalesEpicosMin > 0 Then

RandomCristales = RandomNumber(npc.CristalesEpicosMin, npc.CristalesEpicosMax)
If (UserList(userindex).flags.activoScroll(4)) Then RandomCristales = RandomCristales * (UserList(userindex).Scrolls(4).multScroll)

For i = 1 To RandomCristales

NuevaPos.X = 0
NuevaPos.Y = 0
'Creo el Obj
MiObj.Amount = 1
MiObj.ObjIndex = 1278

NuevaPos.X = 0
NuevaPos.Y = 0
TilelibreCristales npc.Pos, NuevaPos, MiObj
If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
Call MakeObj(SendTarget.toMap, 0, UserList(userindex).Pos.Map, MiObj, UserList(userindex).Pos.Map, NuevaPos.X, NuevaPos.Y)
End If
Next i
End If
End If


End Sub
Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc, ByVal userindex As Integer)
'TIRA TODOS LOS ITEMS DEL NPC
On Error Resume Next

If npc.Invent.NroItems > 0 Then
    Dim i As Long
    Dim loopX As Long
    Dim MiObj As obj
    Dim NumerosUsados As Integer
   
    For i = 1 To npc.Invent.NroItems
        If npc.Invent.Object(i).ObjIndex > 0 Then
    
            NumerosUsados = 0
        
                'Probabilidad del npc
                NumerosUsados = (npc.Invent.Object(i).ProbTirar * 2)
                
                If (UserList(userindex).flags.activoScroll(3)) Then
                    NumerosUsados = NumerosUsados * UserList(userindex).Scrolls(3).multScroll
                End If
                
                If NumerosUsados + (npc.Invent.Object(i).ProbTirar * MultiplicadorDrop) > 200 Then
                    NumerosUsados = 200
                Else
                    NumerosUsados = NumerosUsados + (npc.Invent.Object(i).ProbTirar * MultiplicadorDrop)
                End If
                
                'SUERTE DEL PESONAJE
                If UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) <= 18 Then
                    NumerosUsados = NumerosUsados - 1
                ElseIf UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = 19 Then
                    If NumerosUsados < 200 Then NumerosUsados = NumerosUsados
                ElseIf UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = 20 Then
                    If NumerosUsados < 200 Then NumerosUsados = NumerosUsados + 1
                ElseIf UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = 21 Then
                    If NumerosUsados < 200 Then NumerosUsados = NumerosUsados + 2
                ElseIf UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) >= 22 Then
                    If NumerosUsados + (UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) - 21) <= 200 Then
                        NumerosUsados = NumerosUsados + (UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) - 21)
                    Else
                        NumerosUsados = 200
                    End If
                End If
                
                'Tunica de la suerte/maestria/riqueza
                If UserList(userindex).Invent.ArmourEqpObjIndex = 917 Or UserList(userindex).Invent.ArmourEqpObjIndex = 918 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1456 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1497 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1455 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1496 Then
                    If NumerosUsados < 200 Then NumerosUsados = NumerosUsados + 3
                End If
                
                
            Dim NumeroRandom As Byte
            Dim Ret() As Variant
            Dim DropNuevo As Long
            NumeroRandom = RandomNumber(1, 200)
            
            ' indicar el valor máximo, el minimo y la cantidad de números que se van a generar
            Ret = Generar(200, 1, NumerosUsados)
            
            ' Recorre el array y agrega los números
                For loopX = LBound(Ret) To UBound(Ret)
                    UserList(userindex).flags.Probabilidades(loopX) = Ret(loopX)
                Next loopX
                
            For loopX = 1 To NumerosUsados
                If NumeroRandom = UserList(userindex).flags.Probabilidades(loopX) Then
                    MiObj.Amount = npc.Invent.Object(i).Amount
                    MiObj.ObjIndex = npc.Invent.Object(i).ObjIndex
                    Call LogDrops("Drop: " & UserList(userindex).Name & " dropeo el item " & ObjData(npc.Invent.Object(i).ObjIndex).Name & " a las " & time & " " & Date & "")
                    Call TirarItemAlPiso(npc.Pos, MiObj)
                End If
            Next loopX
        End If
    Next i
End If
End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error Resume Next
'Call LogTarea("Function QuedanItems npcindex:" & NpcIndex & " objindex:" & ObjIndex)

Dim i As Integer
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For i = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
            QuedanItems = True
            Exit Function
        End If
    Next
End If
QuedanItems = False
End Function

Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
On Error Resume Next
'Devuelve la cantidad original del obj de un npc

Dim ln As String, npcfile As String
Dim i As Integer

If Npclist(NpcIndex).Numero > 499 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "NPCs.dat"
End If
 
For i = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)
    If ObjIndex = val(ReadField(1, ln, 45)) Then
        EncontrarCant = val(ReadField(2, ln, 45))
        Exit Function
    End If
Next
                   
EncontrarCant = 50

End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
On Error Resume Next

Dim i As Integer

Npclist(NpcIndex).Invent.NroItems = 0

For i = 1 To MAX_INVENTORY_SLOTS
   Npclist(NpcIndex).Invent.Object(i).ObjIndex = 0
   Npclist(NpcIndex).Invent.Object(i).Amount = 0
Next i

Npclist(NpcIndex).InvReSpawn = 0

End Sub

Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = Npclist(NpcIndex).Invent.Object(slot).ObjIndex

    'Quita un Obj
    If ObjData(Npclist(NpcIndex).Invent.Object(slot).ObjIndex).Crucial = 0 Then
        Npclist(NpcIndex).Invent.Object(slot).Amount = Npclist(NpcIndex).Invent.Object(slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(slot).Amount = 0
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
            
            Call updateNPCInventory(slot, NpcIndex)
        End If
    Else
        Npclist(NpcIndex).Invent.Object(slot).Amount = Npclist(NpcIndex).Invent.Object(slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(slot).Amount = 0
            
            If Not QuedanItems(NpcIndex, ObjIndex) Then
                   
                   Npclist(NpcIndex).Invent.Object(slot).ObjIndex = ObjIndex
                   Npclist(NpcIndex).Invent.Object(slot).Amount = EncontrarCant(NpcIndex, ObjIndex)
                   Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
            
            End If
            
            Call updateNPCInventory(slot, NpcIndex)
        End If
            
            If Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
    
    
    
    End If
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)

'Vuelve a cargar el inventario del npc NpcIndex
Dim loopC As Integer
Dim ln As String
Dim npcfile As String

If Npclist(NpcIndex).Numero > 499 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "NPCs.dat"
End If

Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "NROITEMS"))

For loopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & loopC)
    Npclist(NpcIndex).Invent.Object(loopC).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(loopC).Amount = val(ReadField(2, ln, 45))
    
Next loopC

End Sub




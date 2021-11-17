Attribute VB_Name = "ModFacciones"
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

Dim j As Long

Public Const ExpAlUnirse As Long = 50000
Public Const ExpX100 As Integer = 5000


Public Sub EnlistarArmadaReal(ByVal userindex As Integer)

If UserList(userindex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||173")
    Exit Sub
End If

If UserList(userindex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||174")
    Exit Sub
End If

If Not Ciudadano(userindex) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||175")
    Exit Sub
End If

If UserList(userindex).Faccion.CriminalesMatados < FragsJerarquia(1) Then
    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Para enlistarte en la alianza real tienes que a ver matado " & FragsJerarquia(1) & " criminales." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
Exit Sub
End If

If UserList(userindex).Faccion.Reenlistadas = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||176")
Exit Sub
End If

UserList(userindex).Faccion.ArmadaReal = 1
UserList(userindex).Faccion.Reenlistadas = 1
UserList(userindex).Faccion.RecompensasReal = 0
UserList(userindex).StatusMith.EsStatus = 3
UserList(userindex).flags.PJerarquia = 1
UserList(userindex).flags.SJerarquia = 0
UserList(userindex).flags.TJerarquia = 0
UserList(userindex).flags.CJerarquia = 0
Call SendUserStatux(userindex)

Call SendData(SendTarget.toindex, userindex, 0, "||177")

If UserList(userindex).Faccion.RecibioExpInicialReal = 0 Then
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpAlUnirse
    If UserList(userindex).Stats.Exp > MAXEXP Then _
        UserList(userindex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.toindex, userindex, 0, "||170@" & ExpAlUnirse)
    UserList(userindex).Faccion.RecibioExpInicialReal = 1
    Call CheckUserLevel(userindex)
End If

End Sub
Public Sub RecompensaArmadaReal(ByVal userindex As Integer, Optional Quinta As Boolean = False)
Dim MiObj As Obj
Dim ElYegua As Boolean
    MiObj.Amount = 1
    ElYegua = False
    
If UserList(userindex).Faccion.RecompensasReal = 5 Then
    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Ya eres 5ta jerarquia!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
Exit Sub
End If


If UserList(userindex).Faccion.RecompensasReal = 0 And Quinta = False Then
    If UserList(userindex).Faccion.CriminalesMatados >= FragsJerarquia(1) Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
            UserList(userindex).flags.PJerarquia = 1
            UserList(userindex).flags.SJerarquia = 0
            UserList(userindex).flags.TJerarquia = 0
            UserList(userindex).flags.CJerarquia = 0
            UserList(userindex).Faccion.RecompensasReal = 1
            ElYegua = True
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Debes matar " & FragsJerarquia(1) & " criminales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    End If
    
  ElseIf UserList(userindex).Faccion.RecompensasReal = 1 And Quinta = False Then
    If UserList(userindex).Faccion.CriminalesMatados >= FragsJerarquia(2) Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
            UserList(userindex).Faccion.RecompensasReal = 2
            UserList(userindex).flags.PJerarquia = 0
            UserList(userindex).flags.SJerarquia = 1
            UserList(userindex).flags.TJerarquia = 0
            UserList(userindex).flags.CJerarquia = 0
            ElYegua = True
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Debes matar " & FragsJerarquia(2) & " criminales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    End If
    
  ElseIf UserList(userindex).Faccion.RecompensasReal = 2 And Quinta = False Then
    If UserList(userindex).Faccion.CriminalesMatados >= FragsJerarquia(3) Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
            UserList(userindex).Faccion.RecompensasReal = 3
            UserList(userindex).flags.PJerarquia = 0
            UserList(userindex).flags.SJerarquia = 0
            UserList(userindex).flags.TJerarquia = 1
            UserList(userindex).flags.CJerarquia = 0
            ElYegua = True
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Debes matar " & FragsJerarquia(3) & " criminales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    End If
    
  ElseIf UserList(userindex).Faccion.RecompensasReal = 3 And Quinta = False Then
    If UserList(userindex).flags.CJerarquia = 1 Then Exit Sub
    If UserList(userindex).Faccion.CriminalesMatados >= FragsJerarquia(4) Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
            UserList(userindex).Faccion.RecompensasReal = 4
            UserList(userindex).flags.PJerarquia = 0
            UserList(userindex).flags.SJerarquia = 0
            UserList(userindex).flags.TJerarquia = 0
            UserList(userindex).flags.CJerarquia = 1
            Call SendData(SendTarget.ToAll, 0, 0, "||178@" & UserList(userindex).Name)
            ElYegua = True
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Debes matar " & FragsJerarquia(4) & " criminales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    End If
    
  ElseIf UserList(userindex).Faccion.RecompensasReal = 4 And Quinta = True Then
  
    If Not TieneObjetos(1220, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    ElseIf Not TieneObjetos(1221, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    ElseIf Not TieneObjetos(1222, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    ElseIf Not TieneObjetos(1223, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    ElseIf Not TieneObjetos(1224, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    End If
  
        ElYegua = True
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            UserList(userindex).Faccion.RecompensasReal = 5
            Call SendData(SendTarget.toindex, userindex, 0, "||133")
            Call SendData(SendTarget.ToAll, 0, 0, "||180@" & UserList(userindex).Name)
End If
                  
                  
If ElYegua = False Then Exit Sub
           
Select Case UserList(userindex).Faccion.RecompensasReal
    Case 1

    If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
        
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 951
        Else
            MiObj.ObjIndex = 957
        End If
        
     Else
          
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 950
        Else
            MiObj.ObjIndex = 956
        End If
        
    End If
     
     Case 2
    If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
        
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 953
        Else
            MiObj.ObjIndex = 959
        End If
        
     Else
          
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 952
        Else
            MiObj.ObjIndex = 958
        End If
        
    End If
    
     Case 3
    If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
        
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 955
        Else
            MiObj.ObjIndex = 961
        End If
        
     Else
          
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 954
        Else
            MiObj.ObjIndex = 960
        End If
        
    End If
        
  Case 4
    Dim HechizoIndex As Integer
    HechizoIndex = 0
    
        If UCase$(UserList(userindex).clase) = "MAGO" Then
            HechizoIndex = 60
        ElseIf UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            HechizoIndex = 63
        ElseIf UCase$(UserList(userindex).clase) = "ASESINO" Then
            HechizoIndex = 64
        ElseIf UCase$(UserList(userindex).clase) = "PALADIN" Then
            HechizoIndex = 61
        ElseIf UCase$(UserList(userindex).clase) = "CLERIGO" Then
            HechizoIndex = 62
        End If
        
        If HechizoIndex <> 0 Then
            If Not TieneHechizo(HechizoIndex, userindex) Then
                'Buscamos un slot vacio
                For j = 1 To MAXUSERHECHIZOS
                    If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
                Next j
                    
                If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||181")
                Else
                    UserList(userindex).Stats.UserHechizos(j) = HechizoIndex
                    Call UpdateUserHechizos(False, userindex, CByte(j))
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||182")
            End If
        End If
        
        
            If UCase$(UserList(userindex).clase) = "GUERRERO" Or UCase$(UserList(userindex).clase) = "CAZADOR" Then
                MiObj.ObjIndex = 1219
                
                MiObj.Amount = 1

                If UserList(userindex).Faccion.RecompensasCaos < 4 Then
                    If Not MeterItemEnInventario(userindex, MiObj) Then
                                Call TirarItemAlPiso(UserList(userindex).Pos, MiObj)
                    End If
                End If
                
            End If
        Exit Sub
    
  Case 5
  If Quinta = True Then
    If Not TieneHechizo(54, userindex) Then
        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS
            If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
        Next j
            
        If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||181")
        Else
            UserList(userindex).Stats.UserHechizos(j) = 54
            Call UpdateUserHechizos(False, userindex, CByte(j))
        End If
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||182")
    End If
    Call QuitarObjetos(1220, 20, userindex)
    Call QuitarObjetos(1221, 20, userindex)
    Call QuitarObjetos(1222, 20, userindex)
    Call QuitarObjetos(1223, 20, userindex)
    Call QuitarObjetos(1224, 20, userindex)
End If
    
End Select

MiObj.Amount = 1
 
If UserList(userindex).Faccion.RecompensasReal < 4 Then
    If Not MeterItemEnInventario(userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(userindex).Pos, MiObj)
    End If
End If

    Call CheckUserLevel(userindex)

End Sub

Public Sub ExpulsarFaccionReal(ByVal userindex As Integer)

    UserList(userindex).Faccion.ArmadaReal = 0
    UserList(userindex).flags.PJerarquia = 0
    UserList(userindex).flags.SJerarquia = 0
    UserList(userindex).flags.TJerarquia = 0
    UserList(userindex).flags.CJerarquia = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.toindex, userindex, 0, "||183")
    'Desequipamos la armadura real si está equipada
    If ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal userindex As Integer)

    UserList(userindex).Faccion.FuerzasCaos = 0
    UserList(userindex).flags.PJerarquia = 0
    UserList(userindex).flags.SJerarquia = 0
    UserList(userindex).flags.TJerarquia = 0
    UserList(userindex).flags.CJerarquia = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.toindex, userindex, 0, "||184")
    'Desequipamos la armadura real si está equipada
    If ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
End Sub

Public Function TituloReal(ByVal userindex As Integer) As String
 
Select Case UserList(userindex).Faccion.RecompensasReal
    Case 0
        TituloReal = "Servidor del Rey"
    Case 1
        TituloReal = "Servidor del Rey"
    Case 2
        TituloReal = "Soldado Imperial"
    Case 3
        TituloReal = "Protector del Imperio"
    Case 4
        TituloReal = "Maestro de la Luz"
    Case 5
        TituloReal = "Caballero de la Luz"
End Select
 
End Function

Public Sub EnlistarCaos(ByVal userindex As Integer)


If UserList(userindex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Ya perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Las sombras reinaran en este mundo, largate de aqui ciudadano.!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.RecibioExpInicialReal = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "No permitiré que ningún insecto real ingrese ¡Traidor del Rey!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If Not Criminal(userindex) Then
    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Jajaja tu no eres bienvenido aqui!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.CiudadanosMatados < FragsJerarquia(1) Then
    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos " & FragsJerarquia(1) & " ciudadanos, solo has matado " & UserList(userindex).Faccion.CiudadanosMatados & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.Reenlistadas = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||176")
Exit Sub
End If

UserList(userindex).Faccion.Reenlistadas = 1
UserList(userindex).Faccion.FuerzasCaos = 1
UserList(userindex).Faccion.RecompensasCaos = 0
UserList(userindex).StatusMith.EsStatus = 4
UserList(userindex).flags.PJerarquia = 1
UserList(userindex).flags.SJerarquia = 0
UserList(userindex).flags.TJerarquia = 0
UserList(userindex).flags.CJerarquia = 0
Call SendUserStatux(userindex)

Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Bienvenido a la horda infernal!!!, para recibir tu armadura escribe /recompensa!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))

If UserList(userindex).Faccion.RecibioExpInicialCaos = 0 Then
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpAlUnirse
    If UserList(userindex).Stats.Exp > MAXEXP Then _
        UserList(userindex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.toindex, userindex, 0, "||170@" & ExpAlUnirse)
    UserList(userindex).Faccion.RecibioExpInicialCaos = 1
    Call CheckUserLevel(userindex)
End If

End Sub
Public Sub RecompensaCaos(ByVal userindex As Integer, Optional Quinta As Boolean = False)
Dim MiObj As Obj
Dim ElYegua As Boolean
    MiObj.Amount = 1
    
    ElYegua = False
    
If UserList(userindex).Faccion.RecompensasCaos = 5 Then
    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Ya eres 5ta jerarquia!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
Exit Sub
End If

If UserList(userindex).Faccion.RecompensasCaos = 0 And Quinta = False Then
    If UserList(userindex).Faccion.CiudadanosMatados >= FragsJerarquia(1) Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
            UserList(userindex).flags.PJerarquia = 1
            UserList(userindex).flags.SJerarquia = 0
            UserList(userindex).flags.TJerarquia = 0
            UserList(userindex).flags.CJerarquia = 0
            UserList(userindex).Faccion.RecompensasCaos = 1
            ElYegua = True
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Debes matar " & FragsJerarquia(1) & " criminales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    End If
    
  ElseIf UserList(userindex).Faccion.RecompensasCaos = 1 And Quinta = False Then
    If UserList(userindex).Faccion.CiudadanosMatados >= FragsJerarquia(2) Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
            UserList(userindex).Faccion.RecompensasCaos = 2
            UserList(userindex).flags.PJerarquia = 0
            UserList(userindex).flags.SJerarquia = 1
            UserList(userindex).flags.TJerarquia = 0
            UserList(userindex).flags.CJerarquia = 0
            ElYegua = True
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Debes matar " & FragsJerarquia(2) & " criminales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    End If
    
  ElseIf UserList(userindex).Faccion.RecompensasCaos = 2 And Quinta = False Then
    If UserList(userindex).Faccion.CiudadanosMatados >= FragsJerarquia(3) Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
            UserList(userindex).Faccion.RecompensasCaos = 3
            UserList(userindex).flags.PJerarquia = 0
            UserList(userindex).flags.SJerarquia = 0
            UserList(userindex).flags.TJerarquia = 1
            UserList(userindex).flags.CJerarquia = 0
            ElYegua = True
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Debes matar " & FragsJerarquia(3) & " criminales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    End If
    
  ElseIf UserList(userindex).Faccion.RecompensasCaos = 3 And Quinta = False Then
    If UserList(userindex).flags.CJerarquia = 1 Then Exit Sub
    If UserList(userindex).Faccion.CiudadanosMatados >= FragsJerarquia(4) Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
            UserList(userindex).Faccion.RecompensasCaos = 4
            UserList(userindex).flags.PJerarquia = 0
            UserList(userindex).flags.SJerarquia = 0
            UserList(userindex).flags.TJerarquia = 0
            UserList(userindex).flags.CJerarquia = 1
            ElYegua = True
            Call SendData(SendTarget.ToAll, 0, 0, "||851@" & UserList(userindex).Name)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Debes matar " & FragsJerarquia(4) & " criminales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    End If
    
ElseIf UserList(userindex).Faccion.RecompensasCaos = 4 And Quinta = True Then
  
    If Not TieneObjetos(1220, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    ElseIf Not TieneObjetos(1221, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    ElseIf Not TieneObjetos(1222, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    ElseIf Not TieneObjetos(1223, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    ElseIf Not TieneObjetos(1224, 20, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||179")
        Exit Sub
    End If
  
        ElYegua = True
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        UserList(userindex).Faccion.RecompensasCaos = 5
        Call SendData(SendTarget.toindex, userindex, 0, "||133")
        Call SendData(SendTarget.ToAll, 0, 0, "||852@" & UserList(userindex).Name)
End If
    
If ElYegua = False Then Exit Sub

Select Case UserList(userindex).Faccion.RecompensasCaos
    
    Case 1

    If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
        
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 981
        Else
            MiObj.ObjIndex = 987
        End If
        
     Else
          
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 980
        Else
            MiObj.ObjIndex = 986
        End If
        
    End If
     
     Case 2
     
    If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
        
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 983
        Else
            MiObj.ObjIndex = 989
        End If
        
     Else
          
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 982
        Else
            MiObj.ObjIndex = 988
        End If
        
    End If
    
     Case 3
     
    If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
        
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 985
        Else
            MiObj.ObjIndex = 991
        End If
        
     Else
          
        If UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            MiObj.ObjIndex = 984
        Else
            MiObj.ObjIndex = 990
        End If
        
    End If
        
  Case 4
    Dim HechizoIndex As Integer
    HechizoIndex = 0
    
        If UCase$(UserList(userindex).clase) = "MAGO" Then
            HechizoIndex = 60
        ElseIf UCase$(UserList(userindex).clase) = "DRUIDA" Or UCase$(UserList(userindex).clase) = "BARDO" Then
            HechizoIndex = 63
        ElseIf UCase$(UserList(userindex).clase) = "ASESINO" Then
            HechizoIndex = 64
        ElseIf UCase$(UserList(userindex).clase) = "PALADIN" Then
            HechizoIndex = 61
        ElseIf UCase$(UserList(userindex).clase) = "CLERIGO" Then
            HechizoIndex = 62
        End If
        
        If HechizoIndex <> 0 Then
            If Not TieneHechizo(HechizoIndex, userindex) Then
                'Buscamos un slot vacio
                For j = 1 To MAXUSERHECHIZOS
                    If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
                Next j
                    
                If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||181")
                Else
                    UserList(userindex).Stats.UserHechizos(j) = HechizoIndex
                    Call UpdateUserHechizos(False, userindex, CByte(j))
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||182")
            End If
        End If
        
        
            If UCase$(UserList(userindex).clase) = "GUERRERO" Or UCase$(UserList(userindex).clase) = "CAZADOR" Then
                MiObj.ObjIndex = 1219
                
                MiObj.Amount = 1

                If UserList(userindex).Faccion.RecompensasCaos < 4 Then
                    If Not MeterItemEnInventario(userindex, MiObj) Then
                                Call TirarItemAlPiso(UserList(userindex).Pos, MiObj)
                    End If
                End If
                
            End If
        Exit Sub
    
  Case 5
  
  If Quinta = True Then
    If Not TieneHechizo(55, userindex) Then
        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS
            If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
        Next j
            
        If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||181")
        Else
            UserList(userindex).Stats.UserHechizos(j) = 55
            Call UpdateUserHechizos(False, userindex, CByte(j))
        End If
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||182")
    End If
    
    Call QuitarObjetos(1220, 20, userindex)
    Call QuitarObjetos(1221, 20, userindex)
    Call QuitarObjetos(1222, 20, userindex)
    Call QuitarObjetos(1223, 20, userindex)
    Call QuitarObjetos(1224, 20, userindex)
    Exit Sub
 End If
            
End Select

MiObj.Amount = 1

    If UserList(userindex).Faccion.RecompensasCaos < 4 Then
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(userindex).Pos, MiObj)
        End If
    End If
    
    Call CheckUserLevel(userindex)

End Sub

Public Function TituloCaos(ByVal userindex As Integer) As String
Select Case UserList(userindex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Servidor del Demonio"
    Case 1
        TituloCaos = "Servidor del Demonio"
    Case 2
        TituloCaos = "Mercenario de la Oscuridad"
    Case 3
        TituloCaos = "General de los Infiernos"
    Case 4
        TituloCaos = "Maestro de la Oscuridad"
    Case 5
        TituloCaos = "Caballero de la Oscuridad"
End Select
 
 
End Function


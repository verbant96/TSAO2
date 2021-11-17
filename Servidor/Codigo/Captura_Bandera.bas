Attribute VB_Name = "Captura_Bandera"
Public Sub ComenzarClb()
If HayClb = True Then Exit Sub
SendData SendTarget.toall, 0, 0, "||605"
ResetClb
MinutosClb = 3
HayClb = True
End Sub

Public Sub FinalizarClb()
If HayClb = False Then Exit Sub
SendData SendTarget.toall, 0, 0, "||606"
ResetClb
MinutosClb = 0
HayClb = False
End Sub

Public Sub TiempoClb()
    If MinutosClb > 0 Then
    MinutosClb = MinutosClb - 1
        If MinutosClb = 0 Then
            If ParticipantesClb < 10 Then
                SendData SendTarget.toall, 0, 0, "||607"
                ResetClb
            Else
                SendData SendTarget.toall, 0, 0, "||608"
                SumonearClb
                BloquearEntradas
            End If
        Else
            SendData SendTarget.toall, 0, 0, "||605"
        End If
    End If
End Sub

Public Sub ResetClb()
HayClb = False
SegundosClb = 0
MinutosClb = 0
ParticipantesClb = 0
ComenzoClb = False
PuntoAzul = 0
PuntoRojo = 0


Dim i As Integer

For i = 1 To LastUser
If UserList(i).Pos.Map = MapaCaptura Then
Call WarpUserChar(i, 28, RandomNumber(50, 60), RandomNumber(36, 39), True)
End If
Next i

For i = 1 To 10
PuestoCaptura(i) = 0
Next i

For i = 1 To 20
PosRuna(i).Map = 0
PosRuna(i).X = 0
PosRuna(i).Y = 0
Next i

MapData(MapaCaptura, 20, 12).Blocked = 0
MapData(MapaCaptura, 50, 12).Blocked = 0
MapData(MapaCaptura, 80, 12).Blocked = 0

MapData(MapaCaptura, 20, 89).Blocked = 0
MapData(MapaCaptura, 50, 89).Blocked = 0
MapData(MapaCaptura, 80, 89).Blocked = 0

'LimpiarMapa 0, 166
End Sub

Public Sub SumonearClb()

Dim i As Integer

For i = 1 To 20
PosRuna(i).Map = MapaCaptura
PosRuna(i).X = RandomNumber(15, 85)
PosRuna(i).Y = RandomNumber(13, 88)
Next i

'Sumoneamos a lo negro
For i = 1 To 5
    Call WarpUserChar(PuestoCaptura(i), MapaCaptura, RandomNumber(44, 56), RandomNumber(25, 37), True)
    UserList(PuestoCaptura(i)).flags.EquipoCaptura = "Azul"
    UserList(PuestoCaptura(i)).flags.CvcBlue = 1
    UserList(PuestoCaptura(i)).flags.CvcRed = 0
Next i

For i = 6 To 10
    Call WarpUserChar(PuestoCaptura(i), MapaCaptura, RandomNumber(44, 56), RandomNumber(63, 75), True)
    UserList(PuestoCaptura(i)).flags.EquipoCaptura = "Rojo"
    UserList(PuestoCaptura(i)).flags.CvcBlue = 0
    UserList(PuestoCaptura(i)).flags.CvcRed = 1
Next i
'Sumoneamos a lo negro

Call SendData(SendTarget.ToMap, 0, 166, "!!" & "El objetivo del evento es llevar la bandera que aparece en el centro (o tenga encima un usuario) a las 3 bases que estan atras de la zona de spawn del equipo contrario." & ENDC)

SendData SendTarget.ToMap, 0, 166, "||609"
SendData SendTarget.ToMap, 0, 166, "||610@" & UserList(PuestoCaptura(1)).name & "@" & UserList(PuestoCaptura(2)).name & "@" & UserList(PuestoCaptura(3)).name & "@" & UserList(PuestoCaptura(4)).name & "@" & UserList(PuestoCaptura(5)).name

SendData SendTarget.ToMap, 0, 166, "||611"
SendData SendTarget.ToMap, 0, 166, "||610@" & UserList(PuestoCaptura(6)).name & "@" & UserList(PuestoCaptura(7)).name & "@" & UserList(PuestoCaptura(8)).name & "@" & UserList(PuestoCaptura(9)).name & "@" & UserList(PuestoCaptura(10)).name

SendData SendTarget.ToMap, 0, 166, "||612"

SegundosClb = 30

ComenzoClb = True

Dim Banderita As Obj
Banderita.Amount = 1
Banderita.ObjIndex = BanderaPiso
Call MakeObj(ToMap, 0, MapaCaptura, Banderita, MapaCaptura, 50, 50)


End Sub

Public Sub BuscarPuestoLibre(ByVal Userindex As Integer)

If PuestoCaptura(1) = 0 Then
PuestoCaptura(1) = Userindex
ElseIf PuestoCaptura(2) = 0 Then
PuestoCaptura(2) = Userindex
ElseIf PuestoCaptura(3) = 0 Then
PuestoCaptura(3) = Userindex
ElseIf PuestoCaptura(4) = 0 Then
PuestoCaptura(4) = Userindex
ElseIf PuestoCaptura(5) = 0 Then
PuestoCaptura(5) = Userindex
ElseIf PuestoCaptura(6) = 0 Then
PuestoCaptura(6) = Userindex
ElseIf PuestoCaptura(7) = 0 Then
PuestoCaptura(7) = Userindex
ElseIf PuestoCaptura(8) = 0 Then
PuestoCaptura(8) = Userindex
ElseIf PuestoCaptura(9) = 0 Then
PuestoCaptura(9) = Userindex
ElseIf PuestoCaptura(10) = 0 Then
PuestoCaptura(10) = Userindex
End If

End Sub


Public Sub BorrarPuesto(ByVal Userindex As Integer)

Dim i As Integer

For i = 1 To 10
If PuestoCaptura(i) = Userindex Then
PuestoCaptura(i) = 0
End If
Next i

End Sub


Public Sub BloquearEntradas()
'Equipo Azul
MapData(MapaCaptura, 48, 24).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 48, 24, 1)
MapData(MapaCaptura, 49, 24).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 49, 24, 1)
MapData(MapaCaptura, 50, 24).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 50, 24, 1)
MapData(MapaCaptura, 51, 24).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 51, 24, 1)
MapData(MapaCaptura, 52, 24).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 52, 24, 1)
MapData(MapaCaptura, 48, 38).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 48, 38, 1)
MapData(MapaCaptura, 49, 38).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 49, 38, 1)
MapData(MapaCaptura, 50, 38).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 50, 38, 1)
MapData(MapaCaptura, 51, 38).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 51, 38, 1)
MapData(MapaCaptura, 52, 38).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 52, 38, 1)
MapData(MapaCaptura, 43, 29).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 29, 1)
MapData(MapaCaptura, 43, 30).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 30, 1)
MapData(MapaCaptura, 43, 31).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 31, 1)
MapData(MapaCaptura, 43, 32).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 32, 1)
MapData(MapaCaptura, 43, 33).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 33, 1)
MapData(MapaCaptura, 57, 29).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 29, 1)
MapData(MapaCaptura, 57, 30).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 30, 1)
MapData(MapaCaptura, 57, 31).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 31, 1)
MapData(MapaCaptura, 57, 32).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 32, 1)
MapData(MapaCaptura, 57, 33).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 33, 1)
'Equipo Azul

'Equipo Rojo
MapData(MapaCaptura, 48, 62).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 48, 62, 1)
MapData(MapaCaptura, 49, 62).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 49, 62, 1)
MapData(MapaCaptura, 50, 62).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 50, 62, 1)
MapData(MapaCaptura, 51, 62).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 51, 62, 1)
MapData(MapaCaptura, 52, 62).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 52, 62, 1)
MapData(MapaCaptura, 48, 76).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 48, 76, 1)
MapData(MapaCaptura, 49, 76).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 49, 76, 1)
MapData(MapaCaptura, 50, 76).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 50, 76, 1)
MapData(MapaCaptura, 51, 76).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 51, 76, 1)
MapData(MapaCaptura, 52, 76).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 52, 76, 1)
MapData(MapaCaptura, 43, 67).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 67, 1)
MapData(MapaCaptura, 43, 68).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 68, 1)
MapData(MapaCaptura, 43, 69).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 69, 1)
MapData(MapaCaptura, 43, 70).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 70, 1)
MapData(MapaCaptura, 43, 71).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 71, 1)
MapData(MapaCaptura, 57, 67).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 67, 1)
MapData(MapaCaptura, 57, 68).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 68, 1)
MapData(MapaCaptura, 57, 69).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 69, 1)
MapData(MapaCaptura, 57, 70).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 70, 1)
MapData(MapaCaptura, 57, 71).Blocked = 1
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 71, 1)
'Equipo Rojo

End Sub

Public Sub DesbloquearEntradas()

'Equipo Azul
MapData(MapaCaptura, 48, 24).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 48, 24, 0)
MapData(MapaCaptura, 49, 24).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 49, 24, 0)
MapData(MapaCaptura, 50, 24).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 50, 24, 0)
MapData(MapaCaptura, 51, 24).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 51, 24, 0)
MapData(MapaCaptura, 52, 24).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 52, 24, 0)
MapData(MapaCaptura, 48, 38).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 48, 38, 0)
MapData(MapaCaptura, 49, 38).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 49, 38, 0)
MapData(MapaCaptura, 50, 38).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 50, 38, 0)
MapData(MapaCaptura, 51, 38).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 51, 38, 0)
MapData(MapaCaptura, 52, 38).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 52, 38, 0)
MapData(MapaCaptura, 43, 29).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 29, 0)
MapData(MapaCaptura, 43, 30).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 30, 0)
MapData(MapaCaptura, 43, 31).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 31, 0)
MapData(MapaCaptura, 43, 32).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 32, 0)
MapData(MapaCaptura, 43, 33).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 33, 0)
MapData(MapaCaptura, 57, 29).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 29, 0)
MapData(MapaCaptura, 57, 30).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 30, 0)
MapData(MapaCaptura, 57, 31).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 31, 0)
MapData(MapaCaptura, 57, 32).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 32, 0)
MapData(MapaCaptura, 57, 33).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 33, 0)
'Equipo Azul

'Equipo Rojo
MapData(MapaCaptura, 48, 62).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 48, 62, 0)
MapData(MapaCaptura, 49, 62).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 49, 62, 0)
MapData(MapaCaptura, 50, 62).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 50, 62, 0)
MapData(MapaCaptura, 51, 62).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 51, 62, 0)
MapData(MapaCaptura, 52, 62).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 52, 62, 0)
MapData(MapaCaptura, 48, 76).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 48, 76, 0)
MapData(MapaCaptura, 49, 76).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 49, 76, 0)
MapData(MapaCaptura, 50, 76).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 50, 76, 0)
MapData(MapaCaptura, 51, 76).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 51, 76, 0)
MapData(MapaCaptura, 52, 76).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 52, 76, 0)
MapData(MapaCaptura, 43, 67).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 67, 0)
MapData(MapaCaptura, 43, 68).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 68, 0)
MapData(MapaCaptura, 43, 69).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 69, 0)
MapData(MapaCaptura, 43, 70).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 70, 0)
MapData(MapaCaptura, 43, 71).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 43, 71, 0)
MapData(MapaCaptura, 57, 67).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 67, 0)
MapData(MapaCaptura, 57, 68).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 68, 0)
MapData(MapaCaptura, 57, 69).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 69, 0)
MapData(MapaCaptura, 57, 70).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 70, 0)
MapData(MapaCaptura, 57, 71).Blocked = 0
Call Bloquear(SendTarget.ToMap, 0, MapaCaptura, MapaCaptura, 57, 71, 0)
'Equipo Rojo

End Sub

Public Sub IngresarClb(ByVal Userindex As Integer)

If UserList(Userindex).flags.EnClb = 1 Then Exit Sub

    If HayClb = False Then
            SendData SendTarget.toindex, Userindex, 0, "||613"
        Exit Sub
    ElseIf ParticipantesClb >= 10 Then
            SendData SendTarget.toindex, Userindex, 0, "||614"
        Exit Sub
    ElseIf ComenzoClb = True Then
            SendData SendTarget.toindex, Userindex, 0, "||615"
        Exit Sub
    End If
    
    SendData SendTarget.toindex, Userindex, 0, "||616"
    ParticipantesClb = ParticipantesClb + 1
    BuscarPuestoLibre Userindex
    UserList(Userindex).flags.EnClb = 1
End Sub

Public Sub CorroborarPunto(ByVal Userindex As Integer, EquipoCaptura As String)

Dim i As Integer, PuedeMorir As Byte

PuedeMorir = 0

For i = 1 To 20
    If UserList(Userindex).Pos.Map = PosRuna(i).Map And UserList(Userindex).Pos.X = PosRuna(i).X And UserList(Userindex).Pos.Y = PosRuna(i).Y Then
        PuedeMorir = 1
    End If
Next i

If PuedeMorir = 1 Then
    SendData SendTarget.toindex, Userindex, 0, "||617"
    Call MurioCapturaBomba(Userindex)
End If


If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).OBJInfo.ObjIndex = BanderaPiso Then

    If UserList(Userindex).Pos.X <> 20 And UserList(Userindex).Pos.Y <> 89 Or UserList(Userindex).Pos.X <> 50 And UserList(Userindex).Pos.Y <> 89 Or UserList(Userindex).Pos.X <> 80 And UserList(Userindex).Pos.Y <> 89 Or UserList(Userindex).Pos.X <> 20 And UserList(Userindex).Pos.Y <> 12 Or UserList(Userindex).Pos.X <> 50 And UserList(Userindex).Pos.Y <> 12 Or UserList(Userindex).Pos.X <> 80 And UserList(Userindex).Pos.Y <> 12 Then
    
        Dim Banderita As Obj
        Banderita.Amount = 1
        Banderita.ObjIndex = Bandera
        
        If Not MeterItemEnInventario(Userindex, Banderita) Then
            SendData SendTarget.toindex, Userindex, 0, "||618"
        Exit Sub
        End If
        
            If UserList(Userindex).flags.EquipoCaptura = "Rojo" Then
                SendData SendTarget.ToMap, 0, 166, "||619@" & UserList(Userindex).name & "@" & UserList(Userindex).Pos.X & "@" & UserList(Userindex).Pos.Y
            Else
                SendData SendTarget.ToMap, 0, 166, "||620@" & UserList(Userindex).name & "@" & UserList(Userindex).Pos.X & "@" & UserList(Userindex).Pos.Y
            End If
        
            Call EraseObj(SendTarget.ToMap, Userindex, UserList(Userindex).Pos.Map, 10000, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
    
    End If

End If

    If EquipoCaptura = "Rojo" Then
            If UserList(Userindex).Pos.X = 20 And UserList(Userindex).Pos.Y = 12 Or UserList(Userindex).Pos.X = 50 And UserList(Userindex).Pos.Y = 12 Or UserList(Userindex).Pos.X = 80 And UserList(Userindex).Pos.Y = 12 Then
                If TieneObjetos(Bandera, 1, Userindex) = False Then
                    SendData SendTarget.toindex, Userindex, 0, "||621"
                    Call WarpUserChar(Userindex, MapaCaptura, 50, 69)
                Else
                    PuntoRojo = PuntoRojo + 1
                If PuntoRojo >= 3 Then
                    Call QuitarObjetos(Bandera, 1, Userindex)
                    Call GanadorCaptura("Rojo")
                Else
                    SendData SendTarget.ToMap, 0, 166, "||622"
                    PonerBandera UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y
                    Call WarpUserChar(Userindex, MapaCaptura, 50, 69, False)
                    Call QuitarObjetos(Bandera, 1, Userindex)
                
                    Banderita.Amount = 1
                    Banderita.ObjIndex = BanderaPiso
                    Call MakeObj(ToMap, 0, MapaCaptura, Banderita, MapaCaptura, 50, 50)
                End If
            End If
        End If
    End If

If EquipoCaptura = "Azul" Then
    If UserList(Userindex).Pos.X = 20 And UserList(Userindex).Pos.Y = 89 Or UserList(Userindex).Pos.X = 50 And UserList(Userindex).Pos.Y = 89 Or UserList(Userindex).Pos.X = 80 And UserList(Userindex).Pos.Y = 89 Then
        If TieneObjetos(Bandera, 1, Userindex) = False Then
            SendData SendTarget.toindex, Userindex, 0, "||621"
            Call WarpUserChar(Userindex, MapaCaptura, 50, 31)
        Else
            PuntoAzul = PuntoAzul + 1
            
            If PuntoAzul >= 3 Then
                Call QuitarObjetos(Bandera, 1, Userindex)
                Call GanadorCaptura("Azul")
            Else
                SendData SendTarget.ToMap, 0, 166, "||623"
                PonerBandera UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y
                Call WarpUserChar(Userindex, MapaCaptura, 50, 31, False)
                Call QuitarObjetos(Bandera, 1, Userindex)
                Banderita.Amount = 1
                Banderita.ObjIndex = BanderaPiso
                Call MakeObj(ToMap, 0, MapaCaptura, Banderita, MapaCaptura, 50, 50)
            End If
        End If
    End If
End If

End Sub

Public Sub MurioCaptura(ByVal Userindex As Integer)

If TieneObjetos(Bandera, 1, Userindex) = True Then
    SendData SendTarget.ToMap, 0, 166, "||624"
    Call QuitarObjetos(Bandera, 1, Userindex)
    Dim Banderita As Obj
    Banderita.Amount = 1
    Banderita.ObjIndex = BanderaPiso
    Call MakeObj(ToMap, 0, MapaCaptura, Banderita, MapaCaptura, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
End If

SendData SendTarget.toindex, Userindex, 0, "||625"
UserList(Userindex).Counters.SegundosRevivirCLB = 5

End Sub


Public Sub MurioCapturaBomba(ByVal Userindex As Integer)
If TieneObjetos(Bandera, 1, Userindex) = True Then
SendData SendTarget.ToMap, 0, 166, "||626"
Call QuitarObjetos(Bandera, 1, Userindex)
Dim Banderita As Obj
Banderita.Amount = 1
Banderita.ObjIndex = BanderaPiso
Call MakeObj(ToMap, 0, MapaCaptura, Banderita, MapaCaptura, 50, 50)
End If


Call UserDie(Userindex)

End Sub
Public Sub GanadorCaptura(ByVal equipo As String)

Dim Ganadores As Integer, LlevarTodos As Integer

If equipo = "Rojo" Then
    SendData SendTarget.toall, 0, 0, "||627"
    
    For LlevarTodos = 6 To 10
        Call AgregarPuntos(PuestoCaptura(LlevarTodos), 20)
        SendData SendTarget.toindex, PuestoCaptura(LlevarTodos), 0, "||57@20"
    Next LlevarTodos
End If

If equipo = "Azul" Then
    SendData SendTarget.toall, 0, 0, "||628"
    
    For LlevarTodos = 1 To 5
        Call AgregarPuntos(PuestoCaptura(LlevarTodos), 20)
        SendData SendTarget.toindex, PuestoCaptura(LlevarTodos), 0, "||57@20"
    Next LlevarTodos
End If

For LlevarTodos = 1 To 10
    Call WarpUserChar(PuestoCaptura(i), 28, RandomNumber(50, 60), RandomNumber(36, 39), True)
    UserList(PuestoCaptura(i)).flags.EnClb = 0
    UserList(PuestoCaptura(i)).flags.CvcBlue = 0
    UserList(PuestoCaptura(i)).flags.CvcRed = 0
Next LlevarTodos

ResetClb
MinutosAbrirClb = 60
End Sub

Public Sub PonerBandera(ByVal X As Integer, Y As Integer)

Dim Banderita As Obj
Banderita.Amount = 1
Banderita.ObjIndex = BanderaPiso
Call MakeObj(ToMap, 0, MapaCaptura, Banderita, MapaCaptura, X, Y)
MapData(MapaCaptura, X, Y).Blocked = 1

End Sub

Public Sub MinutosParaAbrirClb()
    If MinutosAbrirClb > 0 Then
        MinutosAbrirClb = MinutosAbrirClb - 1
        If MinutosAbrirClb = 0 Then
            ComenzarClb
        End If
    End If
End Sub

Attribute VB_Name = "DatosUser"
'Este Public Sub se encarga de actualizar todos los datos del usuario
'Separé el paquete de estadisticas en unos 10 o 11 paquetes
'Agregué un Public Sub para actualizar todos los stats a la vez para utilizarlo en el ConnectUser.
'Meti un Public Sub para actualizar el color del nick del usuario
'Actualizamos las auras desde acá
Option Explicit
Public Sub SendUserHP(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[H]" & UserList(userindex).Stats.MaxHP & "," & UserList(userindex).Stats.MinHP
    
    '/Estamos revisando a este usuario? si es así, le enviamos la data a los administradores.
    '/Falta separar el case CHX.
    If RevisandoUsuario = True Then
        If UsuarioRevisado = userindex Then Call ActualizarChori(userindex)
    End If
End Sub
Public Sub SendUserMP(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[M]" & UserList(userindex).Stats.MaxMAN & "," & UserList(userindex).Stats.MinMAN
    
    '/Estamos revisando a este usuario? si es así, le enviamos la data a los administradores.
    '/Falta separar el case CHX.
    If RevisandoUsuario = True Then
        If UsuarioRevisado = userindex Then Call ActualizarChori(userindex)
    End If
End Sub
Public Sub SendUserST(ByVal userindex As Integer)
    If UserList(userindex).Stats.MinSta > UserList(userindex).Stats.MaxSta Then UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
    SendData SendTarget.toindex, userindex, 0, "[S]" & UserList(userindex).Stats.MaxSta & "," & UserList(userindex).Stats.MinSta
End Sub
Public Sub SendUserGLD(ByVal userindex As Integer)
    If UserList(userindex).Stats.GLD > 999999999 Then UserList(userindex).Stats.GLD = 999999999
    SendData SendTarget.toindex, userindex, 0, "[G]" & UserList(userindex).Stats.GLD
End Sub
Public Sub SendUserLVL(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[L]" & UserList(userindex).Stats.ELV
End Sub
Public Sub SendUserEXP(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[E]" & UserList(userindex).Stats.ELU & "," & UserList(userindex).Stats.Exp
End Sub
Public Sub SendUserBANK(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[B]" & UserList(userindex).Stats.Banco
End Sub
Public Sub SendUserNick(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[N]" & UserList(userindex).Name
End Sub
Public Sub SendUserAgilidad(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[A]" & UserList(userindex).Stats.UserAtributos(Agilidad)
End Sub
Public Sub SendUserFuerza(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[F]" & UserList(userindex).Stats.UserAtributos(Fuerza)
End Sub
Public Sub SendUserMontVol(ByVal userindex As Integer)
    SendData SendTarget.toMap, 0, UserList(userindex).Pos.Map, "MVOL" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).flags.levitando
End Sub
Public Sub SendCharData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer)

    Dim Color As Byte
    Color = 0
        
        With UserList(userindex)
            If (.flags.enBatalla) Then
                If (.flags.teamNumber = 1) Then
                    Color = 49
                ElseIf (.flags.teamNumber = 2) Then
                    Color = 30
                ElseIf (.flags.teamNumber = 3) Then
                    Color = 50
                ElseIf (.flags.teamNumber = 4) Then
                    Color = 31
                End If
            Else
                If (UserList(userindex).flags.estado = 1) Then
                    Color = 40
                ElseIf (UserList(userindex).flags.EsPremium = 1) Then
                    Color = 41
                ElseIf (UserList(userindex).flags.GranPoder = 1) Then
                    Color = 42
                ElseIf (UserList(userindex).flags.CvcBlue = 1) Or (UserList(userindex).flags.CastiBlue = 1) Or (UserList(userindex).flags.AramAzul) Then
                    Color = 49
                ElseIf (UserList(userindex).flags.CvcRed = 1) Or (UserList(userindex).flags.CastiRed = 1) Or (UserList(userindex).flags.AramRojo) Then
                    Color = 50
                End If
            End If
        End With
        
        Call SendData(sndRoute, sndIndex, sndMap, "[CD" & UserList(userindex).Char.CharIndex & "," & Color & "," & _
            UserList(userindex).Char.AuraA & "," & UserList(userindex).Char.AuraW & "," & UserList(userindex).Char.AuraE & "," & UserList(userindex).Char.AuraR & "," & UserList(userindex).Char.AuraC & "," & _
            UserList(userindex).flags.levitando & "," & Mod_Ranking.tieneRanking(userindex))
End Sub
Public Sub SendUserReputacion(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[R]" & UserList(userindex).Stats.Reputacione
End Sub
Public Sub ActualizarSlot(ByVal userindex As Integer, ByVal slot As Byte)

   If UserList(userindex).Invent.Object(slot).Amount > 0 Then
    SendData SendTarget.toindex, userindex, 0, "|S1" & slot & "," & UserList(userindex).Invent.Object(slot).Amount
   Else
    Call UpdateUserInv(False, userindex, slot)
   End If

End Sub
Public Sub ActualizarSlotEquipped(ByVal userindex As Integer, ByVal slot As Byte)

    SendData SendTarget.toindex, userindex, 0, "|S2" & slot & "," & UserList(userindex).Invent.Object(slot).Equipped

End Sub
Public Sub ActualizarChori(ByVal userindex As Integer)
    Call SendData(SendTarget.ToAdmins, 0, 0, "CHX" & UserList(UsuarioRevisado).Stats.MaxHP & "," & UserList(UsuarioRevisado).Stats.MinHP & "," & UserList(UsuarioRevisado).Stats.MaxMAN & "," & UserList(UsuarioRevisado).Stats.MinMAN & "," & UserList(UsuarioRevisado).Name)
End Sub
Public Sub SendUserData(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[EZ" & UserList(userindex).Stats.MaxHP & "," & UserList(userindex).Stats.MinHP & "," & UserList(userindex).Stats.MaxMAN & "," & UserList(userindex).Stats.MinMAN & "," & UserList(userindex).Stats.MaxSta & "," & UserList(userindex).Stats.MinSta
End Sub
Public Sub SendUserStats(ByVal userindex As Integer)
    SendData SendTarget.toindex, userindex, 0, "[ES" & UserList(userindex).Stats.MaxHP & "," & UserList(userindex).Stats.MinHP & "," & UserList(userindex).Stats.MaxMAN & "," & UserList(userindex).Stats.MinMAN & "," & UserList(userindex).Stats.MaxSta & "," & UserList(userindex).Stats.MinSta & "," & _
    UserList(userindex).Stats.GLD & "," & UserList(userindex).Stats.ELV & "," & UserList(userindex).Stats.ELU & "," & UserList(userindex).Stats.Exp & "," & UserList(userindex).Name & "," & UserList(userindex).Stats.UserAtributos(Agilidad) & "," & UserList(userindex).Stats.UserAtributos(Fuerza) & "," & _
    UserList(userindex).Stats.Reputacione
End Sub
Public Sub SendUserVariant(ByVal userindex As Integer)

Dim Color As Byte
Color = 0
        
        With UserList(userindex)
            If (.flags.enBatalla) Then
                If (.flags.teamNumber = 1) Then
                    Color = 49
                ElseIf (.flags.teamNumber = 2) Then
                    Color = 30
                ElseIf (.flags.teamNumber = 3) Then
                    Color = 50
                ElseIf (.flags.teamNumber = 4) Then
                    Color = 31
                End If
            Else
                If (UserList(userindex).flags.estado = 1) Then
                    Color = 40
                ElseIf (UserList(userindex).flags.EsPremium = 1) Then
                    Color = 41
                ElseIf (UserList(userindex).flags.GranPoder = 1) Then
                    Color = 42
                ElseIf (UserList(userindex).flags.CvcBlue = 1) Or (UserList(userindex).flags.CastiBlue = 1) Or (UserList(userindex).flags.AramAzul) Then
                    Color = 49
                ElseIf (UserList(userindex).flags.CvcRed = 1) Or (UserList(userindex).flags.CastiRed = 1) Or (UserList(userindex).flags.AramRojo) Then
                    Color = 50
                End If
            End If
        End With
        
        Call SendData(SendTarget.toMap, 0, UserList(userindex).Pos.Map, "XC" & UserList(userindex).Char.CharIndex & "," & Color)
End Sub
Public Sub SendUserAura(ByVal userindex As Integer)
    'Call SendData(SendTarget.toindex, userindex, 0, "AU|" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.AuraA & "," & UserList(userindex).Char.AuraW & "," & UserList(userindex).Char.AuraE & "," & UserList(userindex).Char.AuraR & "," & UserList(userindex).Char.AuraC)
    'Call SendToUserArea(userindex, "AU|" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.AuraA & "," & UserList(userindex).Char.AuraW & "," & UserList(userindex).Char.AuraE & "," & UserList(userindex).Char.AuraR & "," & UserList(userindex).Char.AuraC)
    SendToUserArea userindex, "AU|" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.AuraA & "," & UserList(userindex).Char.AuraW & "," & UserList(userindex).Char.AuraE & "," & UserList(userindex).Char.AuraR & "," & UserList(userindex).Char.AuraC
End Sub
Sub ChangeUserHeading(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal Heading As Byte)

    UserList(userindex).Char.Heading = Heading

    If sndRoute = SendTarget.toMap Then
        Call SendToUserArea(userindex, "|H" & UserList(userindex).Char.CharIndex & "," & Heading)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "|H" & UserList(userindex).Char.CharIndex & "," & Heading)
    End If
    
End Sub
Sub ChangeUserBody(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal Body As Integer)

    UserList(userindex).Char.Body = Body

    If sndRoute = SendTarget.toMap Then
        Call SendToUserArea(userindex, "|B" & UserList(userindex).Char.CharIndex & "," & Body)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "|B" & UserList(userindex).Char.CharIndex & "," & Body)
    End If
    
End Sub
Sub ChangeUserCasco(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal Casco As Integer)

    UserList(userindex).Char.CascoAnim = Casco

    If sndRoute = SendTarget.toMap Then
        Call SendToUserArea(userindex, "|C" & UserList(userindex).Char.CharIndex & "," & Casco)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "|C" & UserList(userindex).Char.CharIndex & "," & Casco)
    End If
    
End Sub
Sub ChangeUserEscudo(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal Escudo As Integer)

    UserList(userindex).Char.ShieldAnim = Escudo

    If sndRoute = SendTarget.toMap Then
        Call SendToUserArea(userindex, "|E" & UserList(userindex).Char.CharIndex & "," & Escudo)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "|E" & UserList(userindex).Char.CharIndex & "," & Escudo)
    End If
    
End Sub
Sub ChangeUserArma(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal Arma As Integer)

    UserList(userindex).Char.WeaponAnim = Arma

    If sndRoute = SendTarget.toMap Then
        Call SendToUserArea(userindex, "|W" & UserList(userindex).Char.CharIndex & "," & Arma)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "|W" & UserList(userindex).Char.CharIndex & "," & Arma)
    End If
    
End Sub
Sub sendUserRank(ByVal userindex As Integer)
    Call SendData(SendTarget.toMap, 0, UserList(userindex).Pos.Map, "RANK" & UserList(userindex).Char.CharIndex & "," & Mod_Ranking.tieneRanking(userindex))
End Sub

Public Sub CargarExperiencia()
    Dim loopC As Long
    
        For loopC = 1 To STAT_MAXELV
            ArrayExp(loopC) = val(GetVar(App.Path & "\Dat\Experiencia.dat", "EXPERIENCIA", "Nivel" & loopC))
        Next loopC
End Sub


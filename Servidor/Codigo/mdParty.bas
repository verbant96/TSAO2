Attribute VB_Name = "mdParty"
Option Explicit

Private Const MAX_MIEMBROS As Byte = 10

Private Type tParty
    Active As Boolean
    Lider As Integer
    Miembros(1 To MAX_MIEMBROS) As Integer
    cantMiembros As Byte
End Type

Private infoParty(1 To 1000) As tParty
Public Sub CreateParty(ByRef userindex As Integer)

        Dim i As Long, newParty As Integer

        '¿Ya tiene party?
        If UserList(userindex).flags.partyIndex <> 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||372")
          Exit Sub
        End If
        
        'Buscamos un index libre
        For i = 1 To UBound(infoParty)
            If (Not infoParty(i).Active) Then
                newParty = i
                Exit For
            End If
        Next i
        
        'Seteamos las variables correspondientes
        resetParty (newParty)
        infoParty(newParty).Active = True
        infoParty(newParty).Lider = userindex
        infoParty(newParty).Miembros(1) = userindex
        infoParty(newParty).cantMiembros = 1
        
        UserList(userindex).flags.partyIndex = newParty
        Call SendData(SendTarget.toindex, userindex, 0, "||373")
End Sub
Public Function miembrosParty(ByVal index As Integer) As Byte
    miembrosParty = infoParty(index).cantMiembros
End Function
Public Sub party_tepearNobleza(ByVal userindex As Integer)

    Dim i As Long, cuantosMeti As Byte
    cuantosMeti = 0
    
    For i = 1 To miembrosParty(UserList(userindex).flags.partyIndex)
        If (cuantosMeti < 5) And (infoParty(UserList(userindex).flags.partyIndex).Miembros(i) > 0) And (Not MapaEspecial(infoParty(UserList(userindex).flags.partyIndex).Miembros(i))) Then
            Call WarpUserChar(infoParty(UserList(userindex).flags.partyIndex).Miembros(i), 141, RandomNumber(46, 54), RandomNumber(52, 58), True)
            cuantosMeti = cuantosMeti + 1
        End If
    Next i

End Sub
Public Sub party_tepearTanaris(ByVal index As Integer)

    Dim i As Long
    
    For i = 1 To miembrosParty(index)
        If (infoParty(index).Miembros(i) > 0) And (UserList(infoParty(index).Miembros(i)).Pos.Map = 141) Then
            Call WarpUserChar(infoParty(index).Miembros(i), 28, 54, 36, True)
        End If
    Next i

End Sub
Public Sub party_entregarInframundo(ByVal index As Integer)

    Dim i As Long, tmpIndex As Integer
    
    For i = 1 To miembrosParty(index)
        tmpIndex = infoParty(index).Miembros(i)
        If (tmpIndex) And (UserList(tmpIndex).Pos.Map = 141) Then
            
            If TieneObjetos(1073, 1, tmpIndex) And TieneObjetos(1074, 1, tmpIndex) And TieneObjetos(1075, 1, tmpIndex) And TieneObjetos(1076, 1, tmpIndex) Then
                Dim Fer As obj
                Fer.ObjIndex = 1077
                Fer.Amount = 1
                
                Call MeterItemEnInventario(tmpIndex, Fer)
                Call SendData(SendTarget.toindex, tmpIndex, 0, "||49")
            End If
                
            Call SendData(SendTarget.toindex, tmpIndex, 0, "||983")
            Call WarpUserChar(tmpIndex, 28, 54, 36, True)
        End If
    Next i

End Sub
Public Sub SoliciteParty(ByRef Leader As Integer, ByRef NewMember As Integer)

                Dim partyIndex As Integer
                partyIndex = UserList(Leader).flags.partyIndex
                
                '¿El que invita tiene party?
                If partyIndex = 0 Then
                    Call SendData(SendTarget.toindex, Leader, 0, "||376")
                    Exit Sub
                End If
                
                '¿Es el lider de la party?
                If infoParty(partyIndex).Lider <> Leader Then
                    Call SendData(SendTarget.toindex, Leader, 0, "||377")
                    Exit Sub
                End If
                
                '¿El miembro a invitar tiene party?
                If UserList(NewMember).flags.partyIndex <> 0 Then
                    Call SendData(SendTarget.toindex, Leader, 0, "||380")
                  Exit Sub
                End If
                
                '¿El miembro a invitar está muerto?
                If UserList(NewMember).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, Leader, 0, "||374")
                    Exit Sub
                End If
        
                '¿Se está autoinvitando?
                If Leader = NewMember Then
                    Call SendData(SendTarget.toindex, Leader, 0, "||375")
                    Exit Sub
                End If
                
                '¿Está muy lejos?
                If Distancia(UserList(NewMember).Pos, UserList(Leader).Pos) > 3 Then
                    Call SendData(SendTarget.toindex, Leader, 0, "||10")
                    Exit Sub
                End If
                
                '¿Ya tiene la solicitud de esta party?
                If UserList(NewMember).flags.PartySolicitud = partyIndex Then
                    Call SendData(SendTarget.toindex, Leader, 0, "||378")
                  Exit Sub
                End If
                
                'Buscamos slot libre
                Dim posIn As Byte
                posIn = slotLibre(partyIndex)
                
                If (posIn <> 0) Then
                    UserList(NewMember).flags.PartySolicitud = partyIndex
                    Call SendData(SendTarget.toindex, NewMember, 0, "||381@" & UserList(Leader).Name)
                    Call SendData(SendTarget.toindex, Leader, 0, "||382@" & UserList(NewMember).Name)
                Else
                    Call SendData(SendTarget.toindex, Leader, 0, "||379")
                End If
End Sub
Public Sub acceptParty(ByRef userindex As Integer)

        Dim solicitudIndex As Integer
        solicitudIndex = UserList(userindex).flags.PartySolicitud
        
        '¿Tiene alguna solicitud de party? o ¿La party ofertada sigue activa?
        If (solicitudIndex = 0) Or (Not infoParty(solicitudIndex).Active) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||383")
            Exit Sub
         End If
         
         '¿Está en alguna party ya?
         If UserList(userindex).flags.partyIndex <> 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||384")
            Exit Sub
         End If
         
         Dim posIn As Byte
         posIn = slotLibre(solicitudIndex)
         
         If (posIn > 0) Then
            infoParty(solicitudIndex).Miembros(posIn) = userindex
            infoParty(solicitudIndex).cantMiembros = infoParty(solicitudIndex).cantMiembros + 1
            UserList(userindex).flags.partyIndex = solicitudIndex
            UserList(userindex).flags.PartySolicitud = 0
            Call SendData(SendTarget.ToPartyArea, userindex, 0, "||386@" & UserList(userindex).Name)
         Else
            Call SendData(SendTarget.toindex, userindex, 0, "||379")
         End If

End Sub
Public Sub cancelParty(ByRef userindex As Integer)

            '¿Ya está una party?, la abandonamos.
            If UserList(userindex).flags.partyIndex <> 0 Then
               Call SendData(SendTarget.ToPartyArea, userindex, 0, "||388@" & UserList(userindex).Name)
                
               infoParty(UserList(userindex).flags.partyIndex).cantMiembros = infoParty(UserList(userindex).flags.partyIndex).cantMiembros - 1
               UserList(userindex).flags.partyIndex = 0
               UserList(userindex).flags.PartySolicitud = 0
               Exit Sub
            End If

            '¿Tiene una solicitud de party?, la cancelamos.
            If UserList(userindex).flags.PartySolicitud <> 0 Then
               UserList(userindex).flags.PartySolicitud = 0
               Call SendData(SendTarget.toindex, userindex, 0, "||387")
               Exit Sub
            Else
                   Call SendData(SendTarget.toindex, userindex, 0, "||383")
               Exit Sub
            End If
         
End Sub
Public Sub closeParty(ByRef userindex As Integer)

        Dim indexParty As Integer
        indexParty = UserList(userindex).flags.partyIndex

        '¿Está en una party?
        If (indexParty = 0) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||389")
            Exit Sub
        End If
        
        '¿Es lider?
        If (infoParty(indexParty).Lider <> userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||377")
            Exit Sub
        End If
         
         
         Call SendData(SendTarget.ToPartyArea, userindex, 0, "||390")
         resetParty (indexParty)
         UserList(userindex).flags.partyIndex = 0
         UserList(userindex).flags.PartySolicitud = 0

End Sub
Public Sub informationParty(ByRef userindex As Integer)

            Dim indexParty As Integer, i As Long
            indexParty = UserList(userindex).flags.partyIndex

            '¿Tiene una party?
            If (indexParty = 0) Then
               Call SendData(SendTarget.toindex, userindex, 0, "||377")
             Exit Sub
            End If
            
            'Informamos
            Call SendData(SendTarget.toindex, userindex, 0, "||391")
            For i = 1 To MAX_MIEMBROS
               If (infoParty(indexParty).Miembros(i) <> 0) Then
                       Call SendData(SendTarget.toindex, userindex, 0, "||392@" & UserList(infoParty(indexParty).Miembros(i)).Name & "@" & UserList(infoParty(indexParty).Miembros(i)).Stats.ELV & "@" & UserList(infoParty(indexParty).Miembros(i)).clase)
                End If
            Next i
End Sub
Public Sub disconnectParty(ByRef userindex As Integer)

    'Si es lider, cerramos la party, de lo contrario solo abandonamos.
    If (infoParty(UserList(userindex).flags.partyIndex).Lider = userindex) Then
        mdParty.closeParty (userindex)
    Else
        mdParty.cancelParty (userindex)
    End If

End Sub
Public Sub doExperience(ByVal userindex As Integer, ByVal Experiencia As Long)
      
      Dim i As Long, indexParty, partyMember As Integer
      indexParty = UserList(userindex).flags.partyIndex
      
      Experiencia = Experiencia / infoParty(indexParty).cantMiembros
      
      For i = 1 To infoParty(indexParty).cantMiembros
            partyMember = infoParty(indexParty).Miembros(i)
        
            If (partyMember <> 0) Then
                If (Distancia(UserList(partyMember).Pos, UserList(userindex).Pos) < 15) Then
                    UserList(partyMember).Stats.Exp = UserList(partyMember).Stats.Exp + Experiencia
                    
                    If UserList(partyMember).Stats.Exp > MAXEXP Then UserList(partyMember).Stats.Exp = MAXEXP
                    Call SendData(SendTarget.toindex, partyMember, 0, "||170@" & PonerPuntos(Experiencia))
                    SendUserEXP (partyMember)
                    Call CheckUserLevel(partyMember)
                End If
            End If
      Next i
        
End Sub
Private Sub resetParty(index As Integer)
    infoParty(index).Active = False
    infoParty(index).Lider = 0
    infoParty(index).cantMiembros = 0
    
    Dim i As Long, partyMember As Integer
    For i = 1 To MAX_MIEMBROS
        partyMember = infoParty(index).Miembros(i)
        
        If (partyMember <> 0) Then
            With UserList(partyMember)
                .flags.partyIndex = 0
                .flags.PartySolicitud = 0
            End With
            
            infoParty(index).Miembros(i) = 0
        End If
    Next i
End Sub
Private Function slotLibre(index As Integer) As Byte

            Dim i As Long, sLib As Byte
            sLib = 0
            
            
            For i = 1 To MAX_MIEMBROS
                If (infoParty(index).Miembros(i) = 0) Then
                    sLib = i
                    Exit For
                End If
            Next i
            
            slotLibre = sLib
End Function


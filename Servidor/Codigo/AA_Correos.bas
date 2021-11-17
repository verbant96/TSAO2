Attribute VB_Name = "AA_Correos"

Public Sub correoIniciarForm(userindex As Integer)

Dim cRemitente As String


Dim comIte      As String
Dim comIl       As Long
Dim comTemp     As String
Dim Temporal As String

        With UserList(userindex)
            For comIl = 1 To 30
                Temporal = ReadField(1, .flags.Correo(comIl), Asc("$"))
                
                If Temporal = "0" Then
                    comTemp = "Nada"
                Else
                    If .flags.NueCorreos(comIl) = 1 Then
                        comTemp = "" & Temporal & " (NUEVO)"
                    Else
                        comTemp = Temporal
                    End If
                End If
                
                comIte = comIte & comTemp & ","
            Next comIl
                
            SendData SendTarget.toindex, userindex, 0, "IFO" & comIte
            comIte = ""
        
            For comIl = 1 To 20
                comTemp = "(Nada)"
                    If .Invent.Object(comIl).ObjIndex > 0 Then comTemp = ObjData(.Invent.Object(comIl).ObjIndex).Name
                comIte = comIte & .Invent.Object(comIl).ObjIndex & "-" & .Invent.Object(comIl).Amount & "-" & comTemp & ","
            Next comIl
        End With
    SendData SendTarget.toindex, userindex, 0, "IDO" & UserList(userindex).Name & "$" & comIte
    comIte = ""
    
        For comIl = 1 To 20
            comIte = comIte & UCase$(UserList(userindex).flags.NombreAmigo(comIl)) & ","
        Next comIl
    SendData SendTarget.toindex, userindex, 0, "IAO" & comIte
    comIte = ""
    
    correoReset userindex
    
End Sub
Public Sub correoEnviarMensaje(userindex As Integer, rData As String)
On Error GoTo Errhandler

    correoReset (userindex)

    With UserList(userindex)
        Dim iMoC As Long, cDatPalOtro As String, itData As String, cNamePutTemp As String, cTempGrh As Integer, Destinatario As String, Mensaje As String, Asunto As String
        
            Destinatario = ReadField(1, rData, Asc("$"))
            Asunto = ReadField(2, rData, Asc("$"))
            Mensaje = ReadField(3, rData, Asc("$"))
            
            If GetVar(CharPath & Destinatario & ".chr", "CORREO", "NUMCORREOS") = 30 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||629")
                Exit Sub
            End If
            
            If UserList(userindex).cComercio.cComercia = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||153")
                Exit Sub
            End If
            
            If Len(Destinatario) <= 0 Then
                    SendData SendTarget.toindex, userindex, 0, "ERRUsuario inexistente"
                Exit Sub
            End If
            
            If FileExist(CharPath & UCase$(Destinatario) & ".chr", vbNormal) = False Then
                SendData SendTarget.toindex, userindex, 0, "ERRUsuario inexistente"
                Exit Sub
            End If
            
            cDatPalOtro = cDatPalOtro & UserList(userindex).Name & "$" & Asunto & "$" & Mensaje & "$" & Date & "$"
        
            For iMoC = 1 To 20
                cNamePutTemp = "(Nada)"
                
                cTempItMo = ReadField(iMoC, rData, Asc(","))
                cTempGrh = "0"
                
                    If ReadField(2, cTempItMo, Asc("-")) > 0 Then
                        .cCorreo.cObj(iMoC).Amount = ReadField(2, cTempItMo, Asc("-"))
                        .cCorreo.cObj(iMoC).ObjIndex = .Invent.Object(iMoC).ObjIndex
                    End If
                    
                    If .cCorreo.cObj(iMoC).ObjIndex > 0 Then
                        cNamePutTemp = ObjData(.cCorreo.cObj(iMoC).ObjIndex).Name
                        cTempGrh = .cCorreo.cObj(iMoC).ObjIndex
                        
                        If ReadField(2, cTempItMo, Asc("-")) > 0 Then
                            If ObjData(.Invent.Object(iMoC).ObjIndex).Intransferible = 1 Or ObjData(.Invent.Object(iMoC).ObjIndex).ItemDios = 1 Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||630")
                                Exit Sub
                                Exit For
                            End If
                        End If
                        
                        If TieneObjetos(.cCorreo.cObj(iMoC).ObjIndex, .cCorreo.cObj(iMoC).Amount, userindex) = False Then
                                correoReset userindex
                                Call SendData(SendTarget.toindex, userindex, 0, "||630")
                          Exit Sub
                          Exit For
                        End If
                        
                    End If
            Next iMoC
            
            'Con esto dejamos de volar el correo de mierda ese.
            For iMoC = 1 To 20
                Dim NameTemporal As String
                Dim AmountTemporal As Integer
                Dim ObjTemporal As Integer
                NameTemporal = "(Nada)"
                AmountTemporal = 0
                ObjTemporal = 0
            
                If .cCorreo.cObj(iMoC).ObjIndex > 0 Then
                    NameTemporal = ObjData(.cCorreo.cObj(iMoC).ObjIndex).Name
                    AmountTemporal = .cCorreo.cObj(iMoC).Amount
                    ObjTemporal = .cCorreo.cObj(iMoC).ObjIndex
                    
                    
                    Call QuitarObjetos(.cCorreo.cObj(iMoC).ObjIndex, .cCorreo.cObj(iMoC).Amount, userindex)
                    Call LogCorreos("" & UserList(userindex).Name & " envio: " & .cCorreo.cObj(iMoC).Amount & " - " & ObjData(.cCorreo.cObj(iMoC).ObjIndex).Name & " (OBJ: " & .cCorreo.cObj(iMoC).ObjIndex & ") a " & Destinatario & "")
                End If
                
                    itData = itData & ObjTemporal & "-" & AmountTemporal & "-" & NameTemporal & ","
            Next iMoC
            
    End With
    
        Dim NumCorreos As Byte
        Dim NueCorreos As String
        Dim NTCR As String
        Dim CorreoTemporal As String
        

        If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Administrador Then
            Call LogGM(UserList(userindex).Name, "Correo: " & UserList(userindex).Name & " quiso enviar por correo " & cDatPalOtro, False)
            Exit Sub
        End If
        
    Dim inDest As Integer
    inDest = NameIndex(Destinatario)
    
    If inDest <= 0 Then
        NumCorreos = GetVar(CharPath & Destinatario & ".chr", "CORREO", "NUMCORREOS")
        NueCorreos = GetVar(CharPath & Destinatario & ".chr", "CORREO", "NUECORREOS")
        Call WriteVar(CharPath & Destinatario & ".chr", "CORREO", "CORREONUM" & NumCorreos + 1, cDatPalOtro)
        Call WriteVar(CharPath & Destinatario & ".chr", "CORREO", "CORREOITEMS" & NumCorreos + 1, itData)
        Call WriteVar(CharPath & Destinatario & ".chr", "CORREO", "NUMCORREOS", NumCorreos + 1)
        
        'Escribimos que tiene un correo nuevo de una forma muy villera
        NTCR = ""
        For iMoC = 1 To 30
            CorreoTemporal = ReadField(iMoC, NueCorreos, Asc(","))
            If iMoC = NumCorreos + 1 Then
                NTCR = NTCR & iMoC & "-1"
            Else
                NTCR = NTCR & iMoC & "-" & ReadField(2, CorreoTemporal, Asc("-")) & ","
            End If
        Next iMoC
        
        Call WriteVar(CharPath & Destinatario & ".chr", "CORREO", "NUECORREOS", NTCR)
        correoReset userindex
    Else
        UserList(inDest).flags.NumCorreos = UserList(inDest).flags.NumCorreos + 1
        UserList(inDest).flags.Correo(UserList(inDest).flags.NumCorreos) = cDatPalOtro
        UserList(inDest).flags.itemsCorreo(UserList(inDest).flags.NumCorreos) = itData
        UserList(inDest).flags.NueCorreos(UserList(inDest).flags.NumCorreos) = 1
        Call SendData(SendTarget.toindex, inDest, 0, "||631")
        correoReset userindex
    End If
    
Errhandler:
    Call LogError("Error al enviar correos.")
End Sub
Public Sub correoLeerMensaje(userindex As Integer, rData As String)
    On Error Resume Next

If rData = 0 Then Exit Sub

correoReset userindex

    With UserList(userindex)

        Dim iMoC As Long, cDatPalOtro As String, cNamePutTemp As String, cTempGrh As Integer, cData As String
        
        cData = UserList(userindex).flags.Correo(rData)
        iData = UserList(userindex).flags.itemsCorreo(rData)
        cDatPalOtro = ""
        
            For iMoC = 1 To 20
                cNamePutTemp = "(Nada)"
                Dim cTempItMo As String
                cTempItMo = ReadField(iMoC, iData, Asc(","))
                
                    If ReadField(2, cTempItMo, Asc("-")) > 0 Then
                        .cCorreo.cObj(iMoC).Amount = ReadField(2, cTempItMo, Asc("-"))
                        .cCorreo.cObj(iMoC).ObjIndex = ReadField(1, cTempItMo, Asc("-"))
                    End If
                    
                    If .cCorreo.cObj(iMoC).ObjIndex > 0 Then
                        cNamePutTemp = ObjData(.cCorreo.cObj(iMoC).ObjIndex).Name
                        cTempGrh = .cCorreo.cObj(iMoC).ObjIndex
                    End If
                    
                cDatPalOtro = cDatPalOtro & cTempGrh & "-" & .cCorreo.cObj(iMoC).Amount & "-" & cNamePutTemp & ","
            Next iMoC
            
        SendData SendTarget.toindex, userindex, 0, "ILO" & cData
        SendData SendTarget.toindex, userindex, 0, "ITO" & cDatPalOtro
        UserList(userindex).flags.NueCorreos(rData) = 0
        
        Call correoRecargarLista(userindex)
        
    End With
    
    
End Sub
Public Sub correoBorrarMensaje(userindex As Integer, rData As String)

If rData = 0 Then Exit Sub

    With UserList(userindex)
    
        If .flags.Correo(rData) = "0" Then Exit Sub

        Dim iMoC As Long, cTemp As Integer, cCorreo As String, nCorreo As String, itCorreo
            cTemp = 0
            

                .flags.Correo(rData) = "0"
                .flags.itemsCorreo(rData) = "0"
                
                For iMoC = 1 To 30
                    cCorreo = ""
                    nCorreo = ""
                    'Cambiamos el numero de correos
                    If .flags.Correo(iMoC) <> "0" Then
                        'Damos temporal
                        cTemp = cTemp + 1
                        cCorreo = .flags.Correo(iMoC)
                        nCorreo = .flags.NueCorreos(iMoC)
                        itCorreo = .flags.itemsCorreo(iMoC)
                        
                        'Ponemos en 0
                        .flags.Correo(iMoC) = 0
                        .flags.NueCorreos(iMoC) = 0
                        .flags.itemsCorreo(iMoC) = 0
                        
                        'Reescribimos
                        .flags.Correo(cTemp) = cCorreo
                        .flags.NueCorreos(cTemp) = nCorreo
                        .flags.itemsCorreo(cTemp) = itCorreo
                    Else
                        'Sino le damos que ya lo leyo
                        .flags.NueCorreos(iMoC) = 0
                    End If
                Next iMoC
                
                    'Seteamos numeros
                    .flags.NumCorreos = cTemp
            
            Call correoRecargarLista(userindex)
            
    End With
    
    
End Sub
Public Sub correoRetirarItems(userindex As Integer, rData As String)
Dim CorreoObj As Obj
Dim lopC As Long

If rData = 0 Or rData > 30 Then Exit Sub

For lopC = 1 To 20
    With UserList(userindex)
        CorreoObj.ObjIndex = .cCorreo.cObj(lopC).ObjIndex
        CorreoObj.Amount = .cCorreo.cObj(lopC).Amount

        If CorreoObj.ObjIndex <> 0 Then
            If Not MeterItemEnInventario(userindex, CorreoObj) Then
                Call TirarItemAlPiso(UserList(userindex).Pos, CorreoObj)
            End If
            
            Call LogRCorreos("" & UserList(userindex).Name & " retiró: " & CorreoObj.Amount & " - " & ObjData(CorreoObj.ObjIndex).Name & "")
        End If
    End With
Next lopC

UserList(userindex).flags.itemsCorreo(rData) = "0"
Call correoLeerMensaje(userindex, rData)
    
End Sub
Private Sub correoRecargarLista(userindex As Integer)
             'CARGAMOS NUEVAMENTE LA LISTA DE MENSAJES
            Dim comIte As String, Temporal As String, comTemp As String
                
            With UserList(userindex)
                    For iMoC = 1 To 30
                            Temporal = ReadField(1, .flags.Correo(iMoC), Asc("$"))
                            If Temporal = "0" Then
                                comTemp = "Nada"
                            Else
                                If .flags.NueCorreos(iMoC) = 1 Then
                                    comTemp = "" & Temporal & " (NUEVO)"
                                Else
                                    comTemp = Temporal
                                End If
                            End If
                            
                            comIte = comIte & comTemp & ","
                        Next iMoC
                SendData SendTarget.toindex, userindex, 0, "IFO" & comIte
                comIte = ""
            End With
End Sub
Public Sub correoReset(userindex As Integer)
Dim comI As Long
    With UserList(userindex)
            For comI = 1 To 20
                .cCorreo.cObj(comI).Amount = 0
                .cCorreo.cObj(comI).ObjIndex = 0
            Next comI
    End With
Exit Sub
End Sub

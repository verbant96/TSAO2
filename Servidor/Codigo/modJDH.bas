Attribute VB_Name = "modJDH"
Option Explicit
 
'***************
'AUTOR: Toyz - Luciano
'FECHA: 14/12/16 - 07:30
'***************
Private Const Tiempo_Cancelamiento As Integer = 180
Private Const Cofre_Abierto As Integer = 10 'Número de cofre abierto.
Private Const Cofre_Cerrado As Integer = 11 'Número de cofre cerrado.

Public HayJDH As Boolean

Private Type tUsuario
    X As Byte
    Y As Byte
End Type

Private Type tJDH
    Activo As Boolean
    Usuarios(1 To 10) As tUsuario
    Conteo As Integer
    Cupos As Byte
    mapa As Integer
    Inscripcion As Long
    Total As Byte
    Restantes As Byte
End Type
 
Private JDH As tJDH
Public Sub Carga_JDH()
    Dim loopC As Long
    Dim loopX As Long
    Dim LoopZ As Long
    Dim DataCofre As obj
 
    DataCofre.Amount = 1
    DataCofre.ObjIndex = Cofre_Cerrado
 
    With JDH
        .Cupos = UBound(.Usuarios())
        .mapa = GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "EVENTO", "Mapa")
        
        For loopC = 1 To .Cupos
            .Usuarios(loopC).X = GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "USUARIO#" & loopC, "X")
            .Usuarios(loopC).Y = GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "USUARIO#" & loopC, "Y")
        Next loopC
        
    End With
End Sub
 
Public Sub Armar_JDH(ByVal Cupos As Byte, ByVal Inscripcion As Long)
    With JDH
        If .Activo = True Then Exit Sub
        
        Dim i As Long
        For i = 1 To LastUser
            UserList(i).flags.EnJDH = False
        Next i
        
        .Cupos = Cupos
        .Inscripcion = Inscripcion
        .Total = .Cupos
        .Restantes = .Total
        .Activo = True
        HayJDH = True
        .Conteo = Tiempo_Cancelamiento
        
        Call SendData(SendTarget.ToAll, 0, 0, "||885@" & .Cupos & "@" & PonerPuntos(.Inscripcion))
    End With
End Sub
Public Function JDH_Activo() As Boolean

    If JDH.Cupos > 0 Then
        JDH_Activo = False
    Else
        JDH_Activo = True
    End If

End Function
Public Sub Entrar_JDH(ByVal ID As Integer)

    With JDH
        If Puede_Entrar(ID) = False Then Exit Sub
        
        Call SendData(SendTarget.toindex, ID, 0, "||886")
        UserList(ID).Stats.GLD = UserList(ID).Stats.GLD - .Inscripcion
        SendUserGLD ID
        
        UserList(ID).flags.EnJDH = True
        UserList(ID).flags.tmpPos = UserList(ID).Pos
        
        Save_Inventory (ID) '//Salvamos el inventario
        WarpUserChar ID, .mapa, .Usuarios(.Cupos).X, .Usuarios(.Cupos).Y, False
        .Cupos = .Cupos - 1
        
        UserList(ID).flags.NotMove = 1
        Call SendData(SendTarget.toindex, ID, 0, "STOPD" & UserList(ID).flags.NotMove)
        
        Dim MiObj As obj
        MiObj.Amount = 5500
        MiObj.ObjIndex = 36
        Call MeterItemEnInventario(ID, MiObj)
        MiObj.ObjIndex = 37
        Call MeterItemEnInventario(ID, MiObj)
        MiObj.ObjIndex = 38
        Call MeterItemEnInventario(ID, MiObj)
        MiObj.ObjIndex = 39
        Call MeterItemEnInventario(ID, MiObj)
        MiObj.ObjIndex = 15
        MiObj.Amount = 1
        Call MeterItemEnInventario(ID, MiObj)
        
        Call SendData(SendTarget.ToAll, 0, 0, "||915@" & UserList(ID).Name)
        
        If .Cupos = 0 Then
            Call SendData(SendTarget.toMap, 0, .mapa, "||887")
            .Conteo = 10
            Call JDH_CofresAzar
            frmMain.JDH.Enabled = True
        End If
        
    End With
End Sub
Private Function Puede_Entrar(ByVal ID As Integer) As Boolean
    Puede_Entrar = False
    If UserList(ID).flags.Muerto > 0 Then
        Call SendData(SendTarget.toindex, ID, 0, "||3")
        Exit Function
    End If
    If UserList(ID).flags.EnJDH Then
        Call SendData(SendTarget.toindex, ID, 0, "||97")
        Exit Function
    End If
    If MapaEspecial(ID) Then
        Call SendData(SendTarget.toindex, ID, 0, "||291")
        Exit Function
    End If
    If JDH.Activo = False Then
        Call SendData(SendTarget.toindex, ID, 0, "||882")
        Exit Function
    End If
    If JDH.Cupos = 0 Then
        Call SendData(SendTarget.toindex, ID, 0, "||883")
        Exit Function
    End If
    If UserList(ID).Stats.GLD < JDH.Inscripcion Then
        Call SendData(SendTarget.toindex, ID, 0, "||663")
        Exit Function
    End If
    If MapInfo(UserList(ID).Pos.Map).Pk Then
        Call SendData(SendTarget.toindex, ID, 0, "||323")
        Exit Function
    End If
    Puede_Entrar = True
End Function
 
Public Sub Contar_JDH()
    Dim loopC As Long
    Dim loopX As Long
    With JDH
        If .Conteo = 0 Then
            .Conteo = -1
            If .Activo = True Then
            
                For loopC = 1 To LastUser
                    If UserList(loopC).flags.NotMove = 1 Then
                        UserList(loopC).flags.NotMove = 0
                        Call SendData(SendTarget.toindex, loopC, 0, "STOPD" & UserList(loopC).flags.NotMove)
                    End If
                Next loopC
                
                If .Cupos = 0 Then
                    Call SendData(SendTarget.toMap, 0, JDH.mapa, "N|Juegos del Hambre> ¡YA!" & FONTTYPE_ORO)
                    SendData SendTarget.toMap, 0, .mapa, "CU0"
                    frmMain.JDH.Enabled = False
                Else
                    Cancelar_JDH
                End If
            End If
        End If
     
        If .Conteo > 0 Then
            If .Cupos = 0 Then _
                Call SendData(SendTarget.toMap, 0, JDH.mapa, "N|Juegos del Hambre> " & .Conteo & FONTTYPE_INFO)
                SendData SendTarget.toMap, 0, .mapa, "CU" & .Conteo
                .Conteo = .Conteo - 1
        End If
    End With
End Sub
Public Sub Muere_JDH(ByVal ID As Integer)

    If (Not UserList(ID).flags.EnJDH) Then Exit Sub
    
    With JDH
        .Restantes = .Restantes - 1
        
        If .Restantes > 1 Then Call SendData(SendTarget.toMap, 0, .mapa, "||889@" & .Restantes)
        Call SendData(SendTarget.toindex, ID, 0, "||888")
        UserList(ID).flags.EnJDH = False
        TirarTodosLosItems ID
        ReLoad_Inventory (ID)
        
        WarpUserChar ID, UserList(ID).flags.tmpPos.Map, UserList(ID).flags.tmpPos.X, UserList(ID).flags.tmpPos.Y, False
        If .Restantes <= 1 Then Finalizar
    End With
End Sub
Private Sub Finalizar()
    Dim loopC As Long
    Dim ID As Integer
    
    With JDH
        For loopC = 1 To LastUser
            If UserList(loopC).flags.EnJDH And UserList(loopC).Pos.Map = .mapa Then
                ID = loopC
            End If
        Next loopC
        
        Call SendData(SendTarget.ToAll, 0, 0, "||891@" & UserList(ID).Name)
        ReLoad_Inventory (ID)
        UserList(ID).flags.EnJDH = False
        
        'Premio
        UserList(ID).Stats.TSPoints = UserList(ID).Stats.TSPoints + 1
        Call SendData(SendTarget.toindex, ID, 0, "||900@1")
        
        WarpUserChar ID, UserList(ID).flags.tmpPos.Map, UserList(ID).flags.tmpPos.X, UserList(ID).flags.tmpPos.Y, False
        Limpiar
    End With
End Sub
Public Sub Cancelar_JDH()
    Dim loopC As Long
    With JDH
        If .Activo = False Then Exit Sub
        For loopC = 1 To LastUser
            If UserList(loopC).flags.EnJDH And UserList(loopC).Pos.Map = .mapa Then
                UserList(loopC).flags.NotMove = 0
                Call SendData(SendTarget.toindex, loopC, 0, "STOPD" & UserList(loopC).flags.NotMove)
                ReLoad_Inventory (loopC)
                UserList(loopC).flags.EnJDH = False
                UserList(loopC).Stats.GLD = UserList(loopC).Stats.GLD + .Inscripcion
                SendUserGLD loopC
                WarpUserChar loopC, UserList(loopC).flags.tmpPos.Map, UserList(loopC).flags.tmpPos.X, UserList(loopC).flags.tmpPos.Y, False
            End If
        Next loopC

        Call SendData(SendTarget.ToAll, 0, 0, "||890")
        Limpiar
    End With
End Sub
Public Sub Desconexion_JDH(ByVal ID As Integer)

    If (Not UserList(ID).flags.EnJDH) Then Exit Sub
    
    With JDH
        TirarTodosLosItems ID
        ReLoad_Inventory (ID)
        UserList(ID).flags.EnJDH = False
        WarpUserChar ID, UserList(ID).flags.tmpPos.Map, UserList(ID).flags.tmpPos.X, UserList(ID).flags.tmpPos.Y, False
        
        If .Cupos > 0 Then
            .Cupos = .Cupos + 1
        Else
            .Restantes = .Restantes - 1
            If .Restantes > 1 Then Call SendData(SendTarget.toMap, 0, .mapa, "||889@" & .Restantes)
            If .Restantes <= 1 Then Finalizar
        End If
    End With
End Sub
Public Sub CambiaMapa_JDH(ByVal ID As Integer)

    If Not UserList(ID).flags.EnJDH Then Exit Sub
    
    With JDH
        TirarTodosLosItems ID
        ReLoad_Inventory (ID)
        UserList(ID).flags.EnJDH = False
        
        If .Cupos > 0 Then
            .Cupos = .Cupos + 1
        Else
            .Restantes = .Restantes - 1
            If .Restantes > 1 Then Call SendData(SendTarget.toMap, 0, .mapa, "||889@" & .Restantes)
            If .Restantes <= 1 Then Finalizar
        End If
    End With
End Sub
Private Sub Limpiar()
    Dim loopC As Long
    With JDH
        .Activo = False
        .Conteo = -1
        .Inscripcion = 0
        .Restantes = 0
        .Total = 0
        Call LimpiarMapa(.mapa)
    End With
    
    HayJDH = False
    frmMain.JDH.Enabled = False
End Sub
Private Sub JDH_CofresAzar()

    Dim CofreCerradito As obj, i As Long, CRandomX As Byte, CRandomY As Byte, BaseUsuarios As Boolean
    CofreCerradito.ObjIndex = 1548
    CofreCerradito.Amount = 1
        
    'Los cofres fijos en el medio:
    Call MakeObj(SendTarget.toMap, 0, JDH.mapa, CofreCerradito, JDH.mapa, 48, 49)
    Call MakeObj(SendTarget.toMap, 0, JDH.mapa, CofreCerradito, JDH.mapa, 50, 49)
    Call MakeObj(SendTarget.toMap, 0, JDH.mapa, CofreCerradito, JDH.mapa, 52, 49)
    
    Call MakeObj(SendTarget.toMap, 0, JDH.mapa, CofreCerradito, JDH.mapa, 49, 50)
    Call MakeObj(SendTarget.toMap, 0, JDH.mapa, CofreCerradito, JDH.mapa, 51, 50)
    
    Call MakeObj(SendTarget.toMap, 0, JDH.mapa, CofreCerradito, JDH.mapa, 48, 51)
    Call MakeObj(SendTarget.toMap, 0, JDH.mapa, CofreCerradito, JDH.mapa, 50, 51)
    Call MakeObj(SendTarget.toMap, 0, JDH.mapa, CofreCerradito, JDH.mapa, 52, 51)
    
    'Ahora tiramos 100 cofres al azar.
    For i = 1 To 100
        CRandomX = RandomNumber(26, 74)
        CRandomY = RandomNumber(25, 73)
        
        BaseUsuarios = (CRandomX >= 42 And CRandomX <= 58) And (CRandomY >= 43 And CRandomY <= 57)
        
                
        'Si esa pos no sirve, buscamos otra.
        Do While (MapData(JDH.mapa, CRandomX, CRandomY).Blocked = 1 Or MapData(JDH.mapa, CRandomX, CRandomY).OBJInfo.ObjIndex > 0 Or MapData(JDH.mapa, CRandomX, CRandomY).userindex > 0 Or BaseUsuarios)
            CRandomX = RandomNumber(26, 74)
            CRandomY = RandomNumber(25, 73)
            
            BaseUsuarios = (CRandomX >= 42 And CRandomX <= 58) And (CRandomY >= 43 And CRandomY <= 57)
        Loop
        
        'Si llegamos acá es porque encontramos donde clavar el cofre.
        Call MakeObj(SendTarget.toMap, 0, JDH.mapa, CofreCerradito, JDH.mapa, CRandomX, CRandomY)
    Next i

End Sub
Public Sub Clickea_Cofre(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userindex As Integer)

    Dim NumCofre As Integer
    Dim NumeritoX As Long, MiObj As obj, PosCofre As WorldPos
    Dim loopC As Long
    
    PosCofre.Map = Map
    PosCofre.X = X
    PosCofre.Y = Y
    
    If Not HayJDH Then Call SendData(SendTarget.toindex, userindex, 0, "||911"): Exit Sub
    If JDH.Conteo <> -1 Then Call SendData(SendTarget.toindex, userindex, 0, "||912"): Exit Sub
    
    If Distancia(PosCofre, UserList(userindex).Pos) > 2 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||910")
        Exit Sub
    End If
    
    For loopC = 1 To 5
        'Recorre los objetos y va descartando
        NumCofre = RandomNumber(1, CofresAzar(2).CantObjs)
        NumeritoX = RandomNumber(1, 100)
    
        Do While (NumeritoX > CofresAzar(2).ObjProbability(NumCofre))
            NumCofre = RandomNumber(1, CofresAzar(2).CantObjs)
            NumeritoX = RandomNumber(1, 100)
        Loop
            'Llegó acá, osea que encontró un obj para tirar.
            MiObj.ObjIndex = CofresAzar(2).ObjIndex(NumCofre)
            MiObj.Amount = CofresAzar(2).ObjAmount(NumCofre)
            
            Call TirarItemAlPiso(PosCofre, MiObj)
    Next loopC
    
    MiObj.ObjIndex = 10
    Call EraseObj(SendTarget.toMap, 0, Map, 10000, Map, X, Y)
    Call MakeObj(SendTarget.toMap, 0, Map, MiObj, Map, X, Y)

End Sub
Public Function JDH_PuedeAtacar() As Boolean

    If JDH.Conteo <> 0 And JDH.Conteo <> -1 Then
        JDH_PuedeAtacar = False
    Else
        JDH_PuedeAtacar = True
    End If

End Function
Private Sub Save_Inventory(ByVal ID As Integer)
    
    '//Guardamos todo el inventario actual del usuario
    Dim loopC As Long
    For loopC = 1 To MAX_INVENTORY_SLOTS
        If UserList(ID).Invent.Object(loopC).ObjIndex > 0 Then Call LogJDH("" & UserList(ID).Name & " ingresó con: " & UserList(ID).Invent.Object(loopC).Amount & " - " & ObjData(UserList(ID).Invent.Object(loopC).ObjIndex).Name)
        UserList(ID).Invent.ExObject(loopC).ObjIndex = UserList(ID).Invent.Object(loopC).ObjIndex
        UserList(ID).Invent.ExObject(loopC).Amount = UserList(ID).Invent.Object(loopC).Amount
    Next loopC
    
    '//Lo desnudamos
    Call LimpiarInventario(ID)
    Call DarCuerpoDesnudo(ID)
        
    Call ChangeUserChar(SendTarget.toMap, 0, UserList(ID).Pos.Map, val(ID), UserList(ID).Char.Body, UserList(ID).Char.Head, UserList(ID).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    Call UpdateUserInv(True, ID, 0)
        
End Sub
Private Sub ReLoad_Inventory(ByVal ID As Integer)

        '//Lo desnudamos
        If UserList(ID).flags.Muerto = 0 Then Call DarCuerpoDesnudo(ID)
        Call LimpiarInventario(ID)
        Call ChangeUserChar(SendTarget.toMap, 0, UserList(ID).Pos.Map, val(ID), UserList(ID).Char.Body, UserList(ID).Char.Head, UserList(ID).Char.Heading, NingunArma, NingunEscudo, NingunCasco)

    '//Devolvemos el inventario que tenía antes de ingresar
    Dim loopC As Long
    For loopC = 1 To MAX_INVENTORY_SLOTS
        If UserList(ID).Invent.ExObject(loopC).ObjIndex > 0 Then Call LogJDH("" & UserList(ID).Name & " salió con: " & UserList(ID).Invent.ExObject(loopC).Amount & " - " & ObjData(UserList(ID).Invent.ExObject(loopC).ObjIndex).Name)
        UserList(ID).Invent.Object(loopC).ObjIndex = UserList(ID).Invent.ExObject(loopC).ObjIndex
        UserList(ID).Invent.Object(loopC).Amount = UserList(ID).Invent.ExObject(loopC).Amount
    Next loopC
    
    '//Ahora le reiniciamos el salvado.
    For loopC = 1 To MAX_INVENTORY_SLOTS
        UserList(ID).Invent.ExObject(loopC).ObjIndex = 0
        UserList(ID).Invent.ExObject(loopC).Amount = 0
        UserList(ID).Invent.ExObject(loopC).Equipped = 0
    Next loopC

    Call UpdateUserInv(True, ID, 0)

End Sub

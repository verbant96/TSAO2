Attribute VB_Name = "Mod_General"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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


Public bK As Long
Public bRK As Long

Public iplst As String
Public banners As String

Public bFogata As Boolean
Public sHKeys() As String
Public Function DirPath(ByVal Path As String) As String
'•Parra: Nuevo Engine v2.0
    Select Case Path
        Case "Graficos"
            DirPath = App.Path & "\Data\GRAFICOS\"
            Exit Function
        
        Case "Sound"
            DirPath = App.Path & "\Data\SOUNDS\WAV\"
            Exit Function
        
        Case "Midi"
            DirPath = App.Path & "\Data\SOUNDS\MIDI\"
            Exit Function
        
        Case "Maps"
            DirPath = App.Path & "\Data\MAPAS\"
            Exit Function
    End Select
End Function

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\Data\" & "GRAFICOS" & "\"
End Function
Public Function DirSound() As String
    DirSound = App.Path & "\Data\SOUNDS\" & "WAV" & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.Path & "\Data\SOUNDS\" & "MIDI" & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.Path & "\Data\" & "MAPAS" & "\"
End Function

Public Function SumaDigitos(ByVal Numero As Integer) As Integer
    'Suma digitos
    Do
        SumaDigitos = SumaDigitos + (Numero Mod 10)
        Numero = Numero \ 10
    Loop While (Numero > 0)
End Function

Public Function SumaDigitosMenos(ByVal Numero As Integer) As Integer
    'Suma digitos, y resta el total de dígitos
    Do
        SumaDigitosMenos = SumaDigitosMenos + (Numero Mod 10) - 1
        Numero = Numero \ 10
    Loop While (Numero > 0)
End Function

Public Function Complex(ByVal Numero As Integer) As Integer
    If Numero Mod 2 <> 0 Then
        Complex = Numero * SumaDigitos(Numero)
    Else
        Complex = Numero * SumaDigitosMenos(Numero)
    End If
End Function

Public Function ValidarLoginMSG(ByVal Numero As Integer) As Integer
    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    
    AuxInteger = SumaDigitos(Numero)
    AuxInteger2 = SumaDigitosMenos(Numero)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim Arch As String
    
    Arch = App.Path & "\Data\INIT\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(Arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(Arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(Arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(Arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(Arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub CargarVersiones()
On Error GoTo errorh:

    Versiones(1) = Val(GetVar(App.Path & "\Data\INIT\" & "versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(GetVar(App.Path & "\Data\INIT\" & "versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(GetVar(App.Path & "\Data\INIT\" & "versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(GetVar(App.Path & "\Data\INIT\" & "versiones.ini", "Init", "Val"))
    Versiones(5) = Val(GetVar(App.Path & "\Data\INIT\" & "versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(GetVar(App.Path & "\Data\INIT\" & "versiones.ini", "E", "Val"))
    Versiones(7) = Val(GetVar(App.Path & "\Data\INIT\" & "versiones.ini", "O", "Val"))
    VersionC = GetVar(App.Path & "\Data\INIT\" & "versiones.ini", "VERSION", "V")
Exit Sub

errorh:
    Call MsgBox("Error cargando versiones")
End Sub
Sub CargarColores()
    Dim archivoC As String
    
    archivoC = App.Path & "\Data\INIT\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 47 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    

'NW
ColoresPJ(22).r = 255
ColoresPJ(22).g = 255
ColoresPJ(22).b = 202
'NW
'Poder
ColoresPJ(20).r = 225
ColoresPJ(20).g = 225
ColoresPJ(20).b = 225
'Poder
'Horda sin enlistar
ColoresPJ(47).r = 227
ColoresPJ(47).g = 141
ColoresPJ(47).b = 150
'Horda sin enlistar
'Alianza sin enlistar
ColoresPJ(46).r = 132
ColoresPJ(46).g = 193
ColoresPJ(46).b = 225
'Alianza sin enlistar
'Horda Enlistado
ColoresPJ(50).r = 255
ColoresPJ(50).g = 0
ColoresPJ(50).b = 0
'Horda Enlistado
'Alianza Enlistado
ColoresPJ(49).r = 0
ColoresPJ(49).g = 128
ColoresPJ(49).b = 255
'Alianza Enlistado
'Neutral
ColoresPJ(48).r = 125
ColoresPJ(48).g = 125
ColoresPJ(48).b = 125
'Neutral

'CONCILIO ALIANZA
ColoresPJ(51).r = 16
ColoresPJ(51).g = 38
ColoresPJ(51).b = 96

'CONCILIO HORDA
ColoresPJ(52).r = 69
ColoresPJ(52).g = 13
ColoresPJ(52).b = 14
End Sub
Sub CargarAuras()
    Dim archivoC As String
    
    archivoC = App.Path & "\Data\INIT\Auras.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
        Call MsgBox("ERROR: no se ha podido cargar las auras. Falta el archivo auras.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    

    Dim XX As Long
    For XX = 1 To GetVar(App.Path & "\Data\INIT\Auras.dat", "INIT", "NumAuras")
        AurasPJ(XX).GrhIndex = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "GrhIndex")
        AurasPJ(XX).r = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Rojo")
        AurasPJ(XX).g = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Verde")
        AurasPJ(XX).b = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Azul")
        AurasPJ(XX).offset = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Offset")
        AurasPJ(XX).Giratoria = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Giratoria")
        AurasPJ(XX).RojoF = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "RojoF")
        AurasPJ(XX).AzulF = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "AzulF")
        AurasPJ(XX).VerdeF = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "VerdeF")
    Next XX
    
End Sub
Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim Arch As String
    
    Arch = App.Path & "\Data\INIT\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(Arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(Arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(Arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(Arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(Arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal text As String, Optional ByVal red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************

If UserConsola = 0 Then
    With RichTextBox
        If (Len(.text)) > 10000 Then .text = ""
        
        .SelStart = Len(RichTextBox.text)
        .SelLength = 0
        
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, Green, Blue)
        
        .SelText = IIf(bCrLf, text, text & vbCrLf)
    
        'RichTextBox.Refresh
    End With
    End If
End Sub
'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).charindex = loopc
        End If
    Next loopc
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
   
    If checkemail And UserEmail = "" Then
        Mensaje.Escribir "Dirección de email invalida"
        Exit Function
    End If
   
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
           Mensaje.Escribir "Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido."
            Exit Function
        End If
    Next loopc
   
    If nombrecuent = "" Then
        Mensaje.Escribir "Ingrese un nombre de cuenta."
        Exit Function
    End If
   
    If UserPassword = "" Then
        Mensaje.Escribir "Ingrese un password."
        Exit Function
    End If
    If Len(nombrecuent) > 30 Then
        Mensaje.Escribir "La cuenta debe tener menos de 30 letras."
        Exit Function
    End If
   
    For loopc = 1 To Len(nombrecuent)
        CharAscii = Asc(mid$(nombrecuent, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            Mensaje.Escribir "Cuenta inválida. El caractér " & Chr$(CharAscii) & " no está permitido."
            Exit Function
        End If
    Next loopc
   
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form

    For Each mifrm In Forms
        Unload mifrm

    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function
Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
 'Set Connected
    Connected = True
    Unload frmAccount
     
    If UserLvl > 50 Then
        frmMain.LvlLbl.ForeColor = vbYellow
    Else
        frmMain.LvlLbl.ForeColor = vbRed
    End If
    
    modoHabla = ";"
    dKeys(1) = 0
    dKeys(2) = 0
    dKeys(3) = 0
    dKeys(4) = 0

    frmMain.ExpBar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Principal_expBar.jpg")
    
    Call SetWindowLong(frmMain.RecTxt.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    Call SetWindowLong(frmMain.PrivatesConsole.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    Call SetWindowLong(frmMain.GlobalConsole.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    Call SetWindowLong(frmMain.ClanConsole.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    
    frmMain.hlst.Clear
    Dim i As Long
    For i = 1 To MAXUSERHECHIZOS
        frmMain.hlst.AddItem "(Nada)"
    Next i
    
    frmMain.hlst.BackColor = RGB(22, 23, 25)
    frmMain.SendTxt.BackColor = RGB(22, 23, 25)
    frmMain.SendCMSTXT.BackColor = RGB(22, 23, 25)
    
    frmMain.ItemName.ForeColor = RGB(186, 179, 169)
    
    frmMain.rep.ForeColor = RGB(186, 179, 169)
    frmMain.Exp.ForeColor = RGB(186, 179, 169)
    
    frmMain.Visible = True
    Call DibujarPuntoMinimap
    Call DibujarMinimap
    Call SaveGameini
    
    frmMain.DyD.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\D&D_N.jpg")
    
    For i = 1 To MAX_INVENTORY_SLOTS
        Call Inventario.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "(Nada)")
    Next i
    
    Call AgregarParticulasyLuces(UserMap)

    If TieneColorMapa = False Then
        day_r_old = 215
        day_g_old = 215
        day_b_old = 215
        base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
    End If
    
End Sub
Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
'***************************************************
    Dim LegalOk As Boolean
    
    If Stopped = 1 Then Exit Sub
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
   If LegalOk Then
        Call SendData("M" & Direccion)
        engine.Char_Move_by_Head UserCharIndex, Direccion
        engine.Engine_MoveScreen Direccion
        UserMeditar = False
        RefreshAllChars
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            charlist(UserCharIndex).Heading = Direccion
            Call SendData("CHEA" & Direccion)
        End If
    End If
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************

    MoveTo RandomNumber(1, 4)
    
End Sub
Sub CheckKeys()
Static LastMovement As Long
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
        
    If UserParalizado Then
        If GetTickCount() - LastMovement > 96 Then
                LastMovement = GetTickCount()
        Else
                Exit Sub
        End If
    
            If GetKeyState(BindKeys(14).KeyCode) < 0 And (dKeys(1) = BindKeys(14).KeyCode) Then
                If BindKeys(14).KeyCode <> 38 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                    If charlist(UserCharIndex).Heading <> 1 Then
                        Call SendData("CHEA" & 1)
                        charlist(UserCharIndex).Heading = 1
                        Exit Sub
                    End If
            End If
       
            'Move Right
            If GetKeyState(BindKeys(17).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 And (dKeys(1) = BindKeys(17).KeyCode) Then
                If BindKeys(17).KeyCode <> 39 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                    If charlist(UserCharIndex).Heading <> 2 Then
                        Call SendData("CHEA" & 2)
                        charlist(UserCharIndex).Heading = 2
                        Exit Sub
                    End If
            End If
       
            'Move down
            If GetKeyState(BindKeys(15).KeyCode) < 0 And (dKeys(1) = BindKeys(15).KeyCode) Then
                If BindKeys(15).KeyCode <> 40 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                    If charlist(UserCharIndex).Heading <> 3 Then
                        Call SendData("CHEA" & 3)
                        charlist(UserCharIndex).Heading = 3
                        Exit Sub
                    End If
            End If
       
            'Move left
            If GetKeyState(BindKeys(16).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 And (dKeys(1) = BindKeys(16).KeyCode) Then
                If BindKeys(16).KeyCode <> 37 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                    If charlist(UserCharIndex).Heading <> 4 Then
                        Call SendData("CHEA" & 4)
                        charlist(UserCharIndex).Heading = 4
                        Exit Sub
                    End If
            End If
            
        Exit Sub
    End If

   
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
            'Move U
            If GetKeyState(BindKeys(14).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 And (dKeys(1) = BindKeys(14).KeyCode) Then
                If BindKeys(14).KeyCode <> 38 And frmMain.SendTxt.Visible = True Then Exit Sub
            
                Call MoveTo(NORTH)
                Exit Sub
            End If
       
            'Move Right
            If GetKeyState(BindKeys(17).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 And (dKeys(1) = BindKeys(17).KeyCode) Then
                If BindKeys(17).KeyCode <> 39 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                Call MoveTo(EAST)
                Exit Sub
            End If
       
            'Move down
            If GetKeyState(BindKeys(15).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 And (dKeys(1) = BindKeys(15).KeyCode) Then
                If BindKeys(15).KeyCode <> 40 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                Call MoveTo(SOUTH)
                Exit Sub
            End If
       
            'Move left
            If GetKeyState(BindKeys(16).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 And (dKeys(1) = BindKeys(16).KeyCode) Then
                If BindKeys(16).KeyCode <> 37 And frmMain.SendTxt.Visible = True Then Exit Sub

                Call MoveTo(WEST)
                Exit Sub
            End If
    End If
End Sub


'TODO : esto no es del tileengine??
Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
    
        Case E_Heading.EAST
            X = 1
    
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
            
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y

    If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
        Exit Sub
    End If
End Sub

'TODO : esto no es del tileengine??
Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************
    Dim loopc As Long
    
    loopc = 1
    Do While charlist(loopc).active And loopc < UBound(charlist)
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function
Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    Dim TempLng As Byte
    Dim TempByte1 As Byte
    Dim TempByte2 As Byte
    Dim TempByte3 As Byte
    Dim i As Byte

    'By Lorwik - www.rincondelao.com.ar
    engine.Particle_Group_Remove_All
    Light.Light_Remove_All
    handle = FreeFile()
    
    Open DirPath("Maps") & "Mapa" & Map & ".map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            For i = 0 To 3
                MapData(X, Y).light_value(i) = False
            Next i
            Get handle, , ByFlags
            MapData(X, Y).luz = 0
            MapData(X, Y).particle_group = 0
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
            
            If ByFlags And 32 Then
               Get handle, , tempint
                MapData(X, Y).particle_group_index = General_Particle_Create(tempint, X, Y, -1)
            End If
            
            'Erase NPCs
            If MapData(X, Y).charindex > 0 Then
                Call EraseChar(MapData(X, Y).charindex)
            End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
            
        Next X
    Next Y
    
    Close handle
    
    Dim pX As Byte
    Dim pY As Byte
    
    For pX = 1 To 100
        For pY = 1 To 100
            If (MapData(pX, pY).charindex > 0) Then ResetCharInfo (MapData(pX, pY).charindex)
            MapData(pX, pY).charindex = 0
            MapData(pX, pY).ObjGrh.GrhIndex = 0
        Next pY
    Next pX
    
    RefreshAllChars
    
    If Map = 999 And frmAccount.Visible = True Then
        Call General_Particle_Create(6, 50, 49, -1)
        Call General_Particle_Create(6, 41, 42, -1)
        Call General_Particle_Create(6, 59, 42, -1)
        Call General_Particle_Create(6, 41, 56, -1)
        Call General_Particle_Create(6, 59, 56, -1)
    End If

    If frmAccount.Visible = True Then Exit Sub
    
    Call AgregarParticulasyLuces(Map)
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map
    
    Call DibujarPuntoMinimap
    Call DibujarMinimap
                
End Sub
'TODO : Reemplazar por la nueva versión, esta apesta!!!
Public Function ReadField(ByVal Pos As Integer, ByVal text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim i As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(text)
        CurChar = mid$(text, i, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = mid$(text, LastPos + 1, (InStr(LastPos + 1, text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = mid$(text, LastPos + 1)
    End If
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.Path & "\Data\INIT\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function

Public Sub CargarServidores()
On Error GoTo errorh
    Dim f As String
    Dim c As Integer
    Dim i As Long
    
    f = App.Path & "\Data\INIT\sinfo.dat"
    c = Val(GetVar(f, "INIT", "Cant"))
    
    ReDim ServersLst(1 To c) As tServerInfo
    For i = 1 To c
        ServersLst(i).Desc = GetVar(f, "S" & i, "Desc")
        ServersLst(i).Ip = Trim$(GetVar(f, "S" & i, "Ip"))
        ServersLst(i).PassRecPort = CInt(GetVar(f, "S" & i, "P2"))
        ServersLst(i).Puerto = CInt(GetVar(f, "S" & i, "PJ"))
    Next i
    CurServer = 1
Exit Sub

errorh:
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Argentum Online")
    End
End Sub
Public Function CurServerIp() As String

  CurServerIp = "177.54.153.240"
  'CurServerIp = "127.0.0.1"

End Function
Public Function CurServerPort() As Integer

    CurServerPort = "5028"

End Function
Sub Main()

On Error Resume Next

Dim strIconPath As String

LastTime = GetTickCount()

HDSerial = GetDriveSerialNumber
AoDefAntiShInitialize
AoDefOriginalClientName = "Tierras Sagradas"
AoDefClientName = App.EXEName
AoDefDetectName = App.EXEName
Set AodefConv = New AoDefenderConverter

Call GenCM("quierovalecuatro") 'clave grh.


    Dim i As Integer, iX As Integer, tX As Integer, DifX As Integer, dNum As String
    
    
    'iX = frmCargando.Inet1.OpenURL("http://www.tierras-sagradas.net/AU/Actualizaciones.txt")
    'tX = LeerInt(App.Path & "\Data\INIT\Update.tsao")
    
    'DifX = iX - tX
    
    'If Not (DifX = 0) Then
    '    If MsgBox("Tu versión no es la actual. ¿Desea ejecutar el actualizador automatico?", vbYesNo, "Tierras Sagradas AO") = vbYes Then
    '        ShellExecute frmCargando.hWnd, "runas", App.Path & "\Launcher TSAO.exe", "", App.Path, vbNormalFocus
    '        End
    '    End If
    'End If
        

'If AoDefChangeName Then
'  Call AoDefClientOn
' End
'End If

If AoDefDebugger Then
    Call AoDefAntiDebugger
   End
End If

    Call WriteClientVer

    If App.PrevInstance Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If

    
    ChDrive App.Path
    ChDir App.Path
    
    Set Light = New clsLight
    
    'Por default usamos el dinámico
    Set SurfaceDB = New clsSurfaceManDynDX8
    
    'cargamos la config del user
    Call frmMenuGral.LoadOptions
        
    frmCargando.Show
    Call frmCargando.establecerProgreso(0)
     
    frmResolucion.Show vbModal, frmCargando
    
    Dim lc As Integer
    Dim LACONCHA As String
    lc = 0
    For lc = 1 To NUMBINDS
        LACONCHA = GetVar(App.Path & "\Data\INIT\" & "Teclas.tsao", "TECLAS", Str(lc))
        BindKeys(lc).KeyCode = Val(ReadField(1, LACONCHA, 44))
        BindKeys(lc).Name = ReadField(2, LACONCHA, 44)
    Next lc
    frmCargando.establecerProgreso (10)
    
    
    Call CargarColores
    Call CargarAuras
    Call General_Load_Interfaces
    frmConnect.limpiarConectar
    frmCargando.Refresh
    frmMain.Socket1.Startup
    frmCargando.establecerProgreso (20)
    Call InicializarNombres
    UserMap = 1
    LoadGrhData
    CargarParticulas
    frmCargando.establecerProgreso (50)
    
    Call CargarParticulas
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarVersiones
    frmCargando.establecerProgreso (70)
    
    CargarCabezas
    CargarCascos
    CargarCuerpos
    CargarFxs
    
    frmCargando.establecerProgreso (80)
    
    modTextos.InitFonts
    modTextos.LoadText
    
    Call engine.Engine_Init
    frmCargando.establecerProgreso (90)
    
    'Inicializamos el sonido
    Call Audio.Initialize(frmMain.hWnd, App.Path & "\Data\SOUNDS\" & "WAV" & "\", App.Path & "\Data\SOUNDS\" & "MIDI" & "\")
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv)
    
    frmCargando.establecerProgreso (100)
    
    For i = 1 To 8
        RangoPRIV(i) = "<Game Master>"
    Next i
        
    RangoPRIV(9) = "<Director de GMs>"
    RangoPRIV(10) = "<Developer>"
    RangoPRIV(11) = "<Sub Administrador>"
    RangoPRIV(12) = "<Administrador>"
    
    EsStatusCOLOR(0) = D3DColorXRGB(ColoresPJ(48).r, ColoresPJ(48).g, ColoresPJ(48).b)
    EsStatusCOLOR(1) = D3DColorXRGB(ColoresPJ(46).r, ColoresPJ(46).g, ColoresPJ(46).b)
    EsStatusCOLOR(2) = D3DColorXRGB(ColoresPJ(47).r, ColoresPJ(47).g, ColoresPJ(47).b)
    EsStatusCOLOR(3) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
    EsStatusCOLOR(4) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
    EsStatusCOLOR(5) = D3DColorXRGB(ColoresPJ(51).r, ColoresPJ(51).g, ColoresPJ(51).b)
    EsStatusCOLOR(6) = D3DColorXRGB(ColoresPJ(52).r, ColoresPJ(52).g, ColoresPJ(52).b)
    EsStatusCOLOR(8) = D3DColorXRGB(ColoresPJ(22).r, ColoresPJ(22).g, ColoresPJ(22).b)
    
    Unload frmCargando
    
    AoDefResult = 0
    frmPres.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Pres1.jpg")
    frmPres.Timer1.Enabled = True
    frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
    Call frmConnect.mostrarConectar(True, True)

    frmMain.InvEqu.Picture = General_Load_Interface_Picture("Centronuevoinventario.jpg")
    frmMain.Picture = General_Load_Interface_Picture("Principal.jpg")

    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
    Sound = Configuracion.Sound
    Musica = Configuracion.Music
    
    Dim IntroMusic As Byte
    IntroMusic = RandomNumber(1, 3)
    If IntroMusic = 1 Then
        Audio.MP3_Play "70"
    ElseIf IntroMusic = 2 Then
        Audio.MP3_Play "73"
    Else
        Audio.MP3_Play "140"
    End If
    
    Dialogos.font = frmMain.font
    
engine.Start


    
Exit Sub
ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    Debug.Print "Contexto:" & err.HelpContext & " Desc:" & err.Description & " Fuente:" & err.Source
    End
End Sub
Function FieldCount(ByRef text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal var As String, ByVal Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, Value, File
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean

    HayAgua = MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
                MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open App.Path & "\Data\INIT\ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(1) = "Ullathorpe"
    Ciudades(2) = "Nix"
    Ciudades(3) = "Banderbill"

    CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
    CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
    CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Bardo"
    ListaClases(6) = "Druida"
    ListaClases(7) = "Paladin"
    ListaClases(8) = "Cazador"

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.meditar) = "Meditar"
    SkillsNames(Skills.Apuñalar) = "Apuñalar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar árboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"
    SkillsNames(Skills.DefensaMagica) = "Defensa Magica"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub

Public Sub LogError(Desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\Data\errores.log" For Append As #nfile
Print #nfile, Desc
Close #nfile
End Sub

Public Sub LogCustom(Desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\Data\custom.log" For Append As #nfile
Print #nfile, Now & " " & Desc
Close #nfile
End Sub
Public Sub DrawGrhtoHdc(desthDC As Long, ByVal grh_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
 
On Error Resume Next
   
Dim file_path As String, src_x As Integer, src_y As Integer, src_width As Integer, src_height As Integer, hdcsrc As Long, PrevObj As Long
 
If grh_index <= 0 Then Exit Sub
 
If GrhData(grh_index).NumFrames <> 1 Then
    
grh_index = GrhData(grh_index).Frames(1)
    
End If
 
        If Extract_File(Graphics, App.Path & "\Data\GRAFICOS\", GrhData(grh_index).FileNum & ".bmp", App.Path & "\Data\GRAFICOS\") Then
            file_path = App.Path & "\Data\GRAFICOS\" & GrhData(grh_index).FileNum & ".bmp"
        End If
       
src_x = GrhData(grh_index).sX
src_y = GrhData(grh_index).sY
src_width = GrhData(grh_index).pixelWidth
src_height = GrhData(grh_index).pixelHeight
                    
hdcsrc = CreateCompatibleDC(desthDC)
         
PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
 
BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
 
DeleteDC hdcsrc
 
End Sub
Public Function SuficientePC() As Boolean
Dim MS As MEMORYSTATUS

MS.dwLength = Len(MS)
GlobalMemoryStatus MS

'si es menor a 0.3 el NW tiene una makina de MIERDA
If ((MS.dwTotalPhys / (1000 ^ 3)) > 0.3) Then
    SuficientePC = True
Else
    SuficientePC = False
End If

End Function
Public Sub DibujarPuntoMinimap()
    
With frmMain
.Puntito.left = UserPos.X - 2
.Puntito.top = UserPos.Y - 3
End With
    
End Sub
Public Sub DibujarMinimap()

    If FileExist(App.Path & "\Data\GRAFICOS\MiniMap\Mapa" & UserMap & ".bmp", vbNormal) Then
        frmMain.Minimap.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\MiniMap\Mapa" & UserMap & ".bmp")
    Else
        frmMain.Minimap.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\MiniMap\Nada.bmp")
    End If

End Sub

Public Function PonerPuntos(Numero As Long) As String
Dim i As Integer
Dim Cifra As String
 
Cifra = Str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next
 
PonerPuntos = left$(PonerPuntos, Len(PonerPuntos) - 1)
 
End Function
Public Function OpenBrowser(strURL As String, lngHwnd As Long)
    OpenBrowser = ShellExecute(lngHwnd, vbNullString, strURL, vbNullString, _
    "c:\", 0)
End Function
Public Function General_Load_Interface_Picture(ByVal PicName As String) As IPicture

On Error GoTo err
'vars
Dim GUIFolder As String
GUIFolder = App.Path & "\Data\GRAFICOS\Principal\"

'vemos si existe la interfas sino cargamos la default
If FileExist(GUIFolder & PicName, vbNormal) Then 'existe la cargamos
    Set General_Load_Interface_Picture = LoadPicture(GUIFolder & PicName)
    'Dest.Picture = LoadPicture(GUIFolder & PicName)
Else 'usamos la default
    Set General_Load_Interface_Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\" & PicName)
    'Dest.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\" & PicName)
End If

Exit Function

'error
err:
LogError "Error al cargar la imagen " & PicName & ", la imagen no se encontro."

End Function
Public Sub General_Load_Interfaces()

Dim N As Integer
Dim i As Integer

N = Val(GetVar(App.Path & "\Data\INIT\Interfaz.dat", "MAIN", "Interfaces"))

ReDim Interfaces(1 To N) As String

For i = 1 To N
    Interfaces(i) = GetVar(App.Path & "\Data\INIT\Interfaz.dat", "INTERFACES", "N" & i)
Next i

End Sub
Public Sub TirarItemMouse()
    Dim tX As Byte
    Dim tY As Byte
    Dim CantidadGG As String
    OfMouse = True
    Call ConvertCPtoTP(frmMain.MouseX, frmMain.MouseY, tX, tY)
    
    Dim Namepos As String, NameReal As String
    
If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
    
    If MapData(tX, tY).charindex > 0 Then
    
        'If MapData(tX, tY).charindex = charindex Then Exit Sub
        
            If charlist(MapData(tX, tY).charindex).NPCNumber = 36 Then
                'depositar
                If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                    SendData ("DEPO" & "," & Inventario.SelectedItem & "," & 1)
                ElseIf Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    CantidadGG = InputBox("Ingresa la cantidad de " & Inventario.ItemName(Inventario.SelectedItem) & " que quieras DEPOSITAR (0 para cancelar):", "¿Cantidad?", "0")
                    If Not IsNumeric(CantidadGG) Then Exit Sub
                    If CantidadGG = 0 Or CantidadGG > 10000 Then Exit Sub
                    SendData ("DEPO" & "," & Inventario.SelectedItem & "," & CantidadGG)
                End If
              Exit Sub
            End If
        
        
            Namepos = InStr(charlist(MapData(tX, tY).charindex).Nombre, "<")
            If Namepos = 0 Then Namepos = Len(charlist(MapData(tX, tY).charindex).Nombre) + 2
            NameReal = left$(charlist(MapData(tX, tY).charindex).Nombre, Namepos - 2)
    
        'transferir
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            If MsgBox("¿Transferirle 1 " & Inventario.ItemName(Inventario.SelectedItem) & " al usuario " & charlist(MapData(tX, tY).charindex).Nombre & "?", vbYesNo, "Confirmacion") = vbYes Then
                Call SendData("DYDTRA" & tX & "," & tY & "," & NameReal & "," & Inventario.SelectedItem & "," & 1)
            End If
        ElseIf Inventario.Amount(Inventario.SelectedItem) > 1 Then
            CantidadGG = InputBox("Ingresa la cantidad de " & Inventario.ItemName(Inventario.SelectedItem) & " que quieras TRANSFERIR a " & charlist(MapData(tX, tY).charindex).Nombre & " (0 para cancelar):", "¿Cantidad?", "0")
            If Not IsNumeric(CantidadGG) Then Exit Sub
            If CantidadGG = 0 Or CantidadGG > 10000 Then Exit Sub
            Call SendData("DYDTRA" & tX & "," & tY & "," & NameReal & "," & Inventario.SelectedItem & "," & CantidadGG)
        End If
        
           MouseRendOK = False
      Exit Sub
        
    Else
        'tirar
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            Call SendData("TR" & Inventario.SelectedItem & "," & 1 & "," & tX & "," & tY)
        ElseIf Inventario.Amount(Inventario.SelectedItem) > 1 Then
                CantidadGG = InputBox("Ingresa la cantidad de " & Inventario.ItemName(Inventario.SelectedItem) & " que quieras TIRAR (0 para cancelar):", "¿Cantidad?", "0")
                If Not IsNumeric(CantidadGG) Then Exit Sub
                If CantidadGG = 0 Or CantidadGG > 10000 Then Exit Sub
                Call SendData("TR" & Inventario.SelectedItem & "," & CantidadGG & "," & tX & "," & tY)
        End If
    End If
End If

End Sub
Public Sub mostrarCuenta()

        frmAccount.Show

          With frmAccount
                .imgCambiarPass.Visible = True
                .imgCrearPersonaje.Visible = True
                .imgSalir4.Visible = True
                
                Dim i As Long
                For i = 0 To CantidadDePersonajes - 1
                    .PJ(i).Visible = True
                Next i
          End With
            
        Audio.StopWave
End Sub
Public Function GetDriveSerialNumber(Optional ByVal DriveLetter As String) As Long
'***************************************************
'Author: Nahuel Casas (Zagen)
'Last Modify Date: 07/12/2009
' 07/12/2009: Zagen - Convertì las funciones, en formulas mas fàciles de modificar.
'***************************************************
    On Error Resume Next
          Dim fso As Object, Drv As Object, DriveSerial As Long
         
          'Creamos el objeto FileSystemObject.
          Set fso = CreateObject("Scripting.FileSystemObject")
         
          'Asignamos el driver principal.
          If DriveLetter <> "" Then
              Set Drv = fso.GetDrive(DriveLetter)
          Else
              Set Drv = fso.GetDrive(fso.GetDriveName(App.Path))
          End If
     
          With Drv
              If .IsReady Then
                  DriveSerial = Abs(.SerialNumber)
              Else    '"Si el driver no està como para empezar ..."
                  DriveSerial = -1
              End If
          End With
         
          'Borramos y limpiamos.
          Set Drv = Nothing
          Set fso = Nothing
    'Seteamos :)
    GetDriveSerialNumber = DriveSerial
         
End Function
Public Sub DarColorCambiante(ByVal charindex As Long)

With charlist(charindex)
    Select Case .EsStatus
        Case 0 '125 125 125
            .AntiguoR = 125
            .AntiguoG = 125
            .AntiguoB = 125
            
        Case 1
            .AntiguoR = 132
            .AntiguoG = 193
            .AntiguoB = 225
            
        Case 2 '227 141 150
            .AntiguoR = 227
            .AntiguoG = 141
            .AntiguoB = 150
            
        Case 3 '0 128 255
            .AntiguoR = 0
            .AntiguoG = 128
            .AntiguoB = 255
        Case 4 '255 0 0
            .AntiguoR = 255
            .AntiguoG = 0
            .AntiguoB = 0
        
        Case 5
            .AntiguoR = 16
            .AntiguoG = 38
            .AntiguoB = 96
            
        Case 6 '69 13 14
            .AntiguoR = 69
            .AntiguoG = 13
            .AntiguoB = 14
            
    End Select
    
            If .color = 40 Then
                'Le damos el color original directamente
                If .ProximoR = 0 And .ProximoG = 0 And .ProximoB = 0 Then
                    .ProximoR = 255
                    .ProximoG = 255
                    .ProximoB = 0
                End If
            
                    'Si ya supero el máximo le damos directamente l color y empezamos a darle paso al azul.
                    If .ProximoR >= 255 And .ProximoG >= 255 And .LlegoAlColor = False Then
                        .ProximoR = 255
                        .ProximoG = 255
                        .ProximoB = 0
                        .LlegoAlColor = True
                    End If
                
                'Empezamos a darle color amarillo
                If ((.ProximoR < 255) Or (.ProximoG < 255)) And .LlegoAlColor = False Then
                    .ProximoR = .ProximoR + 1
                    .ProximoG = .ProximoG + 1
                    .ProximoB = .ProximoB - 1
                                       
                    If .ProximoR >= 255 Then .ProximoR = 255
                    If .ProximoG >= 255 Then .ProximoG = 255
                    If .ProximoB < 0 Then .ProximoB = 0
                    
                'Si ya llego al amarillo, empezamos a darle el color gris.
                ElseIf .LlegoAlColor = True Then
                
                    .ProximoR = .ProximoR - 1
                    .ProximoG = .ProximoG - 1
                    .ProximoB = .ProximoB + 1
                    
                    If .ProximoR <= .AntiguoR Then .ProximoR = .AntiguoR
                    If .ProximoG <= .AntiguoG Then .ProximoG = .AntiguoG
                    If .ProximoB >= .AntiguoB Then .ProximoB = .AntiguoB
                    
                    'Ya llegamos al gris, vamos a darle paso al amarillo
                    If .ProximoR = .AntiguoR And .ProximoG = .AntiguoG And .ProximoB = .AntiguoB Then
                        .LlegoAlColor = False
                    End If
                
                End If
            ElseIf .color = 42 Then
                'Le damos el color original directamente
                If .ProximoR = 0 And .ProximoG = 0 And .ProximoB = 0 Then
                    .ProximoR = 255
                    .ProximoG = 255
                    .ProximoB = 255
                End If
            
                    'Si ya supero el máximo le damos directamente l color y empezamos a darle paso al azul.
                    If .ProximoR >= 255 And .ProximoG >= 255 And .ProximoB >= 255 And .LlegoAlColor = False Then
                        .ProximoR = 255
                        .ProximoG = 255
                        .ProximoB = 255
                        .LlegoAlColor = True
                    End If
                
                'Empezamos a darle color amarillo
                If ((.ProximoR < 255) Or (.ProximoG < 255) Or (.ProximoB < 255)) And .LlegoAlColor = False Then
                    .ProximoR = .ProximoR + 1
                    .ProximoG = .ProximoG + 1
                    .ProximoB = .ProximoB + 1
                                       
                    If .ProximoR >= 255 Then .ProximoR = 255
                    If .ProximoG >= 255 Then .ProximoG = 255
                    If .ProximoB >= 255 Then .ProximoB = 255
                    
                'Si ya llego al amarillo, empezamos a darle el color gris.
                ElseIf .LlegoAlColor = True Then
                
                    .ProximoR = .ProximoR - 1
                    .ProximoG = .ProximoG - 1
                    .ProximoB = .ProximoB - 1
                    
                    If .ProximoR <= .AntiguoR Then .ProximoR = .AntiguoR
                    If .ProximoG <= .AntiguoG Then .ProximoG = .AntiguoG
                    If .ProximoB <= .AntiguoB Then .ProximoB = .AntiguoB
                    
                    'Ya llegamos al gris, vamos a darle paso al amarillo
                    If .ProximoR = .AntiguoR And .ProximoG = .AntiguoG And .ProximoB = .AntiguoB Then
                        .LlegoAlColor = False
                    End If
                
                End If
                
            ElseIf .color = 41 Then
                'Le damos el color original directamente
                If .ProximoR = 0 And .ProximoG = 0 And .ProximoB = 0 Then
                    .ProximoR = 95
                    .ProximoG = 45
                    .ProximoB = 95
                End If
                
                'Empezamos a darle color amarillo
                If .LlegoAlColor = False Then
                    'ROJO
                    If (.ProximoR < 95) Then
                        .ProximoR = .ProximoR + 1
                        If .ProximoR >= 95 Then .ProximoR = 95
                    End If
                    If (.ProximoR > 95) Then
                        .ProximoR = .ProximoR - 1
                        If .ProximoR <= 95 Then .ProximoR = 95
                    End If
                    'ROJO
                    
                    'VERDE
                    If (.ProximoG < 45) Then
                        .ProximoG = .ProximoG + 1
                        If .ProximoG >= 45 Then .ProximoG = 45
                    End If
                    If (.ProximoG > 45) Then
                        .ProximoG = .ProximoG - 1
                        If .ProximoG <= 45 Then .ProximoG = 45
                    End If
                    'VERDE
                    
                    'AZUL
                    If (.ProximoB < 95) Then
                        .ProximoB = .ProximoB + 1
                        If .ProximoB >= 95 Then .ProximoB = 95
                    End If
                    If (.ProximoB > 95) Then
                        .ProximoB = .ProximoB - 1
                        If .ProximoB >= 95 Then .ProximoB = 95
                    End If
                    'AZUL
                    
                    'Si ya supero el máximo le damos directamente l color y empezamos a darle paso al azul.
                    If .ProximoR = 95 And .ProximoG = 45 And .ProximoB = 95 And .LlegoAlColor = False Then
                        .ProximoR = 95
                        .ProximoG = 45
                        .ProximoB = 95
                        .LlegoAlColor = True
                    End If
                    
                'Si ya llego al amarillo, empezamos a darle el color gris.
                ElseIf .LlegoAlColor = True Then
                
                    'ROJO
                    If (.ProximoR < .AntiguoR) Then
                        .ProximoR = .ProximoR + 1
                        If .ProximoR >= .AntiguoR Then .ProximoR = .AntiguoR
                    End If
                    If (.ProximoR > .AntiguoR) Then
                        .ProximoR = .ProximoR - 1
                        If .ProximoR <= .AntiguoR Then .ProximoR = .AntiguoR
                    End If
                    'ROJO
                    
                    'VERDE
                    If (.ProximoG < .AntiguoG) Then
                        .ProximoG = .ProximoG + 1
                        If .ProximoG >= .AntiguoG Then .ProximoG = .AntiguoG
                    End If
                    If (.ProximoG > .AntiguoG) Then
                        .ProximoG = .ProximoG - 1
                        If .ProximoG <= .AntiguoG Then .ProximoG = .AntiguoG
                    End If
                    'VERDE
                    
                    'AZUL
                    If (.ProximoB < .AntiguoB) Then
                        .ProximoB = .ProximoB + 1
                        If .ProximoB >= .AntiguoB Then .ProximoB = .AntiguoB
                    End If
                    If (.ProximoB > .AntiguoB) Then
                        .ProximoB = .ProximoB - 1
                        If .ProximoB >= .AntiguoB Then .ProximoB = .AntiguoB
                    End If
                    'AZUL
                    
                    'Ya llegamos al gris, vamos a darle paso al amarillo
                    If .ProximoR = .AntiguoR And .ProximoG = .AntiguoG And .ProximoB = .AntiguoB Then
                        .LlegoAlColor = False
                    End If
                
                End If
                
            End If
    
End With

End Sub
Private Function LeerInt(ByVal Ruta As String) As Integer
Dim f As Integer
    f = FreeFile
    Open Ruta For Input As f
    LeerInt = Input$(LOF(f), #f)
    Close #f
End Function
Public Sub enviarMacro(text As String)

    Dim primerLetra As String
    primerLetra = left(text, 1)
    If (primerLetra = "/") Then
        Call SendData(text)
    Else
        Call SendData(";" & text)
    End If

End Sub
Public Sub actualizarAL(ByVal charindex As Integer)

                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call DibujarPuntoMinimap
                frmMain.Coord.Caption = Nombredelmapaxx & " (" & UserMap & "," & charlist(charindex).Pos.X & "," & charlist(charindex).Pos.Y & ")"

End Sub

Public Sub resetNPCInventory()

    Dim i As Long
    For i = 1 To MAX_NPC_INVENTORY_SLOTS
        NPCInventory(i).GrhIndex = 0
        NPCInventory(i).Amount = 0
        NPCInventory(i).OBJType = 0
        NPCInventory(i).OBJIndex = 0
        NPCInventory(i).Valor = 0
        NPCInventory(i).Name = ""
        NPCInventory(i).MinHit = 0
        NPCInventory(i).MaxHit = 0
        NPCInventory(i).itemSlot = 0
        NPCInventory(i).C1 = 0
        NPCInventory(i).C2 = 0
        NPCInventory(i).C3 = 0
        NPCInventory(i).C4 = 0
        NPCInventory(i).C5 = 0
        NPCInventory(i).C6 = 0
        NPCInventory(i).C7 = 0
    Next i
    
End Sub

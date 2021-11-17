VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Crear Personaje - Tierras Sagradas"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCrearPersonaje.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   840
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   360
   End
   Begin VB.PictureBox headview 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1755
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   7620
      Width           =   675
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":000C
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":0043
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60000
      Width           =   2700
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00DD
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":00E7
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60000
      Width           =   2700
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00FA
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":010D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60000
      Width           =   2700
   End
   Begin VB.Image imgRaza 
      Height          =   420
      Left            =   900
      Top             =   6600
      Width           =   2340
   End
   Begin VB.Image masRaza 
      Height          =   225
      Left            =   3270
      Top             =   6705
      Width           =   150
   End
   Begin VB.Image menosRaza 
      Height          =   225
      Left            =   750
      Top             =   6705
      Width           =   150
   End
   Begin VB.Image Faccion 
      Height          =   675
      Index           =   1
      Left            =   2280
      Top             =   2715
      Width           =   675
   End
   Begin VB.Image Faccion 
      Height          =   675
      Index           =   0
      Left            =   1335
      Top             =   2715
      Width           =   675
   End
   Begin VB.Label lblBarrita 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5895
      TabIndex        =   5
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label txtNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   2055
      Width           =   3975
   End
   Begin VB.Image Genero 
      Height          =   675
      Index           =   1
      Left            =   2265
      Top             =   4710
      Width           =   690
   End
   Begin VB.Image Genero 
      Height          =   675
      Index           =   0
      Left            =   1335
      Top             =   4710
      Width           =   675
   End
   Begin VB.Image Clase 
      Height          =   600
      Index           =   8
      Left            =   9555
      Top             =   4155
      Width           =   600
   End
   Begin VB.Image Clase 
      Height          =   600
      Index           =   7
      Left            =   9555
      Top             =   3405
      Width           =   600
   End
   Begin VB.Image Clase 
      Height          =   600
      Index           =   6
      Left            =   9555
      Top             =   6375
      Width           =   600
   End
   Begin VB.Image Clase 
      Height          =   600
      Index           =   5
      Left            =   9555
      Top             =   7125
      Width           =   600
   End
   Begin VB.Image Clase 
      Height          =   600
      Index           =   4
      Left            =   9555
      Top             =   4890
      Width           =   600
   End
   Begin VB.Image Clase 
      Height          =   600
      Index           =   3
      Left            =   9555
      Top             =   7845
      Width           =   600
   End
   Begin VB.Image Clase 
      Height          =   600
      Index           =   2
      Left            =   9555
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Clase 
      Height          =   600
      Index           =   1
      Left            =   9555
      Top             =   5640
      Width           =   600
   End
   Begin VB.Image menoshead 
      Height          =   225
      Left            =   750
      Top             =   7170
      Width           =   150
   End
   Begin VB.Image mashead 
      Height          =   225
      Left            =   3270
      Top             =   7170
      Width           =   150
   End
   Begin VB.Image boton 
      Height          =   375
      Index           =   0
      Left            =   5040
      MouseIcon       =   "frmCrearPersonaje.frx":013A
      MousePointer    =   99  'Custom
      Top             =   8580
      Width           =   1980
   End
   Begin VB.Image boton 
      Height          =   420
      Index           =   1
      Left            =   60
      MouseIcon       =   "frmCrearPersonaje.frx":028C
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   420
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Const VK_CAPITAL = &H14
Private keys(0 To 255) As Byte

Private textNombrePJ As String
Private resetBotones As Boolean

Private sumaBarrita As Integer

Private razaSelect As Byte
Private claseSelect As Byte
Private generoSelect As String

Public SkillPoints As Byte

Private Sub Faccion_Click(Index As Integer)
    If Index = 0 Then
       Faccion(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_AlianzaOver.jpg")
       Faccion(1).Picture = Nothing
       UserFaccion = 1
    ElseIf Index = 1 Then
       Faccion(0).Picture = Nothing
       Faccion(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_HordaOver.jpg")
       UserFaccion = 2
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        
        If KeyCode = vbKeyBack And Len(textNombrePJ) = 0 Then Exit Sub
        
        If KeyCode = vbKeyReturn Then
            Call boton_Click(0)
        End If
        
        Dim tmpChr As String
        
        If (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode >= 97 And KeyCode <= 122) Or (KeyCode = 32) Or (KeyCode = vbKeyBack) Then
            If KeyCode = vbKeyBack And Len(textNombrePJ) <> 0 Then
                tmpChr = Asc(mid(textNombrePJ, Len(textNombrePJ), Len(textNombrePJ)))
                If (tmpChr >= 65 And tmpChr <= 90) Then
                    sumaBarrita = sumaBarrita - 5
                Else
                    sumaBarrita = sumaBarrita - 4
                End If
                
                
                textNombrePJ = mid(textNombrePJ, 1, Len(textNombrePJ) - 1)
            Else
                If Len(textNombrePJ) >= 15 Then Exit Sub
                tmpChr = Chr(KeyCode)
                
                If GetKeyState(vbKeyCapital) = 0 Then
                    tmpChr = LCase$(tmpChr)
                    sumaBarrita = sumaBarrita + 4
                Else
                    sumaBarrita = sumaBarrita + 5
                End If
                
                textNombrePJ = textNombrePJ & tmpChr  'convert to character
            End If
            
            txtNombre.Caption = textNombrePJ
            lblBarrita.left = sumaBarrita
        End If
        
End Sub
Function CheckData() As Boolean

If UserName = "" Then
    Mensaje.Escribir "Asigne nombre a su personaje."
    Exit Function
End If

If UserRaza = "" Then
    Mensaje.Escribir "Seleccione la raza del personaje."
    Exit Function
End If

If UserFaccion = 0 Then
    Mensaje.Escribir "Seleccione una facción."
    Exit Function
End If
    
    
If UserSexo = "" Then
    Mensaje.Escribir "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    Mensaje.Escribir "Seleccione la clase del personaje."
    Exit Function
End If

CheckData = True


End Function
Private Sub boton_Click(Index As Integer)
    On Error Resume Next
    
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
         Case 0
           
            UserName = textNombrePJ
           
        If Len(textNombrePJ) < 4 Then
            Mensaje.Escribir "El nombre debe de tener más de 4 caracteres."
            Exit Sub
        End If
         
        If Len(textNombrePJ) >= 16 Then
            Mensaje.Escribir "El nombre debe de tener menos de 15 caracteres."
            Exit Sub
        End If
    
        Dim AllCr As Long
        Dim CantidadEsp As Byte
        Dim thiscr As String
        
        Do
            AllCr = AllCr + 1
            If AllCr > Len(UserName) Then Exit Do
            thiscr = mid(UserName, AllCr, 1)
            If InStr(1, " ", thiscr) = 1 Then
                   CantidadEsp = CantidadEsp + 1
            End If
        Loop
        
        If CantidadEsp > 1 Then
             Mensaje.Escribir "El nombre no puede tener mas de 1 espacio."
             Exit Sub
        End If
        
        If claseSelect <= 0 Or claseSelect > 8 Then
            Mensaje.Escribir "Debes seleccionar una clase para tu personaje."
            Exit Sub
        End If
        
        If razaSelect <= 0 Or razaSelect > 5 Then
            Mensaje.Escribir "Debes seleccionar una raza para tu personaje."
            Exit Sub
        End If
        
        If generoSelect <> "Hombre" And generoSelect <> "Mujer" Then
            Mensaje.Escribir "Debes seleccionar un género para tu personaje."
            Exit Sub
        End If

        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                Mensaje.Escribir "Nombre invalido, se han removido los espacios al final del nombre"
        End If
       
        UserRaza = ListaRazas(razaSelect)
        UserSexo = generoSelect
        UserClase = ListaClases(claseSelect)
       
        'Barrin 3/10/03
        If CheckData() Then
            frmMain.Socket1.HostName = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
   
        Me.MousePointer = 11
        
        EstadoLogin = CrearNuevoPj
        
        If Not frmMain.Socket1.Connected Then
              frmMain.Socket1.Connect
          End If
          
              PJClickeado = UserName
              Call Login
          End If
       
        
    Case 1
      Unload Me
End Select


End Sub
Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function
Private Sub Form_Load()


txtNombre.ForeColor = RGB(185, 169, 146)
txtNombre_DblClick
razaSelect = 1
claseSelect = 0
Actualea = 0
actualizarRazas

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_Main.jpg")
boton(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_ButtonN.jpg")


End Sub
Private Sub MenosHead_Click()

Call Audio.PlayWave(SND_CLICK)

Actualea = Actualea - 1
If Actualea = 273 Then Actualea = 272

If Actualea > MaxEleccion Then
Actualea = MaxEleccion

ElseIf Actualea < MinEleccion Then
Actualea = MinEleccion

End If

End Sub
Private Sub MasHead_Click()

Call Audio.PlayWave(SND_CLICK)

Actualea = Actualea + 1
If Actualea = 273 Then Actualea = 274

If Actualea > MaxEleccion Then
    Actualea = MaxEleccion
ElseIf Actualea < MinEleccion Then
    Actualea = MinEleccion
End If

End Sub
Private Sub menoshead_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menoshead.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_IzquierdaOver.jpg")
    resetBotones = True
End Sub
Private Sub mashead_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mashead.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_DerechaOver.jpg")
    resetBotones = True
End Sub
Private Sub menosRaza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menosRaza.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_IzquierdaOver.jpg")
    resetBotones = True
End Sub
Private Sub masRaza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    masRaza.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_DerechaOver.jpg")
    resetBotones = True
End Sub
Private Sub actualizarRazas()
    
    Select Case razaSelect
        Case 1
            imgRaza.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_Humano.jpg")
        Case 2
            imgRaza.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_Elfo.jpg")
        Case 3
            imgRaza.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_ElfoDrow.jpg")
        Case 4
            imgRaza.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_Gnomo.jpg")
        Case 5
            imgRaza.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_Enano.jpg")
    End Select

End Sub
Private Sub menosRaza_Click()
    
    If (razaSelect = 1) Then
        razaSelect = 5
        actualizarRazas
        Exit Sub
    End If
    
    razaSelect = razaSelect - 1
    actualizarRazas
    Call DameOpciones
    
End Sub
Private Sub masRaza_Click()
    
    If (razaSelect = 5) Then
        razaSelect = 1
        actualizarRazas
        Exit Sub
    End If
    
    razaSelect = razaSelect + 1
    actualizarRazas
    Call DameOpciones
    
End Sub
Private Sub Genero_Click(Index As Integer)

If Index = 0 Then
   Genero(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_HombreOver.jpg")
   Genero(1).Picture = Nothing
   generoSelect = "Hombre"
ElseIf Index = 1 Then
   Genero(0).Picture = Nothing
   Genero(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_MujerOver.jpg")
   generoSelect = "Mujer"
End If

Call DameOpciones

End Sub
Sub DameOpciones()
 
Select Case generoSelect

    Case "Hombre"

        Select Case ListaRazas(razaSelect)
    
            Case "Humano"
                Actualea = 1
                MaxEleccion = 30
                MinEleccion = 1
            
            Case "Elfo"
                Actualea = 101
                MaxEleccion = 113
                MinEleccion = 101
                            
            Case "Elfo Oscuro"
                Actualea = 202
                MaxEleccion = 209
                MinEleccion = 202
                            
            Case "Enano"
                Actualea = 301
                MaxEleccion = 305
                MinEleccion = 301
                            
            Case "Gnomo"
                Actualea = 401
                MaxEleccion = 406
                MinEleccion = 401
                            
            Case Else
                Actualea = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
        
    Case "Mujer"
   
        Select Case ListaRazas(razaSelect)
        
            Case "Humano"
                Actualea = 70
                MaxEleccion = 76
                MinEleccion = 70
                            
            Case "Elfo"
                Actualea = 170
                MaxEleccion = 176
                MinEleccion = 170
                            
            Case "Elfo Oscuro"
                Actualea = 270
                MaxEleccion = 280
                MinEleccion = 270
                            
            Case "Gnomo"
                Actualea = 470
                MaxEleccion = 474
                MinEleccion = 470
                            
            Case "Enano"
                Actualea = 370
                MaxEleccion = 373
                MinEleccion = 370
                        
            Case Else
                Actualea = 70
                MaxEleccion = 70
                MinEleccion = 70
                        
        End Select
End Select
 
End Sub
Private Sub Clase_Click(Index As Integer)

    claseSelect = Index
    actualizarClase
        
End Sub
Private Sub actualizarClase()
    Clase(1).Picture = Nothing
    Clase(2).Picture = Nothing
    Clase(3).Picture = Nothing
    Clase(4).Picture = Nothing
    Clase(5).Picture = Nothing
    Clase(6).Picture = Nothing
    Clase(7).Picture = Nothing
    Clase(8).Picture = Nothing
    
    If claseSelect > 0 Then Clase(claseSelect) = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_" & ListaClases(claseSelect) & "Over.jpg")
    
End Sub
Private Sub Timer1_Timer()
    If lblBarrita.Caption = "" Then
        lblBarrita.Caption = "|"
    Else
        lblBarrita.Caption = ""
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If resetBotones Then
        mashead.Picture = Nothing
        menoshead.Picture = Nothing
        masRaza.Picture = Nothing
        menosRaza.Picture = Nothing
    
        boton(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_ButtonN.jpg")
        boton(1).Picture = Nothing
    End If
End Sub
Private Sub boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then boton(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_ButtonI.jpg")
    If Index = 1 Then boton(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_ReturnOver.jpg")
    resetBotones = True
End Sub
Private Sub boton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then boton(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearPersonaje_ButtonA.jpg")
    resetBotones = True
End Sub

Private Sub Timer2_Timer()
    engine.drawCabezas
End Sub

Private Sub txtNombre_DblClick()
    txtNombre.Caption = ""
    textNombrePJ = ""
    sumaBarrita = 393
    lblBarrita.left = sumaBarrita
End Sub

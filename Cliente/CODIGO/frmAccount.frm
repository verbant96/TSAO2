VERSION 5.00
Begin VB.Form frmAccount 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Tierras Sagradas"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   788.177
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   7545
      Left            =   4217
      Picture         =   "frmAccount.frx":000C
      ScaleHeight     =   503
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   238
      TabIndex        =   2
      Top             =   600
      Width           =   3563
      Begin VB.Image imgCrearPersonaje 
         Height          =   435
         Left            =   0
         Top             =   7050
         Width           =   3570
      End
      Begin VB.Image img_BorrarPJ 
         Height          =   330
         Left            =   3000
         Top             =   1080
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   1
         Left            =   0
         Top             =   885
         Width           =   3570
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   0
         Left            =   0
         MousePointer    =   99  'Custom
         Top             =   120
         Width           =   3570
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   2
         Left            =   0
         Top             =   1635
         Width           =   3570
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   3
         Left            =   0
         Top             =   2385
         Width           =   3570
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   5
         Left            =   0
         Top             =   3885
         Width           =   3570
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   4
         Left            =   0
         Top             =   3135
         Width           =   3570
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   6
         Left            =   0
         Top             =   4635
         Width           =   3570
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   7
         Left            =   0
         Top             =   5385
         Width           =   3570
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   8
         Left            =   0
         Top             =   6135
         Width           =   3570
      End
      Begin VB.Image PJ 
         Height          =   735
         Index           =   9
         Left            =   0
         Top             =   6885
         Width           =   3570
      End
   End
   Begin VB.PictureBox picChar 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   8663
      Picture         =   "frmAccount.frx":1B48
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   0
      Top             =   4395
      Width           =   1857
   End
   Begin VB.Label lblInformacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   8760
      TabIndex        =   4
      Top             =   3810
      Width           =   1695
   End
   Begin VB.Label lblNoticias 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   2415
      Left            =   1440
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblCambiarPassword 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CAMBIAR CONTRASEÑA"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   4765
      TabIndex        =   1
      Top             =   8745
      Width           =   2535
   End
   Begin VB.Image imgPaginaWeb 
      Height          =   480
      Left            =   1568
      Top             =   5955
      Width           =   1650
   End
   Begin VB.Image img_EntrarPJ 
      Height          =   480
      Left            =   8785
      Top             =   5955
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Image imgSalir4 
      Height          =   420
      Left            =   120
      MousePointer    =   99  'Custom
      Picture         =   "frmAccount.frx":20FE
      Top             =   8610
      Width           =   420
   End
   Begin VB.Image imgCambiarPass 
      Height          =   375
      Left            =   4920
      Top             =   8640
      Width           =   2220
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const WS_EX_APPWINDOW               As Long = &H40000
Private Const GWL_EXSTYLE                   As Long = (-20)
Private Const SW_HIDE                       As Long = 0
Private Const SW_SHOW                       As Long = 5
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Dim i As Long
Dim BorrarRandom As String
Dim ElRandom As String

Private m_bActivated As Boolean
 
Private Sub Form_Activate()
    If Not m_bActivated Then
        m_bActivated = True
        Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
        Call ShowWindow(hWnd, SW_HIDE)
        Call ShowWindow(hWnd, SW_SHOW)
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If PJApretado = 99 Then Exit Sub
        SendData ("OOLOGI" & CargarPJ(PJApretado).Nombre & "," & nombrecuent & "," & CodigoRecibido)
        Exit Sub
    End If
End Sub
Private Sub restartButtons()

    For i = 0 To 8
        ButtonPJHover(i) = False
    Next i
    
    ButtonCPHover = False
    ButtonDeleteCharHover = False

    img_EntrarPJ.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\PanelCuenta_Jugar.jpg")
    imgPaginaWeb.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\PanelCuenta_Web.jpg")
    'imgCrearPersonaje.Picture = Nothing
    'img_BorrarPJ.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\PanelCuenta_Borrar.jpg")
    lblCambiarPassword.ForeColor = RGB(161, 148, 128)
    
End Sub
Private Sub Form_Load()
    
    restartButtons
    PJApretado = 99
    
    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\PanelCuenta_Main.jpg")
    
    If CantidadDePersonajes < 10 Then imgCrearPersonaje.Visible = True
    
    img_BorrarPJ.Enabled = False
    img_BorrarPJ.Visible = True
    img_EntrarPJ.Visible = True
    lblInformacion.ForeColor = RGB(175, 159, 112)

End Sub
Private Sub imgCrearPersonaje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCPHover = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    restartButtons
End Sub

Private Sub lblCambiarPassword_Click()
    Call imgCambiarPass_Click
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    restartButtons
End Sub
Private Sub picChar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    restartButtons
End Sub
Private Sub img_EntrarPJ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    img_EntrarPJ.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\PanelCuenta_JugarHover.jpg")
End Sub
Private Sub img_EntrarPJ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    img_EntrarPJ.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\PanelCuenta_JugarPress.jpg")
End Sub
Private Sub imgPaginaWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPaginaWeb.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\PanelCuenta_WebHover.jpg")
End Sub
Private Sub imgPaginaWeb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPaginaWeb.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\PanelCuenta_WebPress.jpg")
    OpenBrowser "http://www.tierras-sagradas.com/", 4
End Sub

Private Sub img_BorrarPJ_Click()
      BorrarRandom = RandomNumber(1000, 9999)
      ElRandom = InputBox("Esta accion no podra ser revertida, para confirmar ingrse el codigo " & BorrarRandom & " para borrar su personaje.", "Borrar Personaje")
        
      If BorrarRandom = ElRandom Then Call SendData("TBRP" & CargarPJ(PJApretado).Nombre & "," & nombrecuent & "," & CodigoRecibido)
End Sub
Private Sub img_EntrarPJ_Click()
    If PJApretado = 99 Then Exit Sub
    SendData ("OOLOGI" & CargarPJ(PJApretado).Nombre & "," & nombrecuent & "," & CodigoRecibido)
End Sub
Private Sub PJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'aca dibujaría la iluminación
    restartButtons
    ButtonPJHover(Index) = True
    
End Sub
Private Sub PJ_Click(Index As Integer)

    If CargarPJ(Index).Existe = True Then
        If Not img_BorrarPJ.Enabled Then img_BorrarPJ.Enabled = True
        img_BorrarPJ.top = 22 + (50 * Index)
        PJApretado = Index
        
        lblInformacion.Caption = CargarPJ(Index).Nombre & " (Nivel " & CargarPJ(Index).Level & ")" & vbCrLf & "<" & CargarPJ(Index).Clase & " " & CargarPJ(Index).Raza & ">"
    End If

End Sub
Private Sub PJ_DblClick(Index As Integer)

    SendData ("OOLOGI" & CargarPJ(Index).Nombre & "," & nombrecuent & "," & CodigoRecibido)

End Sub
Private Sub imgCrearPersonaje_Click()
    On Error Resume Next
    If CargarPJ(9).Existe = True Then
        Mensaje.Escribir "No puedes crear más personajes."
    Else
        Call Audio.PlayWave("click.wav")
        EstadoLogin = Dados
        frmCrearPersonaje.Show , frmAccount
        Audio.StopWave
    End If
End Sub
Private Sub imgSalir4_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmMain.Socket1.Disconnect
    frmMain.Socket1.Cleanup
    
    AoDefResult = 0
    Unload Me
    frmConnect.mostrarConectar (True)
End Sub
Private Sub imgCambiarPass_Click()
    On Error Resume Next
        Call Audio.PlayWave("click.wav")
        Dim anteriorpw As String
        Dim nuevapw As String
        Dim renuevapw As String
        
        anteriorpw = InputBox("Ingrese su actual contraseña:", "Cambiar Password")
        nuevapw = InputBox("Ingrese su nueva contraseña:", "Cambiar Password")
        renuevapw = InputBox("Repita su nueva contraseña:", "Cambiar Password")
        
        If nuevapw <> renuevapw Then
            Mensaje.Escribir "Las passwords que tipeo no coinciden"
            Exit Sub
        End If
        
        If Len(nuevapw) > 15 Then
            Mensaje.Escribir "La password no puede superar los 15 caracteres"
            Exit Sub
        End If
        
        SendData ("REPASS" & nombrecuent & "," & anteriorpw & "," & nuevapw & "," & renuevapw)
End Sub
Private Sub imgSalir4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ButtonSalir = "Iluminado"
End Sub
Private Sub imgSalir4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ButtonSalir = "Apretado"
End Sub
Private Sub imgCambiarPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCambiarPassword.ForeColor = RGB(69, 194, 234)
End Sub
Private Sub img_BorrarPJ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonDeleteCharHover = True
End Sub

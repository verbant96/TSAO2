VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNuevoComercio 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   Icon            =   "frmComercioUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   596
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Can 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4245
      TabIndex        =   10
      Text            =   "1"
      Top             =   6225
      Width           =   1635
   End
   Begin VB.ListBox Ofrecer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4140
      IntegralHeight  =   0   'False
      Left            =   6540
      TabIndex        =   8
      Top             =   1890
      Width           =   2670
   End
   Begin VB.ListBox TusItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4140
      IntegralHeight  =   0   'False
      Left            =   3600
      TabIndex        =   7
      Top             =   1890
      Width           =   2670
   End
   Begin VB.ListBox Oferta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4140
      IntegralHeight  =   0   'False
      ItemData        =   "frmComercioUsuario.frx":000C
      Left            =   390
      List            =   "frmComercioUsuario.frx":000E
      TabIndex        =   5
      Top             =   1890
      Width           =   2895
   End
   Begin VB.TextBox Texto 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   405
      MaxLength       =   150
      TabIndex        =   3
      Text            =   "Escribi aca tu mensaje para el otro usuario y apreta la tecla 'Enter' o clickea en Enviar"
      Top             =   8480
      Width           =   7020
   End
   Begin VB.PictureBox picInv 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   240
      ScaleHeight     =   510
      ScaleWidth      =   540
      TabIndex        =   1
      Top             =   680
      Width           =   540
   End
   Begin RichTextLib.RichTextBox Consola 
      Height          =   780
      Left            =   360
      TabIndex        =   2
      Top             =   7560
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1376
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmComercioUsuario.frx":0010
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label conQuien 
      BackStyle       =   0  'Transparent
      Caption         =   "Con tu vieja"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Command2 
      Height          =   300
      Left            =   6285
      Top             =   3960
      Width           =   255
   End
   Begin VB.Image Command1 
      Height          =   300
      Left            =   6285
      Top             =   3660
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   6180
      Width           =   1785
   End
   Begin VB.Image Command3 
      Height          =   420
      Left            =   6960
      Top             =   6615
      Width           =   1845
   End
   Begin VB.Image cmdAgregarOro 
      Height          =   420
      Left            =   4020
      Top             =   6615
      Width           =   1845
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Envia tu oferta de oro o items al otro usuario."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   720
      Width           =   8175
   End
   Begin VB.Image Res 
      Height          =   420
      Index           =   1
      Left            =   1920
      Top             =   6615
      Width           =   1395
   End
   Begin VB.Image Res 
      Height          =   420
      Index           =   2
      Left            =   345
      Top             =   6615
      Width           =   1395
   End
   Begin VB.Label lblOro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   6180
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   7620
      Top             =   8445
      Width           =   1635
   End
   Begin VB.Label Image2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmNuevoComercio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarOro_Click()

If Not IsNumeric(Can.text) Then Exit Sub
If UserGLD < Can.text Then lblEstado.Caption = "No tienes suficiente oro.": Exit Sub

If Can.text > 0 And Can.text <= 999999999 Then
    Label4.Caption = PonerPuntos(Can.text)
    uOro = Can.text
End If

End Sub
Private Sub Command1_Click()
comAgregarOferta TusItems.ListIndex, Val(Can.text)
End Sub
Private Sub Command2_Click()
    comQuitarOferta Ofrecer.ListIndex, Val(Can.text)
End Sub
Private Sub Command3_Click()
If uOro = "0" And Ofrecer.ListCount = 0 Then
    Mensaje.Escribir "Hace una oferta primero."
Exit Sub
End If
    'Desahablita
    Command3.Enabled = False
    'Cambia lbl
    lblEstado.Caption = "Enviando Ofertas al otro usuario y esperando respuesta..."
    comEnviarOferta
End Sub
Private Sub Form_Load()
    comMensaje "Bienvenido al nuevo sistema de comercio de TSAO, elija los items y haga su oferta, responda la del otro usuario y listo!, para usar el chat escriba el mensaje y aprete ""ENVIAR"" o tan solo ""ENTER"" y listo.", 255, 255, 0, 0, 0
    lblOro.Caption = "0"
    Label4.Caption = "0"
    uOro = 0
    rOro = 0
    
    Oferta.BackColor = RGB(19, 21, 23)
    TusItems.BackColor = RGB(19, 21, 23)
    Ofrecer.BackColor = RGB(19, 21, 23)
    Can.BackColor = RGB(19, 21, 23)
    Consola.BackColor = RGB(19, 21, 23)
    Texto.BackColor = RGB(19, 21, 23)
    
    lblOro.ForeColor = RGB(145, 123, 85)
    Label4.ForeColor = RGB(145, 123, 85)
    Oferta.ForeColor = RGB(145, 123, 85)
    TusItems.ForeColor = RGB(145, 123, 85)
    Ofrecer.ForeColor = RGB(145, 123, 85)
    Can.ForeColor = RGB(145, 123, 85)
    Texto.ForeColor = RGB(145, 123, 85)

    Set form_Moviment = New clsFormMovementManager
    form_Moviment.Initialize Me
    
    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\ComercioPJ.jpg")
    

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Enabled = True Then Image1.Picture = General_Load_Interface_Picture("ComercioPJ_Enviar.jpg")
If Command3.Enabled = True Then Command3.Picture = General_Load_Interface_Picture("ComercioPJ_Ofrecer.jpg")
If cmdAgregarOro.Enabled = True Then cmdAgregarOro.Picture = General_Load_Interface_Picture("ComercioPJ_Oro.jpg")
If Command1.Enabled = True Then Command1.Picture = General_Load_Interface_Picture("ComercioPJ_Flecha_Derecha.jpg")
If Command2.Enabled = True Then Command2.Picture = General_Load_Interface_Picture("ComercioPJ_Flecha_Izquierda.jpg")

If Res(2).Enabled = True Then Res(2).Picture = General_Load_Interface_Picture("ComercioPJ_Aceptar.jpg")
If Res(1).Enabled = True Then Res(1).Picture = General_Load_Interface_Picture("ComercioPJ_Rechazar.jpg")
End Sub
Private Sub Image1_Click()
SendData "VHC" & Texto.text
comMensaje UserName & " >> " & Texto.text, 255, 255, 255, True, False, False
Texto.text = ""
End Sub
Private Sub Image2_Click()
SendData "TCM"
End Sub
Private Sub Oferta_Click()
comDibujarRec Oferta.ListIndex
End Sub
Private Sub Ofrecer_Click()
comDibujarOfe
End Sub
Private Sub Res_Click(Index As Integer)
lblEstado.Caption = "Oferta del otro usuario aceptada."
comRespuesta Index
End Sub
Private Sub Can_Change()
On Error GoTo errHandler
    If Val(Can.text) < 0 Then Can.text = 10000
Exit Sub
errHandler:
    Can.text = "1"
End Sub
Private Sub Can_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
If (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End If
End Sub
Private Sub Texto_Click()

If Texto.text = "Escribi aca tu mensaje para el otro usuario y apreta la tecla 'Enter' o clickea en Enviar" Then
Texto.text = ""
End If

End Sub
Private Sub Texto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
SendData "VHC" & Texto.text
comMensaje "" & UserName & "> " & Texto.text, 255, 0, 0
Texto.text = ""
End If
End Sub
Private Sub TusItems_Click()
comDibujarTusItems TusItems.ListIndex
End Sub


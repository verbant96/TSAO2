VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FranComercio 
   BorderStyle     =   0  'None
   Caption         =   "Comercio"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmNuevoComercio.frx":0000
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   742
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicRe 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   8085
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   4830
      Width           =   510
   End
   Begin VB.PictureBox PicOf 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2970
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   4830
      Width           =   510
   End
   Begin VB.PictureBox PicMis 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2535
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   1365
      Width           =   510
   End
   Begin VB.TextBox Texto 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF80&
      Height          =   210
      Left            =   75
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   7680
      Width           =   10395
   End
   Begin VB.TextBox Can 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Text            =   "1"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.ListBox Oferta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   4230
      ItemData        =   "FrmNuevoComercio.frx":499F1
      Left            =   8640
      List            =   "FrmNuevoComercio.frx":499F3
      TabIndex        =   2
      Top             =   1215
      Width           =   2415
   End
   Begin VB.ListBox Ofrecer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   4230
      ItemData        =   "FrmNuevoComercio.frx":499F5
      Left            =   3525
      List            =   "FrmNuevoComercio.frx":499F7
      TabIndex        =   1
      Top             =   1215
      Width           =   2415
   End
   Begin VB.ListBox TusItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   4230
      ItemData        =   "FrmNuevoComercio.frx":499F9
      Left            =   75
      List            =   "FrmNuevoComercio.frx":499FB
      TabIndex        =   0
      Top             =   1215
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox Consola 
      CausesValidation=   0   'False
      Height          =   1080
      Left            =   75
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   6525
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1905
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmNuevoComercio.frx":499FD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   10680
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   10440
      Top             =   7680
      Width           =   735
   End
   Begin VB.Image Command3 
      Height          =   375
      Left            =   3720
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Image Res 
      Height          =   375
      Index           =   1
      Left            =   8880
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Image Res 
      Height          =   255
      Index           =   2
      Left            =   8880
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Image Command2 
      Height          =   375
      Left            =   2520
      Top             =   3480
      Width           =   975
   End
   Begin VB.Image Command1 
      Height          =   375
      Left            =   2520
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label conQuien 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   135
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Left            =   12600
      TabIndex        =   3
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "FranComercio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
comAgregarOferta TusItems.ListIndex, Val(Can.text)
End Sub
Private Sub Command2_Click()
comQuitarOferta Ofrecer.ListIndex, Val(Can.text)
End Sub
Private Sub Command3_Click()
comEnviarOferta
End Sub
Private Sub Form_Load()
comMensaje "Servidor> Bienvenido al nuevo sistema de Comercio de Khelendor, elija los items y haga su oferta, responda la del otro usuario y listo!, para usar el chat escriba el mensaje y aprete ""ENVIAR"" o tan solo ""ENTER"" y listo.", 255, 255, 255
comMensaje "Servidor> Comerciando con " & cNombre, 255, 255, 255
End Sub
Private Sub Image1_Click()
SendData "VHC" & Texto.text
comMensaje "Dijiste> " & Texto.text, 255, 0, 0
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
comRespuesta Index
End Sub
Private Sub Can_Change()
On Error GoTo ErrHandler
    If Val(Can.text) < 0 Then Can.text = 10000
    If Val(Can.text) > 10000 Then Can.text = "1"
Exit Sub
ErrHandler:
    Can.text = "1"
End Sub
Private Sub Can_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
If (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End If
End Sub
Private Sub Texto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
SendData "VHC" & Texto.text
comMensaje "Dijiste> " & Texto.text, 255, 0, 0
Texto.text = ""
End If
End Sub
Private Sub TusItems_Click()
comDibujarTusItems TusItems.ListIndex
End Sub

VERSION 5.00
Begin VB.Form frmCorreo 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "HOLA"
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   Icon            =   "frmCorreo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   466
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5475
      MaxLength       =   5
      TabIndex        =   12
      Text            =   "1"
      Top             =   5355
      Visible         =   0   'False
      Width           =   700
   End
   Begin VB.ListBox lstObjsEnviar 
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
      Height          =   1785
      Left            =   6300
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox lstObjs 
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
      Height          =   1785
      Left            =   2700
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtMensaje 
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
      Height          =   1890
      Left            =   2715
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1575
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.TextBox txtAsunto 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2715
      MaxLength       =   20
      TabIndex        =   8
      Top             =   900
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.TextBox txtDestinatario 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   195
      MaxLength       =   15
      TabIndex        =   7
      Top             =   900
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.ListBox lstContactos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4080
      ItemData        =   "frmCorreo.frx":000C
      Left            =   165
      List            =   "frmCorreo.frx":004C
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstObjetos 
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
      Height          =   1785
      Left            =   2700
      TabIndex        =   5
      Top             =   3645
      Width           =   2500
   End
   Begin VB.TextBox lblMensaje 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   1860
      Left            =   2715
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmCorreo.frx":0097
      Top             =   1395
      Width           =   6225
   End
   Begin VB.ListBox lstMails 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   5895
      IntegralHeight  =   0   'False
      ItemData        =   "frmCorreo.frx":02C7
      Left            =   120
      List            =   "frmCorreo.frx":0325
      TabIndex        =   0
      Top             =   660
      Width           =   2445
   End
   Begin VB.Image cmdSalir2 
      Height          =   495
      Left            =   8640
      Top             =   0
      Width           =   465
   End
   Begin VB.Image cmdQui 
      Height          =   270
      Left            =   5610
      Top             =   4755
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image cmdAdd 
      Height          =   375
      Left            =   5625
      Top             =   4410
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image cmdRetirar 
      Height          =   330
      Left            =   2700
      Top             =   5625
      Width           =   2490
   End
   Begin VB.Image cmdNuevo 
      Height          =   435
      Left            =   4320
      Top             =   6180
      Width           =   2970
   End
   Begin VB.Image cmdSalir 
      Height          =   495
      Left            =   8640
      Top             =   0
      Width           =   495
   End
   Begin VB.Image cmdGuardar 
      Height          =   570
      Left            =   5355
      Top             =   5385
      Width           =   2625
   End
   Begin VB.Image cmdBorrar 
      Height          =   570
      Left            =   5355
      Top             =   4560
      Width           =   2625
   End
   Begin VB.Image cmdResponder 
      Height          =   570
      Left            =   5355
      Top             =   3705
      Width           =   2625
   End
   Begin VB.Label lblAsunto 
      BackStyle       =   0  'Transparent
      Caption         =   "Hola si dame oro plis"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   975
      Width           =   2535
   End
   Begin VB.Label lblFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "24/11/2808 16:48"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7275
      TabIndex        =   2
      Top             =   630
      Width           =   1935
   End
   Begin VB.Label lblRemitente 
      BackStyle       =   0  'Transparent
      Caption         =   "SuperMiniGnomo"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4140
      TabIndex        =   1
      Top             =   630
      Width           =   1695
   End
   Begin VB.Image cmdSend 
      Height          =   330
      Left            =   4980
      Top             =   5850
      Visible         =   0   'False
      Width           =   1680
   End
End
Attribute VB_Name = "frmCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
correosAgregarItem lstObjs.ListIndex, Val(txtAmount.text)
End Sub
Private Sub cmdQui_Click()
    correosQuitarItem lstObjsEnviar.ListIndex, Val(txtAmount.text)
End Sub
Private Sub cmdRetirar_Click()
    Call SendData("CZR" & lstMails.ListIndex + 1)
End Sub

Private Sub cmdSend_Click()

If Len(frmCorreo.txtDestinatario.text) < 3 Then
    Mensaje.Escribir ("El destinatario debe tener al minimo 3 letras.")
    Exit Sub
End If

If Len(frmCorreo.txtAsunto.text) < 10 And Len(frmCorreo.txtMensaje.text) < 10 Then
    Mensaje.Escribir ("El asunto y el mensaje deben tener un minimo de 10 letras.")
    Exit Sub
End If

correosEnviarItems
Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdResponder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RespMsj_N.jpg")
cmdBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_BorrarMsj_N.jpg")
cmdGuardar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_GuardarMsj_N.jpg")
cmdRetirar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RetirarObj_N.jpg")
cmdNuevo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_NuevoMsj_N.jpg")
cmdSend.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Enviar_N.jpg")
End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_Main.jpg")
cmdResponder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RespMsj_N.jpg")
cmdBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_BorrarMsj_N.jpg")
cmdGuardar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_GuardarMsj_N.jpg")
cmdRetirar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RetirarObj_N.jpg")
cmdNuevo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_NuevoMsj_N.jpg")

lblRemitente.Caption = ""
lblRemitente.Visible = True
lblFecha.Caption = ""
lblFecha.Visible = True
lblMensaje.text = ""
lblMensaje.Visible = True
lblAsunto.Caption = ""
lblAsunto.Visible = True
lstMails.Visible = True
lstObjetos.Clear
lstObjetos.Visible = True
cmdResponder.Visible = True
cmdBorrar.Visible = True
cmdGuardar.Visible = True
cmdNuevo.Visible = True
cmdRetirar.Visible = True

lstObjetos.BackColor = RGB(19, 21, 23)
lblMensaje.BackColor = RGB(19, 21, 23)
lstMails.BackColor = RGB(19, 21, 23)

lstObjetos.ForeColor = RGB(145, 123, 85)
lblMensaje.ForeColor = RGB(145, 123, 85)
lstMails.ForeColor = RGB(145, 123, 85)
lblAsunto.ForeColor = RGB(145, 123, 85)
lblFecha.ForeColor = RGB(145, 123, 85)
lblRemitente.ForeColor = RGB(145, 123, 85)

cmdSend.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Enviar_N.jpg")

'Sacamos todo
cmdSend.Visible = False
cmdAdd.Visible = False
cmdQui.Visible = False
lstObjs.Visible = False
lstObjsEnviar.Visible = False
lstContactos.Visible = False
txtMensaje.Visible = False
txtAsunto.Visible = False
txtDestinatario.Visible = False
txtAmount.Visible = False

End Sub
Private Sub cmdBorrar_Click()
Call SendData("CZB" & lstMails.ListIndex + 1)
End Sub
Private Sub cmdNuevo_Click()

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Main.jpg")

'Ponemos todo en false
lblAsunto.Visible = True
lblRemitente.Visible = False
lblFecha.Visible = False
lblMensaje.Visible = False
lstMails.Visible = False
lstObjetos.Visible = False
cmdResponder.Visible = False
cmdBorrar.Visible = False
cmdGuardar.Visible = False
cmdNuevo.Visible = False
cmdRetirar.Visible = False

'y en true
cmdSend.Visible = True
cmdAdd.Visible = True
cmdQui.Visible = True

lstObjs.Visible = True
lstObjsEnviar.Visible = True
lstContactos.Visible = True

lstObjs.BackColor = RGB(19, 21, 23)
lstObjsEnviar.BackColor = RGB(19, 21, 23)
lstContactos.BackColor = RGB(19, 21, 23)
txtMensaje.BackColor = RGB(19, 21, 23)
txtAsunto.BackColor = RGB(19, 21, 23)
txtDestinatario.BackColor = RGB(19, 21, 23)
txtAmount.BackColor = RGB(19, 21, 23)

lstObjs.ForeColor = RGB(145, 123, 85)
lstObjsEnviar.ForeColor = RGB(145, 123, 85)
lstContactos.ForeColor = RGB(145, 123, 85)
txtMensaje.ForeColor = RGB(145, 123, 85)
txtAsunto.ForeColor = RGB(145, 123, 85)
txtDestinatario.ForeColor = RGB(145, 123, 85)
txtAmount.ForeColor = RGB(145, 123, 85)

txtMensaje.Visible = True
txtMensaje.text = ""

txtAsunto.Visible = True
txtAsunto.text = ""

txtDestinatario.Visible = True
txtDestinatario.text = ""

txtAmount.Visible = True
txtAmount.text = "1"


End Sub
Private Sub cmdResponder_Click()

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Main.jpg")

'Ponemos todo en false
lblAsunto.Visible = True
lblRemitente.Visible = False
lblFecha.Visible = False
lblMensaje.Visible = False
lstMails.Visible = False
lstObjetos.Visible = False
cmdResponder.Visible = False
cmdBorrar.Visible = False
cmdGuardar.Visible = False
cmdNuevo.Visible = False
cmdRetirar.Visible = False

'y en true
cmdSend.Visible = True
cmdAdd.Visible = True
cmdQui.Visible = True

lstObjs.Visible = True
lstObjsEnviar.Visible = True
lstContactos.Visible = True

txtMensaje.Visible = True
txtMensaje.text = ""

txtAsunto.Visible = True
txtAsunto.text = "RE " & lstMails.List(lstMails.ListIndex)

txtDestinatario.Visible = True
txtDestinatario.text = lstMails.List(lstMails.ListIndex)

txtAmount.Visible = True
txtAmount.text = "1"


End Sub
Private Sub cmdSalir2_Click()
    correosCerrar
End Sub
Private Sub lstContactos_Click()
txtDestinatario.text = UCase$(lstContactos.List(lstContactos.ListIndex))
End Sub
Private Sub lstMails_Click()
If lstMails.ListIndex + 1 = 0 Then Exit Sub
If lstMails.ListIndex = CorreoListIndex Then Exit Sub

lblMensaje.text = ""
lblAsunto.Caption = ""
lblFecha.Caption = ""
lblRemitente.Caption = ""
lstObjetos.Clear

CorreoListIndex = lstMails.ListIndex
Call SendData("CZC" & lstMails.ListIndex + 1)

End Sub
Private Sub cmdResponder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdResponder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RespMsj_I.jpg")
End Sub
Private Sub cmdResponder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdResponder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RespMsj_A.jpg")
End Sub
Private Sub cmdBorrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_BorrarMsj_I.jpg")
End Sub
Private Sub cmdBorrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_BorrarMsj_A.jpg")
End Sub
Private Sub cmdGuardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdGuardar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_GuardarMsj_I.jpg")
End Sub
Private Sub cmdGuardar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdGuardar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_GuardarMsj_A.jpg")
End Sub
Private Sub cmdRetirar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdRetirar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RetirarObj_I.jpg")
End Sub
Private Sub cmdRetirar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdRetirar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RetirarObj_A.jpg")
End Sub
Private Sub cmdNuevo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNuevo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_NuevoMsj_I.jpg")
End Sub
Private Sub cmdNuevo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNuevo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_NuevoMsj_A.jpg")
End Sub
Private Sub cmdSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSend.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Enviar_I.jpg")
End Sub
Private Sub cmdSend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSend.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Enviar_A.jpg")
End Sub

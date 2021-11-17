VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sistema de Canjeo"
   ClientHeight    =   5625
   ClientLeft      =   420
   ClientTop       =   315
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "õ.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3840
      TabIndex        =   9
      Text            =   "1"
      Top             =   1600
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   2925
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   8
      Top             =   1300
      Width           =   465
   End
   Begin VB.TextBox lDescripcion 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   800
      Left            =   2950
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3565
      Width           =   1650
   End
   Begin VB.ListBox ListaPremios 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   4140
      IntegralHeight  =   0   'False
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2350
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   4320
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lPuntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   720
      Width           =   765
   End
   Begin VB.Label lAtaque 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3800
      TabIndex        =   5
      Top             =   1975
      Width           =   855
   End
   Begin VB.Label lDef 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3800
      TabIndex        =   4
      Top             =   2355
      Width           =   855
   End
   Begin VB.Label lAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3800
      TabIndex        =   3
      Top             =   2725
      Width           =   855
   End
   Begin VB.Label lDM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3800
      TabIndex        =   2
      Top             =   3100
      Width           =   855
   End
   Begin VB.Label Requiere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   4600
      Width           =   1005
   End
   Begin VB.Image bAceptar 
      Height          =   435
      Left            =   360
      Picture         =   "õ.frx":12807
      Top             =   5000
      Width           =   2355
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
    If MsgBox("¿Estás seguro que deseas canjear " & lCantidad.text & " - " & ListaPremios.text & "?", vbYesNo) = vbYes Then
        Call SendData("SPX" & ListaPremios.ListIndex + 1 & "," & lCantidad.text)
    End If
End Sub
Private Sub Image1_Click()
Unload Me
End Sub
Private Sub lCantidad_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
 Case Asc("0") To Asc("9"), vbKeyDelete, vbKeyBack
Case Else: KeyAscii = 0
 End Select
End Sub
Private Sub lCantidad_Change()

If lCantidad.text = "" Then lCantidad.text = 1
If Not IsNumeric(lCantidad.text) Or lCantidad.text < 0 Or lCantidad.text > 10000 Then lCantidad.text = 1
Requiere.Caption = CantidadCanjeYegua * lCantidad.text

End Sub

Private Sub ListaPremios_Click()
Call SendData("IPX" & ListaPremios.ListIndex + 1)
End Sub
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Call SendData("IPX" & ListaPremios.ListIndex + 1)
bAceptar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Canjear_BcanjearN.jpg")
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Canjear_main.jpg")
End Sub
Private Sub baceptar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bAceptar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Canjear_BcanjearA.jpg")
End Sub
Private Sub baceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bAceptar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Canjear_BcanjearI.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bAceptar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Canjear_BcanjearN.jpg")
End Sub


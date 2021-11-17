VERSION 5.00
Begin VB.Form frmDonador 
   BorderStyle     =   0  'None
   Caption         =   "Premios"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   3780
   End
   Begin VB.TextBox txtPuntos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   420
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   885
      Width           =   2700
   End
   Begin VB.TextBox txtValor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1350
      Width           =   1980
   End
   Begin VB.ListBox lstPremios 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3150
      ItemData        =   "frmDonador.frx":0000
      Left            =   360
      List            =   "frmDonador.frx":005E
      TabIndex        =   1
      Top             =   1551
      Width           =   2700
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1865
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2160
      Width           =   3780
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   6840
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgCanjear 
      Height          =   615
      Left            =   3240
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3120
      Top             =   5160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   9000
      Top             =   5160
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   11040
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmDonador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Call SendData("DPX" & lstPremios.ListIndex + 1)
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Donaciones_Main.jpg")
imgCanjear.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Donaciones_Canjear_N.jpg")
End Sub
Private Sub imgCanjear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCanjear.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Donaciones_Canjear_I.jpg")
End Sub
Private Sub imgCanjear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCanjear.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Donaciones_Canjear_A.jpg")
End Sub
Private Sub imgCanjear_Click()
    If MsgBox("¿Estás seguro que deseas canjear " & lstPremios.text & "?", vbYesNo) = vbYes Then
        Call SendData("DRX" & lstPremios.ListIndex + 1)
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCanjear.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Donaciones_Canjear_N.jpg")
End Sub
Private Sub imgSalir_Click()
Unload Me
End Sub
Private Sub lstPremios_Click()
Call SendData("DPX" & lstPremios.ListIndex + 1)
End Sub

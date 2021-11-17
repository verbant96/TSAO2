VERSION 5.00
Begin VB.Form frmNewTiendaTS 
   BorderStyle     =   0  'None
   Caption         =   "Tienda TS"
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   LinkTopic       =   "Form2"
   Picture         =   "frmNewTiendaTS.frx":0000
   ScaleHeight     =   4740
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image cmdSalir 
      Height          =   255
      Left            =   6600
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblTSPoints 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   4280
      Width           =   975
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   11
      Left            =   5420
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   10
      Left            =   3190
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   9
      Left            =   915
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   8
      Left            =   5420
      Top             =   2875
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   7
      Left            =   3190
      Top             =   2875
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   6
      Left            =   915
      Top             =   2875
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   5
      Left            =   5420
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   4
      Left            =   3190
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   3
      Left            =   915
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   2
      Left            =   5420
      Top             =   1145
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   1
      Left            =   3190
      Top             =   1155
      Width           =   1215
   End
   Begin VB.Image tsCanje 
      Height          =   255
      Index           =   0
      Left            =   915
      Top             =   1155
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewTiendaTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()

    Set form_Moviment = New clsFormMovementManager
    form_Moviment.Initialize Me

    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Main.jpg")
    tsCanje(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Champ.jpg")
    tsCanje(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Luz.jpg")
    tsCanje(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_CopadeOro.jpg")
    tsCanje(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Heroes.jpg")
    tsCanje(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Oscu.jpg")
    tsCanje(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Perseus.jpg")
    tsCanje(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Divina.jpg")
    tsCanje(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Pocion.jpg")
    tsCanje(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Caballo.jpg")
    tsCanje(9).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Suprema.jpg")
    tsCanje(10).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Tiara.jpg")
    tsCanje(11).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_CaballoB.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tsCanje(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Champ.jpg")
    tsCanje(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Luz.jpg")
    tsCanje(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_CopadeOro.jpg")
    tsCanje(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Heroes.jpg")
    tsCanje(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Oscu.jpg")
    tsCanje(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Perseus.jpg")
    tsCanje(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Divina.jpg")
    tsCanje(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Pocion.jpg")
    tsCanje(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Caballo.jpg")
    tsCanje(9).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Suprema.jpg")
    tsCanje(10).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Tiara.jpg")
    tsCanje(11).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_CaballoB.jpg")
End Sub
Private Sub tsCanje_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then tsCanje(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Champ_I.jpg")
    If Index = 1 Then tsCanje(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Luz_I.jpg")
    If Index = 2 Then tsCanje(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_CopadeOro_I.jpg")
    If Index = 3 Then tsCanje(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Heroes_I.jpg")
    If Index = 4 Then tsCanje(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Oscu_I.jpg")
    If Index = 5 Then tsCanje(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Perseus_I.jpg")
    If Index = 6 Then tsCanje(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Divina_I.jpg")
    If Index = 7 Then tsCanje(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Pocion_I.jpg")
    If Index = 8 Then tsCanje(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Caballo_I.jpg")
    If Index = 9 Then tsCanje(9).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Suprema_I.jpg")
    If Index = 10 Then tsCanje(10).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Tiara_I.jpg")
    If Index = 11 Then tsCanje(11).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_CaballoB_I.jpg")
End Sub
Private Sub tsCanje_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then tsCanje(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Champ_A.jpg")
    If Index = 1 Then tsCanje(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Luz_A.jpg")
    If Index = 2 Then tsCanje(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_CopadeOro_A.jpg")
    If Index = 3 Then tsCanje(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Heroes_A.jpg")
    If Index = 4 Then tsCanje(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Oscu_A.jpg")
    If Index = 5 Then tsCanje(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Perseus_A.jpg")
    If Index = 6 Then tsCanje(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Divina_A.jpg")
    If Index = 7 Then tsCanje(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Pocion_A.jpg")
    If Index = 8 Then tsCanje(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Caballo_A.jpg")
    If Index = 9 Then tsCanje(9).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Suprema_A.jpg")
    If Index = 10 Then tsCanje(10).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_Tiara_A.jpg")
    If Index = 11 Then tsCanje(11).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TiendaTS_CaballoB_A.jpg")
End Sub
Private Sub tsCanje_Click(Index As Integer)
    If MsgBox("¿Estás seguro que quieres canjear este objeto?", vbYesNo, "Tierras Sagradas AO") = vbYes Then
        Call SendData("FTSPTS" & Index)
        Unload Me
    End If
End Sub

VERSION 5.00
Begin VB.Form frmGems 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Elegi tu premio"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   6360
      Top             =   120
      Width           =   375
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   5
      Left            =   4320
      Top             =   3240
      Width           =   2385
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   4
      Left            =   120
      Top             =   2160
      Width           =   2385
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   3
      Left            =   4320
      Top             =   2160
      Width           =   2385
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   2
      Left            =   120
      Top             =   1080
      Width           =   2385
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   1
      Left            =   4320
      Top             =   1080
      Width           =   2385
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   0
      Left            =   120
      Top             =   3240
      Width           =   2385
   End
End
Attribute VB_Name = "frmGems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Main.jpg")
PrizeCmd(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Torneo_N.jpg")
PrizeCmd(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Octarina_N.jpg")
PrizeCmd(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Perdon_N.jpg")
PrizeCmd(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Almas_N.jpg")
PrizeCmd(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Renuncia_N.jpg")
PrizeCmd(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Hermandad_N.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PrizeCmd(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Torneo_N.jpg")
PrizeCmd(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Octarina_N.jpg")
PrizeCmd(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Perdon_N.jpg")
PrizeCmd(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Almas_N.jpg")
PrizeCmd(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Renuncia_N.jpg")
PrizeCmd(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Hermandad_N.jpg")
End Sub
Private Sub imgSalir_Click()
Unload Me
End Sub
Private Sub PrizeCmd_Click(Index As Integer)
Call SendData("GEMS" & Index + 1)
Unload Me
End Sub
Private Sub PrizeCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then PrizeCmd(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Torneo_I.jpg")
If Index = 1 Then PrizeCmd(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Octarina_I.jpg")
If Index = 2 Then PrizeCmd(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Perdon_I.jpg")
If Index = 3 Then PrizeCmd(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Almas_I.jpg")
If Index = 4 Then PrizeCmd(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Renuncia_I.jpg")
If Index = 5 Then PrizeCmd(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Hermandad_I.jpg")
End Sub
Private Sub PrizeCmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then PrizeCmd(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Torneo_A.jpg")
If Index = 1 Then PrizeCmd(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Octarina_A.jpg")
If Index = 2 Then PrizeCmd(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Perdon_A.jpg")
If Index = 3 Then PrizeCmd(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Almas_A.jpg")
If Index = 4 Then PrizeCmd(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Renuncia_A.jpg")
If Index = 5 Then PrizeCmd(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Gemas_Hermandad_A.jpg")
End Sub


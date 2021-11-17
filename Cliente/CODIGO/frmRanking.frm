VERSION 5.00
Begin VB.Form frmRanking 
   BorderStyle     =   0  'None
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   Icon            =   "frmRanking.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   247
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   19
      Top             =   3150
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   18
      Top             =   2910
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2670
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   16
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   15
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   14
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   13
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   12
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   11
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   10
      Top             =   1005
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   10
      Left            =   960
      TabIndex        =   9
      Top             =   3135
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   9
      Left            =   960
      TabIndex        =   8
      Top             =   2895
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   8
      Left            =   960
      TabIndex        =   7
      Top             =   2655
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   6
      Top             =   2415
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   5
      Top             =   2175
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   4
      Top             =   1935
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   3
      Top             =   1695
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   2
      Top             =   1455
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   1
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   975
      Width           =   975
   End
   Begin VB.Image Botones 
      Height          =   255
      Index           =   8
      Left            =   2430
      Top             =   4500
      Width           =   810
   End
   Begin VB.Image Botones 
      Height          =   255
      Index           =   7
      Left            =   1455
      Top             =   4500
      Width           =   810
   End
   Begin VB.Image Botones 
      Height          =   255
      Index           =   6
      Left            =   480
      Top             =   4500
      Width           =   810
   End
   Begin VB.Image Botones 
      Height          =   255
      Index           =   5
      Left            =   2430
      Top             =   4080
      Width           =   810
   End
   Begin VB.Image Botones 
      Height          =   255
      Index           =   4
      Left            =   1455
      Top             =   4080
      Width           =   810
   End
   Begin VB.Image Botones 
      Height          =   255
      Index           =   3
      Left            =   480
      Top             =   4080
      Width           =   810
   End
   Begin VB.Image Botones 
      Height          =   255
      Index           =   2
      Left            =   2430
      Top             =   3660
      Width           =   810
   End
   Begin VB.Image Botones 
      Height          =   255
      Index           =   1
      Left            =   1455
      Top             =   3660
      Width           =   810
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3360
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Botones 
      Height          =   255
      Index           =   0
      Left            =   480
      Top             =   3660
      Width           =   810
   End
End
Attribute VB_Name = "frmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub MostrarRanking(rData As String)
    
Dim NickTemporal As String
Dim j As Long
'TOP10.ListItems.Clear
NickTemporal = ""

For j = 1 To 10
    NickTemporal = ReadField(j, rData, Asc(","))
    
    Nombre(j).Caption = ReadField(1, NickTemporal, Asc("-"))
    Puntaje(j).Caption = ReadField(2, NickTemporal, Asc("-"))
Next j
    
End Sub
Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

    Dim j As Long
    For j = 1 To 10
        Nombre(j).Caption = "N/A"
        Puntaje(j).Caption = "0"
    Next j

    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_Main.jpg")
    Botones(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_DuelosN.jpg")
    Botones(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ParejasN.jpg")
    Botones(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RondasN.jpg")
    Botones(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ReputacionN.jpg")
    Botones(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_TorneoN.jpg")
    Botones(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CvcN.jpg")
    Botones(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CastillosN.jpg")
    Botones(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RepuClanesN.jpg")
    Botones(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_FragsN.jpg")
End Sub
Private Sub Image2_Click()
    Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botones(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_DuelosN.jpg")
    Botones(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ParejasN.jpg")
    Botones(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RondasN.jpg")
    Botones(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ReputacionN.jpg")
    Botones(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_TorneoN.jpg")
    Botones(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CvcN.jpg")
    Botones(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CastillosN.jpg")
    Botones(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RepuClanesN.jpg")
    Botones(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_FragsN.jpg")
End Sub
Private Sub Botones_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then Botones(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_DuelosI.jpg")
    If Index = 1 Then Botones(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ParejasI.jpg")
    If Index = 2 Then Botones(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RondasI.jpg")
    If Index = 3 Then Botones(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ReputacionI.jpg")
    If Index = 4 Then Botones(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_TorneoI.jpg")
    If Index = 5 Then Botones(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CvcI.jpg")
    If Index = 6 Then Botones(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CastillosI.jpg")
    If Index = 7 Then Botones(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RepuClanesI.jpg")
    If Index = 8 Then Botones(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_FragsI.jpg")
End Sub
Private Sub Botones_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then Botones(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_DuelosA.jpg")
    If Index = 1 Then Botones(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ParejasA.jpg")
    If Index = 2 Then Botones(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RondasA.jpg")
    If Index = 3 Then Botones(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ReputacionA.jpg")
    If Index = 4 Then Botones(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_TorneoA.jpg")
    If Index = 5 Then Botones(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CvcA.jpg")
    If Index = 6 Then Botones(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CastillosA.jpg")
    If Index = 7 Then Botones(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RepuClanesA.jpg")
    If Index = 8 Then Botones(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_FragsA.jpg")
End Sub
Private Sub Botones_Click(Index As Integer)
    If Index = 0 Then Call SendData("RANKINDuelos")
    If Index = 1 Then Call SendData("RANKINParejas")
    If Index = 2 Then Call SendData("RANKINRondas")
    If Index = 3 Then Call SendData("RANKINReputacion")
    If Index = 4 Then Call SendData("RANKINTorneos")
    If Index = 5 Then Call SendData("RANKINCVCS")
    If Index = 6 Then Call SendData("RANKINCastillos")
    If Index = 7 Then Call SendData("RANKINRepuClanes")
    If Index = 8 Then Call SendData("RANKINFrags")
End Sub
Private Sub Label1_Click()
    Unload Me
End Sub

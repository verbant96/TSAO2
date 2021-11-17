VERSION 5.00
Begin VB.Form frmCastleSiege 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   Icon            =   "frmCastleSiege.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmCastleSiege.frx":10CA
   ScaleHeight     =   5250
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstGuilds 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCastleSiege.frx":18EF9
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   600
      TabIndex        =   5
      Top             =   2140
      Width           =   5655
   End
   Begin VB.Image imgVerDetalles 
      Height          =   450
      Left            =   2500
      Top             =   3880
      Width           =   1935
   End
   Begin VB.Image imgInscripcion 
      Height          =   615
      Left            =   2520
      Top             =   1285
      Width           =   1815
   End
   Begin VB.Image imgInscribirse 
      Height          =   465
      Left            =   2500
      Top             =   3880
      Width           =   1950
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6360
      Top             =   0
      Width           =   495
   End
   Begin VB.Image imgGeneral 
      Height          =   630
      Left            =   350
      Top             =   1285
      Width           =   1770
   End
   Begin VB.Image imgClanesInscriptos 
      Height          =   615
      Left            =   4800
      Top             =   1285
      Width           =   1815
   End
   Begin VB.Label txtInscriptos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "204"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   3260
      Width           =   3375
   End
   Begin VB.Label txtHorario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "04/09/2015  -  12:00 a.m."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   2915
      Width           =   3375
   End
   Begin VB.Label txtLider 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fermín"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label txtClan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<Pachamama>"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2100
      TabIndex        =   0
      Top             =   2205
      Width           =   3375
   End
End
Attribute VB_Name = "frmCastleSiege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub CSList(ByVal Rdata As String)

Dim j As Integer, k As Integer
For j = 0 To lstGuilds.ListCount - 1
    Me.lstGuilds.RemoveItem 0
Next j
k = CInt(ReadField(1, Rdata, 44))

For j = 1 To k
    lstGuilds.AddItem ReadField(1 + j, Rdata, 44)
Next j

End Sub
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

txtClan.Visible = True
txtLider.Visible = True
txtHorario.Visible = True
txtInscriptos.Visible = True
imgVerDetalles.Visible = False
imgInscribirse.Visible = False
lstGuilds.Visible = False
Label1.Visible = False

If Inscripto = 1 Then
  imgInscribirse.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_CancelarInscripcion_Iluminado.jpg")
End If

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_Main_1.jpg")

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Inscripto = 0 Then
    imgInscribirse.Picture = Nothing
End If

imgVerDetalles.Picture = Nothing

imgGeneral.Picture = Nothing
imgInscripcion.Picture = Nothing
imgClanesInscriptos.Picture = Nothing

End Sub
Private Sub imgClanesInscriptos_Click()
txtClan.Visible = False
txtLider.Visible = False
txtHorario.Visible = False
txtInscriptos.Visible = False
lstGuilds.Visible = True
imgInscribirse.Visible = False
Label1.Visible = False


imgVerDetalles.Visible = True
Call SendData("CSI")

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_Main_3.jpg")
End Sub
Private Sub imgGeneral_Click()
txtClan.Visible = True
txtLider.Visible = True
txtHorario.Visible = True
txtInscriptos.Visible = True
lstGuilds.Visible = False
imgInscribirse.Visible = False
imgVerDetalles.Visible = False
Label1.Visible = False

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_Main_1.jpg")

End Sub
Private Sub Image3_Click()
Unload Me
End Sub
Private Sub imgInscribirse_Click()

If Inscripto = 0 Then
 Call SendData("REG")
Else
 Call SendData("ACS")
End If

End Sub
Private Sub imgInscribirse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If Inscripto = 1 Then
  imgInscribirse.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_CancelarInscripcion_Iluminado.jpg")
Else
  imgInscribirse.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_Inscribirse_Iluminado.jpg")
End If

End Sub
Private Sub imgInscribirse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Inscripto = 1 Then
  imgInscribirse.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_CancelarInscripcion_Iluminado.jpg")
Else
  imgInscribirse.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_Inscribirse_Presionado.jpg")
End If

End Sub
Private Sub imgInscripcion_Click()
Label1.Visible = True
txtClan.Visible = False
txtLider.Visible = False
txtHorario.Visible = False
txtInscriptos.Visible = False
lstGuilds.Visible = False
imgVerDetalles.Visible = False

imgInscribirse.Visible = True
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_Main_2.jpg")
End Sub
Private Sub imgVerDetalles_Click()
Call SendData("CLANDETAILS" & lstGuilds.List(lstGuilds.ListIndex))
End Sub
Private Sub imgVerDetalles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgVerDetalles.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_VerDetalles_Iluminado.jpg")
End Sub
Private Sub imgVerDetalles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgVerDetalles.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_VerDetalles_Presionado.jpg")
End Sub
Private Sub imgGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgGeneral.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_General_Iluminado.jpg")
End Sub
Private Sub imgGeneral_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgGeneral.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_General_Presionado.jpg")
End Sub
Private Sub imgInscripcion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgInscripcion.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_Inscripcion_Iluminado.jpg")
End Sub
Private Sub imgInscripcion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgInscripcion.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_Inscripcion_Presionado.jpg")
End Sub
Private Sub imgClanesInscriptos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClanesInscriptos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_ClanesInscriptos_Iluminado.jpg")
End Sub
Private Sub imgClanesInscriptos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClanesInscriptos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CastleSiege_ClanesInscriptos_Presionado.jpg")
End Sub

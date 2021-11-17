VERSION 5.00
Begin VB.Form frmTorneoUsuarios 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   Picture         =   "frmTorneoUsuarios.frx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListaUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2340
      IntegralHeight  =   0   'False
      Left            =   5400
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No se esta organizando ningun torneo actualmente. Podés organizar uno por 400.000 monedas de oro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   350
      Top             =   1300
      Width           =   4815
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2830
      Top             =   2230
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   220
      Top             =   2230
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   5400
      Top             =   110
      Width           =   1815
   End
End
Attribute VB_Name = "frmTorneoUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub PonerListaTorneo(ByVal Rdata As String)

Dim j As Integer, k As Integer
For j = 0 To ListaUsers.ListCount - 1
    Me.ListaUsers.RemoveItem 0
Next j
k = CInt(ReadField(1, Rdata, 44))

For j = 1 To k
    ListaUsers.AddItem ReadField(1 + j, Rdata, 44)
Next j

Me.Show , frmMain


End Sub
Private Sub Image2_Click()
Call SendData("/CTINSC")
End Sub
Private Sub Image1_Click()
Call SendData("TUINFO")
Call SendData("TUINFD")
End Sub
Private Sub Image3_Click()
Unload Me
End Sub
Private Sub Image4_Click()

If MsgBox("¿Desea crear un nuevo torneo?", vbYesNo, "Confirmacion") = vbYes Then
Call SendData("/CTUSER")
Exit Sub
End If

End Sub
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_ActualizarN.jpg")
Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_InscribirN.jpg")
Image3.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_SalirN.jpg")
Image4.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_OrganizarN.jpg")

End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_SalirA.jpg")
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_SalirI.jpg")
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_InscribirA.jpg")
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_InscribirI.jpg")
End Sub
Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_OrganizarA.jpg")
End Sub
Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_OrganizarI.jpg")
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_ActualizarA.jpg")
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_ActualizarI.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_ActualizarN.jpg")
Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_InscribirN.jpg")
Image3.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_SalirN.jpg")
Image4.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\TorneoUser_OrganizarN.jpg")

End Sub

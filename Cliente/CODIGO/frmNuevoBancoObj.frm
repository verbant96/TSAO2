VERSION 5.00
Begin VB.Form frmNuevoBancoObj 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7905
   ClientLeft      =   -30
   ClientTop       =   -375
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmNuevoBancoObj.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox MiOro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   2925
      TabIndex        =   5
      Text            =   "999.999.999"
      Top             =   6690
      Width           =   1785
   End
   Begin VB.TextBox OroBove 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   2925
      TabIndex        =   4
      Text            =   "999.999.999"
      Top             =   6330
      Width           =   1785
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   555
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   3
      Top             =   600
      Width           =   510
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   2940
      TabIndex        =   2
      Text            =   "1"
      Top             =   5745
      Width           =   915
   End
   Begin VB.ListBox List1 
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
      ForeColor       =   &H00FFFFC0&
      Height          =   3735
      Index           =   1
      Left            =   3495
      TabIndex        =   1
      Top             =   1600
      Width           =   2730
   End
   Begin VB.ListBox List1 
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
      ForeColor       =   &H00FFFFC0&
      Height          =   3735
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   1600
      Width           =   2715
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   3390
      Top             =   1005
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   3390
      Top             =   720
      Width           =   210
   End
   Begin VB.Image Salir 
      Height          =   360
      Left            =   2640
      Top             =   7215
      Width           =   1470
   End
   Begin VB.Image RetirarOro 
      Height          =   225
      Left            =   4815
      Top             =   6330
      Width           =   1305
   End
   Begin VB.Image DepositarOro 
      Height          =   225
      Left            =   4815
      Top             =   6690
      Width           =   1305
   End
   Begin VB.Image Retirar 
      Height          =   375
      Left            =   810
      Top             =   5610
      Width           =   1695
   End
   Begin VB.Image Depositar 
      Height          =   375
      Left            =   4230
      Top             =   5610
      Width           =   1695
   End
   Begin VB.Menu cmdMenu 
      Caption         =   "Permitidos"
      Visible         =   0   'False
      Begin VB.Menu addPermitido 
         Caption         =   "Agregar permisos a un usuario para abrir esta boveda."
      End
      Begin VB.Menu delPermitido 
         Caption         =   "Quitar permisos a este usuario para abrir esta boveda."
      End
   End
End
Attribute VB_Name = "frmNuevoBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strInput As String
Dim xcantidadx As Long

Public LastIndexx1 As Integer
Public LastIndexx2 As Integer
Private Sub Depositar_Click()

    If List1(1).List(List1(1).ListIndex) = "Nada" Or _
        List1(1).ListIndex < 0 Then Exit Sub
        LastIndexx2 = List1(1).ListIndex
        If Not Inventario.Equipped(List1(1).ListIndex + 1) Then
            SendData ("DEPB" & "," & List1(1).ListIndex + 1 & "," & cantidad.text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes depositar el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If

List1(0).Clear
List1(1).Clear

Call SendData("CCBG")
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_Main.jpg")
Retirar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_Retirar_N.jpg")
Depositar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_DepositarOBJ_N.jpg")
DepositarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_DepositarOro_N.jpg")
RetirarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_RetirarOro_N.jpg")
Salir.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_Salir_N.jpg")

List1(0).BackColor = RGB(19, 21, 22)
List1(1).BackColor = RGB(19, 21, 22)
cantidad.BackColor = RGB(19, 21, 22)
OroBove.BackColor = RGB(19, 21, 22)
MiOro.BackColor = RGB(19, 21, 22)

List1(0).ForeColor = RGB(145, 123, 85)
List1(1).ForeColor = RGB(145, 123, 85)
cantidad.ForeColor = RGB(145, 123, 85)
OroBove.ForeColor = RGB(145, 123, 85)
MiOro.ForeColor = RGB(145, 123, 85)

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Retirar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_Retirar_N.jpg")
Depositar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_DepositarOBJ_N.jpg")
DepositarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_DepositarOro_N.jpg")
RetirarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_RetirarOro_N.jpg")
Salir.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_Salir_N.jpg")
End Sub
Private Sub List1_Click(Index As Integer)

Dim SR As RECT, DR As RECT, GrhIndex As Long

SR.left = 0
SR.top = 0
SR.Right = 34
SR.bottom = 34

DR.left = 0
DR.top = 0
DR.Right = 34
DR.bottom = 34


Select Case Index
    Case 0
        GrhIndex = UserBancoInventoryB(List1(0).ListIndex + 1).GrhIndex
    Case 1
        GrhIndex = Inventario.GrhIndex(List1(1).ListIndex + 1)
End Select

Call engine.DrawGrhtoHdc(GrhIndex, SR, Picture1)

End Sub

Private Sub Permitidos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu cmdMenu
End If
End Sub
Private Sub Retirar_Click()

    If List1(0).List(List1(0).ListIndex) = "Nada" Or _
        List1(0).ListIndex < 0 Then Exit Sub
        frmNuevoBancoObj.List1(0).SetFocus
        LastIndexx1 = List1(0).ListIndex
        
        SendData ("RETB" & "," & List1(0).ListIndex + 1 & "," & cantidad.text)
        
List1(0).Clear
List1(1).Clear
        
Call SendData("CCBG")
End Sub
Private Sub DepositarOro_Click()
Do
strInput = InputBox("Ingresa la cantidad a depositar", "Depositar", "0")
If StrPtr(xcantidadx) = 0 Then Exit Sub
If Not IsNumeric(strInput) Then Exit Sub

Loop While Not IsNumeric(strInput)

xcantidadx = strInput

Call SendData("CCDO" & xcantidadx)
End Sub
Private Sub RetirarOro_Click()
Do
strInput = InputBox("Ingresa la cantidad a retirar", "Retirar", "0")
If StrPtr(xcantidadx) = 0 Then Exit Sub
If Not IsNumeric(strInput) Then Exit Sub

Loop While Not IsNumeric(strInput)

xcantidadx = strInput

Call SendData("CCRO" & xcantidadx)
End Sub
Private Sub Salir_Click()
    Call SendData("FINCBN")
    Unload Me
End Sub
Private Sub Retirar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Retirar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_Retirar_I.jpg")
End Sub
Private Sub Retirar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Retirar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_Retirar_A.jpg")
End Sub
Private Sub Depositar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Depositar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_DepositarOBJ_I.jpg")
End Sub
Private Sub Depositar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Depositar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_DepositarOBJ_A.jpg")
End Sub
Private Sub RetirarOro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RetirarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_RetirarOro_I.jpg")
End Sub
Private Sub RetirarOro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RetirarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_RetirarOro_A.jpg")
End Sub
Private Sub DepositarOro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DepositarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_DepositarOro_I.jpg")
End Sub
Private Sub DepositarOro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DepositarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_DepositarOro_A.jpg")
End Sub
Private Sub Salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Salir.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_Salir_I.jpg")
End Sub
Private Sub Salir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Salir.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\BovedaClan_Salir_A.jpg")
End Sub

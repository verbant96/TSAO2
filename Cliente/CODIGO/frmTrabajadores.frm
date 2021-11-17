VERSION 5.00
Begin VB.Form frmNobleza 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Noble"
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   Icon            =   "frmTrabajadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3915
      Picture         =   "frmTrabajadores.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   4100
      Width           =   480
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1185
      Picture         =   "frmTrabajadores.frx":0850
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3930
      Picture         =   "frmTrabajadores.frx":1094
      ScaleHeight     =   480
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   615
      Width           =   465
   End
   Begin VB.ListBox lstReq 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H00C0FFFF&
      Height          =   2025
      Index           =   3
      IntegralHeight  =   0   'False
      ItemData        =   "frmTrabajadores.frx":18D8
      Left            =   3015
      List            =   "frmTrabajadores.frx":1918
      TabIndex        =   4
      Top             =   4725
      Width           =   2340
   End
   Begin VB.ListBox lstReq 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H00C0FFFF&
      Height          =   2025
      Index           =   2
      IntegralHeight  =   0   'False
      ItemData        =   "frmTrabajadores.frx":1963
      Left            =   255
      List            =   "frmTrabajadores.frx":19A3
      TabIndex        =   3
      Top             =   4725
      Width           =   2340
   End
   Begin VB.ListBox lstReq 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H00C0FFFF&
      Height          =   1980
      Index           =   1
      IntegralHeight  =   0   'False
      ItemData        =   "frmTrabajadores.frx":19EE
      Left            =   3015
      List            =   "frmTrabajadores.frx":1A2E
      TabIndex        =   2
      Top             =   1230
      Width           =   2340
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1180
      Picture         =   "frmTrabajadores.frx":1A79
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   630
      Width           =   480
   End
   Begin VB.ListBox lstReq 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H00C0FFFF&
      Height          =   1980
      Index           =   0
      IntegralHeight  =   0   'False
      ItemData        =   "frmTrabajadores.frx":22BD
      Left            =   255
      List            =   "frmTrabajadores.frx":22FD
      TabIndex        =   0
      Top             =   1230
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   5280
      Top             =   0
      Width           =   375
   End
   Begin VB.Image cmdConstruir 
      Height          =   405
      Index           =   3
      Left            =   3030
      Top             =   6750
      Width           =   2295
   End
   Begin VB.Image cmdConstruir 
      Height          =   405
      Index           =   2
      Left            =   270
      Top             =   6750
      Width           =   2295
   End
   Begin VB.Image cmdConstruir 
      Height          =   405
      Index           =   1
      Left            =   3030
      Top             =   3255
      Width           =   2295
   End
   Begin VB.Image cmdConstruir 
      Height          =   405
      Index           =   0
      Left            =   270
      Top             =   3255
      Width           =   2295
   End
End
Attribute VB_Name = "frmNobleza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const InterfaceName As String = "Nobleza"
Private Sub cmdConstruir_Click(Index As Integer)

'SendData ClientPacketID.NobleConstruirItem & SeparatorASCII & (Index + 1)

If Index = 0 Then Call SendData("/ITEMNOBLE DIADEMA")
If Index = 1 Then Call SendData("/ITEMNOBLE ARMADURA")
If Index = 2 Then Call SendData("/ITEMNOBLE ESPADA")
If Index = 3 Then Call SendData("/ITEMNOBLE ANILLO")

Unload Me

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

'If Configuracion.Alpha_Interfaz_Transparencia > 0 Then MakeTransparent Me.hWnd, Configuracion.Alpha_Interfaz_Transparencia

Me.Picture = General_Load_Interface_Picture(InterfaceName & "_Main.jpg")

ChangeButtonsNormal

lstReq(0).Clear
lstReq(1).Clear
lstReq(2).Clear
lstReq(3).Clear

End Sub

Private Sub cmdConstruir_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
cmdConstruir(Index).Picture = ChangeButtonState(Apretado, "BConstruir")

End Sub

Private Sub cmdConstruir_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
If cmdConstruir(0).Tag = "0" Then
    Call ChangeButtonsNormal
    cmdConstruir(0).Picture = ChangeButtonState(Iluminado, "BConstruir")
    cmdConstruir(0).Tag = "1"
End If
End If

If Index = 1 Then
If cmdConstruir(1).Tag = "0" Then
    Call ChangeButtonsNormal
    cmdConstruir(1).Picture = ChangeButtonState(Iluminado, "BConstruir")
    cmdConstruir(1).Tag = "1"
End If
End If

If Index = 2 Then
If cmdConstruir(2).Tag = "0" Then
    Call ChangeButtonsNormal
    cmdConstruir(2).Picture = ChangeButtonState(Iluminado, "BConstruir")
    cmdConstruir(2).Tag = "1"
End If
End If

If Index = 3 Then
If cmdConstruir(3).Tag = "0" Then
    Call ChangeButtonsNormal
    cmdConstruir(3).Picture = ChangeButtonState(Iluminado, "BConstruir")
    cmdConstruir(3).Tag = "1"
End If
End If

End Sub

Private Sub cmdConstruir_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

ChangeButtonsNormal

End Sub

Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function

Private Sub ChangeButtonsNormal()

cmdConstruir(0).Picture = ChangeButtonState(BNormal, "BConstruir")
cmdConstruir(1).Picture = ChangeButtonState(BNormal, "BConstruir")
cmdConstruir(2).Picture = ChangeButtonState(BNormal, "BConstruir")
cmdConstruir(3).Picture = ChangeButtonState(BNormal, "BConstruir")

Dim j
For Each j In Me
    j.Tag = "0"
Next

Me.Tag = "0"

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.Tag = "0" Then
    Call ChangeButtonsNormal
    Me.Tag = "1"
End If

End Sub

Private Sub Image1_Click()

Unload Me
 
End Sub

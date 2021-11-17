VERSION 5.00
Begin VB.Form frmBonificadores 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   Icon            =   "frmBonificadores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5040
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblBeneficio 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Index           =   1
      Left            =   1050
      TabIndex        =   1
      Top             =   1500
      Width           =   4155
   End
   Begin VB.Label lblBeneficio 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   675
      Width           =   4155
   End
   Begin VB.Image Bonificacion 
      Height          =   660
      Index           =   1
      Left            =   195
      Top             =   1440
      Width           =   690
   End
   Begin VB.Image Bonificacion 
      Height          =   660
      Index           =   0
      Left            =   195
      Top             =   615
      Width           =   690
   End
End
Attribute VB_Name = "frmBonificadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const InterfaceName As String = "Bonificadores"
Private Sub Bonificacion_Click(Index As Integer)

If MsgBox("¿Elegir este bonificador para tu clase?", vbYesNo) = vbYes Then
 
 If Index = 0 Then
   Call SendData("BOF" & lblBeneficio(0).Caption)
 ElseIf Index = 1 Then
   Call SendData("BOF" & lblBeneficio(1).Caption)
 End If
 
    Unload Me
    Exit Sub
End If

End Sub
Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function
Private Sub ChangeButtonsNormal()

Bonificacion(0).Picture = ChangeButtonState(BNormal, "BAbajo")
Bonificacion(1).Picture = ChangeButtonState(BNormal, "BAbajo")

Dim j
For Each j In Me
    j.Tag = "0"
Next

Me.Tag = "0"

End Sub
Private Sub Bonificacion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Bonificacion(0).Picture = ChangeButtonState(Apretado, "BAbajo")
If Index = 1 Then Bonificacion(1).Picture = ChangeButtonState(Apretado, "BAbajo")
End Sub
Private Sub Bonificacion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
If Bonificacion(0).Tag = "0" Then
    Call ChangeButtonsNormal
    Bonificacion(0).Picture = ChangeButtonState(Iluminado, "BAbajo")
    Bonificacion(0).Tag = "1"
End If
End If

If Index = 1 Then
If Bonificacion(1).Tag = "0" Then
    Call ChangeButtonsNormal
    Bonificacion(1).Picture = ChangeButtonState(Iluminado, "BAbajo")
    Bonificacion(1).Tag = "1"
End If
End If

End Sub

Private Sub Bonificacion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeButtonsNormal
End Sub

Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

lblBeneficio(0).BackColor = RGB(19, 20, 22)
lblBeneficio(1).BackColor = RGB(19, 20, 22)

Me.Picture = General_Load_Interface_Picture(InterfaceName & "_Main.jpg")
ChangeButtonsNormal

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

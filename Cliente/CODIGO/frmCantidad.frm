VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   ClientHeight    =   2415
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCantidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCantidad.frx":000C
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   975
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "0"
      Top             =   855
      Width           =   2520
   End
   Begin VB.Image bSalir 
      Height          =   375
      Left            =   4080
      Top             =   0
      Width           =   375
   End
   Begin VB.Image All 
      Height          =   345
      Left            =   2295
      Picture         =   "frmCantidad.frx":A8DA
      Top             =   1515
      Width           =   1200
   End
   Begin VB.Image Command1 
      Height          =   345
      Left            =   975
      Picture         =   "frmCantidad.frx":EDC7
      Top             =   1515
      Width           =   1200
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.2
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Private Const InterfaceName As String = "TirarObj"

Private Sub bSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
frmCantidad.Visible = False
If OfMouse Then
    SendData "TR" & Inventario.SelectedItem & "," & frmCantidad.Text1.text & "," & tX & "," & tY
    frmCantidad.Text1.text = "0"
Else
    SendData "TI" & Inventario.SelectedItem & "," & frmCantidad.Text1.text
    frmCantidad.Text1.text = "0"
End If
End Sub
Private Sub All_Click()

    frmCantidad.Visible = False
    
        If OfMouse Then
            SendData "TR" & Inventario.SelectedItem & "," & Inventario.Amount(Inventario.SelectedItem) & "," & tX & "," & tY
        Else
            SendData "TI" & Inventario.SelectedItem & "," & Inventario.Amount(Inventario.SelectedItem)
        End If
    
    frmCantidad.Text1.text = "0"

End Sub
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

    Text1.BackColor = RGB(19, 21, 23)
    frmCantidad.Text1.text = 1

'If Configuracion.Alpha_Interfaz_Transparencia > 0 Then MakeTransparent Me.hWnd, Configuracion.Alpha_Interfaz_Transparencia
Me.Picture = General_Load_Interface_Picture("TirarObj_Main.jpg")
ChangeButtonsNormal

End Sub
Private Sub Text1_Change()

If Val(Text1.text) < 0 Then
    Text1.text = MAX_INVENTORY_OBJS
End If

If Val(Text1.text) > MAX_INVENTORY_OBJS And ItemElegido <> FLAGORO Then
    Text1.text = 10000
ElseIf Val(Text1.text) > 200000 And ItemElegido = FLAGORO Then
    Text1.text = 200000
End If

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (Index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Command1.Picture = ChangeButtonState(Apretado, "BAceptar")

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Command1.Tag = "0" Then
    Call ChangeButtonsNormal
    Command1.Picture = ChangeButtonState(Iluminado, "BAceptar")
    Command1.Tag = "1"
End If


End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ChangeButtonsNormal

End Sub

Private Sub All_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    All.Picture = ChangeButtonState(Apretado, "BTodo")

End Sub

Private Sub All_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If All.Tag = "0" Then
    Call ChangeButtonsNormal
    All.Picture = ChangeButtonState(Iluminado, "BTodo")
    All.Tag = "1"
End If


End Sub

Private Sub All_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ChangeButtonsNormal

End Sub
Private Sub MasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

ChangeButtonsNormal

End Sub

Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function

Private Sub ChangeButtonsNormal()

Command1.Picture = ChangeButtonState(BNormal, "BAceptar")
All.Picture = ChangeButtonState(BNormal, "BTodo")

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


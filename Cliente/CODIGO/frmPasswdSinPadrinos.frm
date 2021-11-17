VERSION 5.00
Begin VB.Form frmPasswdSinPadrinos 
   BorderStyle     =   0  'None
   Caption         =   "¡Bienvenido!"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPasswdSinPadrinos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmPasswdSinPadrinos.frx":030A
   ScaleHeight     =   3810
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   370
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   20
      TabIndex        =   2
      ToolTipText     =   "Si perdes tu personaje lo podes recuperar con esto"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Usar '*'"
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   370
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Si perdes tu personaje lo podes recuperar con esto"
      Top             =   840
      Width           =   3495
   End
   Begin VB.Image Salir 
      Height          =   585
      Left            =   2520
      Picture         =   "frmPasswdSinPadrinos.frx":158F5
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Image Entrar 
      Height          =   570
      Left            =   240
      Picture         =   "frmPasswdSinPadrinos.frx":1695D
      Top             =   3120
      Width           =   1185
   End
End
Attribute VB_Name = "frmPasswdSinPadrinos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.2
'
'Copyright (C) 2002 Márquez Pablo Ignacio
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

Option Explicit
Private Sub Check1_Click()
If Text1.PasswordChar = "*" Then
    Text1.PasswordChar = ""
    Text2.PasswordChar = ""
Else
    Text1.PasswordChar = "*"
    Text2.PasswordChar = "*"
End If
End Sub

Private Sub Entrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Entrar.Picture = General_Load_Interface_Picture("PreguntaEntrarA.jpg")
End Sub

Private Sub Entrar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Entrar.Picture = General_Load_Interface_Picture("PreguntaEntrar.jpg")
End Sub

Private Sub Salir_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Text2.PasswordChar = "*"
End Sub

Private Sub Entrar_Click()
        
'    If Not CheckMailString(UserEmail) Then
'            MsgBox "Direccion de mail invalida."
'            Exit Sub
'    End If
    If Len(frmCuentas.Pass) < 4 Then
        Mensaje.Escribir "La clave es muy corta, tiene que tener por lo menos 4 caracteres."
        'MsgBox "La clave es muy corta, tiene que tener por lo menos 4 caracteres."
        'Unload Me
        Exit Sub
    End If
    If Text1.text = "" Then
        Mensaje.Escribir "Introduce una pregunta secreta."
        'MsgBox "Introduce una pregunta secreta."
        'Unload Me
        Exit Sub
    End If
    If Text2.text = "" Then
        Mensaje.Escribir "Introduce una respuesta a la pregunta secreta."
        'Unload Me
        Exit Sub
    End If
    If Text2.text = Text1.text Then
        Mensaje.Escribir "SI LA RESPUESTA SECRETA ES LA MISMA QUE LA PREGUNTA SERA MUY FACIL PARA CUALQUIER USUARIO MAL INTENCIONADO ROBARTE LA CUENTA, POR FAVOR ESCRIBE OTRA RESPUESTA QUE SOLO VOS CONOSCAS."
        Exit Sub
    End If
    
SendData ("NACCNT" & frmCuentas.Cuenta.text & "," & frmCuentas.Pass.text & "," & frmCuentas.Mail.text & "," & Text1.text & "," & Text2.text)

Unload frmCuentas
Unload Me

    If Not frmMain.Socket1.Connected Then
        Mensaje.Escribir "Error: Se ha perdido la conexion con el server."
        Unload Me
    End If

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Interface_Picture("Pregunta.jpg")
Entrar.Picture = General_Load_Interface_Picture("PreguntaEntrar.jpg")
Salir.Picture = General_Load_Interface_Picture("PreguntaSalir.jpg")
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Salir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Salir.Picture = General_Load_Interface_Picture("PreguntaSalirA.jpg")
End Sub

Private Sub Salir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Salir.Picture = General_Load_Interface_Picture("PreguntaSalir.jpg")
End Sub

Private Sub text1_Change()
If Text1.text = "Ejemplo: Mi nombre real es..." Then Text1.text = ""
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'FormDrag Me
End Sub

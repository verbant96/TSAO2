VERSION 5.00
Begin VB.Form frmCustomKeys 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar Teclas"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   20
      Left            =   6240
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   19
      Left            =   6240
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   18
      Left            =   6240
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   17
      Left            =   6240
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   16
      Left            =   6240
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   6240
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   6240
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   6240
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar y Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   4560
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar Teclas por defecto"
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Otros"
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   3960
      TabIndex        =   3
      Top             =   1680
      Width           =   4095
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Foto"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modo Seguro"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Macro Hechizos"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Meditar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Mapa"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Acciones"
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Atacar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tirar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ocultar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Robar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Domar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Equipar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Agarrar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones Personales"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   4095
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar/Ocultar Nombres"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Corregir Posicion"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Activar/Desactivar Musica"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Movimiento"
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Derecha"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Izquierda"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abajo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Arriba"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
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

''
'frmCustomKeys - Allows the user to customize keys.
'Implements class clsCustomKeys
'
'@author Rapsodius
'@date 20070805
'@version 1.0.0
'@see clsCustomKeys

Option Explicit

Private Sub Command1_Click()
Call CustomKeys.LoadDefaults
Dim i As Long

For i = 1 To CustomKeys.Count
    Text1(i).text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
Next i
End Sub

Private Sub Command2_Click()
Dim i As Long

For i = 1 To CustomKeys.Count
If LenB(Text1(i).text) = 0 Then


Mensaje.Escribir "Hay una o mas teclas no validas, por favor verifique."

Exit Sub

Else
Call CustomKeys.SaveCustomKeys
    
End If

Next i

Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = 1 To CustomKeys.Count
        Text1(i).text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub



Private Sub Label9_Click()

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    
    Text1(Index).text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).text)
    
    For i = 1 To CustomKeys.Count
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                Text1(Index).text = "" 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call Text1_KeyDown(Index, KeyCode, Shift)
End Sub

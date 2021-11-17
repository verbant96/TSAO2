VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8985
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   4800
      MouseIcon       =   "frmOpciones.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   6480
      Width           =   2790
   End
   Begin VB.CheckBox CheckHabla 
      BackColor       =   &H00000000&
      Caption         =   "No usar teclado numerico para cambiar modo de Habla."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4920
      TabIndex        =   25
      Top             =   5520
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Configurar Macros"
      Height          =   375
      Left            =   5160
      TabIndex        =   24
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Configurar Teclas"
      Height          =   375
      Left            =   5160
      TabIndex        =   23
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CheckBox Checkmenu 
      BackColor       =   &H00000000&
      Caption         =   "Al hacer click derecho/doble click sobre un usuario desplegar menu contextual."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4920
      TabIndex        =   22
      Top             =   4920
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Controles/Macros"
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   4680
      TabIndex        =   21
      Top             =   3600
      Width           =   3975
   End
   Begin VB.CheckBox OptTrans 
      BackColor       =   &H80000012&
      Caption         =   "Transparencia"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   5280
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Efecto de dia/tarde/noche"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   4200
      Value           =   2  'Grayed
      Width           =   2415
   End
   Begin VB.CheckBox Checktechos 
      BackColor       =   &H00000000&
      Caption         =   "Desvanecimiento en techos"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   3840
      Value           =   2  'Grayed
      Width           =   2415
   End
   Begin VB.CheckBox Checksombras 
      BackColor       =   &H00000000&
      Caption         =   "Sombras"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   4920
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox Checkreflejos 
      BackColor       =   &H00000000&
      Caption         =   "Reflejos"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   4560
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Rendimiento"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Mostrar Cartel de muerte"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox Minimap 
      BackColor       =   &H00000000&
      Caption         =   "Usar Mini Mapa"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox Emoticons 
      BackColor       =   &H00000000&
      Caption         =   "Reemplazar texto por Emoticon (Ejemplo "":P"")"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5040
      TabIndex        =   11
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Extras"
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   4800
      TabIndex        =   10
      Top             =   840
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Musica Activada"
      Height          =   345
      Index           =   0
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1560
      Width           =   2790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sonidos Activados"
      Height          =   345
      Index           =   1
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1080
      Width           =   2790
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de Sonido"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Diálogos de clan"
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   3975
      Begin VB.TextBox txtCantMensajes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "5"
         Top             =   315
         Width           =   450
      End
      Begin VB.OptionButton optPantalla 
         BackColor       =   &H00000000&
         Caption         =   "En pantalla,"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1320
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optConsola 
         BackColor       =   &H00000000&
         Caption         =   "En consola"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mensajes"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3120
         TabIndex        =   6
         Top             =   345
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   345
      Left            =   1200
      MouseIcon       =   "frmOpciones.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6480
      Width           =   2790
   End
   Begin MSComctlLib.Slider Transp 
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   5880
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Max             =   250
      SelStart        =   190
      TickStyle       =   3
      Value           =   190
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones del Juego"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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

Option Explicit

Private Sub Command1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)
    
Select Case Index
    Case 0
        If Musica Then
            Musica = False
            Command1(0).Caption = "Musica Desactivada"
            Audio.StopMidi
        Else
            Musica = True
            Command1(0).Caption = "Musica Activada"
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
        End If
    Case 1
    
        If Sound Then
            Sound = False
            Command1(1).Caption = "Sonidos Desactivados"
            Call Audio.StopWave
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        Else
            Sound = True
            Command1(1).Caption = "Sonidos Activados"
        End If
End Select
End Sub

Private Sub Command2_Click()
Call WriteVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "MenuPJs", frmOpciones.Checkmenu.value)
Call WriteVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Minimap", frmOpciones.Minimap.value)
Call WriteVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Sombras", frmOpciones.Checksombras.value)
Call WriteVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Reflejos", frmOpciones.Checkreflejos.value)
Call WriteVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Emoticons", frmOpciones.Emoticons.value)
Call WriteVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Transparencias", frmOpciones.OptTrans.value)
Call WriteVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Habla", frmOpciones.CheckHabla.value)
Me.Visible = False
End Sub

Private Sub Command3_Click()
Call frmCustomKeys.Show(vbModeless, frmMain)
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Call FrmMacros.Show(vbModeless, frmMain)
End Sub

Private Sub Emoticons_Click()
SendData "/EMOTICONS"
End Sub

Private Sub Form_Load()

'By azthenwok papa
Dim Activado As Integer
    Activado = Val(GetVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "MenuPJs"))
    frmOpciones.Checkmenu.value = Activado
    
    Activado = Val(GetVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Minimap"))
    frmOpciones.Minimap.value = Activado
    
    Activado = Val(GetVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Sombras"))
    frmOpciones.Checksombras.value = Activado
    
    Activado = Val(GetVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Reflejos"))
    frmOpciones.Checkreflejos.value = Activado
    
    Activado = Val(GetVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Emoticons"))
    frmOpciones.Emoticons.value = Activado
    
    Activado = Val(GetVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Transparencias"))
    frmOpciones.OptTrans.value = Activado
    
    Activado = Val(GetVar(App.Path & "\INIT\UserOptions.ini", "Opciones", "Habla"))
    frmOpciones.CheckHabla.value = Activado
'By azthenwok papa
    
    If Musica Then
        Command1(0).Caption = "Musica Activada"
    Else
        Command1(0).Caption = "Musica Desactivada"
    End If
    
    If Sound Then
        Command1(1).Caption = "Sonidos Activados"
    Else
        Command1(1).Caption = "Sonidos Desactivados"
    End If
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Minimap_Click()
If Minimap.value = Checked Then
frmMain.Minimap.Visible = True
End If
If Minimap.value = Unchecked Then
frmMain.Minimap.Visible = False
End If
End Sub

Private Sub optConsola_Click()
    DialogosClanes.Activo = False
End Sub

Private Sub optPantalla_Click()
    DialogosClanes.Activo = True
End Sub

Private Sub OptTrans_Click()
If OptTrans.value = Checked Then
    If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
Else
    Call Aplicar_Transparencia(Me.hWnd, CByte(255))
    End If
End Sub

Private Sub txtCantMensajes_LostFocus()
    txtCantMensajes.Text = Trim$(txtCantMensajes.Text)
    If IsNumeric(txtCantMensajes.Text) Then
        DialogosClanes.CantidadDialogos = Trim$(txtCantMensajes.Text)
    Else
        txtCantMensajes.Text = 5
    End If
End Sub

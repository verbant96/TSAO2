VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creación de un Clan"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGuildFoundation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1050
      Left            =   300
      TabIndex        =   9
      Top             =   4320
      Width           =   5400
   End
   Begin VB.TextBox txtCodex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Index           =   7
      Left            =   525
      TabIndex        =   8
      Top             =   3690
      Width           =   5175
   End
   Begin VB.TextBox txtCodex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Index           =   6
      Left            =   525
      TabIndex        =   7
      Top             =   3315
      Width           =   5175
   End
   Begin VB.TextBox txtCodex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Index           =   5
      Left            =   525
      TabIndex        =   6
      Top             =   2940
      Width           =   5175
   End
   Begin VB.TextBox txtCodex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Index           =   4
      Left            =   525
      TabIndex        =   5
      Top             =   2565
      Width           =   5175
   End
   Begin VB.TextBox txtCodex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Index           =   3
      Left            =   525
      TabIndex        =   4
      Top             =   2190
      Width           =   5175
   End
   Begin VB.TextBox txtCodex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Index           =   2
      Left            =   525
      TabIndex        =   3
      Top             =   1815
      Width           =   5175
   End
   Begin VB.TextBox txtCodex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Index           =   1
      Left            =   525
      TabIndex        =   2
      Top             =   1440
      Width           =   5175
   End
   Begin VB.TextBox txtCodex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Index           =   0
      Left            =   525
      TabIndex        =   1
      Top             =   1065
      Width           =   5175
   End
   Begin VB.TextBox txtClanName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   300
      TabIndex        =   0
      Top             =   435
      Width           =   5400
   End
   Begin VB.Image bSalir 
      Height          =   375
      Left            =   5640
      Top             =   0
      Width           =   375
   End
   Begin VB.Image bFundarClan 
      Height          =   315
      Left            =   1350
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   3300
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
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

Private Sub bFundarClan_Click()

If txtClanName = "" Then Exit Sub

If Len(txtClanName.text) <= 15 Then
    If Not AsciiValidos(txtClanName) Then
        Mensaje.Escribir "Nombre invalido."
        Exit Sub
    End If
Else
        Mensaje.Escribir "Nombre demasiado extenso."
        Exit Sub
End If

Dim fdesc$
    fdesc$ = Replace(txtDescripcion, vbCrLf, "º", , , vbBinaryCompare)
    
    If Not AsciiValidos(fdesc$) Then
           Mensaje.Escribir "La descripcion contiene caracteres invalidos"
        Exit Sub
    End If
    
    Dim k As Integer
    Dim Cont As Integer
    Cont = 0
    For k = 0 To txtCodex.UBound
        If Not AsciiValidos(txtCodex(k)) Then
            Mensaje.Escribir "El codex tiene caracteres invalidos"
            Exit Sub
        End If
        If Len(txtCodex(k).text) > 0 Then Cont = Cont + 1
    Next k
    
    If Cont < 4 Then
            Mensaje.Escribir "Debes definir al menos cuatro mandamientos."
            Exit Sub
    End If
    
    Dim chunk$
    
    If CreandoClan Then
        chunk$ = "CIG" & fdesc$
        chunk$ = chunk$ & "¬" & txtClanName & "¬" & Site & "¬" & Cont
    Else
        chunk$ = "DESCOD" & fdesc$ & "¬" & Cont
    End If
    
    
    
    For k = 0 To txtCodex.UBound
        chunk$ = chunk$ & "¬" & txtCodex(k)
    Next k
    
    
    Call SendData(chunk$)
    
    CreandoClan = False
    
    Unload Me


End Sub

Private Sub bSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
    form_Moviment.Initialize Me

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\FundarClan_Main.jpg")
bFundarClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\FundarClan_Normal.jpg")

Dim i As Long

txtClanName.BackColor = RGB(19, 21, 23)
txtDescripcion.BackColor = RGB(19, 21, 23)

For i = 0 To 7
    txtCodex(i).BackColor = RGB(19, 21, 23)
    txtCodex(i).ForeColor = RGB(145, 123, 85)
Next i
    
txtClanName.ForeColor = RGB(145, 123, 85)
txtDescripcion.ForeColor = RGB(145, 123, 85)

End Sub
Private Sub bFundarClan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bFundarClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\FundarClan_Apretado.jpg")
End Sub
Private Sub bFundarClan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bFundarClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\FundarClan_Iluminado.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bFundarClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\FundarClan_Normal.jpg")
End Sub

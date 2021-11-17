VERSION 5.00
Begin VB.Form frmGuildSubLeader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración del Clan"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Propuestas de alianzas"
      Height          =   495
      Left            =   2040
      MouseIcon       =   "frmGuildSubLeader.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   4080
      MouseIcon       =   "frmGuildSubLeader.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Propuestas de paz"
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmGuildSubLeader.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2895
      Begin VB.ListBox guildslist 
         Height          =   1425
         ItemData        =   "frmGuildSubLeader.frx":03F6
         Left            =   120
         List            =   "frmGuildSubLeader.frx":03F8
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildSubLeader.frx":03FA
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "GuildNews"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildSubLeader.frx":054C
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildSubLeader.frx":069E
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1440
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Height          =   1035
         ItemData        =   "frmGuildSubLeader.frx":07F0
         Left            =   120
         List            =   "frmGuildSubLeader.frx":07F2
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Miembros 
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmGuildSubLeader"
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

Private Sub Command1_Click()

frmCharInfo.frmsolicitudes = True
Call SendData("1HRINFO<" & solicitudes.List(solicitudes.listIndex))

'Unload Me

End Sub

Private Sub Command3_Click()

Dim k$

k$ = Replace(txtguildnews, vbCrLf, "º")

Call SendData("ACTGNEWS" & k$)

End Sub

Private Sub Command4_Click()

frmGuildBrief.EsLeader = True
Call SendData("CLANDETAILS" & guildslist.List(guildslist.listIndex))

'Unload Me

End Sub
Private Sub Command7_Click()
Call SendData("ENVPROPP")
End Sub
Private Sub Command9_Click()
Call SendData("ENVALPRO")
End Sub


Private Sub Command8_Click()
Unload Me
frmMain.SetFocus
End Sub


Public Sub ParseLeaderInfo(ByVal Data As String)

If Me.Visible Then Exit Sub

Dim r%, T%

r% = Val(ReadField(1, Data, Asc("¬")))

For T% = 1 To r%
    guildslist.AddItem ReadField(1 + T%, Data, Asc("¬"))
Next T%

r% = Val(ReadField(T% + 1, Data, Asc("¬")))
Miembros.Caption = "El clan cuenta con " & r% & " miembros."

Dim k%

txtguildnews = Replace(ReadField(T% + k% + 1, Data, Asc("¬")), "º", vbCrLf)

T% = T% + k% + 2

r% = Val(ReadField(T%, Data, Asc("¬")))

For k% = 1 To r%
    solicitudes.AddItem ReadField(T% + k%, Data, Asc("¬"))
Next k%

Me.Show , frmMain

End Sub


Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

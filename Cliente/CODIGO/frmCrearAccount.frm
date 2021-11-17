VERSION 5.00
Begin VB.Form frmCrearAccount 
   BorderStyle     =   0  'None
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   3945
   ControlBox      =   0   'False
   Icon            =   "frmCrearAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearAccount.frx":000C
   ScaleHeight     =   9390
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox respuesta 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      Left            =   480
      TabIndex        =   6
      Top             =   4950
      Width           =   3100
   End
   Begin VB.TextBox pregunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      Left            =   480
      TabIndex        =   5
      Top             =   4200
      Width           =   3100
   End
   Begin VB.TextBox Mail2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      Left            =   480
      MaxLength       =   25
      TabIndex        =   4
      Top             =   3420
      Width           =   3100
   End
   Begin VB.TextBox Nombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      Left            =   480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   460
      Width           =   3100
   End
   Begin VB.TextBox Pass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   25
      PasswordChar    =   "X"
      TabIndex        =   1
      Top             =   1185
      Width           =   3100
   End
   Begin VB.TextBox RePass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   25
      PasswordChar    =   "X"
      TabIndex        =   2
      Top             =   1920
      Width           =   3100
   End
   Begin VB.TextBox Mail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   340
      Left            =   480
      MaxLength       =   25
      TabIndex        =   3
      Top             =   2640
      Width           =   3100
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   600
      Top             =   8280
      Width           =   2625
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   600
      Top             =   7320
      Width           =   2610
   End
End
Attribute VB_Name = "frmCrearAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\CrearCuenta_Main.jpg")
End Sub

Private Sub Image1_Click()
Unload Me

End Sub

Private Sub Image2_Click()

 
If Len(nombre.text) < 4 Then
    MsgBox "El nombre de la cuenta debe de tener mas de 4 caracteres."
    Exit Sub
End If
 
If Len(nombre.text) >= 20 Then
    MsgBox "El nombre de la cuenta debe de tener mas de 20 caracteres."
    Exit Sub
End If

 
If Len(Pass.text) < 6 Then
    MsgBox "El password de la cuenta debe de tener mas de 6 caracteres."
    Exit Sub
End If
 
If Len(Pass.text) >= 25 Then
    MsgBox "El password de la cuenta debe de tener menos de 25 caracteres."
    Exit Sub
End If
        

If Pass <> RePass Then
    MsgBox "Las passwords que tipeo no coinciden", , "Coco rules"
    Exit Sub
End If

If Not CheckMailString(Mail) Then
    MsgBox "Direccion de mail invalida."
    Exit Sub
End If
If Mail.text <> Mail2.text Then
MsgBox "Los emails no coinciden"
Exit Sub
End If
If nombre = "" Or Pass = "" Or RePass = "" Or Mail = "" Or pregunta = "" Or respuesta = "" Then
    MsgBox "Completa todo!"
    Exit Sub
End If
Call SendData("NACCNT" & nombre & "," & Pass & "," & Mail & "," & pregunta & "," & respuesta)

Unload Me
MsgBox "La cuenta fue creada con éxito."
End Sub


VERSION 5.00
Begin VB.Form frmNuevoBanco 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox PIN 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Contraseña 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Cuenta 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Ingresar 
      Caption         =   "Ingresar"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Clave 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox ID 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton CrearCuenta 
      Caption         =   "Crear Cuenta Bancaria"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmNuevoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CrearCuenta_Click()
    Call SendData("CCBN" & Cuenta.text & "," & Contraseña.text & "," & PIN.text)
End Sub
Private Sub Ingresar_Click()
    Call SendData("CCBL" & ID.text & "," & Clave.text)
End Sub

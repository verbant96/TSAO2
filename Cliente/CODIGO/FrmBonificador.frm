VERSION 5.00
Begin VB.Form FrmBonificador 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   Picture         =   "FrmBonificador.frx":0000
   ScaleHeight     =   2220
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label bonificador2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label bonificador1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   240
      Picture         =   "FrmBonificador.frx":183C4
      Top             =   480
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   690
      Left            =   240
      Picture         =   "FrmBonificador.frx":1E490
      Top             =   1320
      Width           =   720
   End
End
Attribute VB_Name = "FrmBonificador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
If MsgBox("¿Estas seguro de elegir este bonificador?", vbYesNo) = vbYes Then
Call SendData("/ELIJOELBONI " & bonificador1.Caption)
Unload Me
End If
End Sub

Private Sub Image2_Click()
If MsgBox("¿Estas seguro de elegir este bonificador?", vbYesNo) = vbYes Then
Call SendData("/ELIJOELBONI " & bonificador2.Caption)
Unload Me
End If

End Sub

Private Sub Label1_Click()
Unload Me
End Sub

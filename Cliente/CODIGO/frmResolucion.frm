VERSION 5.00
Begin VB.Form frmResolucion 
   BorderStyle     =   0  'None
   Caption         =   "Tierras Sagradas - Elegir Resolucion"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   Picture         =   "frmResolucion.frx":0000
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3600
      Top             =   0
      Width           =   495
   End
   Begin VB.Image imgNo 
      Height          =   495
      Left            =   2610
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image imgSi 
      Height          =   495
      Left            =   1350
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmResolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        
        If KeyCode = vbKeyReturn Then
            Call imgSi_Click
        End If
End Sub

Private Sub Image1_Click()
    Resolucion.notModoVentana
    Unload Me
End Sub

Private Sub imgNo_Click()
    Resolucion.SetResolucion
    Unload Me
End Sub

Private Sub imgSi_Click()
    Resolucion.notModoVentana
    Unload Me
End Sub

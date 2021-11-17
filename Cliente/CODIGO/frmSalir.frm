VERSION 5.00
Begin VB.Form frmSalir 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Salir del Juego"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2760
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   255
      Top             =   1530
      Width           =   2640
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   255
      Top             =   525
      Width           =   2640
   End
End
Attribute VB_Name = "frmSalir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Salir_Main.jpg")
End Sub

Private Sub Image1_Click()
    Call SendData("/SALIR")
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
        If KeyCode = vbKeyReturn Then
            Call Image1_Click
        End If
End Sub
Private Sub Image2_Click()
    UnloadAllForms
    engine.Engine_Deinit
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = Nothing
    Image2.Picture = Nothing
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Nothing
    Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Salir_InicioI.jpg")
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Salir_InicioA.jpg")
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = Nothing
    Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Salir_CerrarI.jpg")
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Salir_CerrarA.jpg")
End Sub
Private Sub Image3_Click()
    Unload Me
End Sub

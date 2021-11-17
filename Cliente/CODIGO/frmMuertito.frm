VERSION 5.00
Begin VB.Form frmMuertito 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   Icon            =   "frmMuertito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   300
      Left            =   945
      Top             =   885
      Width           =   2070
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   945
      Top             =   570
      Width           =   2070
   End
End
Attribute VB_Name = "frmMuertito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\cartelMuerte_Main.jpg")
    Set form_Moviment = New clsFormMovementManager
    form_Moviment.Initialize Me
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Image2_Click()
    Call SendData("/REGRESAR")
    Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = Nothing
    Image2.Picture = Nothing
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Nothing
    Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\cartelMuerte_ContinuarI.jpg")
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\cartelMuerte_ContinuarA.jpg")
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = Nothing
    Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\cartelMuerte_RegresarI.jpg")
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\cartelMuerte_RegresarA.jpg")
End Sub


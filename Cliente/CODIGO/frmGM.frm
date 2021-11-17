VERSION 5.00
Begin VB.Form frmGM 
   BorderStyle     =   0  'None
   Caption         =   "Ayuda GM"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   Icon            =   "frmGM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMotivo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   75
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2325
      Width           =   4185
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3960
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   60
      Top             =   4440
      Width           =   4215
   End
End
Attribute VB_Name = "frmGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    If Len(txtMotivo) > 250 Then
        Mensaje.Escribir "Máximo 250 caracteres."
        Exit Sub
    End If
    
    If Len(txtMotivo) < 20 Then
        Mensaje.Escribir "Tamaño mínimo 20 caracteres, no se permiten mensajes del tipo 'GM sum' y parecidos."
    Exit Sub
    End If
    
    Mensaje.Escribir "¡Tu consulta a sido enviada!"
    Call SendData("#" & 0 & "|" & txtMotivo)
    
    Unload Me
End Sub
Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub txtMotivo_Change()
    If Len(txtMotivo) = 250 Then
        Mensaje.Escribir "Tamaño maximo 250 caracteres."
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GM_Main.jpg")
    txtMotivo.BackColor = RGB(21, 22, 24)
    txtMotivo.ForeColor = RGB(185, 169, 146)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = Nothing
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GM_EnviarI.jpg")
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GM_EnviarA.jpg")
End Sub

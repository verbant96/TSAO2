VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMedallas 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   LinkTopic       =   "Form2"
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   3480
   End
   Begin RichTextLib.RichTextBox txtDesc 
      Height          =   450
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   794
      _Version        =   393217
      BackColor       =   -2147483647
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMedallas.frx":0000
   End
   Begin VB.ListBox lstPacks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1650
      IntegralHeight  =   0   'False
      Left            =   800
      TabIndex        =   1
      Top             =   960
      Width           =   4235
   End
   Begin VB.PictureBox picPack 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1640
      Left            =   5040
      ScaleHeight     =   1605
      ScaleWidth      =   1650
      TabIndex        =   0
      Top             =   960
      Width           =   1675
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3330
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image cmdCanjear 
      Height          =   555
      Left            =   2340
      Top             =   4460
      Width           =   2820
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6960
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmMedallas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

    Call SendData("GEPS1")

    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Main.jpg")
    cmdCanjear.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_canjearNormal.jpg")

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCanjear.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_canjearNormal.jpg")
End Sub
Private Sub Image1_Click()
    Unload Me
    Timer1.Enabled = False
End Sub
Private Sub cmdCanjear_Click()
    If MsgBox("¿Estás seguro que deseas canjear " & lstPacks.text & "?", vbYesNo) = vbYes Then
        Call SendData("GEDS" & lstPacks.ListIndex + 1)
        Timer1.Enabled = False
        Unload Me
    End If
End Sub
Private Sub cmdCanjear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCanjear.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_canjearI.jpg")
End Sub
Private Sub cmdCanjear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCanjear.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_canjearPress.jpg")
End Sub
Private Sub lstPacks_Click()
    Call SendData("GEPS" & lstPacks.ListIndex + 1)
End Sub

Private Sub Timer1_Timer()
    Call engine.drawPremiosMedalla
End Sub

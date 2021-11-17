VERSION 5.00
Begin VB.Form frmSubastar 
   BorderStyle     =   0  'None
   Caption         =   "Subasta"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   219
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox StartBid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   345
      Left            =   1680
      TabIndex        =   2
      Text            =   "1000"
      Top             =   5535
      Width           =   1290
   End
   Begin VB.TextBox Amount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Text            =   "1"
      Top             =   5010
      Width           =   1290
   End
   Begin VB.ListBox ItemList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3900
      IntegralHeight  =   0   'False
      ItemData        =   "frmSubastar.frx":0000
      Left            =   300
      List            =   "frmSubastar.frx":0007
      TabIndex        =   0
      Top             =   945
      Width           =   2700
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   0
      Width           =   405
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   320
      Top             =   6030
      Width           =   2655
   End
End
Attribute VB_Name = "frmSubastar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
If Not IsNumeric(Amount.text) Then Exit Sub
If Not IsNumeric(StartBid.text) Then Exit Sub
If ItemList.text = "Nada" Then Exit Sub
 
Call SendData("/INISUB " & ItemList.ListIndex + 1 & " " & Amount.text & " " & StartBid.text & "")
Unload Me
 
End Sub
Private Sub Image2_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

ItemList.BackColor = RGB(19, 21, 23)
Amount.BackColor = RGB(19, 21, 23)
StartBid.BackColor = RGB(19, 21, 23)

Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Subasta_Iniciar_N.jpg")
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Subasta_Main.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Subasta_Iniciar_N.jpg")
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Subasta_Iniciar_A.jpg")
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Subasta_Iniciar_I.jpg")
End Sub

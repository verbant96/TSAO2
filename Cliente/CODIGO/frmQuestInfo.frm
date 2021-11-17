VERSION 5.00
Begin VB.Form frmQuestInfo 
   BorderStyle     =   0  'None
   Caption         =   "Quests Info"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuestInfo.frx":0000
   ScaleHeight     =   4065
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Desc 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFFF&
      Height          =   700
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4275
   End
   Begin VB.Label Tipo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   2470
      Top             =   3380
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   140
      Top             =   3380
      Width           =   1905
   End
   Begin VB.Label GLDPT 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label PosName 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   1290
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label NPCs 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Users 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Call SendData("ACQT" & Numeriyo)
Unload Me
Unload frmQuestSelect
End Sub
Private Sub Image2_Click()
Unload Me
End Sub
Private Sub Form_Load()
Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quest_BAbandonarN.jpg")
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quest_BAceptarN.jpg")

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quest_Main.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quest_BAbandonarN.jpg")
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quest_BAceptarN.jpg")
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quest_BAbandonarA.jpg")
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quest_BAbandonarI.jpg")
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quest_BAceptarA.jpg")
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quest_BAceptarI.jpg")
End Sub

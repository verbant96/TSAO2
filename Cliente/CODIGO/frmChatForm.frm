VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChatForm 
   BorderStyle     =   0  'None
   Caption         =   "pepemago"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   Icon            =   "frmChatForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtChatSend 
      BackColor       =   &H00000040&
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
      ForeColor       =   &H80000005&
      Height          =   555
      Left            =   195
      TabIndex        =   2
      Text            =   "Escribi tu texto aca."
      Top             =   3225
      Width           =   4140
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3240
      Top             =   0
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   2550
      Left            =   195
      TabIndex        =   1
      Top             =   495
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   4498
      _Version        =   393217
      BackColor       =   -2147483647
      BorderStyle     =   0
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmChatForm.frx":000C
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   150
      Width           =   2175
   End
   Begin VB.Image imgCmd 
      Height          =   300
      Index           =   0
      Left            =   3960
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgCmd 
      Height          =   300
      Index           =   1
      Left            =   4250
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmChatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

rtbChat.BackColor = RGB(19, 21, 22)
txtChatSend.BackColor = RGB(19, 21, 22)
txtChatSend.ForeColor = RGB(145, 123, 85)

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Chat_Main.jpg")
Call SetWindowLong(rtbChat.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
End Sub
Private Sub imgCmd_Click(Index As Integer)
Dim i As Long

    If Index = 0 Then
        For i = 1 To 5
            If UCase$(lblName.Caption) = UCase$(NickContacto(i)) Then
                VentanitaMostrar(i) = 0
                RecibioMensaje(i) = False
            End If
        Next i
    
        Me.Visible = False
    ElseIf Index = 1 Then
        For i = 1 To 5
            If UCase$(lblName.Caption) = UCase$(NickContacto(i)) Then
                ChatEnUso(i) = False
                NickContacto(i) = ""
                VentanitaMostrar(i) = 0
                RecibioMensaje(i) = False
                ChatForm(i).rtbChat.text = ""
            End If
        Next i
        
        Me.Visible = False
    End If
End Sub
Private Sub Timer1_Timer()
    rtbChat.Refresh
End Sub
Private Sub txtChatSend_Click()
    txtChatSend.text = ""
End Sub
Private Sub txtChatSend_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
  Call SendData("KKCHAT" & lblName.Caption & "," & txtChatSend.text)
  AddtoRichTextBox rtbChat, "" & frmMain.Label8.Caption & " dice: " & txtChatSend.text & "", 255, 255, 0, True
  txtChatSend.text = ""
End If

End Sub

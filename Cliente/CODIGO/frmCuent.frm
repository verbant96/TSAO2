VERSION 5.00
Begin VB.Form frmCuent 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Tierras Perdidas AO"
   ClientHeight    =   9000
   ClientLeft      =   8310
   ClientTop       =   3645
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCuent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCuent.frx":000C
   Picture         =   "frmCuent.frx":0CD6
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   0
      Width           =   12000
      Begin VB.Image img_EntrarPJ 
         Height          =   300
         Left            =   480
         Top             =   12000
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image img_BorrarPJ 
         Height          =   300
         Left            =   1560
         Top             =   12000
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image imgCrearPersonaje 
         Height          =   495
         Left            =   3900
         Top             =   8085
         Width           =   4230
      End
      Begin VB.Image imgSalir4 
         Height          =   495
         Left            =   300
         Top             =   8070
         Width           =   1350
      End
      Begin VB.Image imgBorrarCuenta 
         Height          =   495
         Left            =   9495
         Top             =   8070
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Image imgCambiarPass 
         Height          =   495
         Left            =   9495
         Top             =   7425
         Width           =   2220
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   9
         Left            =   3120
         Top             =   3720
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   8
         Left            =   7920
         Top             =   3720
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   7
         Left            =   7440
         Top             =   5640
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   6
         Left            =   3600
         Top             =   5640
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   5
         Left            =   3600
         Top             =   1800
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   4
         Left            =   7440
         Top             =   1800
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   3
         Left            =   4560
         Top             =   4680
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   2
         Left            =   6480
         Top             =   4680
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   1
         Left            =   6480
         Top             =   2760
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   0
         Left            =   4560
         Top             =   2760
         Width           =   975
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Accname"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   12840
      TabIndex        =   0
      Top             =   2400
      Width           =   2175
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim PJApretado As Byte
Dim BorrarRandom As String
Dim ElRandom As String
Private Sub img_BorrarPJ_Click()
      BorrarRandom = RandomNumber(1000, 9999)
      ElRandom = InputBox("Ingrese el codigo " & BorrarRandom & " para borrar su personaje:", "Borrar Personaje")
        
      If BorrarRandom = ElRandom Then Call SendData("TBRP" & CargarPJ(PJApretado).nombre & "," & nombrecuent & "," & CodigoRecibido)
End Sub
Private Sub img_EntrarPJ_Click()
    SendData ("OOLOGI" & CargarPJ(PJApretado).nombre & "," & nombrecuent & "," & CodigoRecibido)
    frmMain.Label8.Caption = CargarPJ(PJApretado).nombre
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
For i = 0 To 9
If MostrarTodo(i) = True Then
    MostrarTodo(i) = False
    img_EntrarPJ.Visible = False
    img_BorrarPJ.Visible = False
End If

If CrearAura(i) = True Then
 CrearAura(i) = False
End If
Next i

End Sub
Private Sub PJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

For i = 0 To 9
If CrearAura(i) = True Then
 CrearAura(i) = False
End If
Next i

If MostrarTodo(Index) = True Then
 CrearAura(Index) = False
Else
 CrearAura(Index) = True
    img_EntrarPJ.Visible = False
    img_BorrarPJ.Visible = False
End If

End Sub
Private Sub PJ_Click(Index As Integer)

For i = 0 To 9
If MostrarTodo(i) = True Then
    MostrarTodo(i) = False
End If

If CrearAura(i) = True Then
 CrearAura(i) = False
End If
Next i

If Index = 0 Then
    EntrarX = 290 '+65 borrar
    EntrarY = 180
ElseIf Index = 1 Then
    EntrarX = 418
    EntrarY = 180
ElseIf Index = 2 Then
    EntrarX = 418
    EntrarY = 313
ElseIf Index = 3 Then
    EntrarX = 290
    EntrarY = 313
ElseIf Index = 4 Then
    EntrarX = 477
    EntrarY = 117
ElseIf Index = 5 Then
    EntrarX = 224
    EntrarY = 117
ElseIf Index = 6 Then
    EntrarX = 481
    EntrarY = 379
ElseIf Index = 7 Then
    EntrarX = 224
    EntrarY = 379
ElseIf Index = 8 Then
    EntrarX = 512
    EntrarY = 242
ElseIf Index = 9 Then
    EntrarX = 193
    EntrarY = 242
End If
    
img_EntrarPJ.top = EntrarY
img_EntrarPJ.left = EntrarX
img_BorrarPJ.top = EntrarY
img_BorrarPJ.left = EntrarX + 65

MostrarTodo(Index) = True
img_EntrarPJ.Visible = True
img_BorrarPJ.Visible = True
PJApretado = Index

End Sub
Private Sub imgCrearPersonaje_Click()
    If CargarPJ(9).Existe = True Then
        Mensaje.Escribir "No puedes crear más personajes."
    Else
        Call Audio.PlayWave("click.wav")
    
        EstadoLogin = Dados
        'frmCuent.Visible = False
        frmCrearPersonaje.Show , frmCuent
        Audio.StopWave
        'frmCrearPersonaje.MousePointer = 11
    End If
End Sub
Private Sub imgCambiarPass_Click()

        Call Audio.PlayWave("click.wav")
        Dim anteriorpw As String
        Dim nuevapw As String
        Dim renuevapw As String
        
        anteriorpw = InputBox("Ingrese su actual contraseña:", "Cambiar Password")
        nuevapw = InputBox("Ingrese su nueva contraseña:", "Cambiar Password")
        renuevapw = InputBox("Repita su nueva contraseña:", "Cambiar Password")
        
        If nuevapw <> renuevapw Then
            Mensaje.Escribir "Las passwords que tipeo no coinciden"
            Exit Sub
        End If
        
        SendData ("REPASS" & nombrecuent & "," & anteriorpw & "," & nuevapw & "," & renuevapw)
End Sub
Private Sub Salir4_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmMain.Socket1.Disconnect
    frmMain.Socket1.Cleanup
    Unload frmCuent

    If FileExist(App.Path & "\Data\MAPAS\" & "Mapa" & MapConnect & ".map", vbNormal) Then
       Call SwitchMap(MapConnect)
        day_r_old = 80
        day_g_old = 80
        day_b_old = 80
        base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
   Else
       MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
       Call UnloadAllForms
   End If
    
    AoDefResult = 0
    frmConnect.Visible = True
End Sub
Private Sub imgCrearPersonaje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCP = "Iluminado"
End Sub
Private Sub imgCrearPersonaje_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCP = "Apretado"
End Sub
Private Sub imgCrearPersonaje_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCP = "Normal"
End Sub
Private Sub imgSalir4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonSalir = "Iluminado"
End Sub
Private Sub imgSalir4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonSalir = "Apretado"
End Sub
Private Sub imgSalir4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonSalir = "Normal"
End Sub
Private Sub imgCambiarPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCPass = "Iluminado"
End Sub
Private Sub imgCambiarPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCPass = "Apretado"
End Sub
Private Sub imgCambiarPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCPass = "Normal"
End Sub

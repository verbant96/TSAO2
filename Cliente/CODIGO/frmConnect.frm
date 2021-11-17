VERSION 5.00
Begin VB.Form frmConnect 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   270
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   670.522
   ScaleMode       =   0  'User
   ScaleWidth      =   813.215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1440
      Top             =   960
   End
   Begin VB.Image imgRecAccPin 
      Height          =   615
      Left            =   3881
      Top             =   4583
      Width           =   4260
   End
   Begin VB.Image imgRecuAcc 
      Height          =   675
      Left            =   4884
      Top             =   5481
      Width           =   2265
   End
   Begin VB.Image imgRECACC 
      Height          =   615
      Left            =   3881
      Top             =   3792
      Width           =   4260
   End
   Begin VB.Image imgNEWACC 
      Height          =   615
      Index           =   4
      Left            =   3881
      Top             =   6120
      Width           =   4260
   End
   Begin VB.Image imgNEWACC 
      Height          =   615
      Index           =   3
      Left            =   3881
      Top             =   5320
      Width           =   4260
   End
   Begin VB.Image imgNEWACC 
      Height          =   615
      Index           =   2
      Left            =   3881
      Top             =   4543
      Width           =   4260
   End
   Begin VB.Image imgNEWACC 
      Height          =   615
      Index           =   1
      Left            =   3881
      Top             =   3725
      Width           =   4260
   End
   Begin VB.Image imgNEWACC 
      Height          =   615
      Index           =   0
      Left            =   3881
      Top             =   2921
      Width           =   4260
   End
   Begin VB.Label txtNuevoPIN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   9
      Top             =   5494
      Width           =   4095
   End
   Begin VB.Label txtNuevoPIN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   10
      Top             =   6298
      Width           =   4095
   End
   Begin VB.Label txtNuevaPassword 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   8
      Top             =   4690
      Width           =   4095
   End
   Begin VB.Label txtNuevaPassword 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   7
      Top             =   3886
      Width           =   4095
   End
   Begin VB.Label txtNuevaCuenta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3082
      Width           =   4095
   End
   Begin VB.Image imgNewAccount 
      Height          =   675
      Left            =   4884
      Top             =   7115
      Width           =   2265
   End
   Begin VB.Image imgRecorder 
      Height          =   225
      Left            =   4840
      Top             =   5615
      Width           =   225
   End
   Begin VB.Image imgConectar 
      Height          =   675
      Left            =   4870
      Top             =   6486
      Width           =   2265
   End
   Begin VB.Image imgCrearCuenta 
      Height          =   285
      Left            =   7560
      Top             =   8656
      Width           =   1455
   End
   Begin VB.Image imgSalir 
      Height          =   285
      Left            =   6240
      Top             =   8655
      Width           =   795
   End
   Begin VB.Image imgPass 
      Height          =   615
      Left            =   3885
      Top             =   4596
      Width           =   4260
   End
   Begin VB.Image imgName 
      Height          =   615
      Left            =   3885
      Top             =   3792
      Width           =   4260
   End
   Begin VB.Label lblCreateAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CREAR CUENTA"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   8710
      Width           =   1455
   End
   Begin VB.Label lblSalir 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   8715
      Width           =   735
   End
   Begin VB.Label lblBarrita 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5962
      TabIndex        =   2
      Top             =   4730
      Width           =   135
   End
   Begin VB.Label txtPassword 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4770
      Width           =   4335
   End
   Begin VB.Label txtCuenta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3975
      TabIndex        =   0
      Top             =   3940
      Width           =   4095
   End
   Begin VB.Label txtRecAcc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3969
      TabIndex        =   11
      Top             =   3953
      Width           =   4095
   End
   Begin VB.Label txtRecAccPIN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   4757
      Width           =   4095
   End
   Begin VB.Image imgRecuperarCuenta 
      Height          =   285
      Left            =   3000
      Top             =   8656
      Width           =   2820
   End
   Begin VB.Label lblRecu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¿OLVIDASTE TU CONTRASEÑA?"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2833
      TabIndex        =   5
      Top             =   8710
      Width           =   3015
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pageSelected As Byte

Private txtNewAccount As String
Private txtNewPassword0(1 To 2) As String
Private txtNewPassword1(1 To 2) As String
Private txtPIN0(1 To 2) As String
Private txtPIN1(1 To 2) As String
Private NA_clickeoText As Byte

Private recupSelect(1 To 2) As Boolean
Private txtRecuAccount As String
Private txtRecuPIN(1 To 2) As String


Private modificoBoton As Boolean
Private ClickeoTextCuenta As Boolean
Private ClickeoTextPassw As Boolean
Private TextBoxCuenta As String
Private TextBoxPassw As String
Private TextBoxPasswR As String
Private BarritaTextConnect As Byte

Private Sub guardarCuenta()

    Dim l_file As clsIniReader
    Set l_file = New clsIniReader

    '@ load file
    l_file.Initialize App.Path & "\Data\INIT\UserConfig.ini"
    
    l_file.ChangeValue "OPTIONS", "RECORDAR_CUENTA", Configuracion.recordarCuenta
    
    Configuracion.tmpCuenta = TextBoxCuenta
    Configuracion.tmpPassword = TextBoxPassw
     
    If Configuracion.recordarCuenta Then
        l_file.ChangeValue "OPTIONS", "TMPCUENTA", Configuracion.tmpCuenta
        l_file.ChangeValue "OPTIONS", "TMPPASSWORD", Configuracion.tmpPassword
    Else
        l_file.ChangeValue "OPTIONS", "TMPCUENTA", ""
        l_file.ChangeValue "OPTIONS", "TMPPASSWORD", ""
    End If
    
    l_file.DumpFile App.Path & "\Data\INIT\UserConfig.ini"
    

End Sub
Private Sub imgNEWACC_Click(Index As Integer)
    
    NA_clickeoText = Index
    
    Select Case Index
        Case 0
            txtNuevaCuenta_Click
        Case 1, 2
            txtNuevaPassword_Click (Index - 1)
        Case 3, 4
            txtNuevoPIN_Click (Index - 3)
        Case Else
            Sleep 10
    End Select
    
End Sub
Private Sub imgNEWACC_DblClick(Index As Integer)
    
    NA_clickeoText = Index
    
    Select Case Index
        Case 0
            txtNuevaCuenta_DblClick
        Case 1, 2
            txtNuevaPassword_DblClick (Index - 1)
        Case 3, 4
            txtNuevoPIN_DblClick (Index - 3)
    End Select
    
End Sub
Private Sub imgNewAccount_Click()
    
    imgNewAccount.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_createAccount.jpg"): modificoBoton = True
    
    If (Len(txtPIN0(1)) < 1) Or (Len(txtNewPassword0(1)) < 1) Or (Len(txtNewAccount) < 1) Then
        Mensaje.Escribir "Debes completar todos los campos."
        Exit Sub
    End If
    
    If (txtNewPassword0(1) <> txtNewPassword1(1)) Then
        Mensaje.Escribir "Las contraseñas no coinciden."
        Exit Sub
    End If
    
    If (txtPIN0(1) <> txtPIN0(1)) Then
        Mensaje.Escribir "Los números de PIN ingresados no coinciden."
        Exit Sub
    End If
    
    If Len(txtPIN0(1)) < 4 Then
        Mensaje.Escribir "El PIN tiene que tener por lo menos 4 caracteres."
        Exit Sub
    End If
    
    If Len(txtNewPassword0(1)) < 4 Then
        Mensaje.Escribir "La clave es muy corta, tiene que tener por lo menos 4 caracteres."
        Exit Sub
    End If

    If Not frmMain.Socket1.Connected Then
        Mensaje.Escribir "Error: Se ha perdido la conexion con el server."
    End If
    
    SendData ("NACCNT" & txtNewAccount & "," & txtNewPassword0(1) & "," & txtPIN0(1))
    
    mostrarConectar (True)
    
End Sub
Private Sub imgNewAccount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgNewAccount.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_createAccountHover.jpg"): modificoBoton = True
End Sub
Private Sub imgRECACC_Click()
    recupSelect(1) = True
    recupSelect(2) = False
    
    If txtRecAcc.Caption = "CUENTA" Then txtRecAcc.Caption = "": txtRecuAccount = ""
    recup_escribirDatos
End Sub
Private Sub imgRECACC_DblClick()
    recupSelect(1) = True
    recupSelect(2) = False
    txtRecAcc.Caption = "": txtRecuAccount = ""
    recup_escribirDatos
End Sub
Private Sub imgRecAccPin_Click()
    recupSelect(1) = False
    recupSelect(2) = True
    
    If txtRecAccPIN.Caption = "PIN" Then txtRecAccPIN.Caption = "": txtRecuPIN(1) = "": txtRecuPIN(2) = ""
    recup_escribirDatos
End Sub
Private Sub imgRecAccPin_DblClick()
    recupSelect(1) = False
    recupSelect(2) = True
    
    txtRecAccPIN.Caption = "": txtRecuPIN(1) = "": txtRecuPIN(2) = ""
    recup_escribirDatos
End Sub
Private Sub imgRecorder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Configuracion.recordarCuenta = 0 Then imgRecorder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_recorderI.jpg"): modificoBoton = True
End Sub
Private Sub imgRecorder_Click()
    If Configuracion.recordarCuenta = 0 Then
        Configuracion.recordarCuenta = 1
        imgRecorder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_recorderPress.jpg")
    Else
        Configuracion.recordarCuenta = 0
        imgRecorder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_recorderEmpty.jpg")
    End If
        
    Call guardarCuenta
End Sub
Private Sub imgRecuAcc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRecuAcc.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_RecuperarHover.jpg")
    modificoBoton = True
End Sub
Private Sub imgRecuAcc_Click()
    
    imgRecuAcc.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_RecuperarPress.jpg")
    SendData ("REECUH" & txtRecuAccount & "," & txtRecuPIN(1))
    mostrarConectar (True)
End Sub

Private Sub lblCreateAccount_Click()
    
    imgCrearCuenta_Click
End Sub

Private Sub lblRecu_Click()
    
    imgRecuperarCuenta_Click
End Sub
Private Sub lblSalir_Click()
    
    If pageSelected = 1 Then
        imgSalir_Click
    Else
        mostrarConectar (True)
    End If
End Sub

Private Sub Timer1_Timer()

    If lblBarrita.Caption = "" Then
        lblBarrita.Caption = "|"
    Else
        lblBarrita.Caption = ""
    End If

End Sub
Private Sub imgConectar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgConectar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_ButtonI.jpg")
    modificoBoton = True
End Sub
Private Sub imgConectar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgConectar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_ButtonA.jpg")
    modificoBoton = True
End Sub
Private Sub imgConectar_Click()

    On Error GoTo errh:

        Call guardarCuenta
        nombrecuent = TextBoxCuenta
        passcuent = TextBoxPassw

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            Sleep 30
        End If
        
        UserPassword = TextBoxPassw
       
        If CheckUserData(False) = True Then
            EstadoLogin = LoginAccount
            frmConnect.MousePointer = 99
            frmMain.Socket1.HostAddress = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
        End If
        
Exit Sub
errh:
    Mensaje.Escribir "Ocurrió un error inesperado, asegurate de tener todas las librerias correctamente registradas y vuelve a intentarlo."
End Sub
Private Sub imgCrearCuenta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCreateAccount.ForeColor = RGB(75, 212, 255)
End Sub
Private Sub imgCrearCuenta_Click()
    On Error GoTo errh:

    If pageSelected = 2 Then Exit Sub
    
       lblCreateAccount.ForeColor = RGB(0, 190, 255)

       If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
       End If
       
       EstadoLogin = CrearAccount
        
        frmMain.Socket1.HostAddress = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
        
Exit Sub
errh:
    Mensaje.Escribir "Ocurrió un error inesperado, asegurate de tener todas las librerias correctamente registradas y vuelve a intentarlo."
End Sub

Private Sub imgName_DblClick()
    TextBoxCuenta = ""
    conectar_escribirDatos
End Sub
Private Sub imgPass_DblClick()
     TextBoxPassw = ""
     TextBoxPasswR = ""
     conectar_escribirDatos
End Sub
Private Sub imgName_Click()
     ClickeoTextCuenta = True
     ClickeoTextPassw = False
     conectar_escribirDatos
End Sub
Private Sub imgPass_Click()
     ClickeoTextCuenta = False
     ClickeoTextPassw = True
     conectar_escribirDatos
End Sub
Private Sub imgRecuperarCuenta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If pageSelected <> 3 Then lblRecu.ForeColor = RGB(75, 212, 255)
End Sub
Private Sub imgRecuperarCuenta_Click()
    On Error GoTo errh:
        If pageSelected = 3 Then Exit Sub
            
        lblRecu.ForeColor = RGB(0, 190, 255)
        EstadoLogin = RecuPW

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
        
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect

Exit Sub
errh:
    Mensaje.Escribir "Ocurrió un error inesperado, asegurate de tener todas las librerias correctamente registradas y vuelve a intentarlo."
End Sub
Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSalir.ForeColor = RGB(75, 212, 255)
End Sub
Private Sub imgSalir_Click()

On Error GoTo errh:

    If pageSelected = 1 Then
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        Call UnloadAllForms
        lblSalir.ForeColor = RGB(0, 190, 255)
    Else
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        mostrarConectar (True)
    End If
    
Exit Sub
errh:
    Mensaje.Escribir "Ocurrió un error inesperado, asegurate de tener todas las librerias correctamente registradas y vuelve a intentarlo."
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    

    If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii = vbKeyBack) Or (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyTab) Then

    If pageSelected = 1 Then
        If KeyAscii = vbKeyTab And ClickeoTextCuenta = True Then
            ClickeoTextCuenta = False
            ClickeoTextPassw = True
            conectar_escribirDatos
        Exit Sub
        ElseIf KeyAscii = vbKeyTab And ClickeoTextPassw = True Then
            ClickeoTextCuenta = True
            ClickeoTextPassw = False
            conectar_escribirDatos
        Exit Sub
        End If
        
        If KeyAscii = vbKeyReturn Then
                nombrecuent = TextBoxCuenta
                passcuent = TextBoxPassw
        
                If frmMain.Socket1.Connected Then
                    frmMain.Socket1.Disconnect
                    frmMain.Socket1.Cleanup
                    DoEvents
                End If
               
               
                'update user info
                nombrecuent = TextBoxCuenta
                UserPassword = TextBoxPassw
               
                If CheckUserData(False) = True Then
                    EstadoLogin = LoginAccount
                    frmConnect.MousePointer = 99
                    frmMain.Socket1.HostAddress = CurServerIp
                    frmMain.Socket1.RemotePort = CurServerPort
                    frmMain.Socket1.Connect
                End If
                
                Call guardarCuenta
        Exit Sub
        End If
    
    
        If ClickeoTextCuenta = True Then
        
            If KeyAscii = vbKeyBack And Len(TextBoxCuenta) = 0 Then Exit Sub
        
            If KeyAscii = vbKeyBack And Len(TextBoxCuenta) <> 0 Then
                TextBoxCuenta = mid(TextBoxCuenta, 1, Len(TextBoxCuenta) - 1)
            Else
                If Len(TextBoxCuenta) >= 15 Then Exit Sub
                TextBoxCuenta = TextBoxCuenta & Chr$(KeyAscii)  'convert to character
            End If
            
        ElseIf ClickeoTextPassw = True Then
        
            If KeyAscii = vbKeyBack And Len(TextBoxPassw) = 0 Then Exit Sub
        
            If KeyAscii = vbKeyBack And Len(TextBoxPassw) <> 0 Then
                TextBoxPassw = mid(TextBoxPassw, 1, Len(TextBoxPassw) - 1)
                TextBoxPasswR = mid(TextBoxPasswR, 1, Len(TextBoxPasswR) - 1)
            Else
                If Len(TextBoxPassw) >= 15 Then Exit Sub
                TextBoxPassw = TextBoxPassw & Chr$(KeyAscii)  'convert to character
                TextBoxPasswR = TextBoxPasswR & "*"
            End If
            
        End If
        
        conectar_escribirDatos
    
    ElseIf pageSelected = 2 Then
    
        If KeyAscii = vbKeyTab Then
            If NA_clickeoText <= 4 Then NA_clickeoText = NA_clickeoText + 1
            If NA_clickeoText = 5 Or NA_clickeoText = 99 Then NA_clickeoText = 0
            
            Select Case NA_clickeoText
                Case 0
                    txtNuevaCuenta_DblClick
                Case 1, 2
                    txtNuevaPassword_DblClick (NA_clickeoText - 1)
                Case 3, 4
                    txtNuevoPIN_DblClick (NA_clickeoText - 3)
            End Select
            
            createAccount_escribirDatos
        Exit Sub
        End If
        
        If KeyAscii = vbKeyReturn Then
            imgNewAccount_Click
        Exit Sub
        End If
    
        Select Case NA_clickeoText
        
            Case 0
                If KeyAscii = vbKeyBack And Len(txtNewAccount) = 0 Then Exit Sub
            
                If KeyAscii = vbKeyBack And Len(txtNewAccount) <> 0 Then
                    txtNewAccount = mid(txtNewAccount, 1, Len(txtNewAccount) - 1)
                Else
                    If Len(txtNewAccount) >= 15 Then Exit Sub
                    txtNewAccount = txtNewAccount & Chr$(KeyAscii)  'convert to character
                End If
                
            Case 1
                Call escribirPassword(KeyAscii, txtNewPassword0(1), txtNewPassword0(2), 10)
            
            Case 2
                Call escribirPassword(KeyAscii, txtNewPassword1(1), txtNewPassword1(2), 10)
            
            Case 3
                Call escribirPassword(KeyAscii, txtPIN0(1), txtPIN0(2), 5)
                
            Case 4
                Call escribirPassword(KeyAscii, txtPIN1(1), txtPIN1(2), 5)
        End Select
        
        createAccount_escribirDatos
        
    ElseIf pageSelected = 3 Then
        
        If KeyAscii = vbKeyTab Then
            If (recupSelect(1)) Then
                imgRecAccPin_Click
            Else
                imgRECACC_Click
            End If
        Exit Sub
        End If
        
        If KeyAscii = vbKeyReturn Then
            imgRecuAcc_Click
        Exit Sub
        End If
        
        If (recupSelect(1)) Then
            If KeyAscii = vbKeyBack And Len(txtRecuAccount) = 0 Then Exit Sub
            
            If KeyAscii = vbKeyBack And Len(txtRecuAccount) <> 0 Then
                txtRecuAccount = mid(txtRecuAccount, 1, Len(txtRecuAccount) - 1)
            Else
                If Len(txtRecuAccount) >= 10 Then Exit Sub
                txtRecuAccount = txtRecuAccount & Chr$(KeyAscii)  'convert to character
            End If
        ElseIf (recupSelect(2)) Then
            Call escribirPassword(KeyAscii, txtRecuPIN(1), txtRecuPIN(2), 5)
        End If
        
        recup_escribirDatos
        
    End If
    
End If

End Sub
Private Sub escribirPassword(KeyAscii As Integer, pass1 As String, pass2 As String, ByVal limitLen As Byte)
    
    If KeyAscii = vbKeyBack And Len(pass1) = 0 Then Exit Sub
    
    If KeyAscii = vbKeyBack And Len(pass1) <> 0 Then
        pass1 = mid(pass1, 1, Len(pass1) - 1)
        pass2 = mid(pass2, 1, Len(pass2) - 1)
    Else
        If Len(pass2) >= limitLen Then Exit Sub
        pass1 = pass1 & Chr$(KeyAscii)  'convert to character
        pass2 = pass2 & "*"
    End If
End Sub
Private Sub conectar_escribirDatos()
    

    txtCuenta.Caption = UCase$(TextBoxCuenta)
    txtPassword.Caption = TextBoxPasswR
    
    If ClickeoTextPassw Then
        lblBarrita.top = 353
        lblBarrita.left = 404 + (Len(TextBoxPassw) * 3.8)
    Else
        lblBarrita.top = 293
        lblBarrita.left = 404 + (Len(TextBoxCuenta) * 5.5)
    End If
        

End Sub
Private Sub recup_escribirDatos()
    

    txtRecAcc.Caption = UCase$(txtRecuAccount)
    txtRecAccPIN.Caption = txtRecuPIN(2)
    
    If lblBarrita.Visible = False Then lblBarrita.Visible = True

    
    If recupSelect(1) Then
        lblBarrita.top = 353
        lblBarrita.left = 404 + (Len(txtRecuPIN(2)) * 3.8)
    ElseIf recupSelect(2) Then
        lblBarrita.top = 293
        lblBarrita.left = 404 + (Len(txtRecuAccount) * 5.5)
    End If
        

End Sub
Private Sub createAccount_escribirDatos()
    

    txtNuevaCuenta.Caption = UCase$(txtNewAccount)
    txtNuevaPassword(0).Caption = txtNewPassword0(2)
    txtNuevaPassword(1).Caption = txtNewPassword1(2)
    txtNuevoPIN(0).Caption = txtPIN0(2)
    txtNuevoPIN(1).Caption = txtPIN1(2)
    
    If lblBarrita.Visible = False Then lblBarrita.Visible = True


    Select Case NA_clickeoText
        Case 0
            lblBarrita.top = 228
            lblBarrita.left = 404 + (Len(txtNewAccount) * 5.5)
        
        Case 1
            lblBarrita.top = 228 + (NA_clickeoText * 59)
            lblBarrita.left = 404 + (Len(txtNewPassword0(2)) * 3.8)
            
        Case 2
            lblBarrita.top = 228 + (NA_clickeoText * 59)
            lblBarrita.left = 404 + (Len(txtNewPassword1(2)) * 3.8)
        
        Case 3
            lblBarrita.top = 228 + (NA_clickeoText * 59)
            lblBarrita.left = 404 + (Len(txtPIN0(2)) * 3.8)
            
        Case 4
            lblBarrita.top = 228 + (NA_clickeoText * 59)
            lblBarrita.left = 404 + (Len(txtPIN1(2)) * 3.8)
            
        Case Else
            lblBarrita.Visible = False
    End Select
        

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        Call UnloadAllForms
End If

End Sub
Public Sub limpiarConectar()
    
    If pageSelected = 1 Then
        If Configuracion.recordarCuenta = 1 Then
            imgRecorder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_recorderPress.jpg")
            
            TextBoxCuenta = Configuracion.tmpCuenta
            TextBoxPassw = Configuracion.tmpPassword
            
            TextBoxPasswR = ""
            Dim i As Long
            For i = 1 To Len(TextBoxPassw)
                TextBoxPasswR = TextBoxPasswR & "*"
            Next i
        Else
            imgRecorder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_recorderEmpty.jpg")
        End If
        
        'meter el recordar
        ClickeoTextCuenta = True
        conectar_escribirDatos
    End If
    
End Sub
Private Sub resetConnectForms()

    mostrarConectar (False)
    mostrarNuevaCuenta (False)
    mostrarRecuperarCuenta (False)

End Sub
Public Sub mostrarConectar(ByVal bool As Boolean, Optional bool2 As Boolean = False, Optional bool3 As Boolean = False)
    
    If bool2 Then
        TextBoxCuenta = ""
        TextBoxPassw = ""
        TextBoxPasswR = ""
    End If
    
    If bool Then
        resetConnectForms
        pageSelected = 1
        
        If Not bool3 Then
            limpiarConectar
        Else
            conectar_escribirDatos
        End If
        
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_Main.jpg")
        imgConectar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_ButtonN.jpg")
        
        'Seteamos colores
        txtCuenta.ForeColor = RGB(185, 169, 146)
        txtPassword.ForeColor = RGB(185, 169, 146)
        
        lblSalir.Caption = "SALIR"
        
        lblRecu.ForeColor = RGB(185, 169, 146)
        lblSalir.ForeColor = RGB(185, 169, 146)
        lblCreateAccount.ForeColor = RGB(185, 169, 146)
        
        Me.Show
    End If


    'Seteamos cuenta y pass
    txtCuenta.Visible = bool
    txtPassword.Visible = bool
    imgRecorder.Visible = bool
    imgConectar.Visible = bool

End Sub
Public Sub mostrarNuevaCuenta(ByVal bool As Boolean)
    
    Dim i As Long
    
    If bool Then
        resetConnectForms
        
        NA_clickeoText = 99
        pageSelected = 2
        
        txtNuevaCuenta.Caption = "CUENTA": txtNewAccount = "CUENTA"
        txtNuevaPassword(0).Caption = "PASSWORD": txtNewPassword0(2) = "PASSWORD"
        txtNuevaPassword(1).Caption = "REPETIR PASSWORD": txtNewPassword1(2) = "REPETIR PASSWORD"
        txtNuevoPIN(0).Caption = "PIN": txtPIN0(2) = "PIN"
        txtNuevoPIN(1).Caption = "REPETIR PIN": txtPIN1(2) = "REPETIR PIN"
        lblSalir.Caption = "ATRÁS"
        
        'Seteamos colores
        txtNuevaCuenta.ForeColor = RGB(185, 169, 146)
        
        For i = 0 To 1
            txtNuevaPassword(i).ForeColor = RGB(185, 169, 146)
            txtNuevoPIN(i).ForeColor = RGB(185, 169, 146)
        Next i
        
        lblBarrita.Visible = False
        
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_createAccountMain.jpg")
    End If
    
    For i = 0 To 4
        imgNEWACC(i).Visible = bool
    Next i
        
    imgNewAccount.Visible = bool
    txtNuevaCuenta.Visible = bool
    txtNuevaPassword(0).Visible = bool
    txtNuevaPassword(1).Visible = bool
    txtNuevoPIN(0).Visible = bool
    txtNuevoPIN(1).Visible = bool

End Sub
Public Sub mostrarRecuperarCuenta(ByVal bool As Boolean)
       
    If bool Then
        resetConnectForms
        
        pageSelected = 3
        
        recupSelect(1) = False
        recupSelect(2) = False
        txtRecAcc.Caption = "CUENTA": txtRecuAccount = "CUENTA"
        txtRecAccPIN.Caption = "PIN": txtRecuPIN(1) = "PIN": txtRecuPIN(2) = "PIN"
        lblSalir.Caption = "ATRÁS"
        txtRecAcc.ForeColor = RGB(185, 169, 146)
        txtRecAccPIN.ForeColor = RGB(185, 169, 146)
        
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_RecuperarPersonajeMain.jpg")
        lblBarrita.Visible = False
    End If
    
    txtRecAcc.Visible = bool
    txtRecAccPIN.Visible = bool
    imgRecuAcc.Visible = bool
    imgRECACC.Visible = bool
    imgRecAccPin.Visible = bool

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If pageSelected = 1 Then
        If modificoBoton Then
            If Configuracion.recordarCuenta = 1 Then
                imgRecorder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_recorderPress.jpg")
            Else
                imgRecorder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_recorderEmpty.jpg")
            End If
            
            imgConectar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_ButtonN.jpg")
            modificoBoton = False
        End If
        
    ElseIf pageSelected = 2 Then
        imgNewAccount.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_createAccount.jpg")
    ElseIf pageSelected = 3 Then
        If modificoBoton Then imgRecuAcc.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\PRINCIPAL\Connect_Recuperar.jpg")
    End If
        
        If pageSelected <> 3 Then lblRecu.ForeColor = RGB(185, 169, 146)
        lblSalir.ForeColor = RGB(185, 169, 146)
        If pageSelected <> 2 Then lblCreateAccount.ForeColor = RGB(185, 169, 146)
End Sub

Private Sub txtNuevaCuenta_Click()
    If txtNuevaCuenta.Caption = "CUENTA" Then txtNuevaCuenta.Caption = "": txtNewAccount = ""
    createAccount_escribirDatos
End Sub
Private Sub txtNuevaCuenta_DblClick()
    txtNuevaCuenta.Caption = "": txtNuevaCuenta.Caption = "": txtNewAccount = ""
    createAccount_escribirDatos
End Sub
Private Sub txtNuevaPassword_Click(Index As Integer)
    If txtNuevaPassword(Index).Caption = "PASSWORD" Then txtNuevaPassword(Index).Caption = "": txtNewPassword0(1) = "": txtNewPassword0(2) = ""
    If txtNuevaPassword(Index).Caption = "REPETIR PASSWORD" Then txtNuevaPassword(Index).Caption = "": txtNewPassword1(1) = "": txtNewPassword1(2) = ""
    createAccount_escribirDatos
End Sub
Private Sub txtNuevaPassword_DblClick(Index As Integer)
    If Index = 0 Then txtNuevaPassword(0).Caption = "": txtNewPassword0(1) = "": txtNewPassword0(2) = ""
    If Index = 1 Then txtNuevaPassword(1).Caption = "": txtNewPassword1(1) = "": txtNewPassword1(2) = ""
    createAccount_escribirDatos
End Sub
Private Sub txtNuevoPIN_Click(Index As Integer)
    If txtNuevoPIN(Index).Caption = "PIN" Then txtNuevoPIN(Index).Caption = "": txtPIN0(1) = "": txtPIN0(2) = ""
    If txtNuevoPIN(Index).Caption = "REPETIR PIN" Then txtNuevoPIN(Index).Caption = "": txtPIN1(1) = "": txtPIN1(2) = ""
    createAccount_escribirDatos
End Sub
Private Sub txtNuevoPIN_DblClick(Index As Integer)
    If Index = 0 Then txtNuevoPIN(Index).Caption = "": txtPIN0(1) = "": txtPIN0(2) = ""
    If Index = 1 Then txtNuevoPIN(Index).Caption = "": txtPIN1(1) = "": txtPIN1(2) = ""
    createAccount_escribirDatos
End Sub

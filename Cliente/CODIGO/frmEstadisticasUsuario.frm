VERSION 5.00
Begin VB.Form frmEstadisticasUsuario 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   Icon            =   "frmEstadisticasUsuario.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblParejas 
      BackStyle       =   0  'Transparent
      Caption         =   "8421 jugados (67% victorias)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4530
      TabIndex        =   15
      Top             =   1215
      Width           =   2055
   End
   Begin VB.Label lblDuelos 
      BackStyle       =   0  'Transparent
      Caption         =   "244 jugados (15% victorias)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4530
      TabIndex        =   14
      Top             =   825
      Width           =   1935
   End
   Begin VB.Label lblRondas 
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   1590
      Width           =   615
   End
   Begin VB.Label lblUsuariosMatados 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4.811"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   2370
      Width           =   495
   End
   Begin VB.Label lblMuertes 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   1980
      Width           =   1695
   End
   Begin VB.Label lblQuests 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1.343"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5985
      TabIndex        =   10
      Top             =   3555
      Width           =   495
   End
   Begin VB.Label lblCVCS 
      BackStyle       =   0  'Transparent
      Caption         =   "10.654"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   3150
      Width           =   2055
   End
   Begin VB.Label lblEventos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "22.377"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblReputacion 
      BackStyle       =   0  'Transparent
      Caption         =   "20.000"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   3510
      Width           =   1335
   End
   Begin VB.Label lblJerarquia 
      BackStyle       =   0  'Transparent
      Caption         =   "4ta Jerarquia"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   3135
      Width           =   1335
   End
   Begin VB.Label lblFaccion 
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2775
      Width           =   1815
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      Caption         =   "999999999/999999999"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   2355
      Width           =   2295
   End
   Begin VB.Label lblNivel 
      BackStyle       =   0  'Transparent
      Caption         =   "50 + 20"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1965
      Width           =   1815
   End
   Begin VB.Label lblRaza 
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1575
      Width           =   1815
   End
   Begin VB.Label lblClase 
      BackStyle       =   0  'Transparent
      Caption         =   "Mago"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1185
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Shay"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   735
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6360
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmEstadisticasUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

If formuEstadisticas.Nivel > 50 Then
    lblNivel.Caption = "50 + " & formuEstadisticas.Nivel - 50 & ""
Else
    lblNivel.Caption = formuEstadisticas.Nivel
End If

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\EstadisticasUser_Main.jpg")

With Me
    Me.lblNombre.ForeColor = RGB(185, 169, 146)
    Me.lblClase.ForeColor = RGB(185, 169, 146)
    Me.lblRaza.ForeColor = RGB(185, 169, 146)
    Me.lblNivel.ForeColor = RGB(185, 169, 146)
    Me.lblJerarquia.ForeColor = RGB(185, 169, 146)
    Me.lblDuelos.ForeColor = RGB(185, 169, 146)
    Me.lblParejas.ForeColor = RGB(185, 169, 146)
    Me.lblEventos.ForeColor = RGB(185, 169, 146)
    Me.lblCVCS.ForeColor = RGB(185, 169, 146)
    Me.lblQuests.ForeColor = RGB(185, 169, 146)
    Me.lblMuertes.ForeColor = RGB(185, 169, 146)
    Me.lblUsuariosMatados.ForeColor = RGB(185, 169, 146)
    Me.lblRondas.ForeColor = RGB(185, 169, 146)
    Me.lblExp.ForeColor = RGB(185, 169, 146)
End With

If formuEstadisticas.Faccion = 1 Then
    lblFaccion.ForeColor = &H80&
    lblFaccion.Caption = "HORDA INFERNAL"
ElseIf formuEstadisticas.Faccion = 2 Then
    lblFaccion.ForeColor = &HC00000
    lblFaccion.Caption = "ALIANZA IMPERIAL"
ElseIf formuEstadisticas.Faccion = 0 Then
    lblFaccion.ForeColor = &H404040
    lblFaccion.Caption = "NEUTRAL"
End If
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

VERSION 5.00
Begin VB.Form frmMisionesDiarias 
   Caption         =   "Misiones Diarias"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form2"
   ScaleHeight     =   4230
   ScaleWidth      =   6405
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Información"
      Height          =   1935
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   3495
      Begin VB.TextBox lblInformacion 
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Text            =   "acá va una breve info sobre la misión"
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lblNombre 
         Caption         =   "Mision: Asesinar 10 usuarios."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdReclamar 
      Caption         =   "Reclamar Premio"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "MISIONES DIARIAS"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1440
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label lblComplete 
      Caption         =   "Completado: 0/999"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
End
Attribute VB_Name = "frmMisionesDiarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ParseQuest(ByVal Buffer As String)

    lblNombre.Caption = ReadField(1, Buffer, Asc(","))
    lblInformacion.text = ReadField(2, Buffer, Asc(","))

    Dim Progreso As Long
    Dim Necesario As Long
        Progreso = ReadField(3, Buffer, Asc(","))
        Necesario = ReadField(4, Buffer, Asc(","))
        
        'BarraCompletada.width = (((progreso / 100) / (necesario / 100)) * width barra)
        lblComplete.Caption = "Completado: " & PonerPuntos(Progreso) & "/" & PonerPuntos(Necesario)
        
        If Progreso < Necesario Then
            cmdReclamar.Enabled = False
        End If
        
    
    Me.Show , frmMain
    
End Sub

VERSION 5.00
Begin VB.Form frmEngine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Engine TSAO"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5880
   Icon            =   "frmEngine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "GUARDAR CAMBIOS EN EL MAPA DE FORMA PERMANENTE"
      Height          =   495
      Left            =   120
      TabIndex        =   48
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Frame Frame5 
      Caption         =   "Personaje"
      Height          =   1695
      Left            =   3000
      TabIndex        =   27
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton restCasco 
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   47
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton restWeapon 
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   46
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton restEscu 
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   45
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton restBody 
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   44
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton restCasco 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   43
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton restWeapon 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   42
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton restEscu 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   41
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton restBody 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   40
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton restHead 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   39
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton restHead 
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   38
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtCasco 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Text            =   "1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtWeapon 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Text            =   "1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtShield 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Text            =   "1"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtBody 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   29
         Text            =   "1"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtHead 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   28
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Casco"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Arma"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Escudo:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Body:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Head:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Crear Aura"
      Height          =   975
      Left            =   3000
      TabIndex        =   23
      Top             =   1920
      Width           =   2655
      Begin VB.CommandButton Command4 
         Caption         =   "Crear Aura"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Aurix 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Text            =   "23"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "N° Aura:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Crearla sobre el Personaje"
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Particulas"
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   2775
      Begin VB.TextBox Time 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Namber 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Text            =   "0"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Crearla en el Mapa"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Crear Particula"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Tiempo:"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Numero:"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Crear Luces"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2775
      Begin VB.TextBox Range 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   14
         Text            =   "3"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Blue 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Text            =   "255"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Green 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Text            =   "255"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox red 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "255"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Crear Luz"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Range:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Blue:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Geen:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Red:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox erre 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Text            =   "255"
      Top             =   480
      Width           =   420
   End
   Begin VB.TextBox ge 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "255"
      Top             =   480
      Width           =   420
   End
   Begin VB.TextBox be 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "255"
      Top             =   480
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Luz del Render"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call SendData("/MODMAPINFO RGB " & erre.text & " " & ge.text & " " & be.text)
End Sub

Private Sub Command2_Click()
    Call SendData("/MODMAPINFO LUZ " & Range.text & " " & red.text & " " & Blue.text & " " & Green.text)
End Sub
Private Sub Command3_Click()

If Check1.Value = 0 And Check2.Value = 0 Then MsgBox "Elegí si la queres sobre el mapa o sobre el personaje down."

If Check1.Value = 1 Then
    Call SendData("/MODMAPINFO PART " & Val(Namber.text))
End If

If Check2.Value = 1 Then
    Call SendData("/MOD PART " & Val(Namber.text))
End If


End Sub
Private Sub Command4_Click()
    Call SendData("/MOD AURA " & Aurix.text)
End Sub

Private Sub Command5_Click()
    Call SendData("/GUARDARMAPA")
End Sub
Private Sub restBody_Click(Index As Integer)
    Select Case Index
      Case 0
        txtBody.text = txtBody.text - 1
        
      Case 1
       txtBody.text = txtBody.text + 1
    End Select
    
    Call SendData("/MOD BODY " & Val(txtBody.text))
End Sub

Private Sub restCasco_Click(Index As Integer)
    
    Select Case Index
      Case 0
        txtCasco.text = txtCasco.text - 1
        
      Case 1
       txtCasco.text = txtCasco.text + 1
    End Select
    
    Call SendData("/MOD CASCO " & Val(txtCasco.text))
    
End Sub

Private Sub restEscu_Click(Index As Integer)

    Select Case Index
      Case 0
        txtShield.text = txtShield.text - 1
        
      Case 1
       txtShield.text = txtShield.text + 1
    End Select
    
    Call SendData("/MOD ESCU " & Val(txtShield.text))
    
End Sub
Private Sub restHead_Click(Index As Integer)
    Select Case Index
      Case 0
        txtHead.text = txtHead.text - 1
        
      Case 1
       txtHead.text = txtHead.text + 1
    End Select
    
    Call SendData("/MOD HEAD " & Val(txtHead.text))
End Sub

Private Sub restWeapon_Click(Index As Integer)
    Select Case Index
      Case 0
        txtWeapon.text = txtWeapon.text - 1
        
      Case 1
        txtWeapon.text = txtWeapon.text + 1
    End Select
    
    Call SendData("/MOD ARMA " & Val(txtWeapon.text))
End Sub

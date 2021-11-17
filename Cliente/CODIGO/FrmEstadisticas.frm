VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   7395
   ClientLeft      =   4215
   ClientTop       =   1635
   ClientWidth     =   5625
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInv 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2100
      Left            =   300
      ScaleHeight     =   2100
      ScaleWidth      =   1740
      TabIndex        =   69
      Top             =   1080
      Width           =   1740
   End
   Begin VB.TextBox txtQuestDescription 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   1575
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   5160
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   70
      Top             =   165
      Width           =   255
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   25
      Left            =   2520
      TabIndex        =   68
      Top             =   5533
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   23
      Left            =   2626
      TabIndex        =   67
      Top             =   4870
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   24
      Left            =   2975
      TabIndex        =   66
      Top             =   5201
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   22
      Left            =   1680
      TabIndex        =   65
      Top             =   4515
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   21
      Left            =   2520
      TabIndex        =   64
      Top             =   4185
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   20
      Left            =   4080
      TabIndex        =   63
      Top             =   3074
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   19
      Left            =   2520
      TabIndex        =   62
      Top             =   2730
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   18
      Left            =   2760
      TabIndex        =   61
      Top             =   2385
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   17
      Left            =   2640
      TabIndex        =   60
      Top             =   2085
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   16
      Left            =   1560
      TabIndex        =   59
      Top             =   1725
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   15
      Left            =   1440
      TabIndex        =   58
      Top             =   1402
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image imgHoja 
      Height          =   450
      Index           =   2
      Left            =   4530
      Top             =   480
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgHoja 
      Height          =   450
      Index           =   1
      Left            =   3555
      Top             =   480
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   14
      Left            =   3240
      TabIndex        =   57
      Top             =   5834
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   56
      Top             =   4823
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   55
      Top             =   4488
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   13
      Left            =   2490
      TabIndex        =   54
      Top             =   5834
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   12
      Left            =   3360
      TabIndex        =   53
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   11
      Left            =   2280
      TabIndex        =   52
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   2244
      TabIndex        =   51
      Top             =   4823
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   50
      Top             =   4488
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   49
      Top             =   4150
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   2370
      TabIndex        =   48
      Top             =   3071
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   47
      Top             =   2730
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   46
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   45
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   44
      Top             =   1717
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   43
      Top             =   1402
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   22
      Left            =   4950
      TabIndex        =   42
      Top             =   5040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   21
      Left            =   4950
      TabIndex        =   41
      Top             =   4701
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   20
      Left            =   4950
      TabIndex        =   40
      Top             =   4365
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   19
      Left            =   4950
      TabIndex        =   39
      Top             =   4040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   18
      Left            =   4950
      TabIndex        =   38
      Top             =   3701
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   17
      Left            =   4950
      TabIndex        =   37
      Top             =   3355
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   16
      Left            =   4950
      TabIndex        =   36
      Top             =   3000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   15
      Left            =   4950
      TabIndex        =   35
      Top             =   2662
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   14
      Left            =   4950
      TabIndex        =   34
      Top             =   2323
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   13
      Left            =   4950
      TabIndex        =   33
      Top             =   2002
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   12
      Left            =   4950
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   1
      Left            =   2340
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   2
      Left            =   2340
      TabIndex        =   30
      Top             =   2002
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   3
      Left            =   2340
      TabIndex        =   29
      Top             =   2323
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   4
      Left            =   2340
      TabIndex        =   28
      Top             =   2662
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   5
      Left            =   2340
      TabIndex        =   27
      Top             =   3000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   6
      Left            =   2340
      TabIndex        =   26
      Top             =   3355
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   7
      Left            =   2340
      TabIndex        =   25
      Top             =   3701
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   8
      Left            =   2340
      TabIndex        =   24
      Top             =   4040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   9
      Left            =   2340
      TabIndex        =   23
      Top             =   4365
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   10
      Left            =   2340
      TabIndex        =   22
      Top             =   4701
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   11
      Left            =   2340
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label imgQuestVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label imgQuestVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   3850
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgQuestAbandonar 
      Height          =   540
      Left            =   320
      Top             =   5924
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.Image imgQuestType 
      Height          =   1515
      Left            =   360
      Top             =   3750
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.Label lblAtri 
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   5
      Left            =   4005
      TabIndex        =   17
      Top             =   4740
      Width           =   495
   End
   Begin VB.Label lblAtri 
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   4
      Left            =   2085
      TabIndex        =   16
      Top             =   4740
      Width           =   495
   End
   Begin VB.Label lblAtri 
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   15
      Top             =   4425
      Width           =   495
   End
   Begin VB.Label lblAtri 
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   2805
      TabIndex        =   14
      Top             =   4425
      Width           =   495
   End
   Begin VB.Label lblBonificadores 
      BackStyle       =   0  'Transparent
      Caption         =   "No elegido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   13
      Top             =   6271
      Width           =   4575
   End
   Begin VB.Label lblBonificadores 
      BackStyle       =   0  'Transparent
      Caption         =   "No elegido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   12
      Top             =   6000
      Width           =   4575
   End
   Begin VB.Label lblBonificadores 
      BackStyle       =   0  'Transparent
      Caption         =   "No elegido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   5702
      Width           =   4575
   End
   Begin VB.Label lblAtri 
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   1455
      TabIndex        =   10
      Top             =   4425
      Width           =   480
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Prueba"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   655
      Width           =   1935
   End
   Begin VB.Label lblPuntosDonador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3600
      MouseIcon       =   "FrmEstadisticas.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3647
      Width           =   1215
   End
   Begin VB.Label lblPuntosTorneo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   720
      MouseIcon       =   "FrmEstadisticas.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3647
      Width           =   1215
   End
   Begin VB.Label lblMail 
      BackStyle       =   0  'Transparent
      Caption         =   "a@a.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   2978
      Width           =   2535
   End
   Begin VB.Label lblHogar 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanaris"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   2670
      Width           =   2175
   End
   Begin VB.Label lblGenero 
      BackStyle       =   0  'Transparent
      Caption         =   "Hombre"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   2288
      Width           =   1935
   End
   Begin VB.Label lblClase 
      BackStyle       =   0  'Transparent
      Caption         =   "Mago"
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
      Left            =   3000
      TabIndex        =   3
      Top             =   2002
      Width           =   2295
   End
   Begin VB.Label lblRaza 
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo Oscuro"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   1720
      Width           =   2295
   End
   Begin VB.Label lblReputacion 
      BackStyle       =   0  'Transparent
      Caption         =   "104.250"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   1331
      Width           =   1935
   End
   Begin VB.Image imgExtras 
      Height          =   630
      Left            =   4343
      Top             =   6644
      Width           =   1140
   End
   Begin VB.Image imgHabilidades 
      Height          =   630
      Left            =   1560
      Top             =   6645
      Width           =   1440
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      Caption         =   "50 + 10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2835
      TabIndex        =   0
      Top             =   1035
      Width           =   735
   End
   Begin VB.Image imgQuests 
      Height          =   630
      Left            =   3280
      Top             =   6644
      Width           =   1020
   End
   Begin VB.Image imgGeneral 
      Height          =   630
      Left            =   135
      Top             =   6644
      Width           =   1440
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Me.Picture = General_Load_Interface_Picture("Estadisticas_1_Main.jpg")
imgGeneral.Picture = General_Load_Interface_Picture("Estadisticas_1_Principal_N.jpg")
imgHabilidades.Picture = General_Load_Interface_Picture("Estadisticas_1_Habilidades_N.jpg")
imgQuests.Picture = General_Load_Interface_Picture("Estadisticas_1_Quest_N.jpg")
imgExtras.Picture = General_Load_Interface_Picture("Estadisticas_1_Extras_N.jpg")

'Ponemos visible toda la primera pagina.
lblNombre.Visible = True
lblLvl.Visible = True
lblRaza.Visible = True
lblGenero.Visible = True
lblHogar.Visible = True
lblMail.Visible = True
lblClase.Visible = True
lblReputacion.Visible = True
lblPuntosTorneo.Visible = True
lblPuntosDonador.Visible = True
lblBonificadores(1).Visible = True
lblBonificadores(2).Visible = True
lblBonificadores(3).Visible = True
picInv.Visible = True
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = True
Next

'Vaciamos toda la pagina "extra" y "habilidades"
For i = 0 To 25
lblCounters(i).Visible = False
Next
For i = 1 To NUMSKILLS
Skills(i).Visible = False
Next
imgHoja(1).Visible = False
imgHoja(2).Visible = False

'Vaciamos la pagina "quest"
txtQuestDescription.Visible = False
imgQuestVal1.Visible = False
imgQuestVal2.Visible = False
imgQuestAbandonar.Visible = False
imgQuestType.Visible = False

Call Iniciar_Labels
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Volvemos las imagenes a la normalidad
imgGeneral.Picture = General_Load_Interface_Picture("Estadisticas_1_Principal_N.jpg")
imgHabilidades.Picture = General_Load_Interface_Picture("Estadisticas_1_Habilidades_N.jpg")
imgQuests.Picture = General_Load_Interface_Picture("Estadisticas_1_Quest_N.jpg")
imgExtras.Picture = General_Load_Interface_Picture("Estadisticas_1_Extras_N.jpg")
imgHoja(1).Picture = Nothing
imgHoja(2).Picture = Nothing
imgQuestAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_2_AbandonarQuest_N.jpg")

End Sub
Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills

For i = 1 To NUMATRIBUTOS
    lblAtri(i).Caption = UserAtributos(i)
Next
For i = 1 To NUMSKILLS
    Skills(i).Caption = UserSkills(i)
Next

'### PRIMERA HOJA
lblNombre.Caption = UserEstadisticas.Nombre

If UserEstadisticas.Nivel > 50 Then
lblLvl.Caption = "50 + " & UserEstadisticas.Nivel - 50 & ""
Else
lblLvl.Caption = UserEstadisticas.Nivel
End If

lblRaza.Caption = UserEstadisticas.Raza
lblClase.Caption = UserEstadisticas.Clase
lblGenero.Caption = UserEstadisticas.Genero
lblHogar.Caption = UserEstadisticas.Hogar
lblMail.Caption = UserEstadisticas.Email
lblPuntosTorneo.Caption = UserEstadisticas.PuntosTorneo
lblPuntosDonador.Caption = PonerPuntos(UserEstadisticas.PuntosDonador)
lblBonificadores(1).Caption = UserEstadisticas.Bonif1
lblBonificadores(2).Caption = UserEstadisticas.Bonif2
lblBonificadores(3).Caption = UserEstadisticas.Bonif3
lblReputacion.Caption = UserEstadisticas.UserReputacion

'### EXTRAS - HOJA 1
lblCounters(0).Caption = UserEstadisticas.TorneosParticipados
lblCounters(1).Caption = "0"
lblCounters(2).Caption = "0"
lblCounters(3).Caption = UserEstadisticas.CopasDeOro
lblCounters(4).Caption = UserEstadisticas.CopasDePlata
lblCounters(5).Caption = UserEstadisticas.CopasDeBronce
lblCounters(6).Caption = UserEstadisticas.Eventos
lblCounters(7).Caption = UserEstadisticas.DuelosGanados
lblCounters(8).Caption = UserEstadisticas.DuelosGanados + UserEstadisticas.DuelosPerdidos
lblCounters(9).Caption = UserEstadisticas.ParejasGanadas
lblCounters(10).Caption = UserEstadisticas.ParejasGanadas + UserEstadisticas.ParejasPerdidas
lblCounters(11).Caption = UserEstadisticas.CvcsGanados
lblCounters(12).Caption = UserEstadisticas.MaximasRondas
lblCounters(13).Caption = UserEstadisticas.GuerrasGanadas
lblCounters(14).Caption = UserEstadisticas.GuerrasGanadas + UserEstadisticas.GuerrasPerdidas

'### EXTRAS - HOJA 2
If UserEstadisticas.Alineacion = 1 Then
    lblCounters(15).ForeColor = &H80&
    lblCounters(15) = "HORDA INFERNAL"
ElseIf UserEstadisticas.Alineacion = 2 Then
    lblCounters(15).ForeColor = &HC00000
    lblCounters(15) = "ALIANZA IMPERIAL"
ElseIf UserEstadisticas.Alineacion = 0 Then
    lblCounters(15).ForeColor = &H404040
    lblCounters(15) = "NEUTRAL"
End If

lblCounters(16).Caption = UserEstadisticas.Jerarquia
lblCounters(17).Caption = UserEstadisticas.CiudadanosMatados
lblCounters(18).Caption = UserEstadisticas.NeutralesMatados
lblCounters(19).Caption = UserEstadisticas.CriminalesMatados
lblCounters(20).Caption = UserEstadisticas.Restantes
lblCounters(21).Caption = UserEstadisticas.NPCSMATADOS
lblCounters(22).Caption = UserEstadisticas.MuertesUsuario
lblCounters(23).Caption = UserEstadisticas.QuestCompletadas
lblCounters(24).Caption = "0"
lblCounters(25).Caption = UserEstadisticas.MVPMatados

'### HOJA QUEST
txtQuestDescription.text = "Tipo " & UserEstadisticas.TipoQuest & ": " & UserEstadisticas.DescQuest & " Premio: " & PonerPuntos(UserEstadisticas.PremioOro) & " monedas y " & UserEstadisticas.PremioPuntis & " puntos de torneo."
imgQuestVal1.Caption = UserEstadisticas.CantidadNPCs
imgQuestVal2.Caption = UserEstadisticas.CantidadNPCs - UserEstadisticas.YaMatados
imgQuestType.Picture = General_Load_Interface_Picture("Estadisticas_2_Criaturas.jpg")
imgQuestAbandonar.Picture = General_Load_Interface_Picture("Estadisticas_2_AbandonarQuest_N.jpg")
 
End Sub
Private Sub Image1_Click()
Unload Me
End Sub
Private Sub imgGeneral_Click()
Me.Picture = General_Load_Interface_Picture("Estadisticas_1_Main.jpg")

'Ponemos visible toda la primera pagina.
lblNombre.Visible = True
lblLvl.Visible = True
lblRaza.Visible = True
lblClase.Visible = True
lblGenero.Visible = True
lblHogar.Visible = True
lblMail.Visible = True
lblReputacion.Visible = True
lblPuntosTorneo.Visible = True
lblPuntosDonador.Visible = True
lblBonificadores(1).Visible = True
lblBonificadores(2).Visible = True
lblBonificadores(3).Visible = True
picInv.Visible = True
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = True
Next

'Vaciamos toda la pagina "extra" y "habilidades"
For i = 0 To 25
lblCounters(i).Visible = False
Next
For i = 1 To NUMSKILLS
Skills(i).Visible = False
Next
imgHoja(1).Visible = False
imgHoja(2).Visible = False

'Vaciamos la pagina "quest"
txtQuestDescription.Visible = False
imgQuestVal1.Visible = False
imgQuestVal2.Visible = False
imgQuestAbandonar.Visible = False
imgQuestType.Visible = False
End Sub
Private Sub imgHabilidades_Click()
Me.Picture = General_Load_Interface_Picture("Estadisticas_3_Main.jpg")

'Vaciamos toda la primera pagina.
lblNombre.Visible = False
lblLvl.Visible = False
lblRaza.Visible = False
lblClase.Visible = False
lblGenero.Visible = False
lblHogar.Visible = False
lblMail.Visible = False
lblReputacion.Visible = False
lblPuntosTorneo.Visible = False
lblPuntosDonador.Visible = False
lblBonificadores(1).Visible = False
lblBonificadores(2).Visible = False
lblBonificadores(3).Visible = False
picInv.Visible = False
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = False
Next

'Vaciamos toda la pagina "extra" y ponemos visible "habilidades"
For i = 0 To 25
lblCounters(i).Visible = False
Next
For i = 1 To NUMSKILLS
Skills(i).Visible = True
Next
imgHoja(1).Visible = False
imgHoja(2).Visible = False

'Vaciamos la pagina "quest"
txtQuestDescription.Visible = False
imgQuestVal1.Visible = False
imgQuestVal2.Visible = False
imgQuestAbandonar.Visible = False
imgQuestType.Visible = False
End Sub
Private Sub imgQuests_Click()
Me.Picture = General_Load_Interface_Picture("Estadisticas_2_Main.jpg")

'Vaciamos toda la primera pagina.
lblNombre.Visible = False
lblLvl.Visible = False
lblRaza.Visible = False
lblClase.Visible = False
lblGenero.Visible = False
lblHogar.Visible = False
lblMail.Visible = False
lblReputacion.Visible = False
lblPuntosTorneo.Visible = False
lblPuntosDonador.Visible = False
lblBonificadores(1).Visible = False
lblBonificadores(2).Visible = False
lblBonificadores(3).Visible = False
picInv.Visible = False
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = False
Next

'Vaciamos toda la pagina "extra" y "habilidades"
For i = 0 To 25
lblCounters(i).Visible = False
Next
For i = 1 To NUMSKILLS
Skills(i).Visible = False
Next
imgHoja(1).Visible = False
imgHoja(2).Visible = False

'Mostramos pagina "quest"
If UserEstadisticas.TipoQuest = 0 Then
txtQuestDescription.Visible = True
txtQuestDescription.text = "No estás haciendo ninguna quest."
imgQuestVal1.Visible = False
imgQuestVal2.Visible = False
imgQuestAbandonar.Visible = False
imgQuestType.Visible = False
Else
txtQuestDescription.Visible = True
imgQuestVal1.Visible = True
imgQuestVal2.Visible = True
imgQuestAbandonar.Visible = True
imgQuestType.Visible = True
End If

End Sub
Private Sub imgExtras_Click()
Me.Picture = General_Load_Interface_Picture("Estadisticas_4_1_Main.jpg")
imgHoja(1).Picture = Nothing
imgHoja(2).Picture = Nothing

'Vaciamos toda la primera pagina.
lblNombre.Visible = False
lblLvl.Visible = False
lblRaza.Visible = False
lblGenero.Visible = False
lblClase.Visible = False
lblHogar.Visible = False
lblMail.Visible = False
lblReputacion.Visible = False
lblPuntosTorneo.Visible = False
lblPuntosDonador.Visible = False
lblBonificadores(1).Visible = False
lblBonificadores(2).Visible = False
lblBonificadores(3).Visible = False
picInv.Visible = False
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = False
Next

'Mostramos primera hoja de la pagina "extra" y vaciamos "habilidades"
For i = 0 To 14
lblCounters(i).Visible = True
Next
For i = 15 To 25
lblCounters(i).Visible = False
Next
For i = 1 To NUMSKILLS
Skills(i).Visible = False
Next

imgHoja(1).Visible = True
imgHoja(2).Visible = True

'Vaciamos la pagina "quest"
txtQuestDescription.Visible = False
imgQuestVal1.Visible = False
imgQuestVal2.Visible = False
imgQuestAbandonar.Visible = False
imgQuestType.Visible = False
End Sub
Private Sub imgHoja_Click(Index As Integer)
If Index = 1 Then
Me.Picture = General_Load_Interface_Picture("Estadisticas_4_1_Main.jpg")
imgHoja(1).Picture = Nothing
imgHoja(2).Picture = Nothing

'Vaciamos toda la primera pagina.
lblNombre.Visible = False
lblLvl.Visible = False
lblRaza.Visible = False
lblGenero.Visible = False
lblClase.Visible = False
lblHogar.Visible = False
lblMail.Visible = False
lblReputacion.Visible = False
lblPuntosTorneo.Visible = False
lblPuntosDonador.Visible = False
lblBonificadores(1).Visible = False
lblBonificadores(2).Visible = False
lblBonificadores(3).Visible = False
picInv.Visible = False
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = False
Next

'Mostramos primera hoja de la pagina "extra" y vaciamos "habilidades"
For i = 0 To 14
lblCounters(i).Visible = True
Next
For i = 15 To 25
lblCounters(i).Visible = False
Next
For i = 1 To NUMSKILLS
Skills(i).Visible = False
Next
imgHoja(1).Visible = True
imgHoja(2).Visible = True

'Vaciamos la pagina "quest"
txtQuestDescription.Visible = False
imgQuestVal1.Visible = False
imgQuestVal2.Visible = False
imgQuestAbandonar.Visible = False
imgQuestType.Visible = False

ElseIf Index = 2 Then
Me.Picture = General_Load_Interface_Picture("Estadisticas_4_2_Main.jpg")
imgHoja(1).Picture = Nothing
imgHoja(2).Picture = Nothing

'Vaciamos toda la primera pagina.
lblNombre.Visible = False
lblLvl.Visible = False
lblRaza.Visible = False
lblGenero.Visible = False
lblClase.Visible = False
lblHogar.Visible = False
lblMail.Visible = False
lblPuntosTorneo.Visible = False
lblPuntosDonador.Visible = False
lblBonificadores(1).Visible = False
lblBonificadores(2).Visible = False
lblBonificadores(3).Visible = False
picInv.Visible = False
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = False
Next

'Mostramos primera hoja de la pagina "extra" y vaciamos "habilidades"
For i = 0 To 14
lblCounters(i).Visible = False
Next
For i = 15 To 25
lblCounters(i).Visible = True
Next
For i = 1 To NUMSKILLS
Skills(i).Visible = False
Next
imgHoja(1).Visible = True
imgHoja(2).Visible = True

'Vaciamos la pagina "quest"
txtQuestDescription.Visible = False
imgQuestVal1.Visible = False
imgQuestVal2.Visible = False
imgQuestAbandonar.Visible = False
imgQuestType.Visible = False
End If

End Sub
Private Sub imgQuestAbandonar_Click()
Call SendData("/NOQUEST")
End Sub
Private Sub lblCerrar_Click()
Unload Me
End Sub
Private Sub imgGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgGeneral.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Principal_I.jpg")
End Sub
Private Sub imgGeneral_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgGeneral.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Principal_A.jpg")
End Sub
Private Sub imgHabilidades_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgHabilidades.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Habilidades_I.jpg")
End Sub
Private Sub imgHabilidades_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgHabilidades.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Habilidades_A.jpg")
End Sub
Private Sub imgQuests_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgQuests.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Quest_I.jpg")
End Sub
Private Sub imgQuests_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgQuests.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Quest_A.jpg")
End Sub
Private Sub imgExtras_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgExtras.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Extras_I.jpg")
End Sub
Private Sub imgExtras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgExtras.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Extras_A.jpg")
End Sub
Private Sub imgQuestAbandonar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgQuestAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_2_AbandonarQuest_I.jpg")
End Sub
Private Sub imgQuestAbandonar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgQuestAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_2_AbandonarQuest_A.jpg")
End Sub
Private Sub imgHoja_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 1 Then imgHoja(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_4_Hoja1_I.jpg")
If Index = 2 Then imgHoja(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_4_Hoja2_I.jpg")
End Sub
Private Sub lblPuntosTorneo_Click()
Call SendData("CCANJE")
End Sub

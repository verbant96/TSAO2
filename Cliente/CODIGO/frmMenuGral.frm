VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMenuGral 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   -360
   ClientWidth     =   7875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMenuGral.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Quests_infoDesc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1425
      Left            =   3960
      MultiLine       =   -1  'True
      TabIndex        =   60
      Top             =   1080
      Width           =   3795
   End
   Begin VB.TextBox Quests_qDescription 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3990
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   57
      Top             =   3105
      Visible         =   0   'False
      Width           =   3720
   End
   Begin VB.ListBox Quests_lstQuest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   4965
      IntegralHeight  =   0   'False
      Left            =   90
      TabIndex        =   56
      Top             =   990
      Width           =   3795
   End
   Begin VB.Timer tCredits 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   0
   End
   Begin VB.PictureBox creditos_picPack 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   1440
      ScaleHeight     =   1935
      ScaleWidth      =   1950
      TabIndex        =   54
      Top             =   3360
      Width           =   1980
   End
   Begin VB.ListBox creditos_lstPacks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1980
      IntegralHeight  =   0   'False
      Left            =   285
      TabIndex        =   50
      Top             =   975
      Width           =   1920
   End
   Begin VB.TextBox lDescripcion 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   2445
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   48
      Top             =   3870
      Width           =   1620
   End
   Begin VB.TextBox lCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   105
      Left            =   3495
      TabIndex        =   43
      Text            =   "1"
      Top             =   1800
      Width           =   570
   End
   Begin VB.PictureBox picObj 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   3015
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   41
      Top             =   1620
      Width           =   465
   End
   Begin VB.ListBox ListaPremios 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   4500
      IntegralHeight  =   0   'False
      Left            =   135
      TabIndex        =   39
      Top             =   1215
      Width           =   2250
   End
   Begin VB.TextBox txtAddpoints 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   2580
      TabIndex        =   8
      Text            =   "0"
      Top             =   3390
      Width           =   720
   End
   Begin MSComctlLib.ListView GuildList 
      Height          =   4635
      Left            =   345
      TabIndex        =   0
      Top             =   1095
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Constantia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Faccion"
         Object.Width           =   4180
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nivel"
         Object.Width           =   4260
      EndProperty
   End
   Begin MSComctlLib.ListView lstSolicitudes 
      Height          =   1440
      Left            =   360
      TabIndex        =   13
      Top             =   4455
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   2540
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   2620
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nivel"
         Object.Width           =   2302
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clase"
         Object.Width           =   2751
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Raza"
         Object.Width           =   2672
      EndProperty
   End
   Begin MSComctlLib.ListView Members 
      Height          =   4485
      Left            =   360
      TabIndex        =   14
      Top             =   1380
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   7911
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   3122
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nivel"
         Object.Width           =   3122
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clase"
         Object.Width           =   3043
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Raza"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ListView lstGuildList 
      Height          =   3660
      Left            =   360
      TabIndex        =   15
      Top             =   2040
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   6456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   4180
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Faccion"
         Object.Width           =   4128
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nivel"
         Object.Width           =   4128
      EndProperty
   End
   Begin MSComctlLib.ListView GuildUser_lstMembers 
      Height          =   1425
      Left            =   3930
      TabIndex        =   37
      Top             =   4470
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   2514
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   2064
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ubicacion"
         Object.Width           =   2011
      EndProperty
   End
   Begin MSComctlLib.ListView GuildUser_lstGuildList 
      Height          =   1425
      Left            =   360
      TabIndex        =   38
      Top             =   4470
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   2514
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   2064
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Faccion"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nivel"
         Object.Width           =   2011
      EndProperty
   End
   Begin RichTextLib.RichTextBox creditos_Desc 
      Height          =   1635
      Left            =   4080
      TabIndex        =   53
      Top             =   1200
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   2884
      _Version        =   393217
      BackColor       =   -2147483647
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMenuGral.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgSlideBar_Estadisticas 
      Height          =   315
      Left            =   120
      Top             =   450
      Width           =   1080
   End
   Begin VB.Label Estadisticas_QuestCompletadas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   98
      Top             =   1890
      Width           =   1080
   End
   Begin VB.Label Estadisticas_Muertes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   97
      Top             =   1740
      Width           =   1080
   End
   Begin VB.Label Estadisticas_NPCsAsesinados 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   96
      Top             =   1605
      Width           =   1080
   End
   Begin VB.Label Estadisticas_ParejasGanadas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   95
      Top             =   1455
      Width           =   1080
   End
   Begin VB.Label Estadisticas_DuelosGanados 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   94
      Top             =   1305
      Width           =   1080
   End
   Begin VB.Label Estadisticas_EventosGanados 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   93
      Top             =   1155
      Width           =   1080
   End
   Begin VB.Label Estadisticas_TorneosParticipados 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   92
      Top             =   1005
      Width           =   1080
   End
   Begin VB.Label Estadisticas_lblBonificadores 
      BackStyle       =   0  'Transparent
      Caption         =   "No elegido"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   91
      Top             =   4380
      UseMnemonic     =   0   'False
      Width           =   3495
   End
   Begin VB.Label Estadisticas_lblBonificadores 
      BackStyle       =   0  'Transparent
      Caption         =   "No elegido"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   90
      Top             =   4230
      UseMnemonic     =   0   'False
      Width           =   3495
   End
   Begin VB.Label Estadisticas_lblBonificadores 
      BackStyle       =   0  'Transparent
      Caption         =   "No elegido"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   89
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Estadisticas_HordasMatados 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   88
      Top             =   3765
      Width           =   3000
   End
   Begin VB.Label Estadisticas_AlianzasMatados 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "303"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   960
      TabIndex        =   87
      Top             =   3615
      Width           =   3000
   End
   Begin VB.Label Estadisticas_RequeridosJerarq 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "129"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   86
      Top             =   3330
      Width           =   960
   End
   Begin VB.Label Estadisticas_Jerarquia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   85
      Top             =   3180
      Width           =   3000
   End
   Begin VB.Label Estadisticas_Faccion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   84
      Top             =   3030
      Width           =   3000
   End
   Begin VB.Label Estadisticas_lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   83
      Top             =   2640
      Width           =   2040
   End
   Begin VB.Label Estadisticas_lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   82
      Top             =   2505
      Width           =   2040
   End
   Begin VB.Label Estadisticas_lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   81
      Top             =   2340
      Width           =   2040
   End
   Begin VB.Label Estadisticas_lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   80
      Top             =   2190
      Width           =   2040
   End
   Begin VB.Label Estadisticas_lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   79
      Top             =   2055
      Width           =   2040
   End
   Begin VB.Label Estadisticas_Hogar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tanaris"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   78
      Top             =   1770
      Width           =   2160
   End
   Begin VB.Label Estadisticas_Genero 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hombre"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   77
      Top             =   1620
      Width           =   2160
   End
   Begin VB.Label Estadisticas_Raza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Humano"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   76
      Top             =   1485
      Width           =   2160
   End
   Begin VB.Label Estadisticas_Clase 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mago"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   75
      Top             =   1320
      Width           =   2160
   End
   Begin VB.Label Estadisticas_Reputacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   74
      Top             =   1155
      Width           =   2160
   End
   Begin VB.Label Estadisticas_Nivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "70"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   73
      Top             =   1005
      Width           =   2160
   End
   Begin VB.Image imgSlideBar_Duelos 
      Height          =   315
      Left            =   5700
      Top             =   450
      Width           =   1305
   End
   Begin VB.Image duelos_Ingresar 
      Height          =   240
      Index           =   3
      Left            =   5280
      Top             =   5280
      Width           =   1065
   End
   Begin VB.Label duelos_Jugador8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   375
      Left            =   6120
      TabIndex        =   72
      Top             =   4980
      Width           =   1335
   End
   Begin VB.Label duelos_Jugador7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   4095
      TabIndex        =   71
      Top             =   4980
      Width           =   1440
   End
   Begin VB.Image duelos_Ingresar 
      Height          =   240
      Index           =   2
      Left            =   1530
      Top             =   5280
      Width           =   1065
   End
   Begin VB.Label duelos_Jugador6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   2295
      TabIndex        =   70
      Top             =   4980
      Width           =   1440
   End
   Begin VB.Label duelos_Jugador5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   360
      TabIndex        =   69
      Top             =   4980
      Width           =   1455
   End
   Begin VB.Image duelos_Ingresar 
      Height          =   240
      Index           =   1
      Left            =   5280
      Top             =   3915
      Width           =   1065
   End
   Begin VB.Label duelos_Jugador4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   6120
      TabIndex        =   68
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label duelos_Jugador3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   4095
      TabIndex        =   67
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image duelos_Ingresar 
      Height          =   240
      Index           =   0
      Left            =   1530
      Top             =   3915
      Width           =   1065
   End
   Begin VB.Label duelos_Jugador2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   375
      Left            =   2295
      TabIndex        =   66
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label duelos_Jugador1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   360
      TabIndex        =   65
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label quests_Oro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5700
      TabIndex        =   64
      Top             =   5355
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label quests_Credits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5700
      TabIndex        =   63
      Top             =   5145
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Quests_ptsTS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5700
      TabIndex        =   62
      Top             =   4935
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label quests_ptsTorneo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5700
      TabIndex        =   61
      Top             =   4725
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image Quests_Abandonar 
      Height          =   210
      Left            =   3900
      Top             =   5760
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Label Quests_cursoRestantes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5700
      TabIndex        =   59
      Top             =   3930
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Quests_cursoRequiere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5700
      TabIndex        =   58
      Top             =   3705
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image Quests_Aceptar 
      Height          =   210
      Left            =   3900
      Top             =   2550
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Image imgSlideBar_Quests 
      Height          =   315
      Left            =   1230
      Top             =   450
      Width           =   600
   End
   Begin VB.Image creditos_lblPurchase 
      Height          =   195
      Left            =   240
      Top             =   5760
      Width           =   7380
   End
   Begin VB.Label creditos_lblContent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1785
      Left            =   3840
      TabIndex        =   55
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label creditos_lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   2520
      TabIndex        =   52
      Top             =   1245
      Width           =   855
   End
   Begin VB.Label creditos_lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2520
      TabIndex        =   51
      Top             =   2385
      Width           =   855
   End
   Begin VB.Image imgSlideBar_Creditos 
      Height          =   315
      Left            =   3720
      Top             =   450
      Width           =   690
   End
   Begin VB.Image imgSlideBar_Manual 
      Height          =   315
      Left            =   7020
      Top             =   450
      Width           =   690
   End
   Begin VB.Image Medallas_imgCanjear 
      Height          =   225
      Index           =   5
      Left            =   6030
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Image Medallas_imgCanjear 
      Height          =   225
      Index           =   4
      Left            =   4245
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Image Medallas_imgCanjear 
      Height          =   225
      Index           =   3
      Left            =   6030
      Top             =   2595
      Width           =   1575
   End
   Begin VB.Image Medallas_imgCanjear 
      Height          =   225
      Index           =   2
      Left            =   4245
      Top             =   2595
      Width           =   1575
   End
   Begin VB.Image Medallas_imgCanjear 
      Height          =   225
      Index           =   1
      Left            =   6030
      Top             =   1845
      Width           =   1575
   End
   Begin VB.Image Medallas_imgCanjear 
      Height          =   225
      Index           =   0
      Left            =   4245
      Top             =   1830
      Width           =   1575
   End
   Begin VB.Image Gemas_imgCanjear 
      Height          =   210
      Index           =   4
      Left            =   270
      Top             =   2595
      Width           =   1575
   End
   Begin VB.Image Gemas_imgCanjear 
      Height          =   210
      Index           =   3
      Left            =   2040
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image Gemas_imgCanjear 
      Height          =   210
      Index           =   2
      Left            =   270
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image Gemas_imgCanjear 
      Height          =   210
      Index           =   1
      Left            =   2040
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image Gemas_imgCanjear 
      Height          =   210
      Index           =   0
      Left            =   270
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image imgTSPoint 
      Height          =   600
      Index           =   11
      Left            =   5085
      Top             =   2475
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   600
      Index           =   10
      Left            =   4440
      Top             =   2475
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   600
      Index           =   9
      Left            =   7020
      Top             =   1875
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   600
      Index           =   8
      Left            =   6360
      Top             =   1875
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   600
      Index           =   7
      Left            =   5730
      Top             =   1875
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   600
      Index           =   6
      Left            =   5070
      Top             =   1875
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   600
      Index           =   5
      Left            =   4440
      Top             =   1875
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   675
      Index           =   4
      Left            =   7020
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   675
      Index           =   3
      Left            =   6360
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   675
      Index           =   2
      Left            =   5715
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   675
      Index           =   1
      Left            =   5070
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image imgTSPoint 
      Height          =   675
      Index           =   0
      Left            =   4440
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image Canjes_imgExtras 
      Height          =   225
      Left            =   5280
      Top             =   750
      Width           =   1215
   End
   Begin VB.Image Canjes_imgGM 
      Height          =   225
      Left            =   1200
      Top             =   750
      Width           =   1695
   End
   Begin VB.Image imgSlideBar_Canjes 
      Height          =   315
      Left            =   2520
      Top             =   450
      Width           =   1095
   End
   Begin VB.Label lblTSPoints 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   49
      Top             =   4050
      Width           =   855
   End
   Begin VB.Image bCanjear 
      Height          =   270
      Left            =   2400
      Top             =   5520
      Width           =   1725
   End
   Begin VB.Label lDM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   47
      Top             =   3495
      Width           =   1215
   End
   Begin VB.Label lAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   46
      Top             =   3105
      Width           =   1215
   End
   Begin VB.Label lDef 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   45
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label lAtaque 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   44
      Top             =   2325
      Width           =   1215
   End
   Begin VB.Label Requiere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2415
      TabIndex        =   42
      Top             =   1755
      Width           =   600
   End
   Begin VB.Label lPuntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   40
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Image GuildUser_Abandonar 
      Height          =   255
      Left            =   1890
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image GuildUser_VerDetalles 
      Height          =   255
      Left            =   360
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image GuildDetails_Solicitar 
      Height          =   225
      Left            =   600
      Top             =   5595
      Width           =   6735
   End
   Begin VB.Label GuildDetails_Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   36
      Top             =   1020
      Width           =   3375
   End
   Begin VB.Label GuildDetails_Descripcion 
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1680
      TabIndex        =   35
      Top             =   4440
      Width           =   5535
   End
   Begin VB.Label GuildDetails_Fundador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2310
      TabIndex        =   34
      Top             =   3885
      Width           =   1335
   End
   Begin VB.Label GuildDetails_Faccion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2310
      TabIndex        =   33
      Top             =   3585
      Width           =   1335
   End
   Begin VB.Label GuildDetails_Miembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2310
      TabIndex        =   32
      Top             =   3285
      Width           =   1335
   End
   Begin VB.Label GuildDetails_Reputacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2310
      TabIndex        =   31
      Top             =   2985
      Width           =   1335
   End
   Begin VB.Label GuildDetails_FechaCreacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2310
      TabIndex        =   30
      Top             =   2685
      Width           =   1335
   End
   Begin VB.Label GuildDetails_Nivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2310
      TabIndex        =   29
      Top             =   2370
      Width           =   1335
   End
   Begin VB.Label GuildDetails_subLider 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   28
      Top             =   2100
      Width           =   2175
   End
   Begin VB.Label GuildDetails_Lider 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   27
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label GuildDetails_Codex 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   26
      Top             =   3975
      Width           =   3135
   End
   Begin VB.Label GuildDetails_Codex 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   25
      Top             =   3660
      Width           =   3135
   End
   Begin VB.Label GuildDetails_Codex 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   24
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label GuildDetails_Codex 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   23
      Top             =   3075
      Width           =   3135
   End
   Begin VB.Label GuildDetails_Codex 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   22
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label GuildDetails_Codex 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   21
      Top             =   2475
      Width           =   3135
   End
   Begin VB.Label GuildDetails_Codex 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   20
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label GuildDetails_Codex 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   19
      Top             =   1860
      Width           =   3135
   End
   Begin VB.Image guildLeader_VerDetalles 
      Height          =   225
      Left            =   255
      Top             =   5730
      Width           =   7275
   End
   Begin VB.Label lblCastles 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6105
      TabIndex        =   18
      Top             =   1455
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblCVCWins 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1335
      TabIndex        =   17
      Top             =   1455
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblCVCLosses 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3660
      TabIndex        =   16
      Top             =   1455
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image GuildLeader_Extras 
      Height          =   225
      Left            =   5760
      Top             =   720
      Width           =   1515
   End
   Begin VB.Image GuildLeader_Admin 
      Height          =   225
      Left            =   3120
      Top             =   750
      Width           =   1515
   End
   Begin VB.Image GuildLeader_Main 
      Height          =   225
      Left            =   480
      Top             =   750
      Width           =   1635
   End
   Begin VB.Image GuildLeader_RechazarAll 
      Height          =   495
      Left            =   6240
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Image GuildLeader_RechazarSolic 
      Height          =   495
      Left            =   6240
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Image GuildLeader_AcceptSolic 
      Height          =   495
      Left            =   6240
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Image imgCastillo 
      Height          =   450
      Index           =   4
      Left            =   4440
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Image imgCastillo 
      Height          =   330
      Index           =   3
      Left            =   4440
      Top             =   3330
      Width           =   3135
   End
   Begin VB.Image imgCastillo 
      Height          =   330
      Index           =   2
      Left            =   4440
      Top             =   3030
      Width           =   3135
   End
   Begin VB.Image imgCastillo 
      Height          =   330
      Index           =   1
      Left            =   4440
      Top             =   2745
      Width           =   3135
   End
   Begin VB.Image imgCastillo 
      Height          =   345
      Index           =   0
      Left            =   4440
      Top             =   2415
      Width           =   3135
   End
   Begin VB.Label lblCastillos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ustach Wielu"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   12
      Top             =   2190
      Width           =   2475
   End
   Begin VB.Label lblCastillos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ustach Wielu"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   11
      Top             =   1890
      Width           =   2475
   End
   Begin VB.Label lblCastillos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ustach Wielu"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   10
      Top             =   1590
      Width           =   2475
   End
   Begin VB.Label lblCastillos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ustach Wielu"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   9
      Top             =   1290
      Width           =   2475
   End
   Begin VB.Image cmdAddPoints 
      Height          =   255
      Left            =   2550
      Top             =   3570
      Width           =   795
   End
   Begin VB.Label lblSubLideres 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CleroPowa y xxMagoAlianzaxx"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1260
      TabIndex        =   7
      Top             =   1590
      Width           =   2115
   End
   Begin VB.Label lblLider 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shay"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1230
      TabIndex        =   6
      Top             =   1290
      Width           =   2175
   End
   Begin VB.Label lblMiembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblMaxMiembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   3060
      Width           =   855
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   1860
      Width           =   855
   End
   Begin VB.Label lblRep 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-5.837"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2475
      Width           =   855
   End
   Begin VB.Label lblPuntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1.000"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   2175
      Width           =   855
   End
   Begin VB.Image guildList_VerDetalle 
      Height          =   225
      Left            =   300
      Top             =   5700
      Width           =   7275
   End
   Begin VB.Image imgSlideBar_Clanes 
      Height          =   315
      Left            =   1800
      Top             =   450
      Width           =   675
   End
   Begin VB.Image imgSlideBar_Config 
      Height          =   315
      Left            =   4515
      Top             =   450
      Width           =   1140
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Index           =   16
      Left            =   360
      Top             =   3540
      Width           =   7335
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   15
      Left            =   6990
      Top             =   2670
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   14
      Left            =   6045
      Top             =   2670
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   13
      Left            =   5070
      Top             =   2670
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   12
      Left            =   4110
      Top             =   2670
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   11
      Left            =   3165
      Top             =   2670
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   10
      Left            =   2205
      Top             =   2670
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   9
      Left            =   1245
      Top             =   2670
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   8
      Left            =   270
      Top             =   2670
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   7
      Left            =   6990
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   6
      Left            =   6045
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   5
      Left            =   5070
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   4
      Left            =   4110
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   3
      Left            =   3165
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   2
      Left            =   2205
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   1
      Left            =   1245
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image imgCursor 
      Height          =   615
      Index           =   0
      Left            =   270
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image imgConfCursor 
      Height          =   615
      Left            =   5040
      Top             =   5190
      Width           =   2160
   End
   Begin VB.Image imgConfMac 
      Height          =   615
      Left            =   2880
      Top             =   5190
      Width           =   2160
   End
   Begin VB.Image imgConfTec 
      Height          =   615
      Left            =   720
      Top             =   5190
      Width           =   2160
   End
   Begin VB.Image imgPrivados 
      Height          =   180
      Left            =   5610
      Top             =   4455
      Width           =   180
   End
   Begin VB.Image imgDesplegarMenu 
      Height          =   180
      Left            =   5610
      Top             =   4785
      Width           =   180
   End
   Begin VB.Image imgGlobales 
      Height          =   180
      Left            =   5610
      Top             =   4125
      Width           =   180
   End
   Begin VB.Image imgDobleClick 
      Height          =   180
      Left            =   5610
      Top             =   3795
      Width           =   180
   End
   Begin VB.Image imgNotifyFriend 
      Height          =   180
      Left            =   5610
      Top             =   3465
      Width           =   180
   End
   Begin VB.Image imgMovVentana 
      Height          =   180
      Left            =   5610
      Top             =   3135
      Width           =   180
   End
   Begin VB.Image imgSound 
      Height          =   180
      Left            =   4530
      Top             =   2205
      Width           =   180
   End
   Begin VB.Image imgMusic 
      Height          =   180
      Left            =   1110
      Top             =   2205
      Width           =   180
   End
   Begin VB.Image imgGraphics 
      Height          =   180
      Index           =   2
      Left            =   5850
      Top             =   1275
      Width           =   180
   End
   Begin VB.Image imgGraphics 
      Height          =   180
      Index           =   1
      Left            =   3600
      Top             =   1275
      Width           =   180
   End
   Begin VB.Image imgGraphics 
      Height          =   180
      Index           =   0
      Left            =   1455
      Top             =   1275
      Width           =   180
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   7410
      Top             =   45
      Width           =   375
   End
   Begin VB.Menu mnuLider 
      Caption         =   "Lider"
      Visible         =   0   'False
      Begin VB.Menu mnuExpulsar 
         Caption         =   "Expulsar"
      End
      Begin VB.Menu mnuBoveda 
         Caption         =   "Permisos de Boveda"
         Begin VB.Menu mnuDep 
            Caption         =   "Solo depositar"
         End
         Begin VB.Menu mnuObjs 
            Caption         =   "Permitir retirar objetos"
         End
         Begin VB.Menu mnuGld 
            Caption         =   "Permitir retirar oro"
         End
         Begin VB.Menu mnuFull 
            Caption         =   "Permitir retirar objetos y oro"
         End
      End
      Begin VB.Menu mnus 
         Caption         =   "Sub Liderazgo"
         Begin VB.Menu mnusubhacer 
            Caption         =   "Hacer Sub Lider"
         End
         Begin VB.Menu mnusubsacar 
            Caption         =   "Sacar Sub Lider"
         End
      End
      Begin VB.Menu mnudolider 
         Caption         =   "Pasar Liderazgo"
      End
      Begin VB.Menu mnucontacts 
         Caption         =   "Agregar a contactos"
      End
      Begin VB.Menu mnuMensaje 
         Caption         =   "Enviar un mensaje"
      End
   End
End
Attribute VB_Name = "frmMenuGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Private MusicChanged As Boolean
Public Sub LoadOptions()

    Dim l_file As clsIniReader
    Set l_file = New clsIniReader

    '@ load file
    l_file.Initialize App.Path & "\Data\INIT\UserConfig.ini"
    
    Configuracion.Music = l_file.GetValue("OPTIONS", "Config_Music")
    Configuracion.Sound = l_file.GetValue("OPTIONS", "Config_Sound")
    Configuracion.Graphics = l_file.GetValue("OPTIONS", "Config_Graphics")
    Configuracion.Cursor = l_file.GetValue("OPTIONS", "Config_Cursor")
    Configuracion.Mensajes = l_file.GetValue("OPTIONS", "Config_Mensajes")
    
    Configuracion.DobleClick = l_file.GetValue("OPTIONS", "Config_DobleClick")
    Configuracion.Desactivar_Globales = l_file.GetValue("OPTIONS", "Config_Globales")
    Configuracion.Desactivar_Privados = l_file.GetValue("OPTIONS", "Config_Privados")
    Configuracion.MoverPantalla = l_file.GetValue("OPTIONS", "Config_ModoVentana")
    Configuracion.AnunciarContacto = l_file.GetValue("OPTIONS", "Config_Contactos")
    Configuracion.HablaNumerico = l_file.GetValue("OPTIONS", "Config_HablaNumerico")
    Configuracion.MenuDesplegable = l_file.GetValue("OPTIONS", "Config_MenuDesplegable")
    
    Configuracion.recordarCuenta = l_file.GetValue("OPTIONS", "RECORDAR_CUENTA")
    Configuracion.tmpCuenta = l_file.GetValue("OPTIONS", "TMPCUENTA")
    Configuracion.tmpPassword = l_file.GetValue("OPTIONS", "TMPPASSWORD")
End Sub
Private Sub SaveOptions()

    Dim l_file As clsIniReader
    Set l_file = New clsIniReader
    
    '@ load file
    l_file.Initialize App.Path & "\Data\INIT\UserConfig.ini"
    
    l_file.ChangeValue "OPTIONS", "Config_Music", Configuracion.Music
    l_file.ChangeValue "OPTIONS", "Config_Sound", Configuracion.Sound
    l_file.ChangeValue "OPTIONS", "Config_Graphics", Configuracion.Graphics
    
    l_file.ChangeValue "OPTIONS", "Config_Cursor", Configuracion.Cursor
    
    l_file.ChangeValue "OPTIONS", "Config_DobleClick", Configuracion.DobleClick
    l_file.ChangeValue "OPTIONS", "Config_Mensajes", Configuracion.Mensajes
    l_file.ChangeValue "OPTIONS", "Config_Globales", Configuracion.Desactivar_Globales
    l_file.ChangeValue "OPTIONS", "Config_Privados", Configuracion.Desactivar_Privados
    l_file.ChangeValue "OPTIONS", "Config_ModoVentana", Configuracion.MoverPantalla
    l_file.ChangeValue "OPTIONS", "Config_Contactos", Configuracion.AnunciarContacto
    
    l_file.ChangeValue "OPTIONS", "Config_HablaNumerico", Configuracion.HablaNumerico
    l_file.ChangeValue "OPTIONS", "Config_MenuDesplegable", Configuracion.MenuDesplegable
    
    Sound = Configuracion.Sound
    Musica = Configuracion.Music
    
    If Configuracion.Music = 0 And MusicChanged = True Then
        Audio.MP3_Stop
        Audio.MP3_Destroy
    End If
    
    If Configuracion.Sound = 0 Then
        Audio.StopMidi
        Audio.StopWave
    End If
    
    l_file.DumpFile App.Path & "\Data\INIT\UserConfig.ini"

End Sub
Private Sub mostrarOpciones(ByVal bool As Boolean)
    
    
    If bool Then
        resetAllForms
        
        For i = 0 To 2
            Call AplicarRadio(i)
        Next i
        
        Call AplicarTick(Configuracion.Music, imgMusic)
        Call AplicarTick(Configuracion.Sound, imgSound)
        Call AplicarTick(Configuracion.MoverPantalla, imgMovVentana)
        Call AplicarTick(Configuracion.AnunciarContacto, imgNotifyFriend)
        Call AplicarTick(Configuracion.DobleClick, imgDobleClick)
        Call AplicarTick(Configuracion.Desactivar_Globales, imgGlobales)
        Call AplicarTick(Configuracion.Desactivar_Privados, imgPrivados)
        Call AplicarTick(Configuracion.MenuDesplegable, imgDesplegarMenu)
        
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opciones_Main.jpg")
        MusicChanged = False
    End If
    
    For i = 0 To 2
        imgGraphics(i).Visible = bool
    Next i
    
    imgMusic.Visible = bool
    imgSound.Visible = bool
    imgMovVentana.Visible = bool
    imgNotifyFriend.Visible = bool
    imgDobleClick.Visible = bool
    imgGlobales.Visible = bool
    imgPrivados.Visible = bool
    imgDesplegarMenu.Visible = bool
    imgConfTec.Visible = bool
    imgConfMac.Visible = bool
    imgConfCursor.Visible = bool

End Sub
Private Sub mostrarCursores(ByVal bool As Boolean)
    
    If bool Then
        resetAllForms
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Cursores_Main.jpg")
        
        If (Configuracion.Cursor = 16) Then
            Me.MousePointer = 2
        Else
            Me.MousePointer = 99
            Me.MouseIcon = LoadPicture(App.Path & "\Data\GRAFICOS\Cursores\Cursor" & Configuracion.Cursor & ".ico")
        End If
        
    Else
        Me.MousePointer = vbDefault
    End If
    
    For i = 0 To 16
        imgCursor(i).Visible = bool
    Next i

End Sub
Private Sub mostrarDuelos(ByVal bool As Boolean)

    If bool Then
        resetAllForms
        
        duelos_Jugador1.ForeColor = RGB(145, 123, 85)
        duelos_Jugador2.ForeColor = RGB(145, 123, 85)
        duelos_Jugador3.ForeColor = RGB(145, 123, 85)
        duelos_Jugador4.ForeColor = RGB(145, 123, 85)
        duelos_Jugador5.ForeColor = RGB(145, 123, 85)
        duelos_Jugador6.ForeColor = RGB(145, 123, 85)
        duelos_Jugador7.ForeColor = RGB(145, 123, 85)
        duelos_Jugador8.ForeColor = RGB(145, 123, 85)
        
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Duelos_Main.jpg")
    End If
    
    duelos_Jugador1.Visible = bool
    duelos_Jugador2.Visible = bool
    duelos_Jugador3.Visible = bool
    duelos_Jugador4.Visible = bool
    duelos_Jugador5.Visible = bool
    duelos_Jugador6.Visible = bool
    duelos_Jugador7.Visible = bool
    duelos_Jugador8.Visible = bool
    
    For i = 0 To 3
        duelos_Ingresar(i).Visible = bool
    Next i
End Sub

Private Sub mostrarCanjes(ByVal bool As Boolean)

    If bool Then
        resetAllForms
        
        ListaPremios.BackColor = RGB(22, 23, 25)
        picObj.BackColor = RGB(22, 23, 25)
        lCantidad.BackColor = RGB(22, 23, 25)
        lDescripcion.BackColor = RGB(22, 23, 25)
        
        ListaPremios.ForeColor = RGB(145, 123, 85)
        lCantidad.ForeColor = RGB(145, 123, 85)
        lDescripcion.ForeColor = RGB(145, 123, 85)
        lPuntos.ForeColor = RGB(145, 123, 85)
        lAtaque.ForeColor = RGB(145, 123, 85)
        lDef.ForeColor = RGB(145, 123, 85)
        lAM.ForeColor = RGB(145, 123, 85)
        lDM.ForeColor = RGB(145, 123, 85)
        
        Call SendData("IPX1")
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Canjes_Main.jpg")
    End If
    
    ListaPremios.Visible = bool
    lCantidad.Visible = bool
    lPuntos.Visible = bool
    Requiere.Visible = bool
    picObj.Visible = bool
    lAtaque.Visible = bool
    lDef.Visible = bool
    lAM.Visible = bool
    lDM.Visible = bool
    lDescripcion.Visible = bool
    bCanjear.Visible = bool
    
    For i = 0 To 11
        imgTSPoint(i).Visible = bool
    Next i
    
    lblTSPoints.Visible = bool
    
    Canjes_imgExtras.Visible = bool
    Canjes_imgGM.Visible = bool
    
End Sub
Public Sub mostrarGem_Med(ByVal bool As Boolean)
    
    If bool Then
        resetAllForms
        
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\IntercambioObjetos_Main.jpg")
    End If
    
    For i = 0 To 5
        If i < 5 Then Gemas_imgCanjear(i).Visible = bool
        Medallas_imgCanjear(i).Visible = bool
    Next i
    
    Canjes_imgExtras.Visible = bool
    Canjes_imgGM.Visible = bool
End Sub
Private Sub mostrarQuests(ByVal bool As Boolean)

    If bool Then
        resetAllForms
        
        Quests_qDescription.BackColor = RGB(22, 23, 25)
        Quests_lstQuest.BackColor = RGB(22, 23, 25)
        Quests_infoDesc.BackColor = RGB(22, 23, 25)
        
        Quests_qDescription.ForeColor = RGB(145, 123, 85)
        Quests_lstQuest.ForeColor = RGB(145, 123, 85)
        Quests_infoDesc.ForeColor = RGB(145, 123, 85)
        Quests_cursoRequiere.ForeColor = RGB(145, 123, 85)
        Quests_cursoRestantes.ForeColor = RGB(145, 123, 85)
        quests_Credits.ForeColor = RGB(145, 123, 85)
        quests_Oro.ForeColor = RGB(145, 123, 85)
        quests_ptsTorneo.ForeColor = RGB(145, 123, 85)
        Quests_ptsTS.ForeColor = RGB(145, 123, 85)
        
        Call SendData("INFD1")
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Quests_Main.jpg")
    End If

    Quests_Aceptar.Visible = bool
    Quests_Abandonar.Visible = bool
    Quests_infoDesc.Visible = bool
    Quests_lstQuest.Visible = bool
    Quests_qDescription.Visible = bool
    Quests_cursoRequiere.Visible = bool
    Quests_cursoRestantes.Visible = bool
    quests_Credits.Visible = bool
    quests_Oro.Visible = bool
    quests_ptsTorneo.Visible = bool
    Quests_ptsTS.Visible = bool

End Sub
Private Sub mostrarCreditos(ByVal bool As Boolean)

    If bool Then
        resetAllForms
        creditos_Desc.BackColor = RGB(22, 23, 25)
        creditos_lstPacks.BackColor = RGB(22, 23, 25)
    
        creditos_lstPacks.ForeColor = RGB(145, 123, 85)
        creditos_lblContent.ForeColor = RGB(145, 123, 85)
        
        Call SendData("DPX1")
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Creditos_Main.jpg")
    End If
    
    tCredits.Enabled = bool
    
    creditos_Desc.Visible = bool
    creditos_lstPacks.Visible = bool
    creditos_lblCredits.Visible = bool
    creditos_lblPrice.Visible = bool
    creditos_picPack.Visible = bool
    creditos_lblContent.Visible = bool
    creditos_lblPurchase.Visible = bool
        

End Sub
Private Sub mostrarGuildUser(ByVal bool As Boolean)

    If bool Then
        resetAllForms
        
        txtAddpoints.BackColor = RGB(20, 21, 23)
        GuildUser_lstMembers.BackColor = RGB(20, 21, 23)
        GuildUser_lstMembers.ForeColor = RGB(105, 100, 94)
        GuildUser_lstGuildList.BackColor = RGB(20, 21, 23)
        GuildUser_lstGuildList.ForeColor = RGB(105, 100, 94)
        
        GuildUser_lstMembers.ColumnHeaders(2).Alignment = lvwColumnCenter
        GuildUser_lstMembers.ColumnHeaders(3).Alignment = lvwColumnCenter
        
        GuildUser_lstGuildList.ColumnHeaders(2).Alignment = lvwColumnCenter
        GuildUser_lstGuildList.ColumnHeaders(3).Alignment = lvwColumnCenter
        
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildUser_Main.jpg")
    End If
    
    lblLider.Visible = bool
    lblSubLideres.Visible = bool
    lblNivel.Visible = bool
    lblPuntos.Visible = bool
    lblRep.Visible = bool
    lblMiembros.Visible = bool
    lblMaxMiembros.Visible = bool
    txtAddpoints.Visible = bool
    cmdAddPoints.Visible = bool
    
    GuildUser_Abandonar.Visible = bool
    GuildUser_lstGuildList.Visible = bool
    GuildUser_lstMembers.Visible = bool
    GuildUser_VerDetalles.Visible = bool
    
    For i = 0 To 4
        If i < 4 Then lblCastillos(i).Visible = bool
        imgCastillo(i).Visible = bool
    Next i
    
End Sub
Private Sub mostrarEstadisticas(ByVal bool As Boolean)

    If bool Then
        resetAllForms
        
        
        Estadisticas_AlianzasMatados.ForeColor = RGB(145, 123, 85)
        Estadisticas_Clase.ForeColor = RGB(145, 123, 85)
        Estadisticas_DuelosGanados.ForeColor = RGB(145, 123, 85)
        Estadisticas_EventosGanados.ForeColor = RGB(145, 123, 85)
        Estadisticas_Faccion.ForeColor = RGB(145, 123, 85)
        Estadisticas_Genero.ForeColor = RGB(145, 123, 85)
        Estadisticas_Hogar.ForeColor = RGB(145, 123, 85)
        Estadisticas_HordasMatados.ForeColor = RGB(145, 123, 85)
        Estadisticas_Jerarquia.ForeColor = RGB(145, 123, 85)
        Estadisticas_lblBonificadores(1).ForeColor = RGB(145, 123, 85)
        Estadisticas_lblBonificadores(2).ForeColor = RGB(145, 123, 85)
        Estadisticas_lblBonificadores(3).ForeColor = RGB(145, 123, 85)
        Estadisticas_Muertes.ForeColor = RGB(145, 123, 85)
        Estadisticas_Nivel.ForeColor = RGB(145, 123, 85)
        Estadisticas_NPCsAsesinados.ForeColor = RGB(145, 123, 85)
        Estadisticas_ParejasGanadas.ForeColor = RGB(145, 123, 85)
        Estadisticas_QuestCompletadas.ForeColor = RGB(145, 123, 85)
        Estadisticas_Raza.ForeColor = RGB(145, 123, 85)
        Estadisticas_Reputacion.ForeColor = RGB(145, 123, 85)
        Estadisticas_RequeridosJerarq.ForeColor = RGB(145, 123, 85)
        Estadisticas_TorneosParticipados.ForeColor = RGB(145, 123, 85)
        
        
        For i = 0 To 4
            Estadisticas_lblAtri(i).ForeColor = RGB(145, 123, 85)
        Next i
        
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_Main.jpg")
    End If

    Estadisticas_AlianzasMatados.Visible = bool
    Estadisticas_Clase.Visible = bool
    Estadisticas_DuelosGanados.Visible = bool
    Estadisticas_EventosGanados.Visible = bool
    Estadisticas_Faccion.Visible = bool
    Estadisticas_Genero.Visible = bool
    Estadisticas_Hogar.Visible = bool
    Estadisticas_HordasMatados.Visible = bool
    Estadisticas_Jerarquia.Visible = bool
    Estadisticas_lblBonificadores(1).Visible = bool
    Estadisticas_lblBonificadores(2).Visible = bool
    Estadisticas_lblBonificadores(3).Visible = bool
    Estadisticas_Muertes.Visible = bool
    Estadisticas_Nivel.Visible = bool
    Estadisticas_NPCsAsesinados.Visible = bool
    Estadisticas_ParejasGanadas.Visible = bool
    Estadisticas_QuestCompletadas.Visible = bool
    Estadisticas_Raza.Visible = bool
    Estadisticas_Reputacion.Visible = bool
    Estadisticas_RequeridosJerarq.Visible = bool
    Estadisticas_TorneosParticipados.Visible = bool
    
    
    For i = 0 To 4
        Estadisticas_lblAtri(i).Visible = bool
    Next i
End Sub
Private Sub mostrarGuildDetails(ByVal bool As Boolean)

    If bool Then
        resetAllForms
        
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Detalles.jpg")
    End If
    
    GuildDetails_Nivel.Visible = bool
    GuildDetails_FechaCreacion.Visible = bool
    GuildDetails_Reputacion.Visible = bool
    GuildDetails_Miembros.Visible = bool
    GuildDetails_Faccion.Visible = bool
    GuildDetails_Fundador.Visible = bool
    GuildDetails_Lider.Visible = bool
    GuildDetails_subLider.Visible = bool
    
    For i = 0 To 7
        GuildDetails_Codex(i).Visible = bool
    Next i
    
    GuildDetails_Descripcion.Visible = bool
    GuildDetails_Nombre.Visible = bool
    GuildDetails_Solicitar.Visible = bool
    
End Sub
Private Sub mostrarGuildList(ByVal bool As Boolean)
    
    If bool Then
        resetAllForms
    
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildList_Main.jpg")
        
        GuildList.ColumnHeaders(2).Alignment = lvwColumnCenter
        GuildList.ColumnHeaders(3).Alignment = lvwColumnCenter
        
        GuildList.BackColor = RGB(20, 21, 23)
        GuildList.ForeColor = RGB(105, 100, 94)
    End If
    
    GuildList.Visible = bool
    guildList_VerDetalle.Visible = bool
        

End Sub
Private Sub mostrarExtrasLeader(ByVal bool As Boolean)

    If bool Then
        resetAllForms
        
        lstGuildList.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstGuildList.ColumnHeaders(3).Alignment = lvwColumnCenter
        
        lstGuildList.ForeColor = RGB(105, 100, 94)
        lstGuildList.BackColor = RGB(20, 21, 23)
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildLeader_Extras.jpg")
    End If
    
    guildLeader_VerDetalles.Visible = bool
    lblCastles.Visible = bool
    lblCVCWins.Visible = bool
    lblCVCLosses.Visible = bool
    lstGuildList.Visible = bool
    GuildLeader_Main.Visible = bool
    GuildLeader_Admin.Visible = bool
    GuildLeader_Extras.Visible = bool
End Sub
    
Private Sub mostrarAdminLeader(ByVal bool As Boolean)
    If bool Then
        resetAllForms
        
        Members.ColumnHeaders(2).Alignment = lvwColumnCenter
        Members.ColumnHeaders(3).Alignment = lvwColumnCenter
        Members.ColumnHeaders(4).Alignment = lvwColumnCenter
        
        Members.ForeColor = RGB(105, 100, 94)
        Members.BackColor = RGB(20, 21, 23)
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildLeader_Administracion.jpg")
    End If
    
    Members.Visible = bool
    GuildLeader_Main.Visible = bool
    GuildLeader_Admin.Visible = bool
    GuildLeader_Extras.Visible = bool
    
    
End Sub
Private Sub mostrarMainLeader(ByVal bool As Boolean)

    If bool Then
        resetAllForms
        
        lstSolicitudes.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstSolicitudes.ColumnHeaders(3).Alignment = lvwColumnCenter
        lstSolicitudes.ColumnHeaders(4).Alignment = lvwColumnCenter
        
        lstSolicitudes.ForeColor = RGB(105, 100, 94)
        lstSolicitudes.BackColor = RGB(20, 21, 23)
        txtAddpoints.BackColor = RGB(20, 21, 23)
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildLeader_Main.jpg")
    End If

    lblLider.Visible = bool
    lblSubLideres.Visible = bool
    lblNivel.Visible = bool
    lblPuntos.Visible = bool
    lblRep.Visible = bool
    lblMiembros.Visible = bool
    lblMaxMiembros.Visible = bool
    txtAddpoints.Visible = bool
    cmdAddPoints.Visible = bool
    lstSolicitudes.Visible = bool
    GuildLeader_AcceptSolic.Visible = bool
    GuildLeader_RechazarSolic.Visible = bool
    GuildLeader_RechazarAll.Visible = bool
    GuildLeader_Main.Visible = bool
    GuildLeader_Admin.Visible = bool
    GuildLeader_Extras.Visible = bool
    
    For i = 0 To 4
        If i < 4 Then lblCastillos(i).Visible = bool
        imgCastillo(i).Visible = bool
    Next i
    

End Sub
Private Sub resetAllForms()

    mostrarOpciones (False)
    mostrarCursores (False)
    mostrarGuildList (False)
    mostrarMainLeader (False)
    mostrarAdminLeader (False)
    mostrarExtrasLeader (False)
    mostrarGuildDetails (False)
    mostrarGuildUser (False)
    mostrarCanjes (False)
    mostrarGem_Med (False)
    mostrarCreditos (False)
    mostrarQuests (False)
    mostrarDuelos (False)
    mostrarEstadisticas (False)
    
End Sub

Private Sub duelos_Ingresar_Click(Index As Integer)
    Call SendData("ARE" & Index + 1)
End Sub

Private Sub imgSlideBar_Duelos_Click()
    Call SendData("IDUELOS")
End Sub

Private Sub imgSlideBar_Estadisticas_Click()
    Call SendData("FEST")
End Sub
Private Sub imgSlideBar_Quests_Click()
    Call SendData("IQUEST")
End Sub
Private Sub Quests_Abandonar_Click()
    Call SendData("/NOQUEST")
End Sub
Private Sub Quests_Aceptar_Click()
    Call SendData("ACQT" & Quests_lstQuest.ListIndex + 1)
End Sub
Private Sub Quests_lstQuest_Click()
    Call SendData("INFD" & Quests_lstQuest.ListIndex + 1)
End Sub
Private Sub Canjes_imgGM_Click()
    mostrarGem_Med (True)
End Sub
Private Sub cmdAddPoints_Click()
    Dim cantPuntos As Long
    cantPuntos = Val(txtAddpoints.text)

    If Not IsNumeric(cantPuntos) Then Exit Sub
    If cantPuntos = 0 Then Exit Sub

    Call SendData("ADDPTS" & cantPuntos)
    Call SendData("GLINFO")
End Sub

Private Sub creditos_lstPacks_Click()
    Call SendData("DPX" & creditos_lstPacks.ListIndex + 1)
End Sub
Private Sub Form_Load()
    Set form_Moviment = New clsFormMovementManager
    form_Moviment.Initialize Me
    
    mostrarOpciones (True)
    
End Sub
Private Sub GuildDetails_Solicitar_Click()
    Dim f$

    f$ = "SOLICITUD" & GuildDetails_Nombre
    f$ = f$ & "," & Replace(Replace("", ",", ";"), vbCrLf, "º")
    
    Call SendData(f$)
    
    Unload Me
End Sub
Private Sub Gemas_imgCanjear_DblClick(Index As Integer)
    If MsgBox("¿Estás seguro que quieres canjear este objeto?", vbYesNo, "Tierras Sagradas AO") = vbYes Then
        Call SendData("GEMS" & Index)
        Unload Me
    End If
End Sub

Private Sub imgSlideBar_Creditos_Click()
    Call SendData("DCANJE")
End Sub

Private Sub imgSlideBar_Manual_Click()
    Mensaje.Escribir "Muy pronto habilitaremos la sección Manual con todas las guías del juego."
End Sub

Private Sub Medallas_imgCanjear_Click(Index As Integer)
    If MsgBox("¿Estás seguro que quieres canjear este objeto?", vbYesNo, "Tierras Sagradas AO") = vbYes Then
        Call SendData("GEPS" & Index)
        Unload Me
    End If
End Sub
Private Sub creditos_lblPurchase_Click()
    If MsgBox("¿Estás seguro que deseas canjear " & creditos_lstPacks.text & "?", vbYesNo) = vbYes Then
        Call SendData("DRX" & creditos_lstPacks.ListIndex + 1)
    End If
End Sub
Private Sub GuildLeader_AcceptSolic_Click()
    If lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).text <> "" Then
        Call SendData("ACEPTARI" & lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).text)
        Call SendData("GLINFO")
    End If
End Sub

Private Sub GuildLeader_Admin_Click()
    mostrarAdminLeader (True)
End Sub

Private Sub GuildLeader_Extras_Click()
    mostrarExtrasLeader (True)
End Sub
Private Sub GuildLeader_Main_Click()
    mostrarMainLeader (True)
End Sub
Private Sub GuildLeader_RechazarAll_Click()
    If lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).text <> "" Then
        Call SendData("RECHAZAR" & lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).text & "," & Replace(Replace("asd", ",", " "), vbCrLf, " "))
        lstSolicitudes.ListItems.Remove lstSolicitudes.SelectedItem.Index
            
        Call SendData("GLINFO")
    End If
End Sub
Private Sub GuildLeader_RechazarSolic_Click()
    If lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).text <> "" Then
        Call SendData("RECHAZAR" & lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).text & "," & Replace(Replace("asd", ",", " "), vbCrLf, " "))
        lstSolicitudes.ListItems.Remove lstSolicitudes.SelectedItem.Index
            
        Call SendData("GLINFO")
    End If
End Sub

Private Sub guildLeader_VerDetalles_Click()
    If lstGuildList.SelectedItem.Index <= 0 Then Exit Sub
    If lstGuildList.ListItems.Item(lstGuildList.SelectedItem.Index).text = "" Then Exit Sub
    
    Call SendData("CLANDETAILS" & lstGuildList.ListItems.Item(lstGuildList.SelectedItem.Index).text)
End Sub

Private Sub guildList_VerDetalle_Click()
    If GuildList.SelectedItem.Index <= 0 Then Exit Sub
    If GuildList.ListItems.Item(GuildList.SelectedItem.Index).text = "" Then Exit Sub
    
    Call SendData("CLANDETAILS" & GuildList.ListItems.Item(GuildList.SelectedItem.Index).text)
End Sub

Private Sub GuildUser_Abandonar_Click()
    Call SendData("/SALIRCLAN")
    Unload Me
End Sub
Private Sub GuildUser_VerDetalles_Click()
     If GuildUser_lstGuildList.SelectedItem.Index <= 0 Then Exit Sub
     Call SendData("CLANDETAILS" & GuildUser_lstGuildList.ListItems.Item(GuildUser_lstGuildList.SelectedItem.Index).text)
End Sub
Private Sub bCanjear_Click()
    If MsgBox("¿Estás seguro que deseas canjear " & lCantidad.text & " - " & ListaPremios.text & "?", vbYesNo) = vbYes Then
        Call SendData("SPX" & ListaPremios.ListIndex + 1 & "," & lCantidad.text)
    End If
End Sub
Private Sub Canjes_imgExtras_Click()
    Call SendData("CCANJE")
End Sub
Private Sub imgSlideBar_Canjes_Click()
    Call SendData("CCANJE")
End Sub

Private Sub imgTSPoint_DblClick(Index As Integer)
    If MsgBox("¿Estás seguro que quieres canjear este objeto?", vbYesNo, "Tierras Sagradas AO") = vbYes Then
        Call SendData("FTSPTS" & Index)
    End If
End Sub
Private Sub lCantidad_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
 Case Asc("0") To Asc("9"), vbKeyDelete, vbKeyBack
Case Else: KeyAscii = 0
 End Select
End Sub
Private Sub lCantidad_Change()

If lCantidad.text = "" Then lCantidad.text = 1
If Not IsNumeric(lCantidad.text) Or lCantidad.text < 0 Or lCantidad.text > 10000 Then lCantidad.text = 1
Requiere.Caption = CantidadCanjeYegua * lCantidad.text

End Sub
Private Sub ListaPremios_Click()
    Call SendData("IPX" & ListaPremios.ListIndex + 1)
End Sub
Private Sub imgCastillo_Click(Index As Integer)
    Select Case Index
        Case 0
            Call SendData("/IR 33")
        Case 1
            Call SendData("/IR 31")
        Case 2
            Call SendData("/IR 34")
        Case 3
            Call SendData("/IR 32")
        Case 4
            Call SendData("/IR 35")
    End Select
End Sub
Private Sub imgConfCursor_Click()
    mostrarCursores (True)
End Sub
Private Sub imgConfMac_Click()
    frmMakro.Show , frmMain
End Sub
Private Sub imgConfTec_Click()
    Call frmTeclas.Show(vbModeless, frmMain)
End Sub

Private Sub imgCursor_Click(Index As Integer)

On Error GoTo errorh:

    Configuracion.Cursor = Index
    
    If (Index = 16) Then
        Me.MousePointer = 2
    Else
        Me.MousePointer = 99
        Me.MouseIcon = LoadPicture(App.Path & "\Data\GRAFICOS\Cursores\Cursor" & Index & ".ico")
    End If
    
Exit Sub
errorh:
    Configuracion.Cursor = 16
    Me.MousePointer = 2
    Mensaje.Escribir ("Ocurrió un error, no puedes modificar el cursor de tu computadora.")
End Sub

Private Sub imgDesplegarMenu_Click()
    If Configuracion.MenuDesplegable = 1 Then
        Configuracion.MenuDesplegable = 0
    Else
        Configuracion.MenuDesplegable = 1
    End If
    
    Call AplicarTick(Configuracion.MenuDesplegable, imgDesplegarMenu)
End Sub

Private Sub imgDobleClick_Click()
    If Configuracion.DobleClick = 1 Then
        Configuracion.DobleClick = 0
    Else
        Configuracion.DobleClick = 1
    End If
    
    Call AplicarTick(Configuracion.DobleClick, imgDobleClick)
End Sub

Private Sub imgGlobales_Click()
    If Configuracion.Desactivar_Globales = 1 Then
        Configuracion.Desactivar_Globales = 0
    Else
        Configuracion.Desactivar_Globales = 1
    End If
    
    Call AplicarTick(Configuracion.Desactivar_Globales, imgGlobales)
End Sub
Private Sub imgGraphics_Click(Index As Integer)
    For i = 0 To 2
        imgGraphics(i).Picture = Nothing
    Next i
    
    Configuracion.Graphics = Index
    imgGraphics(Index).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opciones_SelectBox.jpg")
End Sub
Private Sub AplicarRadio(ByVal Index As Byte)
    
    If Configuracion.Graphics = Index Then
        imgGraphics(Index).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opciones_SelectBox.jpg")
    Else
        imgGraphics(Index).Picture = Nothing
    End If

End Sub
Private Sub AplicarTick(ByVal Activate As Byte, aux As Image)
    
    If Activate = 1 Then
        aux.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opciones_CheckBox.jpg")
    Else
        aux.Picture = Nothing
    End If

End Sub
Private Sub imgMovVentana_Click()
    If Configuracion.MoverPantalla = 1 Then
        Configuracion.MoverPantalla = 0
    Else
        Configuracion.MoverPantalla = 1
    End If
    
    Call AplicarTick(Configuracion.MoverPantalla, imgMovVentana)
End Sub

Private Sub imgMusic_Click()
    If Configuracion.Music = 1 Then
        Configuracion.Music = 0
    Else
        Configuracion.Music = 1
    End If
    
    Call AplicarTick(Configuracion.Music, imgMusic)
    MusicChanged = True
End Sub

Private Sub imgNotifyFriend_Click()
    If Configuracion.AnunciarContacto = 1 Then
        Configuracion.AnunciarContacto = 0
    Else
        Configuracion.AnunciarContacto = 1
    End If
    
    Call AplicarTick(Configuracion.AnunciarContacto, imgNotifyFriend)
End Sub

Private Sub imgPrivados_Click()
    If Configuracion.Desactivar_Privados = 1 Then
        Configuracion.Desactivar_Privados = 0
    Else
        Configuracion.Desactivar_Privados = 1
    End If
    
    Call AplicarTick(Configuracion.Desactivar_Privados, imgPrivados)
End Sub
Private Sub imgSlideBar_Clanes_Click()
    Call SendData("GLINFO")
End Sub
Private Sub imgSlideBar_Config_Click()
    mostrarOpciones (True)
End Sub

Private Sub imgSound_Click()
    If Configuracion.Sound = 1 Then
        Configuracion.Sound = 0
    Else
        Configuracion.Sound = 1
    End If
    
    Call AplicarTick(Configuracion.Sound, imgSound)
End Sub
Private Sub imgSalir_Click()
    SaveOptions
    Unload Me
End Sub
Public Sub ParseGuildList(ByVal rData As String)

Dim j As Long, k As Integer
k = CInt(ReadField(1, rData, Asc(",")))

Dim ClanTemporal As String
Dim NombreClan As String
Dim FaccionClan As String
Dim NivelClan As Byte
Dim IndexK As Integer
GuildList.ListItems.Clear
ClanTemporal = ""
IndexK = 1

For j = 1 To k
    ClanTemporal = ReadField(j + 1, rData, Asc(","))
    NombreClan = ReadField(1, ClanTemporal, Asc("-"))
    
    If UCase$(NombreClan) <> UCase$("cerrado" & j & "") Then
        FaccionClan = ReadField(2, ClanTemporal, Asc("-"))
        NivelClan = ReadField(3, ClanTemporal, Asc("-"))
        
        GuildList.ListItems.Add IndexK, , NombreClan
        
        If FaccionClan = 3 Then
            GuildList.ListItems(IndexK).ListSubItems.Add , , "NEUTRAL"
        ElseIf FaccionClan = 4 Or FaccionClan = 5 Then
            GuildList.ListItems(IndexK).ListSubItems.Add , , "ALIANZA"
        ElseIf FaccionClan = 2 Or FaccionClan = 3 Then
            GuildList.ListItems(IndexK).ListSubItems.Add , , "HORDA"
        End If
        
        GuildList.ListItems(IndexK).ListSubItems.Add , , NivelClan
        IndexK = IndexK + 1
    End If
Next j

mostrarGuildList (True)

End Sub
Public Sub ParseLeaderInfo(ByVal Data As String)

Members.ListItems.Clear
lstSolicitudes.ListItems.Clear

lblPuntos.Caption = ReadField(1, Data, Asc("¬"))
lblNivel.Caption = ReadField(2, Data, Asc("¬"))
lblLider.Caption = ReadField(3, Data, Asc("¬"))

If ReadField(4, Data, Asc("¬")) <> "Fermin" And ReadField(5, Data, Asc("¬")) = "Fermin" Then
    lblSubLideres.Caption = ReadField(4, Data, Asc("¬"))
ElseIf ReadField(4, Data, Asc("¬")) = "Fermin" And ReadField(5, Data, Asc("¬")) <> "Fermin" Then
    lblSubLideres.Caption = ReadField(5, Data, Asc("¬"))
ElseIf ReadField(4, Data, Asc("¬")) <> "Fermin" And ReadField(5, Data, Asc("¬")) <> "Fermin" Then
    lblSubLideres.Caption = "" & ReadField(4, Data, Asc("¬")) & " y " & ReadField(5, Data, Asc("¬")) & ""
ElseIf ReadField(4, Data, Asc("¬")) = "Fermin" And ReadField(5, Data, Asc("¬")) = "Fermin" Then
    lblSubLideres.Caption = "-"
End If

lblCastillos(0).Caption = ReadField(6, Data, Asc("¬"))
lblCastillos(1).Caption = ReadField(7, Data, Asc("¬"))
lblCastillos(2).Caption = ReadField(8, Data, Asc("¬"))
lblCastillos(3).Caption = ReadField(9, Data, Asc("¬"))

lblRep.Caption = PonerPuntos(ReadField(10, Data, Asc("¬")))

lblCVCWins.Caption = PonerPuntos(ReadField(11, Data, Asc("¬")))
lblCVCLosses.Caption = PonerPuntos(ReadField(12, Data, Asc("¬")))
lblCastles.Caption = PonerPuntos(ReadField(13, Data, Asc("¬")))

Dim i, cantClanes, cantMiembros, cantSolicitudes As Long

cantClanes = Val(ReadField(14, Data, Asc("¬")))

Dim NombreClan, ClanTemporal, FaccionClan As String
Dim NivelClan, tmpSuma As Byte
Dim IndexK As Integer

lstGuildList.ListItems.Clear
ClanTemporal = ""
IndexK = 1

For i = 1 To cantClanes
    ClanTemporal = ReadField(14 + i, Data, Asc("¬"))
    NombreClan = ReadField(1, ClanTemporal, Asc("$"))
    
    If UCase$(NombreClan) <> UCase$("cerrado" & i & "") Then
        FaccionClan = ReadField(2, ClanTemporal, Asc("$"))
        NivelClan = ReadField(3, ClanTemporal, Asc("$"))
        
        lstGuildList.ListItems.Add IndexK, , NombreClan
        
        If FaccionClan = 3 Then
            lstGuildList.ListItems(IndexK).ListSubItems.Add , , "NEUTRAL"
        ElseIf FaccionClan = 4 Or FaccionClan = 5 Then
            lstGuildList.ListItems(IndexK).ListSubItems.Add , , "ALIANZA"
        ElseIf FaccionClan = 2 Or FaccionClan = 3 Then
            lstGuildList.ListItems(IndexK).ListSubItems.Add , , "HORDA"
        End If
        
        lstGuildList.ListItems(IndexK).ListSubItems.Add , , NivelClan
        IndexK = IndexK + 1
    End If
Next i

tmpSuma = 15 + cantClanes
cantMiembros = Val(ReadField(tmpSuma, Data, Asc("¬")))
lblMiembros.Caption = cantMiembros
lblMaxMiembros.Caption = 4 * Val(cantMiembros)

Dim MiembroTemporal As String
MiembroTemporal = ""

For i = 1 To cantMiembros
    MiembroTemporal = ReadField(tmpSuma + i, Data, Asc("¬"))

    Members.ListItems.Add , , ReadField(1, MiembroTemporal, Asc("$"))
    Members.ListItems(i).bold = True
    
    Members.ListItems(i).ListSubItems.Add , , ReadField(2, MiembroTemporal, Asc("$"))
    Members.ListItems(i).ListSubItems.Add , , ReadField(3, MiembroTemporal, Asc("$"))
    Members.ListItems(i).ListSubItems.Add , , ReadField(4, MiembroTemporal, Asc("$"))

Next i

tmpSuma = (16 + cantClanes) + cantMiembros
cantSolicitudes = Val(ReadField(tmpSuma, Data, Asc("¬")))

For i = 1 To cantSolicitudes
    MiembroTemporal = ReadField(tmpSuma + i, Data, Asc("¬"))
    
    lstSolicitudes.ListItems.Add , , ReadField(1, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(i).bold = True
    
    lstSolicitudes.ListItems(i).ListSubItems.Add , , ReadField(2, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(i).ListSubItems.Add , , ReadField(3, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(i).ListSubItems.Add , , ReadField(4, MiembroTemporal, Asc("$"))
Next i

mostrarMainLeader (True)

End Sub

Public Sub ParseGuildInfo(ByVal Buffer As String)

GuildDetails_Nivel.Caption = ReadField(1, Buffer, Asc("¬"))
GuildDetails_Faccion.Caption = ReadField(2, Buffer, Asc("¬"))
GuildDetails_Reputacion.Caption = PonerPuntos(ReadField(3, Buffer, Asc("¬")))
GuildDetails_Fundador.Caption = ReadField(4, Buffer, Asc("¬"))
GuildDetails_FechaCreacion.Caption = ReadField(5, Buffer, Asc("¬"))
GuildDetails_Lider.Caption = ReadField(6, Buffer, Asc("¬"))

If ReadField(2, Buffer, Asc("¬")) = 2 Then
    GuildDetails_Faccion.ForeColor = &H80&
    GuildDetails_Faccion = "HORDA INFERNAL"
ElseIf ReadField(2, Buffer, Asc("¬")) = 4 Then
    GuildDetails_Faccion.ForeColor = &HC00000
    GuildDetails_Faccion = "ALIANZA IMPERIAL"
ElseIf ReadField(2, Buffer, Asc("¬")) = 3 Then
    GuildDetails_Faccion.ForeColor = &H404040
    GuildDetails_Faccion = "NEUTRAL"
End If

If ReadField(7, Buffer, Asc("¬")) = "Fermin" And ReadField(8, Buffer, Asc("¬")) = "Fermin" Then
    GuildDetails_subLider.Caption = "-"
ElseIf ReadField(7, Buffer, Asc("¬")) <> "Fermin" And ReadField(8, Buffer, Asc("¬")) = "Fermin" Then
    GuildDetails_subLider.Caption = ReadField(7, Buffer, Asc("¬"))
ElseIf ReadField(7, Buffer, Asc("¬")) <> "Fermin" And ReadField(8, Buffer, Asc("¬")) <> "Fermin" Then
    GuildDetails_subLider.Caption = "" & ReadField(7, Buffer, Asc("¬")) & " y " & ReadField(8, Buffer, Asc("¬")) & ""
ElseIf ReadField(7, Buffer, Asc("¬")) = "Fermin" And ReadField(8, Buffer, Asc("¬")) <> "Fermin" Then
    GuildDetails_subLider.Caption = ReadField(8, Buffer, Asc("¬"))
End If

GuildDetails_Miembros.Caption = ReadField(9, Buffer, Asc("¬"))

Dim T As Long

For T = 0 To 7
    GuildDetails_Codex(T).Caption = ReadField(10 + T, Buffer, Asc("¬"))
Next T

Dim des As String

des = ReadField(18, Buffer, Asc("¬"))
GuildDetails_Nombre.Caption = ReadField(19, Buffer, Asc("¬"))
GuildDetails_Descripcion.Caption = Replace(des, "º", vbCrLf)
mostrarGuildDetails (True)

End Sub
Public Sub ParseGuildUserInfo(ByVal Data As String)

GuildUser_lstMembers.ListItems.Clear
GuildUser_lstGuildList.ListItems.Clear

lblPuntos.Caption = ReadField(1, Data, Asc("¬"))
lblNivel.Caption = ReadField(2, Data, Asc("¬"))
lblLider.Caption = ReadField(3, Data, Asc("¬"))

If ReadField(4, Data, Asc("¬")) <> "Fermin" And ReadField(5, Data, Asc("¬")) = "Fermin" Then
    lblSubLideres.Caption = ReadField(4, Data, Asc("¬"))
ElseIf ReadField(4, Data, Asc("¬")) = "Fermin" And ReadField(5, Data, Asc("¬")) <> "Fermin" Then
    lblSubLideres.Caption = ReadField(5, Data, Asc("¬"))
ElseIf ReadField(4, Data, Asc("¬")) <> "Fermin" And ReadField(5, Data, Asc("¬")) <> "Fermin" Then
    lblSubLideres.Caption = "" & ReadField(4, Data, Asc("¬")) & " y " & ReadField(5, Data, Asc("¬")) & ""
ElseIf ReadField(4, Data, Asc("¬")) = "Fermin" And ReadField(5, Data, Asc("¬")) = "Fermin" Then
    lblSubLideres.Caption = "-"
End If

lblCastillos(0).Caption = ReadField(6, Data, Asc("¬"))
lblCastillos(1).Caption = ReadField(7, Data, Asc("¬"))
lblCastillos(2).Caption = ReadField(8, Data, Asc("¬"))
lblCastillos(3).Caption = ReadField(9, Data, Asc("¬"))

lblRep.Caption = PonerPuntos(ReadField(10, Data, Asc("¬")))

Dim i, cantClanes, cantMiembros As Long

cantClanes = Val(ReadField(11, Data, Asc("¬")))

Dim NombreClan, ClanTemporal As String
Dim NivelClan, tmpSuma, FaccionClan As Byte
Dim IndexK As Integer

lstGuildList.ListItems.Clear
ClanTemporal = ""
IndexK = 1

For i = 1 To cantClanes
    ClanTemporal = ReadField(11 + i, Data, Asc("¬"))
    NombreClan = ReadField(1, ClanTemporal, Asc("-"))
    
    If UCase$(NombreClan) <> UCase$("cerrado" & i & "") Then
        FaccionClan = ReadField(2, ClanTemporal, Asc("-"))
        NivelClan = ReadField(3, ClanTemporal, Asc("-"))
        
        GuildUser_lstGuildList.ListItems.Add IndexK, , NombreClan
        
        If FaccionClan = 3 Then
            GuildUser_lstGuildList.ListItems(IndexK).ListSubItems.Add , , "NEUTRAL"
        ElseIf FaccionClan = 4 Or FaccionClan = 5 Then
            GuildUser_lstGuildList.ListItems(IndexK).ListSubItems.Add , , "ALIANZA"
        ElseIf FaccionClan = 2 Or FaccionClan = 3 Then
            GuildUser_lstGuildList.ListItems(IndexK).ListSubItems.Add , , "HORDA"
        End If
        
        GuildUser_lstGuildList.ListItems(IndexK).ListSubItems.Add , , NivelClan
        IndexK = IndexK + 1
    End If
Next i

tmpSuma = 12 + cantClanes
cantMiembros = Val(ReadField(tmpSuma, Data, Asc("¬")))
lblMiembros.Caption = cantMiembros
lblMaxMiembros.Caption = 4 * Val(cantMiembros)

Dim MiembroTemporal As String
MiembroTemporal = ""

For i = 1 To cantMiembros
    MiembroTemporal = ReadField(tmpSuma + i, Data, Asc("¬"))

    GuildUser_lstMembers.ListItems.Add , , ReadField(1, MiembroTemporal, Asc("$"))
    GuildUser_lstMembers.ListItems(i).bold = True
    
    GuildUser_lstMembers.ListItems(i).ListSubItems.Add , , ReadField(2, MiembroTemporal, Asc("$"))
    GuildUser_lstMembers.ListItems(i).ListSubItems.Add , , ReadField(3, MiembroTemporal, Asc("$"))
    GuildUser_lstMembers.ListItems(i).ListSubItems.Add , , ReadField(4, MiembroTemporal, Asc("$"))
Next i

mostrarGuildUser (True)

End Sub

Private Sub members_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    Call SendData("VLKG" & Members.ListItems.Item(Members.SelectedItem.Index).text)
    
    If RetiraObj = 1 And RetiraOro = 1 Then
        mnuGld.Checked = False
        mnuObjs.Checked = False
        mnuDep.Checked = False
        mnuFull.Checked = True
    ElseIf RetiraObj = 0 And RetiraOro = 1 Then
        mnuGld.Checked = True
        mnuObjs.Checked = False
        mnuDep.Checked = False
        mnuFull.Checked = False
    ElseIf RetiraObj = 1 And RetiraOro = 0 Then
        mnuObjs.Checked = True
        mnuGld.Checked = False
        mnuDep.Checked = False
        mnuFull.Checked = False
    ElseIf RetiraObj = 0 And RetiraOro = 0 Then
        mnuDep.Checked = True
        mnuObjs.Checked = False
        mnuGld.Checked = False
        mnuFull.Checked = False
    End If
    
    PopupMenu mnuLider
End If

End Sub
Private Sub mnuDep_Click()
    Call SendData("BOVC" & Members.ListItems.Item(Members.SelectedItem.Index).text & "," & 0)
End Sub
Private Sub mnuGld_Click()
    Call SendData("BOVC" & Members.ListItems.Item(Members.SelectedItem.Index).text & "," & 1)
End Sub
Private Sub mnuObjs_Click()
    Call SendData("BOVC" & Members.ListItems.Item(Members.SelectedItem.Index).text & "," & 2)
End Sub
Private Sub mnuFull_Click()
    Call SendData("BOVC" & Members.ListItems.Item(Members.SelectedItem.Index).text & "," & 3)
End Sub
Private Sub mnuExpulsar_Click()
    Call SendData("ECHARCLA" & Members.ListItems.Item(Members.SelectedItem.Index).text)
    Call SendData("GLINFO")
    mostrarAdminLeader (True)
End Sub
Private Sub mnuSubHacer_Click()
    Call SendData("/SUBLIDER " & Members.ListItems.Item(Members.SelectedItem.Index).text)
End Sub
Private Sub mnuSubSacar_Click()
    Call SendData("/QSUBLIDR " & Members.ListItems.Item(Members.SelectedItem.Index).text)
End Sub
Private Sub mnuDoLider_Click()
    If MsgBox("¿Está seguro que desea pasarle el liderazgo a " & Members.ListItems.Item(Members.SelectedItem.Index).text & "?", vbYesNo) = vbYes Then Call SendData("/HACLIDER " & Members.ListItems.Item(Members.SelectedItem.Index).text)
    Unload Me
End Sub
Private Sub mnucontacts_Click()
    Call SendData("ADDCON" & Members.ListItems.Item(Members.SelectedItem.Index).text)
End Sub
Private Sub mnumensaje_Click()
    TheUser = Members.ListItems.Item(Members.SelectedItem.Index).text
End Sub
Public Sub ParseCanjes(ByVal rData As String)

    Dim cantCanjes As Byte
    cantCanjes = Val(ReadField(1, rData, 44))

    frmMenuGral.ListaPremios.Clear
    For i = 1 To cantCanjes
            frmMenuGral.ListaPremios.AddItem ReadField(i + 1, rData, 44)
    Next i
    
    lPuntos.Caption = ReadField(2 + cantCanjes, rData, 44)
    lblTSPoints.Caption = ReadField(3 + cantCanjes, rData, 44)
    
    mostrarCanjes (True)

End Sub

Public Sub ParseCreditos(ByVal rData As String)

    creditos_lstPacks.Clear
    creditos_lblCredits.Caption = PonerPuntos(ReadField(1, rData, 44))
    For i = 1 To ReadField(2, rData, 44)
        creditos_lstPacks.AddItem ReadField(2 + i, rData, 44)
    Next i
    
    mostrarCreditos (True)

End Sub
Private Sub tCredits_Timer()
     Call engine.DrawDonations
End Sub
Public Sub ParseQuests(ByVal rData As String)

Quests_lstQuest.Clear

Dim cantQuests As Byte
cantQuests = Val(ReadField(1, rData, 44))

For i = 1 To cantQuests
    Quests_lstQuest.AddItem ReadField(1 + i, rData, 44)
Next i

mostrarQuests (True)

End Sub
Public Sub ParseDuelos(ByVal rData As String)

    duelos_Jugador1 = ReadField(1, rData, 44)
    duelos_Jugador2 = ReadField(2, rData, 44)
    duelos_Jugador3 = ReadField(3, rData, 44)
    duelos_Jugador4 = ReadField(4, rData, 44)
    duelos_Jugador5 = ReadField(5, rData, 44)
    duelos_Jugador6 = ReadField(6, rData, 44)
    duelos_Jugador7 = ReadField(7, rData, 44)
    duelos_Jugador8 = ReadField(8, rData, 44)
    mostrarDuelos (True)

End Sub
Public Sub ParseEstadisticas(ByVal rData As String)

        Estadisticas_Nivel.Caption = ReadField(1, rData, 44)
        Estadisticas_Reputacion.Caption = ReadField(2, rData, 44)
        Estadisticas_Clase.Caption = ReadField(3, rData, 44)
        Estadisticas_Raza.Caption = ReadField(4, rData, 44)
        Estadisticas_Genero.Caption = ReadField(5, rData, 44)
        Estadisticas_Hogar.Caption = ReadField(6, rData, 44)
        
        Estadisticas_TorneosParticipados.Caption = ReadField(7, rData, 44)
        Estadisticas_EventosGanados.Caption = ReadField(8, rData, 44)
        Estadisticas_DuelosGanados.Caption = ReadField(9, rData, 44)
        Estadisticas_ParejasGanadas.Caption = ReadField(10, rData, 44)
        Estadisticas_NPCsAsesinados.Caption = ReadField(11, rData, 44)
        Estadisticas_Muertes.Caption = ReadField(12, rData, 44)
        Estadisticas_QuestCompletadas.Caption = ReadField(13, rData, 44)
        
        Estadisticas_lblAtri(0).Caption = ReadField(14, rData, 44)
        Estadisticas_lblAtri(1).Caption = ReadField(15, rData, 44)
        Estadisticas_lblAtri(3).Caption = ReadField(16, rData, 44)
        Estadisticas_lblAtri(4).Caption = ReadField(17, rData, 44)
        Estadisticas_lblAtri(2).Caption = ReadField(18, rData, 44)
        
        Dim tmpAlineacion As Byte
        tmpAlineacion = ReadField(19, rData, 44)
        
        If tmpAlineacion = 1 Then
            Estadisticas_Faccion.ForeColor = &H80&
            Estadisticas_Faccion.Caption = "HORDA INFERNAL"
        ElseIf tmpAlineacion = 2 Then
            Estadisticas_Faccion.ForeColor = &HC00000
            Estadisticas_Faccion.Caption = "ALIANZA IMPERIAL"
        ElseIf tmpAlineacion = 0 Then
            Estadisticas_Faccion.ForeColor = &H404040
            Estadisticas_Faccion.Caption = "NEUTRAL"
        End If
        
        Estadisticas_Jerarquia.Caption = ReadField(20, rData, 44)
        Estadisticas_RequeridosJerarq.Caption = ReadField(21, rData, 44)
        
        Estadisticas_AlianzasMatados.Caption = ReadField(22, rData, 44)
        Estadisticas_HordasMatados.Caption = ReadField(23, rData, 44)
        
        Estadisticas_lblBonificadores(1).Caption = ReadField(24, rData, 44)
        Estadisticas_lblBonificadores(2).Caption = ReadField(25, rData, 44)
        Estadisticas_lblBonificadores(3).Caption = ReadField(26, rData, 44)
        
        mostrarEstadisticas (True)
End Sub

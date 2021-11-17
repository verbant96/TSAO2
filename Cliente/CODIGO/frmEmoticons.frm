VERSION 5.00
Begin VB.Form frmEmoticons 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Emoticons"
   ClientHeight    =   1995
   ClientLeft      =   3795
   ClientTop       =   4905
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmEmoticons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEmoticons.frx":000C
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   525
      Top             =   1560
      Width           =   2940
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   1020
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   22
      Left            =   75
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   21
      Left            =   570
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   20
      Left            =   1050
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   19
      Left            =   1530
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   18
      Left            =   2010
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   17
      Left            =   2490
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   16
      Left            =   3450
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   15
      Left            =   75
      Top             =   540
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   14
      Left            =   570
      Top             =   540
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   13
      Left            =   1050
      Top             =   540
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   12
      Left            =   1530
      Top             =   540
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   11
      Left            =   2010
      Top             =   540
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   10
      Left            =   2490
      Top             =   540
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   9
      Left            =   3450
      Top             =   540
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   8
      Left            =   2955
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   7
      Left            =   570
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   5
      Left            =   1530
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   1980
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   2970
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   2970
      Top             =   540
      Width           =   480
   End
End
Attribute VB_Name = "frmEmoticons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Nothing
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\emojis_CerrarVentanaI.jpg")
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\emojis_CerrarVentanaA.jpg")
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 1
SendData (";" & ":S")

Case 2
SendData (";" & ":(")

Case 3
SendData (";" & ":CA")

Case 4
SendData (";" & ";)")

Case 5
SendData (";" & ":$")

Case 6
SendData (";" & ">.>")

Case 7
SendData (";" & "?")

Case 8
SendData (";" & "!")

Case 9
SendData (";" & "...")

Case 10
SendData (";" & "¬¬")

Case 11
SendData (";" & ":@")

Case 12
SendData (";" & "º_º")

Case 13
SendData (";" & "-_-")

Case 14
SendData (";" & ":3")

Case 15
SendData (";" & "^^")

Case 16
SendData (";" & ":D")

Case 17
SendData (";" & ":P")

Case 18
SendData (";" & "'_'")

Case 19
SendData (";" & ":O")

Case 20
SendData (";" & "xD")

Case 21
SendData (";" & ":'(")

Case 22
SendData (";" & ":)")

End Select

Unload Me

End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

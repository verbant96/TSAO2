VERSION 5.00
Begin VB.Form frmCargando 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Tierras Sagradas"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgProgress 
      Enabled         =   0   'False
      Height          =   105
      Left            =   285
      Top             =   8400
      Width           =   11430
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WS_EX_APPWINDOW               As Long = &H40000
Private Const GWL_EXSTYLE                   As Long = (-20)
Private Const SW_HIDE                       As Long = 0
Private Const SW_SHOW                       As Long = 5

Private m_bActivated As Boolean
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private porcentajeActual As Long
Private Sub Form_Activate()
    If Not m_bActivated Then
        m_bActivated = True
        Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
        Call ShowWindow(hWnd, SW_HIDE)
        Call ShowWindow(hWnd, SW_SHOW)
    End If
End Sub
Public Sub establecerProgreso(ByVal nuevoPorcentaje As Long)
 
If nuevoPorcentaje >= 0 And nuevoPorcentaje <= 100 Then
   Dim indice As Integer, tmpWidth As Integer, i As Long
   indice = (762 * nuevoPorcentaje / 100) - imgProgress.Width
   tmpWidth = imgProgress.Width
   
   For i = 1 To indice
    imgProgress.Width = imgProgress.Width + 1
    Sleep 10
   Next i
   
ElseIf nuevoPorcentaje > 100 Then
    imgProgress.Width = 762
Else
    imgProgress.Width = 0
End If

porcentajeActual = nuevoPorcentaje
 
End Sub
Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Cargando_Main.jpg")
    imgProgress.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Cargando_Bar.jpg")
    porcentajeActual = 0
    imgProgress.Width = 0
End Sub

VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   ControlBox      =   0   'False
   Icon            =   "frmComerciar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   195
      Left            =   1170
      TabIndex        =   8
      Text            =   "1"
      Top             =   6630
      Width           =   990
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000001&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   285
      ScaleHeight     =   570
      ScaleWidth      =   525
      TabIndex        =   2
      Top             =   945
      Width           =   525
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   3990
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   3645
      TabIndex        =   1
      Top             =   1995
      Width           =   2520
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   3990
      Index           =   0
      IntegralHeight  =   0   'False
      ItemData        =   "frmComerciar.frx":000C
      Left            =   390
      List            =   "frmComerciar.frx":000E
      TabIndex        =   0
      Top             =   1995
      Width           =   2535
   End
   Begin VB.Image imgQuit 
      Height          =   495
      Left            =   6120
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click derecho para cerrar la ventana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   6900
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   2
      Left            =   3660
      MouseIcon       =   "frmComerciar.frx":0010
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   6645
      Width           =   2490
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   1
      Left            =   3660
      MouseIcon       =   "frmComerciar.frx":0162
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   6180
      Width           =   2490
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   0
      Left            =   420
      MouseIcon       =   "frmComerciar.frx":02B4
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   6180
      Width           =   2490
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Index           =   3
      Left            =   4170
      TabIndex        =   7
      Top             =   1440
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Index           =   4
      Left            =   4170
      TabIndex        =   6
      Top             =   1200
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Index           =   2
      Left            =   885
      TabIndex        =   5
      Top             =   1275
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Index           =   0
      Left            =   930
      TabIndex        =   4
      Top             =   900
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10.000.000"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Index           =   1
      Left            =   2325
      TabIndex        =   3
      Top             =   1290
      Width           =   810
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************
'*****************************
'*****      Samke       ******
'*****************************
'**************************************************
'**************************************************
'*****      SoHnsalxixon_u2@hotmail.com      ******
'**************************************************
'**************************************************

Private tmpItemSelect As Byte

Private Todo As Byte
Private m_Interval As Integer
Private m_Number As Integer
Private m_Increment As Integer
Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private Sub cantidad_Change()
    If Val(cantidad.text) < 1 Then
        cantidad.text = 1
    End If
    
    If Val(cantidad.text) > MAX_INVENTORY_OBJS Then
        cantidad.text = 1
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus

End Sub

Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me
'If Configuracion.Alpha_Interfaz_Activar > 0 Then MakeTransparent Me.hWnd, Configuracion.Alpha_Interfaz_Transparencia


List1(0).BackColor = RGB(19, 21, 22)
List1(1).BackColor = RGB(19, 21, 22)
cantidad.BackColor = RGB(19, 21, 22)

Label1(0).ForeColor = RGB(145, 123, 85)
Label1(2).ForeColor = RGB(145, 123, 85)
Label1(3).ForeColor = RGB(145, 123, 85)
Label1(4).ForeColor = RGB(145, 123, 85)
cantidad.ForeColor = RGB(145, 123, 85)
List1(0).ForeColor = RGB(145, 123, 85)
List1(1).ForeColor = RGB(145, 123, 85)

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\comerciar.jpg")
m_Number = 1
m_Interval = 30
Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_N.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
    Image1(1).Tag = 1
End If
Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_N.jpg")
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    Call SendData("FINBAN")
    Unload Me
End If
End Sub

Private Sub Image1_Click(Index As Integer)

On Error Resume Next

Select Case Index
    Case 0
        tmpItemSelect = List1(0).ListIndex + 1
        
        If tmpItemSelect <= 0 Then Exit Sub
        LastIndex1 = List1(0).ListIndex
        If UserGLD >= NPCInventory(tmpItemSelect).Valor * Val(cantidad) Then
                SendData ("COMP" & "," & tmpItemSelect & "," & cantidad.text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If

    List1(0).Clear
    List1(1).Clear
        
   Case 1
        If List1(Index).ListIndex + 1 <= 0 Or List1(Index).ListIndex + 1 > MAX_INVENTORY_SLOTS Then Exit Sub
        tmpItemSelect = slotsListaInv(List1(Index).ListIndex + 1)
   
        If tmpItemSelect <= 0 Then Exit Sub
        
            LastIndex2 = List1(1).ListIndex
            If Not Inventario.Equipped(tmpItemSelect) Then
                SendData ("VEND" & "," & tmpItemSelect & "," & cantidad.text)
            Else
                AddtoRichTextBox frmMain.RecTxt, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
            End If
        
        List1(0).Clear
        List1(1).Clear
        
    Case 2
        If List1(Todo).ListIndex >= 0 Then
            If Todo = 1 Then
                cantidad = Label1(2).Caption
            Else
                cantidad = Label1(2).Caption
            End If
        End If
                
End Select

NPCInvDim = 0
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
                Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_I.jpg")
                Image1(0).Tag = 0
                Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
                Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
                Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_I.jpg")
                Image1(1).Tag = 0
                Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
                Image1(0).Tag = 1
        End If
    Case 2
        Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_I.jpg")
        'If Image1(2).Tag = 1 Then
        '        Image1(2).Picture = general_load_interface_picture("Botónokapretado.jpg")
        '        Image1(2).Tag = 0
        'End If
        
End Select
End Sub

Private Sub cantidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgQuit_Click()
    Call SendData("FINBAN")
    Unload Me
End Sub
Private Sub List1_Click(Index As Integer)
Dim SR As RECT, DR As RECT, GrhIndex As Long

On Error Resume Next
 
SR.left = 0
SR.top = 0
SR.Right = 35
SR.bottom = 38

DR.left = 0
DR.top = 0
DR.Right = 35
DR.bottom = 38

Todo = Index

Select Case Index
    Case 0
        If List1(0).ListIndex + 1 = 0 Or List1(0).ListIndex + 1 > MAX_NPC_INVENTORY_SLOTS Then Exit Sub
        tmpItemSelect = slotsListaNPC(List1(0).ListIndex + 1)
        
        If tmpItemSelect <= 0 Then Exit Sub
        
        Label1(0).Caption = NPCInventory(tmpItemSelect).Name
        Label1(1).Caption = PonerPuntos(NPCInventory(tmpItemSelect).Valor)
        Label1(2).Caption = NPCInventory(tmpItemSelect).Amount
        GrhIndex = NPCInventory(tmpItemSelect).GrhIndex
        Select Case NPCInventory(tmpItemSelect).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe: " & NPCInventory(tmpItemSelect).MaxHit
                Label1(4).Caption = "Min Golpe: " & NPCInventory(tmpItemSelect).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & NPCInventory(tmpItemSelect).Def
                Label1(4).Visible = True
            Case 16
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & NPCInventory(tmpItemSelect).Def
                Label1(4).Visible = True
            Case 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & NPCInventory(tmpItemSelect).Def
                Label1(4).Visible = True
        End Select
        Call engine.DrawGrhtoHdc(GrhIndex, SR, Picture1)
    Case 1
        tmpItemSelect = slotsListaInv(List1(1).ListIndex + 1)
        
        If tmpItemSelect <= 0 Then Exit Sub
        
        Label1(0).Caption = Inventario.ItemName(tmpItemSelect)
        Label1(1).Caption = Inventario.Valor(tmpItemSelect)
        Label1(2).Caption = Inventario.Amount(tmpItemSelect)
        GrhIndex = Inventario.GrhIndex(tmpItemSelect)
        Select Case Inventario.OBJType(tmpItemSelect)
            Case 2
                Label1(3).Caption = "Max Golpe: " & Inventario.MaxHit(tmpItemSelect)
                Label1(4).Caption = "Min Golpe: " & Inventario.MinHit(tmpItemSelect)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & Inventario.Def(tmpItemSelect)
                Label1(4).Visible = True
            Case 16
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & Inventario.Def(tmpItemSelect)
                Label1(4).Visible = True
            Case 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & Inventario.Def(tmpItemSelect)
                Label1(4).Visible = True
        End Select
        Call engine.DrawGrhtoHdc(Inventario.GrhIndex(tmpItemSelect), SR, Picture1)
End Select

'Call engine.GrhRenderToHdc(GrhIndex, Picture1.hDC, 0, 0)

End Sub

Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
    Image1(1).Tag = 1
End If
End Sub

Private Sub tmrNumber_Timer()

Const MIN_NUMBER = 1
Const MAX_NUMBER = 10000

    m_Number = m_Number + m_Increment
    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER
    End If

    cantidad.text = Format$(m_Number)
    
End Sub
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_A.jpg")
If Index = 1 Then Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_A.jpg")
If Index = 2 Then Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_A.jpg")
End Sub
Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
If Index = 1 Then Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
If Index = 2 Then Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_N.jpg")
End Sub



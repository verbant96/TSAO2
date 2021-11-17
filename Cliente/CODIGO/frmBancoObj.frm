VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   ControlBox      =   0   'False
   Icon            =   "frmBancoObj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   1170
      TabIndex        =   7
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
      Left            =   405
      TabIndex        =   0
      Top             =   1995
      Width           =   2520
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   6120
      Top             =   0
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   2
      Left            =   3660
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   4170
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Index           =   4
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   1290
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Index           =   0
      Left            =   945
      TabIndex        =   3
      Top             =   900
      Width           =   3120
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

Public Todo As Byte
Private tmpSelectItem As Integer

Private form_Mov As clsFormMovementManager

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private Sub cantidad_Change()
If Val(cantidad.text) < 0 Then
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
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

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

'Cargamos la interfase
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Main.jpg")
Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Retirar_N.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depositar_N.jpg")
Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depo_Todo_N.jpg")

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    Unload Me
    Call SendData("FINBAN")
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Retirar_N.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depositar_N.jpg")
    Image1(1).Tag = 1
End If

Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depo_Todo_N.jpg")

End Sub
Private Sub Image1_Click(Index As Integer)

    On Error Resume Next

Call Audio.PlayWave("click.wav")

Select Case Index
    Case 0
        If List1(Index).List(List1(Index).ListIndex) = "Nada" Or _
            List1(Index).ListIndex < 0 Then Exit Sub
            frmBancoObj.List1(0).SetFocus
            LastIndex1 = List1(0).ListIndex
        
        SendData ("RETI" & "," & List1(0).ListIndex + 1 & "," & cantidad.text)
        
   Case 1
    If Not List1(1).ListIndex + 1 > 0 Or List1(1).ListIndex + 1 > MAX_INVENTORY_SLOTS Then Exit Sub
    tmpSelectItem = slotsListaInv(List1(1).ListIndex + 1)
        If tmpSelectItem < 0 Then Exit Sub
        
        LastIndex2 = List1(1).ListIndex
        
        If Not Inventario.Equipped(tmpSelectItem) Then
            SendData ("DEPO" & "," & tmpSelectItem & "," & cantidad.text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes depositar el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If

    Case 2
        If List1(Todo).ListIndex >= 0 Then
            If Todo = 1 Then
                cantidad = Label1(2).Caption
            Else
                cantidad = Label1(2).Caption
            End If
        End If
End Select

If Index < 2 Then
List1(0).Clear
List1(1).Clear
End If

NPCInvDim = 0
End Sub
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 2 Then Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depo_Todo_A.jpg")
If Index = 1 Then Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depositar_A.jpg")
If Index = 0 Then Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Retirar_A.jpg")
End Sub
Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
                Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Retirar_I.jpg")
                Image1(0).Tag = 0
                Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depositar_N.jpg")
                Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
                Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depositar_I.jpg")
                Image1(1).Tag = 0
                Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Retirar_N.jpg")
                Image1(0).Tag = 1
        End If

    Case 2
        Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depo_Todo_I.jpg")
        
End Select
End Sub

Private Sub Image2_Click()
Unload Me
Call SendData("FINBAN")
End Sub
Private Sub List1_Click(Index As Integer)
Dim SR As RECT, DR As RECT, GrhIndex As Long

SR.left = 0
SR.top = 0
SR.Right = 38
SR.bottom = 35

DR.left = 0
DR.top = 0
DR.Right = 38
DR.bottom = 35

Todo = Index

Select Case Index
    Case 0
        Label1(0).Caption = UserBancoInventory(List1(0).ListIndex + 1).Name
        Label1(2).Caption = UserBancoInventory(List1(0).ListIndex + 1).Amount
        GrhIndex = UserBancoInventory(List1(0).ListIndex + 1).GrhIndex
        Select Case UserBancoInventory(List1(0).ListIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe: " & UserBancoInventory(List1(0).ListIndex + 1).MaxHit
                Label1(4).Caption = "Min Golpe: " & UserBancoInventory(List1(0).ListIndex + 1).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & UserBancoInventory(List1(0).ListIndex + 1).Def
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
    Case 1
        tmpSelectItem = slotsListaInv(List1(1).ListIndex + 1)
        Label1(0).Caption = Inventario.ItemName(tmpSelectItem)
        Label1(2).Caption = Inventario.Amount(tmpSelectItem)
        GrhIndex = Inventario.GrhIndex(tmpSelectItem)
        Select Case Inventario.OBJType(tmpSelectItem)
            Case 2
                Label1(3).Caption = "Max Golpe: " & Inventario.MaxHit(tmpSelectItem)
                Label1(4).Caption = "Min Golpe: " & Inventario.MinHit(tmpSelectItem)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & Inventario.Def(tmpSelectItem)
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
End Select

If GrhIndex = 0 Then
    Picture1.Picture = Nothing
Else
    Call engine.DrawGrhtoHdc(GrhIndex, SR, Picture1)
End If
'Call engine.GrhRenderToHdc(GrhIndex, Picture1.hDC, 0, 0)
'Picture1.Refresh

End Sub
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Retirar_N.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Boveda_Depositar_N.jpg")
    Image1(1).Tag = 1
End If
End Sub

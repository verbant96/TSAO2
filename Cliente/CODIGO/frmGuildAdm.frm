VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGuildAdm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5805
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildAdm.frx":0000
   ScaleHeight     =   5460
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView GuildList 
      Height          =   4200
      Left            =   150
      TabIndex        =   0
      Top             =   645
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   7408
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   5027
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Faccion"
         Object.Width           =   3263
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nivel"
         Object.Width           =   1508
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   5160
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   1800
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   2205
   End
   Begin VB.Image Image3 
      Height          =   5535
      Left            =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Const InterfaceName As String = "RegistredGuilds"

Private Sub Form_Load()

Me.Picture = General_Load_Interface_Picture("RegistredGuilds_Main.jpg")
ChangeButtonsNormal

End Sub
Private Sub Image2_Click()
Unload Me
End Sub
Private Sub Image1_Click()

If GuildList.SelectedItem.Index <= 0 Then Exit Sub
If GuildList.ListItems.Item(GuildList.SelectedItem.Index).Text = "" Then Exit Sub

Call SendData("CLANDETAILS" & GuildList.ListItems.Item(GuildList.SelectedItem.Index).Text)
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = ChangeButtonState(Apretado, "BDetalles")
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Tag = "0" Then
    Call ChangeButtonsNormal
    Image1.Picture = ChangeButtonState(Iluminado, "BDetalles")
    Image1.Tag = "1"
End If
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeButtonsNormal
End Sub
Public Sub ParseGuildList(ByVal Rdata As String)

Dim j As Long, k As Integer
'For j = 0 To GuildsList.ListCount - 1
'    Me.GuildsList.RemoveItem 0
'Next j
k = CInt(ReadField(1, Rdata, Asc(",")))

'For j = 1 To k
'    GuildsList.AddItem ReadField(1 + j, Rdata, 44)
'Next j
'Dim i, o, TotalItems As Long

'TotalItems = frmGuildAdm.GuildsList.ListCount
    
'    For i = 0 To TotalItems
'    For o = 0 To TotalItems
'        If frmGuildAdm.GuildsList.List(i) = "cerrado" & o Then
'            frmGuildAdm.GuildsList.RemoveItem (i)
'        End If
'    Next
'Next

Call Aplicar_Transparencia(Me.hWnd, 240)
Call Aplicar_Transparencia(GuildList.hWnd, 100)


Dim ClanTemporal As String
Dim NombreClan As String
Dim FaccionClan As String
Dim NivelClan As Byte
Dim IndexK As Integer
GuildList.ListItems.Clear
ClanTemporal = ""
IndexK = 1

For j = 1 To k
    ClanTemporal = ReadField(j + 1, Rdata, Asc(","))
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

'SetWindowLong GuildList.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
'GuildList.BackColor = WS_EX_TRANSPARENT

Me.Show vbModal, frmMain

End Sub

Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function

Private Sub ChangeButtonsNormal()

Image1.Picture = ChangeButtonState(BNormal, "BDetalles")

Dim j
For Each j In Me
    j.Tag = "0"
Next

Me.Tag = "0"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.Tag = "0" Then
    Call ChangeButtonsNormal
    Me.Tag = "1"
End If

End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.Tag = "0" Then
    Call ChangeButtonsNormal
    Me.Tag = "1"
End If

End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Arrastrar(Me)
End Sub
Private Function Arrastrar(frmGuildAdm As Form)
    ReleaseCapture
    Call SendMessage(frmGuildAdm.hWnd, &HA1, 2, 0&)
End Function

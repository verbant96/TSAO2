VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim InvRect As RECT
Dim InvD3DRect As D3DRECT
Dim color(3) As Long

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Public InvSelectedItem As Long             'Currently selected item
Dim MouseDownSelectedItem As Long
Dim MouseX As Integer
Dim MouseY As Integer

'Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

Dim ItemMenu As Menu    'Menu to be shown as pop up
Private Sub Class_Initialize()
'***************************************************
'Last Modify Date: 18/05/11
'Author Modify: ShaFTeR
'Extracted to: GSZ-AO by •Parra
'***************************************************
    Set InventoryWindow = frmMain.picInv
    InventoryWindow.AutoRedraw = True
    InventoryWindow.ScaleMode = 3
   
    InvRect.top = 0: InvD3DRect.Y1 = 0
    InvRect.left = 0: InvD3DRect.X1 = 0
    InvRect.bottom = 174: InvD3DRect.Y2 = 174
    InvRect.Right = 174: InvD3DRect.X2 = 174
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub ScrollInventory(ByVal up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Scrolls the graphical inventory up / down
'***************************************************
    'Only move if there is something to show up /down
    If up Then
        If (InventoryOffset + (InventoryWindow.ScaleHeight \ 34)) * (InventoryWindow.ScaleWidth \ 34) < MAX_INVENTORY_SLOTS Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
    
End Sub

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO

End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = UserInventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = UserInventory(slot).MinHit
End Property

Public Property Get Def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    Def = UserInventory(slot).Def
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = UserInventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = UserInventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = UserInventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = UserInventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
   ItemName = UserInventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = UserInventory(slot).Equipped
End Property

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Long, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With

End Sub
Public Sub ActualizarSlotEquipped(ByVal slot As Byte, ByVal eEquipped As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Equipped = eEquipped
    End With
    
End Sub
Public Sub ActualizarSlotCant(ByVal slot As Byte, ByVal eAmount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
    End With
    
End Sub
Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
   
    temp_x = X \ 34 'TilePixelWidth
    temp_y = Y \ 34 'TilePixelHeight
   
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 34) + 1
   
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        ClickItem = TempItem
    End If
    
End Function
Public Sub DrawInventory()
 
    Dim X, Y As Integer
    Dim s(3) As Long
    Dim i As Long
    
    Static re As RECT
    re.left = 0
    re.top = 0
    re.bottom = 174
    re.Right = 174
    
    s(0) = D3DColorXRGB(255, 255, 255)
    s(1) = s(0)
    s(2) = s(0)
    s(3) = s(0)
    

DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
DirectDevice.BeginScene
    
    Call engine.Device_Box_Textured_Render(31570, 0, 0, 174, 174, s, 0, 0)
   
    For i = 1 To UBound(UserInventory)
    
    If MouseOK = True Then
        Call engine.Draw_GrhIndex(MouseItem, MouseX, MouseY)
    End If
 
        If UserInventory(i).GrhIndex Then
            X = ((i - 1) Mod (InventoryWindow.Width / 34)) * 34 + 1
            Y = ((i - 1) \ (InventoryWindow.Width / 34)) * 34
            
             
            If i > 10 And i < 15 Then
                Y = Y + 2
            ElseIf i > 5 And i < 10 Then
                Y = Y + 1
            ElseIf i > 15 And i < 20 Then
                Y = Y + 3
            ElseIf i > 20 Then
                Y = Y + 4
            End If
             
            Select Case i
            
                Case 2, 7, 12, 17, 22
                    X = X + 1
                    
                Case 3, 8, 13, 18, 23
                    X = X + 2
                    
                Case 4, 9, 14, 19, 24
                    X = X + 3
                
                Case 5, 10, 15, 20, 25
                    X = X + 4
            End Select
             
             
            If InvSelectedItem = i Then
                Call engine.Draw_GrhIndex(32758, X, Y)
                'Call engine.Draw_FilledBox(X, Y, 32, 32, D3DColorXRGB(30, 30, 30), D3DColorXRGB(142, 142, 142))
            End If
 
            If DibujadoContinuoInv = True And i = InvSelectedItem And DyDActivado = True Then
                Call engine.Draw_GrhIndex(UserInventory(i).GrhIndex, X, Y)
                Call engine.Draw_GrhInv(UserInventory(i).GrhIndex, frmMain.MouseXInv, frmMain.MouseYInv)
                
                If frmMain.MouseXInv < 0 Or frmMain.MouseXInv < 0 Or frmMain.MouseXInv > InventoryWindow.Width Or frmMain.MouseYInv > InventoryWindow.Height Then
                    Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
                End If
                
            Else
                Call engine.Draw_GrhIndex(UserInventory(i).GrhIndex, X, Y - 1)
            End If
            
            Texto.Engine_Text_Draw X - 1, Y + 3, Str$(UserInventory(i).Amount), D3DColorXRGB(255, 255, 255)
 
            If UserInventory(i).Equipped Then
                Texto.Engine_Text_Draw X + 23, Y + 20, "E", D3DColorXRGB(255, 255, 0)
            End If
        End If
 
    Next i
    
DirectDevice.EndScene
DirectDevice.Present re, ByVal 0, frmMain.picInv.hWnd, ByVal 0


 
End Sub
Public Sub Initialize(ByRef InvPic As PictureBox)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the reference to the picture box on which to render the inventory
'***************************************************
    Set InventoryWindow = InvPic
    
    '\TODO
    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    temp_x = X \ 34
    temp_y = Y \ 34
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 34) + 1
    
    If TempItem <= MAX_INVENTORY_SLOTS Then
        InventoryWindow.ToolTipText = UserInventory(TempItem).Name
      If UserInventory(TempItem).Name = "(None)" Then
        frmMain.ItemName.Caption = "(Nada) - " & UserInventory(TempItem).Amount & ""
      Else
        frmMain.ItemName.Caption = UserInventory(TempItem).Name & " - " & UserInventory(TempItem).Amount
      End If
    End If
    
    If MouseItem = 0 And DibujadoContinuoInv = True Then
        MouseItem = UserInventory(TempItem).GrhIndex
    End If
    
    If MouseItem = 0 And DibujadoContinuoInv = True Then
        MouseItem = UserInventory(TempItem).GrhIndex
    End If
    
    If DyDActivado Then
        MouseX = X
        MouseY = Y
        MouseOK = True
    End If
    
    If PUEDO = True Then
        MouseX = X
        MouseY = Y
        MouseOK = True
        ButtonIN = True
    End If
    
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    If frmMain.picInv.Visible = False Then Exit Sub
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    
      If InvSelectedItem <> MouseDownSelectedItem And MouseDownSelectedItem <> 0 And InvSelectedItem <> 0 Then
        Call SendData("SWAP" & InvSelectedItem & "," & MouseDownSelectedItem)
    End If
    
    If DyDActivado Then
        MouseOK = False
        MouseItem = 0
    End If
    
        MouseItem = 0
    
         If Button = vbKeyRButton Then
            If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
                SendData "EQUI" & Inventario.SelectedItem
            End If

End Sub
Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempItem As Long

    If DyDActivado Then
        PUEDO = True
        MouseItem = 0
        
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    TempItem = ClickItem(X, Y)
       
        If UserInventory(TempItem).GrhIndex Then
            MouseDownSelectedItem = TempItem
            Exit Sub
        End If
    End If
   
    MouseDownSelectedItem = 0
   
End Sub
Public Sub MoverItemInv(equisde As Integer)
    
If equisde = 0 Then

'movemos hacia la izquierda
Call SendData("SWAP" & InvSelectedItem & "," & InvSelectedItem - 1)

ElseIf equisde = 1 Then

'movemos hacia la derecha
Call SendData("SWAP" & InvSelectedItem & "," & InvSelectedItem + 1)

ElseIf equisde = 2 Then

'movemos hacia abajo
Call SendData("SWAP" & InvSelectedItem & "," & InvSelectedItem + 5)

ElseIf equisde = 3 Then

'movemos hacia arriba
If InvSelectedItem - 5 < 0 Then
Call SendData("SWAP" & InvSelectedItem & "," & InvSelectedItem + 15)
Else
Call SendData("SWAP" & InvSelectedItem & "," & InvSelectedItem - 5)
End If

End If

End Sub


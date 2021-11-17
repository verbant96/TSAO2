VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0
Private Type tLight
    RGBColor As D3DCOLORVALUE
    active As Boolean
    map_x As Byte
    map_y As Byte
    Range As Byte
    ID As Long
End Type
 
Private light_list() As tLight
Private NumLights As Long
Dim light_count As Long
Dim light_last As Long
Public Function Create_Light_To_Map(ByVal map_x As Byte, ByVal map_y As Byte, Optional Range As Byte = 3, Optional ByVal red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255)
    NumLights = NumLights + 1
   
    ReDim Preserve light_list(1 To NumLights) As tLight
   
    'Le damos color
    light_list(NumLights).RGBColor.r = red
    light_list(NumLights).RGBColor.g = Green
    light_list(NumLights).RGBColor.B = Blue
   
    'Alpha (Si borras esto RE KB!!)
    light_list(NumLights).RGBColor.a = 255
   
    'Un rango determinado
    light_list(NumLights).Range = Range
   
    'La activamos para que se pueda renderizar
    light_list(NumLights).active = True
   
    'Pos en el mapa
    light_list(NumLights).map_x = map_x
    light_list(NumLights).map_y = map_y
   
    Call LightRender(NumLights)
End Function
Public Function Delete_Light_To_Map(ByVal X As Byte, ByVal Y As Byte)
   
    Dim i As Long
   
    For i = 1 To NumLights
        If light_list(i).map_x = X And light_list(i).map_y = Y Then
            Delete_Light_To_Index i
            Exit Function
        End If
    Next i
 
End Function
Public Function Delete_Light_To_Index(ByVal light_index As Byte)
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim Ya As Integer
    Dim Xa As Integer
 
    light_list(light_index).active = False
 
    min_x = light_list(light_index).map_x - light_list(light_index).Range
    max_x = light_list(light_index).map_x + light_list(light_index).Range
    min_y = light_list(light_index).map_y - light_list(light_index).Range
    max_y = light_list(light_index).map_y + light_list(light_index).Range
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                MapData(Xa, Ya).light_value(0) = 0
                MapData(Xa, Ya).light_value(1) = 0
                MapData(Xa, Ya).light_value(2) = 0
                MapData(Xa, Ya).light_value(3) = 0
            End If
        Next Xa
    Next Ya
   
End Function
Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
   
    Dim CurrentColor As D3DCOLORVALUE
   
    pRadio = cRadio * 32
   
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
   
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
   
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(Round(CurrentColor.r), Round(CurrentColor.g), Round(CurrentColor.B))
        'If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight
    End If
End Function
Private Sub LightRender(ByVal light_index As Integer)
 
    If light_index = 0 Then Exit Sub
    If light_list(light_index).active = False Then Exit Sub
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim color As Long
    Dim Ya As Integer
    Dim Xa As Integer
   
    Dim TileLight As D3DCOLORVALUE
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
   
    Dim XCoord As Integer
    Dim YCoord As Integer
   
    AmbientColor.r = 160
    AmbientColor.g = 160
    AmbientColor.B = 160
   
    LightColor = light_list(light_index).RGBColor
       
    min_x = light_list(light_index).map_x - light_list(light_index).Range
    max_x = light_list(light_index).map_x + light_list(light_index).Range
    min_y = light_list(light_index).map_y - light_list(light_index).Range
    max_y = light_list(light_index).map_y + light_list(light_index).Range
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(1) = LightCalculate(light_list(light_index).Range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)
 
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(3) = LightCalculate(light_list(light_index).Range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)
                       
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(0) = LightCalculate(light_list(light_index).Range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)
   
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(2) = LightCalculate(light_list(light_index).Range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)
               
            End If
        Next Xa
    Next Ya
End Sub
Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function
 
Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, _
                            Optional ByVal Range As Byte = 1, Optional ByVal ID As Long, Optional ByVal red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If InMapBounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, Range, ID, red, Green, Blue
    End If
End Function
 
Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
       
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
   
            Light_Move = True
           
        End If
    End If
End Function
 
Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim addy As Byte
    Dim addx As Byte
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If
 
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
   
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
       
 
 
        Select Case Heading
            Case NORTH
                addy = -1
       
            Case EAST
                addx = 1
       
            Case SOUTH
                addy = 1
           
            Case WEST
                addx = -1
        End Select
       
        nX = map_x + addx
        nY = map_y + addy
       
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
       
            'Move it
            Light_Erase light_index
 
            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
   
            Light_Move_By_Head = True
           
        End If
    End If
End Function
 
Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                        ByVal Range As Long, Optional ByVal ID As Long, Optional ByVal red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
   
    'Make active
    light_list(light_index).active = True
   
        'Le damos color
    light_list(light_index).RGBColor.r = red
    light_list(light_index).RGBColor.g = Green
    light_list(light_index).RGBColor.B = Blue
   
    'Alpha (Si borras esto RE KB!!)
    light_list(light_index).RGBColor.a = 255
   
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).Range = Range
    light_list(light_index).ID = ID
    
End Sub
 
Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function
 
Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
           
    For loop_counter = 1 To light_count
       
        If light_list(loop_counter).active Then
            LightRender loop_counter
        End If
   
    Next loop_counter
End Sub
 
Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function
 
Public Function Light_Find(ByVal ID As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until light_list(loopc).ID = ID
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function
 
Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
   
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
   
    Light_Remove_All = True
End Function
 
Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim Temp As tLight
   
    Light_Erase light_index
   
    light_list(light_index) = Temp
   
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub
 
Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
   
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).Range
    min_y = light_list(light_index).map_y - light_list(light_index).Range
    max_x = light_list(light_index).map_x + light_list(light_index).Range
    max_y = light_list(light_index).map_y + light_list(light_index).Range
   
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = 0
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = 0
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = 0
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = 0
    End If
   
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = 0
            MapData(X, min_y).light_value(2) = 0
        End If
    Next X
   
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = 0
            MapData(X, max_y).light_value(3) = 0
        End If
    Next X
   
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = 0
            MapData(min_x, Y).light_value(3) = 0
        End If
    Next Y
   
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = 0
            MapData(max_x, Y).light_value(1) = 0
        End If
    Next Y
   
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = 0
                MapData(X, Y).light_value(1) = 0
                MapData(X, Y).light_value(2) = 0
                MapData(X, Y).light_value(3) = 0
            End If
        Next Y
    Next X
End Sub

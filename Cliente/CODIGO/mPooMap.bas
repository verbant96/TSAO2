Attribute VB_Name = "mPooMap"
Public Function Map_PosExitsObject(ByVal X As Byte, ByVal Y As Byte) As Integer
 
      '*****************************************************************
      'Checks to see if a tile position has a char_index and return it
      '*****************************************************************

      If (Map_InBounds(X, Y)) Then
            Map_PosExitsObject = MapData(X, Y).ObjGrh.GrhIndex
      Else
            Map_PosExitsObject = 0
      End If
 
End Function

Public Function Char_MapPosExits(ByVal X As Byte, ByVal Y As Byte) As Integer
 
    '*****************************************************************
    'Checks to see if a tile position has a char_index and return it
    '*****************************************************************
   
    Char_MapPosExits = MapData(X, Y).charindex
  
End Function

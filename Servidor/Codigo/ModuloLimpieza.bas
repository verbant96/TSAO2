Attribute VB_Name = "ModuloLimpieza"
Option Explicit
 
Type CleanWorld
   Map As Integer
   X As Integer
   Y As Integer
   Tiempo As Integer
   ObjIndex As Integer
End Type
 
Dim i As Long
 
'VALOR MÁXIMO DE LOS OBJ EN EL PISO
Public Const MAX_OBJS_CLEAR As Integer = 4000
Public tClearWorld(1 To MAX_OBJS_CLEAR) As CleanWorld
Public Sub CleanWorld_Initialize()
    
    For i = 1 To MAX_OBJS_CLEAR
        tClearWorld(i).Map = 0
        tClearWorld(i).ObjIndex = 0
        tClearWorld(i).Tiempo = 0
        tClearWorld(i).X = 0
        tClearWorld(i).Y = 0
    Next i

End Sub
Public Sub CleanWorld_AddItem(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Tiempo As Integer, ByVal ObjIndex As Integer)
 On Error Resume Next
 
'Le damos un indice donde esté desocupado y cerramos el for
For i = 1 To MAX_OBJS_CLEAR
    If tClearWorld(i).Map = 0 And tClearWorld(i).X = 0 And tClearWorld(i).Y = 0 And tClearWorld(i).Tiempo = 0 Then
        With tClearWorld(i)
         .Map = Map
         .X = X
         .Y = Y
         .Tiempo = Tiempo
         .ObjIndex = ObjIndex
        End With
     Exit For
    End If
Next i
 
End Sub
Public Sub CleanWorld_Clear()
On Error Resume Next

For i = 1 To MAX_OBJS_CLEAR
    'Si el indice tiene un objeto
    If tClearWorld(i).Map <> 0 And tClearWorld(i).X <> 0 And tClearWorld(i).Y <> 0 Then
        
        'Restamos un minuto al objeto
        If tClearWorld(i).Tiempo > 0 Then
            tClearWorld(i).Tiempo = tClearWorld(i).Tiempo - 1
        End If
        
        With tClearWorld(i)
            'Si ya se rastrearon el obj (o si lo habian rastreado y ya tiraron otro), borramos la data que habia quedado guardada.
            'Lo hago aca para no seguir haciendo negradas en el GetObj
            If MapData(.Map, .X, .Y).OBJInfo.ObjIndex = 0 Or MapData(.Map, .X, .Y).OBJInfo.ObjIndex <> .ObjIndex Then
                .Map = 0
                .X = 0
                .Y = 0
                .Tiempo = 0
                .ObjIndex = 0
            End If
    
            'Si el objeto llego a 0 minutos ya procedemos a borrarlo a la pija
            If tClearWorld(i).Tiempo = 0 And .Map <> 0 And .X <> 0 And .Y <> 0 Then
                    If MapData(.Map, .X, .Y).OBJInfo.ObjIndex > 0 Then
                        Call EraseObj(ToMap, 0, .Map, 10000, .Map, .X, .Y)
                        .Map = 0
                        .X = 0
                        .Y = 0
                        .Tiempo = 0
                        .ObjIndex = 0
                    End If
            End If
        End With
    End If
Next i
 
End Sub

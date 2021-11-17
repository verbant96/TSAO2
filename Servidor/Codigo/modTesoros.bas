Attribute VB_Name = "modTesoros"
Public CantMapitas As Byte
 
Public MapaTesoro As Integer
Public RecompenzaTesoro As Integer
Public MapaTesoroMap As Integer
Public MapaTesoroX As Integer
Public MapaTesoroY As Integer
Public TiempoTesoro As Integer
Public TesoroContando As Boolean
Public SepuedeDesenterrar As Boolean
Public Const LlaveTesoro As Integer = 1062 'num del mapa
Public ObjetoT As Obj
Public objetoCofreAbierto As Obj
 
Public Sub Tesoros()
       
    ObjetoT.Amount = 1
    ObjetoT.ObjIndex = 11 'Cofre Cerrado
   
    objetoCofreAbierto.Amount = 1
    objetoCofreAbierto.ObjIndex = 10 'Cofre abierto
   
Dim mapitakkk As Byte
CantMapitas = GetVar(App.Path & "\Dat\" & "Tesoros.dat", "MAPAS", "Num")
mapitakkk = RandomNumber(1, CantMapitas)
 
MapaTesoroMap = GetVar(App.Path & "\Dat\" & "Tesoros.dat", "MAPAS", "Mapa" & mapitakkk)
MapaTesoroX = RandomNumber(10, 90)
MapaTesoroY = RandomNumber(10, 90)

If MapData(MapaTesoroMap, MapaTesoroX, MapaTesoroY).Blocked = 1 Then
Call Tesoros
Exit Sub
End If
 
    SepuedeDesenterrar = False
    TesoroContando = False
    TiempoTesoro = 30
    Call SendData(SendTarget.ToAdmins, 0, 0, "||452@" & MapaTesoroMap & "@" & MapaTesoroX & "@" & MapaTesoroY)
End Sub
 
 
Public Sub DondeTesoros()
    Call SendData(SendTarget.toall, 0, 0, "||453@" & MapaTesoro & "@" & MapaTesoroX & "@" & MapaTesoroY)
End Sub
 
Public Sub CofreAbierto()
Call EraseObj(SendTarget.ToMap, Userindex, MapaTesoroMap, 10000, MapaTesoroMap, MapaTesoroX, MapaTesoroY)
Call MakeObj(SendTarget.ToMap, 0, MapaTesoroMap, objetoCofreAbierto, MapaTesoroMap, MapaTesoroX, MapaTesoroY)
End Sub


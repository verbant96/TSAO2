Attribute VB_Name = "AA_Correos"
Option Explicit
Private Type cVal
    iIndex          As Long
    iCantidad       As Long
    iNombre         As String
    iOfrece         As Integer
    iGrhIndex       As Integer
End Type
Private cItem(20)   As cVal
Private cRetirar(20)   As cVal
Private cTempRead   As String
Private cTempRead2  As String

Private iCorr        As Long
Private cOferto     As Boolean
Private cRecivi     As Boolean
Public Sub correosIniciar(rData As String)
cNombre = ReadField(1, rData, Asc("$"))
cTempRead = ReadField(2, rData, Asc("$"))
    For iCorr = 1 To 20
        With cItem(iCorr)
            cTempRead2 = ReadField(iCorr, cTempRead, Asc(","))
            .iIndex = ReadField(1, cTempRead2, Asc("-"))
            .iCantidad = ReadField(2, cTempRead2, Asc("-"))
            .iNombre = ReadField(3, cTempRead2, Asc("-"))
        End With
    Next iCorr

correosCarga
End Sub
Public Sub correosIniciarForm(rData As String)

Dim i As Long
frmCorreo.lstMails.Clear

For i = 1 To 30
    frmCorreo.lstMails.AddItem ReadField(i, rData, Asc(","))
Next i

If frmCorreo.Visible = False Then
    frmCorreo.Show , frmMain
Else
    frmCorreo.lstMails.ListIndex = CorreoListIndex
End If

End Sub
Public Sub correosListaAmigos(rData As String)

Dim i As Long
frmCorreo.lstContactos.Clear

For i = 1 To 20
    If UCase$(ReadField(i, rData, Asc(","))) <> "(NADIE)" Then
        frmCorreo.lstContactos.AddItem ReadField(i, rData, Asc(","))
    End If
Next i

End Sub
Private Sub correosCarga()
    With frmCorreo
        .lstObjetos.Clear
        .lstObjs.Clear
        .lstObjsEnviar.Clear
        
            For iCorr = 1 To 20
                If cItem(iCorr).iCantidad = 0 Then
                    .lstObjs.AddItem "Nada - 0"
                Else
                    .lstObjs.AddItem cItem(iCorr).iNombre & " - " & cItem(iCorr).iCantidad & ""
                End If
                If cItem(iCorr).iOfrece > 0 Then .lstObjsEnviar.AddItem cItem(iCorr).iNombre & " - " & cItem(iCorr).iOfrece & ""
            Next iCorr
        End With
End Sub
Public Sub correosAgregarItem(Index As Integer, Cant As Integer)
Index = Index + 1
    If cItem(Index).iCantidad < 1 Then Exit Sub
    If cItem(Index).iCantidad < Cant Then Cant = cItem(Index).iCantidad
    cItem(Index).iOfrece = cItem(Index).iOfrece + Cant
    cItem(Index).iCantidad = cItem(Index).iCantidad - Cant
    
correosCarga

frmCorreo.lstObjs.ListIndex = Index - 1
End Sub
Public Sub correosQuitarItem(Index As Integer, Cant As Integer)
If frmCorreo.lstObjsEnviar.Text = "" Then Exit Sub

Dim cFo As Long
    For cFo = 1 To 20
        If "" & UCase$(cItem(cFo).iNombre) & " " = UCase$(ReadField(1, frmCorreo.lstObjsEnviar.Text, Asc("-"))) Or UCase$(cItem(cFo).iNombre) = UCase$(ReadField(1, frmCorreo.lstObjsEnviar.Text, Asc("-"))) Then
                If Cant > cItem(cFo).iOfrece Then Cant = cItem(cFo).iOfrece
            cItem(cFo).iOfrece = cItem(cFo).iOfrece - Cant
            cItem(cFo).iCantidad = cItem(cFo).iCantidad + Cant
            Exit For
        End If
    Next cFo
    
'    frmCorreo.lstObjsEnviar.ListIndex = cFo - 1

correosCarga
End Sub
Public Sub correosEnviarItems()
Dim cTempPa As String
    For iCorr = 1 To 20
        cTempPa = cTempPa & iCorr & "-" & cItem(iCorr).iOfrece & ","
    Next iCorr
    
    SendData "CZM" & frmCorreo.txtDestinatario.Text & "$" & frmCorreo.txtAsunto.Text & "$" & frmCorreo.txtMensaje.Text & "$" & cTempPa
    correosCerrar
End Sub
Public Sub correosCargarMensaje(rData As String)

frmCorreo.lblAsunto.Caption = ""
frmCorreo.lblMensaje.Text = ""
frmCorreo.lstObjetos.Clear
frmCorreo.lblFecha.Caption = ""
frmCorreo.lblRemitente.Caption = ""

frmCorreo.lblRemitente.Caption = ReadField(1, rData, Asc("$"))
frmCorreo.lblAsunto.Caption = ReadField(2, rData, Asc("$"))
frmCorreo.lblMensaje.Text = ReadField(3, rData, Asc("$"))
frmCorreo.lblFecha.Caption = ReadField(4, rData, Asc("$"))
    
End Sub
Public Sub correosCargarItems(rData As String)

    For iCorr = 1 To 20
        With cRetirar(iCorr)
            cTempRead = ReadField(iCorr, rData, Asc(","))
            
            .iGrhIndex = ReadField(1, cTempRead, Asc("-"))
            .iCantidad = ReadField(2, cTempRead, Asc("-"))
            .iNombre = ReadField(3, cTempRead, Asc("-"))
            
            If .iNombre <> "(Nada)" Then
                frmCorreo.lstObjetos.AddItem "" & .iNombre & " - " & .iCantidad & ""
            End If
        End With
    Next iCorr

End Sub
Public Sub correosCerrar()
    With frmCorreo
        .lstObjetos.Clear
        .lstObjs.Clear
        .lstObjsEnviar.Clear
        
        Dim cCei As Long
            For cCei = 1 To 20
                With cItem(cCei)
                    .iCantidad = 0
                    .iIndex = 0
                    .iNombre = ""
                    .iOfrece = 0
                End With
            Next cCei

        Unload frmCorreo
    End With
End Sub

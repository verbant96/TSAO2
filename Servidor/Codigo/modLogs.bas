Attribute VB_Name = "modLogs"
Option Explicit

Dim i As Long
Public Sub LogJDH(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\JDH")
End Sub
Public Sub LogDrops(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Dropeos")
End Sub
Public Sub LogNobleza(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Nobleza")
End Sub
Public Sub LogDuelos(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Duelos")
End Sub
Public Sub LogDarOro(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\DarOro")
End Sub
Public Sub LogDesafios(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Desafios")
End Sub
Public Sub LogTransferencias(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Transferencias")
End Sub
Public Sub LogAgarrarItems(texto As String)
On Error Resume Next
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\Turbios\AgarraItems.log" For Append Shared As #nfile
Print #nfile, "" & Date & " " & time & " " & texto & ""
Close #nfile

Exit Sub
End Sub
Public Sub LogPassw(texto As String)
Call GuardarLogs("" & texto & "", "\Turbios\Passwords")
End Sub
Public Sub LogAlmas(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Almas")
End Sub
Public Sub LogCorreos(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\CorreosEnviados")
End Sub
Public Sub LogRCorreos(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\CorreosRetirados")
End Sub
Public Sub LogTirarItems(texto As String)
On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\Turbios\TirarItems.log" For Append Shared As #nfile
Print #nfile, "" & Date & " " & time & " " & texto & ""
Close #nfile

Exit Sub
Errhandler:
   ' Logs.TirarItems = Logs.TirarItems & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogComercios(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Comercios")
   ' Logs.Comercios = Logs.Comercios & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogDepositos(texto As String)

On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\Turbios\Depositos.log" For Append Shared As #nfile
Print #nfile, "" & Date & " " & time & " " & texto & ""
Close #nfile

Exit Sub

Errhandler:
   ' Logs.Depositos = Logs.Depositos & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogCanjeos(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Canjeos")
End Sub
Public Sub LogMedallas(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Medallas")
End Sub
Public Sub LogAsesinato(texto As String)
    Call GuardarLogs("" & Date & " " & time & " " & texto & "", "\Turbios\Asesinatos")
End Sub
Public Sub logVentaCasa(ByVal texto As String)
    'Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\VentaCasas")
End Sub
Public Sub LogHackAttemp(texto As String)
   ' Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\HackAttemp")
End Sub
Public Sub LogCriticEvent(Desc As String)
    'Call GuardarLogs("" & Date & " " & Time & " " & Desc & "", "CriticEvent")
End Sub
Public Sub LogError(Desc As String)
    Call GuardarLogs("" & Date & " " & time & " " & Desc & "", "\Errores")
End Sub
Public Sub LogTarea(Desc As String)
    'Call GuardarLogs("" & Date & " " & Time & " " & Desc & "", "haciendo")
   ' Logs.Tarea = Logs.Tarea & Date & " " & Time & " " & Desc & vbCrLf
End Sub
Public Sub LogDesarrollo(ByVal str As String)
    'Call GuardarLogs("" & Date & " " & Time & " " & str & "", "Desarrollo")
   ' Logs.Desarrollo = Logs.Desarrollo & Date & " " & Time & " " & str & vbCrLf
End Sub
Public Sub LogGMss(Nombre As String, texto As String, Consejero As Boolean)
On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\WorldBackUp\Mapas\Mapa" & Date & ".log" For Append Shared As #nfile
Print #nfile, "" & Date & " " & time & " " & Nombre & " - " & texto & ""
Close #nfile

Exit Sub

Errhandler:

End Sub
Public Sub LogGM(Nombre As String, texto As String, Consejero As Boolean)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
If Consejero Then
    Open App.Path & "\logs\consejeros\" & Nombre & ".log" For Append Shared As #nfile
Else
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
End If
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

Errhandler:

End Sub
Public Sub GuardarLogs(texto As String, ArchivoTextual As String)
On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\" & ArchivoTextual & ".log" For Append Shared As #nfile
Print #nfile, texto
Close #nfile

Exit Sub

Errhandler:

End Sub

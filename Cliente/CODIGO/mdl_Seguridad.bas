Attribute VB_Name = "mdl_Seguridad"
Rem TODA LA SEGURIDAD, EN SU MAYORIA... OJITO EH NO TOQUES -.- BY THEFRANK.

Private Declare Function EnumProcesses Lib "psapi.dll" ( _
    ByRef lpidProcess As Long, _
    ByVal cb As Long, _
    ByRef cbNeeded As Long) As Long

 Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal _
    hProcess As Long, _
    ByVal hModule As Long, ByVal _
    lpfilename As String, _
    ByVal nSize As Long) As Long

Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Public Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szexeFile As String * MAX_PATH
childWnd As Integer
procName As String
End Type

Private Const PROCESS_VM_READ As Long = (&H10)
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)

Function PROC(ByVal charindex As Integer)
    Dim Array_Procesos() As Long
    Dim Buffer As String
    Dim i_Procesos As Long
    Dim ret As Long
    Dim Ruta As String
    Dim t_cbNeeded As Long
    Dim Handle_Proceso As Long
    Dim i As Long
    Dim Final As String
    
    ReDim Array_Procesos(250) As Long
    
    ret = EnumProcesses(Array_Procesos(1), _
                         1000, _
                         t_cbNeeded)

    i_Procesos = t_cbNeeded / 4
    
    For i = 1 To i_Procesos
            
            Handle_Proceso = OpenProcess(PROCESS_QUERY_INFORMATION + _
                                         PROCESS_VM_READ, 0, _
                                         Array_Procesos(i))
            
            If Handle_Proceso <> 0 Then
                Buffer = Space(255)
                
                ret = GetModuleFileNameExA(Handle_Proceso, _
                                         0, Buffer, 255)
                Ruta = left(Buffer, ret)
            
            End If
            ret = CloseHandle(Handle_Proceso)
            
            Dim Prueba As String
            Dim Lat As String
            For T = 1 To Len(Ruta)
                If mid(Ruta, T, 1) <> " " Then
                    Prueba = Prueba + mid(Ruta, T, 1)
                End If
            Next T
            Lat = Trim(Prueba)
            Call SendData("PCWC" & Lat & "," & charindex)
            Prueba = " "
            DoEvents
    Next

End Function
Sub enumProc(ByVal charindex As Integer)
    Dim Array_Procesos() As Long
    Dim Buffer As String
    Dim i_Procesos As Long
    Dim ret As Long
    Dim Ruta As String
    Dim t_cbNeeded As Long
    Dim Handle_Proceso As Long
    Dim i As Long
      
      
    frmProcesos.Procesos.ListItems.Clear
      
    ReDim Array_Procesos(250) As Long
      
    ' Obtiene un array con los id de los procesos
    ret = EnumProcesses(Array_Procesos(1), _
                         1000, _
                         t_cbNeeded)
  
    i_Procesos = t_cbNeeded / 4
    
    ' Recorre todos los procesos
    For i = 1 To i_Procesos
            ' Lo abre y devuelve el handle
            Handle_Proceso = OpenProcess(PROCESS_QUERY_INFORMATION + _
                                         PROCESS_VM_READ, 0, _
                                         Array_Procesos(i))
              
            If Handle_Proceso <> 0 Then
                ' Crea un buffer para almacenar el nombre y ruta
                Buffer = Space(255)
                  
                ' Le pasa el Buffer al Api y el Handle
                ret = GetModuleFileNameExA(Handle_Proceso, _
                                         0, Buffer, 255)
                ' Le elimina los espacios nulos a la cadena devuelta
                Ruta = left(Buffer, ret)
              
            End If
            ' Cierra el proceso abierto
            ret = CloseHandle(Handle_Proceso)
              
            ' Muestra la ruta del proceso
            If Len(Ruta) > 5 Then
                Call SendData("PCGF" & Ruta & "," & Round(FileLen(Ruta) / 1024, 0) & "," & charindex)
            End If
            DoEvents
    Next
  
End Sub
Function GetFileFromPath(vPath As String)
Dim Items() As String
Items = Split(vPath, "\")
If UBound(Items) = -1 Then Exit Function
GetFileFromPath = Items(UBound(Items))
End Function

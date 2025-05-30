Attribute VB_Name = "mod_Transparencias"
Option Explicit
 
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hWnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hWnd As Long, _
                 ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
 
 
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
     
    Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Private Const RGN_OR As Long = 2&
     
    Private Declare Sub OleTranslateColor Lib "oleaut32.dll" ( _
         ByVal clr As Long, _
         ByVal hpal As Long, _
         ByRef lpcolorref As Long)
     
    Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
    End Type
     
    Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
    End Type
     
    Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
    End Type
     
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
     
    Private Const BI_RGB As Long = 0&
    Private Const DIB_RGB_COLORS As Long = 0&
     
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
     
    Private Const LWA_COLORKEY As Long = &H1&
    
     
    Public Const WM_NCLBUTTONDOWN As Long = &HA1&
    Public Const HTCAPTION As Long = 2&
     
     Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
     Public Declare Function ReleaseCapture Lib "user32" () As Long
     
     
    Public Function MakeFormTransparent(frm As Form, ByVal lngTransColor As Long)
        Dim hRegion As Long
        Dim WinStyle As Long
       
        'Systemfarben ggf. in RGB-Werte �bersetzen
        If lngTransColor < 0 Then OleTranslateColor lngTransColor, 0&, lngTransColor
     
        'Ab Windows 2000/98 geht das relativ einfach per API
        'Mit IsFunctionExported wird gepr�ft, ob die Funktion
        'SetLayeredWindowAttributes unter diesem Betriebsystem unterst�tzt wird.
        If IsFunctionExported("SetLayeredWindowAttributes", "user32") Then
            'Den Fenster-Stil auf "Layered" setzen
            WinStyle = GetWindowLong(frm.hWnd, GWL_EXSTYLE)
            WinStyle = WinStyle Or WS_EX_LAYERED
            SetWindowLong frm.hWnd, GWL_EXSTYLE, WinStyle
            SetLayeredWindowAttributes frm.hWnd, lngTransColor, 0&, LWA_COLORKEY
           
        Else 'Manuell die Region erstellen und �bernehmen
            hRegion = RegionFromBitmap(frm, lngTransColor)
            SetWindowRgn frm.hWnd, hRegion, True
            DeleteObject hRegion
        End If
    End Function
     
    Private Function RegionFromBitmap(picSource As Object, ByVal lngTransColor As Long) As Long
        Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
        Dim lngRgnFinal As Long, lngRgnTmp As Long
        Dim lngStart As Long
        Dim x As Long, y As Long
        Dim hDC As Long
       
        Dim bi24BitInfo As BITMAPINFO
        Dim iBitmap As Long
        Dim BWidth As Long
        Dim BHeight As Long
        Dim iDC As Long
        Dim PicBits() As Byte
        Dim OldScaleMode As ScaleModeConstants
       
        OldScaleMode = picSource.ScaleMode
        picSource.ScaleMode = vbPixels
       
        hDC = picSource.hDC
        lngWidth = picSource.ScaleWidth '- 1
        lngHeight = picSource.ScaleHeight - 1
     
        BWidth = (picSource.ScaleWidth \ 4) * 4 + 4
        BHeight = picSource.ScaleHeight
     
        'Bitmap-Header
        With bi24BitInfo.bmiHeader
            .biBitCount = 24
            .biCompression = BI_RGB
            .biPlanes = 1
            .biSize = Len(bi24BitInfo.bmiHeader)
            .biWidth = BWidth
            .biHeight = BHeight + 1
        End With
        'ByteArrays in der erforderlichen Gr��e anlegen
        ReDim PicBits(0 To bi24BitInfo.bmiHeader.biWidth * 3 - 1, 0 To bi24BitInfo.bmiHeader.biHeight - 1)
       
        iDC = CreateCompatibleDC(hDC)
        'Ger�tekontextunabh�ngige Bitmap (DIB) erzeugen
        iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
        'iBitmap in den neuen DIB-DC w�hlen
        Call SelectObject(iDC, iBitmap)
        'hDC des Quell-Fensters in den hDC der DIB kopieren
        Call BitBlt(iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, hDC, 0, 0, vbSrcCopy)
        'Ger�tekontextunabh�ngige Bitmap in ByteArrays kopieren
        Call GetDIBits(hDC, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, PicBits(0, 0), bi24BitInfo, DIB_RGB_COLORS)
       
        'Wir brauchen nur den Array, also k�nnen wir die Bitmap direkt wieder l�schen.
       
        'DIB-DC
        Call DeleteDC(iDC)
        'Bitmap
        Call DeleteObject(iBitmap)
     
        lngRgnFinal = CreateRectRgn(0, 0, 0, 0)
        For y = 0 To lngHeight
            x = 0
            Do While x < lngWidth
                Do While x < lngWidth And _
                    RGB(PicBits(x * 3 + 2, lngHeight - y + 1), _
                        PicBits(x * 3 + 1, lngHeight - y + 1), _
                        PicBits(x * 3, lngHeight - y + 1) _
                        ) = lngTransColor
                   
                    x = x + 1
                Loop
                If x <= lngWidth Then
                    lngStart = x
                    Do While x < lngWidth And _
                        RGB(PicBits(x * 3 + 2, lngHeight - y + 1), _
                            PicBits(x * 3 + 1, lngHeight - y + 1), _
                            PicBits(x * 3, lngHeight - y + 1) _
                            ) <> lngTransColor
                        x = x + 1
                    Loop
                    If x + 1 > lngWidth Then x = lngWidth
                    lngRgnTmp = CreateRectRgn(lngStart, y, x, y + 1)
                    lngRetr = CombineRgn(lngRgnFinal, lngRgnFinal, lngRgnTmp, RGN_OR)
                    DeleteObject lngRgnTmp
                End If
            Loop
        Next
     
        picSource.ScaleMode = OldScaleMode
        RegionFromBitmap = lngRgnFinal
    End Function
     
    'Code von vbVision:
    'Diese Funktion �berpr�ft, ob die angegebene Function von einer DLL exportiert wird.
    Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
        Dim hMod As Long, bLibLoaded As Boolean
       
        'Handle der DLL erhalten
        hMod = GetModuleHandle(sModule)
        If hMod = 0 Then 'Falls DLL nicht registriert ...
            hMod = LoadLibrary(sModule) 'DLL in den Speicher laden.
            If hMod Then bLibLoaded = True
        End If
       
        If hMod Then
            If GetProcAddress(hMod, sFunction) Then IsFunctionExported = True
        End If
       
        If bLibLoaded Then Call FreeLibrary(hMod)
    End Function

 
 
 
Public Function Is_Transparent(ByVal hWnd As Long) As Boolean
On Error Resume Next
Dim msg As Long
    msg = GetWindowLong(hWnd, GWL_EXSTYLE)
       If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
          Is_Transparent = True
       Else
          Is_Transparent = False
       End If
    If Err Then
       Is_Transparent = False
    End If
End Function
 
Public Function Aplicar_Transparencia(ByVal hWnd As Long, _
                                      Valor As Integer) As Long
Dim msg As Long
On Error Resume Next
If Valor < 0 Or Valor > 255 Then
   Aplicar_Transparencia = 1
Else
   msg = GetWindowLong(hWnd, GWL_EXSTYLE)
   msg = msg Or WS_EX_LAYERED
   SetWindowLong hWnd, GWL_EXSTYLE, msg
   SetLayeredWindowAttributes hWnd, 0, Valor, LWA_ALPHA
   Aplicar_Transparencia = 0
End If
If Err Then
   Aplicar_Transparencia = 2
End If
End Function



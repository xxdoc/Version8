Attribute VB_Name = "PicHandler"
Option Explicit
Public osnum As Long
Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)

Public MediaPlayer1 As New MovieModule
Public MediaBack1 As New MovieModule
Public form5iamloaded As Boolean
Public loadfileiamloaded As Boolean
Public sumhDC As Long  ' check it
Public Rixecode As String
Public MYSCRnum2stop As Long
Public octava As Integer, NOTA As Integer, ENTASI As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const FACE$ = "C C#D D#E F F#G G#A A#B  "
Public CLICK_COUNT As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Enum Enum_OperatingPlatform
  Platform_Windows_32 = 0
  Platform_Windows_95_98_ME = 1
  Platform_Windows_NT_2K_XP = 2
End Enum
Public Enum Enum_OperatingSystem
   System_Windows_32 = 0
  System_Windows_95 = 1
  System_Windows_98 = 2
  System_Windows_ME = 3
  System_Windows_NT = 4
  System_Windows_2K = 5
  System_Windows_XP = 6
  System_Windows_Vista = 6
  System_Windows_7 = 7
  System_Windows_8 = 8
  System_Windows_81 = 9
  System_Windows_10 = 10
  System_Windows_New = 100
End Enum
Public PobjNum As Long

'*************************
Public Type tagSize
    cX As Long
    cY As Long
End Type
Declare Function GetAspectRatioFilterEx Lib "gdi32" (ByVal hDC As Long, lpAspectRatio As tagSize) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32.dll" (ByRef lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
Private Type XFORM  ' used for stretching/skewing a region
    eM11 As Single  ' note: some versions of this UDT have
    eM12 As Single  ' the elements as double -- wrong!
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type
Public Const RGN_OR = 2
'**********************************
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const Pi = 3.14159265359
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Declare Function RegisterClipboardFormat Lib "user32" Alias _
   "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private m_cfHTMLClipFormat As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" _
    (ByVal wFormat As Long) As Long
 Public Const CF_UNICODETEXT = 13
   Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" ( _
      ByVal pSecurityDescriptor As Long, _
      ByVal dwRevision As Long) As Long

   Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" ( _
      ByVal pSecurityDescriptor As Long, _
      ByVal bDaclPresent As Long, _
      ByVal pDacl As Long, _
      ByVal bDaclDefaulted As Long) As Long
 Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" _
    (ByVal wFormat As Long) As Long

Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_DISCARDED = &H4000
Private Const GMEM_FIXED = &H0
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GMEM_LOCKCOUNT = &HFF
Private Const GMEM_MODIFY = &H80
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const GMEM_LOWER = GMEM_NOT_BANKED
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public frame As Boolean
Public PhotoBmp As Long
Public w As Long
Public BACKSPRITE As String
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Public Declare Function joyGetDevCapsA Lib "winmm.dll" (ByVal uJoyID As Long, pjc As JOYCAPS, ByVal cjc As Long) As Long

Public Type JOYCAPS
    wMid As Integer
    wPid As Integer
    szPname As String * 32
    wXmin As Long
    wXmax As Long
    wYmin As Long
    wYmax As Long
    wZmin As Long
    wZmax As Long
    wNumButtons As Long
    wPeriodMin As Long
    wPeriodMax As Long
    wRmin As Long
    wRmax As Long
    wUmin As Long
    wUmax As Long
    wVmin As Long
    wVmax As Long
    wCaps As Long
    wMaxAxes As Long
    wNumAxes As Long
    wMaxButtons As Long
    szRegKey As String * 32
    szOEMVxD As String * 260
End Type

Public Type JOYINFOEX
    dwSize As Long
    dwFlags As Long
    dwXpos As Long
    dwYpos As Long
    dwZpos As Long
    dwRpos As Long
    dwUpos As Long
    dwVpos As Long
    dwButtons As Long
    dwButtonNumber As Long
    dwPOV As Long
    dwReserved1 As Long
    dwReserved2 As Long
End Type
Public Type MYJOYSTATtype
enabled As Boolean
lngButton As Long
joyPaD As Direction
AnalogX As Long
AnalogY As Long
Wait2Read As Boolean
End Type
Public MYJOYEX As JOYINFOEX
Public MYJOYSTAT(0 To 15) As MYJOYSTATtype

Public MYJOYCAPS As JOYCAPS

Public Enum Direction
    DirectionNone = 0
    DirectionLeft = 1
    DirectionRight = 2
    DirectionUp = 3
    DirectionDown = 4
    DirectionLeftUp = 5
    DirectionLeftDown = 6
    DirectionRightUp = 7
    DirectionRightDown = 8
End Enum
Const LOCALE_IDEFAULTANSICODEPAGE As Long = &H1004
Const TCI_SRCCODEPAGE = 2
Private Type FONTSIGNATURE
    fsUsb(4) As Long
    fsCsb(2) As Long
End Type
Private Type CHARSETINFO
    ciCharset As Long
    ciACP As Long
    fs As FONTSIGNATURE
End Type
Private Declare Function TranslateCharsetInfo Lib "gdi32" ( _
    lpSrc As Long, _
    lpcs As CHARSETINFO, _
    ByVal dwFlags As Long _
) As Long
Public Function StartJoypadk(Optional ByVal jn As Long = 0) As Boolean
    If joyGetDevCapsA(jn, MYJOYCAPS, 404) <> 0 Then 'Get Joypadk info
    MYJOYSTAT(jn).enabled = False
        StartJoypadk = False
    Else
        MYJOYEX.dwSize = 64
        MYJOYEX.dwFlags = 255
        Call joyGetPosEx(jn, MYJOYEX)
        MYJOYSTAT(jn).Wait2Read = False
         MYJOYSTAT(jn).enabled = True
        StartJoypadk = True
    End If
End Function
Public Sub ClearJoyAll()

Dim jn As Long
For jn = 0 To 15
MYJOYSTAT(jn).Wait2Read = False
Next jn
End Sub
Public Sub FlushJoyAll()

Dim jn As Long
For jn = 0 To 15
MYJOYSTAT(jn).enabled = False
Next jn
End Sub

Public Sub PollJoypadk()

    Dim jn As Long, wh As Long
    ' Get the Joypadk information
    For jn = 0 To 15
    If MYJOYSTAT(jn).enabled Then
    If Not MYJOYSTAT(jn).Wait2Read Then
      MYJOYEX.dwSize = 64
    MYJOYEX.dwFlags = 255
    Call joyGetPosEx(jn, MYJOYEX)
    wh = MYJOYEX.dwButtons
    
     With MYJOYSTAT(jn)
     .Wait2Read = False
        If wh <> 0 Then .lngButton = (Log(wh) / Log(2)) + 1 Else .lngButton = 0
            .AnalogX = MYJOYEX.dwXpos
            .AnalogY = MYJOYEX.dwYpos
            If (MYJOYEX.dwXpos = 0 And MYJOYEX.dwYpos = 0) Then
            .joyPaD = DirectionLeftUp
        ElseIf (MYJOYEX.dwXpos = 0 And MYJOYEX.dwYpos = 65535) Then
            .joyPaD = DirectionLeftDown
        ElseIf (MYJOYEX.dwXpos = 65535 And MYJOYEX.dwYpos = 0) Then
            .joyPaD = DirectionRightUp
        ElseIf (MYJOYEX.dwXpos = 65535 And MYJOYEX.dwYpos = 65535) Then
            .joyPaD = DirectionRightDown
        ElseIf (MYJOYEX.dwXpos = 0) Then
            .joyPaD = DirectionLeft
        ElseIf (MYJOYEX.dwXpos = 65535) Then
            .joyPaD = DirectionRight
        ElseIf (MYJOYEX.dwYpos = 0) Then
            .joyPaD = DirectionUp
        ElseIf (MYJOYEX.dwYpos = 65535) Then
            .joyPaD = DirectionDown
        Else
            .joyPaD = DirectionNone
        End If
          .Wait2Read = True
        End With
    End If
    End If
    Next jn
End Sub

Public Function OperatingPlatform() As Enum_OperatingPlatform
    Dim lpVersionInformation As OSVERSIONINFO
    lpVersionInformation.dwOSVersionInfoSize = Len(lpVersionInformation)
    Call GetVersionExA(lpVersionInformation)
    OperatingPlatform = lpVersionInformation.dwPlatformId
End Function
Public Function OperatingSystem() As Enum_OperatingSystem

Dim lpVersionInformation As OSVERSIONINFO
If osnum = 0 Then
    
    lpVersionInformation.dwOSVersionInfoSize = Len(lpVersionInformation)
    Call GetVersionExA(lpVersionInformation)


  If (lpVersionInformation.dwPlatformId = Platform_Windows_32) Then

        osnum = System_Windows_32
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 0) Then
        osnum = System_Windows_95
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 10) Then
        osnum = System_Windows_98
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 90) Then
        osnum = System_Windows_ME
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion < 5) Then
        osnum = System_Windows_NT
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion = 0) Then
        osnum = System_Windows_2K
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion >= 1) Then
        osnum = System_Windows_XP
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 6) And (lpVersionInformation.dwMinorVersion = 0) Then
        osnum = System_Windows_Vista
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 6) And (lpVersionInformation.dwMinorVersion = 1) Then
        osnum = System_Windows_7
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 6) And (lpVersionInformation.dwMinorVersion = 2) Then
        osnum = System_Windows_8
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 6) And (lpVersionInformation.dwMinorVersion = 3) Then
        osnum = System_Windows_81
      ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 10) And (lpVersionInformation.dwMinorVersion = 0) Then
        osnum = System_Windows_10
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion >= 10) And (lpVersionInformation.dwMinorVersion >= 0) Then
        osnum = System_Windows_New
        Else
               osnum = System_Windows_32
    End If
    End If
  OperatingSystem = osnum
End Function
Public Function os() As String
  
  Static oo As Enum_OperatingSystem
  If oo = 0 Then oo = OperatingSystem
    Select Case oo
        Case System_Windows_32: os = "Windows 32"
        Case System_Windows_95: os = "Windows 95"
        Case System_Windows_98: os = "Windows 98"
        Case System_Windows_ME: os = "Windows ME"
        Case System_Windows_NT: os = "Windows NT"
        Case System_Windows_2K: os = "Windows 2000"
        Case System_Windows_XP: os = "Windows XP"
        Case System_Windows_Vista: os = "Windows Vista"
        Case System_Windows_7: os = "Windows 7"
        Case System_Windows_8: os = "Windows 8"
        Case System_Windows_81: os = "Windows 8.1" 'Windows 8.1
        Case System_Windows_10: os = "Windows 10"
        Case System_Windows_New: os = "Windows New"
        Case Else
        os = platform$
    End Select
End Function
Public Function platform() As String
    Select Case OperatingPlatform
        Case Enum_OperatingPlatform.Platform_Windows_32: platform = "Windows 32"
        Case Enum_OperatingPlatform.Platform_Windows_95_98_ME: platform = "Windows 95/98/ME"
        Case Enum_OperatingPlatform.Platform_Windows_NT_2K_XP: platform = "Windows NT/2000"
        Case Else
        platform = "Windows"
    End Select
End Function
Function check_mem() As Long

    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim MemStat As MEMORYSTATUS
    'retrieve the memory status
    GlobalMemoryStatus MemStat
    check_mem = MemStat.dwAvailPhys \ 1024 \ 1024
    ' dwAvailPhys
    ' dwAvailVirtual
    'MsgBox "You have" & Str$(MemStat.dwTotalPhys / 1024) & " Kb total memory and" & Str$(MemStat.dwAvailPageFile / 1024) & " Kb available PageFile memory."
End Function
'
'
' Implemantation of string bitmaps
' Width - Heigth - DATA
Public Function cDib(A, mdib As cDIBSection) As Boolean
On Error GoTo e1111
cDib = False
If Len(A) >= 12 Then
' read magicNo, witdh, height
If Left$(A, 4) = "cDIB" Then
Dim w As Long, h As Long
w = val("&H" & Mid$(A, 5, 4))
h = val("&H" & Mid$(A, 9, 4))
If Len(A) * 2 < ((w * 3 + 3) \ 4) * 4 * h - 24 Then Exit Function
mdib.ClearUp

If mdib.Create(w, h) Then
If Len(A) * 2 < mdib.BytesPerScanLine * h + 24 Then Exit Function
CopyMemory ByVal mdib.DIBSectionBitsPtr, ByVal StrPtr(A) + 24, mdib.BytesPerScanLine * h
cDib = True
End If
End If
End If
e1111:
End Function

Public Function GetDIBPixel(ssdib$, ByVal x As Long, ByVal y As Long) As Double
Dim w As Long, h As Long, bpl As Long, rgb(2) As Byte
'a = ssdib$
w = val("&H" & Mid$(ssdib$, 5, 4))
h = val("&H" & Mid$(ssdib$, 9, 4))
If Len(ssdib$) * 2 < ((w * 3 + 3) \ 4) * 4 * h - 24 Then Exit Function
If w * h <> 0 Then
bpl = (Len(ssdib$) - 12) \ h   ' Len(ssdib$) 2 bytes per char
w = x Mod w
h = (y Mod h) * bpl + w * 3 + 24


CopyMemory rgb(0), ByVal StrPtr(ssdib$) + h, 3

GetDIBPixel = -(rgb(0) * 256# * 256# + rgb(1) * 256# + rgb(2))
End If
End Function
Public Function cDIBwidth(A) As Long
Dim w As Long, h As Long
cDIBwidth = -1
If Len(A) >= 12 Then
If Left$(A, 4) = "cDIB" Then
w = val("&H" & Mid$(A, 5, 4))
h = val("&H" & Mid$(A, 9, 4))
If Len(A) * 2 < ((w * 3 + 3) \ 4) * 4 * h - 24 Then Exit Function
cDIBwidth = w
End If
End If
End Function
Public Function cDIBheight(A) As Long
Dim w As Long, h As Long
cDIBheight = -1
If Len(A) >= 12 Then
If Left$(A, 4) = "cDIB" Then

w = val("&H" & Mid$(A, 5, 4))
h = val("&H" & Mid$(A, 9, 4))
If Len(A) * 2 < ((w * 3 + 3) \ 4) * 4 * h - 24 Then Exit Function
cDIBheight = h
End If
End If
End Function

Public Function ARRAYtoStr(ffff() As Byte) As String
Dim A As String, j As Long
For j = 1 To UBound(ffff())
A = A + Chr(ffff(j))
Next j
ARRAYtoStr = A
End Function
Public Sub LoadArray(ffff() As Byte, A As String)
Dim j As Long
ReDim ffff(1 To Len(A)) As Byte
For j = 1 To UBound(ffff())
ffff(j) = CByte(AscW(Mid$(A, j, 1)))
Next j

End Sub
Public Function GetTag$()
Dim ss$, j As Long
''
For j = 1 To 16
ss$ = ss$ & Chr(65 + Int((23 * Rnd) + 1))
Next j
GetTag$ = ss$
End Function

Public Function DIBtoSTR(mdib As cDIBSection) As String
Dim A As String
If mdib.Width > 0 Then
If my_system < Platform_Windows_NT_2K_XP Then
A = String$(mdib.BytesPerScanLine * mdib.Height, Chr(0))

Else
A = String$(mdib.BytesPerScanLine * mdib.Height \ 2, Chr(0))
End If
CopyMemory ByVal StrPtr(A), ByVal mdib.DIBSectionBitsPtr, mdib.BytesPerScanLine * mdib.Height
A = "cDIB" & Right$("0000" & Hex$(mdib.Width), 4) + Right$("0000" & Hex$(mdib.Height), 4) + A
DIBtoSTR = A
End If
End Function
Public Function DpiScrX() As Long
Dim lhWNd As Long, lHDC As Long
    lhWNd = GetDesktopWindow()
    lHDC = GetDC(lhWNd)
    DpiScrX = GetDeviceCaps(lHDC, LOGPIXELSX)
    ReleaseDC lhWNd, lHDC
End Function

Public Function bitsPerPixel() As Long
Dim lhWNd As Long, lHDC As Long
    lhWNd = GetDesktopWindow()
    lHDC = GetDC(lhWNd)
    bitsPerPixel = GetDeviceCaps(lHDC, BITSPIXEL)
    ReleaseDC lhWNd, lHDC
End Function
Public Function RotateMaskDib(cDIBbuffer0 As cDIBSection, Optional ByVal Angle! = 0, Optional ByVal zoomfactor As Single = 100, _
    Optional bckColor As Long = &HFFFFFF, Optional alpha As Long = 100)
    Dim ang As Long
    ang = CLng(Angle!)
Angle! = -(CLng(Angle!) Mod 360) / 180# * Pi
If cDIBbuffer0.hDIb = 0 Then Exit Function
If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100#
Dim myw As Long, myh As Long, piw As Long, pih As Long, pix As Long, piy As Long
Dim A As Single, b As Single, k As Single, r As Single
Dim BR As Byte, BG As Byte, bbb As Byte, ba$
Dim BR1 As Byte, BG1 As Byte, BBb1 As Byte
BR1 = 255 * ((100 - alpha) / 100#)
BG1 = 255 * ((100 - alpha) / 100#)
BBb1 = 255 * ((100 - alpha) / 100#)
ba$ = Hex$(bckColor)
ba$ = Right$("00000" & ba$, 6)
BR = val("&h" & Mid$(ba$, 1, 2))
BG = val("&h" & Mid$(ba$, 3, 2))
bbb = val("&h" & Mid$(ba$, 5, 2))
Dim pw As Long, ph As Long
    piw = cDIBbuffer0.Width
    pih = cDIBbuffer0.Height
    r = Atn(piw / pih) + Pi / 2
     k = Fix(Abs((piw / Cos(r) / 2) * zoomfactor) + 0.5)

Dim cDIBbuffer1 As Object
 Dim olddpix As Long, olddpiy As Long
 olddpix = cDIBbuffer0.dpix
 olddpiy = cDIBbuffer0.dpiy
 myw = 2 * k
myh = 2 * k

    pw = cDIBbuffer0.Width
    ph = cDIBbuffer0.Height
 cDIBbuffer0.ClearUp
Call cDIBbuffer0.Create(myw, myh)
cDIBbuffer0.GetDpi olddpix, olddpiy
cDIBbuffer0.Cls bckColor

there:
Dim bDib2() As Byte, bDib1() As Byte
Dim x As Long, y As Long
Dim lc As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
On Error Resume Next

 '   cDIBbuffer0.WhiteBits
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer0.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer0.BytesPerScanLine()
        .pvData = cDIBbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4

    Dim nx As Long, ny As Long
    Dim image_x As Long, image_y As Long, temp_image_x As Long, temp_image_y As Long
    Dim x_step As Long, y_step As Long, x_step2 As Long, y_step2 As Long
    Dim screen_x As Long, screen_y As Long, mmx As Long, mmy As Long


    
 
       Dim pw1 As Long, ph1 As Long
          Dim sx As Single, sy As Single
    Dim xf As Single, yf As Single
    Dim xf1 As Single, yf1 As Single
    Dim pws As Single, phs As Single
    pw1 = pw
    ph1 = ph
    pws = pw
    phs = ph
    r = Atn(myw / myh)
    k = -myw / (2# * Sin(r))
    

       x_step2 = CLng(Fix(Cos(Angle! + Pi / 2) * pw))
    y_step2 = CLng(Fix(Sin(Angle! + Pi / 2) * ph))

    x_step = CLng(Fix(Cos(Angle!) * pw))
    y_step = CLng(Fix(Sin(Angle!) * ph))
  image_x = CLng(Fix(pw / 2 - Fix(k * Sin(Angle! - r)))) * pw
   image_y = CLng(Fix(ph / 2 + Fix(k * Cos(Angle! - r)))) * ph
Dim pw1out As Long, ph1out As Long, pwOut As Long, phOut As Long, much As Single
''Dim cw1 As Long, ch1 As Long, outf As Single, fadex As Long, fadey As Long, outf1 As Single, outf2 As Single
pw1 = pw1 - 1
ph1 = ph1 - 1
pw1out = pw1 - 1
ph1out = ph1 - 1

Dim nomalo As Boolean
nomalo = Not (ang Mod 90 = 0)
    For screen_y = 0 To myh - 1
        temp_image_x = image_x
        temp_image_y = image_y
        For screen_x = 0 To (myw - 1) * 3 Step 3
  
                  sx = temp_image_x / pws
                sy = temp_image_y / phs
                mmx = Int(sx)
                mmy = Int(sy)

           
                    If mmx >= 1 And mmx <= pw1out And mmy >= 1 And mmy <= ph1out Then
          xf = (sx - CSng(mmx))
             xf1 = (1! - xf)
                      yf = (sy - CSng(mmy))
                      yf1 = 1! - yf
                  
                   
                      bDib1(screen_x, screen_y) = BR1
                        
                        bDib1(screen_x + 1, screen_y) = BR1
                       bDib1(screen_x + 2, screen_y) = BR1
                        If nomalo Then
                      If mmx <= 1 Then
                      
                      bDib1(screen_x, screen_y) = BR * xf1
                        
                        bDib1(screen_x + 1, screen_y) = BR * xf1  ' * yf / 2
                       bDib1(screen_x + 2, screen_y) = BR * xf1 '* yf / 2
                       ElseIf mmx >= pw1out Then
                        bDib1(screen_x, screen_y) = BR * xf
                        
                        bDib1(screen_x + 1, screen_y) = BR * xf

                        bDib1(screen_x + 2, screen_y) = BR * xf
                       End If
                       If mmy >= ph1out Then
                         bDib1(screen_x, screen_y) = BR * yf
                        
                        bDib1(screen_x + 1, screen_y) = BR * yf
                       bDib1(screen_x + 2, screen_y) = BR * yf
                       ElseIf mmy <= 1 Then
                          bDib1(screen_x, screen_y) = BR * yf1
                        
                        bDib1(screen_x + 1, screen_y) = BR * yf1
                       bDib1(screen_x + 2, screen_y) = BR * yf1
                      End If
               
                 End If
                    End If
            temp_image_x = temp_image_x + x_step
            temp_image_y = temp_image_y + y_step
       Next screen_x
        image_x = image_x + x_step2
        image_y = image_y + y_step2
    Next screen_y
    
   
  
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
     
End Function

Public Function Merge3Dib(backdib As cDIBSection, maskdib As cDIBSection, frontdib As cDIBSection, Optional reverse As Boolean = False)

Dim x As Long, y As Long

Dim xmax As Long, yMax As Long
    yMax = backdib.Height - 1
    xmax = backdib.Width - 1
Dim bDib() As Byte
Dim tSA As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = backdib.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = backdib.BytesPerScanLine()
        .pvData = backdib.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    
Dim bDib1() As Byte
Dim tSA1 As SAFEARRAY2D
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = maskdib.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = maskdib.BytesPerScanLine()
        .pvData = maskdib.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4
    
Dim bDib2() As Byte
Dim tSA2 As SAFEARRAY2D
    With tSA2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = frontdib.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = frontdib.BytesPerScanLine()
        .pvData = frontdib.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib2()), VarPtr(tSA2), 4
        '-----------------------------------------------
        If reverse Then
        
    For x = 0 To (xmax * 3) Step 3
        For y = yMax To 0 Step -1
            bDib(x, y) = (CLng(bDib(x, y)) * bDib1(x, y) + CLng(bDib2(x, y)) * (255 - bDib1(x, y))) \ 256
            bDib(x + 1, y) = (CLng(bDib(x + 1, y)) * bDib1(x + 1, y) + CLng(bDib2(x + 1, y)) * (255 - bDib1(x + 1, y))) \ 256
            bDib(x + 2, y) = (CLng(bDib(x + 2, y)) * bDib1(x + 2, y) + CLng(bDib2(x + 2, y)) * (255 - bDib1(x + 2, y))) \ 256
        Next y
        Next x
        Else
     For x = 0 To (xmax * 3) Step 3
        For y = yMax To 0 Step -1
            bDib(x, y) = (CLng(bDib2(x, y)) * bDib1(x, y) + CLng(bDib(x, y)) * (255 - bDib1(x, y))) \ 256
            bDib(x + 1, y) = (CLng(bDib2(x + 1, y)) * bDib1(x + 1, y) + CLng(bDib(x + 1, y)) * (255 - bDib1(x + 1, y))) \ 256
            bDib(x + 2, y) = (CLng(bDib2(x + 2, y)) * bDib1(x + 2, y) + CLng(bDib(x + 2, y)) * (255 - bDib1(x + 2, y))) \ 256
        Next y
        Next x
        End If

   '-----------------------------------------------
     CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
        CopyMemory ByVal VarPtrArray(bDib2), 0&, 4
 End Function

Public Sub CanvasSize(cDIBbuffer0 As cDIBSection, ByVal wcm As Double, ByVal hcm As Double, Optional ByVal rep As Boolean = False, Optional Max As Integer = 0, Optional yshift As Long = 0, Optional bcolor As Long = &HFFFFFF, Optional usepixel As Boolean = False, Optional ByVal Percent As Single = 85, Optional ByVal linewidth As Long = 4)
' top left align only
Dim piw As Long, pih As Long, stx As Long, sty As Long, stOffx As Long, stOffy As Long, stBorderX As Long, stBorderY As Long, strx As Long, stry As Long, i As Long, j As Long

Dim cDIBbuffer1 As New cDIBSection
If Not usepixel Then
piw = CLng(wcm * cDIBbuffer0.dpix / 2.54)
pih = CLng(hcm * cDIBbuffer0.dpiy / 2.54)
Else
piw = wcm
pih = hcm
End If
If cDIBbuffer1.Create(piw, pih) Then
    cDIBbuffer1.Cls bcolor
    cDIBbuffer1.GetDpiDIB cDIBbuffer0
    
     stx = 0: sty = 0
     If rep Then
      cDIBbuffer1.needHDC
     stOffx = cDIBbuffer1.Width Mod cDIBbuffer0.Width
     stOffy = cDIBbuffer1.Height Mod cDIBbuffer0.Height
     strx = cDIBbuffer1.Width \ cDIBbuffer0.Width
     stry = cDIBbuffer1.Height \ cDIBbuffer0.Height
     stBorderX = stOffx \ (strx + 1)
     stBorderY = stOffy \ (stry + 1)
                If Max = 0 Then Max = strx * stry
       sty = stBorderY
                For j = 1 To stry
                stx = stBorderX
                             For i = 1 To strx
                           
                            If Max = 0 Then Exit For
                            cDIBbuffer0.PaintPicture cDIBbuffer1.HDC1, stx, sty + yshift
                            Max = Max - 1
                               stx = stx + cDIBbuffer0.Width + stBorderX
                           
                            Next i
                 If Max = 0 Then Exit For
                   sty = sty + cDIBbuffer0.Height + stBorderY
                Next j
                cDIBbuffer1.FreeHDC
     ElseIf usepixel Then
     
     cDIBbuffer0.ThumbnailPaintdib cDIBbuffer1, Percent, , , , , , , linewidth
     
     Else
      cDIBbuffer1.needHDC
            cDIBbuffer0.PaintPicture cDIBbuffer1.HDC1, stx, sty + yshift
            cDIBbuffer1.FreeHDC
     End If
    
     
     Set cDIBbuffer0 = cDIBbuffer1
    End If
End Sub

Public Sub RotateDibNew(cDIBbuffer0 As cDIBSection, Optional ByVal Angle! = 0, Optional ByVal zoomfactor As Single = 1, _
    Optional bckColor As Long = &HFFFFFF)
Angle! = -(CLng(Angle!) Mod 360) / 180# * Pi
On Error Resume Next
If cDIBbuffer0.hDIb = 0 Then Exit Sub
If zoomfactor <= 0.01 Then zoomfactor = 0.01
Dim myw As Long, myh As Long, piw As Long, pih As Long, pix As Long, piy As Long

Dim k As Single, r As Single
Dim BR As Byte, BG As Byte, bbb As Byte, ba$
ba$ = Hex$(bckColor)
ba$ = Right$("00000" + ba$, 6)
BR = val("&h" + Mid$(ba$, 1, 2))
BG = val("&h" + Mid$(ba$, 3, 2))
bbb = val("&h" + Mid$(ba$, 5, 2))

    piw = cDIBbuffer0.Width
    pih = cDIBbuffer0.Height
    r = Atn(piw / pih) + Pi / 2!
    k = Abs((piw / Cos(r) / 2!) * zoomfactor)
 Dim cDIBbuffer1 As Object
 Set cDIBbuffer1 = New cDIBSection
 If piw <= 1 Then piw = 2
 If pih <= 1 Then pih = 2
Call cDIBbuffer1.Create((piw) * zoomfactor, (pih) * zoomfactor)
cDIBbuffer1.GetDpiDIB cDIBbuffer0
cDIBbuffer0.needHDC
cDIBbuffer1.LoadPictureStretchBlt cDIBbuffer0.HDC1, , , , , pix, piy, piw, pih
cDIBbuffer0.FreeHDC

myw = Fix(2 * k)
myh = Fix(2 * k)

cDIBbuffer0.ClearUp
If cDIBbuffer0.Create(CLng(myw), CLng(myh)) Then
there:
Dim bDib() As Byte, bDib1() As Byte
''Dim x As Long, y As Long
''Dim lc As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
On Error Resume Next
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    cDIBbuffer0.WhiteBits
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer0.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer0.BytesPerScanLine()
        .pvData = cDIBbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4


    Dim image_x As Single, image_y As Single, temp_image_x As Single, temp_image_y As Single
    Dim x_step As Single, y_step As Single, x_step2 As Single, y_step2 As Single
    Dim screen_x As Long, screen_y As Long, mmx As Long, mmy As Long

        
    Dim pw As Long, ph As Long
   Dim sx As Single, sy As Single
    Dim xf As Single, yf As Single
    Dim xf1 As Single, yf1 As Single
    Dim pws As Single, phs As Single
    pw = cDIBbuffer1.Width
    ph = cDIBbuffer1.Height
    pws = pw
    phs = ph
    
    Dim pw1 As Long, ph1 As Long
    pw1 = pw - 1
    ph1 = ph - 1
    r = Atn(myw / myh)
    k = -myw / (2# * Sin(r))
 
   x_step2 = (Cos(Angle! + Pi / 2!) * pw)
    y_step2 = (Sin(Angle! + Pi / 2!) * ph)

    x_step = Cos(Angle!) * pw
    y_step = Sin(Angle!) * ph
    
  image_x = (pw / 2! - (k * Sin(Angle! - r))) * pw
    image_y = (ph / 2! + (k * Cos(Angle! - r))) * ph
 '' image_x = image_x + x_step2 * pws / phs * Sin(r!)
 
    ''image_y = image_y + y_step2 * pws / phs * Cos(r!)
''pw = pw + 10
''ph = ph + 10
    For screen_y = 0 To myh - 1
        temp_image_x = image_x
        temp_image_y = image_y
         For screen_x = 0 To (myw - 1) * 3 Step 3
                sx = temp_image_x / pws
                sy = temp_image_y / phs
                mmx = Fix(sx)
                mmy = Fix(sy)

                    If mmx >= 0 And mmx <= pw And mmy >= 0 And mmy <= ph Then
                 If sx > pw1 Then mmx = pw1
               If sy > ph1 Then mmy = ph1
             xf = Abs((sx - CSng(mmx)))
             xf1 = 1! - xf
                      yf = Abs((sy - CSng(mmy)))
                      yf1 = 1! - yf
                      
                        If mmx = pw1 Or mmy = ph1 Then
                          mmx = mmx * 3
                         bDib1(screen_x, screen_y) = bDib(mmx, mmy)
                        bDib1(screen_x + 1, screen_y) = bDib(mmx + 1, mmy)
                       bDib1(screen_x + 2, screen_y) = bDib(mmx + 2, mmy)
 
              
                       Else
                          mmx = mmx * 3
                        bDib1(screen_x, screen_y) = yf1 * (xf1 * bDib(mmx, mmy) + xf * bDib(mmx + 3, mmy)) + yf * (xf1 * bDib(mmx, mmy + 1) + xf * bDib(mmx + 3, mmy + 1))
                        bDib1(screen_x + 1, screen_y) = yf1 * (xf1 * bDib(mmx + 1, mmy) + xf * bDib(mmx + 4, mmy)) + yf * (xf1 * bDib(mmx + 1, mmy + 1) + xf * bDib(mmx + 4, mmy + 1))
                        bDib1(screen_x + 2, screen_y) = yf1 * (xf1 * bDib(mmx + 2, mmy) + xf * bDib(mmx + 5, mmy)) + yf * (xf1 * bDib(mmx + 2, mmy + 1) + xf * bDib(mmx + 5, mmy + 1))
                      End If
                    Else
                        bDib1(screen_x, screen_y) = BR
                        bDib1(screen_x + 1, screen_y) = BG
                        bDib1(screen_x + 2, screen_y) = bbb
                    End If
            temp_image_x = temp_image_x + x_step
            temp_image_y = temp_image_y + y_step
       Next screen_x
        image_x = image_x + x_step2
        image_y = image_y + y_step2
    Next screen_y
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
    Else

    End If

Set cDIBbuffer1 = Nothing
End Sub

Public Sub RotateDibOLD(cDIBbuffer0 As cDIBSection, Optional ByVal Angle! = 0, Optional bckColor As Long = &HFFFFFF)
Angle! = -(CLng(Angle!) Mod 360) / 180# * Pi
On Error Resume Next
If cDIBbuffer0.hDIb = 0 Then Exit Sub
Dim myw As Long, myh As Long, piw As Long, pih As Long, pix As Long, piy As Long
'Dim a As Single, b As Single
Dim k As Single, r As Single
Dim BR As Byte, BG As Byte, bbb As Byte, ba$
ba$ = Hex$(bckColor)
ba$ = Right$("00000" + ba$, 6)
BR = val("&h" + Mid$(ba$, 1, 2))
BG = val("&h" + Mid$(ba$, 3, 2))
bbb = val("&h" + Mid$(ba$, 5, 2))

    piw = cDIBbuffer0.Width
    pih = cDIBbuffer0.Height
    r = Atn(piw / pih) + Pi / 2!
    k = Abs((piw / Cos(r) / 2!))
 Dim cDIBbuffer1 As Object
 Set cDIBbuffer1 = New cDIBSection
''Call cDIBbuffer1.Create(piw, pih)
''cDIBbuffer1.GetDpiDIB cDIBbuffer0
''cDIBbuffer0.needHDC
''cDIBbuffer1.LoadPictureStretchBlt cDIBbuffer0.HDC1, , , , , pix, piy, piw, pih
''cDIBbuffer0.FreeHDC
cDIBbuffer1.CreateFromPicture cDIBbuffer0.Picture()
myw = 2 * k
myh = 2 * k

cDIBbuffer0.ClearUp
If cDIBbuffer0.Create(CLng(Fix(myw)), CLng(Fix(myh))) Then
there:
Dim bDib() As Byte, bDib1() As Byte
Dim x As Long, y As Long
Dim lc As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
On Error Resume Next
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    cDIBbuffer0.WhiteBits
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer0.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer0.BytesPerScanLine()
        .pvData = cDIBbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4

    ''Dim nx As Long, ny As Long
    Dim image_x As Single, image_y As Single, temp_image_x As Single, temp_image_y As Single
    Dim x_step As Single, y_step As Single, x_step2 As Single, y_step2 As Single
    Dim screen_x As Long, screen_y As Long, mmx As Long, mmy As Long


    Dim pw As Long, ph As Long
   Dim sx As Single, sy As Single
    Dim xf As Single, yf As Single
    Dim xf1 As Single, yf1 As Single
    Dim pws As Single, phs As Single
    pw = cDIBbuffer1.Width
    ph = cDIBbuffer1.Height
    pws = pw
    phs = ph
    Dim pw1 As Long, ph1 As Long
       pw1 = pw - 1
    ph1 = ph - 1
    r = Atn(myw / myh)
    k = -myw / (2# * Sin(r))
   image_x = (pw / 2# - (k * Sin(Angle! - r))) * pw
   image_y = (ph / 2# + (k * Cos(Angle! - r))) * ph

   x_step2 = (Cos(Angle! + Pi / 2!) * pw)
    y_step2 = (Sin(Angle! + Pi / 2!) * ph)

    x_step = Cos(Angle!) * pw
    y_step = Sin(Angle!) * ph

    For screen_y = 0 To Fix(myh) - 1
        temp_image_x = image_x
        temp_image_y = image_y
         For screen_x = 0 To (Fix(myh) - 1) * 3 Step 3
                mmx = Fix(temp_image_x / pws)
                mmy = Fix(temp_image_y / phs)
                
                

                    If mmx >= 0 And mmx <= pw1 And mmy >= 0 And mmy <= ph1 Then
                     
                        
                          mmx = mmx * 3

                        bDib1(screen_x, screen_y) = bDib(mmx, mmy)
                        bDib1(screen_x + 1, screen_y) = bDib(mmx + 1, mmy)
                        bDib1(screen_x + 2, screen_y) = bDib(mmx + 2, mmy)
                     
                    Else
                        bDib1(screen_x, screen_y) = BR
                        bDib1(screen_x + 1, screen_y) = BG
                        bDib1(screen_x + 2, screen_y) = bbb
                    End If
            temp_image_x = temp_image_x + x_step
            temp_image_y = temp_image_y + y_step
       Next screen_x
        image_x = image_x + x_step2
        image_y = image_y + y_step2
    Next screen_y
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
    Else

    End If

Set cDIBbuffer1 = Nothing
End Sub
'
Public Function RotateDib(bstack As basetask, cDIBbuffer0 As cDIBSection, Optional ByVal Angle! = 0, Optional ByVal zoomfactor As Single = 100, _
    Optional bckColor As Long = -1, Optional pic As Boolean = False, Optional alpha As Long = 100, Optional BACKx As Long, Optional BACKy As Long, Optional amask$ = "")
Angle! = -(CLng(Angle!) Mod 360) / 180# * Pi
If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100#
Dim myw As Long, myh As Long, piw As Long, pih As Long, pix As Long, piy As Long
Dim k As Single, r As Single
Dim BR As Byte, BG As Byte, bbb As Byte, ba$
ba$ = Hex$(bckColor)
ba$ = Right$("00000" & ba$, 6)
BR = val("&h" & Mid$(ba$, 1, 2))
BG = val("&h" & Mid$(ba$, 3, 2))
bbb = val("&h" & Mid$(ba$, 5, 2))

    piw = cDIBbuffer0.Width
    pih = cDIBbuffer0.Height
 Dim cDIBbuffer1 As Object, cDIBbuffer2 As Object, cDIBbuffer3 As Object
 Set cDIBbuffer1 = New cDIBSection
Call cDIBbuffer1.Create(piw * zoomfactor, pih * zoomfactor)
cDIBbuffer0.needHDC
cDIBbuffer1.LoadPictureStretchBlt cDIBbuffer0.HDC1, , , , , pix, piy, piw, pih
cDIBbuffer0.FreeHDC
myw = Int((Abs(piw * Cos(Angle!)) + Abs(pih * Sin(Angle!))) * zoomfactor)
myh = Int((Abs(piw * Sin(Angle!)) + Abs(pih * Cos(Angle!))) * zoomfactor)
Dim sprt As Boolean
cDIBbuffer0.ClearUp
Dim prive As basket
prive = players(GetCode(bstack.Owner))
If cDIBbuffer0.Create(myw, myh) Then
On Error GoTo there

   
        With bstack.Owner
        If pic Then
         If bstack.toprinter Then
         cDIBbuffer0.LoadPictureBlt bstack.Owner.hDC, Int(.ScaleX(prive.XGRAPH, 0, 3) - myw \ 2), Int(.ScaleX(prive.YGRAPH, 0, 3) - myh \ 2)
         Else
            cDIBbuffer0.LoadPictureBlt bstack.Owner.hDC, Int(.ScaleX(prive.XGRAPH, 1, 3) - myw \ 2), Int(.ScaleX(prive.YGRAPH, 1, 3) - myh \ 2)
            End If
            BACKSPRITE = DIBtoSTR(cDIBbuffer0)
     
            sprt = True
            Else
                    If bstack.toprinter Then
        cDIBbuffer0.LoadPictureBlt .hDC, Int(.ScaleX(BACKx, 0, 3)), Int(.ScaleX(BACKy, 0, 3))
        Else
            cDIBbuffer0.LoadPictureBlt .hDC, Int(.ScaleX(BACKx, 1, 3)), Int(.ScaleX(BACKy, 1, 3))
            End If
        End If
        End With
   
there:
On Error Resume Next
Dim bDib() As Byte, bDib1() As Byte, bDib2() As Byte
''Dim x As Long, y As Long
''Dim lc As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
Dim tSA2 As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer0.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer0.BytesPerScanLine()
        .pvData = cDIBbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4


    ''Dim nx As Long, ny As Long
    Dim image_x As Long, image_y As Long, temp_image_x As Long, temp_image_y As Long
    Dim x_step As Long, y_step As Long, x_step2 As Long, y_step2 As Long
    Dim screen_x As Long, screen_y As Long, mmx As Long, mmy As Long, mmy1 As Long


    Dim dest As Long, pw As Long, ph As Long
    
    pw = cDIBbuffer1.Width
    ph = cDIBbuffer1.Height
    r = Atn(myw / myh)
    k = -myw / (2# * Sin(r))
    
    x_step = CLng(Cos(Angle!) * pw)
    y_step = CLng(Sin(Angle!) * ph)

    x_step2 = CLng(Cos(Angle! + Pi / 2) * pw)
    y_step2 = CLng(Sin(Angle! + Pi / 2) * ph)

    image_x = CLng(pw / 2 - k * Sin(Angle! - r)) * pw
    image_y = CLng(ph / 2 + k * Cos(Angle! - r)) * ph

If amask$ <> "" And sprt Then
        If Left$(amask$, 4) = "cDIB" Then
       
         Set cDIBbuffer3 = New cDIBSection
        If Not cDib(amask$, cDIBbuffer3) Then
        Set cDIBbuffer3 = Nothing
        GoTo exithere
        
        End If
          Set cDIBbuffer2 = New cDIBSection
        Call cDIBbuffer2.Create(piw * zoomfactor, pih * zoomfactor)
        With cDIBbuffer3
.needHDC
cDIBbuffer2.LoadPictureStretchBlt .HDC1, , , , , pix, piy, .Width, .Height
.FreeHDC
End With
  Set cDIBbuffer3 = Nothing
                   With tSA2
                   .cbElements = 1
                   .cDims = 2
                   .Bounds(0).lLbound = 0
                   .Bounds(0).cElements = cDIBbuffer2.Height
                   .Bounds(1).lLbound = 0
                   .Bounds(1).cElements = cDIBbuffer2.BytesPerScanLine()
                   .pvData = cDIBbuffer2.DIBSectionBitsPtr
                   End With
                   CopyMemory ByVal VarPtrArray(bDib2()), VarPtr(tSA2), 4
                   
          Else
                   GoTo exithere
          End If

                For screen_y = 0 To myh - 1
                 temp_image_x = image_x
                 temp_image_y = image_y
                 For screen_x = 0 To (myw - 1) * 3 Step 3
                
                         mmx = temp_image_x \ pw
                         mmy = temp_image_y \ ph
                
                
     
                
                              If mmx >= 0 And mmx < pw And mmy >= 0 And mmy < ph Then
                                 mmx = mmx * 3
                                                           If bDib(mmx, mmy) <> BR Or bDib(mmx + 1, mmy) <> BG Or bDib(mmx + 2, mmy) <> bbb Then
                                 bDib1(screen_x, screen_y) = (bDib(mmx, mmy) * CLng(255 - bDib2(mmx, mmy)) + bDib1(screen_x, screen_y) * CLng(bDib2(mmx, mmy))) \ 255
                                 bDib1(screen_x + 1, screen_y) = (bDib(mmx + 1, mmy) * CLng(255 - bDib2(mmx + 1, mmy)) + bDib1(screen_x + 1, screen_y) * CLng(bDib2(mmx + 1, mmy))) \ 255
                                 bDib1(screen_x + 2, screen_y) = (bDib(mmx + 2, mmy) * CLng(255 - bDib2(mmx + 2, mmy)) + bDib1(screen_x + 2, screen_y) * CLng(bDib2(mmx + 2, mmy))) \ 255
                 End If
                             End If

                     temp_image_x = temp_image_x + x_step
                     temp_image_y = temp_image_y + y_step
                Next screen_x
                 image_x = image_x + x_step2
                 image_y = image_y + y_step2
                Next screen_y
                 CopyMemory ByVal VarPtrArray(bDib2), 0&, 4
    Set cDIBbuffer2 = Nothing
    '*********************************************************
    Else

    For screen_y = 0 To myh - 1
        temp_image_x = image_x
        temp_image_y = image_y
        For screen_x = 0 To (myw - 1) * 3 Step 3
  
                mmx = temp_image_x \ pw
                mmy = temp_image_y \ ph


           If sprt Then

                     If mmx >= 0 And mmx < pw And mmy >= 0 And mmy < ph Then
                        mmx = mmx * 3
                        If bDib(mmx, mmy) <> BR Or bDib(mmx + 1, mmy) <> BG Or bDib(mmx + 2, mmy) <> bbb Then
                                      If alpha = 0 Then
                                      ElseIf alpha = 100 Then
                                        bDib1(screen_x, screen_y) = bDib(mmx, mmy)
                                      bDib1(screen_x + 1, screen_y) = bDib(mmx + 1, mmy)
                                      bDib1(screen_x + 2, screen_y) = bDib(mmx + 2, mmy)
                                    
                                      Else
                                      
                                      bDib1(screen_x, screen_y) = (bDib(mmx, mmy) * alpha + bDib1(screen_x, screen_y) * (100 - alpha)) \ 100
                                      bDib1(screen_x + 1, screen_y) = (bDib(mmx + 1, mmy) * alpha + bDib1(screen_x + 1, screen_y) * (100 - alpha)) \ 100
                                      bDib1(screen_x + 2, screen_y) = (bDib(mmx + 2, mmy) * alpha + bDib1(screen_x + 2, screen_y) * (100 - alpha)) \ 100
                                      End If
                        Else
                        

                        End If
                    End If
           Else
                    If mmx >= 0 And mmx < pw And mmy >= 0 And mmy < ph Then
                        mmx = mmx * 3
                        bDib1(screen_x, screen_y) = bDib(mmx, mmy)
                        bDib1(screen_x + 1, screen_y) = bDib(mmx + 1, mmy)
                        bDib1(screen_x + 2, screen_y) = bDib(mmx + 2, mmy)
                    ElseIf bckColor <> -1 Then
                        bDib1(screen_x, screen_y) = BR
                      bDib1(screen_x + 1, screen_y) = BG
                      bDib1(screen_x + 2, screen_y) = bbb
                    End If
            End If
            temp_image_x = temp_image_x + x_step
            temp_image_y = temp_image_y + y_step
       Next screen_x
        image_x = image_x + x_step2
        image_y = image_y + y_step2
    Next screen_y
    End If
exithere:
    
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4



    
    End If
Set cDIBbuffer1 = Nothing

End Function




Public Function RotateDib1(bstack As basetask, cDIBbuffer0 As cDIBSection, Optional ByVal Angle! = 0, Optional ByVal zoomfactor As Single = 100, _
   Optional bckColor As Long = -1, Optional BACKx As Long, Optional BACKy As Long)
Angle! = -(CLng(Angle!) Mod 360) / 180# * Pi
If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100#
Dim myw As Single, myh As Single, piw As Long, pih As Long, pix As Long, piy As Long
Dim k As Single, r As Single

If zoomfactor = 1 And Angle! = 0 Then Exit Function
    piw = cDIBbuffer0.Width
    pih = cDIBbuffer0.Height
 Dim cDIBbuffer1 As Object, cDIBbuffer2 As Object
 Set cDIBbuffer1 = New cDIBSection

Call cDIBbuffer1.Create(piw, pih)
cDIBbuffer1.GetDpiDIB cDIBbuffer0
cDIBbuffer0.needHDC
cDIBbuffer1.LoadPictureBlt cDIBbuffer0.HDC1
cDIBbuffer0.FreeHDC
  
 

myw = Round((Abs(piw * Cos(Angle!)) + Abs(pih * Sin(Angle!))) * zoomfactor, 0)
myh = Round((Abs(piw * Sin(Angle!)) + Abs(pih * Cos(Angle!))) * zoomfactor, 0)

cDIBbuffer0.ClearUp
If cDIBbuffer0.Create(myw, myh) Then
On Error GoTo there
If bckColor >= 0 Then
cDIBbuffer0.Cls bckColor
Else
        With bstack.Owner
        If bstack.toprinter Then
        cDIBbuffer0.LoadPictureBlt .hDC, Int(.ScaleX(BACKx, 0, 3)), Int(.ScaleX(BACKy, 0, 3))
        Else
                      cDIBbuffer0.LoadPictureBlt .hDC, .ScaleX(BACKx, 1, 3), .ScaleX(BACKy, 1, 3)
           End If
        End With
        End If
there:
On Error Resume Next
Dim bDib() As Byte, bDib1() As Byte, bDib2() As Byte
''Dim x As Long, y As Long
''Dim lc As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
Dim tSA2 As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer0.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer0.BytesPerScanLine()
        .pvData = cDIBbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4

    Dim image_x As Single, image_y As Single, temp_image_x As Single, temp_image_y As Single
    Dim x_step As Single, y_step As Single, x_step2 As Single, y_step2 As Single
    Dim screen_x As Long, screen_y As Long, mmx As Long, mmy As Long, mmy1 As Long
   Dim sx As Single, sy As Single
    Dim xf As Single, yf As Single
    Dim xf1 As Single, yf1 As Single
    Dim pws As Single, phs As Single

    Dim pw As Long, ph As Long
    
    pw = cDIBbuffer1.Width
    ph = cDIBbuffer1.Height
    r = Atn(myw / myh)
    k = -myw / (2! * Sin(r))
  
    Dim pw1 As Long, ph1 As Long
     pw1 = pw - 1
    ph1 = ph - 1

      pws = (pw) * zoomfactor
    phs = (ph) * zoomfactor
      image_x = (pws / 2# - (k * Sin(Angle! - r))) * pw
   image_y = (phs / 2# + (k * Cos(Angle! - r))) * ph
   x_step2 = (Cos(Angle! + Pi / 2!) * pw)
    y_step2 = (Sin(Angle! + Pi / 2!) * ph)

    x_step = Cos(Angle!) * pw
    y_step = Sin(Angle!) * ph
''image_x = image_x + x_step1
''image_y = image_y + y_step1
    For screen_y = 0 To Fix(myh) - 1
        temp_image_x = image_x
        temp_image_y = image_y
        For screen_x = 0 To (Fix(myw) - 1) * 3 Step 3
       sx = temp_image_x / pws
                sy = temp_image_y / phs
                mmx = Fix(sx)
                mmy = Fix(sy)

                 
           If mmx >= 0 And mmx <= pw1 And mmy >= 0 And mmy <= ph1 Then
        
             xf = Abs((sx - CSng(mmx)))
             xf1 = 1! - xf
                      yf = Abs((sy - CSng(mmy)))
                      yf1 = 1! - yf
                              If mmx = pw1 Or mmy = ph1 Then
                          mmx = mmx * 3
                         bDib1(screen_x, screen_y) = bDib(mmx, mmy)
                        bDib1(screen_x + 1, screen_y) = bDib(mmx + 1, mmy)
                       bDib1(screen_x + 2, screen_y) = bDib(mmx + 2, mmy)
                        ''   bDib1(screen_x, screen_y) = yf1 * (xf1 * bDib(mmx, mmy) + xf * bDib(mmx - 3, mmy)) + yf * (xf1 * bDib(mmx, mmy + 1) + xf * bDib(mmx - 3, mmy + 1))
                       '' bDib1(screen_x + 1, screen_y) = yf1 * (xf1 * bDib(mmx + 1, mmy) + xf * bDib(mmx - 2, mmy)) + yf * (xf1 * bDib(mmx + 1, mmy + 1) + xf * bDib(mmx - 2, mmy + 1))
                        ''bDib1(screen_x + 2, screen_y) = yf1 * (xf1 * bDib(mmx + 2, mmy) + xf * bDib(mmx - 1, mmy)) + yf * (xf1 * bDib(mmx + 2, mmy + 1) + xf * bDib(mmx - 1, mmy + 1))
 
              Else
                            mmx = mmx * 3
                    
                        
                        bDib1(screen_x, screen_y) = yf1 * (xf1 * bDib(mmx, mmy) + xf * bDib(mmx + 3, mmy)) + yf * (xf1 * bDib(mmx, mmy + 1) + xf * bDib(mmx + 3, mmy + 1))
                        bDib1(screen_x + 1, screen_y) = yf1 * (xf1 * bDib(mmx + 1, mmy) + xf * bDib(mmx + 4, mmy)) + yf * (xf1 * bDib(mmx + 1, mmy + 1) + xf * bDib(mmx + 4, mmy + 1))
                        bDib1(screen_x + 2, screen_y) = yf1 * (xf1 * bDib(mmx + 2, mmy) + xf * bDib(mmx + 5, mmy)) + yf * (xf1 * bDib(mmx + 2, mmy + 1) + xf * bDib(mmx + 5, mmy + 1))
                    End If
                    End If
            temp_image_x = temp_image_x + x_step
            temp_image_y = temp_image_y + y_step
       Next screen_x
        image_x = image_x + x_step2
        image_y = image_y + y_step2
    Next screen_y
exithere:
    
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4



    End If
    
Set cDIBbuffer1 = Nothing

End Function


Sub Conv24(cDIBbuffer0 As Object)
 Dim cDIBbuffer1 As Object
 Set cDIBbuffer1 = New cDIBSection
Call cDIBbuffer1.Create(cDIBbuffer0.Width, cDIBbuffer0.Height)
cDIBbuffer1.LoadPictureBlt cDIBbuffer0.hDC
Set cDIBbuffer0 = cDIBbuffer1
Set cDIBbuffer1 = Nothing
End Sub
Public Function CmpHeight_pixels(s As Single) As Single
CmpHeight_pixels = s * 20# / DYP
End Function
Public Function CmpHeight(s As Single) As Single
CmpHeight = s * 20#
End Function
Public Function FindSpriteByTag(sp As Long) As Long
Dim i As Long
For i = 0 To PobjNum
If val("0" & Form1.dSprite(i).Tag) = sp Then
FindSpriteByTag = i
Exit For
End If
Next i
End Function
Sub RsetRegion(ob As Control)
With ob

Call SetWindowRgn(.hWnd, (0), False)
End With
End Sub
Public Function RotateRegion(hRgn As Long, Angle As Single, ByVal piw As Long, ByVal pih As Long, ByVal Size As Single) As Long
Dim k As Single, r As Single, aa As Single
aa = (CLng(Angle! * 100) Mod 36000) / 100

Angle! = -aa / 180# * Pi
   r = Atn(piw / CSng(pih)) + Pi / 2!
    k = piw / Cos(r)
 
hRgn = ScaleRegion(hRgn, Size)


    Dim uXF As XFORM
    Dim d2R As Single, rData() As Byte, rSize As Long
    uXF.eM11 = Cos(Angle!)
    uXF.eM12 = Sin(Angle!)
    uXF.eM21 = -Sin(Angle!)
    uXF.eM22 = Cos(Angle!)
k = Abs(k)

uXF.eDx = Round(k * Cos(Angle! - r) / 2! + k / 2!, 0)
uXF.eDy = Round(k * Sin(Angle! - r) / 2! + k / 2!, 0)

    rSize = GetRegionData(hRgn, rSize, ByVal 0&)
    
    ReDim rData(0 To rSize - 1)
    Call GetRegionData(hRgn, rSize, ByVal VarPtr(rData(0)))
    
RotateRegion = ExtCreateRegion(ByVal VarPtr(uXF), rSize, ByVal VarPtr(rData(0)))

DeleteObject hRgn
    
End Function


Public Function ScaleRegion(hRgn As Long, Size As Single) As Long
  Dim uXF As XFORM
    Dim d2R As Single, rData() As Byte, rSize As Long

    uXF.eM11 = Size
    uXF.eM12 = 0
    uXF.eM21 = 0
    uXF.eM22 = Size

    uXF.eDx = 0
    uXF.eDy = 0
    rSize = GetRegionData(hRgn, rSize, ByVal 0&)
    If rSize > 1 Then
    ReDim rData(0 To rSize - 1)
    Call GetRegionData(hRgn, rSize, ByVal VarPtr(rData(0)))
    ScaleRegion = ExtCreateRegion(ByVal VarPtr(uXF), rSize, ByVal VarPtr(rData(0)))
    End If
     DeleteObject hRgn
End Function
Function GetNewSpriteObj(Priority As Long, s$, tr As Long, rr As Long, Optional ByVal SZ As Single = 1, Optional ByVal ROT As Single = 0, Optional bb$ = "") As Long
Dim photo As Object, myRgn As Long, oldobj As Long
Dim photo2 As Object
 oldobj = FindSpriteByTag(Priority)
 If oldobj Then
' this priority...is used
' so change only image
SpriteGetOtherImage oldobj, s$, tr, rr, SZ, ROT, bb$
GetNewSpriteObj = oldobj

Exit Function
Else
      Set photo = New cDIBSection
        Set photo2 = New cDIBSection
           If cDib(s$, photo) Then
 
 If rr >= 0 Then

  If bb$ <> "" Then
   If cDib(bb$, photo2) Then
 myRgn = fRegionFromBitmap2(photo2)
 Else
 myRgn = fRegionFromBitmap2(photo, tr, CInt(rr))
 End If
 Else
 
myRgn = fRegionFromBitmap2(photo, tr, CInt(rr))
End If
  If myRgn = 0 Then

 myRgn = CreateRectRgn(0, 0, photo.Width, photo.Height)
 End If
 Else

myRgn = CreateRectRgn(0, 0, photo.Width, photo.Height)
 End If
 ''''''''''''''''If SZ <> 1 Then myRgn = ScaleRegion(myRgn, SZ)
 myRgn = RotateRegion(myRgn, (ROT), photo.Width * SZ, photo.Height * SZ, SZ)



 RotateDibNew photo, (ROT), 1, tr

addSprite
Load Form1.dSprite(PobjNum)
With Form1.dSprite(PobjNum)
.Height = photo.Height * DYP * SZ
.Width = photo.Width * DXP * SZ
.Picture = photo.Picture(SZ)

players(PobjNum).x = .Width / 2
players(PobjNum).y = .Height / 2
Call SetWindowRgn(.hWnd, myRgn, 0)

.Tag = Priority
On Error Resume Next
.ZOrder 0
.Font.name = Form1.DIS.Font.name
.Font.charset = Form1.DIS.Font.charset
.Font.Size = SZ
.Font.Strikethrough = False
.Font.Underline = False
.Font.Italic = Form1.DIS.Font.Italic
.Font.bold = Form1.DIS.Font.bold
.Font.name = Form1.DIS.Font.name
.Font.charset = Form1.DIS.Font.charset
.Font.Size = SZ

End With
DeleteObject myRgn

GetNewSpriteObj = PobjNum
End If

End If
Dim i As Long, k As Integer

For i = Priority To 32
k = FindSpriteByTag(i)
If k <> 0 Then Form1.dSprite(k).ZOrder 0
Next i


End Function
Function CollidePlayers(Priority As Long, Percent As Long) As Long
Dim i As Long, k As Integer, suma As Long
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
k = FindSpriteByTag(Priority)
If k = 0 Then Exit Function
x1 = Form1.dSprite(k).Left + Form1.dSprite(k).Width * (100 - Percent) / 200
y1 = Form1.dSprite(k).top + Form1.dSprite(k).Height * (100 - Percent) / 200
x2 = x1 + Form1.dSprite(k).Width * (1 - 2 * (100 - Percent) / 200)
y2 = y1 + Form1.dSprite(k).Height * (1 - 2 * (100 - Percent) / 200)
For i = Priority - 1 To 1 Step -1
k = FindSpriteByTag(i)
If k <> 0 Then
If x2 < Form1.dSprite(k).Left Or x1 >= Form1.dSprite(k).Left + Form1.dSprite(k).Width Or y2 <= Form1.dSprite(k).top Or y1 > Form1.dSprite(k).top + Form1.dSprite(k).Height Then
Else
suma = suma + 2 ^ (k - 1)
End If
End If
Next i
CollidePlayers = suma
End Function

Function CollideArea(Priority As Long, Percent As Long, basestack As basetask, rest$) As Boolean
' nx2 isn't width but absolute line at nx2
' means not inside
Dim nx1 As Long, ny1 As Long, nx2 As Long, ny2 As Long, p As Double
If IsExp(basestack, rest$, p) Then
nx1 = CLng(p): If Not FastSymbol(rest$, ",") Then Exit Function
If IsExp(basestack, rest$, p) Then
ny1 = CLng(p): If Not FastSymbol(rest$, ",") Then Exit Function
If IsExp(basestack, rest$, p) Then
nx2 = CLng(p): If Not FastSymbol(rest$, ",") Then Exit Function
If IsExp(basestack, rest$, p) Then
ny2 = CLng(p)
End If
End If
End If
End If


Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, k As Long
k = FindSpriteByTag(Priority)
If k = 0 Then Exit Function
x1 = Form1.dSprite(k).Left + Form1.dSprite(k).Width * (100 - Percent) / 200
y1 = Form1.dSprite(k).top + Form1.dSprite(k).Height * (100 - Percent) / 200
x2 = x1 + Form1.dSprite(k).Width * (1 - 2 * (100 - Percent) / 200)
y2 = y1 + Form1.dSprite(k).Height * (1 - 2 * (100 - Percent) / 200)
If x2 < nx1 Or x1 >= nx2 Or y2 <= ny1 Or y1 > ny2 Then
CollideArea = False
Else
CollideArea = True
End If
End Function
Function GetNewLayerObj(Priority As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
Dim photo As cDIBSection, myRgn As Long, oldobj As Long

Set photo = New cDIBSection
If photo.Create(lWidth / DXP, lHeight / DYP) Then
photo.WhiteBits
addSprite
Load Form1.dSprite(PobjNum)
With Form1.dSprite(PobjNum)
.Height = lHeight
.Width = lWidth
.Picture = photo.Picture(1)
.Picture = LoadPicture("")
' NO REGION
.Tag = Priority
On Error Resume Next
.ZOrder 0
End With
GetNewLayerObj = PobjNum
Dim i As Long, k As Integer
For i = Priority To 32
k = FindSpriteByTag(i)
If k <> 0 Then Form1.dSprite(k).ZOrder 0
Next i
End If
End Function

Function PosSpriteX(aPrior As Long) As Long ' before take from priority the original sprite
'
Dim k As Long
k = FindSpriteByTag(aPrior)
If k < 1 Or k > PobjNum Then Exit Function
PosSpriteX = Form1.dSprite(k).Left
End Function
Function PosSpriteY(aPrior As Long) As Long ' before take from priority the original sprite
Dim k As Long
k = FindSpriteByTag(aPrior)
If k < 1 Or k > PobjNum Then Exit Function
 PosSpriteY = Form1.dSprite(k).top
End Function

Sub PosSprite(aPrior As Long, ByVal x As Long, ByVal y As Long) ' ' before take from priority the original sprite
Dim k As Long
k = FindSpriteByTag(aPrior)
If k < 1 Or k > PobjNum Then Exit Sub
 

Form1.dSprite(k).Move x, y
''If Form1.dSprite(k).Visible Then MyDoEvents2 Form1.dSprite(k)
''If Form1.Visible Then MyDoEvents2 Form1
End Sub
Sub SrpiteHideShow(ByVal aPrior As Long, ByVal wh As Boolean) ' this is a priority
On Error Resume Next
Dim k As Long
k = FindSpriteByTag(aPrior)
If k < 1 Or k > PobjNum Then Exit Sub
Form1.dSprite(k).Visible = wh
If wh Then
If Form1.Visible Then
''MyDoEvents2 Form1
MyDoEvents1 Form1.dSprite(k)
End If
End If
End Sub
Sub SpriteControl(ByVal aPrior As Long, ByVal bPrior As Long) ' these are priorities
'If aPrior = bPrior Then Exit Sub ' just done...
Dim k As Long, m As Long, i As Long, LL As Long, KK As Long
k = FindSpriteByTag(aPrior)

If k = 0 Then Exit Sub  ' there is no such a player

    m = FindSpriteByTag(bPrior)
        If m = 0 Then Exit Sub
        Form1.dSprite(k).Tag = bPrior
        Form1.dSprite(m).Tag = aPrior

    If m < k Then
    For i = m To 32
        k = FindSpriteByTag(i)
        If k <> 0 Then Form1.dSprite(k).ZOrder 0
    Next i
    Else
    For i = k To 32
        m = FindSpriteByTag(i)
        If m <> 0 Then Form1.dSprite(m).ZOrder 0
    Next i
End If
End Sub
Private Sub SpriteGetOtherImage(s As Long, b$, tran As Long, rrr As Long, SZ As Single, ROT As Single, Optional bb$ = "") ' before take from priority the original sprite
Dim photo As Object, myRgn As Long
Dim photo2 As Object
If s < 1 Or s > PobjNum Then Exit Sub

      Set photo = New cDIBSection
       Set photo2 = New cDIBSection
           If cDib(b$, photo) Then
 
 If rrr >= 0 Then
 If bb$ <> "" Then
   If cDib(bb$, photo2) Then
 myRgn = fRegionFromBitmap2(photo2)
 Else
 myRgn = fRegionFromBitmap2(photo, tran, CInt(rrr))
 End If

 
 Else
 myRgn = fRegionFromBitmap2(photo, tran, CInt(rrr))
 End If
  If myRgn = 0 Then
 myRgn = CreateRectRgn(2, 2, photo.Width - 2, photo.Height - 2)
 End If
 
 Else

myRgn = CreateRectRgn(2, 2, photo.Width - 2, photo.Height - 2)


 End If



''If SZ <> 1 Then myRgn = ScaleRegion(myRgn, SZ)

myRgn = RotateRegion(myRgn, (ROT), photo.Width * SZ, photo.Height * SZ, SZ)

 RotateDibNew photo, (ROT), 1, tran
 

 
Dim oldtag As Long


With Form1.dSprite(s)
.Height = photo.Height * DYP * SZ
.Width = photo.Width * DXP * SZ
.Picture = photo.Picture(SZ)
.Left = .Left + players(s).x - .Width / 2
players(s).x = .Width / 2
.top = .top + players(s).y - .Height / 2
players(s).y = .Height / 2
Call SetWindowRgn(.hWnd, myRgn, True)
''''''''''''''''''''''''UpdateWindow .hwnd
 DeleteObject myRgn

End With

End If
End Sub

Sub addSprite()
PobjNum = PobjNum + 1
'
End Sub
Sub ClrSprites()
Dim i As Long
If PobjNum > 0 Then
For i = PobjNum To 1 Step -1
players(i).x = 0: players(i).y = 0
PobjNum = i
Unload Form1.dSprite(PobjNum)
Next i
PobjNum = 0

End If
' pObject

End Sub
Public Function fRegionFromBitmap2(picSource As cDIBSection, Optional lBackColor As Long = &HFFFFFF, Optional RANGE As Integer = 0) As Long
Dim myRgn() As RECT
Dim lReturn   As Long
Dim lRgnTmp   As Long
Dim lSkinRgn  As Long
Dim lStart    As Long
Dim lRow      As Long
Dim lCol      As Long
'............
Dim bDib() As Byte
Dim tSA As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = picSource.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = picSource.BytesPerScanLine()
        .pvData = picSource.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
'.........................
Dim BR As Integer, BG As Integer, bbb As Integer, ba$
ba$ = Hex$(lBackColor)
ba$ = Right$("00000" & ba$, 6)
BR = val("&h" & Mid$(ba$, 1, 2))
BG = val("&h" & Mid$(ba$, 3, 2))
bbb = val("&h" & Mid$(ba$, 5, 2))

'..................................
Dim mmx As Long, mmy As Long, cc As Long

Dim GLHEIGHT, GLWIDTH As Long
    GLHEIGHT = picSource.Height
    GLWIDTH = picSource.Width
    ReDim myRgn(picSource.Height * 4) As RECT
    Dim rectCount As Long, oldrect
    rectCount = -1
  mmy = -1 ''GLHEIGHT
    For lRow = GLHEIGHT - 1 To 0 Step -1
        lCol = 0
        mmx = 0
      mmy = mmy + 1
        Do While lCol < GLWIDTH
            ' Skip all pixels in a row with the same
            ' color as the background color.
            '
            Do While lCol < GLWIDTH
             
            If Abs(bDib(mmx, mmy) - BR) > RANGE Or Abs(bDib(mmx + 1, mmy) - BG) > RANGE Or Abs(bDib(mmx + 2, mmy) - bbb) > RANGE Then Exit Do
               lCol = lCol + 1
                mmx = mmx + 3
            Loop

            If lCol < GLWIDTH Then
                '
                ' Get the start and end of the block of pixels in the
                ' row that are not the same color as the background.
                '
                lStart = lCol
               
                Do While lCol < GLWIDTH
                 If Not (Abs(bDib(mmx, mmy) - BR) > RANGE Or Abs(bDib(mmx + 1, mmy) - BG) > RANGE Or Abs(bDib(mmx + 2, mmy) - bbb) > RANGE) Then Exit Do

                mmx = mmx + 3
                    lCol = lCol + 1
                   
                Loop
                
                If lCol > GLWIDTH Then lCol = GLWIDTH
                If rectCount + 2 >= UBound(myRgn()) Then
                ReDim Preserve myRgn(UBound(myRgn()) * 2)
                End If
                
               oldrect = rectCount
              rectCount = rectCount + 1
              SetRect myRgn(rectCount + 2&), lStart, lRow, lCol, lRow + 1

             ''lCol = GLWIDTH
               
            End If
        Loop
    Next

    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
   
    fRegionFromBitmap2 = c_CreatePartialRegion(myRgn(), 2&, rectCount + 1&, 0&, picSource.Width)

End Function


Public Function fRegionFromBitmap(picSource As cDIBSection, Optional lBackColor As Long = &HFFFFFF, Optional RANGE As Integer = 0) As Long
Dim lReturn   As Long
Dim lRgnTmp   As Long
Dim lSkinRgn  As Long
Dim lStart    As Long
Dim lRow      As Long
Dim lCol      As Long
'............
Dim bDib() As Byte
Dim tSA As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = picSource.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = picSource.BytesPerScanLine()
        .pvData = picSource.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
'.........................
Dim BR As Integer, BG As Integer, bbb As Integer, ba$
ba$ = Hex$(lBackColor)
ba$ = Right$("00000" & ba$, 6)
BR = val("&h" & Mid$(ba$, 1, 2))
BG = val("&h" & Mid$(ba$, 3, 2))
bbb = val("&h" & Mid$(ba$, 5, 2))

'..................................
Dim mmx As Long, mmy As Long, cc As Long

Dim GLHEIGHT, GLWIDTH As Long
    GLHEIGHT = picSource.Height
    GLWIDTH = picSource.Width
lSkinRgn = CreateRectRgn(0, 0, 0, 0)
  mmy = GLHEIGHT

    For lRow = 0 To GLHEIGHT - 1
        lCol = 0
        mmx = 0
      mmy = mmy - 1
        Do While lCol < GLWIDTH
            ' Skip all pixels in a row with the same
            ' color as the background color.
            '
            Do While lCol < GLWIDTH
            If Abs(bDib(mmx, mmy) - BR) > RANGE Or Abs(bDib(mmx + 1, mmy) - BG) > RANGE Or Abs(bDib(mmx + 2, mmy) - bbb) > RANGE Then Exit Do
                lCol = lCol + 1
                mmx = mmx + 3
            Loop

            If lCol < GLWIDTH Then
                '
                ' Get the start and end of the block of pixels in the
                ' row that are not the same color as the background.
                '
                lStart = lCol
                Do While lCol < GLWIDTH
                If Not (Abs(bDib(mmx, mmy) - BR) > RANGE Or Abs(bDib(mmx + 1, mmy) - BG) > RANGE Or Abs(bDib(mmx + 2, mmy) - bbb) > RANGE) Then Exit Do

                mmx = mmx + 3
                    lCol = lCol + 1
                Loop
                If lCol > GLWIDTH Then lCol = GLWIDTH
                '
              
                lRgnTmp = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
                lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                Call DeleteObject(lRgnTmp)
            End If
        Loop
    Next

    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
fRegionFromBitmap = lSkinRgn

End Function

Public Function GetFrequency(Oct As Integer, No As Integer)
Dim lngNote As Long
lngNote = ((Oct - 1) * 12 + No) - 37
GetFrequency = 440 * (2 ^ (lngNote / 12))
End Function
Public Function GetNote(Oct As Integer, No As Integer) As Long
GetNote = Oct * 12 + No
End Function
Public Sub PlayTune(ss$)
Dim octave As Integer, i As Long, v$
Dim note As Integer
Dim silence As Boolean
octave = 4
ss$ = ss$ & " "
For i = 1 To Len(ss$) - 1
v$ = Mid$(ss$, i, 2)
note = InStr(FACE$, UCase(v$))
If note = 24 Then

If silence Then
Sleep beeperBEAT / 2
Else
silence = True
Sleep beeperBEAT + beeperBEAT / 2
End If
Else
If note = 0 Then note = InStr(FACE$, UCase(Left$(v$, 1)) & " ") Else i = i + 1
If note <> 0 Then
' look for number

If Mid$(ss$, i + 1, 1) <> "" Then If InStr("1234567", Mid$(ss$, i + 1, 1)) > 0 Then octave = val(Mid$(ss$, i + 1, 1)): i = i + 1
' no volume control here
silence = False
Beeper GetFrequency(octave, (note + 1) / 2), beeperBEAT
End If
End If
Next i
End Sub
Public Function PlayTuneMIDI(ss$, octave2play As Integer, note2play As Integer, subbeat As Long, volume2play As Long) As Boolean

Dim i As Long, v$, nomore As Boolean, yesvol As Boolean, probe2play As Integer
ss$ = ss$ & " "
note2play = 0
i = 1
If Trim$(ss$) = "" Then note2play = 0: Exit Function
If Asc(ss$) <> 32 Then
v$ = Mid$(ss$, i, 2)
probe2play = InStr(FACE$, UCase(v$))
Else
probe2play = 24
End If

If probe2play = 24 Then

    note2play = 24
    i = i + 1
If Mid$(ss$, i, 1) = "@" Then
            i = i + 1
           If InStr("12345", Mid$(ss$, i, 1)) > 0 Then
           subbeat = val(Mid$(ss$, i, 1))
        i = i + 1
      End If
     End If
     If Mid$(ss$, i, 1) = "V" Then
            i = i + 1
            v$ = ""
        Do While InStr("1234567890", Mid$(ss$, i, 1)) > 0 And (Mid$(ss$, i, 1) <> "")
        v$ = v$ & Mid$(ss$, i, 1)
        i = i + 1
        Loop
        volume2play = val("0" & v$)
     End If
    PlayTuneMIDI = True
    GoTo th
Else
If probe2play = 0 Then
probe2play = InStr(FACE$, UCase(Left$(v$, 1)) & " ")
Else
i = i + 1
End If


If probe2play <> 0 Then
i = i + 1
' look for number
If Mid$(ss$, i, 1) <> "" Then
      If InStr("1234567", Mid$(ss$, i, 1)) > 0 Then
        octave2play = val(Mid$(ss$, i, 1))
         i = i + 1
        End If
        If Mid$(ss$, i, 1) = "@" Then
            i = i + 1
           If InStr("12345", Mid$(ss$, i, 1)) > 0 Then
           subbeat = val(Mid$(ss$, i, 1))
        i = i + 1
      End If
       
    End If
         If Mid$(ss$, i, 1) = "V" Then
            i = i + 1
            v$ = ""
        Do While InStr("1234567890", Mid$(ss$, i, 1)) > 0 And (Mid$(ss$, i, 1) <> "")
        v$ = v$ & Mid$(ss$, i, 1)
        i = i + 1
        Loop
        volume2play = val("0" & v$)
     End If
End If

' so we have it here
note2play = probe2play

PlayTuneMIDI = True
End If
End If
th:
ss$ = Mid$(ss$, 1, Len(ss$) - 1) ' drop space
If i = 1 Then note2play = 0: PlayTuneMIDI = False: Exit Function

ss$ = Mid$(ss$, i)

End Function
Public Sub sThread(ByVal ThID As Long, ByVal Thinterval As Double, ByVal ThCode As String, ByVal where$)
Dim task As TaskInterface
 Set task = New counter
          Set task.Owner = Form1.DIS
          ' not use holdtime yet
          task.Parameters ThID, Thinterval, ThCode, uintnew(-1), where$ ', holdtime
          TaskMaster.AddTask task, tmHigh

End Sub
Public Sub sThreadInternal(BS As basetask, ByVal ThID As Long, ByVal Thinterval As Double, ByVal ThCode As String, holdtime As Double, threadhere$, Nostretch)
Dim task As TaskInterface, bsdady As basetask
Set bsdady = BS.Parent
' above 20000 the thid
 Set task = New myProcess
 
          Set task.Owner = BS.Parent.Owner
          Set task.Process = BS
          
          Set bsdady.LinkThread(ThID) = BS.Process
          Set BS = Nothing
          task.Parameters ThID, Thinterval, ThCode, holdtime, threadhere$, Nostretch
          TaskMaster.rest
          TaskMaster.AddTask task
          DoEvents
          Set task = Nothing
          
          Set bsdady = Nothing
          TaskMaster.RestEnd

End Sub
Public Function SetTextData( _
        ByVal lFormatId As Long, _
         sText As String _
    ) As Boolean
    ' use strptr and lenb
    Dim hMem As Long, lPtr As Long
    Dim lSize As Long
        lSize = LenB(sText)
    hMem = GlobalAlloc(0, lSize + 2)
If (hMem > 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, ByVal StrPtr(sText), lSize + 1
        GlobalUnlock hMem
       If (OpenClipboard(0) <> 0) Then
    
      SetClipboardData lFormatId, hMem
      CloseClipboard
       End If
          
    End If
    

End Function
Public Function HTML(sText As String, _
   Optional sContextStart As String = "<HTML><BODY>", _
   Optional sContextEnd As String = "</BODY></HTML>") As Byte()
   Dim m_sDescription As String
    m_sDescription = "Version:1.0" & vbCrLf & _
                  "StartHTML:aaaaaaaaaa" & vbCrLf & _
                  "EndHTML:bbbbbbbbbb" & vbCrLf & _
                  "StartFragment:cccccccccc" & vbCrLf & _
                  "EndFragment:dddddddddd" & vbCrLf
    Dim A() As Byte, b() As Byte, c() As Byte
   '' sText = "<FONT FACE=Arial SIZE=1 COLOR=BLUE>" + sText + "</FONT>"
   
    A() = Utf16toUtf8(sContextStart & "<!--StartFragment -->")
    b() = Utf16toUtf8(sText)
    c() = Utf16toUtf8("<!--EndFragment -->" & sContextEnd)
   Dim sData As String, mdata As Long, eData As Long, fData As Long

   
    eData = UBound(A()) - LBound(A()) + 1
   mdata = UBound(b()) - LBound(b()) + 1
   fData = UBound(c()) - LBound(c()) + 1
   m_sDescription = Replace(m_sDescription, "aaaaaaaaaa", Format(Len(m_sDescription), "0000000000"))
   m_sDescription = Replace(m_sDescription, "bbbbbbbbbb", Format(Len(m_sDescription) + eData + mdata + fData, "0000000000"))
   m_sDescription = Replace(m_sDescription, "cccccccccc", Format(Len(m_sDescription) + eData, "0000000000"))
   m_sDescription = Replace(m_sDescription, "dddddddddd", Format(Len(m_sDescription) + eData + mdata, "0000000000"))
  Dim all() As Byte, m() As Byte
  ReDim all(Len(m_sDescription) + eData + mdata + fData)
  
  m() = Utf16toUtf8(m_sDescription)
  CopyMemory all(0), m(0), Len(m_sDescription)
  CopyMemory all(Len(m_sDescription)), A(0), eData
  CopyMemory all(Len(m_sDescription) + eData), b(0), mdata
  CopyMemory all(Len(m_sDescription) + eData + mdata), c(0), fData
  HTML = all()
  
End Function

Public Function SimpleHtmlData(ByVal sText As String)
Dim lFormatId As Long, bb() As Byte
lFormatId = RegisterCF
If lFormatId <> 0 Then
If sText = "" Then Exit Function
bb() = HTML(sText)
If CBool(OpenClipboard(0)) Then
   
      Dim hMemHandle As Long, lpData As Long
      
      hMemHandle = GlobalAlloc(0, UBound(bb()) - LBound(bb()) + 10)
      
      If CBool(hMemHandle) Then
               
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            
            CopyMemory ByVal lpData, bb(0), UBound(bb()) - LBound(bb())
            GlobalUnlock hMemHandle
            EmptyClipboard
            SetClipboardData lFormatId, hMemHandle
                        
         End If
      
      End If
   
      Call CloseClipboard
   End If



End If
End Function
Function RegisterCF() As Long


   'Register the HTML clipboard format
   If (m_cfHTMLClipFormat = 0) Then
      m_cfHTMLClipFormat = RegisterClipboardFormat("HTML Format")
   End If
   RegisterCF = m_cfHTMLClipFormat
   
End Function
Public Function SetTextDataLong( _
        ByVal lFormatId As Long, _
         dLong As Long _
    ) As Boolean
    ' use strptr and lenb
    Dim hMem As Long, lPtr As Long
    Dim checkme As Long
    Dim lSize As Long
        lSize = 4
    hMem = GlobalAlloc(0, lSize)
If (hMem > 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, dLong, lSize
        CopyMemory checkme, ByVal lPtr, lSize
        GlobalUnlock hMem
       If (OpenClipboard(0) <> 0) Then
       SetClipboardData lFormatId, hMem
      CloseClipboard
       End If
          
    End If
    

End Function
Public Function SetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean
Dim lSize As Long
Dim lPtr As Long
Dim hMem As Long

    lSize = UBound(bData) - LBound(bData) + 1
    hMem = GlobalAlloc(GMEM_DDESHARE + GMEM_MOVEABLE, lSize)
    If (hMem <> 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, bData(LBound(bData)), lSize
        GlobalUnlock hMem
        OpenClipboard Form1.hWnd
        EmptyClipboard
        If (SetClipboardData(lFormatId, hMem) <> 0) Then
          SetBinaryData = True
        End If
       CloseClipboard
    End If
End Function

Public Function GetClipboardMemoryHandle( _
        ByVal lFormatId As Long _
    ) As Long

    
    ' If the format id is there:
    If (IsClipboardFormatAvailable(lFormatId) <> 0) Then
        ' Get the global memory handIsClipboardFormatAvailable(lFormatId)le to the clipboard data:
       
        GetClipboardMemoryHandle = GetClipboardData(lFormatId)
        
    End If
End Function
Private Function GetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean
' Returns a byte array containing binary data on the clipboard for
' format lFormatID:
Dim hMem As Long, lSize As Long, lPtr As Long
    
    ' Ensure the return array is clear:
    Erase bData
    
    hMem = GetClipboardMemoryHandle(lFormatId)
    ' If success:
    If (hMem <> 0) Then
        ' Get the size of this memory block:
        lSize = GlobalSize(hMem)
        ' Get a pointer to the memory:
        lPtr = GlobalLock(hMem)
        If (lSize > 0) Then
            ' Resize the byte array to hold the data:
            ReDim bData(0 To lSize - 2) As Byte
            ' Copy from the pointer into the array:
            CopyMemory bData(0), ByVal lPtr, lSize - 1
        End If
        ' Unlock the memory block:
        GlobalUnlock hMem
        ' Success:
        GetBinaryData = (lSize > 0)
        ' Don't free the memory - it belongs to the clipboard.
    End If
End Function

Public Function GetTextData(ByVal lFormatId As Long) As String
Dim bData() As Byte, sr As String, sr1 As String
sr1 = Clipboard.GetText(1)
If (OpenClipboard(0) <> 0) Then

        
        If (GetBinaryData(lFormatId, bData())) Then
        sr = bData

            GetTextData = Left$(sr, Len(sr1))
          
        End If

End If
CloseClipboard
End Function
Public Function GetImage() As String
Dim hMem As Long, hDIb As Long
Dim mypic As New cDIBSection
Const CF_DIB = 8
Dim isbitmap As Boolean, okb As Boolean

                     If (OpenClipboard(0) <> 0) Then
                  hMem = GetClipboardData(CF_DIB)
                
                
                If hMem <> 0 Then
              
            hDIb = GlobalLock(hMem)
           
                mypic.ClearUp
       okb = mypic.CreateFromDIB(hDIb)
       If Not okb Then
       hDIb = GlobalUnlock(hMem)
       CloseClipboard
       If Clipboard.GetFormat(2) Then
       mypic.CreateFromPicture Clipboard.GetData(2)
       okb = mypic.Height
       End If
       End If
               If okb Then
               If mypic.bitsPerPixel <> 24 Then Conv24 mypic
               If mypic.dpix = 0 Then mypic.GetDpi 96, 96
               If mypic.Height > 0 And mypic.hDIb <> 0 Then
GetImage = DIBtoSTR(mypic)
      End If
      

                             End If
                           Call GlobalUnlock(hMem)
                         End If
 
CloseClipboard

                    End If

End Function

Public Function MsgBoxN(A$, Optional v As Variant = 5, Optional b$) As Long
AskInput = False
If ASKINUSE Then

Exit Function
End If
    AskTitle$ = b$
    Dim resp As Double
       v = v And &HF
     DialogSetupLang DialogLang
    If DialogLang = 1 Then
        If v = vbRetryCancel Then
        AskOk$ = "RETRY"
        ElseIf v = vbYesNo Then
        AskOk$ = "YES"
        AskCancel$ = "NO"
        ElseIf v = vbOKCancel Then
        AskOk$ = "OK"
        Else
        AskOk$ = "OK"
        AskCancel$ = ""
        End If
        AskText$ = A$ + "..?" + vbCrLf
    Else
             If v = vbRetryCancel Then
        AskOk$ = "епамакгьг"
        ElseIf v = vbYesNo Then
        AskOk$ = "маи"
        AskCancel$ = "ови"
        ElseIf v = vbOKCancel Then
         AskOk$ = "емтанеи"
        Else
        AskCancel$ = ""
        AskOk$ = "емтанеи"
        End If
        AskText$ = A$ + "..;" + vbCrLf
    End If

    resp = Form3.NeoASK(basestack1)
    
 If resp = 0 Then
 
 End If
 
    If v = vbYesNo Then
        If resp = 1 Then MsgBoxN = vbYes Else MsgBoxN = vbNo
    ElseIf v = vbOKCancel Then
        If resp = 1 Then MsgBoxN = vbOK Else MsgBoxN = vbCancel
    ElseIf v = vbRetryCancel Then
        If resp = 1 Then MsgBoxN = vbRetry Else MsgBoxN = vbCancel
    Else
    MsgBoxN = 1
    End If
End Function
Public Function InputBoxN(A$, b$, vv$) As String
Dim resp As Double
If ASKINUSE Then

Exit Function
End If
     DialogSetupLang DialogLang

    AskText$ = A$
    AskTitle$ = b$
    AskInput = True
    AskStrInput$ = Trim$(vv$)
    

    resp = Form3.NeoASK(basestack1)
        If resp = 1 Then InputBoxN = basestack1.soros.PopStr
          AskInput = False
End Function
Public Function ask(A$, Optional retry As Boolean = False) As Double
If Form3.Visible Then
If Form3.WindowState = 1 Then
Form3.Timer1.enabled = False
Form3.Timer1.Interval = 32760
Form3.WindowState = 0
If retry Then
If Form1.Visible Then
ask = MsgBoxN(A$, vbRetryCancel + vbQuestion + vbSystemModal, MesTitle$)
Else
ask = MsgBoxN(A$, vbRetryCancel + vbQuestion + vbSystemModal, MesTitle$)
End If

Else
If Form1.Visible Then
ask = MsgBoxN(A$, vbOKCancel + vbQuestion + vbSystemModal, MesTitle$)
Else
ask = MsgBoxN(A$, vbOKCancel + vbQuestion + vbSystemModal, MesTitle$)
End If
End If
Form3.WindowState = 1
Form3.Timer1.enabled = False
Form3.Timer1.Interval = 100
Exit Function
End If
End If
ask = MsgBoxN(A$, vbOKCancel + vbQuestion + vbSystemModal, MesTitle$)
End Function
Public Function SpellUnicode(A$)
' use spellunicode to get numbers
' and make a ListenUnicode...with numbers for input text
Dim b$, i As Long
For i = 1 To Len(A$) - 1
b$ = b$ & CStr(AscW(Mid$(A$, i, 1))) & ","
Next i
SpellUnicode = b$ & CStr(AscW(Right$(A$, 1)))
End Function
Public Function ListenUnicode(ParamArray aa() As Variant) As String
Dim all$, i As Long
For i = 0 To UBound(aa)
    all$ = all$ & ChrW(aa(i))
Next i
ListenUnicode = all$
End Function
Function Convert2(A$, localeid As Long) As String  ' to feed textboxes
Dim b$, i&
If A$ <> "" Then
For i& = 1 To Len(A$)
b$ = b$ + Left$(StrConv(ChrW$(AscW(Left$(StrConv(Mid$(A$, i, 1) + Chr$(0), 128, localeid), 1))), 64, 1033), 1)

Next i&
Convert2 = b$
End If
End Function
Function Convert3(A$, localeid As Long) As String  ' to feed textboxes
Dim b$, i&
If A$ <> "" Then
If localeid = 0 Then localeid = cLid
For i& = 1 To Len(A$)
b$ = b$ + Left$(StrConv(ChrW$(AscW(Left$(StrConv(Mid$(A$, i, 1) + Chr$(0), 128, 1033), 1))), 64, localeid), 1)

Next i&
Convert3 = b$
End If
End Function
Function Convert2Ansi(A$, localeid As Long) As String
Dim b$, i&
If A$ <> "" Then
For i& = 1 To Len(A$)
b$ = b$ + Left$(StrConv(ChrW$(AscW(Left$(StrConv(Mid$(A$, i, 1) + Chr$(0), 128, localeid), 1))), 64, 1032), 1)

Next i&
Convert2Ansi = b$
End If
End Function
Function GetCodePage(Optional localeid As Long = 1032) As Long
  Dim Buffer As String, Ret&
   Buffer = String$(100, 0)

        Ret = GetLocaleInfoW(localeid, LOCALE_IDEFAULTANSICODEPAGE, StrPtr(Buffer), 10)
If Ret > 0 Then
GetCodePage = val(Mid$(Buffer, 1, 41))
End If
End Function
Function GetCharSet(codepage As Long)
'
 Dim cp As CHARSETINFO
     If TranslateCharsetInfo(ByVal codepage, cp, TCI_SRCCODEPAGE) Then
        GetCharSet = cp.ciCharset
    End If
End Function

Sub SwapVariant(ByRef A As Variant, ByRef b As Variant)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(A), 16
   CopyMemory ByVal VarPtr(A), ByVal VarPtr(b), 16
   CopyMemory ByVal VarPtr(b), t(0), 16
End Sub
Sub SwapVariant2(ByRef A As Variant, ByRef b As mArray, i As Long)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(A), 16
   CopyMemory ByVal VarPtr(A), ByVal b.itemPtr(i), 16
   CopyMemory ByVal b.itemPtr(i), t(0), 16
End Sub
Sub SwapVariant3(ByRef A As mArray, k As Long, ByRef b As mArray, i As Long)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal A.itemPtr(k), 16
   CopyMemory ByVal A.itemPtr(k), ByVal b.itemPtr(i), 16
   CopyMemory ByVal b.itemPtr(i), t(0), 16
End Sub

Private Function c_CreatePartialRegion(rgnRects() As RECT, ByVal lIndex As Long, ByVal uIndex As Long, ByVal leftOffset As Long, ByVal cX As Long, Optional ByVal xFrmPtr As Long) As Long
'' THIS IS Lavolpe ROUTINE
    ' Creates a region from a Rect() array and optionally stretches the region

    On Error Resume Next
    ' Note: Ideally, contiguous rows vertically of equal height & width should
    ' be combined into one larger row. However, thru trial & error I found
    ' that Windows does this for us and taking the extra time to do it ourselves
    ' is too cumbersome & slows down the results.
    
    ' the first 32 bytes of a region contain the header describing the region.
    ' Well, 32 bytes equates to 2 rectangles (16 bytes each), so I'll
    ' cheat a little & use rectangles to store the header
    With rgnRects(lIndex - 2&) ' bytes 0-15
        .Left = 32                      ' length of region header in bytes
        .top = 1                        ' required cannot be anything else
        .Right = uIndex - lIndex + 1&   ' number of rectangles for the region
        .Bottom = .Right * 16&          ' byte size used by the rectangles;
    End With                            ' ^^ can be zero & Windows will calculate
    
    With rgnRects(lIndex - 1&) ' bytes 16-31 bounding rectangle identification
        .Left = leftOffset                  ' left
        .top = rgnRects(lIndex).top         ' top
        .Right = leftOffset + cX            ' right
        .Bottom = rgnRects(uIndex).Bottom   ' bottom
    End With
    ' call function to create region from our byte (RECT) array
    c_CreatePartialRegion = ExtCreateRegion(ByVal xFrmPtr, (rgnRects(lIndex - 2&).Right + 2&) * 16&, rgnRects(lIndex - 2&))
    If Err Then Err.Clear

End Function

Function FoundLocaleId(A$) As Long
If Convert3(Convert2(A$, 1032), 1032) = A$ Then
    FoundLocaleId = 1032
ElseIf Convert3(Convert2(A$, 1033), 1033) = A$ Then
    FoundLocaleId = 1033
ElseIf Convert3(Convert2(A$, cLid), cLid) = A$ Then
 FoundLocaleId = cLid
End If
End Function
Function FoundSpecificLocaleId(A$, this As Long) As Long
If Convert3(Convert2(A$, this), this) = A$ Then FoundSpecificLocaleId = True
End Function
Function ismine1(ByVal A$) As Boolean  '  START A BLOCK
ismine1 = True
A$ = myUcase(A$, True)
Select Case A$
Case "DO", "REPEAT", "PART", "LIB"
Case "епамекабе", "епамакабе", "леяос"
Case Else
ismine1 = False
End Select
End Function
Function ismine2(ByVal A$) As Boolean  ' CAN START A BLOCK OR DO SOMETHING
ismine2 = True
A$ = myUcase(A$, True)
Select Case A$
Case "AFTER", "BACK", "BACKGROUND", "CLASS", "COLOR", "DECLARE", "ELSE", "EVENT", "EVERY", "GLOBAL", "FOR", "FUNCTION", "GROUP", "LAYER", "LOCAL", "MAIN.TASK", "MODULE", "PATH", "PEN", "PRINTER", "PRINTING", "STACK", "START", "TASK.MAIN", "THEN", "THREAD", "TRY", "WIDTH", "WHILE"
Case "аявг", "аккиыс", "цецомос", "цемийо", "цемийг", "цемийес", "циа", "дес", "ейтупытгс", "ейтупысг", "емы", "епипедо", "ивмос", "йахе", "йкасг", "йуяио.еяцо", "лета", "мгла", "олада", "ояисе", "павос", "пема", "пеяихыяио", "сумаятгсг", "сыяос", "тлгла", "топийа", "топийг", "топийес", "тоте", "вяыла"
Case "->"
Case Else
ismine2 = False
End Select
End Function
Function ismine3(ByVal A$) As Boolean  ' CAN START A NEW COMMAND
ismine3 = True
A$ = myUcase(A$, True)
Select Case A$
Case "ELSE", "THEN", "LIB"
Case "аккиыс", "тоте"
Case Else
ismine3 = False
End Select
End Function

Function ismine(ByVal A$) As Boolean
ismine = True
A$ = myUcase(A$, True)
Select Case A$
Case "@(", "$(", "~(", "?", "->"
Case "ABOUT", "ABOUT$", "ABS(", "ADD.LICENCE$(", "AFTER", "ALWAYS", "AND", "ANGLE", "APPDIR$", "APPEND", "APPEND.DOC"
Case "ARRAY$(", "ARRAY(", "AS", "ASC(", "ASCENDING", "ASK$(", "ASK(", "ATN("
Case "BACK", "BACKGROUND", "BACKWARD(", "BASE", "BEEP", "BINARY", "BINARY.AND(", "BINARY.NEG("
Case "BINARY.OR(", "BINARY.ROTATE(", "BINARY.SHIFT(", "BINARY.XOR(", "BITMAPS", "BMP$(", "BOLD"
Case "BOOLEAN", "BORDER", "BREAK", "BROWSER", "BROWSER$", "BYTE", "CALL", "CASE", "CAT"
Case "CDATE(", "CENTER", "CHANGE", "CHARSET", "CHOOSE.COLOR", "CHOOSE.FONT", "CHOOSE.ORGAN"
Case "CHR$(", "CHRCODE$(", "CHRCODE(", "CIRCLE", "CLASS", "CLEAR", "CLIPBOARD", "CLIPBOARD$", "CLIPBOARD.IMAGE$"
Case "CLOSE", "CLS", "CODE", "CODEPAGE", "COLLIDE(", "COLOR", "COLOR(", "COLORS"
Case "COLOUR(", "COMMAND", "COMMAND$", "COMMIT", "COMPARE(", "COMPRESS", "COMPUTER", "COMPUTER$", "CONCURRENT"
Case "CONTINUE", "CONTROL$", "COPY", "COS(", "CTIME(", "CURRENCY", "CURSOR", "CURVE"
Case "DATA", "DATE$(", "DATE(", "DATEFIELD", "DB.PROVIDER", "DB.USER", "DECLARE", "DEF", "DELETE"
Case "DESCENDING", "DESKTOP", "DIM", "DIMENSION(", "DIR", "DIR$", "DIV", "DO"
Case "DOC.LEN(", "DOC.PAR(", "DOC.UNIQUE.WORDS(", "DOC.WORDS(", "DOCUMENT", "DOS", "DOUBLE", "DOWN", "DRAW"
Case "DRAWINGS", "DRIVE$(", "DRIVE.SERIAL(", "DROP", "DRW$(", "DURATION"
Case "EDIT", "EDIT.DOC", "ELSE", "ELSE.IF", "EMPTY", "END", "ENVELOPE$(", "EOF("
Case "ERASE", "ERROR", "ERROR$", "ESCAPE", "EVAL(", "EVAL$(", "EVENT", "EVERY", "EXECUTE", "EXIST(", "EXIST.DIR("
Case "EXIT", "EXPORT", "EXTERN", "FALSE", "FAST", "FIELD", "FIELD$(", "FILE$("
Case "FILE.APP$(", "FILE.NAME$(", "FILE.NAME.ONLY$(", "FILE.PATH$(", "FILE.STAMP(", "FILE.TITLE$(", "FILE.TYPE$(", "FILELEN(", "FILES"
Case "FILL", "FILTER$(", "FIND", "FKEY", "FLOODFILL", "FLUSH", "FONT", "FONTNAME$", "FOR"
Case "FORM", "FORMAT$(", "FORWARD(", "FRAC(", "FRAME", "FREQUENCY(", "FROM", "FUNCTION", "FUNCTION$(", "FUNCTION("
Case "GET", "GLOBAL", "GOSUB", "GOTO", "GRABFRAME$", "GRADIENT", "GREEK", "GROUP"
Case "GROUP.COUNT(", "HEIGHT", "HELP", "HEX", "HEX$(", "HIDE", "HIDE$(", "HIFI", "HIGHWORD("
Case "HILOWWORD(", "HIWORD(", "HOLD", "HTML", "HWND", "ICON", "IF", "IMAGE", "IMAGE.X("
Case "IMAGE.X.PIXELS(", "IMAGE.Y(", "IMAGE.Y.PIXELS(", "IN", "INKEY$", "INKEY(", "INLINE", "INPUT", "INPUT$("
Case "INSERT", "INSTR(", "INT(", "INTEGER", "INTERVAL", "ISLET", "ISNUM", "ITALIC"
Case "JOYPAD", "JOYPAD(", "JOYPAD.ANALOG.X(", "JOYPAD.ANALOG.Y(", "JOYPAD.DIRECTION(", "JPG$(", "KEEP", "KEY$", "KEYBOARD"
Case "KEYPRESS(", "LAMBDA", "LAMBDA$", "LAN$", "LATIN", "LAYER", "LAZY$(", "LCASE$(", "LEFT$(", "LEFTPART$(", "LEGEND", "LEN"
Case "LEN(", "LEN.DISP(", "LET", "LETTER$", "LIB", "LICENCE", "LINE", "LINESPACE", "LINK", "LIST", "LN("
Case "LOAD", "LOAD.DOC", "LOCAL", "LOCALE", "LOCALE$(", "LOCALE(", "LOG(", "LONG", "LOOP"
Case "LOWORD(", "LOWWORD(", "MAIN.TASK", "MARK", "MASTER", "MATCH(", "MAX(", "MAX.DATA$("
Case "MAX.DATA(", "MDB(", "MEDIA", "MEDIA.COUNTER", "MEMBER$(", "MEMBER.TYPE$(", "MEMO", "MEMORY", "MENU"
Case "MENU$(", "MENU.VISIBLE", "MENUITEMS", "MERGE.DOC", "METHOD", "MID$(", "MIN(", "MIN.DATA$(", "MIN.DATA("
Case "MOD", "MODE", "MODULE", "MODULE$", "MODULE(", "MODULES", "MONITOR", "MOTION", "MOTION.W", "MOTION.WX"
Case "MOTION.WY", "MOTION.X", "MOTION.XW", "MOTION.Y", "MOTION.YW", "MOUSE", "MOUSE.ICON", "MOUSE.KEY", "MOUSE.X"
Case "MOUSE.Y", "MOUSEA.X", "MOUSEA.Y", "MOVE", "MOVIE", "MOVIE.COUNTER", "MOVIE.DEVICE$", "MOVIE.ERROR$", "MOVIE.STATUS$"
Case "MOVIES", "MUSIC", "MUSIC.COUNTER", "NAME", "NEW", "NEXT"
Case "NORMAL", "NOT", "NOTHING", "NOW", "NUMBER", "OFF", "OLE", "ON"
Case "OPEN", "OPEN.FILE", "OPEN.IMAGE", "OPTIMIZATION", "OR", "ORDER", "OS$", "OUT", "OUTPUT"
Case "OVER", "OVERWRITE", "PAGE", "PARAGRAPH$(", "PARAGRAPH(", "PARAGRAPH.INDEX(", "PARAM(", "PARAM$(", "PARAMETERS$", "PART", "PASSWORD"
Case "PATH", "PATH$(", "PAUSE", "PEN", "PI", "PIPE", "PIPENAME$(", "PLATFORM$", "PLAY"
Case "PLAYER", "PLAYSCORE", "POINT", "POINT(", "POLYGON", "POS", "POS.X", "POS.Y", "PRINT"
Case "PRINTER", "PRINTERNAME$", "PRINTING", "PRIVATE", "PROFILER", "PROPERTIES", "PROPERTIES$", "PUBLIC", "PUSH", "PUT", "QUOTE$("
Case "RANDOM", "RANDOM(", "READ", "RECORDS(", "RECURSION.LIMIT", "REFER", "REFRESH", "RELEASE", "REM"
Case "REMOVE", "REPEAT", "REPLACE$(", "REPORT", "REPORTLINES", "RESTART", "RETRIEVE", "RETURN", "REVISION"
Case "RIGHT$(", "RIGHTPART$(", "RINSTR(", "ROUND(", "ROW", "SAVE", "SAVE.AS", "SAVE.DOC", "SCALE.X"
Case "SCALE.Y", "SCAN", "SCORE", "SCREEN.PIXELS", "SCREEN.X", "SCREEN.Y", "SCRIPT", "SCROLL", "SEARCH"
Case "SEEK", "SEEK(", "SELECT", "SEQUENTIAL", "SET", "SETTINGS", "SGN(", "SHIFT", "SHIFTBACK", "SHORTDIR$("
Case "SHOW", "SHOW$(", "SIN(", "SINGLE", "SINT(", "SIZE", "SIZE.X(", "SIZE.Y(", "SLOW"
Case "SND$(", "SORT", "SOUND", "SOUNDREC", "SOUNDS", "SPEECH", "SPEECH$(", "SPLIT", "SPRITE"
Case "SPRITE$", "SQRT(", "STACK", "STACK$(", "STACK.SIZE", "STACKITEM$(", "STACKITEM(", "STACKTYPE$(", "START", "STATIC"
Case "STEP", "STEREO", "STOCK", "STOP", "STR$(", "STRING$(", "STRUCTURE", "SUB", "SUBDIR"
Case "SWAP", "SWEEP", "SWITCHES", "TAB", "TAB(", "TABLE", "TAN(", "TARGET"
Case "TARGETS", "TASK.MAIN", "TEMPNAME$", "TEMPORARY$", "TEST", "TEXT", "THEN", "THIS"
Case "THREAD", "THREAD.PLAN", "THREADS", "THREADS$", "TICK", "TIME$(", "TIME(", "TIMECOUNT", "TITLE"
Case "TO", "TODAY", "TONE", "TOP", "TRIM$(", "TRUE", "TRY", "TUNE", "TWIPSX"
Case "TWIPSY", "TYPE", "TYPE$(", "UCASE$(", "UINT(", "UNDER", "UNION.DATA$(", "UNTIL"
Case "UP", "UPDATABLE", "UPDATE", "USE", "USER", "USER.NAME$", "USGN("
Case "VAL(", "VALID(", "VERSION", "VIEW", "VOID", "VOLUME"
Case "WAIT", "WCHAR", "WEAK", "WEAK$(", "WHILE", "WIDE", "WIDTH", "WIN", "WINDOW"
Case "WITH", "WORDS", "WRITABLE(", "WRITE", "WRITER", "X.TWIPS", "XOR", "Y.TWIPS", "адеиас"
Case "адеиасе", "ай(", "айеяаио.дуадийо(", "айеяаиос", "акгхес", "акгхгс", "аккацг", "аккацг$("
Case "аккане", "аккиыс", "аккиыс.ам", "ам", "ама", "амафгтгсг", "амахеыягсг", "амайтгсг", "амакоцио"
Case "амакоцио$", "амакусг.охомгс", "амакусг.у", "амакусг.в", "амакутгс", "амаломг", "амамеысг", "амажояа", "амаье"
Case "амехесе", "амоицла.аявеиоу", "амоицла.еийомас", "амоине", "амтецяаье", "амтицяаье", "амы", "аниа(", "апедысе"
Case "апкос", "апо", "апохгйеусг.ыс", "апой$(", "апойопг", "апок(", "аяца", "аяихлос", "аяихлос.паяацяажоу("
Case "аяис$(", "аяистеяолеяос$(", "аявеиа", "аявеио", "аявеио$(", "аявеиоу.лгйос(", "аявеиоу.сталпа(", "аявг", "аукос"
Case "аукос$(", "аукоу", "ауноуса", "ауто", "ажаияесг", "ажгсе", "баке", "баке.адеиа$(", "басг"
Case "басг(", "басг.паяовос", "басг.вягстгс", "баье", "бектистопоигсг", "бгла", "богхеиа", "цецомос", "целисе", "целисла"
Case "цемийес", "цемийг", "цемийо", "циа", "цяалла$", "цяаллатосеияа", "цяаллатосеияа$", "цяаллесамажояас", "цяаллг"
Case "цяажг$(", "цяаье", "цягцояа", "цымиа", "деийтг.лояжг", "деийтгс", "деийтгс.йол", "деийтгс.у", "деийтгс.в"
Case "деийтгса.у", "деийтгса.в", "деине", "дей(", "дейаен", "дейаен$(", "дем", "дени$(", "денилеяос$(", "дес", "дглосио"
Case "диа", "диабасе", "диацяажг", "диадовийо", "диайопг", "диайоптес", "диалесоу", "диаяйеиа", "диастасг("
Case "диастиво", "диажамеиа", "диажамеиа$", "диажамо", "диажуцг", "диейоье", "дийтуо$", "диояхысе"
Case "дипка", "дипкос", "дойилг", "долг", "дяолеас", "дуадийг.пеяистяожг(", "дуадийо", "дуадийо(", "дуадийо.айеяаио("
Case "дуадийо.амти(", "дуадийо.амтистяожо(", "дуадийо.апо(", "дуадийо.г(", "дуадийо.йаи(", "дуадийо.окисхгсг(", "дуолиса(", "дысе"
Case "еццяажес(", "еццяажо", "еццяажоу.кенеис(", "еццяажоу.лгйос(", "еццяажоу.ломадийес.кенеис(", "еццяажоу.пая(", "еццяаьило(", "ецйуяо(", "еий$("
Case "еийома", "еийома.у(", "еийома.у.сглеиа(", "еийома.в(", "еийома.в.сглеиа(", "еийомес", "еийомидио", "еимая", "еимця"
Case "еисацыцг", "еисацыцг$(", "еисацыцгс", "ейдосг", "ейтекесг", "ейтупысг", "ейтупысгс", "ейтупытгс", "ейтупытгс$", "ейжя(", "ейжя$("
Case "ейжяасг(", "ейжяасг$(", "екецвос", "еккгмийа", "емаомола$", "емхесг", "емйол$", "емйол(", "емтасг", "емтокг$"
Case "емы", "емысе", "емысг.сеияас$(", "енацыцг", "енодос", "енытеяийг", "епам$(", "епамакабе", "епамекабе"
Case "епамы", "епекене", "епекене.цяаллатосеияа", "епекене.ояцамо", "епекене.вяыла", "епицяажг", "епийаияо", "епикене", "епикене.цяаллатосеияа"
Case "епикене.ояцамо", "епикене.вяыла", "епикоцес", "епикоцес$(", "епикоцес.жамеяес", "епикоцг", "епикоцг$(", "епикоцгс", "епипедо"
Case "епистяожг", "епижамеиа", "еполемо", "еуяеиа", "еуяесг", "еуяиа", "ежап(", "ежаялоцг.аявеиоу$(", "ежаялоцг.йат$"
Case "еыс", "г", "гл(", "глеяа$(", "глеяа(", "глеяолгмиа", "гво$(", "гвоцяажгсг"
Case "гвои", "гвос", "хесе", "хесг", "хесг(", "хесг.у", "хесг.в", "хесгдениа("
Case "идиотгтес", "идиотгтес$", "идиытийо", "исвмг", "исвмг$(", "ивмос", "йахаяг", "йахаяо", "йахе", "йаи", "йакесе", "йалпукг"
Case "йаме", "йамомийа", "йат", "йат$", "йатакоцои", "йатакоцос", "йатастасг.таимиас$", "йатавыягсг", "йаты"
Case "йатылисо(", "йеилемо", "йемг", "йемо", "йемтяо", "йеж$(", "йимгсг", "йимгсг.п", "йимгсг.пу"
Case "йимгсг.пв", "йимгсг.у", "йимгсг.уп", "йимгсг.в", "йимгсг.вп", "йкасг", "йкеиди", "йкеисе", "йкилан.у"
Case "йкилан.в", "йол$", "йомсока", "йяата", "йяатгсе", "йяужо$(", "йяуье", "йуйкийа", "йуйкос"
Case "йукисг", "йуяио", "йуяио.еяцо", "йыд(", "йыдийа", "йыдийосекида", "кабг", "кабг(", "кабг.амакоцийо.у("
Case "кабг.амакоцийо.в(", "кабг.йатеухумсг(", "кахос", "кахос$", "кахос.таимиас$", "калда", "калда$", "катимийа", "кенеис", "киста", "коц("
Case "коцийос", "коцистийо", "коцос", "коцос$(", "кс$", "кж(", "лайяус", "ле", "лецако("
Case "лецако.сеияас$(", "лецако.сеияас(", "лецехос", "лецехос.сыяоу", "лецехос.у(", "лецехос.в(", "леходос", "лекос$(", "лекоус.тупос$("
Case "лекыдиа", "леяос", "лес$(", "лета", "летахесг", "летахесг(", "левяи", "лгйос", "лгйос(", "лгйос.елж("
Case "лийяо(", "лийяо.сеияас$(", "лийяо.сеияас(", "лийяос.йатакоцос$(", "лмглг", "лояжг$(", "лоусийг", "лоусийг.летягтгс", "лпип"
Case "лпяоста(", "маи", "меа", "мео", "меои", "меос", "мгла", "мглата", "мглата$"
Case "нейима", "одгциа", "одгцос$(", "охомг", "ойм$(", "олада", "олада.сумоко(", "омола", "омола.аявеиоу$("
Case "омола.аявеиоу.ломо$(", "омола.вягстг$", "ояио.амадяолгс", "ояисе", "ови", "паифеижымг", "паийтгс", "паине", "памта"
Case "памы", "памылисо(", "паяацяажос$(", "паяацяажос(", "паяал(", "паяал$(", "паяахесг$(", "паяахуяо", "паяалетяои$", "паяе", "паяейаяе$"
Case "паяелбокг", "патглемо(", "павос", "педиа", "педио", "педио$(", "пеф$(", "пема", "пеяи"
Case "пеяи$", "пеяихыяио", "пета", "пи", "пимайас", "пимайас$(", "пимайас(", "пимайес", "писы("
Case "пкациа", "пкаисио", "пкатос", "пкатос.сглеиоу", "пкатжояла$", "пкгйтяокоцио", "покуцымо"
Case "пяос", "пяосаялоцгс", "пяосхесе.еццяажо", "пяосхгйг", "пяосыяимо$", "пяовеияо", "пяовеияо$", "пяовеияо.еийома$", "яифа("
Case "яоутима", "яоутимас", "яухлисеис", "яыта$(", "яыта(", "саяысе", "сбгсе", "се"
Case "сеияа", "сеияиайос.дисйоу(", "секида", "семаяио", "сгл", "сгл(", "сглади", "сглеио", "сглеио(", "сглеяа", "статийг", "статийес"
Case "стг", "стгкг", "стгкг(", "стгм", "сто", "стой", "стовои", "стовос", "стяоцц(", "суццяажеас"
Case "суццяажг", "суцйяиме(", "суцйяоусг(", "суцвымеусе.еццяажо", "сулпиесг", "сулпкгяысг", "сум(", "сумаятгсг", "сумаятгсг$("
Case "сумаятгсг(", "сумевисе", "сумхгла", "сус", "сусйеуг.пяобокгс$", "сустгла", "сувмотгта(", "свд$(", "сведиа"
Case "сведио.мглатым", "сыяос", "сыяос$(", "сыяоутупос$(", "сысе", "сысе.еццяажо", "таимиа", "таимиа.летягтгс", "таимиес"
Case "танг", "танимолгсг", "таутисг(", "таутовяомо", "текос", "текос(", "тий", "тиктос.аявеиоу$(", "тилг"
Case "тилг(", "тилгсыяоу$(", "тилгсыяоу(", "типота", "титкос", "тлгла", "тлгла(", "тлгла$", "тлглата", "томос"
Case "тон.еж(", "топийа", "топийес", "топийг", "топийо", "топийо$(", "топийо(", "топос$(", "топос.аявеиоу$("
Case "тоте", "тупос", "тупос$(", "тупос.аявеиоу$(", "тупысе", "туваиос(", "тыяа", "у.сглеиа"
Case "упаявеи(", "упаявеи.йатакоцос(", "уплея(", "упо", "упойатакоцос", "упок", "упокоцистг", "упокоцистгс$", "упокоипо"
Case "уполмгла", "упыяа(", "уьос", "уьос.сглеиоу", "жайекос$(", "жамеяо$(", "жаядиа", "жеяе"
Case "жеяеписы", "жхимоуса", "жиктяо$(", "жомто", "жояла", "жояла$", "жоятос", "жоятысе", "жоятысе.еццяажо"
Case "жяасг", "жымг", "жыто$(", "в.сглеиа", "вая$(", "ваяайтгяес", "ваяане", "ваяйыд$("
Case "ваяйыд(", "вягсг", "вягстг", "вягстгс", "вяомос$(", "вяомос(", "вяыла", "вяыла(", "вяылата"
Case "выяисла", "ьеудес", "ьеудгс", "ьгжио", "ыс"
Case Else
ismine = False
End Select
End Function
Private Function IsNumberQuery(A$, fr As Long, r As Double, lR As Long, skipdecimals As Boolean) As Boolean
Dim SG As Long, sng As Long, n$, ig$, DE$, sg1 As Long, ex$   ', e$
' ti kanei to e$
If A$ = "" Then IsNumberQuery = False: Exit Function
SG = 1
sng = fr - 1
    Do While sng < Len(A$)
    sng = sng + 1
    Select Case Mid$(A$, sng, 1)
    Case " ", "+"
    Case "-"
    SG = -SG
    Case Else
    Exit Do
    End Select
    Loop
n$ = Mid$(A$, sng)

If val("0" & Mid$(A$, sng, 1)) = 0 And Left(Mid$(A$, sng, 1), sng) <> "0" And Left(Mid$(A$, sng, 1), sng) <> "." Then
IsNumberQuery = False

Else
'compute ig$
    If Mid$(A$, sng, 1) = "." And Not skipdecimals Then
    ' no long part
    ig$ = "0"
    DE$ = "."

    Else
    Do While sng <= Len(A$)
        
        Select Case Mid$(A$, sng, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(A$, sng, 1)
        Case "."
        If skipdecimals Then IsNumberQuery = False: Exit Function
        DE$ = "."
        Exit Do
        Case Else
        Exit Do
        End Select
       sng = sng + 1
    Loop
    End If
    ' compute decimal part
    If DE$ <> "" Then
      sng = sng + 1
        Do While sng <= Len(A$)
       
        Select Case Mid$(A$, sng, 1)
        Case " "
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(A$, sng, 1)
        Else
        DE$ = DE$ & Mid$(A$, sng, 1)
        End If
        Case "E", "e" ' ************check it
        If skipdecimals Then Exit Do
             If ex$ = "" Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
   
               Case "е", "Е" ' ************check it
               If skipdecimals Then Exit Do
                         If ex$ = "" Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
        
        
        Case "+", "-"
        If sg1 And Len(ex$) = 1 Then
         ex$ = ex$ & Mid$(A$, sng, 1)
        Else
        Exit Do
        End If
        Case Else
        Exit Do
        End Select
         sng = sng + 1
        Loop
        If ex$ = "E" Or ex$ = "E-" Or ex$ = "E+" Then
        sng = sng - Len(ex$)
        End If
    End If
    If ig$ = "" Then
    IsNumberQuery = False
    lR = 1
    Else
    If SG < 0 Then ig$ = "-" & ig$
    Err.Clear
    On Error Resume Next
    If Len(ex$) = 1 Then
    n$ = ig$ & DE$ & ex$ + "1"
     If IsExp(basestack1, n$, r) Then
    sng = Len(ig$ & DE$ & ex$) - Len(n$)
        fr = 0
    End If
       Else
    n$ = ig$ & DE$ & ex$
    If IsExp(basestack1, n$, r) Then
    sng = Len(ig$ & DE$ & ex$) - Len(n$)
        fr = 0
    Else
    End If
    r = val(ig$ & DE$ & ex$)
    End If
    If Err > 0 Then
    lR = 0
    Else
      'A$ = Mid$(A$, sng)
    lR = sng - fr + 1
       IsNumberQuery = True
    End If
    End If
End If
End Function
Function ValidNum(A$, final As Boolean, Optional cutdecimals As Boolean = False) As Boolean
Dim r As Long
Dim r1 As Long
r1 = 1

Dim v As Double, b$
If final Then
r1 = IsNumberOnly(A$, r1, v, r, cutdecimals)
 r1 = (r1 And Len(A$) <= r) Or (A$ = "")
 
Else
If (A$ = "-") Or A$ = "" Then
r1 = True
Else
 r1 = IsNumberQuery(A$, r1, v, r, cutdecimals)
    If A$ <> "" Then
         If r < 2 Then
                r1 = Not (r <= Len(A$))
                A$ = ""
        Else
                r1 = r1 And Not r <= Len(A$)
                A$ = Mid$(A$, 1, r - 1)
        End If
 End If
 End If
 End If
ValidNum = r1
End Function
Function ValidNumberOnly(A$, r As Double, skipdec As Boolean) As Boolean
ValidNumberOnly = IsNumberOnly(A$, (1), r, (0), skipdec)
End Function
Private Function IsNumberOnly(A$, fr As Long, r As Double, lR As Long, skipdecimals As Boolean) As Boolean
Dim SG As Long, sng As Long, n$, ig$, DE$, sg1 As Long, ex$   ', e$
' ti kanei to e$
If A$ = "" Then IsNumberOnly = False: Exit Function
SG = 1
sng = fr - 1
    Do While sng < Len(A$)
    sng = sng + 1
    Select Case Mid$(A$, sng, 1)
    Case " ", "+"
    Case "-"
    SG = -SG
    Case Else
    Exit Do
    End Select
    Loop
n$ = Mid$(A$, sng)

If val("0" & Mid$(A$, sng, 1)) = 0 And Left(Mid$(A$, sng, 1), sng) <> "0" And Left(Mid$(A$, sng, 1), sng) <> "." Then
IsNumberOnly = False

Else
'compute ig$
    If Mid$(A$, sng, 1) = "." And Not skipdecimals Then
    ' no long part
    ig$ = "0"
    DE$ = "."

    Else
    Do While sng <= Len(A$)
        
        Select Case Mid$(A$, sng, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(A$, sng, 1)
        Case "."
        If skipdecimals Then Exit Do
        DE$ = "."
        Exit Do
        Case Else
        Exit Do
        End Select
       sng = sng + 1
    Loop
    End If
    ' compute decimal part
    If DE$ <> "" Then
      sng = sng + 1
        Do While sng <= Len(A$)
       
        Select Case Mid$(A$, sng, 1)
        Case " "
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(A$, sng, 1)
        Else
        DE$ = DE$ & Mid$(A$, sng, 1)
        End If
        Case "E", "e" ' ************check it
        If skipdecimals Then Exit Do
             If ex$ = "" Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
   
               Case "е", "Е" ' ************check it
               If skipdecimals Then Exit Do
                         If ex$ = "" Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
        
        
        Case "+", "-"
        If sg1 And Len(ex$) = 1 Then
         ex$ = ex$ & Mid$(A$, sng, 1)
        Else
        Exit Do
        End If
        Case Else
        Exit Do
        End Select
         sng = sng + 1
        Loop
        If ex$ = "E" Or ex$ = "E-" Or ex$ = "E+" Then
        sng = sng - Len(ex$)
        End If
    End If
    If ig$ = "" Then
    IsNumberOnly = False
    lR = 1
    Else
    If SG < 0 Then ig$ = "-" & ig$
    r = val(ig$ & DE$ & ex$)
      'A$ = Mid$(A$, sng)
    lR = sng - fr + 1
    IsNumberOnly = True
    End If
End If
End Function
Public Function ScrX() As Long
ScrX = GetSystemMetrics(SM_CXSCREEN) * dv15
End Function
Public Function ScrY() As Long
ScrY = GetSystemMetrics(SM_CYSCREEN) * dv15
End Function
Public Function MyTrimLi(s$, l As Long) As Long
Dim i&
Dim p2 As Long, P1 As Integer, p4 As Long
 If l > Len(s) Then MyTrimLi = Len(s) + 1: Exit Function
 If l <= 0 Then MyTrimLi = 1: Exit Function
  l = l - 1
  i = Len(s)
  p2 = StrPtr(s) + l * 2:  p4 = p2 + i * 2
  For i = p2 To p4 Step 2
  GetMem2 i, P1
  Select Case P1
    Case 32, 160
    Case Else
     MyTrimLi = (i - p2) \ 2 + 1 + l
   Exit Function
  End Select
  Next i
 MyTrimLi = Len(s) + 1
End Function
Public Function MyTrimL(s$) As Long
Dim i&, l As Long
Dim p2 As Long, P1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL = 1: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, P1
  Select Case P1
    Case 32, 160
    Case Else
     MyTrimL = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimL = l + 2
End Function
Public Function excludespace(s$) As Long
Dim i&, l As Long
Dim p2 As Long, P1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, P1
  Select Case P1
    Case 32, 160
    Case Else
     excludespace = (i - p2) \ 2
   Exit Function
  End Select
  Next i

End Function
Function IsLabelAnew(where$, A$, r$, lang As Long) As Long
' for left side...no &

Dim rr&, one As Boolean, c$, gr As Boolean
r$ = ""
' NEW FOR REV 156  - WE WANT TO RUN WITH GREEK COMMANDS IN ANY COMPUTER
Dim i&, l As Long, p3 As Integer
Dim p2 As Long, P1 As Integer, p4 As Long
l = Len(A$): If l = 0 Then IsLabelAnew = 0: lang = 1: Exit Function

p2 = StrPtr(A$): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, P1
  Select Case P1
    Case 13
    
    If i < p4 Then
    GetMem2 i + 2, p3
    If p3 = 10 Then
    IsLabelAnew = 1234
    If i + 6 > p4 Then
    A$ = ""
    Else
    i = i + 4
    Do While i < p4

    GetMem2 i, P1
    If P1 = 32 Or P1 = 160 Then
    i = i + 2
    Else
    GetMem2 i + 2, p3
    If P1 <> 13 And p3 <> 10 Then Exit Do
    i = i + 4
    End If
    Loop
    A$ = Mid$(A$, (i + 2 - p2) \ 2)
    End If
    Else
    If i > p2 Then A$ = Mid$(A$, (i - 2 - p2) \ 2)
    End If
    Else
    If i > p2 Then A$ = Mid$(A$, (i - 2 - p2) \ 2)
    End If
    
    lang = 1
    Exit Function
    Case 32, 160
    Case Else

   Exit For
  End Select
  Next i
    If i > p4 Then A$ = "": IsLabelAnew = 0: Exit Function
  For i = i To p4 Step 2
  GetMem2 i, P1
  If P1 < 256 Then
  Select Case ChrW(P1)
        Case "@"
            If i < p4 And r$ <> "" Then
                GetMem2 i + 2, P1
                where$ = r$
                r$ = ""
            Else
              IsLabelAnew = 0: A$ = Mid$(A$, (i - p2) \ 2): Exit Function
            End If
        Case "?"
        If r$ = "" Then
            r$ = "?"
            i = i + 4
        Else
            i = i + 2
        End If
        A$ = Mid$(A$, (i - p2) \ 2)
        IsLabelAnew = 1
        lang = 1 + CLng(gr)
              
        Exit Function

        Case "."
            If one Then
                Exit For
            ElseIf r$ <> "" And i < p4 Then
                GetMem2 i + 2, P1
                If ChrW(P1) = "." Or ChrW(P1) = " " Then
                If ChrW(P1) = "." And i + 2 < p4 Then
                    GetMem2 i + 4, P1
                    If ChrW(P1) = " " Then i = i + 4: Exit For
                Else
                    i = i + 2
                   Exit For
                End If
            End If
                GetMem2 i, P1
                r$ = r$ & ChrW(P1)
                rr& = 1
            End If
      Case "&"
            If r$ = "" Then
            rr& = 2
            'a$ = Mid$(a$, 2)
            End If
            Exit For
    Case "\", "{" To "~", "^"
          Exit For
        
        Case "0" To "9", "_"
              If one Then

            Exit For
            ElseIf r$ <> "" Then
            r$ = r$ & ChrW(P1)
            '' A$ = Mid$(A$, 2)
            rr& = 1 'is an identifier or floating point variable
            Else
            Exit For
            End If
        Case Is >= "A"
            If one Then
            Exit For
            Else
            r$ = r$ & ChrW(P1)
            rr& = 1 'is an identifier or floating point variable
            End If
        Case "$"
            If one Then Exit For
            If r$ <> "" Then
            one = True
            rr& = 3 ' is string variable
            r$ = r$ & ChrW(P1)
            Else
            Exit For
            End If
        Case "%"
            If one Then Exit For
            If r$ <> "" Then
            one = True
            rr& = 4 ' is long variable
            r$ = r$ & ChrW(P1)
            Else
            Exit For
            End If
            
        Case "("
            If r$ <> "" Then
            If i + 4 <= p4 Then
                GetMem2 i + 2, P1
                GetMem2 i + 2, p3
                If ChrW(P1) + ChrW(p3) = ")@" Then
                    r$ = r$ & "()."
                    i = i + 4
                Else
                    GoTo i1233
                End If
                            Else
i1233:
                                       Select Case rr&
                                       Case 1
                                       rr& = 5 ' float array or function
                                       Case 3
                                       rr& = 6 'string array or function
                                       Case 4
                                       rr& = 7 ' long array
                                       Case Else
                                       Exit For
                                       End Select
                     GetMem2 i, P1
                                        r$ = r$ & ChrW(P1)
                                        i = i + 2
                                      ' A$ = Mid$(A$, 2)
                                   Exit For
                            
                          End If
               Else
                        Exit For
            
            End If
        Case Else
        Exit For
  End Select

        Else
         If one Then
              Exit For
              Else
              gr = True
              r$ = r$ & ChrW(P1)
              rr& = 1 'is an identifier or floating point variable
              End If
    End If


    Next i
  If i > p4 Then A$ = "" Else If (i + 2 - p2) \ 2 > 1 Then A$ = Mid$(A$, (i + 2 - p2) \ 2)
       r$ = myUcase(r$, gr)
       lang = 1 + CLng(gr)

    IsLabelAnew = rr&


End Function
Public Function IsLabelDotSub(where$, A$, rrr$, r$, lang As Long) As Long
' for left side...no &

Dim rr&, one As Boolean, c$, firstdot$, gr As Boolean

rrr$ = ""
r$ = ""
Dim i&, l As Long, p3 As Integer
Dim p2 As Long, P1 As Integer, p4 As Long '', excludesp As Long
  l = Len(A$): If l = 0 Then IsLabelDotSub = 0: lang = 1: Exit Function
p2 = StrPtr(A$): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, P1
  Select Case P1
    Case 13
    
    If i < p4 Then
    GetMem2 i + 2, p3
    If p3 = 10 Then
    IsLabelDotSub = 1234
    If i + 6 > p4 Then
    A$ = ""
    Else
    i = i + 4
    Do While i < p4

    GetMem2 i, P1
    If P1 = 32 Or P1 = 160 Then
    i = i + 2
    Else
    GetMem2 i + 2, p3
    If P1 <> 13 And p3 <> 10 Then Exit Do
    i = i + 4
    End If
    Loop
    A$ = Mid$(A$, (i + 2 - p2) \ 2)
    End If
    Else
    If i > p2 Then A$ = Mid$(A$, (i - 2 - p2) \ 2)
    End If
    Else
    If i > p2 Then A$ = Mid$(A$, (i - 2 - p2) \ 2)
    End If
    
    lang = 1
    Exit Function
    Case 32, 160
    Case Else
     ''excludesp = (i - p2) \ 2
   Exit For
  End Select
  Next i
  
  If i > p4 Then A$ = "": IsLabelDotSub = 0: Exit Function
  
  For i = i To p4 Step 2
  GetMem2 i, P1
  If P1 < 256 Then
  Select Case ChrW(P1)
    Case "@"
            If i < p4 And r$ <> "" Then
            GetMem2 i + 2, P1
            If ChrW(P1) <> "(" Then
              where$ = r$
            r$ = ""
            rrr$ = ""
            Else
              IsLabelDotSub = 0: A$ = firstdot$ + Mid$(A$, (i - p2) \ 2): Exit Function
            End If
            Else
              IsLabelDotSub = 0: A$ = firstdot$ + Mid$(A$, (i - p2) \ 2): Exit Function
            End If
    Case "?"
        If r$ = "" And firstdot$ = "" Then
        rrr$ = "?"
        r$ = rrr$
        i = i + 4
        A$ = Mid$(A$, (i - p2) \ 2) ' mid$(a$, 2)
        IsLabelDotSub = 1
        lang = 1 + CLng(gr)
      
        Exit Function
    
        ElseIf firstdot$ = "" Then
        IsLabelDotSub = 1
        lang = 1 + CLng(gr)
        If lang = 1 Then
        rrr$ = UCase(r$)
        Else
        rrr$ = myUcase(r$)
        End If
    
        A$ = Mid$(A$, (i + 2 - p2) \ 2)
        Exit Function
        Else
        IsLabelDotSub = 0
        A$ = Mid$(A$, (i + 2 - p2) \ 2)
        Exit Function
        End If
    Case "."
            If one Then
            Exit For
            ElseIf r$ <> "" And i < p4 Then
            GetMem2 i + 2, P1
            If ChrW(P1) = "." Or ChrW(P1) = " " Then
            If ChrW(P1) = "." And i + 2 < p4 Then
            
                GetMem2 i + 4, P1
                If ChrW(P1) = " " Then i = i + 4: Exit For
            Else
                i = i + 2
               Exit For
            End If
            End If
            GetMem2 i, P1
            r$ = r$ & ChrW(P1)
            ''A$ = Mid$(A$, 2)
            rr& = 1
            Else
            firstdot$ = firstdot$ + "."
            'A$ = Mid$(A$, 2)
            End If
        Case "\", "{" To "~", "^"
            Exit For
        Case "0" To "9", "_"
           If one Then

            Exit For
            ElseIf r$ <> "" Then
            r$ = r$ & ChrW(P1)
            '' A$ = Mid$(A$, 2)
            rr& = 1 'is an identifier or floating point variable
            Else
            Exit For
            End If
        Case Is >= "A"
            If one Then
            Exit For
            Else
            r$ = r$ & ChrW(P1)
            rr& = 1 'is an identifier or floating point variable
            End If
        Case "$"
            If one Then Exit For
            If r$ <> "" Then
            one = True
            rr& = 3 ' is string variable
            r$ = r$ & ChrW(P1)
            Else
            Exit For
            End If
        Case "%"
            If one Then Exit For
            If r$ <> "" Then
            one = True
            rr& = 4 ' is long variable
            r$ = r$ & ChrW(P1)
            Else
            Exit For
            End If
    Case "("
            If r$ <> "" Then
            If i + 4 <= p4 Then
                GetMem2 i + 2, P1
                GetMem2 i + 2, p3
                If ChrW(P1) + ChrW(p3) = ")@" Then
                    r$ = r$ & "()."
                    i = i + 4
                Else
                    GoTo i123
                End If
                            Else
i123:
                                       Select Case rr&
                                       Case 1
                                       rr& = 5 ' float array or function
                                       Case 3
                                       rr& = 6 'string array or function
                                       Case 4
                                       rr& = 7 ' long array
                                       Case Else
                                       Exit For
                                       End Select
                     GetMem2 i, P1
                                        r$ = r$ & ChrW(P1)
                                        i = i + 2
                                      ' A$ = Mid$(A$, 2)
                                   Exit For
                            
                          End If
               Else
                        Exit For
            
            End If
        Case Else
        Exit For
  End Select
  Else
    If one Then
              Exit For
              Else
              gr = True
              r$ = r$ & ChrW(P1)
              rr& = 1 'is an identifier or floating point variable
              End If
    End If
  Next i
  If i > p4 Then A$ = "" Else If (i + 2 - p2) \ 2 > 1 Then A$ = Mid$(A$, (i + 2 - p2) \ 2)
       rrr$ = firstdot$ + myUcase(r$, gr)
       lang = 1 + CLng(gr)
    IsLabelDotSub = rr&
   'a$ = LTrim(a$)

End Function

Public Function NLtrim$(A$)
If Len(A$) > 0 Then NLtrim$ = Mid$(A$, MyTrimL(A$))
End Function

Public Function allcommands(aHash As sbHash) As Boolean
Dim mycommands(), i As Long
mycommands() = Array("ABOUT", "AFTER", "APPEND", "APPEND.DOC", "BACK", "BACKGROUND", "BASE", "BEEP", "BITMAPS", "BOLD", "BREAK", "BROWSER", "CALL", "CASE", "CAT", "CHANGE", "CHARSET", "CHOOSE.COLOR", "CHOOSE.FONT", "CHOOSE.ORGAN", "CIRCLE", "CLASS", "CLEAR", "CLIPBOARD", "CLOSE", "CLS", "CODEPAGE", "COLOR", "COMMIT", "COMPRESS", "CONTINUE", "COPY", "CURSOR", "CURVE", "DATA", "DB.PROVIDER", "DB.USER" _
, "DECLARE", "DEF", "DELETE", "DESKTOP", "DIM", "DIR", "DIV", "DO", "DOCUMENT", "DOS", "DOUBLE", "DRAW", "DRAWINGS", "DROP", "DURATION", "EDIT", "EDIT.DOC", "ELSE", "ELSE.IF", "EMPTY", "END", "ERASE", "ERROR", "ESCAPE", "EVENT", "EVERY", "EXECUTE", "EXIT", "EXPORT", "FAST", "FIELD", "FILES", "FILL", "FIND", "FKEY", "FLOODFILL", "FLUSH", "FONT", "FOR", "FORM", "FRAME", "FUNCTION", "GET", "GLOBAL" _
, "GOSUB", "GOTO", "GRADIENT", "GREEK", "GROUP", "HEIGHT", "HELP", "HEX", "HIDE", "HOLD", "HTML", "ICON", "IF", "IMAGE", "INLINE", "INPUT", "INSERT", "ITALIC", "JOYPAD", "KEYBOARD", "LATIN", "LAYER", "LEGEND", "LET", "LINE", "LINESPACE", "LINK", "LIST", "LOAD", "LOAD.DOC", "LOCAL", "LOCALE", "LONG", "LOOP", "MAIN.TASK", "MARK", "MEDIA", "MENU", "MERGE.DOC", "METHOD", "MODE", "MODULE" _
, "MODULES", "MONITOR", "MOTION", "MOTION.W", "MOUSE.ICON", "MOVE", "MOVIE", "MOVIES", "MUSIC", "NAME", "NEW", "NEXT", "NORMAL", "ON", "OPEN", "OPEN.FILE", "OPEN.IMAGE", "OPTIMIZATION", "ORDER", "OVER", "OVERWRITE", "PAGE", "PART", "PATH", "PEN", "PIPE", "PLAY", "PLAYER", "POLYGON", "PRINT", "PRINTER", "PRINTING", "PROFILER", "PROPERTIES", "PUSH", "PUT", "READ", "RECURSION.LIMIT" _
, "REFER", "REFRESH", "RELEASE", "REM", "REMOVE", "REPEAT", "REPORT", "RESTART", "RETRIEVE", "RETURN", "SAVE", "SAVE.AS", "SAVE.DOC", "SCAN", "SCORE", "SCREEN.PIXELS", "SCRIPT", "SCROLL", "SEARCH", "SEEK", "SELECT", "SET", "SETTINGS", "SHIFT", "SHIFTBACK", "SHOW", "SLOW", "SORT", "SOUND", "SOUNDREC", "SOUNDS", "SPEECH", "SPLIT", "SPRITE", "STACK", "START", "STATIC", "STEP", "STOCK", "STOP", "STRUCTURE" _
, "SUB", "SUBDIR", "SWAP", "SWEEP", "SWITCHES", "TAB", "TABLE", "TARGET", "TARGETS", "TASK.MAIN", "TEST", "TEXT", "THEN", "THREAD", "THREAD.PLAN", "THREADS", "TITLE", "TONE", "TRY", "TUNE", "UPDATE", "USE", "USER", "VERSION", "VIEW", "VOLUME", "WAIT", "WHILE", "WIDTH", "WIN", "WINDOW", "WITH", "WORDS", "WRITE", "WRITER", "адеиасе", "аккацг", "аккане", "аккиыс", "аккиыс.ам", "ам", "амафгтгсг" _
, "амахеыягсг", "амайтгсг", "амакоцио", "амакусг.охомгс", "амакутгс", "амаломг", "амамеысг", "амажояа", "амаье", "амехесе", "амоицла.аявеиоу", "амоицла.еийомас", "амоине", "амтецяаье", "амтицяаье", "апедысе", "апо", "апохгйеусг.ыс", "апойопг", "аяца", "аявеиа", "аявеио", "аявг", "аукос", "ауноуса", "ажаияесг", "ажгсе", "баке", "басг", "басг.паяовос", "басг.вягстгс", "баье", "бектистопоигсг" _
, "бгла", "богхеиа", "цецомос", "целисе", "цемийес", "цемийг", "цемийо", "циа", "цяаллатосеияа", "цяаллг", "цяаье", "цягцояа", "деийтг.лояжг", "деине", "дейаен", "дес", "диабасе", "диацяажг", "диайопг", "диайоптес", "диалесоу", "диаяйеиа", "диастиво", "диажамеиа", "диажамо", "диажуцг", "диейоье", "диояхысе", "дипка", "дипкос", "дойилг", "долг", "дяолеас", "дысе", "еццяажо", "еийома", "еийомес", "еийомидио" _
, "еисацыцг", "ейдосг", "ейтекесг", "ейтупысг", "ейтупытгс", "екецвос", "еккгмийа", "емхесг", "емтасг", "емы", "емысе", "енацыцг", "енодос", "епамакабе", "епамекабе", "епекене", "епекене.цяаллатосеияа", "епекене.ояцамо", "епекене.вяыла", "епицяажг", "епийаияо", "епикене", "епикене.цяаллатосеияа", "епикене.ояцамо", "епикене.вяыла", "епикоцес", "епикоцг", "епикоцгс" _
, "епипедо", "епистяожг", "епижамеиа", "еполемо", "еуяесг", "глеяолгмиа", "гвоцяажгсг", "гвои", "гвос", "хесе", "хесг", "идиотгтес", "исвмг", "ивмос", "йахаяг", "йахаяо", "йахе", "йакесе", "йалпукг", "йаме", "йамомийа", "йат", "йатакоцои", "йатакоцос", "йатавыягсг", "йеилемо", "йемг", "йимгсг", "йимгсг.п", "йкасг", "йкеиди", "йкеисе", "йомсока", "йяата", "йяатгсе", "йяуье" _
, "йуйкийа", "йуйкос", "йукисг", "йуяио.еяцо", "кабг", "кахос", "катимийа", "кенеис", "киста", "коцос", "лайяус", "ле", "леходос", "лекыдиа", "леяос", "лета", "летахесг", "лмглг", "лоусийг", "лпип", "мео", "мгла", "мглата", "нейима", "охомг", "олада", "омола", "ояио.амадяолгс", "ояисе", "паийтгс", "паине", "памы", "паяахуяо", "паяе", "паяелбокг", "павос", "педио", "пема", "пеяи" _
, "пеяихыяио", "пета", "пимайас", "пимайес", "пкациа", "пкаисио", "пкгйтяокоцио", "покуцымо", "пяос", "пяосхесе.еццяажо", "пяосхгйг", "пяовеияо", "яоутима", "яухлисеис", "с", "саяысе", "сбгсе", "сеияа", "секида", "семаяио", "сгл", "сглади", "сглеио", "статийг", "статийес", "стг", "стгм", "сто", "стой", "стовои", "стовос", "суццяажеас", "суццяажг", "суцвымеусе.еццяажо", "сулпиесг" _
, "сулпкгяысг", "сумаятгсг", "сумевисе", "сумхгла", "сус", "сустгла", "сведиа", "сведио.мглатым", "сыяос", "сысе", "сысе.еццяажо", "таимиа", "таимиес", "танг", "танимолгсг", "текос", "титкос", "тлгла", "тлглата", "томос", "топийа", "топийес", "топийг", "топийо", "тоте", "тупос", "тупысе", "упойатакоцос", "упокоцистг", "жаядиа", "жеяе", "жеяеписы", "жомто", "жояла", "жоятос", "жоятысе" _
, "жоятысе.еццяажо", "жымг", "ваяайтгяес", "ваяане", "вягсг", "вягстг", "вягстгс", "вяыла", "?")
For i = 0 To UBound(mycommands())
    aHash.ItemCreator CStr(mycommands(i)), i + 1
Next i
allcommands = True
End Function

Attribute VB_Name = "Module3"
Option Explicit


Const RC_PALETTE As Long = &H100
Const SIZEPALETTE As Long = 104
Const RASTERCAPS As Long = 38
Dim sapi As Object
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Type PicBmp
    Size As Long
    Type As Long
    hbmp As Long
    hPal As Long
    reserved As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As Long, ByVal hPrinter As Long) As Long
Private Declare Function ResetPrinter Lib "winspool.drv" Alias "ResetPrinterA" (ByVal hPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Type PRINTER_DEFAULTS
        pDataType As String
        pDevMode As DEVMODE
        DesiredAccess As Long
End Type
' New Win95 Page Setup dialogs are up to you
Private Type POINTL
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Private Const MAX_PATH As Long = 1024
Private Const MAX_PATH_UNICODE As Long = 260 * 2 - 1

Private Declare Function GetLongPathName Lib "kernel32" _
   Alias "GetLongPathNameW" _
  (ByVal lpszShortPath As Long, _
   ByVal lpszLongPath As Long, _
   ByVal cchBuffer As Long) As Long
 
Public Function GetLongName(strTest As String) As String
   Dim sLongPath As String
   Dim buff As String
   Dim cbbuff As Long
   Dim result As Long
 
   buff = Space$(MAX_PATH_UNICODE)
   cbbuff = Len(buff)
 
   result = GetLongPathName(StrPtr(strTest), StrPtr(buff), cbbuff)
 
   If result > 0 Then
      sLongPath = Left$(buff, result)
   End If
 
   GetLongName = sLongPath
 
End Function
 


Function PathStrip2root(path$) As String
Dim i As Long
If Len(path$) >= 2 Then
If Mid$(path$, 2, 1) = ":" Then
PathStrip2root = Left$(path$, 2) & "\"
Else
i = InStrRev(path$, Left$(path$, 1))
If i > 1 Then
PathStrip2root = "\" & ExtractPath(Mid$(path$, 2, i))
Else
PathStrip2root = Left$(path$, 1)
End If

End If
End If
End Function

Sub Pprop()
    
    If ThereIsAPrinter = False Then Exit Sub
        
    Dim x As Printer
For Each x In Printers
If x.DeviceName = pname And x.port = port Then Exit For
Next x
Dim gp As Long, Td As PRINTER_DEFAULTS
Call OpenPrinter(x.DeviceName, gp, Td)
If form5iamloaded Then
Call PrinterProperties(Form5.hWnd, gp)
Else
Call PrinterProperties(Form1.hWnd, gp)
End If
Call ResetPrinter(gp, Td)
Call ClosePrinter(gp)
End Sub
Function CreateBitmapPicture(ByVal hbmp As Long, ByVal hPal As Long) As Picture
    Dim r As Long, pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID

    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With pic
        .Size = Len(pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hbmp = hbmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    'Create the picture
    r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function
Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hbmp As Long, hBmpPrev As Long, r As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE

    'Create a compatible device context
    hDCMemory = CreateCompatibleDC(hDCSrc)
    'Create a compatible bitmap
    hbmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    'Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hbmp)

    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    'Does our picture use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        LogPal.palVersion = &H300
        'Number of palette entries
        LogPal.palNumEntries = 256
        'Retrieve the system palette entries
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        'Create the palette
        hPal = CreatePalette(LogPal)
        'Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        'Realize the palette
        r = RealizePalette(hDCMemory)
    End If

    'Copy the source image to our compatible device context
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    'Restore the old bitmap
    hbmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Select the palette
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    'Delete our memory DC
    r = DeleteDC(hDCMemory)

    Set hDCToPicture = CreateBitmapPicture(hbmp, hPal)
End Function

Function DriveType(ByVal path$) As String
    Select Case GetDriveType(path$)
        Case 2
            DriveType = "Μεταθέσιμο"
        Case 3
            DriveType = "Σταθερό"
        Case Is = 4
            DriveType = "Απόμακρο"
        Case Is = 5
            DriveType = "Cd-Rom"
        Case Is = 6
            DriveType = "Προσωρινό στην μνήμη"
        Case Else
            DriveType = "Απροσδιόριστο"
    End Select
End Function

Function DriveTypee(ByVal path$) As String
    Select Case GetDriveType(path$)
        Case 2
            DriveTypee = "Removable"
        Case 3
            DriveTypee = "Drive Fixed"
        Case Is = 4
            DriveTypee = "Remote"
        Case Is = 5
            DriveTypee = "Cd-Rom"
        Case Is = 6
            DriveTypee = "Ram disk"
        Case Else
            DriveTypee = "Unrecognized"
    End Select
End Function
Function DriveSerial(ByVal path$) As Long
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim Serial As Long, VName As String, FSName As String
    'Create buffers
    If Len(path$) = 1 Then path$ = path$ & ":\"
    If Len(path$) = 2 Then path$ = path$ & "\"
    VName = String$(255, Chr$(0))
    FSName = String$(255, Chr$(0))
    'Get the volume information
    GetVolumeInformation path$, VName, 255, Serial, 0, 0, FSName, 255
    'Strip the extra chr$(0)'s
    'VName = Left$(VName, InStr(1, VName, Chr$(0)) - 1)
    'FSName = Left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
 DriveSerial = Serial
End Function

Function WeCanWrite(ByVal path$) As Boolean
Dim SecondTry As Boolean, PP$
On Error GoTo wecant
PP$ = ExtractPath(path$, , True)
PP$ = GetDosPath(PP$)
If PP$ = "" Then
MyEr "Not writable device " & path$, "Δεν μπορώ να γράψω στη συσκευή " & path$
Exit Function
End If
PP$ = PathStrip2root(path$)


    Select Case GetDriveType(PP$)

        Case 2, 3, 4, 6
          WeCanWrite = Not GetAttr(PP$) = vbReadOnly
        Case 5
           WeCanWrite = False
    End Select
   Exit Function
wecant:
                   If Err.Number > 0 Then
                Err.clear
                 MyEr "Not writable device " & path$, "Δεν μπορώ να γράψω στη συσκευή " & path$
            WeCanWrite = False
                Exit Function
                End If

End Function
Public Function VoiceName(ByVal d As Double) As String
On Error Resume Next
Dim o As Object
If Typename(sapi) = "Nothing" Then Set sapi = CreateObject("sapi.spvoice")
If Typename(sapi) = "Nothing" Then VoiceName = "": Exit Function
d = Int(d)
If sapi.getvoices().Count >= d And d > 0 Then
For Each o In sapi.getvoices
d = d - 1
If d = 0 Then VoiceName = o.GetDescription: Exit For
Next o
End If
End Function
Public Function NumVoices() As Long
On Error Resume Next
If Typename(sapi) = "Nothing" Then Set sapi = CreateObject("sapi.spvoice")
If Typename(sapi) = "Nothing" Then NumVoices = -1: Exit Function
If sapi.getvoices().Count > 0 Then
NumVoices = sapi.getvoices().Count
End If
End Function
Public Sub SPEeCH(ByVal A$, Optional BOY As Boolean = False, Optional ByVal vNumber As Long = -1)
Static lastvoice As Long
If vNumber = -1 Then vNumber = lastvoice
On Error Resume Next
If vNumber = 0 Then vNumber = 1
If Typename(sapi) = "Nothing" Then Set sapi = CreateObject("sapi.spvoice")
If Typename(sapi) = "Nothing" Then Beep: Exit Sub
If sapi.getvoices().Count > 0 Then
If sapi.getvoices().Count <= vNumber Or sapi.getvoices().Count < 1 Then vNumber = 1
 With sapi
         Set .Voice = .getvoices.item(vNumber - 1)
       If BOY Then
         .volume = vol
        
         .Rate = 2
       ' boy
         .speak "<pitch absmiddle='25'>" & A$
         Else
         
         'man
       .Rate = 1
       .volume = vol
         .speak "<pitch absmiddle='-5'>" & A$
         End If
       End With
       lastvoice = vNumber
End If
End Sub



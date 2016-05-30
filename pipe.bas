Attribute VB_Name = "Module4"
Option Explicit
Public mSysHandlerWasSet
Public clickMe As Long, clickMe2 As Long
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_NO_BUFFERING = &H20000000
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000

Public Const PIPE_ACCESS_DUPLEX = &H3
Public Const PIPE_ACCESS_INBOUND = &H1
Public Const PIPE_READMODE_MESSAGE = &H2
Public Const PIPE_TYPE_MESSAGE = &H4
Public Const PIPE_WAIT = &H0
Public Const PIPE_NOWAIT = &H1
Public Const WRITE_DAC = &H40000
Public Const PIPE_READMODE_BYTE = &H0
Public Const PIPE_TYPE_BYTE = &H0
Public Const ERROR_NO_DATA = 232&
Public Const NMPWAIT_USE_DEFAULT_WAIT = &H0
Public Const ERROR_PIPE_LISTENING = 536&
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINESCROLL = &HB6
Private Const EM_GETLINECOUNT = 186
Public Const INVALID_HANDLE_VALUE = -1
Declare Function CreateNamedPipe Lib "KERNEL32" Alias "CreateNamedPipeW" (ByVal lpName As Long, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As Any) As Long
Declare Function ConnectNamedPipe Lib "KERNEL32" (ByVal hNamedPipe As Long, lpOverlapped As Long) As Long
Declare Function DisconnectNamedPipe Lib "KERNEL32" (ByVal hNamedPipe As Long) As Long
Declare Function WriteFile Lib "KERNEL32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long) As Long
 Declare Function ReadFile Lib "KERNEL32" (ByVal hFile As Long, lpBuffer As Any, _
      ByVal nNumberOfBytesToRead As Long, _
      lpNumberOfBytesRead As Long, _
      lpOverlapped As Any) As Long
Declare Function WaitNamedPipe Lib "KERNEL32" Alias "WaitNamedPipeW" (ByVal lpNamedPipeName As Long, ByVal nTimeOut As Long) As Long
      
'Declare Function CallNamedPipe Lib "KERNEL32" Alias "CallNamedPipeW" (ByVal lpNamedPipeName As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, ByVal nTimeOut As Long) As Long
Declare Function GetLastError Lib "KERNEL32" () As Long
Declare Function CopyFile Lib "KERNEL32" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
'' MoveFile
Declare Function MoveFile Lib "KERNEL32" Alias "MoveFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Declare Function FlushFileBuffers Lib "KERNEL32" (ByVal hFile As Long) As Long
Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3
Public Const CREATE_NEW = 1
Private Declare Sub CopyMemoryStr Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
   (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
    
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
    Private Const LWA_Defaut         As Long = &H2
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Const GW_HWNDNEXT = 2
Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_CUT As Long = &H300
Private Const WM_COPY As Long = &H301
Private Const WM_PAST As Long = &H302
Private Const EM_CANUNDO = &HC6
Private Const WM_USER = &H400
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_UNDO = WM_USER + 23
Public defWndProc As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
                
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long
Private Declare Function SendMessageAsLong Lib "user32" _
       Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
       ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const SND_APPLICATION = &H80 ' look for application specific association
Private Const SND_ALIAS = &H10000 ' name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000 ' name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1 ' play asynchronously
Private Const SND_FILENAME = &H20000 ' name is a file name
Private Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4 ' lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Private Const SND_NOSTOP = &H10 ' don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000 ' don't wait if the driver is busy
Private Const SND_PURGE = &H40 ' purge non-static events for task
Private Const SND_RESOURCE = &H40004 ' name is a resource name or atom
Private Const SND_SYNC = &H0 ' play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function GetShortPathName Lib "KERNEL32" Alias _
"GetShortPathNameW" (ByVal lpszLongPath As Long, _
ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Type SHFILEOPSTRUCTW
    hWnd As Long
    wFunc As Long
    pFrom As Long 'String
    pTo As Long 'String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As Long  'String
End Type

Public Enum FOACTION
    FO_MOVE = &H1
    FO_COPY = &H2
    FO_DELETE = &H3
    FO_RENAME = &H4
End Enum

Public Enum FOFACTION
    FOF_ALLOWUNDO = &H40
    FOF_CONFIRMMOUSE = &H2
    FOF_FILESONLY = &H80
    FOF_MULTIDESTFILES = &H1
    FOF_NO_CONNECTED_ELEMENTS = &H2000
    FOF_NOCONFIRMATION = &H10
    FOF_NOCONFIRMMKDIR = &H200
    FOF_NOCOPYSECURITYATTRIBS = &H800
    FOF_NOERRORUI = &H400
    FOF_NORECURSION = &H1000
    FOF_RENAMEONCOLLISION = &H8
    FOF_SILENT = &H4
    FOF_SIMPLEPROGRESS = &H100
    FOF_WANTMAPPINGHANDLE = &H20
    FOF_WANTNUKEWARNING = &H4000
End Enum
   
   
Private Declare Function CreateDirectory Lib "KERNEL32" Alias "CreateDirectoryW" (ByVal lpszPath As Long, ByVal lpSA As Long) As Long
   
Public Function PathMakeDirs(ByVal Pathd As String) As Boolean
        Pathd = PurifyPath(Pathd)
      PathMakeDirs = 0 <> CreateDirectory(StrPtr(Pathd), 0&)
   End Function
 
Function PurifyPath(Spath$) As String
Dim A$(), i
If Spath$ = "" Then Exit Function
A$() = Split(Spath, "\")
If isdir(A$(LBound(A$()))) Then i = i + 1
For i = LBound(A$()) + i To UBound(A$())
A$(i) = PurifyName(A$(i))
Next i
If LBound(A()) = UBound(A()) Then
PurifyPath = A$(UBound(A$()))
Else
PurifyPath = ExtractPath(Join(A$, "\") & "\", False)
End If
End Function
Function PurifyName(sStr As String) As String
Const noValidcharList = "\<>:/|"
Dim A$, i As Long, ddt As Boolean
If Len(sStr) > 0 Then
For i = 1 To Len(sStr)
If InStr(noValidcharList, Mid$(sStr, i, 1)) = 0 Then
A$ = A$ & Mid$(sStr, i, 1)
Else
A$ = A$ & "-"
End If

Next i
End If
PurifyName = A$
End Function

Public Sub FixPath(s$)
Dim frm$
If s$ <> "" Then
If Left$(s$, 1) = "." And Mid$(s$, 2, 1) <> "." Then
s$ = mcd + Mid$(s$, 2)
End If
frm$ = ExtractPath(s$)
If frm$ = "" Then
    s$ = mcd + s$
Else
    If Left$(frm$, 2) = "\\" Or Mid$(frm$, 2, 1) = ":" Then
    'root
    Else
    s$ = userfiles$ + s$
    End If
End If
End If
End Sub
Public Function RenameFile(ByVal sSourceFile As String, ByVal sDesFile As String) As Boolean
Dim f$, fd$, flag As Long
If Not CanKillFile(sSourceFile) Then Exit Function
If ExtractType(sSourceFile) = "" Then sSourceFile = sSourceFile + ".gsb"
If ExtractType(sDesFile) = "" Then
If ExtractNameOnly(sDesFile) = ExtractNameOnly(sSourceFile) Then
sDesFile = ExtractNameOnly(sDesFile) + ".bck"
Else
sDesFile = ExtractNameOnly(sDesFile) + ".gsb"
End If
End If
sSourceFile = CFname(sSourceFile)
If sSourceFile = "" Or CFname(sDesFile) <> "" Then
BadFilename
Exit Function
Else
sDesFile = ExtractPath(sSourceFile) + ExtractName(sDesFile)
End If
If Left$(sSourceFile, 2) <> "\\" Then
f$ = "\\?\" + sSourceFile
Else
f$ = sSourceFile
End If
If Left$(sDesFile, 2) <> "\\" Then
fd$ = "\\?\" + sDesFile
Else
fd$ = sDesFile
End If
flag = 1
RenameFile = 0 <> MoveFile(StrPtr(f$), StrPtr(fd$))

End Function


Public Function CanKillFile(FileName$) As Boolean
FixPath FileName$
If Not IsSupervisor Then
    If Left$(FileName$, 1) = "." Then
        CanKillFile = True
    Else
      If strTemp <> "" Then
            If Not mylcasefILE(strTemp) = mylcasefILE(Left$(FileName$, Len(strTemp))) Then
            CanKillFile = mylcasefILE(userfiles) = mylcasefILE(Left$(FileName$, Len(userfiles)))
            Else
            CanKillFile = True
            End If
        Else
            CanKillFile = mylcasefILE(userfiles) = mylcasefILE(Left$(FileName$, Len(userfiles)))
        End If
    End If
Else
    CanKillFile = True
End If

End Function
Public Function MakeACopy(ByVal sSourceFile As String, ByVal sDesFile As String) As Boolean
If Not CanKillFile(sSourceFile) Then Exit Function
Dim f$, fd$, flag As Long
If Left$(sSourceFile, 2) <> "\\" Then
f$ = "\\?\" + sSourceFile
Else
f$ = sSourceFile
End If
If Left$(sDesFile, 2) <> "\\" Then
fd$ = "\\?\" + sDesFile
Else
fd$ = sDesFile
End If


MakeACopy = 0 <> CopyFile(StrPtr(f$), StrPtr(fd$), flag)
End Function

Public Function NeoUnicodeFile(FileName$) As Boolean
Dim hFile, counter
Dim f$, F1$
Sleep 10
If Not CanKillFile(FileName$) Then Exit Function
If Left$(FileName$, 2) <> "\\" Then
f$ = "\\?\" + FileName$
Else
f$ = FileName$
End If
On Error Resume Next
F1$ = Dir(f$)  '' THIS IS THEWORKAROUND FOR THE PROBLEMATIC CREATIFILE (I GOT SOME HANGS)

hFile = CreateFile(StrPtr(f$), GENERIC_WRITE, ByVal 0, ByVal 0, 2, FILE_ATTRIBUTE_NORMAL, ByVal 0)
FlushFileBuffers hFile
Sleep 10

CloseHandle hFile

NeoUnicodeFile = (CFname(GetDosPath(f$)) <> "")

Sleep 10

'need "\\?\" before

'now we can use the getdospath from normal Open File


End Function

Public Function GetDosPath(LongPath As String) As String

Dim s As String
Dim i As Long
Dim PathLength As Long

        i = Len(LongPath) * 2 + 2

        s = String(1024, 0)

        PathLength = GetShortPathName(StrPtr(LongPath), StrPtr(s), i)

        GetDosPath = Left$(s, PathLength)

End Function

Sub PlaySoundNew(f As String)

If f = "" Then
PlaySound 0&, 0&, SND_PURGE
Else
If ExtractType(f) = "" Then f = f & ".WAV"
f = CFname(f)
PlaySound StrPtr(f), ByVal 0&, SND_FILENAME Or SND_ASYNC
End If
End Sub


       
Public Sub SetTrans(oForm As Form, Optional bytAlpha As Byte = 255, Optional lColor As Long = 0, Optional TRMODE As Boolean = False)
    Dim lStyle As Long
    lStyle = GetWindowLong(oForm.hWnd, GWL_EXSTYLE)
    If Not (lStyle And WS_EX_LAYERED) = WS_EX_LAYERED Then _
        SetWindowLong oForm.hWnd, GWL_EXSTYLE, lStyle Or WS_EX_LAYERED
       
    SetLayeredWindowAttributes oForm.hWnd, lColor, bytAlpha, IIf(TRMODE, LWA_COLORKEY Or LWA_Defaut, LWA_Defaut)
    UpdateWindow oForm.hWnd
End Sub


Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function


Function Trans2pipe(pipe$, what$) As Boolean
Dim s$, IL As Long, ok As Long, hPipe As Long, ok2 As Long
s$ = validpipename(pipe$)
Dim b() As Byte
b() = what$
ReDim Preserve b(Len(what$) * 2 + 20) As Byte
Trans2pipe = True
hPipe = CreateFile(StrPtr(s$), GENERIC_WRITE, ByVal 0, ByVal 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)
If hPipe <> INVALID_HANDLE_VALUE Then
ok2 = WriteFile(hPipe, b(0), Len(what$) * 2, IL, ByVal 0)
ok = WaitNamedPipe(StrPtr(s$), 1000)
ok = GetLastError = 0
Trans2pipe = ok2 > 0 Or ok
Else
Trans2pipe = False
End If
CloseHandle hPipe
Sleep 1
End Function
Function validpipename(ByVal A$) As String
Dim b$
A$ = myUcase(A$)
b$ = Left$(A$, InStr(1, A$, "\pipe\", vbTextCompare))
If b$ = "" Then
validpipename = "\\" & strMachineName & "\pipe\" & A$
Else
validpipename = A$
End If
End Function

Function Included(afile$, simple$) As String
Dim A As Document
On Error GoTo inc1
Dim what As Long
Dim st&, pa&, po&
st& = 1

If simple$ = "" Then
Included = ExtractName(afile$)
Else
    Sleep 1
    Set A = New Document
    
    A.LCID = cLid
    A.ReadUnicodeOrANSI afile$, , what
    If InStr(simple$, vbCr) > 0 Then
    'work with any char but using computer locale
    If InStr(1, A.textDoc, simple$, vbTextCompare) > 0 Then
                Included = ExtractName(afile$)
    End If
    Else
    ' work in paragraphs..
    If A.FindStr(simple$, st&, pa&, po&) > 0 Then
                                                                '   'If InStr(1, A$, simple$, vbTextCompare) > 0 Then
            Included = ExtractName(afile$)
    End If
    End If
    Set A = Nothing

End If
inc1:
End Function

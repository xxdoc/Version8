Attribute VB_Name = "Fcall"
' This is a module from Olaf Schmidt changed for M2000 needs
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef RETVAR As Variant) As Long
Private Declare Function GetProcByName Lib "KERNEL32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "KERNEL32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal nOrdinal As Long) As Long
Private Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLibModule As Long) As Long
Private Declare Function lstrlenA Lib "KERNEL32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "KERNEL32" (ByVal lpString As Long) As Long
Private Declare Sub RtlMoveMemory Lib "KERNEL32" (dst As Any, src As Any, ByVal BLen As Long)

Private Enum CALLINGCONVENTION_ENUM
  cc_fastcall
  CC_CDECL
  CC_PASCAL
  CC_MACPASCAL
  CC_STDCALL
  CC_FPFASTCALL
  CC_SYSCALL
  CC_MPWCDECL
  CC_MPWPASCAL
End Enum

Private LibHdls As New FastCollection, VType(0 To 63) As Integer, VPtr(0 To 63) As Long


Public Function stdCallW(sDll As String, sFunc As String, ByVal RetType As Variant, p() As Variant, j As Long)
Dim v(), HRes As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To j - 1 ''UBound(V)
    If VarType(p(i)) = vbString Then
    v(i) = CLng(StrPtr(p(i)))
    VPtr(i) = VarPtr(v(i))
    VType(i) = vbString
    Else
    VType(i) = VarType(v(i))
    VPtr(i) = VarPtr(v(i))
    End If
    
  Next i
  If Left$(func, 1) = "#" Then
  HRes = DispCallFunc(0, GetFuncPtrOrd(sDll, sFunc), CC_STDCALL, CInt(RetType), j, VType(0), VPtr(0), stdCallW)
  Else
  HRes = DispCallFunc(0, GetFuncPtr(sDll, sFunc), CC_STDCALL, CInt(RetType), j, VType(0), VPtr(0), stdCallW)
  End If
  If HRes Then Err.Raise HRes
' p() = v()
 If Typename(stdCallW) = "Null" Then
 stdCallW = vbEmpty
 End If
End Function


Public Function cdeclCallW(sDll As String, sFunc As String, ByVal RetType As Variant, p() As Variant, j As Long)
Dim i As Long, pFunc As Long, v(), HRes As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To j - 1
    If VarType(p(i)) = vbString Then v(i) = StrPtr(p(i))
    VType(i) = VarType(v(i))
    VPtr(i) = VarPtr(v(i))
  Next i
   If Left$(func, 1) = "#" Then
     HRes = DispCallFunc(0, GetFuncPtrOrd(sDll, sFunc), CC_CDECL, CInt(RetType), j, VType(0), VPtr(0), cdeclCallW)
   Else
  HRes = DispCallFunc(0, GetFuncPtr(sDll, sFunc), CC_CDECL, CInt(RetType), j, VType(0), VPtr(0), cdeclCallW)
  End If
  If HRes Then Err.Raise HRes
  If Typename(cdeclCallW) = "Null" Then
  cdeclCallW = vbEmpty
  End If
End Function

Public Function stdCallA(sDll As String, sFunc As String, ByVal RetType As Variant, ParamArray p() As Variant)
Dim i As Long, pFunc As Long, v(), HRes As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To UBound(v)
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbFromUnicode): v(i) = StrPtr(p(i))
    VType(i) = VarType(v(i))
    VPtr(i) = VarPtr(v(i))
  Next i
  
  HRes = DispCallFunc(0, GetFuncPtr(sDll, sFunc), CC_STDCALL, RetType, i, VType(0), VPtr(0), stdCallA)
  
  For i = 0 To UBound(p) 'back-conversion of the ANSI-String-Results
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbUnicode)
  Next i
  If HRes Then Err.Raise HRes
End Function

Public Function cdeclCallA(sDll As String, sFunc As String, ByVal RetType As VbVarType, ParamArray p() As Variant)
Dim i As Long, pFunc As Long, v(), HRes As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To UBound(v)
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbFromUnicode): v(i) = StrPtr(p(i))
    VType(i) = VarType(v(i))
    VPtr(i) = VarPtr(v(i))
  Next i
  
  HRes = DispCallFunc(0, GetFuncPtr(sDll, sFunc), CC_CDECL, RetType, i, VType(0), VPtr(0), cdeclCallA)
  
  For i = 0 To UBound(p) 'back-conversion of the ANSI-String-Results
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbUnicode)
  Next i
  If HRes Then Err.Raise HRes
End Function

Public Function vtblCall(pUnk As Long, ByVal vtblIdx As Long, ParamArray p() As Variant)
Dim i As Long, v(), HRes As Long
  If pUnk = 0 Then Exit Function

  v = p 'make a copy of the params, to prevent problems with VT_ByRef-Members in the ParamArray
  For i = 0 To UBound(v)
    VType(i) = VarType(v(i))
    VPtr(i) = VarPtr(v(i))
  Next i
  
  HRes = DispCallFunc(pUnk, vtblIdx * 4, CC_STDCALL, vbLong, i, VType(0), VPtr(0), vtblCall)
  If HRes Then Err.Raise HRes
End Function

Public Function GetFuncPtr(sLib As String, sFunc As String) As Long

Dim hlib As Long

    If LibHdls.Find(sLib) Then
        hlib = LibHdls.Value
    Else
     
      hlib = LoadLibrary(sLib)
      If hlib = 0 Then Err.Raise vbObjectError, , "Dll not found (or loadable): " & sLib
      LibHdls.AddKey sLib, hlib
    End If
  'End If
  GetFuncPtr = GetProcByName(hlib, sFunc)
  If GetFuncPtr = 0 Then MyEr "EntryPoint not found: " & sFunc & " in: " & sLib, "EntryPoint not found: " & sFunc & " στο: " & sLib
End Function
Public Sub RemoveDll(sLib As String)
If LibHdls.Find(sLib) Then
    FreeLibrary LibHdls.Value
    LibHdls.RemoveWithNoFind sLib
End If
End Sub

Public Function GetFuncPtrOrd(sLib As String, sFunc As String) As Long
Dim hlib As Long
Dim lfunc As Long

lfunc = val(Mid$(sFunc, 2))

    If LibHdls.Find(sLib) Then
        hlib = LibHdls.Value
    Else
      hlib = LoadLibrary(sLib)
      If hlib = 0 Then Err.Raise vbObjectError, , "Dll not found (or loadable): " & sLib
      LibHdls.AddKey sLib, hlib
    End If
   ' End If
  GetFuncPtrOrd = GetProcByOrdinal(hlib, lfunc)
  If GetFuncPtrOrd = 0 Then MyEr "EntryPoint not found: " & sFunc & " in: " & sLib, "EntryPoint not found: " & sFunc & " στο: " & sLib
End Function
Public Function GetBStrFromPtr(lpSrc As Long, Optional ByVal ANSI As Boolean) As String
Dim SLen As Long
  If lpSrc = 0 Then Exit Function
  If ANSI Then SLen = lstrlenA(lpSrc) Else SLen = lstrlenW(lpSrc)
  If SLen Then GetBStrFromPtr = Space$(SLen) Else Exit Function
      
  Select Case ANSI
    Case True: RtlMoveMemory ByVal GetBStrFromPtr, ByVal lpSrc, SLen
    Case Else: RtlMoveMemory ByVal StrPtr(GetBStrFromPtr), ByVal lpSrc, SLen * 2
  End Select
End Function

Public Sub CleanupLibHandles() 'not really needed - but callable (usually at process-shutdown) to clear things up
Dim LibHdl
LibHdls.ToStart
While LibHdls.Done
    FreeLibrary LibHdls.Value
    LibHdls.NextIndex
Wend
'  For Each LibHdl In LibHdls: FreeLibrary LibHdl: Next
  Set LibHdls = Nothing
End Sub
Function IsWine()
Static www As Boolean, wwb As Boolean, hlib As Long
If www Then
Else
Err.Clear
On Error Resume Next
hlib = LoadLibrary("ntdll")
wwb = GetProcByName(hlib, "wine_get_version") <> 0
If hlib <> 0 Then FreeLibrary hlib
If Err.number > 0 Then wwb = False
www = True
End If
IsWine = wwb
End Function


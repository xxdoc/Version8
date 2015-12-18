Attribute VB_Name = "Fcall"
' This is a module from Olaf Schmidt changed for M2000 needs


Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Dst As Any, Src As Any, ByVal BLen As Long)

Private Enum CALLINGCONVENTION_ENUM
  CC_FASTCALL
  CC_CDECL
  CC_PASCAL
  CC_MACPASCAL
  CC_STDCALL
  CC_FPFASTCALL
  CC_SYSCALL
  CC_MPWCDECL
  CC_MPWPASCAL
End Enum

Private LibHdls As New Collection, VType(0 To 63) As Integer, VPtr(0 To 63) As Long

Public Function stdCallW(sDll As String, sFunc As String, ByVal RetType As VbVarType, p() As Variant, j As Long)
Dim V(), HRes As Long
 
  V = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To j - 1 ''UBound(V)
    If VarType(p(i)) = vbString Then V(i) = StrPtr(p(i))
    VType(i) = VarType(V(i))
    VPtr(i) = VarPtr(V(i))
  Next i
  
  HRes = DispCallFunc(0, GetFuncPtr(sDll, sFunc), CC_STDCALL, RetType, j, VType(0), VPtr(0), stdCallW)
  If HRes Then Err.Raise HRes
End Function


Public Function cdeclCallW(sDll As String, sFunc As String, ByVal RetType As VbVarType, ParamArray p() As Variant)
Dim i As Long, pFunc As Long, V(), HRes As Long
 
  V = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To UBound(V)
    If VarType(p(i)) = vbString Then V(i) = StrPtr(p(i))
    VType(i) = VarType(V(i))
    VPtr(i) = VarPtr(V(i))
  Next i
  
  HRes = DispCallFunc(0, GetFuncPtr(sDll, sFunc), CC_CDECL, RetType, i, VType(0), VPtr(0), cdeclCallW)
  If HRes Then Err.Raise HRes
End Function

Public Function stdCallA(sDll As String, sFunc As String, ByVal RetType As VbVarType, ParamArray p() As Variant)
Dim i As Long, pFunc As Long, V(), HRes As Long
 
  V = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To UBound(V)
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbFromUnicode): V(i) = StrPtr(p(i))
    VType(i) = VarType(V(i))
    VPtr(i) = VarPtr(V(i))
  Next i
  
  HRes = DispCallFunc(0, GetFuncPtr(sDll, sFunc), CC_STDCALL, RetType, i, VType(0), VPtr(0), stdCallA)
  
  For i = 0 To UBound(p) 'back-conversion of the ANSI-String-Results
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbUnicode)
  Next i
  If HRes Then Err.Raise HRes
End Function

Public Function cdeclCallA(sDll As String, sFunc As String, ByVal RetType As VbVarType, ParamArray p() As Variant)
Dim i As Long, pFunc As Long, V(), HRes As Long
 
  V = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To UBound(V)
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbFromUnicode): V(i) = StrPtr(p(i))
    VType(i) = VarType(V(i))
    VPtr(i) = VarPtr(V(i))
  Next i
  
  HRes = DispCallFunc(0, GetFuncPtr(sDll, sFunc), CC_CDECL, RetType, i, VType(0), VPtr(0), cdeclCallA)
  
  For i = 0 To UBound(p) 'back-conversion of the ANSI-String-Results
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbUnicode)
  Next i
  If HRes Then Err.Raise HRes
End Function

Public Function vtblCall(pUnk As Long, ByVal vtblIdx As Long, ParamArray p() As Variant)
Dim i As Long, V(), HRes As Long
  If pUnk = 0 Then Exit Function

  V = p 'make a copy of the params, to prevent problems with VT_ByRef-Members in the ParamArray
  For i = 0 To UBound(V)
    VType(i) = VarType(V(i))
    VPtr(i) = VarPtr(V(i))
  Next i
  
  HRes = DispCallFunc(pUnk, vtblIdx * 4, CC_STDCALL, vbLong, i, VType(0), VPtr(0), vtblCall)
  If HRes Then Err.Raise HRes
End Function

Public Function GetFuncPtr(sDll As String, sFunc As String) As Long
Static hLib As Long, sLib As String
  If sLib <> sDll Then 'just a bit of caching, to make resolving libHdls faster
    sLib = sDll
    On Error Resume Next
      hLib = 0
      hLib = LibHdls(sLib)
    On Error GoTo 0
    
    If hLib = 0 Then
      hLib = LoadLibrary(sLib)
      If hLib = 0 Then Err.Raise vbObjectError, , "Dll not found (or loadable): " & sLib
      LibHdls.Add hLib, sLib '<- cache it under the dll-name for the next call
    End If
  End If
  GetFuncPtr = GetProcAddress(hLib, sFunc)
  If GetFuncPtr = 0 Then MyEr "EntryPoint not found: " & sFunc & " in: " & sLib, "EntryPoint not found: " & sFunc & " στο: " & sLib
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
  For Each LibHdl In LibHdls: FreeLibrary LibHdl: Next
  Set LibHdls = Nothing
End Sub


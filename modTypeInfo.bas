Attribute VB_Name = "modTypeInfo"
Option Explicit
Private Declare Function lstrlenA Lib "KERNEL32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "KERNEL32" (ByVal lpString As Long) As Long
Private Declare Sub RtlMoveMemory Lib "KERNEL32" (dst As Any, src As Any, ByVal BLen As Long)

' modTypeInfo - enumerate object members and get member infos
'
' low level COM project by [rm] 2005


Private Type TYPEATTR
    guid(15)                As Byte
    LCID_DEF                    As Long
    dwReserved              As Long
    memidConstructor        As Long
    memidDestructor         As Long
    pstrSchema              As Long
    cbSizeInstance          As Long
    typekind                As Long
    cFuncs                  As Integer
    cVars                   As Integer
    cImplTypes              As Integer
    cbSizeVft               As Integer
    cbAlignment             As Integer
    wTypeFlags              As Integer
    wMajorVerNum            As Integer
    wMinorVerNum            As Integer
    tdescAlias              As Long
    idldescType             As Long
End Type

Private Type FUNCDESC
    memid                   As Long
    lprgscode               As Long
    lprgelemdescParam       As Long
    funcking                As Long
    invkind                 As Long
    callconv                As Long
    cParams                 As Integer
    cParamsOpt              As Integer
    oVft                    As Integer
    cScodes                 As Integer
    elemdescFunc            As Long
    wFuncFlags              As Integer
End Type

Public Type fncinf
    name                    As String
    addr                    As Long
    params                  As Integer
End Type

Public Type enmeinf
    name                    As String
    invkind                 As invokekind
    params                  As Integer
End Type


Private Declare Sub CpyMem Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal dwLen As Long)


Private Declare Sub SysFreeString Lib "oleaut32" ( _
    ByVal bstr As Long)

Public Enum invokekind
    INVOKE_FUNC = &H1
    INVOKE_PROPERTY_GET = &H2
    INVOKE_PROPERTY_PUT = &H4
    INVOKE_PROPERTY_PUTREF = &H8
End Enum

Public Function GetObjMembers(obj As Object) As enmeinf()
    Dim vtblObj         As Long, ret            As Long
    Dim vtblTpInf       As Long, vtblTpInfV(21) As Long
    Dim pptInfo         As Long, rgBstrNames    As Long
    Dim ppFuncDesc      As Long, fncdsc         As FUNCDESC
    Dim pAttr           As Long, attr           As TYPEATTR
    Dim NameLen         As Long, strName        As String
    Dim pGetTpInf       As Long
    Dim cFnc            As Integer
    Dim oInfo           As Object
    Dim iunk            As IUnknown

    Dim udeRet()        As enmeinf
    ReDim udeRet(0) As enmeinf

    ' VTable of passed object
    vtblObj = ObjPtr(obj)
    CpyMem vtblObj, ByVal vtblObj, 4
    CpyMem pGetTpInf, ByVal vtblObj + 4 * 4, 4

    ' IDispatch->GetTypeInfo
    ret = CallPointer(pGetTpInf, ObjPtr(obj), 0, LCID_DEF, VarPtr(pptInfo))
    If ret Then Exit Function

    CpyMem oInfo, pptInfo, 4
    Set iunk = oInfo
    CpyMem oInfo, 0&, 4

    ' VTable of ITypeInfo
    CpyMem vtblTpInf, ByVal pptInfo, 4
    ' IUnknown(3) + ITypeInfo(19) = 22
    CpyMem vtblTpInfV(0), ByVal vtblTpInf, 22 * 4

    ' ITypeInfo->GetTypeAttr
    ret = CallPointer(vtblTpInfV(3), pptInfo, VarPtr(pAttr))
    If ret Then Exit Function

    ' get TypeAttributes struct
    CpyMem attr, ByVal pAttr, Len(attr)

    ' ITypeInfo->ReleaseTypeAttr
    ret = CallPointer(vtblTpInfV(19), pptInfo, VarPtr(pAttr))
    If ret Then Debug.Print "Couldn't release TypeAttr"

    ' go through all members
    For cFnc = 0 To attr.cFuncs - 1

        ' ITypeInfo->GetFuncDesc
        ret = CallPointer(vtblTpInfV(5), pptInfo, cFnc, VarPtr(ppFuncDesc))
        If ret Then GoTo NextItem

        ' read function descriptor struct
        CpyMem fncdsc, ByVal ppFuncDesc, Len(fncdsc)

        ' ITypeInfo->ReleaseFuncDesc
        ret = CallPointer(vtblTpInfV(20), pptInfo, ppFuncDesc)

        ' ITypeInfo->GetNames for current member
        ret = CallPointer(vtblTpInfV(12), pptInfo, fncdsc.memid, VarPtr(rgBstrNames), 0, 0, 0)
        If ret Then GoTo NextItem

        ' read its name (Unicode)
        CpyMem NameLen, ByVal rgBstrNames - 4, 4
        strName = Space$(NameLen / 2)
        CpyMem ByVal StrPtr(strName), ByVal rgBstrNames, NameLen
        SysFreeString rgBstrNames

        udeRet(UBound(udeRet)).name = strName
        udeRet(UBound(udeRet)).invkind = fncdsc.invkind
        udeRet(UBound(udeRet)).params = fncdsc.cParams

        ReDim Preserve udeRet(UBound(udeRet) + 1) As enmeinf

NextItem:
    Next

    GetObjMembers = udeRet

    Set iunk = Nothing
End Function

Public Function GetFncInfo(obj As Object, fnc As String) As fncinf
    Dim vtblObj         As Long, ret            As Long
    Dim vtblTpInf       As Long, vtblTpInfV(21) As Long
    Dim pptInfo         As Long, rgBstrNames    As Long
    Dim ppFuncDesc      As Long, fncdsc         As FUNCDESC
    Dim pAttr           As Long, attr           As TYPEATTR
    Dim NameLen         As Long, strName        As String
    Dim pGetTpInf       As Long
    Dim cFnc            As Integer
    Dim oInfo           As Object
    Dim iunk            As IUnknown

    ' VTable of passed object
    vtblObj = ObjPtr(obj)
    CpyMem vtblObj, ByVal vtblObj, 4
    CpyMem pGetTpInf, ByVal vtblObj + 4 * 4, 4

    ' IDispatch->GetTypeInfo
    ret = CallPointer(pGetTpInf, ObjPtr(obj), 0, LCID_DEF, VarPtr(pptInfo))
    If ret Then Exit Function

    CpyMem oInfo, pptInfo, 4
    Set iunk = oInfo
    CpyMem oInfo, 0&, 4

    ' VTable of ITypeInfo
    CpyMem vtblTpInf, ByVal pptInfo, 4
    ' IUnknown(3) + ITypeInfo(19) = 22
    CpyMem vtblTpInfV(0), ByVal vtblTpInf, 22 * 4

    ' ITypeInfo->GetTypeAttr
    ret = CallPointer(vtblTpInfV(3), pptInfo, VarPtr(pAttr))
    If ret Then Exit Function

    ' get TypeAttributes struct
    CpyMem attr, ByVal pAttr, Len(attr)

    ' ITypeInfo->ReleaseTypeAttr
    ret = CallPointer(vtblTpInfV(19), pptInfo, VarPtr(pAttr))

    ' go through all members
    For cFnc = 0 To attr.cFuncs - 1

        ' ITypeInfo->GetFuncDesc
        ret = CallPointer(vtblTpInfV(5), pptInfo, cFnc, VarPtr(ppFuncDesc))
        If ret Then GoTo NextItem

        ' read function descriptor struct
        CpyMem fncdsc, ByVal ppFuncDesc, Len(fncdsc)

        ' ITypeInfo->ReleaseFuncDesc
        ret = CallPointer(vtblTpInfV(20), pptInfo, ppFuncDesc)

        ' ITypeInfo->GetNames for current member
        Dim mmname$
        mmname$ = Space$(200)
          ret = CallPointer(vtblTpInfV(12), pptInfo, fncdsc.memid, StrPtr(mmname$), 0, 0, 0)
        If ret Then GoTo NextItem
strName = GetBStrFromPtr(VarPtr(mmname$))

    
        If StrComp(strName, fnc, vbTextCompare) = 0 Then
            With GetFncInfo
                .name = strName
                .params = fncdsc.cParams
                CpyMem .addr, ByVal vtblObj + fncdsc.oVft, 4
                Exit For
            End With
        End If

NextItem:
    Next

    Set iunk = Nothing
End Function
Private Function GetBStrFromPtr(lpSrc As Long, Optional ByVal ANSI As Boolean) As String
Dim SLen As Long
  If lpSrc = 0 Then Exit Function
  If ANSI Then SLen = lstrlenA(lpSrc) Else SLen = lstrlenW(lpSrc)
  If SLen Then GetBStrFromPtr = Space$(SLen) Else Exit Function
      
  Select Case ANSI
    Case True: RtlMoveMemory ByVal GetBStrFromPtr, ByVal lpSrc, SLen
    Case Else: RtlMoveMemory ByVal StrPtr(GetBStrFromPtr), ByVal lpSrc, SLen * 2
  End Select
End Function
Public Function VTableEntry(obj As Object, ByVal entry As Integer) As Long
    Dim pVTbl       As Long
    Dim lngEntry    As Long

    pVTbl = ObjPtr(obj)
    CpyMem pVTbl, ByVal pVTbl, 4

    CpyMem lngEntry, ByVal pVTbl + &H1C + entry * 4 - 4, 4
    VTableEntry = lngEntry
End Function

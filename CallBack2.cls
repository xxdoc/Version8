VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallBack2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'class for call back
Private objref As basetask
Private mineHere$
Public Event ErrorAtt(ByRef recover As Variant)
Private mObjRef As Object, mName$
Private reghash As New FastCollection

Dim locked As Boolean
Friend Sub NoPublic(bstack As basetask, where$)
        Set objref = New basetask
        Set objref.Sorosref = bstack.soros
        Set objref.Owner = bstack.Owner
        mineHere$ = where$
End Sub
Property Let ShowGui(RHS As Boolean)
If Forms.count = 0 Then Exit Property
If RHS Then
Form1.Visible = True
Form1.WindowState = 0
If Form1.top >= Screen.Height Then Form1.Move 0, 0
Else
If Form1.WindowState = 0 Then
If Not (QRY Or GFQRY Or Forms.count > 5) Then Form1.Visible = False
End If
End If
End Property
Public Sub EventRef(RHS As Object, Optional name As String)
Set mObjRef = RHS
mName$ = name$
End Sub
Sub MethodType(ByVal mName$, mType$)
mName$ = myUcase(mName$)
reghash.AddKey mName$, mType$
End Sub
Public Sub CallGlobal(modulename$, ParamArray a())
If Status <> 0 Then Exit Sub
Dim ohere$
ohere$ = here$
Dim i As Long, mS As New mStiva, ar As mArray, j As Long
Dim manydim As Long, maxitems As Long, probe As Long, v() As Variant, probelow As Long
Dim oldsoros As mStiva
Set oldsoros = objref.soros
Set objref.Sorosref = mS
For i = 0 To UBound(a())
Select Case VarType(a(i))
Case Is >= vbArray
' only with error control
If Typename(a(i)) = "Variant()" Then
Set ar = New mArray
v() = a(i)
On Error Resume Next
For j = 1 To 60
    probe = UBound(v, j)
    If Err Then Exit For
Next j
manydim = j - 1
On Error Resume Next
Dim items As Long
For j = 1 To manydim
    
    probe = UBound(v, j)
    If Err Then Exit For
    probelow = LBound(v, j)
    ar.PushDim probe - probelow + 1
Next j
ar.PushEnd
'ar.RevOrder = True
ar.CopySerialize v()
mS.DataObj ar
Set ar = Nothing
Erase v()
Else
mS.DataValLong 0
End If
Case vbEmpty
mS.DataValLong 0
Case vbVariant, vbString
mS.DataStr CStr(a(i))
Case Else
mS.DataVal CDbl(a(i))
End Select
Next i
PushStage objref, 0
Set mS = Nothing
here$ = ""
If subHash.ExistKey(myUcase(modulename$, True)) Then
    probe = interpret(objref, modulename$)
Else
   
    Dim what
        what = False
    If mObjRef Is Nothing Then
    RaiseEvent ErrorAtt(what)
        If what = True Then Reset
    Else
        On Error Resume Next
        If mName$ <> "" Then
            mObjRef.Run mName$ + "_ErrorAtt", what
            If Err Then Err.Clear: mObjRef.Run mName$ + ":ErrorAtt", what
            
        Else
            mObjRef.Run "ErrorAtt", what
        End If
        Err.Clear
        If what = True Then
            Reset
        Else
            Err.Raise 12345, , "Not Found"
        End If
    End If
    
End If
PopStage objref
Set objref.Sorosref = oldsoros
Set oldsoros = Nothing
here$ = ohere$
End Sub
Public Sub CallModule(modulename$, ParamArray a())
Static once As Boolean, where As Long, what
If Status <> 0 Then Exit Sub
If Not GetSubFullName2(myUcase(modulename$, True), full$, where) Then
'MyEr "Not Found", "��� �������"
    what = False
    
    If mObjRef Is Nothing Then
    RaiseEvent ErrorAtt(what)
       If what = True Then Reset
    Else
        On Error Resume Next
        If mName$ <> "" Then
            mObjRef.Run mName$ + "_ErrorAtt", what
            If Err Then Err.Clear: mObjRef.Run mName$ + ":ErrorAtt", what
            
        Else
            mObjRef.Run "ErrorAtt", what
        End If
        Err.Clear
        If what = True Then
            Reset
        Else
        MyEr "Not Found " + modulename$, "��� ������� " + modulename$
        End If
    End If
Exit Sub
End If

If once Then Exit Sub
Dim ohere$
once = True
ohere$ = here$
here$ = full$


Dim i As Long, mS As New mStiva, ar As mArray, j As Long
Dim manydim As Long, maxitems As Long, probe As Long, v() As Variant, probelow As Long
Dim IhaveSignature As Boolean, sign$, sign2$, dummy$, look As Long
If reghash.Find(myUcase(modulename$, True)) Then
sign$ = reghash.Value
sign2$ = sigh$
IhaveSignature = True
End If
' no byrefernece yet
Dim oldsoros As mStiva
Set oldsoros = objref.soros
Set objref.Sorosref = mS


For i = 0 To UBound(a())

Select Case VarType(a(i))
Case Is >= vbArray
' only with error control
If Typename(a(i)) = "Variant()" Then
Set ar = New mArray
v() = a(i)
On Error Resume Next
For j = 1 To 60
    probe = UBound(v, j)
    If Err Then Exit For
Next j
manydim = j - 1
On Error Resume Next
For j = manydim To 1 Step -1
    
    probe = UBound(v, j)
    If Err Then Exit For
    probelow = LBound(v, j)
    ar.PushDim probe - probelow + 1
Next j
ar.PushEnd
ar.RevOrder = True

ar.CopySerialize v()
''
If IhaveSignature Then
look = IsLabelOnly(sign$, dummy$)
If look = 2 Then ' byref
look = IsLabelOnly(sign$, dummy$)
look = GlobalVar("b" + CStr(var2used) + "." + dummy$, True)
ElseIf look > 4 Then
FastSymbol sign, ")", True
mS.DataObj ar
Else
SyntaxError
GoTo there
End If
Else
mS.DataObj ar
End If
Set ar = Nothing
Erase v()
Else
mS.DataValLong 0
End If
Case vbEmpty
mS.DataValLong 0
Case vbVariant, vbString
mS.DataStr CStr(a(i))
Case Else

mS.DataVal CDbl(a(i))

End Select
Next i
PushStage objref, 0
objref.soros.MergeTop mS

probe = FastCallModule(objref, where)

If probe <> 1 Then
    If mObjRef Is Nothing Then
    RaiseEvent ErrorAtt(what)
       If what = True Then Reset
    Else
        On Error Resume Next
        If mName$ <> "" Then
            mObjRef.Run mName$ + "_ErrorAtt", what
            If Err Then Err.Clear: mObjRef.Run mName$ + ":ErrorAtt", what
            
        Else
            mObjRef.Run "ErrorAtt", what
        End If
        Err.Clear
        If what = True Then
            Reset
        Else
            Err.Raise 12345, , "Not Found"
        End If
    End If

End If

   


there:
PopStage objref

Set objref.Sorosref = oldsoros
Set oldsoros = Nothing
here$ = ohere$
once = False
End Sub

Public Sub ExecuteStatement(ByVal code$)
Dim what$, pos As Long, func As Boolean, param$
If Status <> 0 Then Exit Sub
If IsLabelOnly(code$, what$) > 4 Then
    func = True
ElseIf what$ = "" Then
    Exit Sub
ElseIf LTrim$(code$) = "" Then
' one call
Call Run(what$)
Exit Sub
End If
Do
 pos = 1
  drop$ = aheadstatus(code$, True, pos)
  If drop$ = "" And pos > Len(code$) Then Exit Sub
  If param$ <> "" Then
  param$ = param$ + ", " + Left$(code$, pos - 1)
  Else
  param$ = Left$(code$, pos - 1)
  End If
  code$ = Mid$(code$, pos)
  Loop Until Not FastSymbol(code$, ",")
  If func Then
  Call Run(what$ + " " + param$ + ")")
  Else
  Call Run(what$ + " " + param$)
  End If
End Sub
Public Function Eval(code$) As Variant
Dim pos As Long, p As Double, s As String, drop$
 pos = 1
  drop$ = aheadstatus(code$, True, pos)
  If drop$ = "" And pos > Len(code$) Then Exit Function
  If InStr(drop$, "N") Then
  If IsExp(objref, Left$(code$, pos), p) Then
  If Not objref.lastobj Is Nothing Then
  Set Eval = objref.lastobj
  
  Else
    Eval = p
   End If
  End If
  Else
  If IsStrExp(objref, Left$(code$, pos), s$) Then
  If Not objref.lastobj Is Nothing Then
      Set Eval = objref.lastobj
  Else
    Eval = s$
  End If
  End If
  End If
Set objref.lastobj = Nothing
End Function
Public Sub AddExecCode(code$)
If locked Then Exit Sub
Call Run(code$, False)
End Sub
Public Function Run(ByVal code$, Optional clean As Boolean = True) As Boolean
Dim ohere$, stay As Variant
If Status <> 0 Then Exit Function
If elevatestatus2 Then Exit Function
elevatestatus2 = 1
If clean Then PushStage objref, 0
ohere$ = here$
here$ = mineHere$
NoAction = True
NOEXECUTION = False
objref.toprinter = False
MOUT = False
ResetBreak
Run = True
        LastErNum = 0: LastErNum1 = 0
        LastErName = ""
        LastErNameGR = ""

If Execute(objref, code$, False) <> 1 Then
Run = fasle

If Not App.StartMode = vbSModeStandalone Then
                 MOUT = True
                 ProcTitle objref, Chr$(34) + Chr$(34) + ",0", 0
End If
            
            
                Dim x As Form
                    For Each x In Forms
             
                    If Typename$(x) = "GuiM2000" Then Unload x
                    Next
                    Set x = Nothing
If NERR Then
If mObjRef Is Nothing Then
    RaiseEvent ErrorAtt(stay)
        If stay = True Then Reset Else Err.Raise 12345, , "Not Found"
    Else
        On Error Resume Next
        If mName$ <> "" Then
            mObjRef.Run mName$ + "_ErrorAtt", stay
            If Err Then Err.Clear: mObjRef.Run mName$ + ":ErrorAtt", stay
            
        Else
            mObjRef.Run "ErrorAtt", stay
        End If
        Err.Clear
        If stay = True Then
            Reset
        Else
             Err.Raise 12345, , "Not Found"
        End If
    End If
If Not stay Then GoTo finalize
End If
    objref.toprinter = False
    If MOUT Then
            NOEXECUTION = False
            ResetBreak
            MOUT = interpret(objref, "new: flush: clear")
            MOUT = False
        Else
        
        If NOEXECUTION Then
                closeAll
        End If

        End If
        
MOUT = True
NoAction = False
If ExTarget Then Run = False
End If
finalize:
If NERR Then
' problem
PopStage objref

End If
here$ = ohere$

If clean Then PopStage objref
elevatestatus2 = 0
Dim exityes
If NERR Then NERR = False: Err.Raise 50500, , "Fatsl Error"
End Function
Property Get ErrorGr() As String
ErrorGr = LastErNameGR
End Property
Property Get ErrorEn() As String
ErrorEn = LastErName
End Property
Property Get Status() As Long
If elevatestatus2 Then
Status = -elevatestatus2
Else
Status = LastErNum
End If
End Property
Public Sub Reset()

    NERR = False
        LastErNum = 0: LastErNum1 = 0
        LastErName = ""
        LastErNameGR = ""
 

        
End Sub

Private Sub Class_Initialize()
        Set objref = New basetask
        Set objref.Owner = Form1.DIS
        mineHere$ = "" 'HERE$
End Sub

Private Sub Class_Terminate()
Set mObjRef = Nothing
mName$ = name$
Set objref.Owner = Nothing
Set objref.Sorosref = Nothing
Set objref = Nothing
GarbageFlush
End Sub

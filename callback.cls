VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "callback"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'class for call back

Private objref As basetask
Private mineHere$
Public Event ErrorAtt(ByRef recover As Variant)
Dim locked As Boolean
Friend Sub NoPublic(bstack As basetask, where$)

        Set objref = New basetask
        Set objref.Sorosref = bstack.soros
        Set objref.Owner = bstack.Owner
        mineHere$ = where$
End Sub
Public Sub Cli(LoadFileAndSwitches$, Optional Prompt$ = ">")
Dim ohere$
If elevatestatus Then Exit Sub
elevatestatus = 1
ohere$ = here$
here$ = mineHere$
If LoadFileAndSwitches$ = "" Then newshow objref
NoAction = True
NOEXECUTION = False
objref.toprinter = False
MOUT = False
ResetBreak
LastErNum = 0: LastErNum1 = 0
LastErName = ""
LastErNameGR = ""
s$ = LoadFileAndSwitches$
If Not ISSTRINGA(s$, cLine) Then
cLine = mylcasefILE(Trim(s$))
Else
cLine = mylcasefILE(cLine)
End If
While Left$(cLine, 1) = Chr(34) And Right$(cLine, 1) = Chr(34) And Len(cLine) > 2
cLine = Mid$(cLine, 2, Len(cLine) - 2)
Wend
If ExtractType(cLine) <> "gsb" Then cLine = ""
If cLine <> "" Then
para$ = ExtractPath(cLine) + ExtractName(cLine)
cLine = Trim$(Mid$(cLine, Len(para$) + 1))
s$ = cLine + " " + s$
cLine = para$
ElseIf s$ <> "" Then
para$ = Trim$(s$)
End If



Switches para$

Form1.MyPrompt cLine, Prompt$
finalize:
here$ = ohere$
elevatestatus = 0
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
Public Sub ExecuteStatement(ByVal code$)
Dim what$, pos As Long, func As Boolean, param$
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
If elevatestatus Then Exit Function
elevatestatus = 1
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
                 MOUT = True
End If
If clean Or MOUT Or ExTarget Then
                Dim x As Form
                    For Each x In Forms
             
                    If Typename$(x) = "GuiM2000" Then Unload x
                    Next
                    Set x = Nothing
If NERR Then

RaiseEvent ErrorAtt(stay)
If Not stay Then GoTo finalize
End If
If Not objref Is Nothing Then
    objref.toprinter = False
Else
    Run = False: GoTo finalize
End If
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
here$ = ohere$

If clean Then
If Not objref Is Nothing Then PopStage objref Else Run = False
End If
elevatestatus = 0
End Function
Property Get ErrorGr() As String
ErrorGr = LastErNameGR
End Property
Property Get ErrorEn() As String
ErrorEn = LastErName
End Property
Property Get Status() As Long
If elevatestatus Then
Status = -elevatestatus
Else
Status = LastErNum
End If
End Property
Public Sub Reset()
elevatestatus = 0
    NERR = False
        LastErNum = 0: LastErNum1 = 0
        LastErName = ""
        LastErNameGR = ""
 'If Not App.StartMode = vbSModeStandalone Then
 
NeoSubMain
 
' End If

        
End Sub

Private Sub Class_Initialize()

 If Not l_complete Then NeoSubMain
        Set objref = New basetask
        Set objref.Owner = Form1.DIS
        mineHere$ = "" 'HERE$
End Sub

Private Sub Class_Terminate()
Set objref.Owner = Nothing
Set objref.Sorosref = Nothing
Set objref = Nothing
GarbageFlush
'If App.StartMode = vbSModeStandalone Then Exit Sub
TerminateM2000
terminatefinal
l_complete = False

End Sub

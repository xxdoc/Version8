VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "basetask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Basetask is a module task of M2000
' This is a utility soros for passing parammeters to functions
Public addlen As Long, NoFuncError As Boolean, CallW As String, fHere$, strg As Boolean
Public nokillvars As Boolean
Public IFCTRL As Boolean, jump As Boolean
Public Vars As Long
Public commnum As Long
Public connectnow As Boolean
Public LastExecute As Long
Public LastOnce As Boolean
Public priveflag As Boolean
Private parentSoros As mStiva ' reference to parent..stiva
Private MyParent As basetask  'reference to parent
Private myOwner As Object
Private mtoback As Boolean
Private mtoprinter As Boolean
Private mitalic As Boolean   ' todo
Private mCharSet As Integer  ' todo
Private mBold As Boolean   ' todo
Private mSize As Single    ' TODO
Private mFontname As String 'TODO
Private Mchild As Boolean
Private mtolayer As Long
Private myRun As Boolean
Private mprocess As TaskInterface ' reference to current process maybe same to parrent...
Private mprocessOk As Boolean ' if i have a process class
Private threadsnames As GarbageClass   ' δουλεύει
Public TaskMain As Boolean
Public OriginalName$
Private mGroupName$
Public UseGroupname$
Public LastClick As Long
Public Look2Parent As Boolean
Private Const dotcolmax = 10
Private dotcol(dotcolmax) As String, curdotcol As Long
Public StaticCollection As FastCollection
Private topdotlevel As Long
Public reflimit As Long
Public IndexSub As Long
Public lastobj As Object
Public FuncObj As Object
Public FuncRec As String
Public FuncValue As Variant
Public LastValue As Variant
Public LastMouse1 As Variant
Public LastMouse2 As String
Public LastState As Boolean
Private RetSoros As mStiva  ' for basic mimic functions- for education only
Public SubLevel As Long
Public OriginalCode As Long
Public StaticInUse$
Private mIamlambda As Boolean
Private mIamAnEvent As Boolean
Public NoError As Boolean
Public Sub CopyStrip(bst As basetask)
With bst
.StaticInUse = StaticInUse
.OriginalCode = OriginalCode
Set .StaticCollection = StaticCollection
Set .Sorosref = soros
End With

End Sub
Public Sub SetVar(ByVal name$, var As Variant)

    If Not StaticCollection.Find(name$) Then
   ' If StaticCollection.Count = 1 Then Stop
    StaticCollection.AddKey name$
    End If
   StaticCollection.Value = var
End Sub
Public Sub SetVarobJ(ByVal name$, var As Variant)
    If StaticCollection Is Nothing Then Set StaticCollection = New FastCollection
    If Not StaticCollection.Find(name$) Then StaticCollection.AddKey name$
    Set StaticCollection.ValueObj = var
End Sub

Function AlterVar(ByVal name$, p As Double, ss$, inttype As Boolean) As Boolean

If Not StaticCollection.Find(name$) Then Exit Function
  With StaticCollection
  If .IsObj Then
    If Typename(.ValueObj) = "mHandler" Then
        If .ValueObj.t1 = 3 Then
        Dim mm As mArray
            Set mm = .ValueObj.objref
            If mm.Arr Then
                mm.Compute p, ss$
            End If
            Set mm = Nothing
        End If
    End If
  Else
            Select Case ss$
            Case "!!"
            .Value = -1 - (.Value <> 0)
            Case "++"
            .Value = .Value + 1
            Case "--"
            .Value = .Value - 1
            Case "-!"
            .Value = -.Value
            Case "=", "g"
            .Value = p
            Case "+"
            .Value = p + .Value
            Case "*"
            .Value = p * .Value
            Case "-"
            .Value = .Value - p
            Case "/"
            If p = 0 Then
            
             Exit Function
            End If
            .Value = .Value / p
            Case "!"
            .Value = -1 - (.Value <> 0)
            End Select
            If inttype Then
            .Value = Int(.Value)
            Else
            .Value = MyRound(.Value, 13)
            End If
            End If
            End With
   AlterVar = True
End Function
Function MyRound(ByVal x, Optional d As Double = 0) As Double
Dim i, n
  i = Abs(Int(d)): If i > 13 Then i = 13
  n = Sgn(x) * 0.5

Select Case i
Case 0
MyRound = Fix(x + n)
Case 1
MyRound = Fix(x * 10# + n) / 10#
Case 2
MyRound = Fix(x * 100# + n) / 100#
Case 3
MyRound = Fix(x * 1000# + n) / 1000#
Case 4
MyRound = Fix(x * 10000# + n) / 10000#
Case 5
MyRound = Fix(x * 100000# + n) / 100000#
Case 6
MyRound = Fix(x * 1000000# + n) / 1000000#
Case 7
MyRound = Fix(x * 10000000# + n) / 10000000#
Case 8
MyRound = Fix(x * 100000000# + n) / 100000000#
Case 9
MyRound = Fix(x * 1000000000# + n) / 1000000000#
Case 10
MyRound = Fix(x * 10000000000# + n) / 10000000000#
Case 11
MyRound = Fix(x * 100000000000# + n) / 100000000000#
Case 12
MyRound = Fix(x * 1000000000000# + n) / 1000000000000#
Case 13
MyRound = Fix(x * 10000000000000# + n) / 10000000000000#
End Select
End Function
Public Function ExistVar(ByVal name$) As Boolean
        ExistVar = StaticCollection.ExistKey(name$)
End Function
Public Sub ReadVar(ByVal name$, var As Variant)
    Dim i As Long, vv As Variant
    With StaticCollection
        If .Find(name$) Then
            If .IsObj Then
                    Set var = .ValueObj
                    
          If Left$(name$, 1) = "%" Then Set .ValueObj = Nothing
            Else
                    var = .Value
            End If
        End If
    End With
End Sub
Public Sub DropNdot(n As Long)
If n = 0 Then Exit Sub
If curdotcol - n < 0 Then
    curdotcol = 0
Else
    curdotcol = curdotcol - n
End If
End Sub

Public Sub MoveNameDot(a$)
If a$ = "" Then Exit Sub
If curdotcol < dotcolmax Then
    curdotcol = curdotcol + 1
    dotcol(curdotcol) = a$
    
Else
    MyEr "Max count for dot names", "Μέγιστος αριθμός τελείας"
End If
End Sub
Public Function GetDotNew(a$, dotnum&) As Boolean
If Look2Parent Then GetDotNew = Parent.GetDotNew(a$, dotnum&): Exit Function

If curdotcol = 0 Then
If UseGroupname <> "" Then
 If InStr(a$, "(") > 0 Then
 If AscW(Left$(a$, 1)) > 255 Then
 a$ = "ΑΥΤΟ." + a$
 Else
 a$ = "THIS." + a$
 End If
 Else
 a$ = UseGroupname + a$
 End If

 
 End If
Else
newrun:
If curdotcol < dotnum Then
'' wrong
If UseGroupname = "" Then

Else

            
            a$ = UseGroupname + a$
 End If
Else

  If dotcol(dotnum) = "THIS" Then
    If UseGroupname <> "" Then
      a$ = UseGroupname + a$
    Else
    a$ = "THIS." + a$
    End If
  Else
    a$ = dotcol(dotnum) + "." + a$

    End If
    If Asc(a$) = 46 Then
    dotnum = 1
    While Asc(Mid$(a$, dotnum + 1) + " ") = 46
    dotnum = dotnum + 1
    Wend
    a$ = Mid$(a$, dotnum + 1)
    GoTo newrun
    End If
    End If
    End If
    GetDotNew = True
End Function

Public Function GetDot(a$, st&) As Boolean
Dim i&
If curdotcol = 0 Then a$ = Mid$(a$, 2): Exit Function
Do While Mid$(a$, st&, 1) = " "
st& = st& + 1
If st& > Len(a$) Then Exit Function
Loop

For i& = st& + 1 To Len(a$)
If Mid$(a$, i&, 1) <> "." Then i& = i& + 1: Exit For
Next i&
i& = i& - 2
a$ = Left$(a$, st& - 1) + dotcol(i& - st& + 1) + Mid$(a$, i&)
If Not Err Then GetDot = True
End Function



Public Property Let myitalic(ByVal RHS As Boolean)
mitalic = RHS
End Property

Public Property Get myitalic() As Boolean
myitalic = mitalic
End Property
Public Property Let myCharSet(ByVal RHS As Integer)
mCharSet = RHS
End Property

Public Property Get myCharSet() As Integer
myCharSet = mCharSet
End Property
Public Property Let myBold(ByVal RHS As Boolean)
mBold = RHS
End Property

Public Property Get myBold() As Boolean
myBold = mBold
End Property
Public Property Let toprinter(ByVal RHS As Boolean)
mtoprinter = RHS
End Property

Public Property Get toprinter() As Boolean
toprinter = mtoprinter
End Property
Public Property Let tolayer(ByVal RHS As Long)
mtolayer = RHS
End Property

Public Property Get tolayer() As Long
tolayer = mtolayer
End Property

Public Property Let toback(ByVal RHS As Boolean)
mtoback = RHS
End Property

Public Property Get toback() As Boolean
toback = mtoback
End Property


Public Property Get IamRunning() As Boolean
IamRunning = myRun
End Property
Public Property Get IamChild() As Boolean
IamChild = Mchild
End Property

Public Property Set Parent(dady As basetask)
Set MyParent = dady
Mchild = True
Me.toback = dady.toback
Me.toprinter = dady.toprinter
Me.tolayer = dady.tolayer
' αν θέλω μαζί με την σύνδεση του Parent
' δίνω και την Sorosref από την Soros του Parent (την δική του ή από τον πατέρα του)
End Property
Public Property Get Parent() As basetask
Set Parent = MyParent

End Property
Public Property Set Sorosref(a As mStiva)
Set parentSoros = a
End Property
Public Property Get RetStackTotal() As Long
If RetSoros Is Nothing Then
Else
RetStackTotal = RetSoros.Total
End If
End Property

Public Property Get RetStack() As mStiva
If RetSoros Is Nothing Then Set RetSoros = New mStiva
Set RetStack = RetSoros
End Property
Public Function SorosNothing() As Boolean
SorosNothing = parentSoros Is Nothing
End Function
Public Property Get soros() As mStiva
If parentSoros Is Nothing Then
Set parentSoros = New mStiva
End If
Set soros = parentSoros
End Property
Public Property Set Owner(a As Object)
Set myOwner = a
End Property
Public Property Get Owner() As Object
If mtoprinter Then
Set Owner = Form1.PrinterDocument1
Else
Set Owner = myOwner
End If

On Error Resume Next
Err.Clear
If Not Owner Is Nothing Then
mitalic = Owner.Font.Italic
mBold = Owner.Font.bold
mCharSet = Owner.Font.charset
myRun = True
Else
myRun = False
End If
Err.Clear
End Property
Public Property Get IamThread() As Boolean
IamThread = mprocessOk
End Property
Public Sub EndRun()
' μόνο από το Myprocess.tick
myRun = False
End Sub
Public Property Set Process(a As TaskInterface)
If mprocess Is a Then
Set mprocess = Nothing
mprocessOk = True
Else
Set mprocess = a

mprocessOk = True
End If
End Property
Public Property Get Process() As TaskInterface
Set Process = mprocess
End Property
Private Sub checkialive()
If threadsnames Is Nothing Then Set threadsnames = New GarbageClass
End Sub

Public Property Set LinkThread(Threadid As Long, a As TaskInterface)
Dim aa As mThreadref
checkialive
If threadsnames.Find(Threadid) Then
Set aa = threadsnames.ValueObj
Set aa.thread = a
End If
End Property

Public Sub PushThread(Threadid As Long, threadvar As String)
Dim aa As New mThreadref
checkialive
Set aa.thread = Nothing  ' no thread
aa.Id = Threadid
aa.VarName = threadvar
If Not threadsnames.Find(Threadid) Then threadsnames.AddKey Threadid
Set threadsnames.ValueObj = aa

End Sub
Public Function Exist(Threadid As Long, ThreadName As String) As Boolean
'Dim aa As mThreadref
If threadsnames Is Nothing Then Exist = False: Exit Function
If threadsnames.Find(Threadid) Then
 With threadsnames.ValueObj
    Exist = .VarName = ThreadName
End With
End If

End Function

Public Function ThrowOne(Threadid As Long) As Boolean
Dim aa As mThreadref
On Error Resume Next
If threadsnames Is Nothing Then Exit Function
' this is mandatory
If TaskMaster Is Nothing Then Exit Function
If TaskMaster.ThrowOne(Threadid) Then ThrowOne = True
If Err.Number > 0 Then Err.Clear: Exit Function
' maybe the task isn't on our list - because we send message from other thread!
If threadsnames.Find(Threadid) Then
    Set aa = threadsnames.ValueObj

    threadsnames.Remove Threadid
    threadsnames.PrevIndex
    aa.thread.Dispose tmAllTasks
    ThrowOne = True
    Set aa.thread = Nothing
Else
    Err.Clear
    ThrowOne = Me.Parent.ThrowOne(Threadid)
End If
Set aa = Nothing
End Function
Public Function ThreadsStr(Optional ByVal lang As Long = 0)
'' use threadsnames

Dim aa As mThreadref, a$
If threadsnames Is Nothing Then Exit Function

If threadsnames.Count > 0 Then
threadsnames.ToStart
While threadsnames.Done
' For Each aa In threadsnames
Set aa = threadsnames.ValueObj
If Left$(aa.VarName, 1) = "_" Then
If lang = 0 Then
a$ = a$ & "ΕΣΩΤΕΡΙΚΟ ΝΗΜΑ "
Else
a$ = a$ & "INTERNAL THREAD "
End If
Else
a$ = a$ & aa.VarName + Str$(aa.Id) & " "
End If
threadsnames.NextIndex
Wend
Set aa = Nothing
'Next aa
End If
If Me.IamChild Then
If Not Me.Parent Is Nothing Then
ThreadsStr = Trim$(a$ & Me.Parent.ThreadsStr(lang))
Else
ThreadsStr = ""
End If
Else
ThreadsStr = Trim$(a$)
End If
End Function
Public Function ThreadsNumber() As Long
'' use threadsnames
Dim aa As mThreadref, kk
If threadsnames Is Nothing Then ThreadsNumber = kk: Exit Function
If threadsnames.Count > 0 Then
threadsnames.ToStart
While threadsnames.Done
Set aa = threadsnames.ValueObj
If Left$(aa.VarName, 5) <> "_multi" Then kk = kk + 1
threadsnames.NextIndex
Wend
Set aa = Nothing
End If
ThreadsNumber = kk
End Function
Public Sub ThrowThreads(Optional idexclude As Long = -1)
If threadsnames Is Nothing Then Exit Sub
Dim aa As mThreadref, Tid As Long
Dim thatProcess As TaskInterface
Dim thatBasetask As basetask
CONT:
If threadsnames.Count > 0 Then
threadsnames.ToStart
While threadsnames.Done
Set aa = threadsnames.ValueObj
If idexclude <> aa.Id Then
If Not TaskMaster Is Nothing Then
        If TaskMaster.ThrowOne(aa.Id) Then
        ThrowOne aa.Id
        ElseIf Not IamThread Then
        threadsnames.RemoveWithNoFind
        GoTo CONT
        End If
Else
ThrowOne aa.Id
End If
End If
threadsnames.NextIndex

Wend
Set aa = Nothing
End If

End Sub

Private Sub Class_Initialize()
'Set parentSoros = New mStiva ' maybe this will change
Set Owner = Nothing
Set lastobj = Nothing
Set FuncObj = Nothing
End Sub

Public Function IsOptional() As Boolean
If parentSoros.LookNull Then
    IsOptional = False
ElseIf parentSoros.PopType = ">" Then
    parentSoros.drop 1
    IsOptional = True
End If
End Function
Public Function IsNumber(r As Double) As Boolean
If parentSoros.LookNull Then
r = parentSoros.PopVal
IsNumber = True
ElseIf parentSoros.PopType = "N" Then
r = parentSoros.PopVal
IsNumber = True
ElseIf parentSoros.PopType = "L" Then
r = parentSoros.PopValLong
IsNumber = True
End If
End Function
Public Function IsObjectRef(r As Object) As Boolean
If parentSoros.LookNull Then
Set r = Nothing
IsObjectRef = True
ElseIf parentSoros.PopType = "*" Then
Set r = parentSoros.PopObj
IsObjectRef = True
End If
End Function
Public Function IsString(r As String) As Boolean
If parentSoros.LookNull Then
r = parentSoros.PopStr
IsString = True
ElseIf parentSoros.PopType = "S" Then
r = parentSoros.PopStr
IsString = True
End If
End Function
Public Function IsInStackObject(r As Object) As Boolean  ' not used yet
If soros.LookNull Then
Set r = Nothing
IsInStackObject = True
ElseIf soros.PopType = "*" Then
Set r = parentSoros.PopObj
IsInStackObject = True
End If
End Function
Public Function IsInRetStackNumber(r As Double) As Boolean
If RetSoros Is Nothing Then Exit Function
If RetSoros.LookNull Then
r = RetSoros.PopVal
IsInRetStackNumber = True
ElseIf RetSoros.PopType = "N" Then
r = RetSoros.PopVal
IsInRetStackNumber = True
End If
End Function
Public Function IsInRetStackString(r As String) As Boolean
If RetSoros Is Nothing Then Exit Function
If RetSoros.LookNull Then
r = RetSoros.PopStr
IsInRetStackString = True
ElseIf RetSoros.PopType = "S" Then
r = RetSoros.PopStr
IsInRetStackString = True
End If

End Function
Public Function IsInStackLong(l As Long) As Boolean
If soros.PopType = "L" Then
    l = soros.PopValLong
    IsInStackLong = True
End If
End Function
Public Function IsInStackNumber(r As Double) As Boolean
If soros.LookNull Then
r = soros.PopVal
IsInStackNumber = True
ElseIf soros.PopType = "N" Then
r = soros.PopVal
IsInStackNumber = True
End If
End Function
Public Function IsInStackString(r As String) As Boolean
If soros.LookNull Then
r = soros.PopStr
IsInStackString = True
ElseIf soros.PopType = "S" Then
r = soros.PopStr
IsInStackString = True
End If
End Function
Private Sub Class_Terminate()
' debug.print "basestack terminate"
Me.ThrowThreads
If IamChild Then Set MyParent.lastobj = lastobj
Set lastobj = Nothing
Set myOwner = Nothing
Set mprocess = Nothing
Set parentSoros = Nothing
Set MyParent = Nothing
Set StaticCollection = Nothing
End Sub



Public Property Get GroupName() As String
GroupName = mGroupName
End Property

Public Property Let GroupName(ByVal vNewValue As String)

''If InStr(vNewValue, "..") > 0 Then Stop
If vNewValue <= "." Then
mGroupName = ""
Else
''''If Right$(vNewValue, 1) <> "." Then Stop
mGroupName = vNewValue
End If
End Property

Public Property Get IamAnEvent() As Variant
IamAnEvent = mIamAnEvent
End Property

Public Property Let IamAnEvent(ByVal vNewValue As Variant)
mIamAnEvent = vNewValue
End Property

Friend Function lastobjIndirect(var()) As Variant
If lastobj Is Nothing Then
ElseIf TypeOf lastobj Is mHandler Then
With lastobj
Select Case .indirect
    Case 0 To var2used
        On Error Resume Next
        Set lastobjIndirect = var(.indirect)
    Case Else
    
       Set lastobjIndirect = lastobj.objref
    End Select
End With
Else
Set lastobjIndirect = lastobj
End If

End Function


Public Property Get IamLambda() As Variant
IamLambda = mIamlambda
End Property

Public Property Let IamLambda(ByVal RHS As Variant)
mIamlambda = RHS
End Property

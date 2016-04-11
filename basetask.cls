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
Public addlen As Long
Public nokillvars As Boolean
Public IFCTRL As Boolean, jump As Boolean
Public Vars As Long, vname As Long
Public LastComm$
Public LastExecute As Long
Public LastOnce As Boolean
''Public LastLoop As Boolean
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
Private threadsnames As New Collection  ' δουλεύει
Public TaskMain As Boolean
Public OriginalName$
Private mGroupName$
Public UseGroupname$
Public LastClick As Long
Public Look2Parent As Boolean
Private dotcollection As Collection
Public StaticCollection As Collection
Private topdotlevel As Long
Public reflimit As Long
Public IndexSub As Long
Public lastobj As Object
Public FuncObj As Object
Public FuncValue As Variant
Public LastValue As Variant
Public LastMouse1 As Variant
Public LastMouse2 As String
Public LastState As Boolean
Private RetSoros As New mStiva  ' for basic mimic functions- for education only
Public SubLevel As Long
Public OriginalCode As Long
Public StaticInUse$
Private mIamAnEvent As Boolean
Public Sub SetVar(ByVal name$, var As Variant)
    On Error Resume Next
    StaticCollection.Add var, name$
    If Err.Number > 0 Then
    StaticCollection.Remove name$
    StaticCollection.Add var, name$
    End If
   
   '' Debug.Print "Put to ", Me.StaticInUse, " Var ", NAME$, " type ", Typename$(VAR)
End Sub
Function AlterVar(ByVal name$, p As Double, ss$, inttype As Boolean) As Boolean
Dim var As Variant
  var = StaticCollection(name$)
    StaticCollection.Remove name$
            Select Case ss$
            Case "!!"
            var = -1 - (var <> 0)
            Case "++"
            var = var + 1
            Case "--"
            var = var - 1
            Case "-!"
            var = -var
            Case "=", "g"
            var = p
            Case "+"
            var = p + var
            Case "*"
            var = p * var
            Case "-"
            var = var - p
            Case "/"
            If p = 0 Then
            
             Exit Function
            End If
            var = var / p
            Case "!"
            var = -1 - (var <> 0)
            End Select
            If inttype Then
            var = Int(var)
            Else
            var = MyRound(var, 13)
            End If
   StaticCollection.Add var, name$
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
    Dim i As Long
    On Error Resume Next
    If StaticCollection.Count = 0 Then Exit Function
    If Err Then Exit Function
    i = VarType(StaticCollection(name$))
    If Err.Number > 0 Then Exit Function
    ExistVar = True
End Function
Public Sub ReadVar(ByVal name$, var As Variant)
    Dim i As Long, vv As Variant
    On Error Resume Next
    i = VarType(StaticCollection(name$))
    If Err.Number Then var = Empty: Exit Sub
    ''Debug.Print NAME$, OriginalName, OriginalCode, HERE$, mGroupName
    If i = vbObject Then
            Set var = StaticCollection(name$)
            StaticCollection.Remove name$  '' we do a remove
            Set vv = Nothing
            StaticCollection.Add vv, name$
            ' objects are popped
    Else
            var = StaticCollection(name$)
    End If

End Sub
Public Sub DropNdot(n As Long)
''On Error Resume Next
If n = 0 Then Exit Sub
Dim i As Long
If dotcollection.Count - n < 0 Then
Set dotcollection = New Collection
Else
For i = dotcollection.Count To dotcollection.Count - n + 1 Step -1
dotcollection.Remove dotcollection.Count
Next i
End If
End Sub
Public Sub MoveNameDot(a$)
If a$ = "" Then Exit Sub
If dotcollection.Count = 0 Then
dotcollection.Add a$
Else
dotcollection.Add a$, , , dotcollection.Count
End If
End Sub
Public Function GetDotNew(a$, dotnum&) As Boolean
If Look2Parent Then GetDotNew = Parent.GetDotNew(a$, dotnum&): Exit Function

On Error Resume Next
If dotcollection.Count = 0 Then
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
On Error Resume Next
If dotcollection.Count < dotnum Then
'' wrong
If UseGroupname = "" Then

Else
           ' If InStr(A$, "(") > 0 Then
            ' If AscW(Left$(A$, 1)) > 255 Then
 'A$ = "ΑΥΤΟ." + A$
 'Else
 'A$ = "THIS." + A$
 'End If
  '          Else
            
            a$ = UseGroupname + a$
   '         End If

 End If
Else

  If dotcollection(dotnum) = "THIS" Then
    If UseGroupname <> "" Then
      a$ = UseGroupname + a$
    Else
    a$ = "THIS." + a$
    End If
  Else
    a$ = dotcollection(dotnum) + "." + a$

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
    GetDotNew = Err.Number = 0
End Function
Public Function GetDot(a$, st&) As Boolean
On Error Resume Next
Dim i&
Err.Clear
If dotcollection.Count = 0 Then a$ = Mid$(a$, 2): Exit Function
Do While Mid$(a$, st&, 1) = " "
st& = st& + 1
If st& > Len(a$) Then Exit Function
Loop

For i& = st& + 1 To Len(a$)
If Mid$(a$, i&, 1) <> "." Then i& = i& + 1: Exit For
Next i&
i& = i& - 2
' from st& to i& replace name
a$ = Left$(a$, st& - 1) + dotcollection(i& - st& + 1) + Mid$(a$, i&)

If Not Err Then GetDot = True
End Function

Public Property Let myitalic(ByVal rhs As Boolean)
mitalic = rhs
End Property

Public Property Get myitalic() As Boolean
myitalic = mitalic
End Property
Public Property Let myCharSet(ByVal rhs As Integer)
mCharSet = rhs
End Property

Public Property Get myCharSet() As Integer
myCharSet = mCharSet
End Property
Public Property Let myBold(ByVal rhs As Boolean)
mBold = rhs
End Property

Public Property Get myBold() As Boolean
myBold = mBold
End Property
Public Property Let toprinter(ByVal rhs As Boolean)
mtoprinter = rhs
End Property

Public Property Get toprinter() As Boolean
toprinter = mtoprinter
End Property
Public Property Let tolayer(ByVal rhs As Long)
mtolayer = rhs
End Property

Public Property Get tolayer() As Long
tolayer = mtolayer
End Property

Public Property Let toback(ByVal rhs As Boolean)
mtoback = rhs
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
Public Property Get RetStack() As mStiva
Set RetStack = RetSoros
End Property
Public Property Get soros() As mStiva
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
Public Property Set LinkThread(Threadid As Long, a As TaskInterface)
Dim aa As mThreadref
Set aa = threadsnames(CStr(Threadid))
Set aa.thread = a

End Property

Public Sub PushThread(Threadid As Long, threadvar As String)
Dim aa As New mThreadref

Set aa.thread = Nothing  ' no thread
aa.Id = Threadid
aa.VarName = threadvar
threadsnames.Add aa, CStr(Threadid)

End Sub
Public Function Exist(Threadid As Long, ThreadName As String) As Boolean
Dim aa As mThreadref
On Error Resume Next
Exist = threadsnames(CStr(Threadid)).VarName = ThreadName
If Err.Number > 0 Then
Exist = False

End If

End Function

Public Function ThrowOne(Threadid As Long) As Boolean
Dim aa As mThreadref
On Error Resume Next
' this is mandatory
If TaskMaster Is Nothing Then Exit Function
If TaskMaster.ThrowOne(Threadid) Then ThrowOne = True
If Err.Number > 0 Then Err.Clear: Exit Function
' maybe the task isn't on our list - because we send message from other thread!
Set aa = threadsnames(CStr(Threadid))
If Err.Number = 0 Then
threadsnames.Remove CStr(Threadid)
aa.thread.Dispose tmAllTasks
ThrowOne = True
Dim bb As basetask
Set aa.thread = Nothing
Else
Err.Clear
ThrowOne = Me.Parent.ThrowOne(Threadid)

End If
Err.Clear
'If Me.IamChild Then Me.Parent.ThrowOne Threadid
Set aa = Nothing
End Function
Public Function ThreadsStr(Optional ByVal lang As Long = 0)
'' use threadsnames
Dim aa As mThreadref, a$
If threadsnames.Count > 0 Then
For Each aa In threadsnames
If Left$(aa.VarName, 1) = "_" Then
If lang = 0 Then
a$ = a$ & "ΕΣΩΤΕΡΙΚΟ ΝΗΜΑ "
Else
a$ = a$ & "INTERNAL THREAD "
End If
Else
a$ = a$ & aa.VarName + Str$(aa.Id) & " "
End If
Next aa
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
Dim aa As mThreadref, KK
If threadsnames.Count > 0 Then
For Each aa In threadsnames
If Left$(aa.VarName, 5) <> "_multi" Then KK = KK + 1
Next aa
End If

ThreadsNumber = KK
End Function
Public Sub ThrowThreads(Optional idexclude As Long = -1)

Dim aa As mThreadref, Tid As Long
Dim thatProcess As TaskInterface
Dim thatBasetask As basetask

If threadsnames.Count > 0 Then
For Each aa In threadsnames
If idexclude <> aa.Id Then
If Not TaskMaster Is Nothing Then
If TaskMaster.ThrowOne(aa.Id) Then
ThrowOne aa.Id
End If
Else
ThrowOne aa.Id
End If
End If
Next aa
End If

End Sub

Private Sub Class_Initialize()
Set parentSoros = New mStiva ' maybe this will change
Set dotcollection = New Collection
Set StaticCollection = New Collection
Set Owner = Nothing
Set lastobj = Nothing
Set FuncObj = Nothing
End Sub
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
If RetSoros.LookNull Then
r = RetSoros.PopVal
IsInRetStackNumber = True
ElseIf RetSoros.PopType = "N" Then
r = RetSoros.PopVal
IsInRetStackNumber = True
End If
End Function
Public Function IsInRetStackString(r As String) As Boolean
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
Set dotcollection = Nothing
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

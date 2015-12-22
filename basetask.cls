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
Public vars As Long, VName As Long
Public arrs As Long, anames As Long
Public LastComm$
Public LastExecute As Long
Public LastOnce As Boolean
''Public LastLoop As Boolean
Private uSoros As New mStiva
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
Private mLineSpace As Long
Private Mchild As Boolean
Private mtolayer As Long
Private myRun As Boolean
Private mprocess As TaskInterface ' reference to current process maybe same to parrent...
Private mprocessOk As Boolean ' if i have a process class
Private threadsnames As New Collection  ' ��������
Public mMap As New Collection
Public TaskMain As Boolean
Public OriginalName$
Public mGroupName$
Public UseGroupname$
Public LastClick As Long
Public Look2Parent As Boolean
Private dotcollection As New Collection
Private topdotlevel As Long
Public reflimit As Long
Public IndexSub As Long
Public LastObj As Object
Public FuncObj As Object
Public FuncValue As Variant
Public LastValue As Variant
Public LastMouse1 As Variant
Public LastMouse2 As String
Public LastState As Boolean
Private RetSoros As New mStiva  ' for basic mimic functions- for education only
Public SubLevel As Long
Public OriginalCode As Long

Public Sub DropNdot(n As Long)
''On Error Resume Next
If n = 0 Then Exit Sub
Dim i As Long
If dotcollection.Count - n < 0 Then
Set dotcollection = New Collection
Else
For i = dotcollection.Count To dotcollection.Count - n + 1 Step -1
dotcollection.remove dotcollection.Count
Next i
End If
End Sub
Public Sub MoveNameDot(A$)
If A$ = "" Then Exit Sub
If dotcollection.Count = 0 Then
dotcollection.Add A$
Else
dotcollection.Add A$, , , dotcollection.Count
End If
End Sub
Public Function GetDotNew(A$, dotnum&) As Boolean
If Look2Parent Then GetDotNew = Parent.GetDotNew(A$, dotnum&): Exit Function

On Error Resume Next
If dotcollection.Count = 0 Then
If UseGroupname <> "" Then
 If InStr(A$, "(") > 0 Then
 If AscW(Left$(A$, 1)) > 255 Then
 A$ = "����." + A$
 Else
 A$ = "THIS." + A$
 End If
 Else
 A$ = UseGroupname + A$
 End If

 
 End If
Else
newrun:
On Error Resume Next
If dotcollection.Count < dotnum Then
'' wrong
If UseGroupname = "" Then

Else
            If InStr(A$, "(") > 0 Then
             If AscW(Left$(A$, 1)) > 255 Then
 A$ = "����." + A$
 Else
 A$ = "THIS." + A$
 End If
            Else
            
            A$ = UseGroupname + A$
            End If

 End If
Else
    A$ = dotcollection(dotnum) + "." + A$
    If Asc(A$) = 46 Then
    dotnum = 1
    While Asc(Mid$(A$, dotnum + 1) + " ") = 46
    dotnum = dotnum + 1
    Wend
    A$ = Mid$(A$, dotnum + 1)
    GoTo newrun
    End If
    End If
    End If
    GetDotNew = Err.Number = 0
End Function
Public Function GetDot(A$, st&) As Boolean
On Error Resume Next
Dim i&
Err.clear
If dotcollection.Count = 0 Then A$ = Mid$(A$, 2): Exit Function
Do While Mid$(A$, st&, 1) = " "
st& = st& + 1
If st& > Len(A$) Then Exit Function
Loop

For i& = st& + 1 To Len(A$)
If Mid$(A$, i&, 1) <> "." Then i& = i& + 1: Exit For
Next i&
i& = i& - 2
' from st& to i& replace name
A$ = Left$(A$, st& - 1) + dotcollection(i& - st& + 1) + Mid$(A$, i&)

If Not Err Then GetDot = True
End Function

Public Property Get LineSpace() As Long  ' in twips
LineSpace = mLineSpace
End Property

Public Property Let LineSpace(ByVal rhs As Long)
mLineSpace = rhs
End Property

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

Public Property Get Map(Id As String) As Object  'variant class
Set Map = mMap(Id)  ' get a reference write or read....and then throw the reference...
End Property


Public Function MapDelone(Id As String) As Boolean
MapDelone = False
On Error GoTo mmmm

 mMap.remove Id
 MapDelone = True
mmmm:
End Function
Public Sub MapFlush()
Dim i As Long
If mMap.Count > 0 Then
For i = 1 To mMap.Count
mMap.remove 1
Next i
End If
End Sub

Public Function MapAdd(anObj As Object) As String
Static counter As Long 'so we can "delete" a pointer...without reindexing
Dim t$
counter = counter + 1
t$ = CStr(counter)
mMap.Add anObj, t$
MapAdd = t$
End Function
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
' �� ���� ���� �� ��� ������� ��� Parent
' ���� ��� ��� Sorosref ��� ��� Soros ��� Parent (��� ���� ��� � ��� ��� ������ ���)
End Property
Public Property Get Parent() As basetask
Set Parent = MyParent

End Property
Public Property Set Sorosref(A As mStiva)
Set parentSoros = A
End Property
Public Property Get RetStack() As mStiva
Set RetStack = RetSoros
End Property
Public Property Get Soros() As mStiva
Set Soros = parentSoros
End Property
Public Property Set Owner(A As Object)
Set myOwner = A
End Property
Public Property Get Owner() As Object
If mtoprinter Then
Set Owner = Form1.PrinterDocument1
Else
Set Owner = myOwner
End If

On Error Resume Next
Err.clear
If Not Owner Is Nothing Then
mitalic = Owner.Font.Italic
mBold = Owner.Font.bold
mCharSet = Owner.Font.charset
myRun = True
Else
myRun = False
End If
Err.clear
End Property
Public Property Get IamThread() As Boolean
IamThread = mprocessOk
End Property
Public Sub EndRun()
' ���� ��� �� Myprocess.tick
myRun = False
End Sub
Public Property Set Process(A As TaskInterface)
If mprocess Is A Then
Set mprocess = Nothing
mprocessOk = True
Else
Set mprocess = A

mprocessOk = True
End If
End Property
Public Property Get Process() As TaskInterface
Set Process = mprocess

End Property
Public Property Set LinkThread(Threadid As Long, A As TaskInterface)
Dim aa As mThreadref
Set aa = threadsnames(CStr(Threadid))
Set aa.thread = A

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
If TaskMaster.ThrowOne(Threadid) Then ThrowOne = True
If Err.Number > 0 Then Err.clear: Exit Function
' maybe the task isn't on our list - because we send message from other thread!
Set aa = threadsnames(CStr(Threadid))
If Err.Number = 0 Then
threadsnames.remove CStr(Threadid)
aa.thread.Dispose tmAllTasks
ThrowOne = True
Dim bb As basetask
Set aa.thread = Nothing
Else
Err.clear
ThrowOne = Me.Parent.ThrowOne(Threadid)

End If
Err.clear
'If Me.IamChild Then Me.Parent.ThrowOne Threadid
Set aa = Nothing
End Function
Public Function ThreadsStr(Optional ByVal Lang As Long = 0)
'' use threadsnames
Dim aa As mThreadref, A$
If threadsnames.Count > 0 Then
For Each aa In threadsnames
If Left$(aa.VarName, 1) = "_" Then
If Lang = 0 Then
A$ = A$ & "��������� ���� "
Else
A$ = A$ & "INTERNAL THREAD "
End If
Else
A$ = A$ & aa.VarName + Str$(aa.Id) & " "
End If
Next aa
End If
If Me.IamChild Then
If Not Me.Parent Is Nothing Then
ThreadsStr = Trim$(A$ & Me.Parent.ThreadsStr(Lang))
Else
ThreadsStr = ""
End If
Else
ThreadsStr = Trim$(A$)
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
Set parentSoros = uSoros ' maybe this will change
End Sub
Public Function IsNumber(r As Double) As Boolean
If parentSoros.LookNull Then
r = parentSoros.PopVal
IsNumber = True
ElseIf parentSoros.PopType = "N" Then
r = parentSoros.PopVal
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
If Soros.LookNull Then
Set r = Nothing
IsInStackObject = True
ElseIf Soros.PopType = "*" Then
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
Public Function IsInStackNumber(r As Double) As Boolean
If Soros.LookNull Then
r = Soros.PopVal
IsInStackNumber = True
ElseIf Soros.PopType = "N" Then
r = Soros.PopVal
IsInStackNumber = True
End If
End Function
Public Function IsInStackString(r As String) As Boolean
If Soros.LookNull Then
r = Soros.PopStr
IsInStackString = True
ElseIf Soros.PopType = "S" Then
r = Soros.PopStr
IsInStackString = True
End If
End Function
Private Sub Class_Terminate()
' debug.print "basestack terminate"
Me.ThrowThreads
If IamChild Then Set MyParent.LastObj = LastObj
Set LastObj = Nothing
Set myOwner = Nothing
Set mprocess = Nothing
Set parentSoros = Nothing
Set MyParent = Nothing

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
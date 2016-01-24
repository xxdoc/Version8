VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim VarRef As String
Dim ArrRef As String
Dim FuncRef As String
Dim LocalRef As String  ' used only for classes
Dim OwnSoros As mStiva
Dim mName As String
Dim mFloatGroupName As String
Dim FloatGroup As Boolean
Dim bank As mArray
Dim mTeam As Boolean
Dim LinkToGroup As Object
Dim mTag As String
Public Function PrepareSoros(ByRef V()) As mStiva
Dim i As Long, AllSoros As New mStiva, s As String, VV As mStiva
For i = 1 To OwnSoros.Total
s = OwnSoros.StackItem(i)

AllSoros.DataStr s
If Left$(s$, 1) = "*" Then
 Set VV = V(Val(Split(s)(1))).PrepareSoros(V())
AllSoros.MergeBottom VV
End If
Next
Set PrepareSoros = AllSoros

End Function

Public Function types() As String
Dim i As Long, s$
i = OwnSoros.Total
On Error Resume Next
While i > 0
If InStr(OwnSoros.StackItem(i), ") ") > 0 Then
s$ = "A" + s$
Else
s$ = "V" + s$
End If
i = i - 1

Wend
types = s$
End Function
Public Sub PeekItem(cursor As Long, ITEM)
'' used in the float mode
Dim JJ As Variant
If IsObject(bank.ITEM(cursor)) Then
Set ITEM = bank.ITEM(cursor)
Else
If mTeam Then
        If cursor = bank.ITEM(1) * 2 + 3 Then
            ' we need common definitions for functions and modules all elements
                
                
                LinkToGroup.PeekItem 1, JJ
                LinkToGroup.PeekItem JJ * 2 + 3, ITEM
                        Exit Sub
            ElseIf cursor = bank.ITEM(1) * 2 + 2 Then
                    LinkToGroup.PeekItem 1, JJ
                    LinkToGroup.PeekItem JJ * 2 + 2, ITEM
                          Exit Sub
            End If
End If

ITEM = bank.ITEM(cursor)
End If

End Sub
Public Sub PokeItem(cursor As Long, ITEM)
' used in the float mode
If IsObject(ITEM) Then
Set bank.ITEM(cursor) = ITEM
Else
' always last two items are for local vars and functions
If mTeam And cursor >= bank.ITEM(1) * 2 + 2 Then
''Stop
'' do nothing
Else
 bank.ITEM(cursor) = ITEM
 End If
End If
End Sub

Public Sub BeginFloat(Total As Long)
FloatGroup = True
Set bank = New mArray
flitems = Total
bank.PushDim Total * 2
bank.PushEnd
End Sub
Public Sub BeginElement(Total As Long)
FloatGroup = False
Set bank = New mArray
flitems = Total
bank.PushDim Total
bank.PushEnd
End Sub
Public Property Get IamFloatGroup()
IamFloatGroup = FloatGroup
End Property
Public Property Get IamTeam()
IamTeam = mTeam
End Property
Public Sub EndFloat()
If FloatGroup Then bank.SerialItem (0), 0, 7
FloatGroup = False
Set bank = Nothing
End Sub
Public Property Set LinkRef(a As Object)
If FloatGroup Then
mTeam = False
Set LinkToGroup = a
If a Is Nothing Then Exit Property
'erase functions,,,
 bank.ITEM(bank.ITEM(1) * 2 + 3) = ""
mTeam = True
Else
' only when is in an array as Float Group
End If
End Property
Public Property Get Link() As Object
Set Link = LinkToGroup
End Property
Public Property Set Sorosref(a As mStiva)
If FloatGroup Then
' we can't do anything with stored values
Else
Set OwnSoros = a
End If
End Property
Public Property Get soros() As mStiva
Set soros = OwnSoros
End Property

Public Function VarlistTotal(V(), skip As Boolean) As String
Dim i As Long, s As String, Sm As String

For i = 1 To OwnSoros.Total
s = OwnSoros.StackItem(i)
If Left$(s, 1) = "*" Then
Sm = Sm + V(Val(Split(s)(1))).VarlistTotal(V(), 0)
End If
Next
If skip Then
VarlistTotal = VarRef + Sm$
Else
VarlistTotal = Replace(VarRef + Sm$, Chr$(1) + "_ ", Chr$(1) + "_ " + mName)
End If
End Function
Public Property Let Varlist(ByVal vNewValue As String)
VarRef$ = vNewValue
End Property
Public Property Get Varlist() As String
Varlist = VarRef$
End Property
Public Function FunclistTotal(V(), skip As Boolean) As String
Dim i As Long, s As String, Sm As String
For i = 1 To OwnSoros.Total
s = OwnSoros.StackItem(i)
If Left$(s, 1) = "*" Then
Sm = Sm + V(Val(Split(s)(1))).FunclistTotal(V(), False)
End If
Next
If skip Then
FunclistTotal = FuncRef + Sm
Else
FunclistTotal = Replace(FuncRef + Sm$, Chr$(2), Chr$(2) + mName$)
End If
End Function
Public Property Get FuncList() As String

FuncList = FuncRef$
End Property
Public Property Let FuncList(ByVal vNewValue As String)
FuncRef$ = vNewValue
End Property
Public Function ArrlistTotal(V(), skip As Boolean) As String
    Dim i As Long, s As String, Sm As String
    For i = 1 To OwnSoros.Total
    s = OwnSoros.StackItem(i)
    If Left$(s, 1) = "*" Then
    Sm = Sm + V(Val(Split(s)(1))).ArrlistTotal(V(), 0)
    End If
    Next
    If skip Then
    ArrlistTotal = ArrRef + Sm$
    Else
    ArrlistTotal = Replace(ArrRef + Sm$, Chr$(13) + "__", Chr$(13) + "__" + mName)
    End If
End Function
Public Property Get ArrList() As String
ArrList = ArrRef$
End Property

Public Property Let ArrList(ByVal vNewValue As String)
ArrRef$ = vNewValue
End Property
Public Property Get LocalList() As String
LocalList = LocalRef$
End Property

Public Property Let LocalList(ByVal vNewValue As String)
LocalRef$ = vNewValue
End Property
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
Value = 0
End Property
Public Property Let Value(DUMMY As Variant)
'
End Property

Public Property Get FloatGroupName() As Variant
FloatGroupName = mFloatGroupName
End Property

Public Property Let FloatGroupName(ByVal vNewValue As Variant)
mFloatGroupName = vNewValue
End Property

Public Property Get GroupName() As Variant
GroupName = mName
End Property

Public Property Let GroupName(ByVal vNewValue As Variant)
mName = vNewValue
End Property


Private Sub Class_Initialize()
Set OwnSoros = New mStiva
End Sub

Private Sub Class_Terminate()
If IamFloatGroup Then EndFloat
Set LinkToGroup = Nothing
End Sub

Public Property Get edittag() As String
If mTeam Then
edittag = LinkToGroup.edittag
Else
edittag = mTag
End If
End Property

Public Property Let edittag(ByVal vNewValue As String)
If Not mTeam Then mTag = vNewValue
End Property

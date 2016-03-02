VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'
Dim mEnabled As Boolean
Private PriveSpace() As GenItem
Dim MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Integer)
Private mParamBlock As String, mParams As Long
Private Hash() As Long
Public Property Get Params() As Long
Params = mParams
End Property
Public Property Get ParamsRead() As String
ParamsRead = mParamBlock
End Property
Public Sub ParamBlock(rhs As String, many As Long)
mParamBlock = rhs
mParams = many
End Sub
Public Property Get Enabled() As Boolean
Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal rhs As Boolean)
mEnabled = rhs
End Property
Private Sub Class_Initialize()
    MaxSpace = 10
    ReDim PriveSpace(MaxSpace) As GenItem, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
    mEnabled = True
End Sub
Public Sub BypassInit(mMaxSpace As Long)
    MaxSpace = mMaxSpace
    ReDim PriveSpace(MaxSpace) As GenItem, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
End Sub
Friend Sub CopySpaceUp(ByRef anySpave() As GenItem, ByRef anyhash() As Long, anyTopLim As Long)
anySpave() = PriveSpace()
anyhash() = Hash()
anyTopLim = toplim
End Sub
Friend Sub CopySpaceDown(ByRef anySpave() As GenItem, ByRef anyhash() As Long, anyTopLim As Long)
PriveSpace() = anySpave()
Hash() = anyhash()
toplim = anyTopLim
End Sub

Friend Sub ReadVar(where, name$, func As String)
    name$ = PriveSpace(where).Key
    func = PriveSpace(where).NDX
End Sub
Friend Sub ReadVarEx(where, name$, func As String, Extra As Long)
    name$ = PriveSpace(where).Key
    func = PriveSpace(where).NDX
    Extra = PriveSpace(where).firsthash
End Sub

Private Sub ExpandHash()

    Dim i As Long
    hashlen = hashlen * 2 + 3
    
    ReDim Hash(hashlen) As Long
    For i = 0 To MaxSpace \ 2 - 1

    place HashFunc2(i), i
    Next i
End Sub
Friend Sub ReduceHash(newTop As Long)
    Dim i As Long
    
    For i = toplim To newTop Step -1
       With PriveSpace(i)
        If .Pleft > 0 Then
            Hash(.lastpos) = .Pleft
            .Key = ""
            .Pleft = 0
        Else
            Hash(.lastpos) = 0
            .Key = ""
        End If

        .NDX = ""
            
        End With

    Next i
    If newTop <= 50 Then
            If MaxSpace > 50 Then
                        MaxSpace = 50
                         ReDim Preserve PriveSpace(MaxSpace) As GenItem
                            
                            ReDim Hash(MaxSpace * 2 + 3)
                           hashlen = MaxSpace * 2 + 3
                         '  reHash...
                         toplim = newTop - 1
                        For i = 0 To toplim
                            place HashFunc2(i), i
                        Next i
                Else
                toplim = newTop - 1
            End If
            
    Else
            If MaxSpace - newTop + 1 > 2 * newTop + 2 Then
            
                MaxSpace = 2 * (newTop + 1) + 1
                
                ReDim Preserve PriveSpace(MaxSpace) As GenItem
                
                ReDim Hash(MaxSpace * 2 + 3)
               hashlen = MaxSpace * 2 + 3
             '  reHash...
             toplim = newTop - 1
            For i = 0 To toplim
                place HashFunc2(i), i
            Next i
        
            Else
 toplim = newTop - 1
    End If
   End If
End Sub
Private Function Malloc() As Long

    If toplim + 1 >= MaxSpace Then
           '' expand hash
           MaxSpace = MaxSpace * 2
           ReDim Preserve PriveSpace(MaxSpace) As GenItem
           If MaxSpace > hashlen Then ExpandHash
    End If
        toplim = toplim + 1
        Malloc = toplim
End Function
Property Get CurMaxSpace()
CurMaxSpace = MaxSpace
End Property
Property Get Count()
Count = toplim + 1
End Property
Friend Function find(Key As String, func As String) As Boolean
Dim k As Long
func = ""
'If Right$(Key, 2) = "()" Then Stop
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    func = PriveSpace(k).NDX: find = True: Exit Function
    End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Function StandBy(Key As String, NewKey As String, func As String)
Dim k As Long

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    PriveSpace(k).NDX = func
    PriveSpace(k).Key = NewKey
    PriveSpace(k).firsthash = HashFunc0(NewKey)
                ReDim Hash(MaxSpace * 2 + 3)
               hashlen = MaxSpace * 2 + 3
             '  reHash...
            For k = 0 To toplim
                place HashFunc2(k), k
            Next k
            StandBy = True
            Exit Function
               
 End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0

 End If
            
End Function
Friend Function ExistKey(Key As String) As Boolean
Dim k As Long
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then ExistKey = True: Exit Function
  '  If k = PriveSpace(k).Pleft - 1 Then Stop: Exit Do
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function


Friend Sub GenItemCreator(Key As String, func As String)
Dim a As Long
a = Malloc()
With PriveSpace(a)
    .Key = Key
    .NDX = func
    .firsthash = HashFunc0(Key)
End With
place HashFunc(Key), a
End Sub
Property Get Percent()
Percent = Count / hashlen
End Property
Private Sub place(ByVal b, ByVal a)
Dim k As Long

    k = Hash(b)
    Hash(b) = a + 1
    PriveSpace(a).Pleft = k
    PriveSpace(a).lastpos = b
End Sub
Private Sub ab(ByVal a As Long, b As Long)
b = (b + a + 1024 * a) Xor &H83F3CF
End Sub
Public Function HashFunc0(a$) As Long
Dim i&, p2 As Long, P1 As Integer, p4 As Long, k As Long
  p4 = Len(a$) * 2 - 2
  p2 = StrPtr(a$): p4 = p2 + p4
   For i = p4 To p4 - 6 Step -2
  GetMem2 i, P1
  HashFunc0 = (HashFunc0 + P1 + 1024& * P1) Xor &H83F3CF
  Next i
  k = 2
  For i = p4 - 8 To p2 Step -2
  GetMem2 i, P1
  HashFunc0 = (HashFunc0 + P1 + 1024& * P1) Xor &H83F3CF
 i = i - 2 * k
  k = k + 1
  Next i
 
End Function
Private Function HashFunc2(where As Long)
HashFunc2 = MyMod(uintnew(PriveSpace(where).firsthash Or &HAFFA000 + Len(PriveSpace(where).Key) + hashlen + 1), hashlen)
End Function

Private Function HashFunc(a$)
HashFunc = MyMod(uintnew(HashFunc0(a$) Or &HAFFA000 + Len(a$) + hashlen + 1), hashlen)
End Function
Private Function MyMod(r1 As Double, po) As Double
MyMod = Sgn(r1) * (Int(Abs(r1)) - Int(Int(Abs(r1) / Abs(Int(po))) * Abs(Int(po))))
End Function
Private Function uintnew(ByVal a As Double) As Double
If a > CDbl(2147483647) Then a = CDbl(2147483647)
If a < CDbl(-2147483648#) Then a = CDbl(-2147483648#)
If a < 0 Then
uintnew = CDbl(&H7FFFFFFF) + (CDbl(&H7FFFFFFF) + a) + 2
Else
uintnew = a
End If
End Function
Public Sub Upgrade(usethis$)
Dim i As Long, a$
For i = 0 To toplim
a$ = PriveSpace(i).NDX
If a$ <> "" Then
a$ = Left$(a$, rinstr(a$, "}", 0))
If a$ <> "" Then
PriveSpace(i).NDX = a$ + usethis$
End If
End If
Next i
End Sub

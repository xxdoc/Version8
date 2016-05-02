VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' George's FastCollection
' Only Add, and Find.
' Key as string may be a hex number (used for enumarates)
' No same key.
' We can remove key
' We can add key (only unique keys)
' return variant (can be anything)
'**************
'* Propertiea:
'*      Count
'*'* from 1 to Count we can access all the members of FastCollection
'*     .Index=1  used for input/output
'* We can get Key or Value
'*      .KeyToString
'*      .Value
'* We can change Valus
'*    set .ValueObj = objectX
'*     .ValueVar
'**************************
'* Methods
'* AddKey "10" [, variant]
'* ToKey "10"
'* SwapValues "10", "30"
'*
'* .Done ..operation is ok
'* .Remove 10
'* .AddKey without parameter add 1 to lastkey
'* .LastKey
'* .LastIndex
'* .Frwd  (for index, exit if not .done)
'* .Back  (for index, exit if not .done)
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type item
    Key As String
    iValue As Variant
    firsthash As Long
    lastpos As Long
    Pleft As Long  ' a list
End Type
Private PriveSpace() As item
Dim MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private lastfind As String
Private lastkey As String
Private Hash() As Long
Public Done As Boolean
Public Index As Long
Private ParentIndex As Long
Private Sub Class_Initialize()
    MaxSpace = 5
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
End Sub
Public Sub AddKey(RHS, Optional aValue As Variant)
    Index = -1
    lastkey = CStr(RHS)
    Done = False
    If Not IsMissing(aValue) Then
    
   
        If IsObject(aValue) Then
        Value = -1
        Set ValueObj = aValue
        Else
        Value = aValue
        End If
    Else
        Value = aValue
    End If
End Sub
Public Sub ToKey(RHS)
    Done = Find(RHS)
    If Done Then lastkey = CStr(RHS)
End Sub
Public Sub Remove(RHS)
Dim k As Long, k1 As Long, vvv As Variant, top1 As Long
top1 = toplim
Done = Find(RHS)
If Done Then
    If Index = toplim Then

        ReduceHash top1

    Else
        If ParentIndex > 0 Then
            PriveSpace(ParentIndex).Pleft = PriveSpace(Index).Pleft
        Else
            Hash(PriveSpace(Index).lastpos) = PriveSpace(Index).Pleft
        End If
        SwapVariant vvv, PriveSpace(top1).iValue
        PriveSpace(Index).iValue = -1
        PriveSpace(Index) = PriveSpace(top1)
        Dim new_item As item
        PriveSpace(top1) = new_item
        SwapVariant vvv, PriveSpace(Index).iValue
        If Hash(PriveSpace(Index).lastpos) - 1 = top1 Then
            Hash(PriveSpace(Index).lastpos) = Index + 1
        Else
            k = Hash(PriveSpace(Index).lastpos) - 1
            If k < 1 Then
                Hash(PriveSpace(Index).lastpos) = Index + 1
            Else
                Do
                k1 = k
                k = PriveSpace(k).Pleft - 1
                Loop Until k = top1 Or k < 0
                PriveSpace(k1).Pleft = Index + 1
            End If
        End If
        toplim = toplim - 1
       End If
End If
End Sub
Property Let Value(RHS As Variant)
Done = False
If Index = -1 Then
ItemCreator lastkey, RHS

Else
'PriveSpace(Index).iValue = -1
    PriveSpace(Index).iValue = RHS
End If
Done = True
End Property
Property Set ValueObj(RHS As Variant)
Done = False
If Index = -1 Then
    ItemCreator lastkey, RHS
Else
    Set PriveSpace(Index).iValue = RHS
End If
Done = True
End Property

Property Get Value() As Variant
Done = False
If Index = -1 Then
ElseIf Not IsObject(PriveSpace(Index).iValue) Then
    Value = PriveSpace(Index).iValue
End If
End Property
Property Get ValueObj() As Variant
Done = False
If Index = -1 Then
ElseIf IsObject(PriveSpace(Index).iValue) Then
    Set ValueObj = PriveSpace(Index).iValue
    Done = True
End If

End Property
Private Sub ExpandHash()
    hashlen = hashlen * 2 + 3
    Debug.Print "ReHash"
    ReHash
End Sub
Private Sub ReHash()
    Dim i As Long
    ReDim Hash(hashlen) As Long
    For i = 0 To toplim
        place HashFunc2(i), i
    Next i
End Sub

Private Sub ReduceHash(newTop As Long)
    If newTop <= 40 Then
            If MaxSpace > 40 Then
                MaxSpace = 40
                ReDim Preserve PriveSpace(MaxSpace) As item
                
                ReDim Hash(MaxSpace * 2 + 3)
                hashlen = MaxSpace * 2 + 3
                
                ReHash
            End If
    Else
            If MaxSpace - newTop + 1 > 2 * newTop + 2 Then
            MaxSpace = 2 * (newTop + 1) + 1
            ReDim Preserve PriveSpace(MaxSpace) As item
            ReDim Hash(MaxSpace * 2 + 3)
            hashlen = MaxSpace * 2 + 3
            ReHash
    End If
   End If
End Sub
Function ChangeKey(key1, key2)

    Dim Key As String, NewKey As String
    Dim k As Long
    If key1 = key2 Then ChangeKey = True
    Key = CStr(key1)
    If Len(Key) = 0 Then Exit Function
    If ExistKey(key2) Then Exit Function
    NewKey = CStr(key2)
    If Len(NewKey) = 0 Then Exit Function
    
    k = Hash(HashFunc(Key)) - 1
    If k >= 0 Then
    Do
            If PriveSpace(k).Key = Key Then
                PriveSpace(k).Key = NewKey
                PriveSpace(k).firsthash = HashFunc0(NewKey)
                ReDim Hash(MaxSpace * 2 + 3)
                hashlen = MaxSpace * 2 + 3
                ReHash
                ChangeKey = True
                Exit Function
            End If
            k = PriveSpace(k).Pleft - 1
    Loop Until k < 0
    End If
End Function

Sub SwapValues(key1, key2)
Dim k As Long
If Find(key1) Then
k = Index
If Find(key2) Then
    SwapVariant PriveSpace(k).iValue, PriveSpace(Index).iValue
End If
End If
End Sub
Sub SwapNext(key1)
Dim k As Long
On Error Resume Next
If Done Then
    k = Index
    If Find(key1) Then
        SwapVariant PriveSpace(k).iValue, PriveSpace(Index).iValue
    End If
End If
End Sub
Sub SwapByIndex(RHS As Long)
On Error Resume Next
If Done Then
    SwapVariant PriveSpace(RHS).iValue, PriveSpace(Index).iValue
    Index = RHS
End If
End Sub


Private Function Malloc() As Long
    If toplim + 1 >= MaxSpace Then
        MaxSpace = MaxSpace * 2
        ReDim Preserve PriveSpace(MaxSpace) As item
        If MaxSpace > hashlen Then ExpandHash
    End If
    toplim = toplim + 1
    Malloc = toplim
End Function
Property Get Count()
    Count = toplim + 1
End Property
Property Get LastKnown() As String
    LastKnown = lastfind
End Property
Function IsObj() As Boolean
If Index = -1 Then
Else
IsObj = IsObject(PriveSpace(Index).iValue)
End If
End Function
Friend Function Find(RHS) As Boolean
Dim k As Long, Key As String
Key = CStr(RHS)
ParentIndex = 0
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    Find = True: lastfind = Key: Index = k:  Exit Function
    End If
     ParentIndex = k
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Function ExistKey(RHS) As Boolean
Dim k As Long, Key As String
Key = CStr(RHS)
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then ExistKey = True: Index = k: Exit Function
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Private Function ExistKey2(Key As String, nkey As Long) As Boolean
Dim k As Long
If Len(Key) = 0 Then ExistKey2 = True: Exit Function
k = Hash(HashFunc1(nkey, Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then ExistKey2 = True: Exit Function
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function

Private Sub ItemCreator(Key As String, nValue As Variant)
Dim A As Long, kk As Long
Done = False
kk = HashFunc0(Key)
If ExistKey2(Key, kk) Then Exit Sub  ' done is false
A = Malloc()

With PriveSpace(A)
    .Key = Key
    If IsObject(nValue) Then
    Set .iValue = nValue
    Else
    .iValue = nValue
    End If
    .firsthash = kk
End With
place HashFunc1(kk, Key), A
Index = A
Done = True
End Sub

Property Get Percent()
Percent = Count / hashlen
End Property
Private Sub place(ByVal b, ByVal A)
    Dim k As Long
    k = Hash(b)
    Hash(b) = A + 1
    PriveSpace(A).Pleft = k
    PriveSpace(A).lastpos = b
End Sub
Private Function HashFunc0(A$) As Long
 Dim i&, p2 As Long, P1 As Integer, p4 As Long, k As Long, p3 As Long
  p4 = Len(A$) * 2 '- 2
  p2 = StrPtr(A$): p4 = p2 + p4
   For i = p4 To p4 - 6 Step -2
  GetMem2 i, P1
  p3 = UINT(P1)
  HashFunc0 = (HashFunc0 + p3 + 1024& * p3) Xor &H83F3CF
  Next i
  k = 2
  For i = p4 - 8 To p2 Step -2
  GetMem2 i, P1
  p3 = UINT(P1)
  HashFunc0 = (HashFunc0 + p3 + 1024& * p3) Xor &H83F3CF
 i = i - 2 * k
  k = k + 1
  Next i
End Function
Private Function HashFunc2(where As Long)
HashFunc2 = MyMod(uintnew(PriveSpace(where).firsthash Or &HAFFA000) + Len(PriveSpace(where).Key) + hashlen + 1, hashlen)
End Function
Private Function HashFunc1(nkey As Long, A$)
HashFunc1 = MyMod(uintnew(nkey Or &HAFFA000) + Len(A$) + hashlen + 1, hashlen)
End Function
Private Function HashFunc(A$)
HashFunc = MyMod(uintnew(HashFunc0(A$) Or &HAFFA000) + Len(A$) + hashlen + 1, hashlen)
End Function
Private Function UINT(ByVal A As Integer) As Long
 Dim b As Integer
 b = A And &HFFFF
 If b < 0 Then
 UINT = CLng(&H10000 + b)
 Else
 UINT = CLng(b)
 End If
 
 End Function

Private Function MyMod(r1 As Double, po) As Double
MyMod = Sgn(r1) * (Int(Abs(r1)) - Int(Int(Abs(r1) / Abs(Int(po))) * Abs(Int(po))))
End Function
Private Function uintnew(ByVal A As Double) As Double
If A > CDbl(2147483647) Then A = CDbl(2147483647)
If A < CDbl(-2147483648#) Then A = CDbl(-2147483648#)
If A < 0 Then
uintnew = 4294967296# + A
Else
uintnew = A
End If
End Function

Private Sub SwapVariant(ByRef A As Variant, ByRef b As Variant)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(A), 16
   CopyMemory ByVal VarPtr(A), ByVal VarPtr(b), 16
   CopyMemory ByVal VarPtr(b), t(0), 16
End Sub
' only for tests
Private Function GetRefCount(ByRef vvv) As Long
    If Not IsObject(vvv) Then Exit Function
    Dim obj As IUnknown
    Set obj = vvv
    If obj Is Nothing Then Exit Function
    CopyMemory GetRefCount, ByVal (ObjPtr(obj)) + 4, 4
    GetRefCount = GetRefCount - 3
    Set obj = Nothing
End Function

Public Property Get KeyToString2() As Variant
If Index > -1 Then
    KeyToString2 = PriveSpace(Index).Key + Str$(PriveSpace(Index).lastpos) + Str$(Int(Percent * 1000) / 100) + Str$(hashlen) + Str$(Count)
End If
End Property
Public Property Get KeyToString() As Variant
If Index > -1 Then
    KeyToString = PriveSpace(Index).Key
End If
End Property
Public Sub ToStart()
Index = 0
Done = Count > 0
End Sub
Public Sub ToEnd()
Index = Count - 1
Done = Count > 0
End Sub

Public Sub NextIndex()
Done = False
    If Index = -1 Then
 
    Else
        If Index < toplim Then Done = True: Index = Index + 1
    End If

End Sub
Public Sub PrevIndex()
Done = False
    If Index = -1 Then
    
    Else
        If Index > 0 Then Done = True: Index = Index - 1
    End If

End Sub
Private Sub Class_Terminate()

Erase PriveSpace()
End Sub

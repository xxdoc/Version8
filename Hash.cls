VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 Private Type item
 Key As String
  NDX As Long
  KeyType As Long  ' 0 as is , 1, reference/don't delete
 firsthash As Long
 lastpos As Long
 Pleft As Long  ' a list
 End Type
'
Private PriveSpace() As item
Dim maxspace As Long
Dim hashlen As Long
Dim toplim As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Integer)

Private Hash() As Long
Private Sub Class_Initialize()
    maxspace = 10
    ReDim PriveSpace(maxspace) As item, Hash(maxspace * 2 + 3)
    hashlen = maxspace * 2 + 3
    toplim = -1

End Sub
Friend Sub ReadVar(where, name$, num As Long)
    name$ = PriveSpace(where).Key
    num = PriveSpace(where).NDX
End Sub


Private Sub ExpandHash()

    Dim I As Long
    hashlen = hashlen * 2 + 3
    
    ReDim Hash(hashlen) As Long
    For I = 0 To maxspace \ 2 - 1

    place HashFunc2(I), I
    Next I
End Sub
Friend Sub ReduceHash(newTop As Long, v())
    Dim I As Long
    
    For I = toplim To newTop Step -1
       With PriveSpace(I)
        
    
        If .Pleft > 0 Then
            Hash(.lastpos) = .Pleft
            .Key = ""
            .Pleft = 0
        Else
            Hash(.lastpos) = 0
            .Key = ""
        End If
        If .KeyType = 0 And .NDX >= 0 Then
            If IsObject(v(.NDX)) Then Set v(.NDX) = Nothing
            v(.NDX) = Empty
            End If
        .KeyType = 0
        .NDX = -1
            
        End With

    Next I
    If maxspace - newTop + 1 > 2 * newTop + 2 Then
        maxspace = 2 * (newTop + 1) + 1
        ReDim Preserve PriveSpace(maxspace) As item
        
        ReDim Hash(maxspace * 2 + 3)
       hashlen = maxspace * 2 + 3
     '  reHash...
     toplim = newTop - 1
    For I = 0 To toplim
        place HashFunc2(I), I
    Next I

Else
 toplim = newTop - 1
    End If
   
End Sub

Private Function Malloc() As Long

    If toplim + 1 >= maxspace Then
           '' expand hash
           maxspace = maxspace * 2
           ReDim Preserve PriveSpace(maxspace) As item
           If maxspace > hashlen Then ExpandHash
    End If
        toplim = toplim + 1
        Malloc = toplim
End Function
Property Get Count()
Count = toplim + 1
End Property
Friend Function find(Key As String, num As Long) As Boolean
Dim k As Long
num = 0

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    ' Debug.Print "var ", Key
    num = PriveSpace(k).NDX: find = True: Exit Function
    End If
 '   If k = PriveSpace(k).Pleft - 1 Then Stop: Exit Do
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

Friend Function findRebound(Key As String, newnum As Long, v()) As Boolean
Dim k As Long
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
With PriveSpace(k)
    If .Key = Key Then
    If .KeyType = 0 Then
        If IsObject(v(.NDX)) Then Set v(.NDX) = Nothing
        v(.NDX) = Empty
        .KeyType = 1
    End If
    .NDX = newnum
    ' must be keytype=1
     findRebound = True: Exit Function
    End If
   ' If k = .Pleft - 1 Then Stop: Exit Do
    End With
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Sub ItemCreator(Key As String, num As Long, Optional isref As Boolean = False)
Dim a As Long
a = Malloc()
If Right$(Key, 2) = "()" Then Stop
With PriveSpace(a)
    .Key = Key
    .NDX = num
    If isref Then
        .KeyType = 1
    Else
        .KeyType = 0
    End If
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
Dim I&, p2 As Long, P1 As Integer, p4 As Long, k As Long
  p4 = Len(a$) * 2 - 2
  p2 = StrPtr(a$): p4 = p2 + p4
   For I = p4 To p4 - 6 Step -2
  GetMem2 I, P1
  HashFunc0 = (HashFunc0 + P1 + 1024& * P1) Xor &H83F3CF
  Next I
  k = 2
  For I = p4 - 8 To p2 Step -2
  GetMem2 I, P1
  HashFunc0 = (HashFunc0 + P1 + 1024& * P1) Xor &H83F3CF
 I = I - 2 * k
  k = k + 1
  Next I
 
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
Function FindLast()

End Function

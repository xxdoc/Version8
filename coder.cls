VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "coder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This is the Coder utility"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Function GOOD(ByRef counter As Long, ByVal key As String, ByVal item As String) As String
Dim i As Long, j As Long
j = Len(key)
If ReduceKey(item) Then
For i = 1 To Len(item)


Mid(item, i, 1) = ChrW(AscW(Mid$(item, i, 1)) Xor Asc(Mid$(key, (counter Mod j) + 1)))

counter = counter + 1
Next i
Else
For i = 1 To Len(item)

Mid$(item, i, 1) = Chr(Asc(Mid$(key, (counter Mod j) + 1, 1)) Xor (Asc(Mid$(item, i, 1))))

counter = counter + 1
Next i
End If
GOOD = item
End Function
Private Sub ExpandKey(key As String)
Dim i As Long, b As Long, K As Long, c As Long, cx1 As Long, c2 As Long, c1 As Long

Randomize
b = Len(key)
key = String$(Len(key), " ") + key
For i = 1 To Len(key) - 1 Step 2
K = Int(8 * Rnd)
c = AscW(Mid$(key, b + i \ 2 + 1, 1))
c1 = UINT(c And &HF00) + &H4000
c2 = UINT(c And &HF000) \ 8 + &H4000


c = c And &HFF
Mid$(key, i, 2) = ChrW(c1 + 128 + K + (c And &HF) * 8) + ChrW(c2 + 128 + K + (c And &HF0) / 2)
Next i

End Sub
Private Function UINT(ByVal a As Long) As Long
 Dim b As Integer
 b = a And &HFFFF
 If b < 0 Then
 UINT = CLng(&H10000 + b)
 Else
 UINT = CLng(b)
 End If
 
 End Function
Private Function ReduceKey(key As String) As Boolean
Dim i As Long, j As Long, c1 As Long, c2 As Long, noANSI As Boolean
On Error Resume Next
If (Len(key) Mod 2 = 1) Or key = "" Then Exit Function
noANSI = Not ((UINT(AscW(Mid$(key, 2, 1))) And &H4000) = 0 Or (UINT(AscW(Mid$(key, i + 1, 1))) And &H4000) = 0)
If noANSI Then
For i = 1 To Len(key) - 1 Step 2
c2 = UINT(AscW(Mid$(key, i + 1, 1)))
c1 = UINT(AscW(Mid$(key, i, 1)))
Mid$(key, i \ 2 + 1, 1) = ChrW((c1 And &HF00) + (c1 And &H78) \ 8 + (c2 And &H1E00) * 8 + (c2 And &H78) * 2)
Next i
Else
For i = 1 To Len(key) - 1 Step 2
Mid$(key, i \ 2 + 1, 1) = Chr((Asc(Mid$(key, i, 1)) And &H78) \ 8 + (Asc(Mid$(key, i + 1, 1)) And &H78) * 2)
Next i
End If

key = Mid$(key, 1, Len(key) \ 2)
ReduceKey = noANSI
End Function
Private Function ReduceKey1(key As String) As Boolean
Dim i As Long, j As Long
If (Len(key) Mod 2 = 1) Or key = "" Then Exit Function
For i = 1 To Len(key) - 1 Step 2
Mid$(key, i \ 2 + 1, 1) = Chr((Asc(Mid$(key, i, 1)) And &H78) \ 8 + (Asc(Mid$(key, i + 1, 1)) And &H78) * 2)
Next i
key = Mid$(key, 1, Len(key) \ 2)
ReduceKey1 = False
End Function
Private Function Bad(ByRef counter As Long, key As String, item As String) As String
Dim i As Long, j As Long
j = Len(key)
For i = 1 To Len(item)
Mid(item, i, 1) = ChrW(AscW(Mid$(item, i, 1)) Xor Asc(Mid$(key, (counter Mod j) + 1)))

counter = counter + 1
Next i

 ExpandKey item
Bad = item
End Function
Public Function decryptline(ByVal Data As String, ByVal MasterKey As String, ByRef Start As Long) As String
decryptline = GOOD(Start, MasterKey, Data)
End Function

Public Function encryptline(ByVal Data As String, ByVal MasterKey As String, ByRef Start As Long) As String
encryptline = Bad(Start, MasterKey, Data)
End Function
Public Function must(sss$) As String
must = decryptline(sss$, "MISTAKE TRY AGAIN", (Len(sss$) / 2) Mod 33)
End Function
Public Function must1(sss$) As String
must1 = encryptline(sss$, "MISTAKE TRY AGAIN", Len(sss$) Mod 33)
End Function





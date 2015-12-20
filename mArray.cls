VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mmname As String
Dim dnum As Long
Dim limit(10) As Long  ' max 9 dimensions
Dim dtable(10) As Long
Dim actualdata() As Variant
Dim maxitems As Long
Dim maxitems1 As Long
Public IHaveClass As Boolean
Public GroupRef As Object   ' get a floating Group, maybe with groups inside....as floating group
Public CodeName As String
Private oldmaxitems As Long
Public Sub CopyArray(k As mArray)
'copy this to k
Dim V As Variant, i As Long
k.SerialItem V, Me.UpperMonoLimit, 8
k.StartResize
For i = 0 To dnum - 1
k.PushDim limit(i)
Next i
k.PushEnd
If Right$(k.arrname, 2) = "%(" Then
For i = 0 To Me.UpperMonoLimit

If IsObject(Me.item(i)) Then
Set k.item(i) = Me.item(i)
Else
k.item(i) = Int(Me.item(i))
End If
Next i
Else
''''''''''****************************
For i = 0 To Me.UpperMonoLimit

If IsObject(Me.item(i)) Then
Set k.item(i) = Me.item(i)
Else
k.item(i) = Me.item(i)
End If
Next i
End If

End Sub

Public Function UpperMonoLimit()
UpperMonoLimit = maxitems1
End Function
Public Property Let arrname(aName As String)
mmname = aName
End Property
Public Property Get arrname() As String
arrname = mmname
End Property
Public Sub PushDim(nDim As Long)
If dnum < 10 Then
limit(dnum) = nDim
maxitems = maxitems * nDim
dnum = dnum + 1
End If
maxitems1 = maxitems - 1
End Sub
Public Sub PushEnd()
On Error GoTo there11
If dnum = 0 Then
maxitems = 1
maxitems1 = 0
ReDim actualdata(maxitems) As Variant
Else
ReDim Preserve actualdata(maxitems) As Variant
Dim i As Long, mx As Long
mx = maxitems
If limit(i) > 0 Then
For i = 0 To dnum - 1
    dtable(i) = mx / limit(i)
    mx = dtable(i)
Next i
Else
GoTo there11
End If

End If
Exit Sub
there11:
dnum = 0
maxitems = 1
maxitems1 = 0
ReDim actualdata(maxitems)
End Sub
Public Sub StartResize()
oldmaxitems = maxitems1
dnum = 0 ' no actions
maxitems = 1
maxitems1 = 0
End Sub
Public Function SerialItem(item As Variant, cursor As Long, command As Long) As Boolean
If command = 1 Then
If dnum = 0 Then Exit Function
If cursor < 0 Or cursor >= maxitems Then Exit Function
actualdata(cursor) = item
ElseIf command = 2 Then
If dnum = 0 Then Exit Function
If cursor < 0 Or cursor >= maxitems Then Exit Function
item = actualdata(cursor)
ElseIf command = 3 Then

    Fill item
ElseIf command = 4 Then
    cursor = maxitems
ElseIf command = 5 Then
cursor = dnum
item = dnum
ElseIf command = 6 Then
If cursor > 9 Or cursor < 0 Then
item = 0
Else
item = limit(cursor)
End If
ElseIf command = 7 Then  'erase all data
Dim aa As Variant
Fill aa
ElseIf command = 8 Then
If dnum = 1 Then
    ReDim Preserve actualdata(cursor) As Variant
    maxitems = cursor
    limit(0) = maxitems
End If
ElseIf command = 9 Then
If dnum = 1 Then
    ReDim Preserve actualdata(cursor) As Variant
    maxitems = cursor - 1
    limit(0) = maxitems
End If
maxitems1 = maxitems - 1
End If
SerialItem = True

End Function

Public Function PushOffset(curitem As Long, curdim As Long, nDim As Long) As Boolean
If curdim >= dnum Then
' error...
PushOffset = False
Else
If nDim >= limit(curdim) Then Exit Function
curitem = curitem + dtable(curdim) * nDim
PushOffset = True
End If
End Function
Public Property Get itemPtr(curitem As Long) As Long
itemPtr = VarPtr(actualdata(curitem))
End Property
Public Property Get item(curitem As Long) As Variant
On Error GoTo there1
If IsObject(actualdata(curitem)) Then
Set item = actualdata(curitem)
Else
item = actualdata(curitem)
End If
Exit Property
there1:

End Property
Public Property Let item(curitem As Long, item As Variant)
On Error Resume Next
actualdata(curitem) = item

End Property
Public Property Set item(curitem As Long, item As Variant)
On Error Resume Next
Set actualdata(curitem) = item

End Property

Private Sub Fill(item As Variant)
Dim aa As Object
If dnum > 0 Then
Dim i As Long
If IHaveClass Then
For i = 0 To maxitems - 1
CopyGroup GroupRef, aa
Set actualdata(i) = aa
Next i

Else
If Typename(item) = "Empty" Then
ReDim actualdata(maxitems)
''maxitems = 0

Else
For i = 0 To maxitems - 1
actualdata(i) = item
Next i
End If
End If
End If
End Sub
Sub MakeitObject2(var As Variant)
Dim aa As Object
Set aa = New Group
Set var = aa
End Sub
Sub CopyGroup(myobject As Object, myobject2 As Object)
'Special Copy
 Dim ps As New mStiva, V As Long, s$, frm$, vvl As Variant, x1 As Long, SS$
Dim subgroup As Object, sub2 As Object
 Set myobject2 = New Group
 
             With myobject
                If .IamFloatGroup Then
         
                 .PeekItem 1, vvl
                 myobject2.BeginFloat vvl + 2
                 For x1 = 0 To vvl * 2 + 2
                    .PeekItem x1, vvl
                    If Typename(vvl) = "Group" Then
                    Set subgroup = vvl
                    CopyGroup subgroup, sub2
                    Set subgroup = Nothing
                    Set vvl = sub2
                    Set sub2 = Nothing
                    End If
                    myobject2.PokeItem x1, vvl
     
                Next x1
                

                Set myobject2.LinkRef = myobject
                End If
            End With
         
End Sub
Private Sub Class_Initialize()
dnum = 0 ' no actions
maxitems = 1
maxitems1 = 0
ReDim actualdata(0)
CodeName = GetAnyName$
End Sub

Private Sub Class_Terminate()
Erase actualdata()
End Sub
Private Function GetAnyName$()
Dim SS$, j As Long
For j = 1 To 4
SS$ = SS$ & Chr(65 + Int((23 * Rnd) + 1))
Next j
GetAnyName$ = ChrW(&H1FFF) + SS$
End Function

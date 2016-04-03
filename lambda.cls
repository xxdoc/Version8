VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lambda functions in M2000
' in running state copy nonlocal variables to current list
Public name$
Public Code$
Private Type Vars
    Value As Variant
    There As Long
End Type
Private Total As Long, fixed As Long, used As Long
Private myVar() As Vars, myHash As New Hash

Sub CopyToVar(bstack As basetask, glb As Boolean, v() As Variant)
Dim i As Long, nm$, where As Long, dummy As Variant, x1 As Long
If used > 0 Then
For i = 0 To used - 1
    myHash.ReadVar i, nm$, where
    If IsObject(myVar(where).Value) Then
    dummy = 0
    myVar(where).There = GlobalVar(nm$, dummy, , glb)
    
    If Typename(myVar(where).Value) = "lambda" Then
         
                        If glb Then
                        GlobalSub nm$ + "()", "CALL EXTERN " & CStr(myVar(where).There)
                    Else
                        GlobalSub HERE$ & "." & bstack.GroupName & nm$ + "()", "CALL EXTERN " & CStr(myVar(where).There)
                End If
    
    
    End If
    Set v(myVar(where).There) = myVar(where).Value
    Else
    myVar(where).There = GlobalVar(nm$, myVar(where).Value, , glb)
    End If
Next i
End If
End Sub
Sub CopyFromVar(v() As Variant)
Dim i As Long
For i = 0 To used - 1
    If IsObject(v(myVar(i).There)) Then
        Set myVar(i).Value = v(myVar(i).There)
        Set v(myVar(i).There) = Nothing
    Else
        myVar(i).Value = v(myVar(i).There)
    End If
    myVar(i).There = 0
Next i

End Sub

Sub FeedNonLocal(aName$, val As Variant)
Dim where As Long
If Not myHash.Find(aName$, where) Then
where = GetFree()
myHash.ItemCreator aName$, where
End If
If IsObject(val) Then
'' maybe here we can copy
'' check first the type
    Set myVar(where).Value = val
Else
    myVar(where).Value = val
End If
End Sub

Private Sub Class_Initialize()
Total = 10
ReDim myVar(Total) As Vars
End Sub
Private Function GetFree() As Long
    If used >= Total Then
        Total = Total * 2
        ReDim Preserve myVar(Total) As Vars
    End If
    GetFree = used
    used = used + 1
End Function

Private Sub Class_Terminate()
Dim i As Long
If used > 0 Then
For i = used - 1 To 0
If IsObject(myVar(i).Value) Then
Set myVar(i).Value = Nothing
End If
Next i
Erase myVar()
End If
End Sub
Public Sub CopyTo(A As lambda, v() As Variant)
' only for new
Dim B As lambda, VR As Variant, where As Long
Set A = New lambda
A.Code$ = Me.Code$
Dim i As Long
If used > 0 Then
For i = 0 To used - 1
myHash.ReadVar i, nm$, where
If myVar(where).There <> 0 Then
If Typename(v(myVar(where).There)) = "lambda" Then
    With v(myVar(where).There)
    .CopyTo B, v()
    End With
    Set VR = B
    A.FeedNonLocal nm$, VR
    Set VR = Nothing
    Set B = Nothing
Else
    A.FeedNonLocal nm$, v(myVar(where).There)
End If
Else
If Typename(myVar(where).Value) = "lambda" Then
    myVar(where).Value.CopyTo B, v()
    Set VR = B
    A.FeedNonLocal nm$, VR
    Set VR = Nothing
    Set B = Nothing
Else
    A.FeedNonLocal nm$, myVar(where).Value
End If
End If
Next i
End If
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stdCallFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mLib As String
Dim mFunction As String
Dim mConctruct As String
Dim RetType2Stack As Long   ' 0 is long, 1 is string without zeros
Dim t() As Long
Dim paramnum As Long


Public Sub CallThis(ByVal vNewValue As String, paramList As String, Lang As Long)
Dim l As Long, check$, m As Long
l = InStr(vNewValue, ".")
If l <> 0 Then
mLib = Left$(vNewValue, l - 1)
mFunction = Mid$(vNewValue, l + 1)
Else
mLib = vNewValue
mFunction = "" ' what about that??
End If
mConctruct = paramList
Do While paramList <> ""
If FastSymbol(paramList, "&") Then l = 16 Else l = 0
 If IsLabelSymbolNew(paramList, "ÌÁÊÑÕÓ", "LONG", Lang) Then
 If IsPureLabel(paramList, check$) = 1 Then
 'OK
 m = 2
 Else
 paramList = "error"
 Exit Sub
 End If
 Else
 m = IsPureLabel(paramList, check$)
 End If
PushType m + l   'above 16 is byref
If IsArrayItem(paramnum) Then
If Not FastSymbol(paramList, ")") Then Exit Do
End If
If Not FastSymbol(paramList, ",") Then Exit Do
If check$ = "" Then Exit Do
If paramnum = 63 Then Exit Do
Loop
paramList = Trim$(paramList)   ' if we have somethig then that is an error
End Sub
Public Property Get LIB() As String
LIB = mLib
End Property

Public Property Get FUNC() As String
FUNC = mFunction
End Property

Public Function ReadType(there As Long) As Long
ReadType = (t(there) And &HF)
End Function
Public Function IsArrayItem(there As Long) As Boolean
IsArrayItem = (t(there) And &HF) > 4
End Function
Public Function CheckType(there As Long, ByVal thistype As Long) As Boolean
Select Case thistype
Case 1, 4, 5, 7
thistype = 1
Case Else
thistype = 2
End Select
Select Case (t(there) And &HF)
Case 1, 4, 5, 7
CheckType = thistype = 1
Case Else
CheckType = thistype = 2
End Select
End Function
Public Function IsByRef(there As Long) As Boolean
IsByRef = (t(there) And &HF0) <> 0
End Function


Private Sub PushType(ParamType As Long)
Dim hlp As Long
If paramnum + 1 >= UBound(t()) - LBound(t()) Then
hlp = UBound(t()) * 2
ReDim Preserve t(hlp)
End If
paramnum = paramnum + 1
t(paramnum) = ParamType


End Sub

Private Sub Class_Initialize()
ReDim t(16) As Long
paramnum = -1
End Sub

Public Property Get Count() As Long
Count = paramnum + 1
End Property

Public Property Get RetType() As Long
RetType = RetType2Stack
End Property

Public Property Let RetType(ByVal vNewValue As Long)
RetType2Stack = vNewValue
End Property

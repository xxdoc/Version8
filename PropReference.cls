VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim ObjectRef As Long
Dim vtref As Long
Dim mIndex As Variant
Public UseIndex As Boolean
Dim Init As Boolean
Dim IamMethod  As Boolean
Dim hardlink As Object, useobj As Boolean
Public Arr As Boolean
Sub Construct(ref As Long, vtrefnumber As Long)
If useobj Then Exit Sub
Init = True
ObjectRef = ref
vtref = vtrefnumber
End Sub
Sub ConstructObj(obj As Object, vtrefnumber As Long)
Init = True
Set hardlink = obj
useobj = True
'ObjectRef = ref
vtref = vtrefnumber
End Sub

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
If Init Then
If UseIndex Then
If useobj Then
 Value = ReadPropIndexObj(hardlink, vtref, mIndex)
Else
 Value = ReadPropIndex(ObjectRef, vtref, mIndex)
 End If

Else
If useobj Then
 Value = ReadPropObj(hardlink, vtref)
Else
 Value = ReadProp(ObjectRef, vtref)
 End If
 End If
Else
' ??????????
End If
End Property
Public Property Let Value(vNewValue As Variant)
If Init Then
If Not UseIndex Then
If useobj Then
WritePropObj hardlink, vtref, vNewValue
Else
WriteProp ObjectRef, vtref, vNewValue
End If
Else
If useobj Then
WritePropIndexObj hardlink, vtref, vNewValue, mIndex
Else
WritePropIndex ObjectRef, vtref, vNewValue, mIndex
End If
End If
End If

End Property
Public Property Get isMethod() As Boolean
isMethod = IamMethod
End Property

Public Property Let isMethod(ByVal vNewValue As Boolean)
IamMethod = vNewValue
End Property

Private Sub Class_Terminate()
Set hardlink = Nothing
End Sub
Public Property Get Index() As Variant
Index = mIndex
End Property

Public Property Let Index(ByVal rhs As Variant)
mIndex = rhs
End Property

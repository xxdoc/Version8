VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mTypeName As String
Private mItemStr As String
Private mItemDouble As Double
Private mItemObject As Object
Private IamNull As Boolean
 Property Let Typename(ByRef rhs As String)
mTypeName = rhs
End Property

 Property Get Typename() As String
Typename = mTypeName
End Property

 Property Let ItemStr(ByRef rhs As String)
 IamNull = False
mItemStr = rhs
End Property

 Property Get ItemStr() As String
ItemStr = mItemStr
End Property
 Property Let ItemDouble(ByVal rhs As Double)
 IamNull = False
mItemDouble = rhs
End Property

Property Get ItemDouble() As Double
ItemDouble = mItemDouble
End Property

Private Sub Class_Initialize()
IamNull = True
End Sub
Property Get NoValue() As Boolean
NoValue = IamNull
End Property

Public Property Set ObjectRef(rhs As Object)
IamNull = rhs Is Nothing
    Set mItemObject = rhs
End Property
Public Property Get ObjectRef() As Object
Set ObjectRef = mItemObject
End Property

Private Sub Class_Terminate()
' Ensure that if this is a last reference to object...
' then a class terminate happen
Set mItemObject = Nothing
End Sub

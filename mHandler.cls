VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Generic Handler
Public objref As Object
Public T1 As Variant, Indirect As Long
Public ReadOnly As Boolean

Private Sub Class_Initialize()
Indirect = -1
End Sub

''Public Par1 As Variant
Private Sub Class_Terminate()
Set objref = Nothing
End Sub
Public Sub CopyTo(A As mHandler)
Set A = New mHandler
A.T1 = T1
''A.Par1 = Par1   ' key or index ???
' no copy for reference
A.Indirect = Indirect
Set A.objref = objref
End Sub

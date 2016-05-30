VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mPtr As Long
Private mItems As Long
Private ItemLen As Long
Private mBytes As Long
Private ErrorMes As Long
Public structref As FastCollection
Public UseStruct As Boolean
Private mNorun As Boolean

Sub Costruct(ByteLen As Long, Optional Times As Long = 1, Optional clearme As Long = 0, Optional executable As Boolean = False)
If mItems <> 0 Then ErrorMes = 1: Exit Sub
If Times < 1 Then ErrorMes = 2:
If ByteLen < 1 Then ErrorMes = ErrorMes + 4
If ErrorMes <> 0 Then Exit Sub
ItemLen = ByteLen
mItems = Times
mNorun = Not executable
mPtr = Blockmalloc(ByteLen * Times, clearme, mNorun)
If mNorun Then If mPtr <> 0 Then mBytes = ByteLen * Times
End Sub
Property Get ItemSize()
ItemSize = ItemLen
End Property
Public Sub ClearError()
    ErrorMes = 0
End Sub
Public Property Get Status()
    Status = ErrorMes
End Property
Public Sub ResizeItems(NewSize As Long, Optional clearme As Long = 0)
If Not mNorun Then Exit Sub
If mBytes = 0 Then Exit Sub
Dim oPtr As Long
oPtr = mPtr
If NewSize > 0 Then
    oPtr = BlockResizemalloc(mPtr, ItemLen * NewSize, clearme)
    If oPtr <> 0 Then
        mItems = NewSize
        mPtr = oPtr
    Else
        ErrorMes = 3
    End If
Else
    ErrorMes = 2
End If
End Sub
Public Property Get GetBytePtr(offset As Double) As Long
If ErrorMes = 0 And mItems > 0 Then
If offset < 0 Or offset >= mItems * ItemLen Then
    ErrorMes = 4
Else
    GetBytePtr = signlong(uintnew(mPtr) + offset)
    End If
Else
    GetBytePtr = 0  ' null
End If
End Property
Public Property Get GetPtr(offset As Double) As Long
If ErrorMes = 0 And mItems > 0 Then
If offset < 0 Or offset >= mItems Then
    ' zero
    ErrorMes = 4
Else
    GetPtr = cUlng(uintnew(mPtr) + ItemLen * offset)
    End If
Else
    GetPtr = 0  ' null
End If
End Property
Public Property Get SizeByte() As Long
If ErrorMes = 0 And mItems > 0 Then
If mNorun Then
    SizeByte = BlockSize(mPtr)
    Else
    SizeByte = mBytes
    End If
End If
End Property
Public Property Get items() As Long
If ErrorMes = 0 And mItems > 0 And ItemLen > 0 Then
If mNorun Then
    items = BlockSize(mPtr) \ ItemLen
    Else
    items = mBytes \ ItemLen
    End If
    
End If
End Property
Public Function ValidArea(ByVal ThisStart As Long, ByVal ThisWidth As Double) As Boolean
Dim A As Double
' ThisStart is a sign long which hold same bits with unsign address
If ErrorMes = 0 And mItems > 0 Then
A = uintnew(CDbl(mPtr))  ' make mptr an unsign as double
If ThisStart = 0 Then
    ThisStart = A
Else
    ThisStart = uintnew(CDbl(ThisStart))
End If
' we have to make both unsign and then check them
If ThisStart - A >= 0 Then
If mNorun Then
If ThisStart + ThisWidth <= A + BlockSize(mPtr) Then ValidArea = True
Else
If ThisStart + ThisWidth <= A + mBytes Then ValidArea = True
End If
End If
End If

End Function
Public Function ValidArea2(ByVal ThisOffsetU As Double, ByVal ThisWidth As Double) As Boolean
' ThisOffsetU is unsign
If ErrorMes = 0 And mItems > 0 Then
If ThisOffsetU >= 0 Then
If mNorun Then
If ThisOffsetU * ItemLen + ThisWidth <= BlockSize(mPtr) Then ValidArea2 = True
Else
If ThisOffsetU * ItemLen + ThisWidth <= mBytes Then ValidArea2 = True
End If
End If
End If

End Function

Private Sub Class_Terminate()
Set structref = Nothing
If mPtr <> 0 Then
If NoRun Then
If BlockSize(mPtr) > 0 Then
    BlockFree mPtr
End If
Else
        BlockFreeVirtual mPtr, mBytes
        mBytes = 0
End If
End If
mPtr = 0
End Sub

Public Property Get NoRun() As Boolean
NoRun = mNorun
End Property


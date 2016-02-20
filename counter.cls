VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TaskInterface
   Private Type SECURITY_ATTRIBUTES
           nLength As Long
           lpSecurityDescriptor As Long
           bInheritHandle As Long
   End Type
   
    Const SECURITY_DESCRIPTOR_MIN_LENGTH = (20)
    Const SECURITY_DESCRIPTOR_REVISION = (1)
Dim jobnameID As Long
' load code or just use it as a counter
' load code first time
Dim m_duration As Double
Dim mYspace As Double
    Dim myholdtime As Double
Dim mycode As String
Dim MyMode As Long
Dim mybusy As Boolean
Public hPipe As Long
Private pipename As String  '= "\\.\pipe\bigtest"
Private mytarget As String
Private myBUFvar As String, res As Long
Private myBase As TaskInterface
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private myProcess As basetask

Private Property Get TaskInterface_Process() As basetask
Set TaskInterface_Process = myProcess
End Property

Private Property Set TaskInterface_Process(aProcess As basetask)
Set myProcess = aProcess
Set myProcess.Process = myBase
End Property
Private Sub Class_Terminate()
If hPipe <> 0 Then
res = DisconnectNamedPipe(hPipe)
 CloseHandle hPipe
End If
End Sub

Private Property Let TaskInterface_interval(ByVal rhs As Double)
mYspace = rhs
m_duration = uintnew(CDbl(timeGetTime)) + mYspace
End Property

Private Property Get TaskInterface_interval() As Double
'
TaskInterface_interval = mYspace
End Property

' IMPLEMENTED PROPERTIES

Private Property Set TaskInterface_Owner(rhs As Object)
' Usage: Private Property Set TaskInterface_Owner(RHS As Form)
  
  'Validating type
  'If TypeOf RHS Is Form Then
    Set myBase.Owner = rhs
  'Else
   ' Error tmTypeMisMatch
  'End If
End Property


' IMPLEMENTED METHODS

Private Sub TaskInterface_Parameters(ParamArray Values() As Variant)
'' Usage: Private Sub TaskInterface_Parameters(Color As Long, Count As Long)
On Error GoTo poulos
  ' Verifing parameter count
  If UBound(Values) = 4 Then
    
    On Error Resume Next
   jobnameID = CLng(Values(0)) ' this is the hanlde
   mYspace = CDbl(Values(1)) '
 '  m_duration = uintnew(CDbl(timeGetTime)) '+ mYspace
  pipename = "\\.\pipe\" & Values(2) + ChrW$(0)
'pipename = validpipename(CStr(Values(2)))
   myholdtime = CDbl(Values(3))

   
    m_duration = uintnew(CDbl(timeGetTime)) + myholdtime
   mytarget$ = Values(4)
   ' AND A VARIABLE....TO PLACE DATA
   Dim openMode, pipeMode As Long
    Dim SA As SECURITY_ATTRIBUTES
'Create the named pipe
   Const GMEM_FIXED = &H0
    Const GMEM_ZEROINIT = &H40
    Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
    Dim pSD As Long
  pSD = GlobalAlloc(GPTR, SECURITY_DESCRIPTOR_MIN_LENGTH)
      res = InitializeSecurityDescriptor(pSD, SECURITY_DESCRIPTOR_REVISION)
      res = SetSecurityDescriptorDacl(pSD, -1, 0, 0)
      SA.nLength = LenB(SA)
      SA.lpSecurityDescriptor = pSD
      SA.bInheritHandle = True
openMode = PIPE_ACCESS_INBOUND 'PIPE_ACCESS_DUPLEX ' Or WRITE_DAC ' PIPE_ACCESS_DUPLEX | WRITE_DAC
pipeMode = PIPE_NOWAIT Or PIPE_TYPE_BYTE Or PIPE_READMODE_BYTE
hPipe = CreateNamedPipe(StrPtr(pipename), openMode, pipeMode, 10, 10000, 2000, 400, SA)

If INVALID_HANDLE_VALUE = hPipe Then
myBase.Done = True: Exit Sub
End If
    If Err.Number > 0 Then
      On Error GoTo 0

    End If
    On Error GoTo 0
    ' do something
  Else
poulos:

  End If
  
End Sub
Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = ab() 'StrConv(ab, vbUnicode)
    
End Function
Private Sub TaskInterface_Tick()
 Dim x As Integer, mycnt As Long
Dim lpBuffer(2048) As Byte
On Error GoTo there
If m_duration <= uintnew(CDbl(timeGetTime)) Then
   m_duration = uintnew(CDbl(timeGetTime)) + mYspace
   If hPipe <> 0 Then
  ' mybuf$ = String$(2001, Chr(0))
   If ConnectNamedPipe(hPipe, ByVal 0) = 0 Then
    If GetLastError = ERROR_PIPE_LISTENING Then

    Exit Sub
    End If
   Else
   Exit Sub
   End If
    Erase lpBuffer()
  res = ReadFile(hPipe, lpBuffer(0), 2043, mycnt, ByVal 0&)


  If res = 1 Then

  If mycnt > 0 Then

  
    Thing mytarget$, Left$(BytesToStr(lpBuffer()), mycnt / 2)
    DisconnectNamedPipe hPipe
  End If
  ElseIf GetLastError = ERROR_NO_DATA Then
  Stop
  ' ERROR_NO_DATA
  
  End If
' ByVal stringbuffer, 10, numread, 0
   End If
   End If
   Exit Sub
there:
    myBase.Done = True
End Sub



' DELEGATED PROPERTIES



Private Property Let TaskInterface_Done(ByVal rhs As Boolean)
  myBase.Done = rhs
End Property

Private Property Get TaskInterface_Done() As Boolean
On Error Resume Next
 TaskInterface_Done = myBase.Done
End Property


Private Property Get TaskInterface_Owner() As Object
 Set TaskInterface_Owner = myBase.Owner
End Property

Public Property Let TaskInterface_Priority(ByVal Value As PriorityLevel)
  myBase.Priority = Value
End Property

Public Property Get TaskInterface_Priority() As PriorityLevel
On Error Resume Next
  TaskInterface_Priority = myBase.Priority
End Property
Private Property Let TaskInterface_ID(ByVal rhs As Long)
If rhs = jobnameID Then
myBase.Done = True
End If
End Property

Private Property Get TaskInterface_ID() As Long
TaskInterface_ID = jobnameID
End Property
Private Property Let TaskInterface_busy(ByVal rhs As Boolean)
mybusy = rhs
End Property

Private Property Get TaskInterface_busy() As Boolean
TaskInterface_busy = mybusy
End Property

' DELEGATED METHODS

Private Sub TaskInterface_Dispose(ByVal Action As DisposeAction)
  myBase.Dispose Action
  Set myBase = Nothing
End Sub




' PRIVATE ROUTINES

Private Sub Class_Initialize()
  Set myBase = New TaskBase
'
End Sub





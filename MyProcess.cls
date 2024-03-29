VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Implements TaskInterface

Private mybuf$
' load code or just use it as a counter
' load code first time
Private myPriority As PriorityLevel
Private myOwner As Object
Private myDone As Boolean
Private m_duration As Double
Private mYspace As Double
Private myHere$
Private myid As Long
Private mycode As String
Private MyMode As Long
Private mybusy As Boolean
Public hPipe As Long  ' �� ������� ���� ���;;
Private pipename As String  '= "\\.\pipe\bigtest"
Private myholdtime As Double
Private myBUFvar As String, res As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private myProcess As basetask  ' myprocess.process.id is the handler
Private myProcessPureParent As basetask
Private mNostretch As Boolean


Private Property Get TaskInterface_Process() As basetask
Set TaskInterface_Process = myProcess
End Property

Private Property Set TaskInterface_Process(aProcess As basetask)
If aProcess Is Nothing Then
    Set myOwner = Nothing
   myProcessPureParent.ThrowOne myid
    TaskInterface_Done = True
Else
Set myProcess = aProcess
Set myProcessPureParent = aProcess.Parent
While myProcessPureParent.IamThread = True
Set myProcessPureParent = myProcessPureParent.Parent
Wend

If myProcess.Process Is Me Then
myDone = True
Else
Set myProcess.Process = Me
Set myOwner = myProcess.Parent.Owner
' ������� ��� ������� (�� ��������� ���  �� ������)  �� �� ����� ����
' ���� ���� ���� �� ��������� �� ���� �� ��� ��� myProcess ��� ���
' � ��� ���� (�� �� this) �� ��� �� me.
MyMode = 1
End If
End If
End Property
Private Sub Class_Terminate()
If hPipe <> 0 Then
res = DisconnectNamedPipe(hPipe)
 CloseHandle hPipe
End If
If MyMode > 0 Then
myProcess.ThrowOne myid
    Set myProcess.Process = Nothing ' ��������...��� �� �����������!

    End If
    If Not myProcess Is Nothing Then
     Set myProcess.Owner = Nothing
     Set myProcess = Nothing
     End If
        
    Set myProcessPureParent = Nothing
End Sub

Private Property Let TaskInterface_interval(ByVal RHS As Double)
mYspace = RHS
m_duration = 0
If mYspace = 0 Then

Else
m_duration = uintnew(CDbl(timeGetTime)) + mYspace
End If
End Property

Private Property Get TaskInterface_interval() As Double
'

TaskInterface_interval = mYspace
End Property

' IMPLEMENTED PROPERTIES

Private Property Set TaskInterface_Owner(RHS As Object)

  Set myOwner = RHS
End Property

' IMPLEMENTED METHODS

Private Sub TaskInterface_Parameters(ParamArray Values() As Variant)
'' Usage: Private Sub TaskInterface_Parameters(Color As Long, Count As Long)
On Error GoTo poulos
  ' Verifing parameter count
  If UBound(Values) = 5 Then
    
    On Error Resume Next
    myid = CLng(Values(0))  ' this is the hanlde
   mYspace = CDbl(Values(1)) '
   mycode = CStr(Values(2))
   myholdtime = CDbl(Values(3)) ' ��� �� after ����
    m_duration = uintnew(CDbl(timeGetTime)) + myholdtime
   myHere$ = CStr(Values(4))  ' change from 5 to 4
   mNostretch = CBool(Values(5))
    If Err.Number > 0 Then
      On Error GoTo 0
    
    End If
    On Error GoTo 0
    ' do something
  Else
poulos:
    ' not defined yet
  End If
  
End Sub

Private Sub TaskInterface_TickolD()
 Dim x As Long, mycnt As Long, ohere$, look As Boolean
 Static offset As Double
 Dim NowProcess As basetask
 Set NowProcess = myProcess
Set NowProcess.Owner = myOwner
SetText myOwner
If MyMode = 0 Then
' ����� ��� ���� ���������� � �������!
myDone = True
End If
If NowProcess Is Nothing Then
myDone = True
End If
         
 If myDone Then GoTo there

 If m_duration <= uintnew(CDbl(timeGetTime)) Then
        m_duration = uintnew(CDbl(timeGetTime)) + mYspace

            If mybuf$ = "" Then mybuf$ = mycode: NowProcess.IFCTRL = 0: NowProcess.jump = False
            ohere$ = here$
            here$ = myHere$
 ''Set NowProcess.Owner = myOwner ' ������ �� ���� �������... ��� ��������???
            ' always...start run from here owner setting.
            x = 1
            look = True


        Call executeblock(x, NowProcess, mybuf$, False, look, (Interrupted))
        If Not (NowProcess Is Nothing) Then Set myOwner = NowProcess.Owner
        '' resync thread if we need more time new from 7 rev 19
        If Len(mybuf$) < 3 Then
        mybuf$ = ""
    '' If Not mNostretch And Not Interrupted Then
            ''      If m_duration < uintnew(CDbl(timeGetTime)) Then
   ''         mYspace = uintnew(CDbl(timeGetTime)) - m_duration - mYspace
  ''          End If
    
   ''  End If
        Else
      
    If Interrupted Then m_duration = 0      ' do it in past;;
    If Not look Then GoTo conthere
     End If
       '' If Not NowProcess.Parent.IamRunning Then
       If LastErNum <> 0 Or x = 0 Then
       myDone = True
       End If
        If x = 3 Then x = 1: look = False
        If ProcessEnd Then Exit Sub
        
        If (x = 1 Or (x = 2 And mybuf$ = "")) And look = True Then
        
            If NowProcess.Parent.Exist(myid, "_multi") Then
              TaskMaster.RestEnd1
              TaskMaster.RestEnd
              
            myProcessPureParent.ThrowThreads   '  ������� ��� ����� ��� ������� �� done

            MyMode = 0
            myDone = True
           ''
            GoTo there
            Else
            GoTo there
            End If
        End If
conthere:
If NOEXECUTION Then GoTo there
                here$ = ohere$
        If x = 0 Then
        myDone = True
        
        End If
If MOUT Then GoTo there

   


   End If
    If myDone Then GoTo there
   Exit Sub
there:

On Error GoTo 0

    Set myOwner = Nothing

   TaskInterface_Dispose tmThisTask
   
End Sub



Private Sub TaskInterface_Tick()
On Error Resume Next
 Dim x As Long, mycnt As Long, ohere$, look As Boolean
 Static offset As Double
 Dim NowProcess As basetask
 Set NowProcess = myProcess

If MyMode = 0 Then
' ����� ��� ���� ���������� � �������!
myDone = True
End If
If NowProcess Is Nothing Then
myDone = True
End If
         
 If myDone Then GoTo there

 If m_duration <= uintnew(CDbl(timeGetTime)) Then
        m_duration = uintnew(CDbl(timeGetTime)) + mYspace

            If mybuf$ = "" Then mybuf$ = mycode: NowProcess.IFCTRL = 0: NowProcess.jump = False
            ohere$ = here$
            here$ = myHere$
            x = 1
            look = True
            
Set NowProcess.Owner = myOwner
If Interrupted Then SetTextBasketBack myOwner, players(GetCode(myOwner))
 Call executeblock(x, NowProcess, mybuf$, False, look, Interrupted)
  If Not (NowProcess Is Nothing) Then Set myOwner = NowProcess.Owner
        If Len(mybuf$) < 3 Then
        mybuf$ = ""
        Else
      
     If Interrupted Then m_duration = 0           ' do it in past;;
     
    If Not look Then GoTo conthere
     End If
       '' If Not NowProcess.Parent.IamRunning Then
       If LastErNum <> 0 Or x = 0 Then
       myDone = True
       End If
        If x = 3 Then x = 1: look = False
        If ProcessEnd Then Exit Sub
        
        If (x = 1 Or (x = 2 And mybuf$ = "")) And look = True Then
        
            If NowProcess.Parent.Exist(myid, "_multi") Then
              TaskMaster.RestEnd1
              TaskMaster.RestEnd
              
            myProcessPureParent.ThrowThreads   '  ������� ��� ����� ��� ������� �� done

            MyMode = 0
            myDone = True
           ''
            GoTo there
            Else
            GoTo there
            End If
        End If
conthere:
If NOEXECUTION Then GoTo there
                here$ = ohere$
        If x = 0 Then
        myDone = True
        
        End If
If MOUT Then GoTo there

   


   End If
    If myDone Then GoTo there
    
   Exit Sub
there:
On Error GoTo 0

    Set myOwner = Nothing

   TaskInterface_Dispose tmThisTask
   
End Sub


' DELEGATED PROPERTIES
Private Property Let TaskInterface_Done(ByVal RHS As Boolean)
myDone = RHS
If MyMode > 0 Then
If Not myProcess.IamRunning Then TaskInterface_Dispose tmAllTasks
End If
End Property

Private Property Get TaskInterface_Done() As Boolean
On Error Resume Next
 TaskInterface_Done = myDone
End Property


Private Property Get TaskInterface_Owner() As Object
 Set TaskInterface_Owner = myOwner
End Property

Public Property Let TaskInterface_Priority(ByVal Value As PriorityLevel)
  myPriority = Value
End Property


Public Property Get TaskInterface_Priority() As PriorityLevel
  TaskInterface_Priority = myPriority
End Property
Private Property Let TaskInterface_ID(ByVal RHS As Long)
If RHS = myid Then
myDone = True
Else
' ��� �������
End If
End Property

Private Property Get TaskInterface_ID() As Long
TaskInterface_ID = myid
End Property
Private Property Let TaskInterface_busy(ByVal RHS As Boolean)
mybusy = RHS
End Property

Private Property Get TaskInterface_busy() As Boolean
TaskInterface_busy = mybusy
End Property

' DELEGATED METHODS

Private Sub TaskInterface_Dispose(ByVal Action As DisposeAction)
On Error Resume Next
myDone = True
On Error Resume Next
If Not TaskMaster Is Nothing Then TaskMaster.rest
    Set myOwner = Nothing
If MyMode > 0 Then
 If Not myProcess.IamRunning Then
MyMode = 0

   Set myProcess.Process = Nothing

   If Not ProcessEnd Then Set myProcess = Nothing
       
    End If
   End If

If Not TaskMaster Is Nothing Then TaskMaster.RestEnd
End Sub
Private Function ProcessEnd() As Boolean
On Error Resume Next
If myProcess Is Nothing Then
ProcessEnd = True
Else
myProcess.EndRun
End If
End Function






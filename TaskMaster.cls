VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' The TaskMaster handles queing and timing of the
' various tasks.
Dim mymes As Long
Dim tickdrop As Long
Dim holdtasks As Boolean
Public Enum DisposeAction
  tmThisTask
  tmAllTasks
End Enum

Public Enum PriorityLevel
  tmLow
  tmHigh
End Enum


Public MusicTaskNum As Long
Public OnlyMusic As Boolean
Private myRunningList As Collection
Private myMasterList As Collection
Private myInterval As Long
Private myTimerHandle As Long
Private myTimerIsRunning As Boolean


' PUBLIC PROPERTIES
Public Property Get PlayMusic() As Boolean
' check if music is there

PlayMusic = MusicTaskNum > 0
Exit Sub

Dim i As Long
rest
       If myMasterList.Count > 0 Then
       For i = 1 To myMasterList.Count
       If myMasterList(i).Id < 100 Then
 PlayMusic = True
GoTo thh11
       End If
       
       Next i
        
        End If
       If myRunningList.Count > 0 Then
       For i = 1 To myRunningList.Count
       If myRunningList(i).Id < 100 Then
        PlayMusic = True
 GoTo thh11
       End
       End If
       
       Next i
        
        End If
thh11:
        RestEnd
 
End Property
Public Property Get QueueCount() As Long
  QueueCount = myRunningList.Count + myMasterList.Count
End Property


Public Property Get Interval() As Long
  Interval = myInterval
End Property


Public Property Let Interval(ByVal Value As Long)
  myInterval = Abs(Value)
End Property


Public Property Get Processing() As Boolean
  Processing = myTimerIsRunning
End Property

Public Function ThrowOne(thisid As Long) As Boolean
Dim i As Long
rest
       If myMasterList.Count > 0 Then
       For i = 1 To myMasterList.Count
       If myMasterList(i).Id = thisid Then
       myMasterList.remove i
       ThrowOne = True
GoTo thh
       End If
       
       Next i
        
        End If
       If myRunningList.Count > 0 Then
       For i = 1 To myRunningList.Count
       If myRunningList(i).Id = thisid Then
       
       myRunningList(i).Done = True
       myRunningList.remove i
       ThrowOne = True
 GoTo thh
       End
       End If
       
       Next i
        
        End If
        ' SECONDTIME
               If myMasterList.Count > 0 Then
       For i = 1 To myMasterList.Count
       If myMasterList(i).Id = thisid Then
       myMasterList.remove i
       ThrowOne = True
GoTo thh
       End If
       
       Next i
        
        End If
       If myRunningList.Count > 0 Then
       For i = 1 To myRunningList.Count
       If myRunningList(i).Id = thisid Then
       myRunningList(i).Done = True
       myRunningList.remove i
       ThrowOne = True
 GoTo thh
       End
       End If
       
       Next i
        
        End If
       
thh:
        RestEnd
        If Me.QueueCount <= 0 Then
        StopProcess
        End If
End Function
Private Function remove(thisid As Long) As Boolean
Dim i As Long
rest

       If myRunningList.Count > 0 Then
       For i = 1 To myRunningList.Count
       If myRunningList(i).Id = thisid Then
       myRunningList.remove i
       remove = True
       RestEnd
       Exit Function
       End
       End If
       
       Next i
        
        End If
        RestEnd
End Function

' PUBLIC METHODS

Public Sub AddTask(ByRef NewTask As TaskInterface, Optional Priority As PriorityLevel)
  ' Add to list
 
  If Priority = tmHigh And myMasterList.Count > 0 Then
    myMasterList.Add NewTask, "", 1

  Else
    myMasterList.Add NewTask

  End If

    StartTimer

End Sub
Private Property Get getone(thisid As Long) As TaskInterface
Dim i As Long
rest
       If myMasterList.Count > 0 Then
       For i = 1 To myMasterList.Count
       If myMasterList(i).Id = thisid Then
      Set getone = myMasterList(i)
GoTo thh
       End If
       
       Next i
        
        End If
       If myRunningList.Count > 0 Then
       For i = 1 To myRunningList.Count
       If myRunningList(i).Id = thisid Then
        Set getone = myRunningList(i)
 GoTo thh
       End
       End If
       
       Next i
        
        End If
thh:
        RestEnd
End Property
Public Sub Message(Mes As Long, code As Long, Optional PARAM As Long = 0)
Dim ok As Boolean
Dim task As TaskInterface
Set task = getone(Mes)
ok = False
If Not (task Is Nothing) Then
        Select Case code
        Case 0 ' hold
            task.busy = True
        Case 1
            task.Dispose tmAllTasks
            ThrowOne Mes
        Case 2 ' Restart
            task.busy = False
        Case 3
        task.Interval = PARAM
        
        End Select
        

        If ok Then code = -1
End If
End Sub




Public Sub Dispose()
  On Error Resume Next
        StopTimer
        Do While myMasterList.Count > 0
          myMasterList(1).Dispose tmAllTasks
          myMasterList.remove 1
        Loop
            Do While myRunningList.Count > 0
        
          myRunningList(1).Dispose tmAllTasks
          myRunningList.remove 1
        Loop
   
End Sub

Public Sub TimerTick()
If tickdrop > 0 Then Exit Sub
Dim this As TaskInterface
  If myMasterList.Count = 0 Then Exit Sub
  
  Do
  
  Set this = myMasterList(1)
  myMasterList.remove 1

  
 If OnlyMusic Then
 If this.Id > 100 Then
 AddTask this
 Exit Sub
 End If
 ElseIf holdtasks Then
 AddTask this, this.Priority
 Exit Sub
 End If
  myRunningList.Add this
  Loop While this.Id < 0 And myMasterList.Count <> 0
  If this.Id < 0 Then Exit Sub
  
  If Not this.busy And Not this.Done Then

  this.Tick
  End If
  If remove(this.Id) Then
  If this.Done Then

    this.Dispose tmThisTask
  Else
If Not NOEXECUTION Then
AddTask this, this.Priority
Else
this.Dispose tmThisTask
End If
    
  End If
  Else
    If this.Done Then

  Else

''  AddTask this, this.Priority
    
  End If
    this.Dispose tmThisTask
  End If

  If myTimerIsRunning = False Or QueueCount = 0 Then
    StopTimer

       ElseIf QueueCount = 1 And taskmainonly Then
    Sleep 50
  End If

  If LastErNum <> 0 Then Dispose
End Sub
Public Sub TimerTickNow()
Dim this As TaskInterface
  If myMasterList.Count = 0 Then Exit Sub
  
  Do
  
  Set this = myMasterList(1)
  myMasterList.remove 1

  
 If OnlyMusic Then
 If this.Id > 100 Then
 AddTask this
 Exit Sub
 End If
 ElseIf holdtasks Then
 AddTask this, this.Priority
 Exit Sub
 End If
  myRunningList.Add this
  Loop While this.Id < 0 And myMasterList.Count <> 0
  If this.Id < 0 Then Exit Sub
  
  If Not this.busy And Not this.Done Then

  this.Tick
  End If
  If remove(this.Id) Then
  If this.Done Then

    this.Dispose tmThisTask
  Else
If Not NOEXECUTION Then
AddTask this, this.Priority
Else
this.Dispose tmThisTask
End If
    
  End If
  Else
    If this.Done Then
    this.Dispose tmThisTask
  Else

    AddTask this, this.Priority
    
  End If
    this.Dispose tmThisTask
  End If

  If myTimerIsRunning = False Or QueueCount = 0 Then
    StopTimer

       ElseIf QueueCount = 1 And taskmainonly Then
    Sleep 50
  End If

  If LastErNum <> 0 Then Dispose
End Sub

Public Sub StopProcess()
  myTimerIsRunning = False
  StopTimer

End Sub


Public Sub StartProcess()
  StartTimer
End Sub


' PRIVATE ROUTINES

Private Sub Class_Initialize()
  Set myMasterList = New Collection
  Set myRunningList = New Collection

  myInterval = 5

End Sub


Private Sub Class_Terminate()
  If QueueCount > 0 Then
    Dispose
  End If
  Set myMasterList = Nothing
  Set myRunningList = Nothing
End Sub


Private Sub StartTimer()
  If myTimerHandle = 0 Then
  If QueueCount > 0 Then
 ''MasterTimer = 0
  
    myTimerHandle = SetTimer(0&, 0&, myInterval, AddressOf TaskMasterTick)
    myTimerIsRunning = (myTimerHandle <> 0)
  End If
  End If
End Sub


Private Sub StopTimer()
  If myTimerHandle <> 0 Then
  '' MasterTimer = 0
    KillTimer 0&, myTimerHandle
    myTimerHandle = 0
    myTimerIsRunning = False
  End If
End Sub

Public Sub rest()
'
''holdtasks = True
tickdrop = tickdrop + 1
End Sub
Public Sub RestEnd()
tickdrop = tickdrop - 1
If tickdrop <= 0 Then tickdrop = 0: holdtasks = False
If tickdrop > 0 Then If QueueCount = 0 Then tickdrop = 0:: holdtasks = False
End Sub
Public Sub Rest1()
'

tickdrop = tickdrop + 1
End Sub
Public Sub RestEnd1()
tickdrop = 0
End Sub

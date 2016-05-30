VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M2000"
   ClientHeight    =   570
   ClientLeft      =   -47955
   ClientTop       =   48315
   ClientWidth     =   1365
   Icon            =   "SMALL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   1365
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   930
      Top             =   360
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hideme As Boolean
Private foundform5 As Boolean
Private reopen4 As Boolean, reopen2 As Boolean
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function GetModuleHandleW Lib "KERNEL32" (ByVal lpModuleName As Long) As Long


Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Long, ByVal lpProcName As String) As Long


Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long


Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
    Private Const GWL_WNDPROC = -4
    Private m_Caption As String


Public Property Get CaptionW() As String
    CaptionW = m_Caption
End Property


Public Property Let CaptionW(ByRef NewValue As String)
    Static WndProc As Long, VBWndProc As Long
    m_Caption = NewValue
    ' get window procedures if we don't have
    '     them
ttl = True

    If WndProc = 0 Then
        ' the default Unicode window procedure
        WndProc = GetProcAddress(GetModuleHandleW(StrPtr("user32")), "DefWindowProcW")
        ' window procedure of this form
        VBWndProc = GetWindowLongA(hWnd, GWL_WNDPROC)
    End If
    ' ensure we got them


    If WndProc <> 0 Then
        ' replace form's window procedure with t
        '     he default Unicode one
        SetWindowLongW hWnd, GWL_WNDPROC, WndProc
        ' change form's caption
        SetWindowTextW hWnd, StrPtr(m_Caption)
        ' restore the original window procedure
        SetWindowLongA hWnd, GWL_WNDPROC, VBWndProc
    Else
        ' no Unicode for us
        Caption = m_Caption
        
    End If
End Property
' usage sample



'** Function **

'Private onlyone As Boolean
Public Function ask(bstack As basetask, A$) As Double
If ASKINUSE Then Exit Function
DialogSetupLang DialogLang
AskText$ = A$
ask = NeoASK(bstack)

End Function
Public Function NeoASK(bstack As basetask) As Double
If ASKINUSE Then Exit Function
Dim oldesc As Boolean
oldesc = escok
'using AskTitle$, AskText$, AskCancel$, AskOk$, AskDIB$
Static once As Boolean
If once Then Exit Function
once = True
ASKINUSE = True
If TypeOf Screen.ActiveForm Is GuiM2000 Then Screen.ActiveForm.UNhookMe

Dim INFOONLY As Boolean
k1 = 0
If AskTitle$ = "" Then AskTitle$ = MesTitle$
If AskCancel$ = "" Then INFOONLY = True
If AskOk$ = "" Then AskOk$ = "OK"
If Form1.Visible Then
MyDoEvents1 Form1
Sleep 1
NeoMsgBox.Show , Form1
Else
If form5iamloaded Then
MyDoEvents1 Form5
Sleep 1
NeoMsgBox.Show , Form5
Else
NeoMsgBox.Show
End If
End If
On Error Resume Next
''SleepWait3 10
Sleep 1
If Form1.Visible Then
Form1.Refresh
ElseIf form5iamloaded Then
Form5.Refresh
Else
MyDoEvents
End If
Sleep 1
While Not NeoMsgBox.Visible
    MyDoEvents
Wend
NeoMsgBox.ZOrder 0
If AskInput Then
NeoMsgBox.gList3.SetFocus
End If
    
  If bstack.ThreadsNumber = 0 Then
    On Error Resume Next
    If Not (bstack.toback Or bstack.toprinter) Then If bstack.Owner.Visible Then bstack.Owner.Refresh
    End If
    If Not NeoMsgBox.Visible Then
    NeoMsgBox.Visible = True
    MyDoEvents
    End If
    Dim mycode As Double, oldcodeid As Double, x As Form
mycode = Rnd * 12312314
oldcodeid = ModalId
 For Each x In Forms
                            If x.Visible And x.name = "GuiM2000" Then
                     
                           If x.Enablecontrol Then
                               x.Modal = mycode
                                x.Enablecontrol = False
                            End If
                            End If
                    Next x
                     Set x = Nothing
If INFOONLY Then
NeoMsgBox.command1(0).SetFocus
End If
ModalId = mycode
Do
        mywaitOld bstack, 5
      Sleep 1
Loop Until NOEXECUTION Or Not ASKINUSE
 ModalId = mycode
k1 = 0
 BLOCKkey = True
While KeyPressed(&H1B) ''And UseEsc

ProcTask2 bstack
NOEXECUTION = False
Wend
BLOCKkey = False
AskTitle$ = ""
Dim z As Form
 Set z = Nothing

           For Each x In Forms
            If x.Visible And x.name = "GuiM2000" Then
            If Not x.Enablecontrol Then x.TestModal mycode
          If x.Enablecontrol Then Set z = x
            End If
            Next x
             Set x = Nothing
          If Typename(z) = "GuiM2000" Then
            z.ShowmeALL
            z.SetFocus
            Set z = Nothing
          End If
          ModalId = oldcodeid
If INFOONLY Then
NeoASK = 1
Else
NeoASK = Abs(AskCancel$ = "") + 1
End If
If NeoASK = 1 Then
If AskInput Then
bstack.soros.PushStr AskStrInput$
End If
End If
AskCancel$ = ""
once = False
ASKINUSE = False
INK$ = ""
On Error Resume Next
If Not bstack.Owner Is Nothing Then
If bstack.Owner.Visible Then
bstack.Owner.SetFocus
End If
End If
  escok = oldesc
End Function
Sub mywait(bstack As basetask, PP As Double, Optional SLEEPSHORT As Boolean = False)
Dim p As Boolean, e As Boolean
On Error Resume Next
If bstack.Process Is Nothing Then
''If extreme Then MyDoEvents1 Form1
If PP = 0 Then Exit Sub
Else

Err.Clear
p = bstack.Process.Done
If Err.number = 0 Then
e = True
If p <> 0 Then
Exit Sub
End If
End If
End If

PP = PP + CDbl(timeGetTime)

Do


'If TaskMaster.Processing And Not bstack.TaskMain Then
        'If Not bstack.toprinter Then bstack.Owner.Refresh
      
      '  If Not TaskMaster Is Nothing Then TaskMaster.TimerTickNow
        If Form1.DIS.Visible And Not bstack.toprinter Then
        MyDoEvents0 Form1.DIS
   
        Else
        MyDoEvents0 Me
        End If
'Els
        'MyDoEvents
'End If
If SLEEPSHORT Then Sleep 1
If e Then
p = bstack.Process.Done
If Err.number = 0 Then
If p <> 0 Then
Exit Do
End If
End If
End If
Loop Until PP <= CDbl(timeGetTime) Or NOEXECUTION Or MOUT

                       If exWnd <> 0 Then
                MyTitle$ bstack
                End If
End Sub

Private Sub mywaitOld(bstack As basetask, PP As Double)
Dim p As Boolean, e As Boolean

On Error Resume Next
If bstack.ThreadsNumber = 0 Then GoTo cont1
If bstack.Process Is Nothing Then
''If extreme Then MyDoEvents
If PP = 0 Then Exit Sub
Else

Err.Clear
p = bstack.Process.Done
If Err.number = 0 Then
e = True
If p <> 0 Then
Exit Sub
End If
End If
End If
cont1:
PP = PP + CDbl(timeGetTime)

Do


If TaskMaster.Processing And Not bstack.TaskMain Then
        If Not bstack.toprinter Then bstack.Owner.Refresh
        TaskMaster.TimerTick
       ' SleepWait 1
       MyDoEvents
       
Else
        ' SleepWait 1
        MyDoEvents
        End If
If e Then
p = bstack.Process.Done
If Err.number = 0 Then
If p <> 0 Then
Exit Do
End If
End If
End If
Loop Until PP <= CDbl(timeGetTime) Or NOEXECUTION Or MOUT

                       If exWnd <> 0 Then
                MyTitle$ bstack
                End If
            
End Sub


Private Sub Form_Activate()
If QRY Then
If Form1.Visible Then Form1.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If QRY Or GFQRY Then
If Form1.Visible Then Form1.SetFocus
ElseIf KeyCode = 27 And ASKINUSE Then
    NOEXECUTION = True
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If QRY Or GFQRY Then
If Form1.Visible Then Form1.SetFocus
End If
If Not BLOCKkey Then INK$ = INK$ & Chr(KeyAscii)
End Sub

Private Sub Form_Load()
Debug.Assert (InIDECheck = True)
ttl = True
Timer1.Interval = 10000
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
If exWnd <> 0 Then
Form1.IEUP ("")
Cancel = True
Exit Sub
End If
Timer1.enabled = False
NOEXECUTION = True
ExTarget = True
INK$ = Chr(27)
If Not TaskMaster Is Nothing Then
TaskMaster.Dispose
End If
NOEDIT = True
MOUT = True
Cancel = True
Else
ttl = False
End If


End Sub

Private Sub Form_Resize()

 hideme = (Me.WindowState = 1)
 If hideme Then
 reopen2 = False
 reopen4 = False
 If Form4.Visible Then Form4.Visible = False: reopen4 = True
 If Form3.Visible Then If trace Then Form2.Visible = False: reopen2 = True
 End If
 Timer1.enabled = Timer1.Interval < 10000
 
End Sub



Private Sub Timer1_Timer()
' On Error Resume Next
Dim x As Form, z As Long
If DIALOGSHOW Or ASKINUSE Or ModalId <> 0 Then
Timer1.enabled = False
Exit Sub
End If
Timer1.enabled = False
Timer1.Interval = 20
If Not hideme Then
If Not Form1.Visible Then
If foundform5 Then
Form5.Visible = True
'DoEvents
End If
If Not ttl Then
ttl = True
z = Form1.top
Form1.top = 0
If Not IsSelectorInUse Then Form1.Show , Form5
Else
If Not IsSelectorInUse Then Form1.Show , Form5
End If
'DoEvents
End If

'Sleep 500
If Form1.Visible And Not IsSelectorInUse Then
'Form1.ZOrder
If Not trace Then reopen2 = False
If vH_title$ = "" Then reopen4 = False
If reopen4 Then Form4.Show , Form1: Form4.Visible = True
If reopen2 Then Form2.Show , Form1: Form2.Visible = True
   For Each x In Forms
       If Typename$(x) = "GuiM2000" Then
       If x.Visible Then
       x.Visible = False
       x.Show , Form1
       End If
       End If
       Next
       
Sleep 1
If Forms.Count > 5 Then

Else
Form1.SetFocus
End If
Form1.ZOrder 0
Sleep 2
 Set x = Nothing
End If
Else
If Not ((exWnd <> 0) Or AVIRUN Or IsSelectorInUse) Then
Form1.Hide
If Form5.Visible Then Form5.Visible = False: foundform5 = True
End If


End If
End Sub

Public Function InIDECheck() As Boolean
    m_bInIDE = True
    InIDECheck = True
End Function

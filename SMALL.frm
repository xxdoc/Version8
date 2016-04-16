VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M2000"
   ClientHeight    =   765
   ClientLeft      =   -47955
   ClientTop       =   48315
   ClientWidth     =   1530
   Icon            =   "SMALL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   1530
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   360
      Top             =   240
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private hideme As Boolean
Private foundform5 As Boolean
Private reopen4 As Boolean, reopen2 As Boolean
Private Declare Function GetTextMetrics Lib "gdi32" _
Alias "GetTextMetricsA" (ByVal hDC As Long, _
lpMetrics As TEXTMETRIC) As Long
Private Type TEXTMETRIC
tmHeight As Long
tmAscent As Long
tmDescent As Long
tmInternalLeading As Long
tmExternalLeading As Long
tmAveCharWidth As Long
tmMaxCharWidth As Long
tmWeight As Long
tmOverhang As Long
tmDigitizedAspectX As Long
tmDigitizedAspectY As Long
tmFirstChar As Byte
tmLastChar As Byte
tmDefaultChar As Byte
tmBreakChar As Byte
tmItalic As Byte
tmUnderlined As Byte
tmStruckOut As Byte
tmPitchAndFamily As Byte
tmCharSet As Byte
End Type
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Dim tm As TEXTMETRIC
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long


Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long


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
Public Function InternalLeadingSpace() As Long
On Error Resume Next
    GetTextMetrics hDC, tm
  With tm
InternalLeadingSpace = (tm.tmInternalLeading = 0) Or Not (tm.tmInternalLeading > 0)
End With
End Function
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
Static Once As Boolean
If Once Then Exit Function
Once = True
ASKINUSE = True
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
    Dim mycode As Variant
mycode = Rnd * 12312314

For Each x In Forms
If x.Visible And x.name = "GuiM2000" Then

If Not x.enabled = False Then
x.Modal = mycode
x.enabled = False
End If
End If
Next x
If INFOONLY Then
NeoMsgBox.command1(0).SetFocus
End If
Do

        mywait bstack, 5
      Sleep 1
Loop Until NOEXECUTION Or Not ASKINUSE
k1 = 0
 BLOCKkey = True
While KeyPressed(&H1B) ''And UseEsc

ProcTask2 bstack
NOEXECUTION = False
Wend
BLOCKkey = False
AskTitle$ = ""
For Each x In Forms
If x.Visible And x.name = "GuiM2000" Then
x.TestModal mycode
End If
Next x
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
Once = False
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
Private Sub mywait(bstack As basetask, PP As Double)
Dim p As Boolean, e As Boolean

On Error Resume Next
If bstack.ThreadsNumber = 0 Then GoTo cont1
If bstack.Process Is Nothing Then
''If extreme Then MyDoEvents
If PP = 0 Then Exit Sub
Else

Err.Clear
p = bstack.Process.Done
If Err.Number = 0 Then
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
If Err.Number = 0 Then
If p <> 0 Then
Exit Do
End If
End If
End If
Loop Until PP <= CDbl(timeGetTime) Or NOEXECUTION Or MOUT

                       If exWnd <> 0 Then
                mytitle$ bstack
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
'''Debug.Print "FORM3 LOADED"
ttl = True
'Icon = Form2.Icon
'hideme = True
 CaptionW = ""
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
If Not TaskMaster Is Nothing Then
TaskMaster.Dispose
End If
ttl = False
End If


End Sub

Private Sub Form_Resize()

 hideme = Me.WindowState = 1
 If hideme Then
 reopen2 = False
 reopen4 = False
 If Form4.Visible Then Form4.Visible = False: reopen4 = True
 If Form3.Visible Then If trace Then Form2.Visible = False: reopen2 = True
 
 End If
'Debug.Print "RESIZE ME"
 Timer1.enabled = True
End Sub

Private Sub Timer1_Timer()
' On Error Resume Next
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
If Not IsSelectorInUse Then Form1.Show , Form5
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

Form1.SetFocus
Form1.ZOrder 0

End If
Else
If Not ((exWnd <> 0) Or AVIRUN Or IsSelectorInUse) Then
Form1.Hide
If Form5.Visible Then Form5.Visible = False: foundform5 = True
End If


End If
End Sub
Sub StoreFont(aName$, aSize As Single, ByVal aCharset As Long)
On Error Resume Next
Form3.Font.Size = aSize
If Err.Number > 0 Then aSize = 12: Form3.Font.Size = aSize
    Form3.FontName = aName$
    Form3.Font.bold = True
    Form3.Font.Italic = True
    Form3.Font.charset = aCharset
        Form3.FontName = aName$
    Form3.Font.bold = True
    Form3.Font.Italic = True
    Form3.Font.charset = aCharset
    Form3.Font.Size = aSize
    aSize = Form3.Font.Size '' return
End Sub

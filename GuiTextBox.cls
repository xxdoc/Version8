VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' If you copy and paste code
' don't forget to set default property to vartext (changed from text)
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Event ValidString(ThatString As String, setpos As Long)
Event KeyDown(KeyCode As Integer, shift As Integer)
Event LostFocus()
Event GotFocus()
Event KeyPress(KeyAscii As Integer)
Event MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Event SpinnerValue(ThisValue As Long)
Private mprompt As String
Private mkind As String
Public SkipProcAtExitFocus As Boolean
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_DISCARDED = &H4000
Private Const GMEM_FIXED = &H0
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GMEM_LOCKCOUNT = &HFF
Private Const GMEM_MODIFY = &H80
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetClipboardData Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Declare Function SetClipboardData Lib "user32" _
    (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function OpenClipboard Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" _
    () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private pad$
    ' for Gui purposes
Public CtrlName As String
Dim mIndex As Long
Dim Callback As GuiM2000
Public Property Get GetCallBack() As Form
Set GetCallBack = Callback
End Property
 Public Sub Construct(there As Form, name$)
 mIndex = -1
 Set Callback = there
 On Error Resume Next
 Set glistN = there.Controls.Add("m2000.glist", name$)
 If Err.Number > 0 Then
 Set glistN = there.Controls.Add("m2000.glist", GetTag$)
 End If
 CtrlName = name$
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(name$).TabStop = True
End Sub
Public Sub ConstructArray(there As Form, name$, i As Long)
 Set Callback = there
 On Error Resume Next
 If i = 0 Then
 Set glistN = there.Controls.Add("m2000.glist", name$)
 Else
  Set glistN = there.Controls.Add("m2000.glist", name$ + "_" + LTrim(Str$(i)))

 End If
 If Err.Number > 0 Then
 Set glistN = there.Controls.Add("m2000.glist", GetTag$)
 End If
 CtrlName = name$
 mIndex = i
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(name$).TabStop = True
End Sub
Public Property Get Text() As String
If glistN.listcount = 0 Then
glistN.additem mprompt
Else
Text = glistN.List(0)
End If
End Property
Public Property Get vartext() As String
If glistN.listcount = 0 Then
glistN.additem ""
Else
vartext = Mid$(glistN.List(0), 1 + Len(mprompt))
End If
End Property
Public Property Let vartext(ByVal rhs As String)
glistN.restrictLines = 1
glistN.HeadLine = "" ' no info
glistN.HeadlineHeight = 0

If glistN.listcount = 0 Then
glistN.additem mprompt + rhs
Else
glistN.List(0) = mprompt + rhs
End If
glistN.SelStart = Len(mprompt + rhs) + 1
glistN.ShowMe
End Property
Public Property Let Text(ByVal rhs As String)
'Dim oldright As Long
'oldright = glistN.LeftMarginPixels
glistN.restrictLines = 1
glistN.HeadLine = "" ' no info
glistN.HeadlineHeight = 0
If glistN.listcount = 0 Then
glistN.additem rhs
Else
glistN.List(0) = rhs
End If
'glistN.LeftMarginPixels = oldright

glistN.ShowMe2
End Property
Public Property Get Info() As String

Info = glistN.HeadLine

End Property

Public Property Let Info(ByVal rhs As String)
Locked = False
glistN.restrictLines = 0
glistN.HeadLineHeightMinimum = glistN.HeightPixels
glistN.HeadLine = rhs

glistN.ShowMe
Locked = True
End Property

Public Property Get top() As Long
top = glistN.TopTwips
End Property
Public Property Let top(rhs As Long)
glistN.TopTwips = rhs
End Property
Public Sub SetUp()
Dim ff As New StdFont
With Callback
ff.name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With

With glistN
.Enabled = True
Info = ""
If .listcount = 0 Then Text = ""
.NoCaretShow = True
.VerticalCenterText = True
If Not .CenterText Then .LeftMarginPixels = 8
.BackColor = rgb(200, 120, 60)
Set .Font = ff
.ForeColor = Callback.Controls(0).ForeColor
.dcolor = Callback.Controls(0).dcolor
.ShowMe
End With
Locked = False
Enabled = True
Retired

End Sub


Public Property Get Prompt() As String
Prompt = mprompt
End Property

Public Property Let Prompt(ByVal rhs As String)
mprompt = rhs
glistN.PromptLineIdent = Len(rhs)
Text = rhs
End Property



Private Sub glistN_RegisterGlist(this As gList)
If Not glistN.Spinner Then this.NoWheel = True
On Error Resume Next
'glistN.Parent.hookme this
If Err.Number > 0 Then this.NoWheel = True

End Sub



Private Sub glistN_ChangeListItem(item As Long, content As String)
If Not Enabled Then Exit Sub
Dim this As String, thispos As Long, vr(2)
thispos = glistN.SelStart - Len(mprompt)
this = Mid$(content, 1 + Len(mprompt))
'RaiseEvent ValidString(this, thispos)
vr(0) = this
vr(1) = thispos
   If mIndex >= 0 Then
   Callback.CallbackNow CtrlName + ".ValidString(" + CStr(mIndex) + ")", vr()
   Else
      Callback.CallbackNow CtrlName + ".ValidString()", vr()
      End If
      
thispos = vr(1) + Len(mprompt)
content = mprompt + vr(0)
If thispos <> glistN.SelStart And this <> "" Then glistN.SelStart = thispos: glistN.ShowMe
End Sub

Public Sub ResetPan()
glistN.PanPos = 0
End Sub

Private Sub glistN_ChangeSelStart(thisselstart As Long)
If thisselstart <= Len(mprompt) Then thisselstart = Len(mprompt) + 1
End Sub

Public Property Get ThisKind() As String
ThisKind = mkind
End Property

Public Property Let ThisKind(ByVal rhs As String)
mkind = rhs
glistN.LastLinePart = " " & rhs
End Property


Public Property Get Enabled() As Boolean
Enabled = glistN.EditFlag
End Property

Public Property Let Enabled(ByVal rhs As Boolean)
glistN.EditFlag = rhs
glistN.NoCaretShow = Not rhs
End Property
'


Private Sub glistN_CheckGotFocus()
If Not Locked Then
If glistN.listindex < 0 And glistN.listcount > 0 Then glistN.ListindexPrivateUse = 0
RaiseEvent GotFocus
If Enabled Then
glistN.ForeColor = rgb(0, 0, 0)
glistN.BackColor = rgb(255, 255, 255)
glistN.NoCaretShow = False:
glistN.PanPos = 0
Else
glistN.ForeColor = rgb(255, 255, 255)
glistN.BackColor = rgb(100, 200, 160)
glistN.NoCaretShow = True
End If
glistN.ShowMe
End If

End Sub
Public Sub Retired()
glistN.ForeColor = rgb(255, 255, 255)
glistN.BackColor = rgb(200, 120, 60)
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
End Sub

Private Sub glistN_CheckLostFocus()
If Not Locked Then
If Not SkipProcAtExitFocus Then
glistN.ForeColor = rgb(255, 255, 255)
glistN.BackColor = rgb(200, 120, 60)
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
Else
'
End If
End If
RaiseEvent LostFocus
End Sub

Private Sub glistN_HeaderSelected(Button As Integer)
' INFO SELECTED
Text = Text
End Sub

Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)
Dim vr(2)
vr(0) = KeyCode
vr(1) = shift
   If mIndex >= 0 Then
   Callback.CallbackNow CtrlName + ".KeyDown(" + CStr(mIndex) + ")", vr()
   Else
      Callback.CallbackNow CtrlName + ".KeyDown()", vr()
      End If
KeyCode = vr(0)
shift = vr(1)
If KeyCode = vbKeyUp Then
ElseIf KeyCode = vbKeyDown Then
ElseIf KeyCode = vbKeyV And (shift And 3) = 2 Then
If Not Enabled Then Exit Sub
pad$ = GetTextData(13)
If pad$ = "" Then pad$ = Clipboard.GetText(1)
If pad$ <> "" Then vartext = pad$
End If

End Sub

Public Property Get SelStart() As Long
SelStart = glistN.SelStart
End Property

Public Property Let SelStart(ByVal rhs As Long)
glistN.SelStart = rhs
End Property

Public Property Get Visible() As Boolean
Visible = glistN.Visible
End Property

Public Property Let Visible(ByVal rhs As Boolean)
glistN.Visible = rhs
End Property
Public Sub SetFocus()
On Error Resume Next
glistN.SetFocus
End Sub

Private Sub glistN_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, shift, x, y)
End Sub

Private Sub glistN_Selected(item As Long)
If item >= 0 Then
   '' If (Not mCallback Is Nothing) Then mCallback.Press Index
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Click(" + CStr(mIndex) + ")"
   Else
      Callback.Callback CtrlName + ".Click()"
      End If
    End If
End Sub

Private Sub glistN_Selected2(item As Long)
If item >= 0 Then
  If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Enter(" + CStr(mIndex) + ")"
   Else
      Callback.Callback CtrlName + ".Enter()"
      End If
End If
End Sub


Private Sub glistN_SyncKeyboard(item As Integer)
RaiseEvent KeyPress(item)
End Sub
Private Sub glistN_SpinnerValue(ThatValue As Long)
''RaiseEvent SpinnerValue(ThatValue)
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".SpinnerValue(" + CStr(mIndex) + "," + LTrim(Str(ThatValue)) + ")"
   Else
      Callback.Callback CtrlName + ".SpinnerValue(" + LTrim(Str(ThatValue)) + ")"
      End If
     '' vartext = CStr(ThatValue)
End Sub
Public Property Get Locked() As Boolean
Locked = Not glistN.Enabled
End Property

Public Property Let Locked(ByVal rhs As Boolean)
glistN.Enabled = Not rhs
End Property
Public Sub Spinner(ShowBar As Boolean, MinValue As Long, maxvalue As Long, LargeValue As Long)
If Not ShowBar Then
glistN.Spinner = False
glistN.ShowBar = False
glistN.EnabledBar = False
Else
glistN.Enabled = False
glistN.restrictLines = 1
glistN.Spinner = True
glistN.SetSpin MinValue, maxvalue, LargeValue
glistN.Enabled = True
glistN.EnabledBar = True
glistN.ShowBar = True

End If
glistN.Vertical = True
End Sub


Public Property Get Value() As Long
If glistN.Spinner Then
Value = glistN.ValueSilent
End If

End Property

Public Property Let Value(ByVal rhs As Long)
If glistN.Spinner Then
glistN.ValueSilent = rhs
End If

End Property



Public Property Get MaxCharLength() As Long
MaxCharLength = glistN.maxchar
End Property

Public Property Let MaxCharLength(ByVal rhs As Long)
glistN.maxchar = rhs
End Property


Private Function GetTextData(ByVal lFormatId As Long) As String
'' Addition to myTextBox for ctrl+V

Dim bData() As Byte, sr As String, sr1 As String
sr1 = Clipboard.GetText(1)
If (OpenClipboard(0) <> 0) Then

        
        If (GetBinaryData(lFormatId, bData())) Then
        sr = bData

            GetTextData = Left$(sr, Len(sr1))
          
        End If

End If
CloseClipboard
End Function
Private Function GetClipboardMemoryHandle( _
        ByVal lFormatId As Long _
    ) As Long

    

    If (IsClipboardFormatAvailable(lFormatId) <> 0) Then
         
        GetClipboardMemoryHandle = GetClipboardData(lFormatId)
        
    End If
End Function
Private Function GetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean

Dim hMem As Long, lSize As Long, lPtr As Long
    

    Erase bData
    
    hMem = GetClipboardMemoryHandle(lFormatId)

    If (hMem <> 0) Then

        lSize = GlobalSize(hMem)

        lPtr = GlobalLock(hMem)
        If (lSize > 0) Then

            ReDim bData(0 To lSize - 2) As Byte

            CopyMemory bData(0), ByVal lPtr, lSize - 1
        End If

        GlobalUnlock hMem

        GetBinaryData = (lSize > 0)

    End If
End Function

Private Sub glistN_UnregisterGlist()
If Not glistN.Spinner Then glistN.NoWheel = True
On Error Resume Next
If Typename(glistN.Parent) = "GuiM2000" Then glistN.NoWheel = True: Exit Sub
glistN.Parent.hookme Nothing
If Err.Number > 0 Then glistN.NoWheel = True

End Sub

Public Sub Move(x, y, Optional w, Optional h)
If IsMissing(w) Then
glistN.MoveTwips CLng(x), CLng(y), 0, 0
ElseIf IsMissing(h) Then
glistN.MoveTwips CLng(x), CLng(y), CLng(w), 0
Else
glistN.MoveTwips CLng(x), CLng(y), CLng(w), CLng(h)
End If
End Sub
Public Sub deconstruct()
glistN.Shutdown
Set Callback = Nothing
Set glistN = Nothing
End Sub
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set glistN.Font = aa
End If
End Sub

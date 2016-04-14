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
    (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" _
    () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Pad$
    ' for Gui purposes
Public CtrlName As String
Private mesCtrlName As String
Dim mIndex As Long, mStayOpen As Boolean
Dim Callback As GuiM2000
Dim white As Long, black As Long, dimm1 As Long, dimm2 As Long
Public Property Get GetCallBack() As Form
Set GetCallBack = Callback
End Property
Public Sub Construct(there As Form, name$)
 mIndex = -1
 Set Callback = there
 On Error Resume Next
 Set glistN = there.Controls.Add("m2000.glist", name$)
  CtrlName = name$
 If Err.Number > 0 Then
 Set glistN = there.Controls.Add("m2000.glist", GetTag$)
  CtrlName = GetTag$
 End If

 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(CtrlName).TabStop = True
 mesCtrlName = name$
End Sub
Public Sub ConstructArray(there As Form, name$, i As Long)
 Set Callback = there
 On Error Resume Next

  Set glistN = there.Controls.Add("m2000.glist", name$ + "_" + LTrim(Str$(i)))
CtrlName = name$ + "_" + LTrim(Str$(i))

 If Err.Number > 0 Then
 Set glistN = there.Controls.Add("m2000.glist", GetTag$)
 CtrlName = GetTag$
 End If
 mIndex = i
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(CtrlName).TabStop = True
 mesCtrlName = name$
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

Public Property Get top() As Variant
top = Callback.Controls(CtrlName).top
End Property
Public Property Let top(ByVal rhs As Variant)
Callback.Controls(CtrlName).top = rhs
End Property
Public Property Get Left() As Variant
Left = Callback.Controls(CtrlName).Left
End Property

Public Property Let Left(ByVal rhs As Variant)
Callback.Controls(CtrlName).Left = rhs
End Property


Public Sub SetUp()
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
Dim ff As New StdFont
With Callback
white = .Controls(1).ForeColor
ff.name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With

With glistN
.enabled = True
Info = ""
If .listcount = 0 Then Text = ""
.NoCaretShow = True
.VerticalCenterText = True
If Not .CenterText Then .LeftMarginPixels = 8 Else .LeftMarginPixels = 0
.BackColor = rgb(200, 120, 60)
Set .Font = ff
.ForeColor = white
.dcolor = Callback.Controls(1).dcolor
.ShowMe
End With
Locked = False
enabled = True
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


Private Sub Class_Terminate()
Set glistN = Nothing
End Sub

Private Sub glistN_RegisterGlist(this As gList)
If Not glistN.Spinner Then this.NoWheel = True
On Error Resume Next
Set LastGlist = this
If Err.Number > 0 Then this.NoWheel = True
End Sub



Private Sub glistN_ChangeListItem(item As Long, Content As String)
If Not enabled Then Exit Sub
Dim this As String, thispos As Long, VR(2)
thispos = glistN.SelStart - Len(mprompt)
this = Mid$(Content, 1 + Len(mprompt))
'' RaiseEvent ValidString(this, thispos)
VR(0) = this
VR(1) = thispos
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".ValidString(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".ValidString()", VR()
      End If
      
thispos = VR(1) + Len(mprompt)
Content = mprompt + VR(0)
If thispos <> glistN.SelStart And VR(0) <> "" Then glistN.SelStart = thispos: glistN.ShowMe
End Sub

Public Sub ResetPan()
glistN.PanPos = 0
End Sub

Private Sub glistN_ChangeSelStart(thisselstart As Long)
If thisselstart <= Len(mprompt) Then thisselstart = Len(mprompt) + 1
'glistN.ShowPan
With glistN
.ShowPan
.ExternalCursor .SelStart, .List(.ListIndex), .ForeColor
End With
End Sub

Public Property Get ThisKind() As String
ThisKind = mkind
End Property

Public Property Let ThisKind(ByVal rhs As String)
mkind = rhs
glistN.LastLinePart = " " & rhs
End Property


Public Property Get enabled() As Boolean
enabled = glistN.EditFlag
End Property

Public Property Let enabled(ByVal rhs As Boolean)
glistN.EditFlag = rhs
glistN.NoCaretShow = Not rhs
End Property
'


Private Sub glistN_CheckGotFocus()
If Not Locked Then
If glistN.ListIndex < 0 And glistN.listcount > 0 Then glistN.ListindexPrivateUse = 0
' RaiseEvent GotFocus
If enabled Then
glistN.ForeColor = black
glistN.BackColor = white
glistN.NoCaretShow = False:
glistN.PanPos = 0
Else
glistN.ForeColor = white
glistN.BackColor = dimm1
glistN.NoCaretShow = True
End If
glistN.ShowMe
End If

End Sub
Public Sub Retired()
If mStayOpen Then
glistN.ForeColor = black
glistN.BackColor = white
glistN.NoCaretShow = False
Else
glistN.ForeColor = white
glistN.BackColor = dimm2
glistN.NoCaretShow = True
End If
 glistN.PanPos = 0: glistN.ShowMe
End Sub

Private Sub glistN_CheckLostFocus()
If Not Locked Then
If mStayOpen Then
glistN.NoCaretShow = True

If glistN.Visible Then glistN.PanPos = 0: glistN.PrepareToShow
Exit Sub
End If
If Not SkipProcAtExitFocus Then
If glistN.BackStyle <> 1 Then
glistN.ForeColor = white
glistN.BackColor = dimm2
End If
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
Else
'
End If
End If
' RaiseEvent LostFocus
End Sub

Private Sub glistN_HeaderSelected(Button As Integer)
' INFO SELECTED
Text = Text
End Sub

Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)
Dim VR(2)
VR(0) = KeyCode
VR(1) = shift
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".KeyDown(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".KeyDown()", VR()
      End If
KeyCode = VR(0)
shift = VR(1)
If KeyCode = vbKeyUp Then
ElseIf KeyCode = vbKeyDown Then
ElseIf KeyCode = vbKeyV And (shift And 3) = 2 Then
If Not enabled Then Exit Sub
Pad$ = GetTextData(13)
If Pad$ = "" Then Pad$ = Clipboard.GetText(1)
If Pad$ <> "" Then vartext = Pad$
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
' RaiseEvent MouseMove(Button, shift, x, y)
End Sub

Private Sub glistN_Selected(item As Long)
If item >= 0 Then
   '' If (Not mCallback Is Nothing) Then mCallback.Press Index
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click()"
      End If
    End If
End Sub

Private Sub glistN_Selected2(item As Long)
If item >= 0 Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Enter(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Enter()"
      End If
End If
End Sub


Private Sub glistN_SyncKeyboard(item As Integer)
' RaiseEvent KeyPress(item)
End Sub
Private Sub glistN_SpinnerValue(ThatValue As Long)
''' RaiseEvent SpinnerValue(ThatValue)
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".SpinnerValue(" + Str(mIndex) + "," + LTrim(Str(ThatValue)) + ")"
   Else
      Callback.Callback mesCtrlName + ".SpinnerValue(" + LTrim(Str(ThatValue)) + ")"
      End If
     '' vartext = str(ThatValue)
End Sub
Public Property Get Locked() As Boolean
Locked = Not glistN.enabled
End Property

Public Property Let Locked(ByVal rhs As Boolean)
glistN.enabled = Not rhs
End Property
Public Sub Spinner(ShowBar As Boolean, MinValue As Long, maxvalue As Long, LargeValue As Long)
If Not ShowBar Then
glistN.Spinner = False
glistN.ShowBar = False
glistN.EnabledBar = False
Else
glistN.enabled = False
glistN.restrictLines = 1
glistN.Spinner = True
glistN.SetSpin MinValue, maxvalue, LargeValue
glistN.enabled = True
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

            GetTextData = Mid$(sr, 1, Len(sr1))
          
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
Set LastGlist = Nothing
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

If Not glistN Is Nothing Then glistN.Shutdown
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
Public Property Get Index() As Long
Index = mIndex
End Property

Public Property Let Index(ByVal rhs As Long)
mIndex = rhs
End Property
Public Property Let ShowAlways(ByVal vNewValue As Variant)
mStayOpen = CBool(vNewValue)
If mStayOpen Then
glistN.ForeColor = rgb(0, 0, 0)
glistN.BackColor = rgb(255, 255, 255)
End If
End Property
Sub GetFocus()
On Error Resume Next
Callback.Controls(CtrlName).SetFocus
End Sub
Sub Transparent()
Dim x  As Long, y As Long
glistN.ForeColor = black
glistN.BackStyle = 1
glistN.GetLeftTop x, y
glistN.RepaintFromOut Callback.image, x, y
End Sub
Sub ColorsReset()
white = rgb(255, 255, 255)
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
End Sub
Sub Colors(white1, Optional black1, Optional dimm11, Optional dimm21)

white = mycolor(white1)
If Not IsMissing(black1) Then
black = mycolor(black1)
If Not IsMissing(dimm11) Then
dimm1 = dimm11
If Not IsMissing(dimm21) Then dimm2 = dimm21
End If
End If
If enabled Then
glistN.ForeColor = white
Else
glistN.ForeColor = black
End If

'Retired
End Sub

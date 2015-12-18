VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dropdownlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents TextBox As myTextBox
Attribute TextBox.VB_VarHelpID = -1
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Public UseOnlyTheList As Boolean
' you have to make false the tab stop of glist
Event AutoCompleteDone(ByVal this$)
Event PickOther(ByVal this$)
Dim Iam As Boolean
Public AutoComplete As Boolean
Public NoCaseSensitive As Boolean
Public Property Set Container(glistNN As gList)
Set glistN = glistNN
glistNN.Enabled = True
glistNN.NoPanRight = True
glistNN.SingleLineSlide = True
glistNN.LeaveonChoose = True
glistNN.LeftMarginPixels = 8
glistNN.VerticalCenterText = True
glistNN.StickBar = True ' try with false - or hold shift to engage false
glistNN.NoFreeMoveUpDown = True
If Not TextBox Is Nothing Then
TextBox.Prompt = ChrW(&H25BC) + " "
End If
End Property
Public Property Get ListText()
ListText = glistN.ListValue
End Property
Public Property Let Label(A$)
TextBox.Prompt = A$ + " " + ChrW(&H25BC) & " "
TextBox.glistN.ShowMe
End Property
Private Sub glistN_RegisterGlist(this As gList)
Dim mo As Object
Set mo = glistN.Parent
If Not mo Is Nothing Then
On Error Resume Next
mo.hookme this
If Err.Number > 0 Then this.NoWheel = True
End If
Set mo = Nothing
End Sub




Private Sub Class_Terminate()
Set TextBox = Nothing
Set glistN = Nothing
End Sub

Private Sub glistN_CheckGotFocus()
TextBox.SkipProcAtExitFocus = True
If UseOnlyTheList Then If TextBox.Enabled Then TextBox.glistN.NoCaretShow = False
End Sub
Private Sub glistN_CheckLostFocus()

If TextBox.SkipProcAtExitFocus Then
TextBox.Retired
TextBox.SkipProcAtExitFocus = False
End If
Iam = True
End Sub

Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)

If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then

Iam = True
glistN.BypassLeaveonChoose = True
TextBox.SetFocus
End If
End Sub

Private Sub glistN_PanLeftRight(Direction As Boolean)
If Direction = 1 Then
Text = glistN.List(glistN.listindex)
End If
End Sub


Private Sub textbox_KeyDown(KeyCode As Integer, shift As Integer)
Dim b$
Iam = False
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
If glistN.Visible Then glistN.Enabled = True: TextBox.SkipProcAtExitFocus = True: glistN.SetFocus: Exit Sub
    KeyCode = 0
    glistN.Visible = True
    glistN.ZOrder
    TextBox.SkipProcAtExitFocus = True
    glistN.SetFocus
  If AutoComplete Then
  glistN.moveto Text
  If glistN.listindex < 0 Then
  glistN.listindex = 0
  Else
  glistN.ShowThis glistN.listindex + 1
  
  End If
  Else
        If glistN.LastSelected >= 0 Then
        glistN.listindex = glistN.LastSelected
    Else
        glistN.listindex = 0
    End If
End If

ElseIf KeyCode = vbKeyTab Then
TextBox.SkipProcAtExitFocus = False
glistN.Visible = False
ElseIf AutoComplete And KeyCode = vbKeyRight And TextBox.SelStart >= Len(TextBox.Text) Then
b$ = Left$(TextBox.glistN.LastLinePart, 1)
If b$ <> "" Then TextBox.glistN.LastLinePart = Mid$(TextBox.glistN.LastLinePart, 2)
TextBox.Text = TextBox.Text + b$
End If


End Sub


Private Sub TextBox_KeyPress(KeyAscii As Integer)
Iam = False
End Sub

Private Sub textbox_LostFocus()

If UseOnlyTheList Then
''If Iam And TextBox.Enabled Then TextBox.glistN.NoCaretShow = False
Else
TextBox.glistN.NoCaretShow = False
End If
TextBox.glistN.ShowMe
glistN.BypassLeaveonChoose = False
If Iam Then
glistN.Visible = False
Else
If Not glistN.Enabled Then glistN.Visible = False: glistN.Enabled = True
End If

End Sub
Private Sub textbox_GotFocus()
 TextBox.glistN.NoCaretShow = Not TextBox.Enabled
'' TextBox.glistN.BackColor = RGB(100, 200, 160): TextBox.glistN.ShowMe
If glistN.Visible Then
Iam = False
glistN.Enabled = False
glistN.BypassLeaveonChoose = False
End If
Dim DUMMY As Long
End Sub
Private Sub glistN_Selected2(item As Long)
If glistN.ListSep(item) Then Exit Sub
If item >= 0 Then
TextBox.vartext = glistN.List(item): TextBox.SelStart = Len(TextBox.Text) + 1
TextBox.SetFocus
ElseIf item < -1 Then
glistN.Visible = False
End If
End Sub


Private Sub textbox_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If TextBox.glistN.DblClick Then
Iam = False

 TextBox.glistN.ShowMe2

 End If
If Iam Then Exit Sub

If (Button And 1) = 1 And Not glistN.Visible Then
Iam = True
If UseOnlyTheList Then If TextBox.Enabled Then TextBox.glistN.NoCaretShow = False:    TextBox.glistN.ShowMe2
TextBox.Locked = True
glistN.Visible = True
glistN.ZOrder
glistN.LeaveonChoose = True
 If AutoComplete Then
  glistN.moveto Text
  If glistN.listindex < 0 Then
  glistN.listindex = 0
  Else
  glistN.ShowThis glistN.listindex + 1
  End If
  Else
       If glistN.LastSelected >= 0 Then
        glistN.listindex = glistN.LastSelected
    Else
        glistN.listindex = 0
    End If
    End If
    TextBox.SkipProcAtExitFocus = True
    glistN.SetFocus
TextBox.Locked = False
Iam = False
End If
End Sub
Private Sub glistN_SyncKeyboard(item As Integer)
Static f As Long
f = glistN.FindItemStartWidth(Chr$(item), Not NoCaseSensitive, f + 1)
If f >= 0 Then
glistN.ScrollTo f - glistN.lines / 2, f + 1
 RaiseEvent PickOther(glistN.ListValue)
Else
f = -1
End If
End Sub

Private Sub TextBox_ValidString(ThatString As String, setpos As Long)
Dim l As Boolean
If Not AutoComplete Then Exit Sub
Dim f As Long
Iam = True
If ThatString = "" And TextBox.glistN.LastLinePart <> "" Then
f = glistN.FindItemStartWidth(TextBox.glistN.LastLinePart, Not NoCaseSensitive, 0)

Else
f = glistN.FindItemStartWidth(ThatString, Not NoCaseSensitive, 0)
End If
If f >= 0 Then
If Not NoCaseSensitive Then ThatString = Left$(glistN.List(f), Len(ThatString))
TextBox.glistN.LastLinePart = Mid$(glistN.List(f), Len(ThatString) + 1)
RaiseEvent AutoCompleteDone(glistN.List(f))
glistN.Enabled = False
glistN.ScrollTo f - glistN.lines / 2, f + 1
glistN.Enabled = True
glistN.ShowMe

Else
If UseOnlyTheList Then
l = TextBox.glistN.Enabled
TextBox.glistN.Enabled = False
TextBox.glistN.LastLinePart = TextBox.vartext + TextBox.glistN.LastLinePart
ThatString = ""
setpos = 0
TextBox.glistN.Enabled = l
Else
TextBox.glistN.LastLinePart = ""
End If
End If
End Sub

Public Property Get Text() As String
Text = TextBox.vartext + TextBox.glistN.LastLinePart
End Property

Public Property Let Text(ByVal rhs As String)
 TextBox.glistN.LastLinePart = ""
TextBox.vartext = rhs

End Property
Public Property Get Locked() As Boolean
Locked = TextBox.Locked
End Property

Public Property Let Locked(ByVal rhs As Boolean)
If rhs = TextBox.Locked Then Exit Property
If rhs Then
If Not TextBox.Locked Then
glistN.BypassLeaveonChoose = True
TextBox.SetFocus
glistN.BypassLeaveonChoose = False
End If
End If
TextBox.Locked = rhs
End Property
Public Sub additem(A$)
glistN.additemFast A$
Text = A$
End Sub

Public Sub additemFast(A$)
glistN.additemFast A$
End Sub
Public Sub Shutdown()
'new
glistN.Shutdown
glistN.Enabled = False
TextBox.Locked = True

End Sub

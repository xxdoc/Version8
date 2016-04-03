VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Textbox As myTextBox
Attribute Textbox.VB_VarHelpID = -1
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Public WithEvents Pad As Form
Attribute Pad.VB_VarHelpID = -1
Private mEnabled As Boolean
Dim Iam As Boolean
'' aytocomplete isn't check yet for GuiDropDown...
'' next...time...
Public AutoComplete As Boolean
Public NoCaseSensitive As Boolean
' gui own
Public CtrlName As String
Dim mIndex As Long, mStayOpen As Boolean
Dim Callback As GuiM2000
Public Property Get GetCallBack() As Form
Set GetCallBack = Callback
End Property
 Public Sub Construct(there As Form, name$)
 mIndex = -1
 Set Callback = there
 Set Pad = New GuiM2000

 On Error Resume Next
 Set glistN = Pad.Controls(1)
 If Err.Number > 0 Then
 Set glistN = Pad.Controls(1)
 End If
 With Pad
.gList2.HeadLine = ""
.gList2.HeadLine = ""
.gList2.HeadlineHeight = .gList2.HeightPixels
End With
 Set Textbox = New myTextBox
 Set Textbox.Container = there.Controls.Add("m2000.glist", name$)
 CtrlName = name$
 If Err.Number > 0 Then
 Set Textbox.Container = there.Controls.Add("m2000.glist", GetTag$)
 CtrlName = GetTag$
 End If

 there.AddGuiControl Me
 glistN.Dynamic
 Textbox.glistN.Dynamic
 there.Controls(name$).TabStop = True
End Sub
Public Sub ConstructArray(there As Form, name$, i As Long)
 Set Callback = there
  Set Pad = New GuiM2000

 On Error Resume Next
 Set glistN = Pad.Controls(1)
 If Err.Number > 0 Then
 Set glistN = Pad.Controls(1)
 End If
 With Pad
.gList2.HeadLine = ""
.gList2.HeadLine = ""
.gList2.HeadlineHeight = .gList2.HeightPixels
End With
 Set Textbox = New myTextBox
 On Error Resume Next
 If i = 0 Then
 Set glistN = there.Controls.Add("m2000.glist", name$)
 CtrlName = name$
 Else
  Set glistN = there.Controls.Add("m2000.glist", name$ + "_" + LTrim(Str$(i)))
CtrlName = name$ + "_" + LTrim(Str$(i))
 End If
 If Err.Number > 0 Then
 Set glistN = there.Controls.Add("m2000.glist", GetTag$)
 CtrlName = GetTag$
 End If
 
 mIndex = i
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(name$).TabStop = True
End Sub
Public Sub Move(x, y, Optional w, Optional h)
With Textbox
If IsMissing(w) Then
.glistN.MoveTwips CLng(x), CLng(y), 0, 0
ElseIf IsMissing(h) Then
.glistN.MoveTwips CLng(x), CLng(y), CLng(w), 0
Else
.glistN.MoveTwips CLng(x), CLng(y), CLng(w), CLng(h)
End If
Pad.Move Pad.Left, Pad.top, .glistN.widthtwips, .glistN.HeightTwips * 5
End With
glistN.restrictLines = 5
glistN.MoveTwips 0, 0, Pad.Width, Pad.Height
End Sub
Public Sub deconstruct()
If Not glistN Is Nothing Then glistN.Shutdown
Set glistN = Nothing
If Not Pad Is Nothing Then Unload Pad
If Not Textbox Is Nothing Then Textbox.Dereference
Set Textbox = Nothing
Set Callback = Nothing
Set Pad = Nothing
End Sub
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set Textbox.glistN.Font = aa
End If
End Sub
Public Property Get Index() As Long
Index = mIndex
End Property

Public Property Set Container(glistNN As gList)
Set glistN = glistNN
glistNN.enabled = True
glistNN.NoPanRight = True
glistNN.SingleLineSlide = True
glistNN.LeaveonChoose = True
glistNN.LeftMarginPixels = 8
glistNN.VerticalCenterText = True
glistNN.StickBar = True ' try with false - or hold shift to engage false
glistNN.NoFreeMoveUpDown = True
If Not Textbox Is Nothing Then
Textbox.Prompt = ChrW(&H25BC) + " "
End If
End Property
Public Property Get ListText()
ListText = glistN.ListValue
End Property
Public Property Let Label(a$)
Textbox.Prompt = a$ + " " + ChrW(&H25BC) & " "
Textbox.glistN.ShowMe
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
Set Textbox = Nothing
Set glistN = Nothing
If Not Pad Is Nothing Then Unload Pad
Set Pad = Nothing
End Sub

Private Sub glistN_CheckGotFocus()
Textbox.SkipProcAtExitFocus = True

Textbox.glistN.NoCaretShow = True
glistN.ForeColor = rgb(0, 0, 0)
glistN.BackColor = rgb(255, 255, 255)
glistN.NoCaretShow = False:
glistN.PanPos = 0
Textbox.glistN.dropkey = False
End Sub
Private Sub glistN_CheckLostFocus()

If Textbox.SkipProcAtExitFocus Then
Textbox.Retired
Textbox.SkipProcAtExitFocus = False
Else
Pad.Visible = False
End If
Iam = True

End Sub

Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)

If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then

Iam = True
glistN.BypassLeaveonChoose = True
Textbox.SetFocus
End If
End Sub

Private Sub glistN_PanLeftRight(Direction As Boolean)
If Direction = True Then
Text = glistN.List(glistN.ListIndex)
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".DblClick(" + Str(mIndex) + "," + Str(glistN.ListIndex) + ")"
   Else
      Callback.Callback CtrlName + ".DblClick(" + Str(glistN.ListIndex) + ")"
      End If
   Textbox.SetFocus
End If
End Sub
Private Sub glistN_ScrollMove(item As Long)
'
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Move(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback CtrlName + ".Move(" + Str(item) + ")"
      End If
End Sub

Private Sub glistN_ScrollSelected(item As Long, y As Long)
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Scroll(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback CtrlName + ".Scroll(" + Str(item) + ")"
      End If
End Sub


Public Property Get List(Index) As Variant
On Error Resume Next
List = glistN.List(CLng(Index))
End Property
Private Sub glistN_Selected(item As Long)
If glistN.PanPos <> 0 Then glistN.PanPos = 0
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Click(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback CtrlName + ".Click(" + Str(item) + ")"
      End If
End Sub

Private Sub textbox_KeyDown(KeyCode As Integer, shift As Integer)
Dim b$
Dim gl As Long
Iam = False
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
If glistN.Visible Then glistN.enabled = True: Textbox.SkipProcAtExitFocus = True: glistN.SetFocus: Exit Sub
    KeyCode = 0
    gl = glistN.listcount
     If gl = 0 Then Exit Sub
     
Pad.Width = CLng(Callback.Controls(CtrlName).Width / 1.618 * dv15) \ dv15
If gl < 5 Then

Pad.Height = Callback.Controls(CtrlName).Height * gl
glistN.restrictLines = gl

Else
Pad.Height = Callback.Controls(CtrlName).Height * 5
glistN.restrictLines = 5
End If
glistN.MoveTwips 0, 0, Pad.Width, Pad.Height
    Textbox.SkipProcAtExitFocus = True
    Textbox.glistN.NoCaretShow = True
    Textbox.glistN.ShowPan
    glistN.LeaveonChoose = True
   ' glistN.SetFocus
  If AutoComplete Then
  glistN.moveto Text
  If glistN.ListIndex < 0 Then
  glistN.ListIndex = 0
  Else
  glistN.ShowPan ' glistN.ListIndex + 1
  
  End If
  Else
        If glistN.LastSelected >= 0 Then
        glistN.ListIndex = glistN.LastSelected
    Else
    glistN.moveto Text
        If glistN.ListIndex < 0 Then
  glistN.ListIndex = 0
  Else
  glistN.ShowPan ' glistN.ListIndex + 1
  
  End If
    End If
Callback.PopUpPos Pad, Left + Callback.Controls(CtrlName).Width - Pad.Width, top + Callback.Controls(CtrlName).Height



End If

ElseIf KeyCode = vbKeyTab Then
Textbox.SkipProcAtExitFocus = False
Pad.Visible = False
'glistN.Visible = False
ElseIf AutoComplete And KeyCode = vbKeyRight And Textbox.SelStart >= Len(Textbox.Text) Then
b$ = Myleft$(Textbox.glistN.LastLinePart, 1)
If b$ <> "" Then Textbox.glistN.LastLinePart = Mid$(Textbox.glistN.LastLinePart, 2)
Textbox.Text = Textbox.Text + b$
End If


End Sub
Function Myleft$(a As Variant, pos As Variant)
Myleft$ = Mid$(a, 1, pos)
End Function

Private Sub TextBox_KeyPress(KeyAscii As Integer)
Iam = False
End Sub

Private Sub textbox_LostFocus()


Textbox.glistN.NoCaretShow = True


Textbox.glistN.ShowMe
glistN.BypassLeaveonChoose = False
If Iam Then
Pad.Visible = False
'glistN.Visible = False
Else
If Not glistN.enabled Then
'glistN.Visible = False:
glistN.enabled = True
End If
End If

End Sub
Private Sub textbox_GotFocus()
 Textbox.glistN.NoCaretShow = Not Textbox.enabled
 Textbox.glistN.ShowMe
If Pad.Visible Then
Iam = False
glistN.enabled = False
glistN.BypassLeaveonChoose = False
End If
Dim dummy As Long
End Sub
Private Sub glistN_Selected2(item As Long)
If glistN.ListSep(item) Then Exit Sub
If item >= 0 Then
Textbox.vartext = glistN.List(item): Textbox.SelStart = Len(Textbox.Text) + 1

Textbox.SetFocus

   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".DblClick(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback CtrlName + ".DblClick(" + Str(item) + ")"
      End If
ElseIf item < -1 Then
'glistN.Visible = False
Pad.Visible = False
End If
End Sub


Private Sub textbox_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim gl As Long
If Textbox.glistN.DblClick Then
Iam = False

 Textbox.glistN.ShowMe2

 End If

If (Button And 1) = 1 And Not Pad.Visible Then
gl = glistN.listcount
If gl = 0 Then Exit Sub

Textbox.glistN.dropkey = True
If glistN.Visible Then glistN.enabled = True: Textbox.SkipProcAtExitFocus = True: glistN.SetFocus: Exit Sub
Pad.Width = CLng(Callback.Controls(CtrlName).Width / 1.618 * dv15) \ dv15
If gl < 5 Then

Pad.Height = Callback.Controls(CtrlName).Height * gl
glistN.restrictLines = gl

Else
Pad.Height = Callback.Controls(CtrlName).Height * 5
glistN.restrictLines = 5
End If
glistN.MoveTwips 0, 0, Pad.Width, Pad.Height
    Textbox.SkipProcAtExitFocus = True
    Textbox.glistN.NoCaretShow = True
    Textbox.glistN.ShowPan
    glistN.LeaveonChoose = True
   ' glistN.SetFocus
  If AutoComplete Then
        glistN.moveto Text
        If glistN.ListIndex < 0 Then
        glistN.ListIndex = 0
        Else
        glistN.ShowPan ' glistN.ListIndex + 1
        
         End If
  Else
        If glistN.LastSelected >= 0 Then
             glistN.ListIndex = glistN.LastSelected
    Else
            glistN.moveto Text
                If glistN.ListIndex < 0 Then
                glistN.ListIndex = 0
                Else
                glistN.ShowPan ' glistN.ListIndex + 1
                
                End If
    End If
    End If
Callback.PopUpPos Pad, Left + Callback.Controls(CtrlName).Width - Pad.Width, top + Callback.Controls(CtrlName).Height


Button = 0
Textbox.glistN.dropkey = False
'Iam = False
End If
End Sub
Private Sub glistN_SyncKeyboard(item As Integer)
Static f As Long
f = glistN.FindItemStartWidth(Chr$(item), Not NoCaseSensitive, f + 1)
If f >= 0 Then
glistN.ScrollTo f - glistN.lines / 2, f + 1
    If mIndex >= 0 Then
   Callback.Callback CtrlName + ".PickOther(" + Str(mIndex) + ",{" + glistN.ListValue + "})"
   Else
      Callback.Callback CtrlName + ".PickOther({" + glistN.ListValue + "})"
      End If
Else
f = -1
End If
End Sub

Private Sub TextBox_ValidString(ThatString As String, setpos As Long)
Dim l As Boolean
If Not AutoComplete Then Exit Sub
Dim f As Long
Iam = True
If ThatString = "" And Textbox.glistN.LastLinePart <> "" Then
f = glistN.FindItemStartWidth(Textbox.glistN.LastLinePart, Not NoCaseSensitive, 0)

Else
f = glistN.FindItemStartWidth(ThatString, Not NoCaseSensitive, 0)
End If
If f >= 0 Then
If Not NoCaseSensitive Then ThatString = Myleft$(glistN.List(f), Len(ThatString))
Textbox.glistN.LastLinePart = Mid$(glistN.List(f), Len(ThatString) + 1)
'RaiseEvent AutoCompleteDone(glistN.List(f))
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".AutoCompleteDone(" + Str(mIndex) + ",{" + glistN.List(f) + "})"
   Else
      Callback.Callback CtrlName + ".AutoCompleteDone({" + glistN.List(f) + "})"
      End If

glistN.enabled = False
glistN.ScrollTo f - glistN.lines / 2, f + 1
glistN.enabled = True
glistN.ShowMe

Else
'If UseOnlyTheList Then

If False Then
l = Textbox.glistN.enabled
Textbox.glistN.enabled = False
Textbox.glistN.LastLinePart = Textbox.vartext + Textbox.glistN.LastLinePart
ThatString = ""
setpos = 0
Textbox.glistN.enabled = l
Else
Textbox.glistN.LastLinePart = ""
End If
End If
End Sub

Public Property Get Text() As String
Text = Textbox.vartext + Textbox.glistN.LastLinePart
End Property

Public Property Let Text(ByVal rhs As String)
 Textbox.glistN.LastLinePart = ""
Textbox.vartext = rhs

End Property
Public Property Get Locked() As Boolean
Locked = Textbox.Locked
End Property

Public Property Let Locked(ByVal rhs As Boolean)
If rhs = Textbox.Locked Then Exit Property
If rhs Then
If Not Textbox.Locked Then
glistN.BypassLeaveonChoose = True
Textbox.SetFocus
glistN.BypassLeaveonChoose = False
End If
End If
Textbox.Locked = rhs
End Property
Public Sub additem(a$)
glistN.additemFast a$
Text = a$
End Sub

Public Sub additemFast(a$)
glistN.additemFast a$
End Sub
Public Sub Shutdown()
'new
glistN.Shutdown
glistN.enabled = False
Textbox.Locked = True

End Sub

Public Sub SetUp()
white = rgb(255, 255, 255)
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)

Dim ff As New StdFont
Dim ff1 As New StdFont
With Callback
ff.name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
ff1.name = .CtrlFontName
ff1.charset = Form1.DIS.Font.charset
ff1.Size = .CtrlFontSize
ff1.bold = .CtrlFontBold


End With
Set Textbox.glistN.Font = ff1
With glistN
.SingleLineSlide = True
.NoPanLeft = True
.NoPanRight = False
.Vertical = True
.VerticalCenterText = True
.Text = mListText
.BackColor = dimm2
.MoveTwips 0, 0, Pad.Width, Pad.Height

Set .Font = ff
.ForeColor = Callback.Controls(1).ForeColor
.dcolor = Callback.Controls(1).dcolor
.NoCaretShow = False
.enabled = True
.StickBar = True
enabled = True
.ShowMe
End With

End Sub

Public Property Let ListText(ByVal vNewValue As Variant)
mListText = vNewValue
glistN.Text = mListText
End Property
Public Property Get enabled() As Boolean
enabled = Textbox.glistN.enabled
End Property

Public Property Let enabled(ByVal rhs As Boolean)
'glistN.HideCaretOnexit = Not rhs
'If Not rhs Then
'glistN.Noflashingcaret = False
'EditDoc = EditDoc
'Show

'End If


Textbox.glistN.enabled = rhs
End Property
Public Property Get Left() As Variant
Left = Callback.Controls(CtrlName).Left
End Property

Public Property Let Left(ByVal rhs As Variant)
Callback.Controls(CtrlName).Left = rhs
End Property
Public Property Get top() As Variant
top = Callback.Controls(CtrlName).top
End Property
Public Property Let top(ByVal rhs As Variant)
Callback.Controls(CtrlName).top = rhs
End Property
Public Sub Sort(Optional SkipChar As Long = 1)
If glistN.listcount < 2 Then Exit Sub
Dim doc As Document, m As Long
Set doc = New Document
doc.textDocFast = glistN.Text
m = glistN.ListIndex
If SkipChar < 1 Then SkipChar = 1
doc.SortDoc SkipChar, 1, glistN.listcount
glistN.Text = doc.textDoc
If m > 0 Then glistN.ListindexPrivateUse = m
glistN.ShowMe2
Set doc = Nothing
End Sub
Public Property Get Find(ByVal FindThis$) As Variant
    Find = glistN.Find(FindThis)
End Property
Public Property Get ListIndex() As Variant

ListIndex = glistN.ListIndex
End Property

Public Property Let ListIndex(ByVal vNewValue As Variant)
glistN.ListIndex = CLng(vNewValue)
glistN.ShowPan

End Property
Public Property Let ListIndex2(ByVal vNewValue As Variant)
glistN.ListindexPrivateUse = CLng(vNewValue)
End Property
Public Property Get listcount() As Variant
listcount = glistN.listcount
End Property




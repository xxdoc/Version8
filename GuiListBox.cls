VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Private mEnabled As Boolean
Public CtrlName As String
Dim mListText As String
Dim mIndex As Long, mStayOpen As Boolean
Dim Callback As GuiM2000, Point2Me As Object
' colors
Dim white As Long, black As Long, dimm1 As Long, dimm2 As Long, trans As Boolean, noeditme As Boolean

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
Public Sub deconstruct()
Set Point2Me = Nothing
glistN.Shutdown
Set Callback = Nothing
Set glistN = Nothing
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

Private Sub Class_Terminate()
Set Point2Me = Nothing
End Sub

Private Sub glistN_BlockCaret(item As Long, blockme As Boolean, skipme As Boolean)
If trans And noeditme Then blockme = True  ': glistN.NoCaretShow = True
End Sub

Private Sub glistN_ChangeListItem(item As Long, Content As String)
If Not mEnabled Then Exit Sub
Dim thispos As Long, vr(3)
thispos = glistN.SelStart
vr(0) = Content
vr(1) = thispos
vr(2) = item
   If mIndex >= 0 Then
   Callback.CallbackNow CtrlName + ".ValidList(" + Str(mIndex) + ")", vr()
   Else
      Callback.CallbackNow CtrlName + ".ValidList()", vr()
      End If
      
thispos = vr(1)
Content = vr(0)
If thispos <> glistN.SelStart And Content <> "" Then glistN.SelStart = thispos: glistN.ShowMe
End Sub

Private Sub glistN_CheckGotFocus()
If Not locked Then
If glistN.ListIndex < 0 And glistN.listcount > 0 Then glistN.ListindexPrivateUse = 0
If Not trans Then
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
End If
glistN.ShowMe
End If

End Sub
Public Sub Retired()
If Not trans Then
glistN.ForeColor = white
glistN.BackColor = dimm2
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
End If
End Sub

Private Sub glistN_CheckLostFocus()
If Not locked Then
If mStayOpen Then
If glistN.Visible Then glistN.PanPos = 0: glistN.PrepareToShow
Exit Sub
End If
If Not trans Then
glistN.ForeColor = white
glistN.BackColor = dimm2
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.PrepareToShow
End If
End If

End Sub
Public Sub SetFocus()
On Error Resume Next
glistN.SetFocus
End Sub
Public Sub SetUp()
white = rgb(255, 255, 255)
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)

Dim ff As New StdFont
With Callback
ff.name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With

With glistN
.SingleLineSlide = True
.NoPanLeft = True
.NoPanRight = False
.Vertical = True
.VerticalCenterText = True
.Text = mListText
.BackColor = dimm2

Set .Font = ff
.ForeColor = Callback.Controls(0).ForeColor
.dcolor = Callback.Controls(0).dcolor
.NoCaretShow = False
.enabled = True
.StickBar = True
enabled = True
.ShowMe
End With
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
Public Property Get locked() As Boolean
locked = Not glistN.enabled
End Property

Public Property Let locked(ByVal rhs As Boolean)
glistN.enabled = Not rhs
End Property

Private Sub glistN_ExposeListcount(cListCount As Long)
' use with readlistitem

If Not mEnabled Then Exit Sub
Dim vr(1)

vr(0) = cListCount

   If mIndex >= 0 Then
   Callback.CallbackNow CtrlName + ".ListCount(" + Str(mIndex) + ")", vr()
   Else
      Callback.CallbackNow CtrlName + ".ListCount()", vr()
      End If
cListCount = vr(0)
End Sub

Private Sub glistN_GetBackPicture(pic As Object)
Set pic = Point2Me
End Sub

Private Sub glistN_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vr(2)
vr(0) = KeyCode
vr(1) = Shift
   If mIndex >= 0 Then
   Callback.CallbackNow CtrlName + ".KeyDown(" + Str(mIndex) + ")", vr()
   Else
      Callback.CallbackNow CtrlName + ".KeyDown()", vr()
      End If
KeyCode = vr(0)
Shift = vr(1)
End Sub

Private Sub glistN_PanLeftRight(Direction As Boolean)
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".DblClick(" + Str(mIndex) + "," + Str(glistN.ListIndex) + ")"
   Else
      Callback.Callback CtrlName + ".DblClick(" + Str(glistN.ListIndex) + ")"
      End If
End Sub

Private Sub glistN_ReadListItem(item As Long, Content As String)
If Not mEnabled Then Exit Sub
Dim vr(1)


vr(0) = Content
   If mIndex >= 0 Then
   Callback.CallbackNow CtrlName + ".External(" + Str(mIndex) + "," + Str(item) + ")", vr()
   Else
      Callback.CallbackNow CtrlName + ".External(" + Str(item) + ")", vr()
      End If
    
Content = vr(0)


End Sub

Private Sub glistN_RefreshDesktop()

If Callback.Visible Then Callback.Refresh

End Sub

Private Sub glistN_RemoveOne(that As String)
'' this is the unicode key
Dim vr(1)


vr(0) = that
   If mIndex >= 0 Then
   Callback.CallbackNow CtrlName + ".Unicode(" + Str(mIndex) + ")", vr()
   Else
      Callback.CallbackNow CtrlName + ".Unicode()", vr()
      End If
    
that = vr(0)

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

Private Sub glistN_Selected(item As Long)
'
If glistN.PanPos <> 0 Then glistN.PanPos = 0
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Click(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback CtrlName + ".Click(" + Str(item) + ")"
      End If
End Sub

Private Sub glistN_Selected2(item As Long)
'
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".DblClick(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback CtrlName + ".DblClick(" + Str(item) + ")"
      End If
End Sub

Private Sub gListN_SpecialColor(rgbcolor As Long)

If Not mEnabled Then Exit Sub
Dim vr(1)

vr(0) = -rgbcolor

   If mIndex >= 0 Then
   Callback.CallbackNow CtrlName + ".Color(" + Str(mIndex) + ")", vr()
   Else
      Callback.CallbackNow CtrlName + ".Color(", vr()
      End If
rgbcolor = mycolor(vr(0))
End Sub
Function Control() As gList
    Set Control = glistN
End Function

Public Property Get ListText() As Variant
ListText = mListText
End Property

Public Property Let ListText(ByVal vNewValue As Variant)
mListText = vNewValue

End Property
Public Property Get enabled() As Boolean
enabled = mEnabled
End Property

Public Property Let enabled(ByVal rhs As Boolean)
mEnabled = rhs
End Property

Public Property Get Linespace() As Variant
Linespace = dv15 * (glistN.addpixels) / 2
End Property

Public Property Let Linespace(ByVal vNewValue As Variant)
    vNewValue = (vNewValue \ 30) * 60
 
glistN.addpixels = vNewValue / dv15
If glistN.Visible Then glistN.ShowMe
End Property
Public Sub Refresh()
glistN.CalcAndShowBar
If glistN.Visible Then glistN.ShowMe2
End Sub

Public Property Get List(Index) As Variant
On Error Resume Next
List = glistN.List(CLng(Index))
End Property

Public Property Let List(Index, ByVal vNewValue As Variant)
On Error Resume Next
glistN.List(CLng(Index)) = CStr(vNewValue)
glistN.PrepareToShow
End Property
Sub Add(Content$)
glistN.additem Content$
End Sub
Sub AddFast(Index As Long, Content$)
glistN.additemFast Content$
End Sub
Sub Insert(Content$)
glistN.additemAtListIndex Content$
End Sub
Sub Clear()
glistN.Clear
If glistN.Visible Then glistN.ShowMe2
End Sub

Sub Delete(Index As Long)
glistN.Removeitem Index
End Sub

Public Property Get ListIndex() As Variant
ListIndex = glistN.ListIndex
End Property

Public Property Let ListIndex(ByVal vNewValue As Variant)
glistN.ListIndex = CLng(vNewValue)
glistN.ShowThis CLng(vNewValue) + 1
End Property
Public Property Let ListIndex2(ByVal vNewValue As Variant)
glistN.ListindexPrivateUse = CLng(vNewValue)
End Property
Public Property Get listcount() As Variant
listcount = glistN.listcount
End Property
Public Sub ShowThis(Index As Long, Optional noselect)
glistN.ShowThis Index, CBool(noselect)
End Sub

Public Property Get Text() As Variant
Text = glistN.Text
End Property

Public Property Let Text(ByVal vNewValue As Variant)
glistN.Text = CStr(vNewValue)
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


Public Property Get External() As Variant
External = glistN.BlockItemcount
End Property

Public Property Let External(ByVal vNewValue As Variant)
glistN.BlockItemcount = CBool(vNewValue)
End Property
Public Property Let DisplayLines(ByVal vNewValue As Variant)
glistN.restrictLines = CLng(vNewValue)
End Property
Sub ClorosReset()
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
Retired
End Sub
Property Let Transparent(ok As Boolean)
trans = ok
If ok Then
glistN.NoCaretShow = True
Set Point2Me = Callback.Picture
glistN.SkipForm = False
glistN.BackStyle = 1
Else
glistN.NoCaretShow = False
Set Point2Me = Nothing
glistN.SkipForm = True
glistN.BackStyle = 0

End If

End Property
Property Let Center(ok As Boolean)
    glistN.CenterText = ok
    glistN.ShowMe2
End Property
Property Let Edit(ok As Boolean)
    glistN.EditFlag = ok
    glistN.NoCaretShow = Not ok
    noeditme = Not ok
    glistN.ShowMe2
End Property
Private Sub glistN_RegisterGlist(this As gList)
On Error Resume Next
Set LastGlist = this
If Err.Number > 0 Then this.NoWheel = True
End Sub

Private Sub glistN_SyncKeyboard(item As Integer)
Dim vr(1)
vr(0) = item
   If mIndex >= 0 Then
   Callback.CallbackNow CtrlName + ".Ascii(" + Str(mIndex) + ")", vr()
   Else
      Callback.CallbackNow CtrlName + ".Ascii()", vr()
      End If
    
item = vr(0)
End Sub

Private Sub glistN_UnregisterGlist()
On Error Resume Next
Set LastGlist = Nothing
If Err.Number > 0 Then glistN.NoWheel = True
End Sub
Property Let ScrollTo(there As Long)
glistN.ScrollTo there
End Property
Property Get ScrollFrom()
ScrollFrom = glistN.ScrollFrom
End Property
Property Let ShowBar(ok As Boolean)
glistN.ShowBar = ok
If ok Then
End If
End Property

Public Property Let ShowAlways(ByVal vNewValue As Variant)
mStayOpen = CBool(vNewValue)
End Property
Public Property Get Find(ByVal FindThis$) As Long
    Find = glistN.Find(FindThis)
End Property


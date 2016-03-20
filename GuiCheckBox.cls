VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' From Tools make changed property as default

Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Dim check1focus As Boolean
Dim mCaptext As String
Dim mValue As Boolean
'' FOR EVERY GuiSomething
Public CtrlName As String
Dim mIndex As Long
Dim Callback As GuiM2000
'' for every GuiSomething
Public Property Get Index() As Long
Index = mIndex
End Property

Public Property Let Index(ByVal rhs As Long)
mIndex = rhs
End Property
Public Sub Construct(there As Form, name$)
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
Public Sub ConstructContainer(there As Form, name$, containerCtrl As Object)
Set Callback = there
 Set glistN = there.Controls.Add("m2000.glist", name$, containerCtrl)
there.AddGuiControl Me
CtrlName = nam$
 glistN.Dynamic
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
Shutdown
Set Callback = Nothing
  Set glistN = Nothing
End Sub
Public Sub SetUp()
Dim ff As New StdFont
With Callback
ff.name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With
With glistN
glistN.clear
.additem "  " + mCaptext
.NoCaretShow = True
.NoPanRight = False
.VerticalCenterText = True
.restrictLines = 1
.MenuItem 1, True, False, mValue
.listindex = 0
Set .Font = ff
.ForeColor = Callback.Controls(0).ForeColor
.dcolor = Callback.Controls(0).dcolor
.LeftMarginPixels = .HeightPixels
.ShowMe
.Enabled = True
End With
End Sub
Public Property Let Caption(myCap As String)
'
mCaptext = myCap
If Not glistN Is Nothing Then
glistN.List(0) = "  " + mCaptext
glistN.ShowMe
End If
End Property
Public Property Let CheckReset(rhs As Boolean)
' without event
mValue = rhs
If Not glistN Is Nothing Then
glistN.ListSelected(0) = mValue
glistN.ShowMe
End If
End Property



Private Sub Class_Terminate()
Set glistN = Nothing
End Sub

Private Sub glistN_CheckGotFocus()

check1focus = True
End Sub

Private Sub gListN_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
'
If check1focus Then
glistN.FillThere thisHDC, thisrect, rgb(100, 200, 160)
Else
glistN.FillThere thisHDC, thisrect, rgb(200, 120, 60)
End If
skip = False
End Sub


Private Sub glistN_CheckLostFocus()
check1focus = False
glistN.ShowMe
End Sub
Private Sub gListN_MenuChecked(item As Long)
' always item is 1  so 1-1..0 (base 0)
' RaiseEvent Changed(glistN.ListSelected(0))
If item >= 0 Then
  If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Changed(" + CStr(mIndex) + "," + CStr(glistN.ListSelected(0)) + ")"
   Else
      Callback.Callback CtrlName + ".Changed(" + CStr(glistN.ListSelected(0)) + ")"
      End If
End If
End Sub
Public Property Get Checked() As Boolean
If Not glistN Is Nothing Then
Checked = glistN.ListSelected(0)
Else
' error
End If
End Property

Sub Shutdown()
glistN.Shutdown
glistN.Enabled = False
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
Private Sub glistN_OnResize()
glistN.LeftMarginPixels = glistN.HeightPixels
End Sub

Private Sub glistN_PanLeftRight(Direction As Boolean)
  If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Click(" + CStr(mIndex) + "," + CStr(Checked) + ")"
   Else
      Callback.Callback CtrlName + ".Click(" + CStr(Checked) + ")"
      End If
End Sub

Private Sub glistN_Selected2(item As Long)
  If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Click(" + CStr(mIndex) + "," + CStr(Checked) + ")"
   Else
      Callback.Callback CtrlName + ".Click(" + CStr(Checked) + ")"
      End If
End Sub
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

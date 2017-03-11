VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Dim mCaptext As String
Public CtrlName As String
Private mesCtrlName As String, mlocked As Boolean
Dim mIndex As Long
Dim Callback As GuiM2000
Dim white As Long, dimm1 As Long, dimm2 As Long
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
Public Sub ConstructContainer121(there As Form, name$, containerCtrl As Object)
Set Callback = there
 Set glistN = there.Controls.Add("m2000.glist", name$, containerCtrl)
there.AddGuiControl Me
CtrlName = name$
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
white = Callback.Controls(1).ForeColor
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
.NoCaretShow = True
.NoPanRight = False
.restrictLines = 1
.CenterText = True
.VerticalCenterText = True
.Text = mCaptext
.BackColor = rgb(200, 120, 60)
Set .Font = ff
.ForeColor = Callback.Controls(1).ForeColor
.dcolor = Callback.Controls(1).dcolor

.enabled = True
.ShowMe
End With
End Sub

Private Sub Class_Initialize()
mCaptext = "Button"
mIndex = -1
End Sub

Private Sub Class_Terminate()
Set glistN = Nothing
End Sub

Private Sub glistN_BlinkNow(Face As Boolean)
'
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Blink(" + Str(mIndex) + "," + Str(Face) + ")"
   Else
      Callback.Callback mesCtrlName + ".Blink(" + Str(Face) + ")"
      End If
End Sub

Private Sub glistN_CheckGotFocus()
If mlocked Then Exit Sub
glistN.BackColor = dimm1
glistN.ShowMe
End Sub

Private Sub glistN_CheckLostFocus()

If Me.enabled Then
If Not mlocked Then glistN.BackColor = dimm2
glistN.PanPos = 0
glistN.ShowMe

End If
End Sub

Private Sub glistN_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
glistN.ShowMe
End Sub

Private Sub glistN_PanLeftRight(Direction As Boolean)

 
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click()"
      End If
    
End Sub

Private Sub glistN_RegisterGlist(this As gList)
this.NoWheel = True
End Sub

Private Sub glistN_Selected(item As Long)
If mlocked Then
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
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click()"
      End If
    End If
End Sub
'
'
Public Property Let Caption(myCap As String)
'
mCaptext = myCap
If Not glistN Is Nothing Then
If glistN.CenterText Then
glistN.List(0) = mCaptext
Else
glistN.List(0) = "  " + mCaptext
End If
glistN.ShowMe
End If
End Property
Public Property Get Caption() As String
Caption = mCaptext
End Property

Public Property Get Index() As Long
Index = mIndex
End Property

Public Property Let Index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Property Get enabled() As Long
enabled = glistN.enabled
End Property

Public Property Let enabled(ByVal RHS As Long)
If glistN Is Nothing Then Exit Property
glistN.enabled = RHS
End Property
Sub ColorsReset()
white = rgb(255, 255, 255)
glistN.ForeColor = rgb(255, 255, 255)
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
End Sub
Sub Colors(white1, Optional dimm11, Optional dimm21)

glistN.ForeColor = mycolor(white1)

If Not IsMissing(dimm11) Then
dimm1 = dimm11
If Not IsMissing(dimm21) Then dimm2 = dimm21
End If
'Retired
End Sub
Public Sub Shutdown()
On Error Resume Next
If Not glistN Is Nothing Then
glistN.Shutdown
glistN.enabled = False
End If
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
Public Sub Curve(RHS As Single)
glistN.Curve False, RHS
End Sub
Public Property Get top() As Variant
top = Callback.Controls(CtrlName).top
End Property
Public Property Let top(ByVal RHS As Variant)
Callback.Controls(CtrlName).top = RHS
End Property
Public Property Get Left() As Variant
Left = Callback.Controls(CtrlName).Left
End Property

Public Property Let Left(ByVal RHS As Variant)
Callback.Controls(CtrlName).Left = RHS
End Property
Sub GetFocus()
On Error Resume Next
Callback.Controls(CtrlName).SetFocus
End Sub


Public Property Get Locked() As Variant
Locked = mlocked
End Property

Public Property Let Locked(ByVal vNewValue As Variant)
mlocked = vNewValue
glistN.dropkey = mlocked
glistN.TabStop = Not mlocked
glistN.BackColor = mycolor(dimm1)
glistN.PanPos = 0
glistN.ShowMe
End Property
Sub Show()
glistN.ShowPan
End Sub
Public Property Let Blink(ByVal vNewValue As Variant)
glistN.BlinkTime = vNewValue
End Property
Sub Transparent()
Dim x  As Long, y As Long
glistN.BackStyle = 1
glistN.GetLeftTop x, y
glistN.RepaintFromOut Callback.image, x, y
End Sub

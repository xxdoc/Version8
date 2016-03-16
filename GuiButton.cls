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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Dim mCaptext As String
Public CtrlName As String
Dim mIndex As Long
Dim Callback As GuiM2000
Public Property Get GetCallBack() As Form
Set GetCallBack = Callback
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
.NoCaretShow = True
.NoPanRight = False
.restrictLines = 1
.CenterText = True
.VerticalCenterText = True
.Text = mCaptext
.BackColor = rgb(200, 120, 60)
Set .Font = ff
.ForeColor = Callback.Controls(0).ForeColor
.dcolor = Callback.Controls(0).dcolor

.Enabled = True
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

Private Sub glistN_CheckGotFocus()
''Callback.Callback CtrlName + ".GotFocus()"
If mModalId <> ModalId And ModalId <> 0 Then Exit Sub
glistN.BackColor = rgb(100, 200, 160)
glistN.ShowMe
End Sub

Private Sub glistN_CheckLostFocus()
If Me.Enabled Then
''Callback.Callback CtrlName + ".LostFocus()"
glistN.BackColor = rgb(200, 120, 60)
glistN.PanPos = 0
glistN.ShowMe

End If
End Sub

Private Sub glistN_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
glistN.ShowMe
End Sub

Private Sub glistN_PanLeftRight(Direction As Boolean)

 
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Click(" + CStr(mIndex) + ")"
   Else
      Callback.Callback CtrlName + ".Click()"
      End If
    
''    Callback.Callback CtrlName + ".PanLeftRight(" + CStr(Direction) + ")"
End Sub

Private Sub glistN_RegisterGlist(this As gList)
this.NoWheel = True
End Sub

Private Sub glistN_Selected2(item As Long)
If item >= 0 Then
   '' If (Not mCallback Is Nothing) Then mCallback.Press Index
   If mIndex >= 0 Then
   Callback.Callback CtrlName + ".Click(" + CStr(mIndex) + ")"
   Else
      Callback.Callback CtrlName + ".Click()"
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

Public Property Let Index(ByVal rhs As Long)
mIndex = rhs
End Property
Public Property Get Enabled() As Long
Enabled = glistN.Enabled
End Property

Public Property Let Enabled(ByVal rhs As Long)
glistN.Enabled = rhs
End Property

Public Sub Shutdown()
On Error Resume Next
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
Public Sub Curve(rhs As Single)
glistN.Curve False, rhs
End Sub

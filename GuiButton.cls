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
Dim Callback As Form
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
Public Sub ConstructContainer(there As Form, name$, containerCtrl As Object)
Set Callback = there
 Set glistN = there.Controls.Add("m2000.glist", name$, containerCtrl)
there.AddGuiControl Me
CtrlName = nam$
 glistN.Dynamic
End Sub
Public Sub Move(x, y, w, h)
glistN.MoveTwips CLng(x), CLng(y), CLng(w), CLng(h)
End Sub
Public Sub deconstruct()
Shutdown
Set Callback = Nothing
  Set glistN = Nothing
End Sub
Public Sub SetUp()
Dim ff As New StdFont
With Callback.Controls(0).Font
ff.name = .name
ff.charset = .charset
ff.bold = .bold
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
.FontSize = 12
.ForeColor = Callback.Controls(0).ForeColor
.dcolor = Callback.Controls(0).dcolor

.Enabled = True
.ShowMe
End With
End Sub

Private Sub Class_Initialize()
mCaptext = "Button"
End Sub

Private Sub Class_Terminate()
Set glistN = Nothing
End Sub

Private Sub glistN_CheckGotFocus()
''Callback.Callback CtrlName + ".GotFocus()"
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
''glistN.BackColor = rgb(100, 200, 160)
    Callback.Callback CtrlName + ".PanLeftRight(" + CStr(Direction) + ")"
''glistN.PanPos = 0
''glistN.ShowMe
End Sub

Private Sub glistN_RegisterGlist(this As gList)
this.NoWheel = True
End Sub

Private Sub glistN_Selected2(item As Long)
If item >= 0 Then
   '' If (Not mCallback Is Nothing) Then mCallback.Press Index
      Callback.Callback CtrlName + ".Click()"
    End If
End Sub
'
'
Public Property Let caption(myCap As String)
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
Public Property Get caption() As String
caption = mCaptext
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


VERSION 5.00
Begin VB.Form GuiM2000 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   Caption         =   "aaa"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "GuiM2000.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame ResizeMark 
      Appearance      =   0  'Flat
      BackColor       =   &H003B3B3B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   8475
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   873
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   3881787
      ForeColor       =   16777215
      CapColor        =   16777215
   End
End
Attribute VB_Name = "GuiM2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Dim setupxy As Single
Dim Lx As Single, ly As Single, dr As Boolean
Dim scrTwips As Long
Dim bordertop As Long, borderleft As Long
Dim allwidth As Long, itemWidth As Long
Private ExpandWidth As Boolean, lastfactor As Single
Private myEvent As mEvent
Private GuiControls As New Collection
'Dim gList1 As gList
Dim onetime As Boolean, PopupOn As Boolean
Dim alfa As New GuiButton
Public MyName$
Public ModuleName$
Public prive As Long
Private ByPassEvent As Boolean
Private mIndex As Long
Private mSizable As Boolean
Public Relax As Boolean
Private MarkSize As Long
Public MY_BACK As New cDIBSection
Dim CtrlFont As New StdFont
Dim novisible As Boolean
Private mModalId As Double, mModalIdPrev As Double
Public IamPopUp As Boolean
Private mEnabled As Boolean
Public Sub AddGuiControl(widget As Object)
GuiControls.Add widget
End Sub
Public Sub TestModal(alfa As Double)
If mModalId = alfa Then
mModalId = mModalIdPrev
mModalIdPrev = 0
Enablecontrol = True
End If
End Sub
Property Get Modal() As Double
    Modal = mModalId
End Property
Property Let Modal(RHS As Double)
mModalIdPrev = mModalId
mModalId = RHS
End Property

Public Property Let Enablecontrol(RHS As Boolean)
If RHS = False Then UnHook hWnd '  And Not Me Is Screen.ActiveForm Then UnHook hWnd
If Len(MyName$) = 0 Then Exit Property
'If rhs = Fals Then UnHook hWnd
If mEnabled = False And RHS = True Then Me.enabled = True
mEnabled = RHS

Dim w As Object
If Controls.Count > 0 Then
For Each w In Me.Controls
If w Is gList2 Then
gList2.enabled = RHS
gList2.mousepointer = 0
ElseIf w.Visible Then
w.enabled = RHS
If TypeOf w Is gList Then w.TabStop = RHS
End If
Next w
End If
Me.enabled = RHS
End Property
Public Property Get Enablecontrol() As Boolean
If Len(MyName$) = 0 Then Enablecontrol = False: Exit Property
Enablecontrol = mEnabled


End Property


Property Get NeverShow() As Boolean
NeverShow = Not novisible
End Property
Friend Property Set EventObj(aEvent As Object)
Set myEvent = aEvent
End Property

Public Sub Callback(b$)
If ByPassEvent Then
CallEventFromGuiOne Me, myEvent, b$
Else
CallEventFromGui Me, myEvent, b$
End If
End Sub
Public Sub CallbackNow(b$, VR())

CallEventFromGuiNow Me, myEvent, b$, VR()
End Sub


Public Sub ShowmeALL()
Dim w As Object
If Controls.Count > 0 Then
For Each w In Controls
If w.enabled Then w.Visible = True
    
Next w
End If
gList2.PrepareToShow
End Sub
Public Sub RefreshALL()
Dim w As Object
If Controls.Count > 0 Then
For Each w In Controls
If w.Visible Then
If TypeOf w Is gList Then w.ShowMe2
End If
Next w
End If
Refresh
End Sub

Private Sub Form_Click()
If gList2.Visible Then gList2.SetFocus
If Index > -1 Then
    Callback MyName$ + ".Click(" + CStr(Index) + ")"
Else
    Callback MyName$ + ".Click()"
End If
End Sub

Private Sub Form_Activate()
If PopupOn Then PopupOn = False
If novisible Then Hide: Unload Me
If gList2.HeadLine <> "" Then If ttl Then Form3.CaptionW = gList2.HeadLine: Form3.Refresh
MarkSize = 4
ResizeMark.Width = MarkSize * dv15
ResizeMark.Height = MarkSize * dv15
ResizeMark.Left = Width - MarkSize * dv15
ResizeMark.top = Height - MarkSize * dv15

ResizeMark.BackColor = GetPixel(Me.hDC, 0, 0)
ResizeMark.Visible = Sizable
If Sizable Then ResizeMark.ZOrder 0
If Typename(ActiveControl) = "gList" Then
Hook hWnd, ActiveControl
Else
Hook hWnd, Nothing
End If
End Sub
Private Sub Form_Deactivate0()
If PopupOn Then
UnHook hWnd

Exit Sub
End If
If IamPopUp Then
If mModalId = ModalId And ModalId <> 0 Then
        
        If Visible Then Hide
       
        ModalId = 0
            novisible = False
End If
Else
    If mModalId = ModalId And ModalId <> 0 Then
        If Visible Then
            On Error Resume Next
            Me.SetFocus
        Else
        UnHook hWnd
            If mModalId <> 0 Then ModalId = 0
 
            
        End If
    
    Else
    UnHook hWnd
    End If
   
    End If
End Sub


Private Sub Form_Deactivate()
            UnHook hWnd
If PopupOn Then

Exit Sub
End If
If IamPopUp Then
If mModalId = ModalId And ModalId <> 0 Then
If Visible Then Hide
ModalId = 0
novisible = False
End If
Else
If mModalId = ModalId And ModalId <> 0 Then If Not Visible Then If mModalId <> 0 Then ModalId = 0
End If

End Sub


Private Sub Form_Initialize()
mEnabled = True
End Sub

Private Sub Form_LostFocus()
If Index > -1 Then
    Callback MyName$ + ".LostFocus(" + CStr(Index) + ")"
Else
    Callback MyName$ + ".LostFocus()"
End If
If HOOKTEST <> 0 Then
UnHook hWnd
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then



Relax = True
If Index > -1 Then
    Callback MyName$ + ".MouseDown(" + CStr(Index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
    Callback MyName$ + ".MouseDown(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If



Relax = False
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then
Relax = True

If Index > -1 Then
Callback MyName$ + ".MouseMove(" + CStr(Index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback MyName$ + ".MouseMove(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then

Relax = True

If Index > -1 Then
Callback MyName$ + ".MouseUp(" + CStr(Index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback MyName$ + ".MouseUp(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If mModalId = ModalId And ModalId <> 0 Then
    If Visible Then Hide
    If mModalId <> 0 Then ModalId = 0
    Cancel = True
    novisible = False
ElseIf mModalId <> 0 And Visible Then
    mModalId = mModalIdPrev
    mModalIdPrev = 0
    If mModalId > 0 Then
        Cancel = True
    Else
      '    Set LastGlist = Nothing

    End If
Else
mModalIdPrev = 0
'Set LastGlist = Nothing
End If
End Sub

Private Sub Form_Resize()
gList2.MoveTwips 0, 0, Me.Width, gList2.HeightTwips
ResizeMark.Move Width - ResizeMark.Width, Height - ResizeMark.Height
End Sub


Private Sub gList2_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, Skip As Boolean)
If item = -1 Then
FillThere thisHDC, thisrect, gList2.CapColor
FillThereMyVersion thisHDC, thisrect, &H999999
Skip = True
End If
End Sub
Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If gList2.DoubleClickCheck(Button, item, x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor, -1) Then
    ByeBye
End If
End Sub
Sub ByeBye()
Dim var(1) As Variant
var(1) = CLng(0)
If mIndex > -1 Then
CallEventFromGuiNow Me, myEvent, MyName$ + ".Unload(" + CStr(mIndex) + ")", var()
Else
CallEventFromGuiNow Me, myEvent, MyName$ + ".Unload()", var()
End If
If var(0) = 0 Then
                     If ttl Then
                     Form3.CaptionW = ""
                     If Form3.WindowState = 1 Then Form3.WindowState = 0
               
                    Unload Form3
             End If
                              Unload Me
                      End If
End Sub

Private Sub Form_Load()

If onetime Then
novisible = True
Exit Sub
End If
onetime = True
' try0001
Set LastGlist = Nothing
scrTwips = Screen.TwipsPerPixelX
' clear data...
lastfactor = 1
setupxy = 20
gList2.enabled = True
gList2.CapColor = rgb(255, 160, 0)
gList2.FloatList = True
gList2.MoveParent = True
gList2.HeadLine = ""
gList2.HeadLine = "Form"
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
gList2.TabStop = False
With gList2.Font
CtrlFont.name = .name
CtrlFont.Size = .Size
CtrlFont.bold = .bold
End With
gList2.FloatLimitTop = ScrY() - 600
gList2.FloatLimitLeft = ScrX() - 450

End Sub

Private Sub Form_Unload(Cancel As Integer)
UNhookMe
Set myEvent = Nothing

If prive <> 0 Then
players(prive).used = False
players(prive).MAXXGRAPH = 0  '' as a flag
prive = 0
End If
Dim w As Object
If GuiControls.Count > 0 Then
For Each w In GuiControls
    w.deconstruct
Next w
End If
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
Private Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT
CopyFromLParamToRect A, thatRect

FillBack thathDC, A, thatbgcolor
End Sub

Private Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT, b As Long
b = 2 * lastfactor
If b < 2 Then b = 2
If setupxy - b < 0 Then b = setupxy \ 4 + 1
CopyFromLParamToRect A, thatRect
A.Left = b
A.Right = setupxy - b
A.top = b
A.Bottom = setupxy - b
FillThere thathDC, VarPtr(A), 0
b = 5 * lastfactor
A.Left = b
A.Right = setupxy - b
A.top = b
A.Bottom = setupxy - b
FillThere thathDC, VarPtr(A), rgb(255, 160, 0)
End Sub

Public Property Get TITLE() As Variant
TITLE = gList2.HeadLine
End Property

Public Property Let TITLE(ByVal vNewValue As Variant)
' A WORKAROUND TO CHANGE TITLE WHEN FORM IS DISABLED BY A MODAL FORM
On Error Resume Next
Dim oldenable As Boolean
oldenable = gList2.enabled
gList2.enabled = True
gList2.HeadLine = ""
If Trim(vNewValue) = "" Then vNewValue = " "
gList2.HeadLine = vNewValue
gList2.HeadlineHeight = gList2.HeightPixels
If oldenable = False Then gList2.ShowMe
gList2.enabled = oldenable
End Property
Public Property Get Index() As Long
Index = mIndex
End Property

Public Property Let Index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Sub CloseNow()
Dim w As Object
    If mModalId = ModalId And ModalId <> 0 Then
        ModalId = 0
      If Visible Then Hide
    Else
    mModalId = 0
    For Each w In GuiControls
    If Typename(w) Like "Gui*" Then
    w.deconstruct
    End If
Next w
Set w = Nothing
         If ttl Then
                     Form3.CaptionW = ""
                     If Form3.WindowState = 1 Then Form3.WindowState = 0
               
                    Unload Form3
             End If

Unload Me
    End If
End Sub
Public Function Control(Index) As Object
On Error Resume Next
Set Control = Controls(Index)
If Err > 0 Then Set Control = Me
End Function
Public Sub Opacity(mAlpha, Optional mlColor = 0, Optional mTRMODE = 0)
SetTrans Me, CInt(Abs(mAlpha)) Mod 256, CLng(mycolor(mlColor)), CBool(mTRMODE)
End Sub
Public Sub Hold()
MY_BACK.ClearUp
If MY_BACK.Create(Form1.Width / DXP, Form1.Height / DYP) Then
MY_BACK.LoadPictureBlt hDC
If MY_BACK.bitsPerPixel <> 24 Then Conv24 MY_BACK
End If
End Sub
Public Sub Release()
MY_BACK.PaintPicture hDC
End Sub


Public Property Get ByPass() As Variant
ByPass = ByPassEvent
End Property

Public Property Let ByPass(ByVal vNewValue As Variant)
ByPassEvent = CBool(vNewValue)
End Property
Property Get TitleHeight() As Variant
TitleHeight = gList2.Height
End Property
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set gList2.Font = aa
gList2.Height = gList2.HeadlineHeightTwips
lastfactor = gList2.HeadlineHeight / 30
setupxy = 20 * lastfactor
 gList2.Dynamic

End If
End Sub
Public Sub CtrlFontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)

If ThisFontName <> "" Then

CtrlFont.name = ThisFontName

If ThisMode > 7 Then CtrlFont.Size = ThisMode Else CtrlFont = 7
CtrlFont.bold = ThisBold

End If
End Sub
Public Property Get CtrlFontName()
    CtrlFontName = CtrlFont.name
End Property
Public Property Get CtrlFontSize()
    CtrlFontSize = CtrlFont.Size
End Property
Public Property Get CtrlFontBold()
    CtrlFontBold = CtrlFont.bold
End Property





Private Sub gList2_RefreshDesktop()
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
End Sub
Public Sub PopUp(vv As Variant, ByVal x As Variant, ByVal y As Variant)
Dim var1() As Variant, retobject As Object, that As Object
ReDim var1(0 To 1)
Dim var2() As String
ReDim var2(0 To 0)

x = x + Left
y = y + top
Set that = vv
If Me Is that Then Exit Sub
If that.Visible Then
If Not that.enabled Then Exit Sub
End If
If x + that.Width > ScrX() Then
If y + that.Height > ScrY() Then
that.Move ScrX() - that.Width, ScrY() - that.Height
Else
that.Move ScrX() - that.Width, y
End If
ElseIf y + that.Height > ScrY() Then
that.Move x, ScrY() - Height
Else
that.Move x, y
End If
var1(1) = 1
Set var1(0) = Me
that.IamPopUp = True
CallByNameFixParamArray that, "Show", VbMethod, var1(), var2(), 2
Set that = Nothing
Set var1(0) = Nothing
'Show
MyDoEvents

End Sub
Public Sub PopUpPos(vv As Variant, ByVal x As Variant, ByVal y As Variant)
Dim that As Object
x = x + Left
y = y + top
Set that = vv
If Me Is that Then Exit Sub
If that.Visible Then
If Not that.enabled Then Exit Sub
End If
If x + that.Width > ScrX() Then
If y + that.Height > ScrY() Then
that.Move ScrX() - that.Width, ScrY() - that.Height
Else
that.Move ScrX() - that.Width, y
End If
ElseIf y + that.Height > ScrY() Then
that.Move x, ScrY() - Height
Else
that.Move x, y
End If
that.ShowmeALL
'If ModalId <> 0 Then
PopupOn = True

that.Show , Me

End Sub
Public Sub hookme(This As gList)
Set LastGlist = This
End Sub

Private Sub ResizeMark_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If Sizable And Not dr Then
    x = x + ResizeMark.Left
    y = y + ResizeMark.top
    If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
    
    dr = Button = 1
    ResizeMark.mousepointer = vbSizeNWSE
    Lx = x
    ly = y
    If dr Then Exit Sub
    
    End If
    
End If
End Sub

Private Sub ResizeMark_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim addy As Single, addx As Single
If Not Relax Then
    x = x + ResizeMark.Left
    y = y + ResizeMark.top
    If Button = 0 Then If dr Then Me.mousepointer = 0: dr = False: Relax = False: Exit Sub
    Relax = True
    If dr Then
         If y < (Height - 150) Or y >= Height Then addy = (y - ly) Else addy = dv15 * 5
         If x < (Width - 150) Or x >= Width Then addx = (x - Lx) Else addx = dv15 * 5
         If Width + addx >= 1800 And Width + addx < ScrX() Then
             If Height + addy >= 1800 And Height + addy < ScrY() Then
                Lx = x
                ly = y
                Move Left, top, Width + addx, Height + addy
                If Index > -1 Then
                    Callback MyName$ + ".Resize(" + CStr(Index) + ")"
                Else
                    Callback MyName$ + ".Resize()"
                End If
            End If
        End If
        Relax = False
        Exit Sub
    Else
        If Sizable Then
            If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
                    dr = Button = 1
                    ResizeMark.mousepointer = vbSizeNWSE
                    Lx = x
                    ly = y
                    If dr Then Relax = False: Exit Sub
                Else
                    ResizeMark.mousepointer = 0
                    dr = 0
                End If
            End If
    End If
Relax = False
End If
End Sub

Public Property Get Sizable() As Variant
Sizable = mSizable
End Property

Public Property Let Sizable(ByVal vNewValue As Variant)
mSizable = vNewValue
ResizeMark.enabled = vNewValue
If ResizeMark.enabled Then
ResizeMark.Visible = Me.Visible
Else
ResizeMark.Visible = False
End If
End Property
Public Property Let SizerWidth(ByVal vNewValue As Variant)
If vNewValue \ dv15 > 1 Then
    MarkSize = vNewValue \ dv15
    With ResizeMark
    .Width = MarkSize * dv15
    .Height = MarkSize * dv15
    .Move Width - .Width, Height - .Height
    End With
End If
End Property

Public Property Get header() As Variant
header = gList2.Visible
End Property

Public Property Let header(ByVal vNewValue As Variant)
gList2.Visible = vNewValue
End Property
Sub GetFocus()
On Error Resume Next
Me.SetFocus
End Sub
Public Sub UNhookMe()
Set LastGlist = Nothing
UnHook hWnd
End Sub

VERSION 5.00
Begin VB.Form GuiM2000 
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
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   0
      TabIndex        =   0
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
Dim Lx As Long, ly As Long, dr As Boolean
Dim scrTwips As Long
Dim bordertop As Long, borderleft As Long
Dim allwidth As Long, itemWidth As Long
Private ExpandWidth As Boolean, lastfactor As Single
Private myEvent As mEvent
Private GuiControls As New Collection
'Dim gList1 As gList
Dim onetime As Boolean
Dim alfa As New GuiButton
Public MyName$
Public ModuleName$
Public Sub AddGuiControl(widget As Object)
GuiControls.Add widget
End Sub

Friend Property Set EventObj(aEvent As Object)
Set myEvent = aEvent
End Property

Public Sub Callback(b$)
CallEventFromGui Me, myEvent, b$

End Sub
Public Sub CallbackNow(b$, vr())
CallEventFromGuiNow Me, myEvent, b$, vr()

End Sub


Public Sub ShowmeALl()
Dim w As Object
If Controls.Count > 0 Then
For Each w In Controls
If w.Enabled Then w.Visible = True
    
Next w
End If
End Sub


Private Sub Form_Click()
gList2.SetFocus
CallEventFromGui Me, myEvent, MyName$ + ".Click()"
End Sub

Private Sub Form_Resize()
gList2.MoveTwips 0, 0, Me.Width, gList2.HeightTwips
End Sub

Private Sub Form_Terminate()
Dim w As Object
If GuiControls.Count > 0 Then
For Each w In GuiControls
    w.deconstruct
Next w
End If
End Sub

Private Sub gList2_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = -1 Then
FillThere thisHDC, thisrect, gList2.CapColor
FillThereMyVersion thisHDC, thisrect, &H999999
skip = True
End If
End Sub
Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If gList2.DoubleClickCheck(Button, item, x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor, -1) Then
                      Unload Me
End If
End Sub


Private Sub Form_Load()
If onetime Then Exit Sub
onetime = True
scrTwips = Screen.TwipsPerPixelX
' clear data...
lastfactor = 1
setupxy = 20
gList2.Enabled = True
gList2.CapColor = rgb(255, 160, 0)
gList2.FloatList = True
gList2.MoveParent = True
gList2.HeadLine = ""
gList2.HeadLine = "My Caption"
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
gList2.TabStop = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set myEvent = Nothing

End Sub

Public Sub PrintMe(ParamArray aa() As Variant)
Dim i As Long
For i = LBound(aa()) To UBound(aa())
Print aa(i),
Next i
Print
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
Private Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT
CopyFromLParamToRect a, thatRect
FillBack thathDC, a, thatbgcolor
End Sub

Public Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT, b As Long
b = 2
CopyFromLParamToRect a, thatRect
a.Left = b
a.Right = setupxy - b
a.top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), 0
b = 5
a.Left = b
a.Right = setupxy - b
a.top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), rgb(255, 160, 0)
End Sub

Public Property Get TITLE() As Variant
TITLE = gList2.HeadLine
End Property

Public Property Let TITLE(ByVal vNewValue As Variant)
gList2.HeadLine = ""
gList2.HeadLine = vNewValue
gList2.HeadlineHeight = gList2.HeightPixels
End Property

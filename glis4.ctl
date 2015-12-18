VERSION 5.00
Begin VB.UserControl gList 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   FillColor       =   &H80000002&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   MousePointer    =   1  'Arrow
   OLEDropMode     =   1  'Manual
   PropertyPages   =   "glis4.ctx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   7245
   Begin VB.Timer Timer2bar 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2745
      Top             =   2565
   End
   Begin VB.Timer Timer1bar 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1950
      Top             =   1710
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5475
      Top             =   3585
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5505
      Top             =   1035
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5940
      Top             =   2595
   End
End
Attribute VB_Name = "gList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'm2000 ver 2.1
Option Explicit
Dim waitforparent As Boolean
Dim havefocus As Boolean, UKEY$
Dim DUMMY As Long
Private Type Myshape
Visible As Boolean
hatchType As Long
top As Long
Left As Long
Width As Long
Height As Long
End Type
Private mynum$
Public overrideTextHeight As Long
Public AutoHide As Boolean, NoWheel As Boolean
Private Shape1 As Myshape, Shape2 As Myshape, Shape3 As Myshape
Private Type RECT
        Left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type
Private Type itemlist
    selected As Boolean  ' use this for multiselect or checked
    Checked As Boolean  ' use this to use list item as menu
    radiobutton As Boolean  ' use this to checked like radio buttons ..with auto unselect between to lines...or all list if not lines foundit
    content As String
    contentID As String
    line As Boolean
End Type
Private fast As Boolean
Private Declare Function GdiFlush Lib "gdi32" () As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Const PS_NULL = 5
Private Const PS_SOLID = 0
Public restrictLines As Long
Private nowx As Single, nowy As Single
Private marvel As Boolean
Private Const DT_BOTTOM As Long = &H8&
Private Const DT_CALCRECT As Long = &H400&
Private Const DT_CENTER As Long = &H1&
Private Const DT_EDITCONTROL As Long = &H2000&
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_EXPANDTABS As Long = &H40&
Private Const DT_EXTERNALLEADING As Long = &H200&
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const DT_INTERNAL As Long = &H1000&
Private Const DT_LEFT As Long = &H0&
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_NOCLIP As Long = &H100&
Private Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
Private Const DT_NOPREFIX As Long = &H800&
Private Const DT_PATH_ELLIPSIS As Long = &H4000&
Private Const DT_PREFIXONLY As Long = &H200000
Private Const DT_RIGHT As Long = &H2&
Private Const DT_SINGLELINE As Long = &H20&
Private Const DT_TABSTOP As Long = &H80&
Private Const DT_TOP As Long = &H0&
Private Const DT_VCENTER As Long = &H4&
Private Const DT_WORDBREAK As Long = &H10&
Private Const DT_WORD_ELLIPSIS As Long = &H40000

Const m_def_Text = ""
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = 0
Const m_def_Enabled = False
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_dcolor = &H333333
Const m_def_CapColor = &HAAFFBB
Const m_def_Showbar = True
Const m_def_sync = ""

Dim m_sync As String
Dim m_backcolor As Long
Dim m_ForeColor As Long
'Dim m_Enabled As Boolean
Dim m_font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_CapColor As Long
Dim m_dcolor As Long

Dim m_showbar As Boolean
Dim mlist() As itemlist
Dim topitem As Long
Dim itemcount As Long
Dim Mselecteditem As Long
Event selected(item As Long)
Event SelectedMultiAdd(item As Long)
Event SelectedMultiSub(item As Long)
Event Selected2(item As Long)
Event softSelected(item As Long)
Event Maybelanguage()
Event MouseUp(x As Single, y As Single)
Event SpecialColor(rgbcolor As Long)
Event RemoveOne(that As String)
Event PushMark2Undo(that As String)
Event PushUndoIfMarked()
Event addone(that As String)
Event MayRefresh(ok As Boolean)
Event CheckGotFocus()
Event CheckLostFocus()
Event DragData(ThatData As String)
Event DragPasteData(ThatData As String)
Event DropOk(ok As Boolean)
Event DropFront(ok As Boolean)
Event ScrollMove(item As Long)
Event OutPopUp(x As Single, y As Single, myButton As Integer)
Event SplitLine()
Event LineUp()
Event LineDown()
Event MarkIn()
Event MarkOut()
Event MarkDestroyAny()
Event MarkDestroy()
Event MarkDelete(preservecursor As Boolean)
Event WordMarked(ThisWord As String)
Event ShowExternalCursor()
Event ChangeSelStart(thisselstart As Long)
Event ReadListItem(item As Long, content As String)
Event ChangeListItem(item As Long, content As String)
Event HeaderSelected(Button As Integer)
Event BlockCaret(item As Long, blockme As Boolean, skipme As Boolean)
Event ScrollSelected(item As Long, y As Long)
Event MenuChecked(item As Long)
Event PromptLine(ThatLine As Long)
Event PanLeftRight(Direction As Boolean)
Event GetBackPicture(pic As Object)
Event KeyDown(KeyCode As Integer, shift As Integer)
Event KeyDownAfter(KeyCode As Integer, shift As Integer)
Event SyncKeyboard(item As Integer)
Event find(key As String, where As Long, skip As Boolean)
Event ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Event ExposeListcount(cListCount As Long)
Event ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
Event MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Event SpinnerValue(ThatValue As Long)
Event RegisterGlist(this As gList)
Event UnregisterGlist()
Private state As Boolean
Private secreset As Boolean
Private scrollme As Long
Private scrolledit As Long
Private ly As Long, dr As Boolean
Private drc As Boolean
Private scrTwips As Long
Private cY As Long
Private cX As Long
Dim myt As Long
Dim mytPixels As Long
Public BarColor As Long
Public BarHatch As Long
Public BarHatchColor As Long
Public LeaveonChoose As Boolean
Public BypassLeaveonChoose As Boolean
Public LastSelected As Long
Public NoPanLeft As Boolean
Public NoPanRight As Boolean
Private LastVScroll As Long
Private FreeMouse As Boolean
Public NoCaretShow As Boolean

Dim valuepoint As Long, minimumWidth As Long
Dim mValue As Long, mmax As Long, mmin As Long, mLargeChange As Long  ' min 1
Dim mSmallChange As Long  ' min 1
Dim mVertical As Boolean
Dim OurDraw As Boolean, GetOpenValue As Long
Dim lastX As Single, LastY As Single

Private mjumptothemousemode As Boolean
Private mpercent As Single
Private barwidth As Long
Private NoFire As Boolean
Public addpixels As Long
Public StickBar As Boolean
Dim Hidebar As Boolean
Dim myEnabled As Boolean
Public WrapText As Boolean
Public CenterText As Boolean
Public VerticalCenterText As Boolean
Private mHeadline As String
Private mHeadlineHeight As Long
Private mHeadlineHeightTwips As Long
Public MultiSelect As Boolean
Public LeftMarginPixels As Long
Dim Buffer As Long
Public FloatList As Boolean
Public MoveParent As Boolean
Public BlockItemcount As Boolean
Private useFloatList As Boolean
Public HeadLineHeightMinimum As Long
Private mPreserveNpixelsHeaderRight As Long
Public AutoPanPos As Boolean   ' used if we have no EditFlag
Public FloatLimitLeft As Long
Public FloatLimitTop As Long
Public mEditFlag As Boolean
Public SingleLineSlide As Boolean
Private mSelstart As Long
Private caretCreated As Boolean
Public MultiLineEditBox As Boolean
Public NoScroll As Boolean
Public MarkNext As Long  ' 0 - markin, 1- Markout
Public Noflashingcaret As Boolean
Public NoFreeMoveUpDown As Boolean  ' if true then keyup and keydown scroll up down the list
Public PromptLineIdent As Long ' to be a console we need prompt line to have some chars untouch perhaps this ">"
Public LastLinePart As String
Public Spinner As Boolean ' if true and restrictline =1 - we have events for up down values
Public maxchar As Long ' for non multiline
Public WordCharLeft As String
Public WordCharRight As String
Public WordCharRightButIncluded As String
Public DropEnabled As Boolean
Public DragEnabled As Boolean
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&) ' not NT?
Private Const DWL_ANYTHREAD& = 0
Const LOCALE_ILANGUAGE = 1
Private Declare Function PeekMessageW Lib "user32" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Const WM_KEYFIRST = &H100
 Const WM_KEYLAST = &H108
 Private Type POINTAPI
    x As Long
    y As Long
End Type
 Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Dim doubleclick As Long
Dim mlx As Long, mly As Long
Public SkipForm As Boolean
Public dropkey As Boolean
Public Function GetLastKeyPressed() As Long
Dim Message As Msg
If mynum$ <> "" Then
GetLastKeyPressed = -1
ElseIf PeekMessageW(Message, 0, WM_KEYFIRST, WM_KEYLAST, 0) Then
        GetLastKeyPressed = Message.wParam
    Else
        GetLastKeyPressed = -1
    End If
    Exit Function
End Function
Public Property Let HeadlineHeight(ByVal rhs As Long)
If HeadLine <> "" Then
mHeadlineHeight = rhs
mHeadlineHeightTwips = rhs * scrTwips

Else
mHeadlineHeight = 0
mHeadlineHeightTwips = 0

End If
End Property

Public Property Get HeadlineHeight() As Long
If HeadLine <> "" Then
HeadlineHeight = mHeadlineHeight
Else
HeadlineHeight = 0

End If
End Property

Public Property Let HeadLine(ByVal rhs As String)
If mHeadline = "" Then
' reset headlineheight
mHeadline = rhs
HeadlineHeight = UserControlTextHeight() / scrTwips

Exit Property

End If
mHeadline = rhs
End Property

Public Property Get HeadLine() As String
HeadLine = mHeadline
End Property
Public Sub PrepareToShow(Optional delay As Single = 10)
 barwidth = UserControlTextWidth("W")
 CalcAndShowBar1
Timer1.Enabled = False
If delay < 1 Then delay = 1
If fast Then
fast = False
Timer1.Interval = delay
Else
Timer1.Interval = delay * 5
End If
Timer1.Enabled = True
End Sub
Public Sub PressSoft()
secreset = False
RaiseEvent Selected2(SELECTEDITEM - 1)
End Sub
Public Property Get ScrollFrom() As Long
    ScrollFrom = topitem
End Property
Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal rhs As Integer)
    m_BorderStyle = rhs
    
 If BackStyle = 0 Then UserControl.BorderStyle = m_BorderStyle Else UserControl.BorderStyle = 0
    PropertyChanged "BorderStyle"
End Property
Public Property Get sync() As String
    sync = m_sync
End Property
Public Property Let sync(ByVal New_sync As String)
    If Ambient.UserMode Then Err.Raise 393
    m_sync = New_sync
    PropertyChanged "sync"
End Property
Public Property Get hWnd() As Long
hWnd = UserControl.hWnd
End Property
Public Property Let Text(ByVal new_text As String)
clear True
If new_text <> "" Then
If Right$(new_text, 2) <> vbCrLf And new_text <> "" Then
new_text = new_text + vbCrLf
End If
Dim mpos As Long, b$
Do
b$ = GetStrUntilB(mpos, vbCrLf, new_text)
additemFast b$  ' and blank lines
Loop Until mpos > Len(new_text) Or mpos = 0
End If
If UserControl.Ambient.UserMode = False Then
Repaint
SELECTEDITEM = 0
CalcAndShowBar
ShowMe
End If

PropertyChanged "Text"
End Property
Public Property Let ListText(ByVal new_text As String)
clear True
If Right$(new_text, 2) <> vbCrLf And new_text <> "" Then
new_text = new_text + vbCrLf
End If
Dim mpos As Long, b$
Do
b$ = GetStrUntilB(mpos, vbCrLf, new_text)

If Left$(b$, 1) <> "_" Then
additemFast b$
Else
b$ = Mid$(b$, 2)
If b$ = "" Then
addsep
Else
additemFast b$
menuEnabled(itemcount - 1) = False
End If
End If
Loop Until mpos > Len(new_text) Or mpos = 0
Repaint
SELECTEDITEM = 0
CalcAndShowBar
ShowMe
End Property
Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
Dim i As Long
For i = 0 To listcount - 2
Text = Text + List(i) + vbCrLf
Next i
Text = Text + List(i)
End Property
Public Sub ScrollTo(ThatTopItem As Long, Optional this As Long = -2)
On Error GoTo scroend
topitem = ThatTopItem
If topitem < 0 Then topitem = 0
If this > -2 Then
SELECTEDITEM = this
End If
CalcAndShowBar1
Timer1.Enabled = True
scroend:
End Sub
Public Sub ScrollToSilent(ThatTopItem As Long, Optional this As Long = -2)
On Error GoTo scroend
topitem = ThatTopItem
If topitem < 0 Then topitem = 0
If this > -2 Then
SELECTEDITEM = this
End If
If BarVisible Then Redraw ShowBar
Timer1.Enabled = True
scroend:
End Sub
Public Sub CalcAndShowBar()
CalcAndShowBar1
ShowMe2
End Sub
Private Sub CalcAndShowBar1()
Dim oldvalue As Long, oldmax As Long
oldvalue = Value
oldmax = Max
On Error GoTo calcend
state = True

   On Error Resume Next
            Err.clear
    If Not Spinner Then
            If listcount - 1 - lines < 1 Then
            Max = 1
            Else
            Max = listcount - 1 - lines
            largechange = lines
            End If
            If Err.Number > 0 Then
                Value = listcount - 1
                Max = listcount - 1
            End If
                      Value = topitem
        End If

state = False
If listcount < lines + 2 Then
BarVisible = False
Else
Redraw Hidebar

End If
calcend:
End Sub
Public Property Get ListValue() As String
' this was text before
If SELECTEDITEM <= 0 Then Else ListValue = List(listindex)
End Property

Public Property Get listcount() As Long
Dim thatlistcount As Long
RaiseEvent ExposeListcount(thatlistcount)
If thatlistcount > 0 Then
listcount = thatlistcount
Else
  listcount = itemcount
  End If
End Property
Public Property Let ShowBar(ByVal rhs As Boolean)
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
Else
myt = UserControlTextHeight() + addpixels * scrTwips
End If
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
    m_showbar = rhs
    barwidth = UserControlTextWidth("W")
    
    state = True
    Value = 0
    state = False
 
    If listcount >= lines Then
BarVisible = (m_showbar Or StickBar Or AutoHide) Or Hidebar
Else
Redraw (m_showbar Or StickBar Or AutoHide) Or Hidebar
End If
   
'RepaintScrollBar
End Property
Public Property Get ShowBar() As Boolean
If Hidebar Then
ShowBar = True ' TEMPORARY USE
Else

    ShowBar = m_showbar Or StickBar Or AutoHide
    End If
End Property

Public Property Let BackColor(ByVal rhs As OLE_COLOR)

    m_backcolor = rhs
UserControl.BackColor = rhs
  PropertyChanged "BackColor"
    
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_backcolor 'UserControl.Backcolor
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal rhs As OLE_COLOR)
    m_ForeColor = rhs
    UserControl.ForeColor = rhs
    PropertyChanged "ForeColor"
End Property
Public Property Get CapColor() As OLE_COLOR
    CapColor = m_CapColor
  
End Property

Public Property Let CapColor(ByVal rhs As OLE_COLOR)
    m_CapColor = rhs
    PropertyChanged "CapColor"
End Property
Public Property Get dcolor() As OLE_COLOR
    dcolor = m_dcolor
  
End Property

Public Property Let dcolor(ByVal rhs As OLE_COLOR)
    m_dcolor = rhs
    PropertyChanged "dcolor"
End Property
Public Property Get Enabled() As Boolean
    Enabled = myEnabled
End Property
Public Property Let Enabled(ByVal rhs As Boolean)
 myEnabled = rhs
    PropertyChanged "Enabled"
    On Error Resume Next
    Dim MM$, mo As Control, nm$, cnt$, p As Long
    
''new position
If Not waitforparent Then Exit Property
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
'' old position
If UserControl.Parent Is Nothing Then Exit Property
If Err.Number > 0 Then Exit Property
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
mo.TabStop = rhs
End Property

Public Property Get Font() As Font

Dim i As Integer
 Set Font = m_font
End Property

Public Property Set Font(New_Font As Font)
    Set m_font = New_Font
Set UserControl.Font = m_font
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
Else

myt = UserControlTextHeight() + addpixels * scrTwips
End If

HeadlineHeight = UserControlTextHeight() / scrTwips
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
    PropertyChanged "Font"
End Property
Public Sub CalcNewFont()
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
Else

myt = UserControlTextHeight() + addpixels * scrTwips
End If
HeadlineHeight = UserControlTextHeight() / scrTwips
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
If listindex >= 0 Then
CalcAndShowBar1
    ShowThis listindex + 1
Else
    ShowMe True
End If

End Sub

Public Property Get FontSize() As Single

  FontSize = m_font.Size
 
End Property

Public Property Let FontSize(New_FontSize As Single)
     If New_FontSize < 6 Then
  m_font.Size = 6
     Else
m_font.Size = New_FontSize
End If

If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
Else
myt = UserControlTextHeight() + addpixels * scrTwips
End If
'HeadlineHeight = UserControlTextHeight() / SCRTWIPS
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips

End Property

Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle

End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
On Error Resume Next
    m_BackStyle = New_BackStyle
    If m_BackStyle = 0 Then UserControl.BorderStyle = m_BorderStyle Else UserControl.BorderStyle = 0
    PropertyChanged "BackStyle"
 
End Property



Private Sub usercontrol_GotFocus1()
Dim YYT As Long
YYT = myt
DrawMode = vbCopyPen
If SELECTEDITEM > 0 Then
If SELECTEDITEM - topitem - 1 <= lines Then
If BackStyle = 1 Then
'If BarVisible Then
'Line (scrollme + SCRTWIPS, (selecteditem - topitem) * yyt)-(scrollme + UserControl.Width - barwidth, (selecteditem - topitem - 1) * yyt), 0, B
'Else
Line (scrollme + scrTwips, (SELECTEDITEM - topitem) * YYT)-(scrollme + UserControl.Width, (SELECTEDITEM - topitem - 1) * YYT), 0, B
'End If
Else
'If BarVisible Then
'Line (scrollme, (selecteditem - topitem) * yyt)-(scrollme + UserControl.Width - barwidth, (selecteditem - topitem - 1) * yyt), 0, B
'Else
Line (scrollme, (SELECTEDITEM - topitem) * YYT)-(scrollme + UserControl.Width, (SELECTEDITEM - topitem - 1) * YYT), 0, B
'End If


End If
End If
End If
DrawMode = vbCopyPen
Timer1.Interval = 40
Timer1.Enabled = True
End Sub

Public Sub LargeBar1KeyDown(KeyCode As Integer, shift As Integer)
Timer1.Enabled = False
If listindex < 0 Then
Else
PressKey KeyCode, shift
End If
End Sub

Private Sub Timer1bar_Timer()
processXY lastX, LastY
End Sub

Private Sub timer2bar_Timer()
If m_showbar Or Shape1.Visible Or Spinner Then Redraw
End Sub


Private Sub UserControl_GotFocus()

RaiseEvent CheckGotFocus
havefocus = True

SoftEnterFocus
If Not NoWheel Then RaiseEvent RegisterGlist(Me)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
If dropkey Then KeyAscii = 0: Exit Sub
Dim bb As Boolean, KK$
If listindex < 0 Then
Else
    If Not state Then
        If KeyAscii = 13 And myEnabled And Not MultiLineEditBox Then
            KeyAscii = 0
            If SELECTEDITEM < 0 Then
                
            ElseIf SELECTEDITEM > 0 Then
                secreset = False
                RaiseEvent Selected2(SELECTEDITEM - 1)
            End If
        ElseIf KeyAscii = 27 Then  ' can be used if not enabled...to quit
            KeyAscii = 0
            SELECTEDITEM = -1
            secreset = False
             RaiseEvent Selected2(-2)
     Else
        If myEnabled Then
        If maxchar = 0 Or (maxchar > Len(List(SELECTEDITEM - 1)) Or MultiLineEditBox) Then
         RaiseEvent SyncKeyboard(KeyAscii)
         If KeyAscii > 31 And SELECTEDITEM > 0 Then
            If EditFlag Then
            bb = Enabled
            Enabled = False
            RaiseEvent PushUndoIfMarked
            RaiseEvent MarkDelete(False)
            Enabled = bb
            End If
            If EditFlag And KeyAscii > 32 And KeyAscii <> 127 Then
            If UKEY$ <> "" Then
            KK$ = UKEY$
            UKEY$ = ""
            Else
  KK$ = GetKeY(KeyAscii)
  End If
  
             RaiseEvent RemoveOne(KK$)
            If SelStart = 0 Then mSelstart = 1
           
            SelStartEventAlways = SelStart + 1
   
                List(SELECTEDITEM - 1) = Left$(List(SELECTEDITEM - 1), SelStart - 2) + KK$ + Mid$(List(SELECTEDITEM - 1), SelStart - 1)
            
         
            End If
         End If
         End If
         End If
    End If
End If
KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Timer1.Interval = 30
If Not Enabled Then Exit Sub
If listcount > 0 Then
  ShowMe2
Else
  ShowMe
End If
refresh

End Sub

Private Sub Timer2_Timer()
If drc Then
If topitem > 0 Then
topitem = topitem - 1
 SELECTEDITEM = topitem + 1

Timer1.Interval = 0
Timer1.Interval = 100
  Timer1.Enabled = True
 End If
Else
If topitem + 1 < listcount - lines Then
topitem = topitem + 1
 If topitem + lines + 1 <= listcount Then SELECTEDITEM = topitem + lines + 1
Timer1.Interval = 0
Timer1.Interval = 100
  Timer1.Enabled = True
  End If
End If
state = True
 On Error Resume Next
 Err.clear

    If SELECTEDITEM >= listcount Then
 Value = listcount - 1
  state = False
  Exit Sub
        Else
    Value = topitem
    End If
    state = False
 If Timer2.Enabled = False Then
If SELECTEDITEM - topitem > 0 And SELECTEDITEM - topitem - 1 <= lines And cX > 0 And cX < UserControl.ScaleWidth Then
 If SELECTEDITEM > 0 Then
         If Not BlockItemcount Then
             REALCUR List(SELECTEDITEM - 1), cX - scrollme, DUMMY, mSelstart, True
             mSelstart = mSelstart + 1
RaiseEvent ChangeSelStart(mSelstart)
             End If
 RaiseEvent selected(SELECTEDITEM)
 End If
 End If
 Else
 Timer3.Enabled = True
 End If
End Sub





Private Sub Timer3_Timer()
Timer3.Enabled = False
DOT3
End Sub
Private Sub DOT3()
If SELECTEDITEM > listcount Then
Timer3.Enabled = False
Exit Sub
End If
If SELECTEDITEM > 0 Then
' why???
'ShowMe2
RaiseEvent ScrollSelected(SELECTEDITEM, cY * myt)

End If
End Sub


Public Sub SoftEnterFocus()


FreeMouse = True
state = Not Enabled
Noflashingcaret = Not Enabled
If EditFlag Then
If Not Spinner Then state = Not MultiLineEditBox
End If
RaiseEvent ShowExternalCursor
If Not Timer1.Enabled Then PrepareToShow 5
End Sub

Private Sub SoftExitFocus()
If Not havefocus Then Exit Sub
Noflashingcaret = True
state = True ' no keyboard input

secreset = False
Timer2.Enabled = False
FreeMouse = False

If (Not BypassLeaveonChoose) And LeaveonChoose Then
If Not MultiLineEditBox Then If EditFlag And caretCreated Then caretCreated = False: DestroyCaret
SELECTEDITEM = -1: RaiseEvent Selected2(-2)
End If
If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar

RaiseEvent ShowExternalCursor
state = False
End Sub



Private Sub UserControl_Initialize()
Buffer = 100
Set m_font = UserControl.Font
ReDim mlist(0 To Buffer)
scrTwips = Screen.TwipsPerPixelX

DrawWidth = 1
DrawStyle = 0
NoPanLeft = True
NoPanRight = True
clear
maxchar = 50
WordCharLeft = " ,."
WordCharRight = " ,."
BarColor = &H63DFFE  '&HC3C3C3
Shape1.hatchType = 1
mlx = -1000
mly = -1000
End Sub

Private Sub UserControl_InitProperties()
 BackColor = m_def_BackColor
   ForeColor = m_def_ForeColor
    CapColor = m_def_CapColor
 dcolor = m_def_dcolor
mValue = 0
mmin = 0
mVertical = False
mjumptothemousemode = False
minimumWidth = 60
mLargeChange = 1
mSmallChange = 1
mmax = 100
mpercent = 0.07
NoPanLeft = True
NoPanRight = True

End Sub
Public Sub PressKey(KeyCode As Integer, shift As Integer, Optional NoEvents As Boolean = False)

If shift <> 0 And KeyCode = 16 Then Exit Sub

Timer1.Enabled = False
'Timer1.Interval = 1000
Dim lastlistindex As Long, bb As Boolean
lastlistindex = listindex
If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
If MarkNext = 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
End If
If KeyCode = 93 Then
' you have to clear myButton, here keycode
RaiseEvent OutPopUp(nowx, nowy, KeyCode)
End If
Select Case KeyCode
Case vbKeyHome
If EditFlag Then
mSelstart = 1
Else
ShowThis 1
       Do While Not (Not ListSep(listindex) Or listindex = listcount - 1)
    ShowThis SELECTEDITEM + 1
    Loop
    If ListSep(listindex) Then listindex = lastlistindex
    RaiseEvent ChangeSelStart(SelStart)
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
End If
Case vbKeyEnd
If EditFlag Then
mSelstart = Len(List(listindex)) + 1
Else
    ShowThis listcount
    Do While Not (Not ListSep(listindex) Or listindex = 0)
        ShowThis SELECTEDITEM - 1
    Loop
    If ListSep(listindex) Then listindex = lastlistindex
    RaiseEvent ChangeSelStart(SelStart)
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
End If
Case vbKeyPageUp
    If SELECTEDITEM - lines < 0 Then
       If SELECTEDITEM - 1 > 0 Then
       ShowThis SELECTEDITEM - 1
       Else
       PrepareToShow 5
           If shift <> 0 Then If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
           
       shift = 0: KeyCode = 0: Exit Sub
       End If
    Else
        If topitem < SELECTEDITEM - (lines + 1) \ 2 Then
            If topitem = 0 Then
                ShowThis SELECTEDITEM - 1
            Else
                ShowThis topitem
            End If
        Else
            ShowThis SELECTEDITEM - (lines + 1) \ 2 - 1
        End If
    End If
    While ListSep(listindex) And Not listindex = 0
        ShowThis SELECTEDITEM - 1
    Wend
    If ListSep(listindex) Then listindex = lastlistindex
    RaiseEvent ChangeSelStart(SelStart)
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
         If shift <> 0 Then If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)

     shift = 0: KeyCode = 0: Exit Sub
Case vbKeyUp
If Spinner Then Exit Sub
    Do
    
        ShowThis SELECTEDITEM - 1
        
    Loop Until Not ListSep(listindex) Or listindex = 0
    
    If ListSep(listindex) Then listindex = lastlistindex
' FIND RIGHT SELSTART...

    RaiseEvent ChangeSelStart(SelStart)
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
       If shift <> 0 Then
   ' KeyCode = 0
    PrepareToShow 5
    If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
    Else
      RaiseEvent MarkDestroyAny
    MarkNext = 0
    If NoFreeMoveUpDown Then ShowMe2: Exit Sub
    End If
      shift = 0: KeyCode = 0: Exit Sub
    
Case vbKeyDown
If Spinner Then Exit Sub
    Do
     
    ShowThis SELECTEDITEM + 1

    Loop Until Not ListSep(listindex) Or listindex = listcount - 1
    If ListSep(listindex) Then listindex = lastlistindex
    SelStartEventAlways = SelStart
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
    If shift <> 0 Then
    'KeyCode = 0
    PrepareToShow 5
  If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
    Else
    RaiseEvent MarkDestroyAny
    MarkNext = 0
    If NoFreeMoveUpDown Then ShowMe2: Exit Sub
    End If
  
      KeyCode = 0: Exit Sub
Case vbKeyPageDown
''RaiseEvent ScrollMove(topitem)
    If SELECTEDITEM + (lines + 1) \ 2 >= listcount Then
     'FindRealCursor SELECTEDITEM + 1
     If listcount > SELECTEDITEM Then
    ShowThis SELECTEDITEM + 1
    Else
     PrepareToShow 5
        If shift <> 0 Then If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
    shift = 0: KeyCode = 0: Exit Sub
    End If
    ElseIf (SELECTEDITEM - topitem) <= (lines + 1) \ 2 Then
    If topitem + (lines + 1) + 1 <= listcount Then
    ShowThis topitem + (lines + 1) + 1
    Else
    ShowThis SELECTEDITEM + 1
    End If
    Else
    ShowThis SELECTEDITEM + (lines + 1) \ 2 + 1
    End If
    While ListSep(listindex) And Not listindex = listcount - 1
    ShowThis SELECTEDITEM + 1
    Wend
    If ListSep(listindex) Then listindex = lastlistindex
    RaiseEvent ChangeSelStart(SelStart)
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
    If shift <> 0 Then If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
     shift = 0: KeyCode = 0: Exit Sub
Case vbKeySpace

If SELECTEDITEM > 0 Then
If EditFlag Then
If mSelstart = 0 Then mSelstart = 1
 If maxchar = 0 Or (maxchar > Len(List(SELECTEDITEM - 1)) Or MultiLineEditBox) Then
 bb = Enabled
 Enabled = False
     RaiseEvent PushUndoIfMarked
     RaiseEvent MarkDelete(False)
 Enabled = bb
List(SELECTEDITEM - 1) = Left$(List(SELECTEDITEM - 1), SelStart - 1) & " " & Mid$(List(SELECTEDITEM - 1), SelStart)
 RaiseEvent RemoveOne(" ")
SelStartEventAlways = SelStart + 1
KeyCode = 0
PrepareToShow 10
End If
Exit Sub
Else

If (MultiSelect Or ListMenu(SELECTEDITEM - 1)) Then
If ListRadio(SELECTEDITEM - 1) And ListSelected(SELECTEDITEM - 1) Then
' do nothing
Else
ListSelected(SELECTEDITEM - 1) = Not ListSelected(SELECTEDITEM - 1)
' from 1 to listcount
If MultiSelect Then
   If ListSelected(SELECTEDITEM - 1) Then
    RaiseEvent SelectedMultiAdd(SELECTEDITEM)
    Else
    RaiseEvent SelectedMultiSub(SELECTEDITEM)
    End If
Else
RaiseEvent MenuChecked(SELECTEDITEM)
End If
End If
End If
End If
End If
Case vbKeyLeft
If EditFlag Then

If MultiLineEditBox Then
If SelStart > 1 Then
mSelstart = SelStart - 1
RaiseEvent MayRefresh(bb)
If bb Then ShowMe2
ElseIf listindex > 0 Then
ShowThis SELECTEDITEM - 1
If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
mSelstart = Len(List(listindex)) + 1
End If
ElseIf SelStart > 1 Then
mSelstart = SelStart - 1
End If
End If
Case vbKeyRight
If EditFlag Then
If MultiLineEditBox Then
If SelStart <= Len(List(SELECTEDITEM - 1)) Then
mSelstart = SelStart + 1
RaiseEvent MayRefresh(bb)
If bb Then ShowMe2
ElseIf listindex < listcount - 1 Then
ListindexPrivateUse = listindex + 1
mSelstart = 0
If (SELECTEDITEM - topitem) > lines + 1 Then topitem = topitem + 1
If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)


End If
Else
If SelStart <= Len(List(SELECTEDITEM - 1)) Then mSelstart = SelStart + 1
End If
End If
Case vbKeyDelete
If EditFlag Then
If mSelstart = 0 Then mSelstart = 1
If SelStart > Len(List(SELECTEDITEM - 1)) Then
If listcount > SELECTEDITEM Then
If Not NoEvents Then

RaiseEvent LineDown
RaiseEvent addone(vbCrLf)
End If
End If
Else
 
 RaiseEvent addone(Mid$(List(SELECTEDITEM - 1), SelStart, 1))
List(SELECTEDITEM - 1) = Left$(List(SELECTEDITEM - 1), SelStart - 1) + Mid$(List(SELECTEDITEM - 1), SelStart + 1)
ShowMe2
End If
End If

Case vbKeyBack

If EditFlag Then
    If SelStart > 1 Then
        SelStart = SelStart - 1  ' make it a delete because we want selstart to take place before list() take value
     
        RaiseEvent addone(Mid$(List(SELECTEDITEM - 1), SelStart, 1))
      

        List(SELECTEDITEM - 1) = Left$(List(SELECTEDITEM - 1), SelStart - 1) + Mid$(List(SELECTEDITEM - 1), SelStart + 1)
        ShowMe2  'refresh now
    Else
        If mSelstart = 0 Then mSelstart = 1
        
        If Not NoEvents Then RaiseEvent LineUp
    End If
End If
Case vbKeyReturn
If MultiLineEditBox Then

RaiseEvent SplitLine
RaiseEvent RemoveOne(vbCrLf)
End If
End Select
If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
End If
KeyCode = 0
SelStartEventAlways = SelStart
Me.PrepareToShow 5
KeyCode = 0
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, shift As Integer)
Dim i As Long
If KeyCode = 18 Then
RaiseEvent Maybelanguage
ElseIf KeyCode = 16 And shift <> 0 Then
RaiseEvent Maybelanguage
ElseIf KeyCode = vbKeyV Then
Exit Sub
End If
i = -1
If shift <> 4 And mynum$ <> "" Then
On Error Resume Next
If Left$(mynum$, 1) = "0" Then
i = Val(mynum$)
Else
i = Val(mynum$)
End If
mynum$ = ""
Else
i = GetLastKeyPressed
End If

 If i <> -1 And i <> 94 Then
 UKEY$ = ChrW(i)
 Else
UKEY$ = ""
 End If
End Sub

Private Sub UserControl_LostFocus()

doubleclick = 0
If Not NoWheel Then RaiseEvent UnregisterGlist
RaiseEvent CheckLostFocus
If myEnabled Then SoftExitFocus
havefocus = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
' cut area
If dropkey Then Exit Sub
nowx = x
nowy = y

If (Button And 2) = 2 Then Exit Sub
If myt = 0 Then Exit Sub
FreeMouse = True
Dim YYT As Long, oldbutton As Integer
If mHeadlineHeightTwips = 0 Then
YYT = y \ myt
Else
    If y < mHeadlineHeightTwips Then
        If y < 0 Then
        YYT = -1
        Else
        YYT = 0
        End If
    Else
    YYT = (y - mHeadlineHeightTwips) \ myt + 1
    End If
End If
If YYT < 0 Then YYT = 0
If (YYT >= 0 And (YYT < listcount Or listcount = 0) And myEnabled) Then

oldbutton = Button

If mHeadline <> "" And Timer2.Enabled = False Then
    If YYT = 0 Then ' we move in mHeadline
        ' -1 is mHeadline
        ' headline listen clicks if  list is disabled...
        RaiseEvent ExposeItemMouseMove(Button, -1, CLng(x) / scrTwips, CLng(y) / scrTwips)
        If (x < Width - mPreserveNpixelsHeaderRight) Or (mPreserveNpixelsHeaderRight = 0) Then RaiseEvent HeaderSelected(Button)
        If oldbutton <> Button Then
        Button = 0
        Exit Sub
        End If
    ElseIf myEnabled Then
        RaiseEvent ExposeItemMouseMove(Button, topitem + YYT - 1, CLng(x) / scrTwips, CLng(y - (YYT - 1) * myt - mHeadlineHeightTwips) / scrTwips)
    End If
ElseIf myEnabled Then
    RaiseEvent ExposeItemMouseMove(Button, topitem + YYT, CLng(x) / scrTwips, CLng(y - YYT * myt) / scrTwips)
End If
If oldbutton <> Button Then Exit Sub
End If
YYT = YYT + (mHeadline <> "")
lastX = x
LastY = y

If (x > Width - barwidth) And BarVisible And EnabledBar And Button = 1 Then
If Vertical Then
GetOpenValue = valuepoint - y + mHeadlineHeightTwips
Else
'GetOpenValue = valuepoint - x ' NOT USED HERE
End If
If processXY(lastX, LastY, False) And myEnabled Then
FreeMouse = False
End If
Timer3.Enabled = False
Else
cX = x
If Not dr Then ly = x
dr = True

If cY = y \ myt Then Timer3.Enabled = False: cY = y \ myt
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If dropkey Then Exit Sub
Static PX As Long, PY As Long

If Abs(PX - x) <= 60 And Abs(PY - y) <= 60 Then Exit Sub
PX = x
PY = y

RaiseEvent MouseMove(Button, shift, x, y)
If myt = 0 Or Not myEnabled Then Exit Sub
If (Button And 2) = 2 Then Exit Sub
Dim tListcount As Long

tListcount = listcount
Static TIMESTAMP As Double

If TIMESTAMP = 0 Or (TIMESTAMP - Timer) > 1 Then TIMESTAMP = Timer

If (TIMESTAMP + 0.02) > Timer And shift = 0 Then Exit Sub
TIMESTAMP = Timer


If Not FreeMouse Then Exit Sub
If Button = 0 Then If mousepointer < 2 Then mousepointer = 1
If (x > Width - barwidth) And tListcount > lines + 1 And Not BarVisible Then
Hidebar = True: BarVisible = True
ElseIf (x < Width - barwidth) And Button = 0 And BarVisible And (StickBar Or AutoHide) Then
Hidebar = False
BarVisible = False
End If
If OurDraw Then
barMouseMove Button, shift, x, y
Exit Sub
End If
cX = x
Timer3.Enabled = False
Dim YYT As Long, oldbutton As Integer
If mHeadlineHeightTwips = 0 Then
YYT = y \ myt
Else
    If y < mHeadlineHeightTwips Then
        If y < 0 Then
        YYT = -1
        Else
        YYT = 0
        End If
    Else
    YYT = (y - mHeadlineHeightTwips) \ myt + 1
    End If
End If
oldbutton = Button
If (Button And 3) > 0 And useFloatList And FloatList Then FloatListMe useFloatList, x, y: Button = 0 Else If mousepointer > 1 Then mousepointer = 1
If mHeadline <> "" Then
If YYT = 0 Then ' we move in mHeadline
' -1 is mHeadline

If (Button And 3) > 0 And FloatList And Not useFloatList Then FloatListMe useFloatList, x, y: Button = 0
RaiseEvent ExposeItemMouseMove(Button, -1, CLng(x) / scrTwips, CLng(y) / scrTwips)
Else
RaiseEvent ExposeItemMouseMove(Button, topitem + YYT - 1, CLng(x) / scrTwips, CLng(y - (YYT - 1) * myt) / scrTwips)
End If
Else
RaiseEvent ExposeItemMouseMove(Button, topitem + YYT, CLng(x) / scrTwips, CLng(y - YYT * myt) / scrTwips)
End If
If oldbutton <> Button Then Exit Sub
YYT = YYT + (mHeadline <> "")
If (Button And 3) = 0 Then

If YYT >= 0 And YYT <= lines Then
If topitem + YYT < tListcount Then
secreset = False
End If
End If
ElseIf dr Then

     If MultiLineEditBox And (Button = 1) And secreset Then
            If MarkNext > 3 Then
       
                ElseIf MarkNext = 0 Then
                MarkNext = 1
                RaiseEvent MarkIn
                End If
     End If
If (SELECTEDITEM <> (topitem + YYT + 1)) And SELECTEDITEM >= 0 And Button <> 0 Then secreset = False
' special for M2000  (StickBar And x > Width / 2)
If shift = 0 And ((Not scrollme > 0) And (x > Width / 2) Or Not SingleLineSlide) And StickBar And MarkNext = 0 And tListcount > lines + 1 Then
If Abs(LastY - y) < scrTwips * 2 Then LastY = y: Exit Sub
Hidebar = True
CalcAndShowBar1
   If LastY < y Then
      y = scrTwips * 2
      Else
      y = ScaleHeight - scrTwips
      End If
     
           
            If Abs(lastX - x) < scrTwips * 4 Or Not MultiLineEditBox Then
             lastX = x
            LastY = y
            
            If Vertical Then
 
            GetOpenValue = valuepoint - y + mHeadlineHeightTwips
    
            Else
          '  GetOpenValue = valuepoint - x ' NO USED HERE
            End If
       

         
            If processXY(lastX, LastY, True) Then
            FreeMouse = False
            End If
            Timer3.Enabled = False
            Exit Sub
            Else
          
            If YYT >= 0 And YYT <= lines Then shift = 1: GoTo there1
            End If
            
End If
If mHeadline <> "" And y < mHeadlineHeightTwips Then
' we sent twips not pixels
' move...me..??

ElseIf (y - mHeadlineHeightTwips) < myt / 2 And (topitem + YYT > 0) Then
'scroll up


drc = True
 Timer2.Enabled = True
 
ElseIf y > ScaleHeight - myt \ 2 And (tListcount <> 1) Then

drc = False
 Timer2.Enabled = True
ElseIf YYT >= 0 And YYT <= lines Then
there1:

                If MultiLineEditBox And (Button = 1) Then
                If MarkNext = 1 Then
                shift = 1

                RaiseEvent MarkOut
                ElseIf shift = 0 And MarkNext = 2 Then
                MarkNext = 0  ' so markNext=2 we have a complete marked text
                RaiseEvent MarkDestroy
                End If
                End If
If Timer2.Enabled Then
 Timer2.Enabled = False
 
End If
If topitem + YYT < tListcount Then

If (cX > ScaleWidth / 4 And cX < ScaleWidth * 3 / 4) And scrollme = 0 Then x = ly

        If Not SELECTEDITEM = topitem + YYT + 1 Then
            
            SELECTEDITEM = topitem + YYT + 1
            
             If Not BlockItemcount Then
             REALCUR List(SELECTEDITEM - 1), cX - scrollme, DUMMY, mSelstart, True
      
              mSelstart = mSelstart + 1
              
                RaiseEvent ChangeSelStart(mSelstart)
            End If
 If MultiLineEditBox And (Button = 1) Then
                If shift = 1 And MarkNext = 0 Then
                MarkNext = 1
                RaiseEvent MarkIn
                ElseIf shift = 1 And MarkNext = 1 Then
                
                RaiseEvent MarkOut
                End If
     End If
      
            If StickBar Or AutoHide Then DOT3
            
            If x - ly > 0 And Not NoPanRight Then
            scrollme = (x - ly)
            ElseIf x - ly < 0 And Not NoPanLeft Then
             scrollme = (x - ly)
            Else
            If Not EditFlag Then scrollme = 0
            End If
         'Timer1.Enabled = True
         If Not EditFlag Then If scrollme > 0 Then scrollme = 0
          '
           
        ElseIf cY <> YYT Then
            cY = YYT
            Timer3.Enabled = True
        Else
         If Not BlockItemcount Then
             REALCUR List(SELECTEDITEM - 1), cX - scrollme, DUMMY, mSelstart, True
              mSelstart = mSelstart + 1
              ' maybe this can change
RaiseEvent ChangeSelStart(mSelstart)
             End If
              If MultiLineEditBox And (Button = 1) Then
                  If shift = 1 And MarkNext = 0 Then
                      MarkNext = 1
                      RaiseEvent MarkIn
                            ElseIf shift = 1 And MarkNext = 1 Then
                                RaiseEvent MarkOut
                End If
                End If
               If x - ly > 0 And Not NoPanRight Then
            scrollme = (x - ly)
            ElseIf x - ly < 0 And Not NoPanLeft Then
             scrollme = (x - ly)
            Else
          If Not EditFlag Then scrollme = 0
            End If
            
         '   If scrollme > 0 Then scrollme = 0
            'Timer3.Enabled = False
            Timer1.Interval = 20
            Timer1.Enabled = True
            Timer3.Enabled = False
        End If

End If
End If
End If

End Sub
Public Sub CheckMark()
' if shift =0
    If MarkNext >= 1 Then
    If MarkNext < 4 Then
                MarkNext = 0  ' so markNext=2 we have a complete marked text
                RaiseEvent MarkDestroy
                ShowMe2
                Else
                MarkNext = MarkNext - 1
                End If
      End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
'If Not myt = mytPixels * SCRTWIPS Then Stop
If dropkey Then Exit Sub
If Button = 1 Then mlx = CLng(x / scrTwips): mly = CLng(y / scrTwips): RaiseEvent MouseUp(x / scrTwips, y / scrTwips)
If (Button And 2) = 2 Then
x = nowx
y = nowy
End If
useFloatList = False
If myt = 0 Then Exit Sub
Timer1bar.Interval = 100
Timer1bar.Enabled = False
If OurDraw Then
OurDraw = False
Exit Sub
End If
Timer2.Enabled = False
If Not (FreeMouse Or Not myEnabled) Then Exit Sub

With UserControl
 If (x < 0 Or y < 0 Or x > .Width Or y > .Height) And (LeaveonChoose And Not BypassLeaveonChoose) Then
If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar
 SELECTEDITEM = -1
 RaiseEvent Selected2(-2)
 Exit Sub
 End If
End With
cX = x
If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar
If Timer3.Enabled Then cY = y: DOT3
Timer3.Enabled = False
If Timer2.Enabled Then
 Timer2.Enabled = False
 End If
Dim YYT As Long
  If dr Then
                    ly = 0
            
                    If scrollme < -myt Then
                        RaiseEvent PanLeftRight(False)
                    ElseIf scrollme > myt Then
                        RaiseEvent PanLeftRight(True)
                    Else
                    dr = False
                    GoTo jump1
                    End If
                 If Not EditFlag Then scrollme = 0
                    Timer1.Enabled = True
                    dr = False
                End If
jump1:
If mHeadlineHeightTwips = 0 Then
YYT = y \ myt
Else
    If y < mHeadlineHeightTwips Then
        If y < 0 Then
        YYT = -1
        Else
        YYT = 0
        End If
    Else
    YYT = (y - mHeadlineHeightTwips) \ myt + 1
    End If
End If


If YYT = -1 Then Button = 0
If mHeadline <> "" And YYT = 0 Then Button = 0
YYT = YYT + (mHeadline <> "")

If YYT >= 0 And YYT <= lines Then


If topitem + YYT < listcount Then

If (Button And 3) > 0 And myEnabled Then


    If secreset Then
        ' this is a double click
        secreset = False
         If Not ListSep(topitem + YYT) Then
         If MarkNext = 0 And EditFlag Then
         
      MarkWord
      
      Else
      RaiseEvent Selected2(SELECTEDITEM - 1)
      Exit Sub
                End If
        
        
        End If
        
    Else
        Timer1.Enabled = False
        If (((SELECTEDITEM <> (topitem + YYT + 1)) And Not secreset) Or EditFlag) And Not ListSep(topitem + YYT) Then
             SELECTEDITEM = topitem + YYT + 1 ' we have a new selected item
             ' compute selstart always
             If Not BlockItemcount Then
             
             REALCUR List(SELECTEDITEM - 1), cX - scrollme, DUMMY, mSelstart, True
              mSelstart = mSelstart + 1
RaiseEvent ChangeSelStart(mSelstart)

             End If
              RaiseEvent selected(SELECTEDITEM)  ' broadcast
              
         End If
    '     If Shift = 0 Then CheckMark

         If SELECTEDITEM = topitem + YYT + 1 Then
                        If MultiSelect Or ListMenu(SELECTEDITEM - 1) Then
                            If (x / scrTwips > 0) And (x / scrTwips < LeftMarginPixels) Then
                                If ListRadio(SELECTEDITEM - 1) And ListSelected(SELECTEDITEM - 1) Then
                                ' do nothing
                                Else
                                ListSelected(SELECTEDITEM - 1) = Not ListSelected(SELECTEDITEM - 1)
                                If MultiSelect Then
                                If ListSelected(SELECTEDITEM - 1) Then
                                RaiseEvent SelectedMultiAdd(SELECTEDITEM)
                                Else
                                RaiseEvent SelectedMultiSub(SELECTEDITEM)
                                End If
                                Else
                                RaiseEvent MenuChecked(SELECTEDITEM)
                                End If
                                End If
                            End If
                            
                        End If

End If

End If
If secreset = False Then If shift = 0 Then CheckMark
If Not Enabled Then Exit Sub
secreset = True
ShowMe2
 If Button = 2 Then
RaiseEvent OutPopUp(x, y, Button)

End If
''
End If
'End If

End If
End If
End Sub



Private Sub UserControl_OLECompleteDrag(Effect As Long)
If Effect = 0 Then
' CANCEL...
If marvel Then
RaiseEvent MarkDestroy
ShowMe2
End If
ElseIf Effect = vbDropEffectMove Then
If marvel Then
RaiseEvent PushUndoIfMarked
RaiseEvent MarkDelete(False)
End If
End If
Effect = 0
End Sub

Private Sub UserControl_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single)
Dim something$, ok As Boolean
If dropkey Then Exit Sub
If (Effect And 3) > 0 Then
If data.GetFormat(vbCFText) Or data.GetFormat((13)) Then

If (Button And 1) = 0 Then
    If (shift And 2) = 2 Then
        Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectMove
            End If
        End If
End If
End If
RaiseEvent DropOk(ok)
If marvel Then

Else
RaiseEvent MarkDestroyAny
ok = True
End If
If ok Then
        If data.GetFormat(13) Then
          
          something$ = data.GetData(13)
          Else
        
            something$ = data.GetData(vbCFText)
            End If
something$ = Replace(something$, ChrW(0), "")

If marvel Then
RaiseEvent DropFront(ok)
If ok Then
RaiseEvent selected(SELECTEDITEM)

    RaiseEvent DragPasteData(something$)
 
   If Effect = vbDropEffectMove Then
 RaiseEvent addone(something$)
 
   RaiseEvent MarkDelete(True)
    RaiseEvent RemoveOne("")
    Else

        RaiseEvent MarkDestroyAny
    End If
Else
If Effect = vbDropEffectMove Then
    RaiseEvent addone(something$)
    RaiseEvent PushMark2Undo(something$)
    RaiseEvent MarkDelete(True)
    
Else
    RaiseEvent MarkDestroyAny
End If
    RaiseEvent selected(SELECTEDITEM)
    RaiseEvent DragPasteData(something$)
    
End If
Else
RaiseEvent selected(SELECTEDITEM)
RaiseEvent DragPasteData(something$)

End If
marvel = False



Else
Effect = 0
End If

End Sub

Private Sub UserControl_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single, state As Integer)
If dropkey Then Exit Sub
If Not DropEnabled Then Effect = 0: Exit Sub
Dim tListcount As Long, YYT As Long
  If TaskMaster.QueueCount > 0 Then
              TaskMaster.RestEnd1
   TaskMaster.TimerTick
TaskMaster.rest
        End If
tListcount = listcount
 If state = vbOver Then
 
If mHeadline <> "" And y < mHeadlineHeightTwips Then
' we sent twips not pixels
' move...me..??

ElseIf (y - mHeadlineHeightTwips) < myt / 2 And (topitem + YYT > 0) Then
                drc = True
                Timer2.Enabled = True
        
        ElseIf y > ScaleHeight - myt \ 2 And (tListcount <> 1) Then
                drc = False
                Timer2.Enabled = True
        Else
                Timer2.Enabled = False
             '  If marvel Then
                                 MovePos x, y
                              
                              
                               
              '  End If
            If data.GetFormat(vbCFText) Or data.GetFormat((13)) Then
                        If (shift And 2) = 2 Then
                            Effect = vbDropEffectCopy
                        Else
                            Effect = vbDropEffectMove
                        End If
                Else
                    Effect = vbDropEffectNone
            End If
            End If
ElseIf state = vbLeave Then
        Timer3.Enabled = True
        Effect = vbDropEffectNone
        
ElseIf state = vbEnter Then
       
        
     MovePos x, y
   ShowMe2
                             
                               
        End If
        
             If data.GetFormat(vbCFText) Or data.GetFormat((13)) Then
                    If (shift And 2) = 2 Then
                       Effect = vbDropEffectCopy
                       Else
                           Effect = vbDropEffectMove
                           End If
            Else
                Effect = vbDropEffectNone
        End If
      
End Sub



Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
        If TaskMaster.QueueCount > 0 Then
              TaskMaster.RestEnd1
   TaskMaster.TimerTick
TaskMaster.rest
        End If
End Sub

Private Sub UserControl_OLEStartDrag(data As DataObject, AllowedEffects As Long)
If dropkey Then Exit Sub
If Not DragEnabled Then Exit Sub
Dim aa() As Byte, this$
RaiseEvent DragData(this$)
aa = this$ & ChrW$(0)
 data.SetData aa(), 13
data.SetData aa(), vbCFText
 AllowedEffects = vbDropEffectCopy + vbDropEffectMove
 
End Sub
Public Sub MovePos(ByVal x As Single, ByVal y As Single)
Dim DUMMY As Long, YYT As Long, M_CURSOR As Long

If mHeadlineHeightTwips = 0 Then
YYT = y \ myt + 1
Else
    If y < mHeadlineHeightTwips Then
        If y < 0 Then
        Exit Sub
        Else
        YYT = 1
        End If
    Else
    YYT = (y - mHeadlineHeightTwips) \ myt + 1
    End If
End If
YYT = YYT - 1
If topitem + YYT < listcount Then
REALCUR List(topitem + YYT), x - scrollme, DUMMY, M_CURSOR
ListindexPrivateUse = topitem + YYT
If listindex = -1 Then
        If itemcount = 0 Then
        additemFast ""
        End If
        ListindexPrivateUse = 0

End If
SelStart = M_CURSOR + 1

Else
ListindexPrivateUse = listcount - 1
            If listindex = -1 Then
            If itemcount = 0 Then
            additemFast ""
            End If
            ListindexPrivateUse = 0
            
            End If
SelStart = Len(List(listindex)) + 1

End If
RaiseEvent selected(SELECTEDITEM)
RaiseEvent ChangeSelStart(SelStart)
ExternalCursor SelStart, List(listindex)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 m_sync = .ReadProperty("sync", m_def_sync)
 NoFire = True
Value = .ReadProperty("Value", 0)
Max = .ReadProperty("Max", 100)
Min = .ReadProperty("Min", 0)
largechange = .ReadProperty("LargeChange", 1)
smallchange = .ReadProperty("SmallChange", 1)
Percent = .ReadProperty("Percent", 0.07)
Vertical = .ReadProperty("Vertical", False)
jumptothemousemode = .ReadProperty("JumptoTheMouseMode", False)
NoFire = False
 Set Font = .ReadProperty("Font", Ambient.Font)
 
   
    myEnabled = .ReadProperty("Enabled", m_def_Enabled)
 
 
    BackStyle = .ReadProperty("BackStyle", m_def_BackStyle)
  BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
      
   m_showbar = .ReadProperty("ShowBar", m_def_Showbar)
    dcolor = .ReadProperty("dcolor", m_def_dcolor)
      BackColor = .ReadProperty("BackColor", m_def_BackColor)
    ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
      CapColor = .ReadProperty("CapColor", m_def_CapColor)

   Text = .ReadProperty("Text", m_def_Text)

   End With
   If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
Else

myt = UserControlTextHeight() + addpixels * scrTwips
End If
HeadlineHeight = UserControlTextHeight() / scrTwips
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
waitforparent = True
End Sub



Private Sub UserControl_Show()
If Not design() Then
'CalcAndShowBar
fast = True
SoftEnterFocus

End If
End Sub

Private Sub UserControl_Terminate()
If LastGlist Is Me Then Set LastGlist = Nothing
waitforparent = True
Set m_font = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

With PropBag
     .WriteProperty "sync", m_sync, m_def_sync
    .WriteProperty "Value", Value, 0
    .WriteProperty "Max", Max, 100
    .WriteProperty "Min", Min, 0
    .WriteProperty "LargeChange", largechange, 1
    .WriteProperty "SmallChange", smallchange, 1
    .WriteProperty "Percent", Percent, 0.07
    .WriteProperty "Vertical", Vertical, False
    .WriteProperty "JumptoTheMouseMode", jumptothemousemode, False
    .WriteProperty "Font", m_font, Ambient.Font
    .WriteProperty "Enabled", myEnabled, m_def_Enabled
    .WriteProperty "BackStyle", m_BackStyle, m_def_BackStyle
   .WriteProperty "BorderStyle", m_BorderStyle, m_def_BorderStyle
    .WriteProperty "ShowBar", m_showbar, m_def_Showbar
    .WriteProperty "dcolor", dcolor, m_def_dcolor
     .WriteProperty "Backcolor", BackColor, m_def_BackColor
       .WriteProperty "ForeColor", ForeColor, m_def_ForeColor
    .WriteProperty "CapColor", CapColor, m_def_CapColor

      .WriteProperty "Text", Text, ""
      End With

End Sub
Property Get listindex() As Long
If SELECTEDITEM < 0 Then
listindex = -1  ' CHANGED
Else
listindex = SELECTEDITEM - 1
End If
End Property
Property Let listindex(item As Long)
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Property
If Err.Number > 0 Then Exit Property
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
If listcount <= lines + 1 Then
BarVisible = False
Else
Redraw m_showbar And mo.Visible

End If


If item < listcount Then SELECTEDITEM = item + 1
If SELECTEDITEM > 0 Then
RaiseEvent softSelected(SELECTEDITEM)
Else
SELECTEDITEM = 0
End If
End Property
Private Sub FloatListMe(state As Boolean, x As Single, y As Single)
Static preX As Single, preY As Single
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Sub
If Err.Number > 0 Then Exit Sub
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
If Not state Then
preX = x
preY = y
state = True
mousepointer = 0
doubleclick = 0
Else
If mo.Visible Then
mousepointer = 5
If MoveParent Then
If (mo.Parent.top + (y - preY) < 0) Then preY = y + mo.Parent.top
If (mo.Parent.Left + (x - preX) < 0) Then preX = x + mo.Parent.Left
If ((mo.Parent.top + y - preY) > FloatLimitTop) And FloatLimitTop > 0 Then preY = mo.Parent.top + y - FloatLimitTop

If ((mo.Parent.Left + x - preX) > FloatLimitLeft) And FloatLimitLeft > 0 Then preX = mo.Parent.Left + x - FloatLimitLeft
mo.Parent.Move mo.Parent.Left + (x - preX), mo.Parent.top + (y - preY)
Else
mo.ZOrder
If (mo.top + (y - preY) < 0) Then preY = y + mo.top
If (mo.Left + (x - preX) < 0) Then preX = x + mo.Left
If ((mo.top + y - preY) > FloatLimitTop) And FloatLimitTop > 0 Then preY = mo.top + y - FloatLimitTop

If ((mo.Left + x - preX) > FloatLimitLeft) And FloatLimitLeft > 0 Then preX = mo.Left + x - FloatLimitLeft

mo.Move mo.Left + (x - preX), mo.top + (y - preY)
End If
End If
If Me.BackStyle = 1 Then ShowMe2
End If
End Sub
Property Get List(item As Long) As String
Dim that$
If itemcount = 0 Or BlockItemcount Then
RaiseEvent ReadListItem(item, that$)
List = that$
Exit Property
End If
If item < 0 Then Exit Property
If item >= listcount Then
Err.Raise vbObjectError + 1050
Else
List = mlist(item).content
End If
End Property

Property Let List(item As Long, ByVal b$)
On Error GoTo nnn1
If itemcount = 0 Or BlockItemcount Then
If List(item) <> b$ Then RaiseEvent ChangeListItem(item, b$)
Exit Property
End If
If item >= 0 Then
With mlist(item)
If Not .content = b$ Then RaiseEvent ChangeListItem(item, b$)
.content = b$
.line = False
.selected = False
End With
End If
nnn1:
End Property
Property Let menuEnabled(item As Long, ByVal rhs As Boolean)
If item >= 0 Then
With mlist(item)
.line = Not rhs   ' The line flag used as enabled flag, in reverse logic
End With
End If
End Property
Property Let ListSep(item As Long, ByVal rhs As Boolean)
If item >= 0 Then
With mlist(item)
.line = rhs
End With
End If
End Property
Property Get ListSep(item As Long) As Boolean
Dim skip As Boolean, blockit As Boolean
RaiseEvent BlockCaret(item, blockit, skip)
If skip Then
ListSep = blockit
Exit Property
End If
If itemcount = 0 Or BlockItemcount Then Exit Property
If item >= 0 Then
With mlist(item)
ListSep = .line
End With
End If
End Property

Property Let ListSelected(item As Long, ByVal b As Boolean)
Dim first As Long, last As Long
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then

If mlist(item).radiobutton Then
        ' erase all
        first = item
        While first > 0 And mlist(first).radiobutton
        first = first - 1
        Wend
        If Not mlist(first).radiobutton Then first = first + 1
        last = item
        While last < listcount - 1 And mlist(last).radiobutton
        last = last + 1
        Wend
        If Not mlist(last).radiobutton Then last = last - 1
        For first = first To last
        mlist(first).selected = False
        Next first
End If
With mlist(item)
.selected = b
End With
End If
End If
End Property
Property Let ListSelectedNoRadioCare(item As Long, ByVal b As Boolean)
Dim first As Long, last As Long
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then

With mlist(item)
.selected = b
End With
End If
End If
End Property
Property Get ListSelected(item As Long) As Boolean

If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mlist(item)
ListSelected = .selected
End With
End If
End If
End Property
Property Let ListRadio(item As Long, ByVal b As Boolean)
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mlist(item)
.radiobutton = b
End With
End If
End If
End Property
Property Get ListRadio(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mlist(item)
ListRadio = .radiobutton
End With
End If
End If
End Property
Property Get ListMenu(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mlist(item)
ListMenu = .radiobutton Or .Checked
End With
End If
End If
End Property
Property Let ListChecked(item As Long, ByVal b As Boolean)
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mlist(item)
.Checked = b
End With
End If
End If
End Property
Property Get ListChecked(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mlist(item)
ListChecked = .Checked
End With
End If
End If
End Property
Public Sub moveto(ByVal key As String)
Dim i As Long
For i = 0 To listcount - 1
If List(i) Like key Then Exit For
Next i
If i < listcount Then
listindex = i
End If
End Sub
Public Function FindItemStartWidth(ByVal key As String, NoCase As Boolean, ByVal offset) As Long
Dim i As Long, j As Long
j = Len(key)
i = -1
FindItemStartWidth = -1
If j = 0 Then Exit Function
If NoCase Then
For i = offset To listcount - 1
If StrComp(Left$(List(i), j), key, vbTextCompare) = 0 Then Exit For
Next i
Else
For i = offset To listcount - 1
If StrComp(Left$(List(i), j), key, vbBinaryCompare) = 0 Then Exit For
Next i
End If
If i < listcount Then
FindItemStartWidth = i
End If
End Function
Public Function find(ByVal key As String) As Long
Dim i As Long, skipme As Boolean
i = -1
RaiseEvent find(key, i, skipme)
If skipme Then find = i: Exit Function
find = -1
For i = 0 To listcount - 1
If List(i) Like key Then Exit For
Next i
If i < listcount Then
find = i
End If
End Function
Public Sub ShowThis(ByVal item As Long, Optional noselect As Boolean)
Dim MM$, mo As Control, nm$, cnt$, p As Long
On Error GoTo skipthis
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Sub
If Err.Number > 0 Then Exit Sub
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If

If listcount <= lines + 1 Then
BarVisible = False
Else
BarVisible = m_showbar And mo.Visible

End If
If item > 0 And item <= listcount Then
If MultiLineEditBox Then FindRealCursor item
    If item - topitem > 0 And item - topitem <= lines + 1 Then
    
        SELECTEDITEM = item
        
            If SELECTEDITEM = listcount Then
            state = True
            Value = Max
            state = False
            End If
            
        
    Else
    If item < lines / 2 Then topitem = 0 Else topitem = item - lines / 2

CalcAndShowBar1
        SELECTEDITEM = item
        ''ShowMe
     End If
   If Not noselect Then If Not Timer1.Enabled Then PrepareToShow 10
End If
If noselect Then
SELECTEDITEM = 0: ShowMe2

 
  End If
skipthis:
End Sub
Public Sub RepaintScrollBar()
If m_showbar Or StickBar Or AutoHide Or Shape1.Visible Or Spinner Then Redraw
If Not BarVisible Then refresh
End Sub
Public Sub clear(Optional ByVal interface As Boolean = False)
SELECTEDITEM = -1
LastSelected = -2
itemcount = 0
If hWnd <> 0 Then HideCaret (hWnd)
state = True
mValue = 0  ' if here you have an error you forget to apply VALUE as default property
showshapes
LastVScroll = 1

'max = 0
state = False
topitem = 0
Buffer = 100
ReDim mlist(0 To Buffer)

If interface Then
 '   barvisible = False
    ShowMe
End If
End Sub
Public Sub ClearClick()
SELECTEDITEM = -1
secreset = False
End Sub

Public Function DblClick() As Boolean
DblClick = secreset
secreset = False
End Function


Private Sub UserControl_Resize()
'If Not design() Then
CalcAndShowBar
'End If
End Sub
Public Sub additem(A$)
Dim i As Long

If itemcount = Buffer Then
Buffer = Buffer * 2
ReDim Preserve mlist(0 To Buffer)
End If
itemcount = itemcount + 1
With mlist(itemcount - 1)
.content = A$
.line = False
.selected = False
End With
Timer1.Enabled = False
Timer1.Interval = 100
Timer1.Enabled = True
End Sub
Public Sub addsep()
Dim i As Long

If itemcount = Buffer Then
Buffer = Buffer * 2
ReDim Preserve mlist(0 To Buffer)
End If
itemcount = itemcount + 1
ListSep(itemcount - 1) = True
Timer1.Enabled = False
Timer1.Interval = 100
Timer1.Enabled = True
End Sub
Public Sub additemFast(A$)
Dim i As Long
If itemcount = Buffer Then
Buffer = Buffer * 2
ReDim Preserve mlist(0 To Buffer)
End If
itemcount = itemcount + 1
With mlist(itemcount - 1)
.content = A$
.line = False
.selected = False
End With
End Sub
Public Sub Removeitem(ByVal ii As Long)
Dim i As Long
If ii = itemcount - 1 Then
Else
For i = ii + 1 To itemcount - 1
mlist(i - 1).content = mlist(i).content
mlist(i - 1).line = mlist(i).line
mlist(i - 1).selected = mlist(i).selected

Next i

End If
itemcount = itemcount - 1

If listcount < 0 Then
itemcount = 0
clear
Exit Sub
End If
If itemcount < Buffer \ 2 And Buffer > 100 Then
Buffer = Buffer \ 2
ReDim Preserve mlist(0 To Buffer)
End If
SELECTEDITEM = 0
If listcount <= lines + 1 Then
BarVisible = False
Else
'If BorderStyle = 0 Then
'LargeBar1.Move Width - barwidth, 0, barwidth, Height
'Else
'LargeBar1.Move ScaleWidth + SCRTWIPS - barwidth, 0, barwidth, ScaleHeight
'End If
' check this
''' Redraw m_showbar
End If
Timer1.Enabled = True


End Sub
Public Sub ShowMe(Optional visibleme As Boolean = False)
 Dim REALX As Long, REALX2 As Long, myt1
If visibleme Then
CalcAndShowBar1
Timer1.Enabled = True: Exit Sub
End If
If listcount = 0 And HeadLine = "" Then
    Repaint
    Exit Sub
End If

Dim i As Long, j As Long, g$, nr As RECT, fg As Long, hnr As RECT, skipme As Boolean, nfg As Long
If MultiSelect And LeftMarginPixels < mytPixels Then LeftMarginPixels = mytPixels
Repaint
CurrentY = 0
nr.top = 0
nr.Left = 0
nr.Bottom = mytPixels + 1
hnr.Bottom = mytPixels + 1

nr.Right = Width / scrTwips
hnr.Right = Width / scrTwips

If mHeadline <> "" Then
nr.Bottom = HeadlineHeight

RaiseEvent ExposeRect(-1, VarPtr(nr), UserControl.hDC, skipme)
nr.Bottom = HeadlineHeight
CalcRectHeader UserControl.hDC, mHeadline, hnr, DT_CENTER
If Not skipme Then
If hnr.Bottom < HeadLineHeightMinimum Then
hnr.Bottom = HeadLineHeightMinimum
End If


If mHeadlineHeight <> hnr.Bottom Then
HeadlineHeight = hnr.Bottom
nr.Bottom = mHeadlineHeight
End If
FillBack UserControl.hDC, nr, CapColor
End If
hnr.top = (nr.Bottom - hnr.Bottom) \ 2
hnr.Bottom = nr.Bottom - hnr.top
hnr.Left = 0
hnr.Right = nr.Right
PrintLineControlHeader UserControl.hDC, mHeadline, hnr, DT_CENTER

     nr.top = nr.Bottom
nr.Bottom = nr.top + mytPixels + 1
End If
If AutoPanPos Then

If SelStart = 0 Then SelStart = 1
scrollme = 0
again123:

REALX = UserControlTextWidth(Mid$(List(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
REALX2 = scrollme + REALX
If Not NoScroll Then If REALX2 > Width * 0.8 Then scrollme = scrollme - Width * 0.2:  GoTo again123
If REALX2 < 0 Then
If Not NoScroll Then scrollme = scrollme + Width * 0.2: GoTo again123

End If
End If
If SingleLineSlide Then
nr.Left = LeftMarginPixels
Else
nr.Left = scrollme / scrTwips + LeftMarginPixels
End If
If listcount = 0 Then
BarVisible = False
Exit Sub
End If
If SELECTEDITEM > 0 Then
topitem = 0

       j = SELECTEDITEM - lines / 2 - 1

    If j < 0 Then j = 0
    If listcount <= lines + 1 Then
       topitem = 0
    Else
    If j + lines > listcount Then
    If listcount - lines >= 0 Then
    topitem = listcount - lines - 1
    End If
    Else
        topitem = j
    End If
        state = True
            On Error Resume Next
            Err.clear
    If Not Spinner Then
            If listcount - 1 - lines < 1 Then
            Max = 1
            Else
            Max = listcount - 1 - lines
            End If
            If Err.Number > 0 Then
                Value = listcount - 1
                Max = listcount - 1
            End If
                      Value = j

        End If
        state = False
    
    End If
   
Else
    state = True
        On Error Resume Next
        Err.clear
        If Not Spinner Then
        Max = listcount - 1
        If Err.Number > 0 Then
            Value = listcount - 1
            Max = listcount - 1
        End If
        End If
    state = False
    
End If
  
    '
j = topitem + lines
If j >= listcount Then j = listcount - 1
'Text1 = ""


    If listcount = 0 Then
        ' DO NOTHING
    Else

       CurrentX = scrollme

  DrawStyle = vbSolid
  fg = Me.ForeColor
  
  If havefocus Then
  caretCreated = False
  DestroyCaret
  End If
        For i = topitem To j
        
        RaiseEvent ExposeRect(i, VarPtr(nr), UserControl.hDC, skipme)
        If Not skipme Then
             If i = SELECTEDITEM - 1 And Not NoCaretShow And Not ListSep(i) Then

nr.Left = scrollme / scrTwips + LeftMarginPixels
              nfg = fg
  RaiseEvent SpecialColor(nfg)
  If nfg <> fg Then Me.ForeColor = nfg
             If mEditFlag Then
                If nfg = fg Then Me.ForeColor = fg
             ElseIf nfg = fg Then
                    If Me.BackColor = 0 Then
                    Me.ForeColor = &HFFFFFF
                    Else
                    Me.ForeColor = 0
                    End If
            End If
            
                    If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                                   myMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i, True
                        End If

                 PrintLineControlSingle UserControl.hDC, List(i), nr
                 Me.ForeColor = fg
             Else
                 If ListSep(i) And List(i) = "" Then
                   hnr.Left = 0
                   hnr.Right = nr.Right
                   hnr.top = nr.top + mytPixels \ 2
                   hnr.Bottom = hnr.top + 1
                   FillBack UserControl.hDC, hnr, ForeColor
                Else
                   If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                                 myMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i
                    End If
                    If ListSep(i) Then
                        ForeColor = dcolor
                    Else
         
               '    ForeColor = fg
                    End If
                    PrintLineControlSingle UserControl.hDC, List(i), nr
                    
                End If
      
             End If
                     If SingleLineSlide Then
nr.Left = LeftMarginPixels
Else
nr.Left = scrollme / scrTwips + LeftMarginPixels
End If
        End If
     nr.top = nr.top + mytPixels
nr.Bottom = nr.top + mytPixels + 1
 ForeColor = fg
    Next i
  
 ''''''''' PrintLineControl UserControl.HDC, g$, nr
    'Print g$
'#  DrawStyle = vbInvisible
 DrawMode = vbInvert

 myt1 = myt - scrTwips
    If SELECTEDITEM > 0 Then
        If SELECTEDITEM - topitem - 1 <= lines And Not ListSep(SELECTEDITEM - 1) Then
                If Not NoCaretShow Then
                                If EditFlag And Not BlockItemcount Then
                                If SelStart = 0 Then SelStart = 1
                                        DrawStyle = vbSolid
                                 If CenterText Then
                                        ' (UserControl.ScaleWidth- LeftMarginPixels * scrTwips-UserControlTextWidth(list$(selecteitem-1)))/2
                                          REALX = UserControlTextWidth(Mid$(List(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips + (UserControl.ScaleWidth - LeftMarginPixels * scrTwips - UserControlTextWidth(List$(SELECTEDITEM - 1))) / 2
                                          REALX2 = scrollme / 2 + REALX
                                            Else
                                   REALX = UserControlTextWidth(Mid$(List(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
                                    REALX2 = scrollme + REALX
                                   End If
                                  If Noflashingcaret Or Not havefocus Then
    
                          Line (scrollme + REALX, (SELECTEDITEM - topitem - 1) * myt + myt1 + mHeadlineHeightTwips)-(scrollme + REALX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips), ForeColor
    
                           Else
                                   ShowMyCaretInTwips REALX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips
                                   End If
                                   If Not NoScroll Then If REALX2 > Width * 0.8 Then scrollme = scrollme - Width * 0.2: PrepareToShow 10
                                   If REALX2 < 0 Then
                                   If Not NoScroll Then
                                     scrollme = scrollme + Width * 0.2
                                   
:
                                   PrepareToShow 10
                                   End If
                                   End If
                                    Else
                                         DrawStyle = vbInvisible
                                
                                        If BackStyle = 1 Then
                            
                                            Line (scrTwips, (SELECTEDITEM - topitem) * myt + mHeadlineHeightTwips)-(scrollme + UserControl.Width - 2 * scrTwips, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips), 0, B
                                        Else
                                            Line (0, (SELECTEDITEM - topitem) * myt + mHeadlineHeightTwips)-(scrollme + UserControl.Width, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips), 0, B
                                     
                                        End If
                                End If
                End If
        Else
        HideCaret (hWnd)
        End If
    End If
    
    CurrentY = 0
    CurrentX = 0
    
    DrawMode = vbCopyPen
    


End If
 DrawStyle = vbSolid
    LastVScroll = Value
RepaintScrollBar
End Sub
Public Sub ShowMe2()
Dim YYT As Long, nr As RECT, j As Long, i As Long, skipme As Boolean, fg As Long, hnr As RECT, nfg As Long
 Dim REALX As Long, REALX2 As Long, myt1
If listcount = 0 And HeadLine = "" Then
Repaint
HideCaret (hWnd)
Exit Sub
End If
If MultiSelect And LeftMarginPixels < mytPixels Then LeftMarginPixels = mytPixels
Repaint

YYT = myt
nr.top = 0
nr.Left = 0 '
hnr.Left = 0  ' no scrolling
nr.Bottom = mytPixels + 1
hnr.Bottom = mytPixels + 1
nr.Right = Width / scrTwips
hnr.Right = Width / scrTwips

If mHeadline <> "" Then
nr.Bottom = HeadlineHeight
RaiseEvent ExposeRect(-1, VarPtr(nr), UserControl.hDC, skipme)
nr.Bottom = HeadlineHeight
CalcRectHeader UserControl.hDC, mHeadline, hnr, DT_CENTER
If Not skipme Then

If mHeadlineHeight <> hnr.Bottom Then
HeadlineHeight = hnr.Bottom
nr.Bottom = mHeadlineHeight
End If
FillBack UserControl.hDC, nr, CapColor
End If
hnr.top = (nr.Bottom - hnr.Bottom) \ 2
hnr.Bottom = nr.Bottom - hnr.top
hnr.Left = 0
hnr.Right = nr.Right
PrintLineControlHeader UserControl.hDC, mHeadline, hnr, DT_CENTER

nr.top = nr.Bottom
nr.Bottom = nr.top + mytPixels + 1
End If
If AutoPanPos Then

If SelStart = 0 Then SelStart = 1
scrollme = 0
again123:

REALX = UserControlTextWidth(Mid$(List(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
REALX2 = scrollme + REALX
If Not NoScroll Then If REALX2 > Width * 0.8 Then scrollme = scrollme - Width * 0.2:  GoTo again123
If REALX2 < 0 Then
If Not NoScroll Then scrollme = scrollme + Width * 0.2: GoTo again123

End If
End If
          
            


If SingleLineSlide Then
nr.Left = LeftMarginPixels
Else
nr.Left = scrollme / scrTwips + LeftMarginPixels
End If
j = topitem + lines
If j >= listcount Then j = listcount - 1

If listcount = 0 Then
BarVisible = False

Exit Sub
Else

 DrawStyle = vbSolid

  If havefocus Then
  caretCreated = False
  DestroyCaret
  End If
fg = Me.ForeColor
For i = topitem To j
CurrentX = scrollme
CurrentY = 0
  RaiseEvent ExposeRect(i, VarPtr(nr), UserControl.hDC, skipme)
  If Not skipme Then
  If i = SELECTEDITEM - 1 And Not NoCaretShow And Not ListSep(i) Then
    nfg = fg
  RaiseEvent SpecialColor(nfg)
  If nfg <> fg Then Me.ForeColor = nfg
  nr.Left = scrollme / scrTwips + LeftMarginPixels
  If mEditFlag Then
   If nfg = fg Then Me.ForeColor = fg
  ElseIf nfg = fg Then
  If Me.BackColor = 0 Then
  Me.ForeColor = &HFFFFFF
  Else
  Me.ForeColor = 0
  End If
  End If

   If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
 myMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i, True
 End If

   PrintLineControlSingle UserControl.hDC, List(i), nr
 If nfg = fg Then Me.ForeColor = fg
 Else
    nfg = fg
  RaiseEvent SpecialColor(nfg)
 If ListSep(i) And List(i) = "" Then
 hnr.Left = 0
 hnr.Right = nr.Right
 hnr.top = nr.top + mytPixels \ 2
 hnr.Bottom = hnr.top + 1
 FillBack UserControl.hDC, hnr, ForeColor
 Else

 If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
 myMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i
 End If
  If ListSep(i) Then
 ForeColor = dcolor
 Else
 If nfg = fg Then ForeColor = fg
   If SELECTEDITEM - 1 = i And nfg <> fg Then
   Me.ForeColor = nfg  'uintnew(&HFFFFFF) - uintnew(nfg)
   End If
 End If

 PrintLineControlSingle UserControl.hDC, List(i), nr
 End If

   End If
 If SingleLineSlide Then
nr.Left = LeftMarginPixels
Else
nr.Left = scrollme / scrTwips + LeftMarginPixels
End If
 
  End If
 
nr.top = nr.top + mytPixels
nr.Bottom = nr.top + mytPixels + 1
ForeColor = fg
Next i

 myt1 = myt - scrTwips
' DrawStyle = vbInvisible
DrawMode = vbInvert
If SELECTEDITEM > 0 Then

    If SELECTEDITEM - topitem - 1 <= lines And SELECTEDITEM > topitem And Not ListSep(SELECTEDITEM - 1) Then
       '' cY = yyt * (i - topitem + 1) 'CurrentY
        
        If Not NoCaretShow Then
                 If EditFlag And Not BlockItemcount Then
                    If SelStart = 0 Then SelStart = 1
                                             DrawStyle = vbSolid
                                          If CenterText Then
                                        ' (UserControl.ScaleWidth- LeftMarginPixels * scrTwips-UserControlTextWidth(list$(selecteitem-1)))/2
                                          REALX = UserControlTextWidth(Mid$(List(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips + (UserControl.ScaleWidth - LeftMarginPixels * scrTwips - UserControlTextWidth(List$(SELECTEDITEM - 1))) / 2
                                          REALX2 = scrollme / 2 + REALX
                                            Else
                                   REALX = UserControlTextWidth(Mid$(List(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
                                    REALX2 = scrollme + REALX
                                   End If
                                   
                                  If Noflashingcaret Or Not havefocus Then
                                 
                                  Line (scrollme + REALX, (SELECTEDITEM - topitem - 1) * myt + myt1 + mHeadlineHeightTwips)-(scrollme + REALX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips), ForeColor
                      
                                Else
                                   ShowMyCaretInTwips REALX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips
                                   End If
                                   If Not NoScroll Then If REALX2 > Width * 0.8 Then scrollme = scrollme - Width * 0.2: PrepareToShow 10
                                   If REALX2 < 0 Then
                                    If Not NoScroll Then
                                     scrollme = scrollme + Width * 0.2
                    
:
                                   PrepareToShow 10
                                   End If
                                   End If
                           Else
                                   DrawStyle = vbInvisible
                                   
                                   If BackStyle = 1 Then
                       
                                           Line (scrTwips, (SELECTEDITEM - topitem) * YYT + mHeadlineHeightTwips)-(0 + UserControl.Width, (SELECTEDITEM - topitem - 1) * YYT + mHeadlineHeightTwips - scrTwips / 2), 0, B
                         
                                   Else
                       
                                         Line (0, (SELECTEDITEM - topitem) * YYT + mHeadlineHeightTwips)-(0 + UserControl.Width, (SELECTEDITEM - topitem - 1) * YYT + mHeadlineHeightTwips), 0, B
                           
                                   End If
                End If

        
        End If
        Else

        HideCaret (hWnd)
    End If
Else


End If

 DrawStyle = vbSolid
DrawMode = vbCopyPen
CurrentY = 0
CurrentX = 0
End If
RepaintScrollBar
End Sub

Property Get lines() As Long
Dim l As Long
On Error GoTo ex1
 myt = UserControlTextHeight() + addpixels * scrTwips
If restrictLines > 0 Then
l = restrictLines - 1
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips - 1) / restrictLines

Else
l = Int((UserControl.ScaleHeight - mHeadlineHeightTwips) / myt) - 1
End If
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
ex1:
If l <= 0 Then
l = 0
End If

lines = l
End Property


Private Sub LargeBar1_Change()

If Not state Then



    topitem = Value
  
RaiseEvent ScrollMove(topitem)
Timer1.Enabled = True

LastVScroll = Value

End If
End Sub
Public Sub RepaintOld7_18()
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) \ restrictLines
Else
myt = UserControlTextHeight() + addpixels * scrTwips
End If
'HeadlineHeight = UserControlTextHeight() / SCRTWIPS
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
On Error GoTo th1
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName
If Err.Number > 0 Then
'DestroyCaret
Exit Sub
End If
nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Sub
If Err.Number > 0 Then Exit Sub
If cnt$ <> "" Then
Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
If UserControl.Parent.Picture.handle <> 0 And BackStyle = 1 Then

If Me.BorderStyle = 1 Then
CurrentY = 0
    CurrentX = 0
Line (0, 0)-(ScaleWidth - scrTwips, ScaleHeight - scrTwips), Me.BackColor, B
UserControl.PaintPicture UserControl.Parent.Picture, scrTwips, scrTwips, Width - 2 * scrTwips, Height - 2 * scrTwips, mo.Left, mo.top, Width - 2 * scrTwips, Height - 2 * scrTwips
    CurrentY = 0
    CurrentX = 0
Else
UserControl.PaintPicture UserControl.Parent.Picture, 0, 0, , , mo.Left, mo.top

End If

ElseIf BackStyle = 1 Then
Dim mmo As PictureBox
RaiseEvent GetBackPicture(mmo)
If Not mmo Is Nothing Then
If mmo.Picture.handle <> 0 Then
    UserControl.PaintPicture mmo.Picture, 0, 0, , , mo.Left, mo.top
    If Me.BorderStyle = 1 Then
    CurrentY = 0
        CurrentX = 0
    Line (0, 0)-(ScaleWidth - scrTwips, ScaleHeight - scrTwips), Me.BackColor, B
        CurrentY = 0
        CurrentX = 0
    End If
End If
End If
Else
th1:
UserControl.Cls
End If
End Sub
Public Sub Repaint()
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) \ restrictLines
Else
myt = UserControlTextHeight() + addpixels * scrTwips
End If
'HeadlineHeight = UserControlTextHeight() / SCRTWIPS
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
On Error GoTo th1
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName
If Err.Number > 0 Then
'DestroyCaret
Exit Sub
End If
nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If Not waitforparent Then Exit Sub
If UserControl.Parent Is Nothing Then Exit Sub
If Err.Number > 0 Then Exit Sub
If cnt$ <> "" Then
Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
If BackStyle = 1 Then
    If Not SkipForm Then
        If UserControl.Parent.Picture.handle <> 0 Then
            If Me.BorderStyle = 1 Then
                    CurrentY = 0
                    CurrentX = 0
                    Line (0, 0)-(ScaleWidth - scrTwips, ScaleHeight - scrTwips), Me.BackColor, B
                    UserControl.PaintPicture UserControl.Parent.Picture, scrTwips, scrTwips, Width - 2 * scrTwips, Height - 2 * scrTwips, mo.Left, mo.top, Width - 2 * scrTwips, Height - 2 * scrTwips
                    CurrentY = 0
                    CurrentX = 0
            Else
                    UserControl.PaintPicture UserControl.Parent.Picture, 0, 0, , , mo.Left, mo.top
            End If
            Else
            If Me.BorderStyle = 1 Then
                CurrentY = 0
                CurrentX = 0
                Line (0, 0)-(ScaleWidth - scrTwips, ScaleHeight - scrTwips), Me.BackColor, B
                UserControl.PaintPicture UserControl.Parent.image, scrTwips, scrTwips, Width - 2 * scrTwips, Height - 2 * scrTwips, mo.Left, mo.top, Width - 2 * scrTwips, Height - 2 * scrTwips
                CurrentY = 0
                CurrentX = 0
            Else
                UserControl.PaintPicture UserControl.Parent.image, 0, 0, , , mo.Left, mo.top
            End If
        End If
    Else
        Dim mmo As Object, isfrm As Boolean
        RaiseEvent GetBackPicture(mmo)
        If Not mmo Is Nothing Then
            If mmo.image.handle <> 0 Then
                UserControl.PaintPicture mmo.image, 0, 0, , , mo.Left - mmo.Left, mo.top - mmo.top
                If Me.BorderStyle = 1 Then
                CurrentY = 0
                    CurrentX = 0
                Line (0, 0)-(ScaleWidth - scrTwips, ScaleHeight - scrTwips), Me.BackColor, B
                    CurrentY = 0
                    CurrentX = 0
                End If
            End If
        End If
    End If
Else
th1:
UserControl.Cls
End If
End Sub
Private Function GetStrUntilB(pos As Long, ByVal sStr As String, fromStr As String, Optional RemoveSstr As Boolean = True) As String
Dim i As Long
If fromStr = "" Then GetStrUntilB = "": Exit Function
If pos <= 0 Then pos = 1
If pos > Len(fromStr) Then
    GetStrUntilB = ""
Exit Function
End If
i = InStr(pos, fromStr, sStr)
If (i < 1 + pos) And Not ((i > 0) And RemoveSstr) Then
    GetStrUntilB = ""
    pos = Len(fromStr) + 1
Else
    GetStrUntilB = Mid$(fromStr, pos, i - pos)
    If RemoveSstr Then
        pos = i + Len(sStr)
    Else
        pos = i
    End If
End If
End Function
Function design() As Boolean
If listcount = 0 Then
   '      barvisible = False

Cls
If UserControl.Ambient.UserMode = False Then
CurrentX = scrTwips
CurrentY = scrTwips
Print UserControl.Ambient.DisplayName

CurrentX = 0
CurrentY = 0
End If
design = True
End If
End Function
Private Sub LargeBar1_Scroll()
If Not state Then
 topitem = Value
RaiseEvent ScrollMove(topitem)
Timer1.Enabled = True
LastVScroll = Value
End If
End Sub
Public Function UserControlTextWidthPixels(A$) As Long
Dim nr As RECT
If Len(A$) > 0 Then
CalcRect UserControl.hDC, A$, nr
UserControlTextWidthPixels = nr.Right
End If
End Function
Public Function UserControlTextWidth(A$) As Long
Dim nr As RECT
CalcRect UserControl.hDC, A$, nr
UserControlTextWidth = nr.Right * scrTwips
End Function
Private Function UserControlTextHeight() As Long
Dim nr As RECT
If overrideTextHeight = 0 Then
CalcRect1 UserControl.hDC, "fj", nr
UserControlTextHeight = nr.Bottom * scrTwips
Exit Function
End If
UserControlTextHeight = overrideTextHeight

End Function

Private Sub PrintLineControlSingle(mHdc As Long, c As String, r As RECT)
' this is our basic print routine
Dim that As Long, cc As String
If CenterText Then that = DT_CENTER
If VerticalCenterText Then that = that Or DT_VCENTER
If WrapText Then
DrawText mHdc, StrPtr(c), -1, r, DT_WORDBREAK Or DT_NOPREFIX Or DT_MODIFYSTRING Or that
Else
If LastLinePart <> "" Then
cc = c + LastLinePart
   DrawText mHdc, StrPtr(cc), -1, r, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that
Else

    DrawText mHdc, StrPtr(c), -1, r, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that
    End If
    End If
    
    End Sub
Private Sub PrintLineControlHeader(mHdc As Long, c As String, r As RECT, Optional that As Long = 0)
' this is our basic print routine

DrawText mHdc, StrPtr(c), -1, r, DT_WORDBREAK Or DT_NOPREFIX Or DT_MODIFYSTRING Or that

    
    End Sub
  Private Sub CalcRectHeader(mHdc As Long, c As String, r As RECT, Optional that As Long = 0)
r.top = 0
r.Left = 0
If r.Right = 0 Then r.Right = UserControl.Width / scrTwips
DrawText mHdc, StrPtr(c), -1, r, DT_CALCRECT Or DT_WORDBREAK Or DT_NOPREFIX Or DT_MODIFYSTRING Or that
End Sub
Private Sub PrintLineControl(mHdc As Long, c As String, r As RECT)

    DrawText mHdc, StrPtr(c), -1, r, DT_NOPREFIX Or DT_NOCLIP

End Sub
Private Sub PrintLinePixels(dd As Object, c As String)
Dim r As RECT    ' print to a picturebox as label
r.Right = dd.ScaleWidth
r.Bottom = dd.ScaleHeight
DrawText dd.hDC, StrPtr(c), -1, r, DT_NOPREFIX Or DT_WORDBREAK
End Sub
Private Sub CalcRect(mHdc As Long, c As String, r As RECT)
r.top = 0
r.Left = 0
Dim that As Long
If CenterText Then that = DT_CENTER
If VerticalCenterText Then that = that Or DT_VCENTER
If WrapText Then
If r.Right = 0 Then r.Right = UserControl.Width / scrTwips
DrawText mHdc, StrPtr(c), -1, r, DT_CALCRECT Or DT_WORDBREAK Or DT_NOPREFIX Or DT_MODIFYSTRING Or that
Else
    DrawText mHdc, StrPtr(c), -1, r, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that
    End If

End Sub
Private Sub CalcRect1(mHdc As Long, c As String, r As RECT)
r.top = 0
r.Left = 0

If WrapText Then
If r.Right = 0 Then r.Right = UserControl.Width / scrTwips - LeftMarginPixels

DrawText mHdc, StrPtr(c), -1, r, DT_CALCRECT Or DT_WORDBREAK Or DT_NOPREFIX
Else
    DrawText mHdc, StrPtr(c), -1, r, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP
    End If

End Sub

Public Function SpellUnicode(A$)
' use spellunicode to get numbers
' and make a ListenUnicode...with numbers for input text
Dim b$, i As Long
For i = 1 To Len(A$) - 1
b$ = b$ & CStr(AscW(Mid$(A$, i, 1))) & ","
Next i
SpellUnicode = b$ & CStr(AscW(Right$(A$, 1)))
End Function
Public Function ListenUnicode(ParamArray aa() As Variant) As String
Dim all$, i As Long
For i = 0 To UBound(aa)
    all$ = all$ & ChrW(aa(i))
Next i
ListenUnicode = all$
End Function

Public Sub RepaintFromOut(parentpic As StdPicture, myleft As Long, mytop As Long)
On Error GoTo th1

If parentpic.handle <> 0 Then
UserControl.PaintPicture parentpic, 0, 0, , , myleft, mytop
Else
th1:
'UserControl.Cls
End If
End Sub
Private Sub Redraw(ParamArray status())

If EnabledBar Then
Dim fakeLargeChange As Long, newheight As Long, newtop As Long
Dim b As Boolean, nstatus As Boolean
Timer2bar.Enabled = False
If UBound(status) >= 0 Then
nstatus = CBool(status(0))
Else
nstatus = Shape1.Visible
End If
With UserControl
If mHeadline <> "" Then
newheight = .Height - mHeadlineHeightTwips
newtop = mHeadlineHeightTwips
Else
newheight = .Height
End If

If newheight <= 0 Then

Else
        minimumWidth = (1 - (Max - Min) / (largechange + Max - Min)) * newheight * (1 - Percent * 2) + 1
        If minimumWidth < 60 Then
        
        mLargeChange = Round(-(Max - Min) / ((60 - 1) / newheight / (1 - Percent * 2) - 1) - Max + Min) + 1
        
        minimumWidth = (1 - (Max - Min) / (largechange + Max - Min)) * newheight * (1 - Percent * 2) + 1
        End If
        valuepoint = (Value - Min) / (largechange + Max - Min) * (newheight * (1 - 2 * Percent)) + newheight * Percent

       Shape Shape1, Width - barwidth, newtop + valuepoint, barwidth, minimumWidth
       Shape Shape2, Width - barwidth, newtop + newheight * (1 - Percent), barwidth, newheight * Percent ' newtop + newheight * Percent - scrTwips
        Shape Shape3, Width - barwidth, newtop, barwidth, newheight * Percent   ' left or top
End If
End With
If UBound(status) >= 0 Then
b = (CBool(status(0)) Or Spinner) And listcount > lines
If Not Shape1.Visible = b Then
Shape1.Visible = b
Shape2.Visible = b
Shape3.Visible = b

End If
End If

End If
End Sub
Private Property Get largechange() As Long
If mLargeChange < 1 Then mLargeChange = 1
largechange = mLargeChange
End Property

Private Property Let largechange(ByVal rhs As Long)
If rhs < 1 Then rhs = 1
mLargeChange = rhs
showshapes
PropertyChanged "LargeChange"
End Property
Public Property Get smallchange() As Long
smallchange = mSmallChange
End Property

Private Property Let smallchange(ByVal rhs As Long)
If rhs < 1 Then rhs = 1
mSmallChange = rhs
showshapes
PropertyChanged "SmallChange"
End Property
Private Property Get Max() As Long
Max = mmax
End Property

Private Property Let Max(ByVal rhs As Long)
If Min > rhs Then rhs = Min
If mValue > rhs Then mValue = rhs  ' change but not send event
If rhs = 0 Then rhs = 1
mmax = rhs
showshapes
PropertyChanged "Max"
End Property

Private Property Get Min() As Long
Min = mmin
End Property
Public Sub SetSpin(low As Long, high As Long, stepbig As Long)
If Spinner Then
mpercent = 0.33
mmax = high
mmin = low
mLargeChange = (Max - Min) * 0.2
mSmallChange = stepbig
mjumptothemousemode = True
End If
End Sub

Private Property Let Min(ByVal rhs As Long)
If Max <= rhs Then rhs = Max
If mValue < rhs Then mValue = rhs  ' change but not send event

mmin = rhs
showshapes
PropertyChanged "LargeChange"
PropertyChanged "Min"
End Property
Public Property Get EnabledBar() As Boolean
EnabledBar = Not NoFire
End Property

Public Property Let EnabledBar(ByVal rhs As Boolean)
If Not myEnabled Then Exit Property
NoFire = Not EnabledBar
Shape1.Visible = Not NoFire
Shape2.Visible = Not NoFire
Shape3.Visible = Not NoFire
Shape Shape1
Shape Shape2
Shape Shape3
If Not NoFire = True Then Timer1.Enabled = True
End Property
Public Property Get Value() As Long
Value = mValue
End Property
Public Property Get Visible() As Boolean
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Property
If Err.Number > 0 Then Exit Property
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If

Visible = mo.Visible
End Property

Public Property Get TopTwips() As Long
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Property
If Err.Number > 0 Then Exit Property
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
TopTwips = CLng(mo.top)
End Property
Public Property Let Visible(ByVal rhs As Boolean)
Dim MM$, mo As Control, nm$, cnt$, p As Long
On Error Resume Next
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)

If UserControl.Parent Is Nothing Then Exit Property
If Err.Number > 0 Then Exit Property
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
If mo.Visible = rhs Then Exit Property
mo.Visible = rhs

End Property
Public Property Let TopTwips(ByVal rhs As Long)
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Property
If Err.Number > 0 Then Exit Property
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
mo.Move mo.Left, CSng(rhs)
End Property
Public Property Get HeightTwips() As Long
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Property
If Err.Number > 0 Then Exit Property
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
HeightTwips = CLng(mo.Height)
End Property
Public Property Let HeightTwips(ByVal rhs As Long)
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Property
If Err.Number > 0 Then Exit Property
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
mo.Move mo.Left, mo.top, mo.Width, rhs
End Property
Public Sub MoveTwips(ByVal mleft As Long, ByVal mtop As Long, mWidth As Long, mHeight As Long)
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Sub
If Err.Number > 0 Then Exit Sub
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
mo.Move mleft, mtop, mWidth, mHeight
End Sub
Public Sub ZOrder(Optional ByVal rhs As Long = 0)
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Sub
If Err.Number > 0 Then Exit Sub
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
mo.ZOrder rhs
End Sub

Public Sub SetFocus()
Dim MM$, mo As Control, nm$, cnt$, p As Long
MM$ = UserControl.Ambient.DisplayName

nm$ = GetStrUntilB(p, "(", MM$ & "(", True)
cnt$ = GetStrUntilB(p, ")", MM$, True)
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Sub
If Err.Number > 0 Then Exit Sub
If cnt$ <> "" Then

Set mo = UserControl.Parent.Controls(nm$).item(CInt(cnt$))
Else
Set mo = UserControl.Parent.Controls(nm$)
End If
mo.SetFocus
End Sub
Public Property Let Value(ByVal rhs As Long)
' Dim oldvalue As Long
If rhs < Min Then rhs = Min
If rhs > Max Then rhs = Max
If state And Spinner Then
'don't fix the value
Else
mValue = rhs
End If
showshapes

If Not Spinner Then
If Not NoFire Then
LargeBar1_Change
End If
Else

RaiseEvent SpinnerValue(mmax - mValue + mmin)
Redraw

'UserControl.refresh
End If
PropertyChanged "Value"
End Property
Public Property Let ValueSilent(ByVal rhs As Long)
If Spinner Then
' no events
If rhs < Min Then rhs = Min
If rhs > Max Then rhs = Max
mValue = Max - rhs + Min
showshapes
End If
End Property
Public Property Get ValueSilent() As Long
ValueSilent = Max - mValue + Min
End Property
Private Property Get BarVisible() As Boolean
BarVisible = Shape1.Visible
End Property
Private Property Let BarVisible(ByVal rhs As Boolean)
If Not myEnabled Then
Exit Property
End If
If listcount = 0 Then rhs = False
Shape1.Visible = rhs Or Spinner
Shape2.Visible = rhs Or Spinner
Shape3.Visible = rhs Or Spinner
Shape Shape1
Shape Shape2
Shape Shape3
If Not NoFire = True Then Timer1.Enabled = True
End Property

Private Sub showshapes()
If m_showbar Or StickBar Or Spinner Or AutoHide Then
Timer2bar.Enabled = True
End If
End Sub
Public Property Get Percent() As Single
Percent = mpercent
End Property

Public Property Let Percent(ByVal rhs As Single)
mpercent = rhs
PropertyChanged "Percent"
End Property
Private Sub UserControl_KeyDown(KeyCode As Integer, shift As Integer)
If dropkey Then shift = 0: KeyCode = 0: Exit Sub
Dim i&
If shift = 4 Then
If KeyCode = 18 Then
If mynum$ = "" Then mynum$ = "0"
KeyCode = 0
Exit Sub
End If
Select Case KeyCode
Case vbKeyAdd
mynum$ = "&h"
Case vbKey0 To vbKey9
mynum$ = mynum$ + Chr$(KeyCode - vbKey0 + 48)
Case vbKeyNumpad0 To vbKeyNumpad9
mynum$ = mynum$ + Chr$(KeyCode - vbKeyNumpad0 + 48)
Case vbKeyA To vbKeyF
If Left$(mynum$, 1) = "&" Then
mynum$ = mynum$ + Chr$(KeyCode - vbKeyNumpad0 + 65)
Else
mynum$ = ""
End If
Case Else
mynum$ = ""
End Select

Exit Sub
End If

mynum$ = ""
If shift <> 0 And KeyCode = 0 Then Exit Sub
RaiseEvent KeyDown(KeyCode, shift)
If (KeyCode = 0) Or Not (Enabled Or state) Then Exit Sub
If SELECTEDITEM < 0 Then
SELECTEDITEM = topitem + 1: ShowMe2
If Not EditFlag Then: KeyCode = 0
End If
LargeBar1KeyDown KeyCode, shift
If EnabledBar Then
Select Case KeyCode
Case vbKeyLeft, vbKeyUp
If Spinner Then
If shift Then
Value = Value - 1
Else
Value = Value - mSmallChange
End If
Else
Value = Value - mSmallChange
End If
Case vbKeyPageUp
Value = Value - largechange
Case vbKeyRight, vbKeyDown
If Spinner Then
If shift Then
Value = Value + 1
Else
Value = Value + mSmallChange
End If
Else
If Value + largechange + 1 <= Max Then
Value = Value + mSmallChange
End If
End If
Case vbKeyPageDown
Value = Value + largechange
End Select
End If

i = GetLastKeyPressed
 If i <> -1 And i <> 94 Then
 UKEY$ = ChrW(i)
 Else

 End If
End Sub
Public Property Get Vertical() As Boolean
Vertical = mVertical
End Property

Public Property Let Vertical(ByVal rhs As Boolean)
rhs = True ' intercept
mVertical = rhs
showshapes
PropertyChanged "Vertical"
End Property

Public Property Get jumptothemousemode() As Boolean
jumptothemousemode = mjumptothemousemode
End Property

Public Property Let jumptothemousemode(ByVal rhs As Boolean)
mjumptothemousemode = rhs
End Property
Private Function processXY(ByVal x As Single, ByVal y As Single, Optional rep As Boolean = True) As Boolean
Timer1bar.Enabled = False
Dim checknewvalue As Long, newheight As Long
With UserControl
If mHeadline <> "" Then
newheight = .Height - mHeadlineHeightTwips
y = y - mHeadlineHeightTwips
Else
newheight = .Height
End If

If minimumWidth < 60 Then minimumWidth = 60  ' 4 x scrtwips
' value must have real max ...minimum MAX-60
If Vertical Then
' here minimumwidth is minimumheight
If y >= valuepoint - scrTwips And y <= minimumWidth + valuepoint - scrTwips Then
' is our scroll bar
OurDraw = Not rep

ElseIf y > newheight * Percent And y < newheight * (1 - Percent) Then
'  we are inside so take a largechange
processXY = True

        If y < valuepoint Then
         ' jump to mouse position at page (or fakepage )
                    If mjumptothemousemode Then
                     y = (y \ minimumWidth + 1) * minimumWidth - minimumWidth
                     Else
                    y = valuepoint - minimumWidth
                    End If
        Else
         ' jump to mouse position at page (or fakepage )
                If mjumptothemousemode Then
                y = (y \ minimumWidth - 1) * minimumWidth + minimumWidth
                Else
                y = valuepoint + minimumWidth
                End If
        End If
            If y < newheight * Percent Then y = newheight * Percent
            If y > Round(newheight * (1 - Percent)) - minimumWidth + newheight * Percent Then
            y = newheight * (1 - Percent) - minimumWidth
            End If
            checknewvalue = Round((y - newheight * Percent) * (Max - Min) / ((newheight * (1 - Percent) - minimumWidth) - newheight * Percent)) + Min
            If checknewvalue = Value And mjumptothemousemode Then
                 ' do nothing
                
            Else
    
                Value = checknewvalue
                If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5 ' autorepeat
                Timer1bar.Enabled = True
            End If
ElseIf y >= newheight * (1 - Percent) And y <= newheight Then ' is right button
processXY = True
checknewvalue = Value + mSmallChange
If checknewvalue = Value Then
' do nothing
Else
Value = checknewvalue
If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5  '
Timer1bar.Enabled = True
End If

ElseIf y < newheight * Percent - scrTwips Then
processXY = True
checknewvalue = Value - mSmallChange
' is  left button
If checknewvalue = Value Then
' do nothing
Else
Value = checknewvalue

If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5 ' autorepeat
Timer1bar.Enabled = True
End If
End If

ElseIf Not Timer1bar.Enabled Then
If x >= valuepoint - scrTwips And x <= minimumWidth + valuepoint - scrTwips Then
' is our scroll bar
OurDraw = Not rep

ElseIf x > .Width * Percent And x < .Width * (1 - Percent) Then
processXY = True
'  we are inside so take a largechange
        If x < valuepoint Then
                If mjumptothemousemode Then
                  x = (x \ minimumWidth + 1) * minimumWidth - minimumWidth
                Else
                x = valuepoint - minimumWidth
                End If
        Else
                If mjumptothemousemode Then
                x = (x \ minimumWidth - 1) * minimumWidth + minimumWidth
                Else
                x = valuepoint + minimumWidth
                End If
        End If
            If x < .Width * Percent Then x = .Width * Percent
            If x > Round(.Width * (1 - Percent)) - minimumWidth + .Width * Percent Then
            x = .Width * (1 - Percent) - minimumWidth
            End If
            checknewvalue = Round((x - .Width * Percent) * (Max - Min) / ((.Width * (1 - Percent) - minimumWidth) - .Width * Percent)) + Min
            If checknewvalue = Value And mjumptothemousemode Then
            ' do nothing
            Else
            Value = checknewvalue
            If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5  ' autorepeat
            Timer1bar.Enabled = True
            End If
ElseIf x >= .Width * (1 - Percent) And x <= .Width Then
processXY = True
checknewvalue = Value + mSmallChange
If checknewvalue = Value Then
' do nothing
Else
Value = checknewvalue
If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5 ' autorepeat
Timer1bar.Enabled = True
End If
' is right button
ElseIf x < .Width * Percent - scrTwips Then
processXY = True
checknewvalue = Value - mSmallChange
If checknewvalue = Value Then
' do nothing
Else
Value = checknewvalue
If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5  ' autorepeat
Timer1bar.Enabled = True
' is  left button
End If
End If

End If
End With
End Function
Private Sub barMouseMove(Button As Integer, shift As Integer, x As Single, ByVal y As Single)
If Not EnabledBar Then Exit Sub
Dim ForValidValue As Long, newheight As Long

If OurDraw Then
If Button = 1 Then
Timer1bar.Interval = 5000
'timer2bar.enabled = False
If minimumWidth < 60 Then minimumWidth = 60  ' 4 x scrtwips

With UserControl


If Vertical Then

If mHeadline <> "" Then
y = y - mHeadlineHeightTwips
newheight = .Height - mHeadlineHeightTwips
Else
newheight = .Height
End If
        ForValidValue = y + GetOpenValue 'ForValidValue + valuepoint
        If ForValidValue < newheight * Percent Then
        ForValidValue = newheight * Percent
        Value = Min
        ElseIf ForValidValue > ((newheight * (1 - Percent) - minimumWidth)) Then
        ForValidValue = ((newheight * (1 - Percent) - minimumWidth))
        Value = Max
        Else

         Value = Round((ForValidValue - newheight * Percent) * (Max - Min) / ((newheight * (1 - Percent) - minimumWidth) - newheight * Percent)) + Min
         
        End If
    

Else

         ForValidValue = x + GetOpenValue
        If ForValidValue < .Width * Percent Then
        ForValidValue = .Width * Percent
        Value = Min
        ElseIf ForValidValue > ((.Width * (1 - Percent) - minimumWidth)) Then
        ForValidValue = ((.Width * (1 - Percent) - minimumWidth))
        Value = Max
        Else
        Value = Round((ForValidValue - .Width * Percent) * (Max - Min) / ((.Width * (1 - Percent) - minimumWidth) - .Width * Percent)) + Min
        
        End If
      
End If
showshapes
'Redraw


End With
If Not Spinner Then
If Not NoFire Then LargeBar1_Scroll
Else

RaiseEvent SpinnerValue(mmax - mValue + mmin)
End If
End If
End If
End Sub
Public Sub MenuItem(ByVal item As Long, Checked As Boolean, radiobutton As Boolean, firstState As Boolean, Optional Id$)
' Using MenuItem we want glist to act as a menu with checked and radio buttons
item = item - 1  ' from 1...to listcount as input
' now from 0 to listcount-1
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 And item < listcount Then
If LeftMarginPixels < mytPixels Then LeftMarginPixels = mytPixels
mlist(item).Checked = Checked ' means that can be checked
mlist(item).contentID = Id$
ListSelected(item) = firstState
mlist(item).radiobutton = radiobutton ' one of the group can be checked
End If
End If
End Sub
Public Function GetMenuId(Id$, pos As Long) As Boolean
' return item number with that id$
' work only in the internal list
Dim i As Long
If itemcount > 0 And Not BlockItemcount Then
For i = 0 To itemcount - 1
If mlist(i).contentID = Id$ Then pos = i: Exit For
Next i
End If
GetMenuId = Not (i = itemcount)
End Function
Property Get Id(item As Long) As String
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 And item < listcount Then
Id = mlist(item).contentID
End If
End If
End Property
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
Private Sub myMark(thathDC As Long, radius As Long, x As Long, y As Long, item As Long, Optional reverse As Boolean = False) ' circle
'
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
Dim th As RECT
th.Left = x - radius
th.top = y - radius
th.Right = x + radius
th.Bottom = y + radius
Dim old_brush As Long, old_pen As Long, my_brush As Long

    If Not ListChecked(item) Then
                If reverse Then
                  my_brush = CreateSolidBrush(0)
                Else
                   my_brush = CreateSolidBrush(m_ForeColor)
                End If
            FillRect thathDC, th, my_brush
            DeleteObject my_brush
             radius = radius - 2
             If radius = 0 Then radius = 1
        Else
        radius = 4
        End If
             
        th.Left = x - radius
        th.top = y - radius
        th.Right = x + radius
        th.Bottom = y + radius

        

        If ListSelected(item) Then
            If reverse Then
                my_brush = CreateSolidBrush(0)  '
            Else
                my_brush = CreateSolidBrush(m_dcolor)  'm_CapColor
            End If
        Else
        If reverse Then
            my_brush = CreateSolidBrush(&HFFFFFF)
        Else
            my_brush = CreateSolidBrush(m_backcolor)
        End If
             End If
     FillRect thathDC, th, my_brush
DeleteObject my_brush
 



End Sub


Public Property Get widthtwips() As Long

widthtwips = UserControl.Width
End Property
Public Property Get WidthPixels() As Long
WidthPixels = UserControl.Width / scrTwips
End Property
Public Property Get HeightPixels() As Long
HeightPixels = UserControl.Height / scrTwips
End Property
Public Sub REALCUR(ByVal s$, ByVal probeX As Single, realpos As Long, usedCharLength As Long, Optional notextonly As Boolean = False)
' for a probeX (maybe a cursor position or a wrapping point) we want to know for a S$, what is the real posistion in realpos,
' and how match is the length of S$ in the left side of that position

Dim n As Long, st As Long, st1 As Long, st0 As Long
'probeX = probeX - scrollme
'If Not notextonly Then probeX = probeX - UserControlTextWidth("W") ' Else' probeX = probeX + 2 * scrTwips

n = UserControlTextWidth(s$)

If CenterText Then
probeX = scrollme / 2 + probeX - LeftMarginPixels * scrTwips - (UserControl.ScaleWidth - LeftMarginPixels * scrTwips - n) / 2 + 2 * scrTwips
Else
probeX = probeX - LeftMarginPixels * scrTwips + 2 * scrTwips
End If

If probeX > n Then
If s$ = "" Then
realpos = 0
usedCharLength = 1
Else
realpos = n
usedCharLength = Len(s$)
End If
Else
If probeX <= 30 Then
realpos = 0
usedCharLength = 0
Exit Sub
End If
st = Len(s$)
st1 = st + 1
st0 = 1
While st > st0 + 1
st1 = (st + st0) \ 2
If probeX >= UserControlTextWidth(Mid$(s$, 1, st1)) Then
st0 = st1
Else
st = st1
End If
Wend
If probeX > UserControlTextWidth(Mid$(s$, 1, st1)) Then
st1 = st1 + 1
Else
If st1 = 2 Then
If probeX < UserControlTextWidth(Mid$(s$, 1, 1)) Then st1 = 1
End If
End If
Do
st1 = st1 - 1
s$ = Mid$(s$, 1, st1)  '
realpos = UserControlTextWidth(s$)
Loop While realpos > probeX And Len(s$) > 1
usedCharLength = Len(s$)
End If
End Sub
Public Function Pixels2Twips(pixels As Long) As Long
Pixels2Twips = pixels * scrTwips
End Function
Public Function BreakLine(data As String, datanext As String, Optional thatTwipsPreserveRight As Long = -1) As Boolean
Dim i As Long, k As Long, m As Long
If thatTwipsPreserveRight = -1 Then
m = widthtwips
Else
m = widthtwips - thatTwipsPreserveRight
End If
REALCURb data, m, k, i, True
datanext = Mid$(data, 1, i)
data = Mid$(data, i + 1)

' lets see if we have space in data
If Len(data) > 0 Then
    If Right$(datanext, 1) <> " " And Left$(data, 1) <> " " Then
    ' we have a broken word
    m = InStrRev(datanext, " ")
    If m > 0 Then
    ' we have a space inside datanext
    If m > 1 Then
    data = Mid$(datanext, m + 1) + data
    datanext = Left$(datanext, m)
    Else
    ' do nothing, we will have nothing for this line if we take the word
    End If
    Else
    ' do nothing it is a big word...
    m = InStrRev(datanext, "\")
    If m > 1 Then
    data = Mid$(datanext, m + 1) + data
    datanext = Left$(datanext, m)
    Else
    ' do nothing, we will have nothing for this line if we take the word
    End If
    End If
    End If
    
    i = 1
    If data <> " " Or data$ = "" Then
    While Left$(data, i) = " "
    i = i + 1
    Wend
    End If
    datanext = datanext + Mid$(data, 1, i - 1)
    data = Mid$(data, i)
    
End If
BreakLine = data <> ""
End Function
Public Sub REALCURb(ByVal s$, ByVal probeX As Single, realpos As Long, usedCharLength As Long, Optional notextonly As Boolean = False)
' this is for breakline only
Dim n As Long, st As Long, st1 As Long, st0 As Long

If Not notextonly Then probeX = probeX - UserControlTextWidth("W") ' Else' probeX = probeX + 2 * scrTwips
n = UserControlTextWidth(s$)


probeX = probeX - 2 * LeftMarginPixels * scrTwips - 2 * scrTwips

If probeX > n Then
If s$ = "" Then
realpos = 0
usedCharLength = 1
Else
realpos = n
usedCharLength = Len(s$) + 1
End If
Else
If probeX <= 30 Then
realpos = 0
usedCharLength = 1
Exit Sub
End If
st = Len(s$)
st1 = st + 1
st0 = 1
While st > st0 + 1
st1 = (st + st0) \ 2
If probeX >= UserControlTextWidth(Mid$(s$, 1, st1)) Then
st0 = st1
Else
st = st1
End If
Wend

If probeX > UserControlTextWidth(Mid$(s$, 1, st1 + 1)) Then
st1 = st1 + 1
Else
If probeX < UserControlTextWidth(Mid$(s$, 1, st1)) Then st1 = st0  ' new in m2000
If st1 = 2 Then

If probeX < UserControlTextWidth(Mid$(s$, 1, 1)) Then st1 = 1
End If
End If
s$ = Mid$(s$, 1, st1)  '
realpos = UserControlTextWidth(s$)
usedCharLength = Len(s$)
End If
End Sub


Property Let ListindexPrivateUse(item As Long)
If item < listcount Then
SELECTEDITEM = item + 1
Else
SELECTEDITEM = 0
End If
End Property


Private Property Get SELECTEDITEM() As Long
SELECTEDITEM = Mselecteditem
End Property

Private Property Let SELECTEDITEM(ByVal rhs As Long)
If rhs > listcount And rhs > 0 Then
rhs = 0

If rhs > listcount Then Exit Property
End If
Mselecteditem = rhs
End Property

Public Property Get PanPos() As Long
PanPos = scrollme

End Property
Public Property Get PanPosPixels() As Long
If scrollme <> 0 Then PanPosPixels = scrollme / scrTwips
End Property
Public Property Let PanPos(ByVal rhs As Long)
scrollme = rhs
End Property

Public Sub refresh()
Dim A As Long
Shape Shape1
Shape Shape2
Shape Shape3
A = GdiFlush()
UserControl.refresh
End Sub
Public Property Get PreserveNpixelsHeaderRightTwips() As Long
PreserveNpixelsHeaderRightTwips = mPreserveNpixelsHeaderRight
End Property

Public Property Let PreserveNpixelsHeaderRightTwips(ByVal rhs As Long)
mPreserveNpixelsHeaderRight = rhs
End Property


Public Property Get SelStart() As Long
SelStart = mSelstart
End Property
Public Property Let SelStartEventAlways(ByVal rhs As Long)
Dim checkline As Long
RaiseEvent PromptLine(checkline)
If PromptLineIdent > 0 And (listindex = checkline) And PromptLineIdent >= rhs Then rhs = PromptLineIdent + 1

mSelstart = rhs

RaiseEvent ChangeSelStart(rhs)
mSelstart = rhs
End Property
Public Property Let SelStart(ByVal rhs As Long)
Dim checkline As Long
RaiseEvent PromptLine(checkline)
If PromptLineIdent > 0 And (listindex = checkline) And PromptLineIdent >= rhs Then rhs = PromptLineIdent + 1
If Not (mSelstart = rhs) Then
mSelstart = rhs
RaiseEvent ChangeSelStart(rhs)
mSelstart = rhs

Else
mSelstart = rhs
End If
End Property
Private Sub ShowMyCaretInTwips(x1 As Long, y1 As Long)

If hWnd <> 0 Then
 With UserControl
 If Not caretCreated Then

 CreateCaret hWnd, 0, .ScaleX(1, 1, 3) + 2, .ScaleY(myt, 1, 3) - 2: caretCreated = True
 End If
' we can set caret pos if we don't have the focus

SetCaretPos .ScaleX(x1, 1, 3), .ScaleY(y1, 1, 3)
ShowCaret (hWnd)


End With
End If
End Sub




Public Property Get EditFlag() As Boolean
EditFlag = mEditFlag

End Property

Public Property Let EditFlag(ByVal rhs As Boolean)
mEditFlag = rhs
If Not rhs Then If hWnd <> 0 Then DestroyCaret: caretCreated = False
End Property
Public Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long, Optional ByVal offsetx As Long = 0)
Dim A As RECT
CopyFromLParamToRect A, thatRect
A.Bottom = A.Bottom - 1
A.Left = A.Left + offsetx
FillBack thathDC, A, thatbgcolor
End Sub
Public Sub WriteThere(thatRect As Long, aa$, ByVal offsetx As Long, ByVal offsety As Long, thiscolor As Long)
Dim A As RECT, fg As Long
CopyFromLParamToRect A, thatRect
If A.Left > Width Then Exit Sub
A.Right = WidthPixels
A.Left = A.Left + offsetx
A.top = A.top + offsety
fg = ForeColor
ForeColor = thiscolor
    DrawText UserControl.hDC, StrPtr(aa$), -1, A, DT_NOPREFIX Or DT_NOCLIP
    ForeColor = fg
End Sub
Public Property Get FontBold() As Boolean
FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal rhs As Boolean)
UserControl.FontBold = rhs
CalcNewFont
 PropertyChanged "Font"
End Property

Public Property Get charset() As Integer
charset = UserControl.Font.charset
End Property

Public Property Let charset(ByVal rhs As Integer)
UserControl.Font.charset = rhs
CalcNewFont
 PropertyChanged "Font"
End Property
Public Sub ExternalCursor(ByVal ExtSelStart, that$)

 Dim REALX As Long, REALX2 As Long, myt1
 
 myt1 = myt - scrTwips
If ExtSelStart <= 0 Then ExtSelStart = 1
                                             DrawStyle = vbSolid
             
                                   REALX = UserControlTextWidth(Mid$(that$, 1, ExtSelStart - 1)) + LeftMarginPixels * scrTwips
                                    REALX2 = scrollme + REALX
                                    If (Not marvel) And (havefocus And Not Noflashingcaret) Then
                                    ShowMyCaretInTwips REALX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips + scrTwips
                                    Else
                   DrawMode = vbInvert
                Line (REALX2, (SELECTEDITEM - topitem - 1) * myt + myt1 + mHeadlineHeightTwips + scrTwips)-(REALX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips + scrTwips), ForeColor
                 DrawMode = vbCopyPen
                                 End If
                                   If Not NoScroll Then If REALX2 > Width * 0.8 Then scrollme = scrollme - Width * 0.2: PrepareToShow 10
                                   If REALX2 < 0 Then
                              If Not NoScroll Then
                              scrollme = scrollme + Width * 0.2
                                   PrepareToShow 10
                                   End If
                                   End If

End Sub

Private Sub FindRealCursor(ByVal tothere As Long)
' from listindex to tothere
' No center text
tothere = tothere - 1
If tothere = listindex Then Exit Sub
Dim thatwidth As Long, c$, dummy1 As Long
If SelStart < 1 Then
c$ = List(listindex)
Else
c$ = Left$(List(listindex), SelStart - 1)
End If
thatwidth = UserControlTextWidth(c$) + LeftMarginPixels * scrTwips
REALCUR List(tothere), thatwidth, (dummy1), mSelstart, True
mSelstart = mSelstart + 1
End Sub

Public Sub Shutdown()
waitforparent = False
Timer1.Interval = 10000
Timer2.Interval = 10000
Timer3.Interval = 10000
Enabled = False
End Sub

Public Sub DragNow()
marvel = True
UserControl.OLEDrag
marvel = False
End Sub
Private Sub MarkWord()
If listindex < 0 Then Exit Sub
Dim one$
Dim mline$, pos As Long, Epos As Long, oldselstart As Long
mline$ = List(listindex)
'Enabled = False
pos = SelStart
If pos <> 0 Then
Dim mypos As Long, ogt As String, this$
Epos = pos
Do While pos > 0
If InStr(1, WordCharLeft, Mid$(mline$, pos, 1)) Then Exit Do
pos = pos - 1
Loop
Do While Epos <= Len(mline$)
one$ = Mid$(mline$, Epos, 1)
If InStr(1, WordCharRightButIncluded, one$) Then Epos = Epos + 1: Exit Do
If InStr(1, WordCharRight, one$) Then Exit Do
Epos = Epos + 1
Loop
If (Epos - pos - 1) > 0 Then
this$ = Mid$(mline$, pos + 1, Epos - pos - 1)
RaiseEvent WordMarked(this$)
If this = "" Then Exit Sub
oldselstart = SelStart
MarkNext = 0
If (oldselstart - pos - 1) > (Epos - oldselstart) Then
SelStart = pos + 1
RaiseEvent MarkIn
MarkNext = 1
SelStart = Epos
RaiseEvent MarkOut
Else
SelStart = Epos
RaiseEvent MarkIn
SelStart = pos + 1
MarkNext = 1
RaiseEvent MarkOut
SelStart = pos + 1
End If

ShowMe2
End If
End If
'Enabled = True

End Sub
Public Sub MarkALL()
MarkNext = 0
ListindexPrivateUse = 0
SelStart = 0
RaiseEvent selected(listindex + 1)
RaiseEvent MarkIn
MarkNext = 1
ListindexPrivateUse = listcount - 1
SelStart = Len(List(listindex)) + 1
RaiseEvent selected(listindex + 1)
RaiseEvent MarkOut
ShowMe2
End Sub
Public Sub ShowPan()
Dim ll As Long
If listcount > 0 Then
    If listindex >= 0 Then
            If (listindex - topitem) >= 0 And (listindex - topitem) < lines Then
                    If SelStart = 0 Then
                    ll = scrollme
                    Else
                    ll = UserControlTextWidthPixels(Left$(List(listindex), SelStart)) + scrollme
                    End If
                    If ll < WidthPixels Then
                    ShowMe2
                    Exit Sub
                    ElseIf ll >= 0 Then
                    ShowMe2
                    Exit Sub
                    End If
           
            End If
    End If
End If
ShowMe
End Sub

Public Property Get mousepointer() As Integer
mousepointer = UserControl.mousepointer
End Property

Public Property Let mousepointer(ByVal rhs As Integer)
UserControl.mousepointer = rhs
End Property
Function GetKeY(ascii As Integer) As String
    Dim Buffer As String, Ret As Long
    Buffer = String$(514, 0)
    Dim r&
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      r = Val("&H" & Right(Hex(r), 4))
    Ret = GetLocaleInfo(r, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    If Ret > 0 Then
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, CLng(Val("&h" + Left$(Buffer, Ret - 1))))))
    Else
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, 1033)))
    End If
End Function

Public Function LineTopOffsetPixels()
Dim nr As RECT, A$
A$ = "fg"
CalcRect1 UserControl.hDC, A$, nr
LineTopOffsetPixels = (mytPixels - nr.Bottom) / 2
End Function


Private Sub Shape(A As Myshape, Optional Left As Long = -1, Optional top As Long = -1, Optional Width As Long = -1, Optional Height As Long = -1)
If Left <> -1 Then A.Left = Left
If top <> -1 Then A.top = top
If Width <> -1 Then A.Width = Width
If Height <> -1 Then A.Height = Height
Dim th As RECT, my_brush As Long, br2 As Long
If A.Visible Then
With th
.top = A.top / scrTwips
.Left = A.Left / scrTwips
.Bottom = .top + A.Height / scrTwips
.Right = .Left + A.Width / scrTwips
End With

 br2 = CreateSolidBrush(BarHatchColor)
   
   If A.hatchType = 1 Then

    SetBkColor UserControl.hDC, BarColor
 my_brush = CreateHatchBrush(BarHatch, BarHatchColor)

  FillRect UserControl.hDC, th, my_brush
 Else
  my_brush = CreateSolidBrush(BarColor)
  FillRect UserControl.hDC, th, my_brush
End If

FrameRect UserControl.hDC, th, br2

  DeleteObject my_brush
  DeleteObject br2
End If
End Sub

Function DoubleClickCheck(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long, ByVal Xorigin As Long, ByVal Yorigin As Long, setupxy As Long, itemline As Long) As Boolean
' doubleclick
Static Lx As Long, ly As Long
If item = itemline Then
   If Abs(x - Xorigin) < setupxy And Abs(y - Yorigin) < setupxy Then
      FloatList = False
      mousepointer = 1
            If Button = 1 Then
                doubleclick = doubleclick + 1
                ''If mlx <> lx Or mly <> ly Then
                          Lx = mlx
                          ly = mly
                      ''       doubleclick = 1
              ''  Else
                        
                 If Lx <> -1000 And ly <> -1000 Then
                        doubleclick = doubleclick + 1
                            If doubleclick > 1 Then DoubleClickCheck = True: Exit Function
    End If
                            Button = 0
                       
            ''End If
    Else
''   If mlx <> lx Or mly <> ly Then doubleclick = 0
    End If
    Else
 mlx = -1000
      mly = -1000
        doubleclick = 0
       FloatList = True
       mousepointer = 5

    End If
End If
End Function


Public Property Get Parent() As Variant
On Error GoTo there
If UserControl.Parent Is Nothing Then Exit Property
Set Parent = UserControl.Parent
there:
End Property


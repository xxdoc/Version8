VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   9765
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TextP0.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6345
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   0
      Picture         =   "TextP0.frx":0582
      ScaleHeight     =   405
      ScaleWidth      =   390
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox dSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   0
      Left            =   6465
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   675
      ScaleWidth      =   780
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox PrinterDocument1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   7995
      ScaleHeight     =   1140
      ScaleWidth      =   1185
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4965
      Visible         =   0   'False
      Width           =   1185
   End
   Begin M2000.gList gList1 
      Height          =   1575
      Left            =   7485
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   2778
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
   End
   Begin M2000.gList List1 
      Height          =   1920
      Left            =   7740
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   3387
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser view1 
      Height          =   2280
      Left            =   4620
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3135
      Visible         =   0   'False
      Width           =   3015
      ExtentX         =   5318
      ExtentY         =   4022
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox DIS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5640
      Left            =   780
      MouseIcon       =   "TextP0.frx":06CC
      MousePointer    =   1  'Arrow
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   5640
      ScaleWidth      =   5640
      TabIndex        =   5
      Top             =   300
      Visible         =   0   'False
      Width           =   5640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public fState As Long
Public lockme As Boolean
Public WithEvents TEXT1 As TextViewer
Attribute TEXT1.VB_VarHelpID = -1
Public EditTextWord As Boolean
' by default EditTextWord is false, so we look for identifiers not words
Private pad$, s$
Private LastDocTitle$, Para1 As Long, PosPara1 As Long, Para2 As Long, PosPara2 As Long, Para3 As Long, PosPara3 As Long
Public ShadowMarks As Boolean
Private nochange As Boolean
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal psString As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Public MY_BACK As New cDIBSection
Private mynum$
Dim OneOnly As Boolean
Public WithEvents HTML As HTMLDocument
Attribute HTML.VB_VarHelpID = -1
''Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
     ByVal nIndex As Integer, ByVal dwNewLong As Long) As Long
''Private Const GWL_STYLE = (-16)
Private DisStack As New basetask
Private MeStack As New basetask
Dim lookfirst As Boolean, look1 As Boolean
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
Public Point2Me As Object

Private Declare Function GetCommandLineW Lib "kernel32" () As Long

Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long

Public Function CommandW() As String
Static MM$
If MM$ <> "" Then CommandW = MM$: Exit Function
If m_bInIDE Then
MM$ = command
Else
Dim Ptr As Long: Ptr = GetCommandLineW
    If Ptr Then
        PutMem4 VarPtr(CommandW), SysAllocStringLen(Ptr, lstrlenW(Ptr))
     If AscW(CommandW) = 34 Then
       CommandW = Mid$(CommandW, InStr(CommandW, """ ") + 2)
       Else
            CommandW = Mid$(CommandW, InStr(CommandW, " ") + 1)
        End If
    End If
    End If
    If MM$ = "" And command <> "" Then CommandW = command Else CommandW = MM$
End Function


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




Private Sub DIS_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single, state As Integer)
  If TaskMaster.QueueCount > 0 Then
              TaskMaster.RestEnd1
   TaskMaster.TimerTick
TaskMaster.rest
        End If
End Sub

Private Sub dSprite_GotFocus(Index As Integer)
If lockme Then TEXT1.SetFocus: Exit Sub
End Sub

Private Sub dSprite_OLEDragOver(Index As Integer, data As DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single, state As Integer)
  If TaskMaster.QueueCount > 0 Then
              TaskMaster.RestEnd1
   TaskMaster.TimerTick
TaskMaster.rest
        End If
End Sub

Private Sub Form_Activate()
If ASKINUSE Then Me.ZOrder 1
releasemouse = True
End Sub
Private Sub Form_GotFocus()
UseEsc = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, shift As Integer)
Dim i As Long
 If List1.LeaveonChoose Then Exit Sub
clickMe = -1
i = -1
If KeyCode = vbKeyV Then
Exit Sub
End If
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
 If i <> -1 Then UKEY$ = ""
 End If

End Sub

Private Sub Form_LostFocus()
DestroyCaret
UseEsc = False
End Sub

Private Sub Form_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single, state As Integer)
  If TaskMaster.QueueCount > 0 Then
              TaskMaster.RestEnd1
   TaskMaster.TimerTick
TaskMaster.rest
        End If
End Sub

Private Sub gList1_ChangeListItem(item As Long, content As String)
Dim i As Long

If nochange Then
nochange = True
i = TEXT1.SelLength
Form1mn1Enabled = i > 1
Form1mn2Enabled = i > 1
Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
Form1sdnEnabled = i > 0 And (TEXT1.Length - TEXT1.SelStart) > i
Form1supEnabled = i > 0 And TEXT1.SelStart > i
Form1mscatEnabled = Form1sdnEnabled Or Form1supEnabled
Form1rthisEnabled = Form1mscatEnabled
nochange = False
End If
End Sub

Private Sub gList1_ChangeSelStart(thisselstart As Long)
Dim i As Long

If gList1.Enabled Then
i = TEXT1.SelLength
Form1mn1Enabled = i > 1
Form1mn2Enabled = i > 1
Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
Form1sdnEnabled = i > 0 And (TEXT1.Length - TEXT1.SelStart) > i
Form1supEnabled = i > 0 And TEXT1.SelStart > i
Form1mscatEnabled = Form1sdnEnabled Or Form1supEnabled
Form1rthisEnabled = Form1mscatEnabled
End If
End Sub

Private Sub gList1_GetBackPicture(pic As Object)
Set pic = Point2Me
End Sub

Private Sub gList1_HeaderSelected(Button As Integer)
Dim i As Long

If Not gList1.Enabled Then Exit Sub

If TEXT1.UsedAsTextBox Then Exit Sub
i = TEXT1.SelLength
Form1mn1Enabled = i > 1
Form1mn2Enabled = i > 1
Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
Form1sdnEnabled = i > 0 And (TEXT1.Length - TEXT1.SelStart) > TEXT1.SelLength
Form1supEnabled = i > 0 And TEXT1.SelStart > TEXT1.SelLength
Form1mscatEnabled = Form1sdnEnabled Or Form1supEnabled
Form1rthisEnabled = Form1mscatEnabled

MyPopUp.Up

''Form1.PopupMenu Form1.aaaa
End Sub

Private Sub gList1_KeyDownAfter(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyTab Then
KeyCode = 0
End If
End Sub

Private Sub gList1_MarkOut()
Pack1
End Sub
Public Sub Pack1()
Dim i As Long

i = TEXT1.SelLength
Form1mn1Enabled = i > 1
Form1mn2Enabled = i > 1
Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
Form1sdnEnabled = i > 0 And (TEXT1.Length - TEXT1.SelStart) > i
Form1supEnabled = i > 0 And TEXT1.SelStart > i
Form1mscatEnabled = Form1sdnEnabled Or Form1supEnabled
Form1rthisEnabled = Form1mscatEnabled
End Sub

Private Sub gList1_SyncKeyboard111(KeyAscii As Integer)
If KeyAscii = 9 Then KeyAscii = 0: Exit Sub
If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
End Sub
Private Sub gList1_OutPopUp(x As Single, y As Single, myButton As Integer)
If TEXT1.UsedAsTextBox Then Exit Sub
MyPopUp.Up x + gList1.Left, y + gList1.top

''Form1.PopupMenu Form1.aaaa, , X + gList1.Left, Y + gList1.top
myButton = 0
End Sub

Public Sub helpmeSub()
If Not EditTextWord Then
If Trim(TEXT1.SelText) <> "" Then
ffhelp myUcase(Trim(TEXT1.SelText), True)
Else
vHelp
End If
Else
If abt Then
feedback$ = Trim(TEXT1.SelText)
feednow$ = FeedbackExec$
CallGlobal feednow$
Else
vHelp
End If
End If
End Sub

Private Sub ffhelp(A$)
If Left$(A$, 1) < "Α" Then
fHelp basestack1, A$, True
Else
fHelp basestack1, A$
End If
End Sub



Private Sub List1_ListError(code As Long)
Dim DUMMY As Long
List1.listindex = -1
List1.LeaveonChoose = False
List1.Visible = False
If List1.Tag <> "" Then
If QRY Or GFQRY Then
Else

DUMMY = interpret(basestack1, List1.Tag)
Me.KeyPreview = True
End If
End If
MyEr "Menu Error " & CStr(code), "Λάθος στην ΕΠΙΛΟΓΗ αριθμός " & CStr(code)
End Sub


Private Function HTML_oncontextmenu() As Boolean
HTML_oncontextmenu = False
End Function



Private Sub HTML_onkeydown()
Select Case view1.Document.parentWindow.event.KeyCode
Case vbKeyF1
IEUP homepage$
Form1.KeyPreview = False
Case vbKeyEscape
If escok Then
IEUP ""
While KeyPressed(&H1B)
MyDoEvents
refresh
Wend
INK$ = ""
UINK$ = ""
End If

End Select
End Sub

Private Sub list1_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = List1.listindex Then
' is our cursor
'If item = 0 Then Stop
List1.FillThere thisHDC, thisrect, &HFFFFFF, -List1.LeftMarginPixels  ' or black in reverse
List1.WriteThere thisrect, List1.List(item), List1.PanPos / dv15, List1.addpixels / 2, 0
skip = True
Else
skip = False
End If
End Sub


Private Sub List1_PanLeftRight(Direction As Boolean)
Dim DUMMY As Boolean
If List1.Tag <> "" Then
If QRY Or GFQRY Then
Else

DUMMY = interpret(basestack1, List1.Tag)
Me.KeyPreview = True
End If
Else
List1.LeaveonChoose = False
List1.Visible = False

End If
End Sub

Private Sub List1_Selected2(item As Long)
Dim DUMMY As Boolean

If List1.Tag <> "" Then
If QRY Or GFQRY Then
Else

DUMMY = interpret(basestack1, List1.Tag)
Me.KeyPreview = True
End If
Else
List1.LeaveonChoose = False
List1.Visible = False
End If
End Sub

Private Sub List1_SyncKeyboard(item As Integer)
'refresh
MyDoEvents2
INK$ = INK$ & Chr(item)
End Sub

Public Sub mn5sub()
 If Not EditTextWord Then
 ' check if { } is ok...
 If Not blockCheck(TEXT1.Text, DialogLang) Then Exit Sub
 End If

CancelEDIT = True
MyDoEvents
NOEDIT = True
End Sub

Public Sub mscatsub()
''
Dim l As Long, w As Long, s$, TempLcid As Long, OldLcid As Long
Dim eL As Long, eW As Long, SAFETY As Long, tt$

w = TEXT1.mdoc.MarkParagraphID
eW = w
TEXT1.SelStartSilent = TEXT1.SelStart  'MOVE CHARPOS TO SELSTART

eL = TEXT1.Charpos  ' charpos maybe is in the start or the end of block
s$ = TEXT1.SelText
OldLcid = TEXT1.mdoc.LCID
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then TEXT1.mdoc.LCID = TempLcid

l = eL + 1
If EditTextWord Then
Do
If TEXT1.mdoc.FindWord(s$, True, w, l) Then
tt$ = TEXT1.mdoc.TextParagraph(w)
Mid$(tt$, l, Len(s$)) = s$
TEXT1.mdoc.ReWritePara w, tt$
Else
w = 1
l = 0
SAFETY = SAFETY + 1
End If
Loop Until (w = eW And l = eL) Or SAFETY = 2

Else
Do
If TEXT1.mdoc.FindIdentifier(s$, True, w, l) Then
tt$ = TEXT1.mdoc.TextParagraph(w)
Mid$(tt$, l, Len(s$)) = s$
TEXT1.mdoc.TextParagraph(w) = tt$
Else
w = 1
l = 0
SAFETY = SAFETY + 1
End If
Loop Until (w = eW And l = eL) Or SAFETY = 2

End If
TEXT1.mdoc.LCID = OldLcid
TEXT1.mdoc.WrapAgainColor
TEXT1.mdoc.WrapAgain

TEXT1.Render
End Sub

Public Sub rthissub()
Dim l As Long, w As Long, s$, TempLcid As Long, OldLcid As Long
Dim eL As Long, eW As Long, SAFETY As Long, tt$, w1 As Long, i1 As Long
Dim neo$, mDoc10 As Document, addthat As Long
w = TEXT1.mdoc.MarkParagraphID
eW = w
TEXT1.SelStartSilent = TEXT1.SelStart  'MOVE CHARPOS TO SELSTART
eL = TEXT1.Charpos  ' charpos maybe is in the start or the end of block
s$ = Trim$(TEXT1.SelText)
TEXT1.SelStartSilent = TEXT1.SelStart
eL = TEXT1.Charpos  ' charpos maybe is in the start or the end of block

If pagio$ = "GREEK" Then
neo$ = InputBoxN("Αλλαγή Λέξης", "Συγγραφή Κειμένου", s$)
Else
neo$ = InputBoxN("Replace Word", "Text Editor", s$)
End If
If neo$ = "" Then Exit Sub
OldLcid = TEXT1.mdoc.LCID
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then TEXT1.mdoc.LCID = TempLcid
If Len(neo$) >= Len(s$) Then
    Set mDoc10 = New Document
    mDoc10 = neo$
    w1 = 0
    i1 = 0
    
    If EditTextWord Then
        If mDoc10.FindWord(s$, True, w1, i1) Then addthat = i1 - 1: If Len(neo$) = Len(s$) And addthat = 0 Then Exit Sub
    Else
        If mDoc10.FindIdentifier(s$, True, w1, i1) Then addthat = i1 - 1: If Len(neo$) = Len(s$) And addthat = 0 Then Exit Sub
    End If
    
End If

i1 = eL
l = i1 + addthat
w1 = w
If EditTextWord Then
TEXT1.glistN.dropkey = True
Do
If TEXT1.mdoc.FindWord(s$, True, w, l) Then
If SAFETY And w = w1 Then

If l = i1 Then
 TEXT1.SelLengthSilent = 0
TEXT1.mdoc.MarkParagraphID = w
TEXT1.glistN.Enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.Enabled = True
TEXT1.SelLength = Len(s$)
TEXT1.AddUndo ""
TEXT1.SelText = neo$
TEXT1.RemoveUndo neo$
Exit Do
ElseIf l - addthat < i1 Then
i1 = i1 + Len(neo$) - Len(s$)
Else

End If
End If
TEXT1.SelLengthSilent = 0
TEXT1.mdoc.MarkParagraphID = w
TEXT1.glistN.Enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.Enabled = True
TEXT1.SelLength = Len(s$)
TEXT1.AddUndo ""
TEXT1.SelText = neo$
TEXT1.RemoveUndo neo$
l = l + Len(neo$)

Else
w = 1
l = 0
SAFETY = SAFETY + 1
End If
Loop Until SAFETY = 2
TEXT1.glistN.dropkey = False

Else
''If l > 0 Then l = l - 1
TEXT1.glistN.dropkey = True
Do
If TEXT1.mdoc.FindIdentifier(s$, True, w, l) Then
If SAFETY And w = w1 Then

If l = i1 Then
 TEXT1.SelLengthSilent = 0
TEXT1.mdoc.MarkParagraphID = w
TEXT1.glistN.Enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.Enabled = True
TEXT1.SelLength = Len(s$)
TEXT1.AddUndo ""
TEXT1.SelText = neo$
TEXT1.RemoveUndo neo$
Exit Do
ElseIf l - addthat < i1 Then
i1 = i1 + Len(neo$) - Len(s$)
Else

End If
End If
TEXT1.SelLengthSilent = 0
TEXT1.mdoc.MarkParagraphID = w
TEXT1.glistN.Enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.Enabled = True
TEXT1.SelLength = Len(s$)
TEXT1.AddUndo ""
TEXT1.SelText = neo$
TEXT1.RemoveUndo neo$
l = l + Len(neo$)

Else
w = 1
l = 0
SAFETY = SAFETY + 1
End If
Loop Until SAFETY = 2
TEXT1.glistN.dropkey = False
End If
TEXT1.mdoc.LCID = OldLcid
TEXT1.mdoc.WrapAgainColor
TEXT1.Render

End Sub

Public Sub sdnSub()
Dim b$
s$ = TEXT1.SelText
If s$ = "" Then s$ = b$
SearchDown s$
End Sub
Sub SearchDown(s$, Optional anystr As Boolean = False)
Dim l As Long, w As Long, TempLcid As Long, OldLcid As Long
w = TEXT1.mdoc.MarkParagraphID   ' this is the not the order
TEXT1.SelStartSilent = TEXT1.SelStart
l = TEXT1.Charpos + 1

OldLcid = TEXT1.mdoc.LCID
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then TEXT1.mdoc.LCID = TempLcid
If EditTextWord Or anystr Then
    If anystr Then
  If Not TEXT1.mdoc.FindStrDown(s$, w, l) Then GoTo sdnOut
  Else
    If Not TEXT1.mdoc.FindWord(s$, True, w, l) Then GoTo sdnOut
    End If
Else
    If Not TEXT1.mdoc.FindIdentifier(s$, True, w, l) Then GoTo sdnOut
End If
TEXT1.SelLengthSilent = 0
TEXT1.mdoc.MarkParagraphID = w
TEXT1.glistN.Enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.Enabled = True
TEXT1.SelLength = Len(s$)
sdnOut:
TEXT1.mdoc.LCID = OldLcid
End Sub

Public Sub supsub()
Dim b$
s$ = TEXT1.SelText
If s$ = "" Then s$ = b$
Searchup s$
End Sub
Sub Searchup(s$, Optional anystr As Boolean = False)
Dim l As Long, w As Long, TempLcid As Long, OldLcid As Long
w = TEXT1.mdoc.MarkParagraphID
TEXT1.SelStartSilent = TEXT1.SelStart - (TEXT1.SelLength > 1)
l = TEXT1.Charpos + 1
OldLcid = TEXT1.mdoc.LCID
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then TEXT1.mdoc.LCID = TempLcid
If EditTextWord Or anystr Then
   If anystr Then
   If Not TEXT1.mdoc.FindStrUp(s$, w, l) Then GoTo sdupOut
   Else
       If Not TEXT1.mdoc.FindWord(s$, False, w, l) Then GoTo sdupOut
    End If
Else
    If Not TEXT1.mdoc.FindIdentifier(s$, False, w, l) Then GoTo sdupOut
End If
TEXT1.SelLengthSilent = 0
TEXT1.mdoc.MarkParagraphID = w
TEXT1.glistN.Enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.Enabled = True
TEXT1.SelLength = Len(s$)
sdupOut:
TEXT1.mdoc.LCID = OldLcid
End Sub
Public Function InIDECheck() As Boolean
    m_bInIDE = True
    InIDECheck = True
End Function


Private Sub DIS_GotFocus()
If lockme Then TEXT1.SetFocus: Exit Sub
Dim dX As Long, dy As Long
clickMe2 = -1
End Sub

Private Sub DIS_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyPause Then
Form_KeyDown KeyCode, shift
End If
If Not NOEDIT Then
End If

End Sub
Public Sub GiveASoftBreak(Sorry As Boolean)
clickMe2 = -1
' Try first with escape
If Sorry Then
Form_KeyDown vbKeyPause, (0)
Else  'CTRL C
Form_KeyDown &HFFFE, (0)
End If

End Sub

Private Sub DIS_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
If lockme Then Exit Sub
MOUB = Button

If Not NoAction Then
NoAction = True
''Dim xx&, yy&, , ox&, oy&
''SetText DIS
''GetXY DIS, ox&, oy&
Dim sel&
If Button > 0 And Targets Then


sel& = ScanTarget(q(), CLng(x), CLng(y), 0)

If sel& >= 0 Then

If Button = 1 Then


Select Case q(sel&).Id Mod 100
Case Is < 10
If Not interpret(DisStack, (q(sel&).Comm)) Then Beep
Case Else
INK$ = q(sel&).Comm
End Select


Else

End If

End If
If Not nomore Then NoAction = False

End If
End If

End Sub

Private Sub DIS_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If lockme Then Exit Sub

MOUB = Button

If NOEDIT = True And (exWnd = 0 Or Button) Then
Me.KeyPreview = True
End If
End Sub

Private Sub DIS_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)

If lockme Then
If Not NOEDIT Then TEXT1.SetFocus
Exit Sub
End If
MOUB = 0
End Sub






Private Sub dSprite_MouseDown(Index As Integer, Button As Integer, shift As Integer, x As Single, y As Single)
Dim p As Long, u2 As Long
If lockme Then Exit Sub
 MOUB = Button
'If button > 0 Then

' look for one and only target
If Not NoAction Then
NoAction = True
Dim sel&
p = Val("0" & dSprite(Index).Tag)
With players(p)
    u2 = .uMineLineSpace * 2

        If Button > 0 And Targets Then

        sel& = ScanTarget(q(), CLng(x), CLng(y), Index)
            If sel& >= 0 Then
                If Button = 1 Then
                '' If QRY Then LCTC dSprite(Index), oy&, ox&, ins& Else LCT dSprite(Index), oy&, ox&
                Select Case q(sel&).Id Mod 100
                Case Is < 10
                If Not interpret(DisStack, "LAYER " & dSprite(Index).Tag + " {" + vbCrLf + q(sel&).Comm + vbCrLf & "}") Then Beep
                Case Else
                INK$ = q(sel&).Comm
                End Select
               ''' If QRY Then LCTC dSprite(Index), oy&, ox&, ins& Else LCT dSprite(Index), oy&, ox&

End If
End If


If Not nomore Then NoAction = False

End If
End With
End If

End Sub

Private Sub dSprite_MouseMove(Index As Integer, Button As Integer, shift As Integer, x As Single, y As Single)
If lockme Then Exit Sub
MOUB = Button
If NOEDIT = True And (exWnd = 0 Or Button) Then
Me.KeyPreview = True
End If
End Sub

Private Sub dSprite_MouseUp(Index As Integer, Button As Integer, shift As Integer, x As Single, y As Single)
If lockme Then Exit Sub
MOUB = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
Dim i As Long
Form1.Font.charset = GetCharSet(GetCodePage(GetLCIDFromKeyboard))
'If BLOCKkey Then Stop
Static ctrl As Boolean, noentrance As Boolean
If KeyCode = 13 And List1.Visible And (Not List1.LeaveonChoose) And Not QRY Then
KeyCode = 0
List1.PressSoft
Exit Sub
End If
If KeyCode = 13 And trace Then
If Not STq Then STbyST = True: KeyCode = 0

End If
clickMe = HighLow(CLng(shift), CLng(KeyCode))
If clickMe2 = -2 Then clickMe2 = clickMe
If clickMe = 27 And escok Then
NOEXECUTION = True
If exWnd <> 0 Then
MyDoEvents
    nnn$ = "bye bye"
    exWnd = 0
    
    End If
If view1.Visible Then
MyDoEvents

view1.Navigate "about:blank"
Sleep 50
view1.Visible = False

 End If

End If

If clickMe2 <> -1 Then KeyCode = 0: Exit Sub

If BLOCKkey Then Exit Sub
If noentrance Then
KeyCode = 0
Exit Sub
End If
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


Select Case KeyCode
Case vbKeyE, vbKeyD
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = INK$ & "ΣΥΓΓΡΑΦΗ "
Else
INK$ = INK$ & "EDIT "
End If
End If
End If
Case vbKeyA
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If LASTPROG$ <> "" Then
If pagio$ = "GREEK" Then
INK$ = "ΣΩΣΕ ΕΝΤΟΛΗ$" & vbCr
Else
INK$ = "SAVE COMMAND$" & vbCr
End If
End If
End If
End If
Case vbKeyS
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = "ΣΩΣΕ "
Else
INK$ = "SAVE "
End If
End If
End If
Case vbKeyL
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = "ΛΙΣΤΑ "
Else
INK$ = "LOAD "
End If
End If
End If
Case vbKeyF
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = "ΦΟΡΤΩΣΕ "
Else
INK$ = "FILES "
End If
End If
End If
Case vbKeyP, vbKeyT
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = "ΤΥΠΩΣΕ "
Else
INK$ = "PRINT "
End If
End If
End If
Case vbKeyM
If ctrl And (shift And &H2) = 2 Then
 If QRY Then
 If pagio$ = "GREEK" Then
INK$ = "ΤΜΗΜΑΤΑ "
Else
INK$ = "MODULES "
End If
 End If
End If
Case vbKeyU
If ctrl And (shift And &H2) = 2 Then
 If QRY Then
 If pagio$ = "GREEK" Then
INK$ = "ΡΥΘΜΙΣΕΙΣ " + vbCr
Else
INK$ = "SETTINGS " + vbCr
End If
 End If
End If
Case vbKeyN
If ctrl And (shift And &H2) = 2 Then
 If QRY Then
 If pagio$ = "GREEK" Then
INK$ = "ΤΜΗΜΑΤΑ ? " + vbCr
Else
INK$ = "MODULES ? " + vbCr
End If
 End If
End If
Case vbKeyTab
    If (shift And 1) = 1 Then
    INK$ = INK$ & Chr$(6)
    KeyCode = 0
    End If
Case vbKeyV
    If ctrl And (shift And &H2) = 2 Then
        pad$ = GetTextData(CF_UNICODETEXT)
    
        If pad$ <> "" Then
          '  For i = 1 To Len(pad$)
              '  If Asc(Mid$(pad$, i, 1)) > 31 Then
                INK$ = pad$
              '  Exit Sub
             '   End If
           ' Next i
        End If
    End If
KeyCode = 0
Exit Sub
Case vbKeyC, &HFFFE
If (ctrl And (shift And &H2) = 2) Or KeyCode = &HFFFE Then
If QRY Then
INK$ = INK$ & "CLS" & Chr$(13)
Else
KeyCode = 0
If Form4.Visible Then
Form4.Visible = False
    If TEXT1.Visible Then
        TEXT1.SetFocus
    Else
        Form1.SetFocus
    End If
End If
EXECSTOP
End If
End If
Case vbKeyPause  '(this is the break key!!!!!'
If QRY Or GFQRY Then
If Form4.Visible Then Form4.Visible = False
i = MOUT
If ASKINUSE Then
If BreakMe Then Exit Sub
Unload NeoMsgBox: ASKINUSE = False: Exit Sub
End If
BreakMe = True
If MsgBoxN("Break Key - Hard Reset" + vbCrLf + "Μ2000 - Execution Stop / Τερματισμός Εκτέλεσης", vbYesNo, MesTitle$) <> vbNo Then

MOUT = i

If AVIRUN Then AVI.GETLOST
On Error Resume Next
noentrance = True
NoAction = True
If Not TaskMaster Is Nothing Then TaskMaster.Dispose: MyEr "", ""

If Me.Visible Then Me.SetFocus
closeAll ' we closed all files
QRY = False
GFQRY = False
escok = True
INK$ = Chr$(27) + Chr$(27)
If MOUT = False Then
NOEXECUTION = True
MOUT = True
Else
MOUT = False
End If
If List1.Visible Then
List1.Tag = ""
List1.Visible = False
List1.LeaveonChoose = False
INK$ = ""
End If
noentrance = False

End If
BreakMe = False
End If
KeyCode = 0
Case vbKeyLeft
INK$ = INK$ & Chr(0) + Chr(75)
Case vbKeyRight
INK$ = INK$ & Chr(0) + Chr(77)
Case vbKeyUp
INK$ = INK$ & Chr(0) + Chr(72)
Case vbKeyDown
INK$ = INK$ & Chr(0) + Chr(80)
Case vbKeyInsert
INK$ = INK$ & Chr(0) + Chr(82)
Case vbKeyDelete
INK$ = INK$ & Chr(0) + Chr(83)
Case vbKeyEscape
If List1.LeaveonChoose Then Exit Sub
INK$ = INK$ & Chr(27)
If escok Then
If AVIRUN Then
AVI.GETLOST
End If
NOEXECUTION = True
End If
Case vbKeyF1 To vbKeyF12
If FKey >= 0 Then FKey = KeyCode - vbKeyF1 + 1
If Abs(FKey) = 1 And ctrl And (shift And &H2) = 2 Then
FKey = 0: KeyCode = 0: vHelp
ElseIf FKey = 1 And (shift And 1) Then
FKey = 13
ElseIf FKey = 4 And ctrl And QRY Then
interpret DisStack, "END"
End If

Case vbKeyControl

ctrl = True
KeyCode = 0
Exit Sub
Case Else
If ctrl And (shift And &H2) = 2 And lckfrm = 0 And KeyCode <> 3 And KeyCode <> 16 Then
If escok Then
STq = False
STEXIT = False
STbyST = True
Form2.Show , Form1
Form2.Label1(0) = HERE$
Form2.Label1(1) = "..."
Form2.Label1(2) = "..."
    Form2.gList3(2).BackColor = &H3B3B3B
    TestShowCode = False
     TestShowSub = ""
 TestShowStart = 0
     Set Form2.Process = basestack1
   stackshow basestack1
Form1.Show , Form5
trace = True
End If
End If
End Select

ctrl = False
 If List1.LeaveonChoose Then Exit Sub
 If KeyCode = 91 Then Exit Sub
i = GetLastKeyPressed
 If i <> -1 And i <> 94 Then UKEY$ = ChrW(i) Else If i <> -1 Then UKEY$ = ""
 If List1.Visible Then
 Else
KeyCode = 0
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 And view1.Visible Then view1.SetFocus: KeyAscii = 0: Exit Sub
If clickMe2 = -2 And clickMe <> -1 Then clickMe2 = clickMe
If clickMe2 <> -1 And Not List1.LeaveonChoose Then KeyAscii = 0: Exit Sub

If Right$(INK$, 1) = Chr$(6) And KeyAscii = 9 Then

Else
If mynum$ <> "" Then
If mynum$ <> "" Then Exit Sub
End If
If UKEY$ <> "" Then
INK$ = INK$ & UKEY$
UKEY$ = ""
Else
If KeyAscii = 22 Then
KeyAscii = 0
Else
INK$ = INK$ & GetKeY(KeyAscii)
End If
End If
End If
End Sub
Private Sub EXECSTOP()
Dim iamhere As Boolean
If iamhere Then Exit Sub
iamhere = True
NOEXECUTION = False
If MsgBoxN("[Ctrl + C] Μ2000 - Execution Stop / Τερματισμός Εκτέλεσης", vbYesNo, MesTitle$) = vbYes Then
extreme = False
If AVIRUN Then
AVI.GETLOST
End If
 On Error Resume Next
'noentrance = True
If TaskMaster.QueueCount > 0 Then TaskMaster.Dispose
NoAction = True
Close ' we closed all files
escok = True
If QRY Then
INK$ = Chr$(27) + Chr$(27)
MyDoEvents
End If
QRY = False
RRCOUNTER = 0
REFRESHRATE = 25
INK$ = Chr$(27) + Chr$(27)
NOEXECUTION = True
'trace = True
If lckfrm > 0 Then
If MOUT = False Then
NOEXECUTION = True
MOUT = True
Else
MOUT = False
End If
End If

'noentrance = False
End If
EmptyClipboard

iamhere = False
End Sub

Public Sub HideMouse()
MouseIcon = Form1.Picture2.Picture
mousepointer = 99
End Sub

Private Sub Form_Load()

Set TEXT1 = New TextViewer

Set TEXT1.Container = gList1

TEXT1.glistN.DragEnabled = False ' only drop - we can change this from popup menu
TEXT1.glistN.Enabled = False
TEXT1.FileName = ""
TEXT1.glistN.addpixels = 0
TEXT1.showparagraph = False
TEXT1.EditDoc = True

TEXT1.glistN.LeftMarginPixels = 10
With TEXT1.glistN
.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", "@")
.WordCharRight = ConCat(":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^")
.WordCharRightButIncluded = "("

End With
List1.LeftMarginPixels = 4
List1.NoPanRight = False
List1.SingleLineSlide = True
Dim s$
Set DisStack.Owner = DIS

List1.BypassLeaveonChoose = False

Set MeStack.Owner = Me


ThereIsAPrinter = IsPrinter
If ThereIsAPrinter Then

pname = Printer.DeviceName
port = Printer.port

End If

dset
If LoadFont(GetCurDir(True) & "TT6492M_.TTF") Then
defFontname = "monospace 821 greek bt"
MYFONT = defFontname
myBold = True
Else
MYFONT = "Tahoma"
defFontname = MYFONT
myBold = False
End If
myCharSet = 0
With Form1
.Font.name = MYFONT
.Font.Strikethrough = False
.Font.Underline = False
.Font.bold = myBold
MYFONT = .Font.name
    .Font.charset = myCharSet
    .DIS.Font.charset = myCharSet
    .DIS.Font.name = MYFONT
    .DIS.Font.bold = myBold
    .TEXT1.Font.charset = myCharSet
    .TEXT1.Font.name = MYFONT
    .TEXT1.Font.bold = myBold
    
    .List1.charset = myCharSet
    .List1.Font.name = MYFONT
    .List1.FontBold = myBold
     
End With


''DIS.Visible = False
Debug.Assert (InIDECheck = True)
s$ = CommandW
If Not ISSTRINGA(s$, cLine) Then
cLine = mylcasefILE(Trim(s$))
Else
cLine = mylcasefILE(cLine)
End If
While Left$(cLine, 1) = Chr(34) And Right$(cLine, 1) = Chr(34) And Len(cLine) > 2
cLine = Mid$(cLine, 2, Len(cLine) - 2)
Wend
If ExtractType(cLine) <> "gsb" Then cLine = ""
If cLine <> "" Then
para$ = ExtractPath(cLine) + ExtractName(cLine)
cLine = Trim$(Mid$(cLine, Len(para$) + 1))
s$ = cLine + " " + s$
cLine = para$
ElseIf s$ <> "" Then
para$ = Trim$(s$)
End If



Switches para$
    
    l_complete = True
  
 
111:

  On Error Resume Next
  Dim i As Long
  
      For i = 0 To Controls.Count - 1
     If Typename(Controls(i)) <> "Menu" Then Controls(i).TabStop = False
      Next i
 
End Sub




Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
If lockme Then Exit Sub
MOUB = Button
clickMe2 = -1

If NoAction Then Exit Sub
NoAction = True
Dim sel&

If Button > 0 And Targets Then
sel& = ScanTarget(q(), CLng(x), CLng(y), -1)

If sel& >= 0 Then

If Button = 1 Then


Select Case q(sel&).Id Mod 100
Case Is < 10

If Not interpret(MeStack, (q(sel&).Comm)) Then Beep
Case Else
INK$ = q(sel&).Comm
End Select


Else

End If

End If
If Not nomore Then NoAction = False

End If

End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If lockme Then Exit Sub
If Button > 0 Then MOUB = Button
'moux = x
'mouy = y
'If Not toback Then Exit Sub
If NOEDIT = True And (exWnd = 0 Or Button) Then
Me.KeyPreview = True
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If lockme Then Exit Sub
 MOUB = 0
'moux = x
'mouy = y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = NoAction
End Sub

Private Sub iForm_Resize()
DIS.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Public Sub Up()
UpdateWindow hWnd
End Sub

Sub something()

Set basestack1.Owner = DIS
Set DisStack.Owner = DIS

On Error Resume Next
Const HWND_BROADCAST = &HFFFF&
Const WM_FONTCHANGE = &H1D
Dim pn As Long, A As New cDIBSection
AutoRedraw = True
If OneOnly Then Exit Sub
OneOnly = True


escok = False
Sleep 10
FK$(13) = "ΣΥΓΓΡΑΦΕΑΣ"
'Me.WindowState = 2
''Hide
Me.WindowState = 0
Sleep 10
If Me.WindowState = 0 Then
Me.Move 0, ScrY(), (ScrX() - 1), (ScrY() - 1)
End If
Sleep 10

basestack1.myCharSet = 0
    Font.charset = basestack1.myCharSet
    basestack1.Owner.Font.charset = basestack1.myCharSet
    TEXT1.Font.charset = basestack1.myCharSet
    List1.Font.charset = basestack1.myCharSet
   ' List2.Font.CharSet = basestack1.myCharSet

basestack1.Owner.Move 0, 0, ScaleWidth, ScaleHeight


''mmx = basestack1.Owner.Width
''mmy = basestack1.Owner.Height

If NoAction Then Exit Sub
Dim DUMMY As Boolean, i
NOEXECUTION = False
'
myBreak basestack1

If cLine <> "" Then
LASTPROG$ = cLine
'
Original basestack1, " : NEW : CLEAR : " & "TITLE " & """" + ExtractNameOnly(cLine) + """" & ", 0"
Else
Original basestack1, " : NEW : CLEAR :" & "TITLE " & """" & "M2000" & """"
End If

s_complete = True

Dim helpcnt As Long, qq$

''MyDoEvents

Dim mybasket As basket
mybasket = players(DisForm)
PlaceBasket DIS, mybasket
Do
Do
escok = True
MOUT = True

MyDoEvents

 If cLine = "" Then

   If trace Then
    Form2.Label1(0) = HERE$
    Form2.Label1(1) = "..."
    Form2.Label1(2) = "..."
    Form2.gList3(2).BackColor = &H3B3B3B
    TestShowCode = False
     TestShowSub = ""
 TestShowStart = 0
      Set Form2.Process = basestack1
   stackshow basestack1
Form2.ComputeNow

   
    End If
   
     If Not Form1.Visible Then Form1.Show , Form5
    If Not releasemouse Then Form1.SetFocus
  
    NORUN1 = False

 players(DisForm) = mybasket
 ''reset refresh system
  REFRESHRATE = 25
  k1 = 0
    QUERY basestack1, ">", qq$, (mybasket.mx * 4), True
      mybasket = players(DisForm)
If basestack1.Owner.Visible = True Then basestack1.Owner.refresh Else basestack1.Owner.Visible = True

    FK$(13) = "ΣΥΓΓΡΑΦΕΑΣ"
    INK$ = ""
    mybasket.pageframe = 0
    MYSCRnum2stop = holdcontrol(DIS, mybasket)
    HoldReset 1, mybasket
If CommandW = "" And qq$ = "" Then helpcnt = helpcnt + 1

        If helpcnt > 4 Then
    If basestack1.Owner.Font.charset <> 161 Then
    qq$ = " HELP": helpcnt = -100000
    Else
    qq$ = " ΒΟΗΘΕΙΑ": helpcnt = -100000
    End If
    End If
 crNew basestack1, mybasket

 
Else
   sHelp "", "", 0, 0
   qq$ = "LOAD" & """" + cLine + """"
   If Len(Left$(cLine, rinstr(cLine, "\"))) > 0 Then

    mcd = Left$(cLine, rinstr(cLine, "\"))
   End If
   cLine = ""
End If

If Not MOUT Then NOEXECUTION = False: ResetBreak: MOUT = interpret(basestack1, "START"): qq$ = "": mybasket = players(DisForm)

Loop Until qq$ <> ""

NoAction = True
NOEXECUTION = False
basestack1.toprinter = False
MOUT = False
ResetBreak
players(DisForm) = mybasket
If Not interpret(basestack1, qq$) Then
mybasket = players(DisForm)
If NERR Then Exit Do
    basestack1.toprinter = False
    If MOUT Then
            NOEXECUTION = False
            ResetBreak
            MOUT = interpret(basestack1, "START"): qq$ = ""
            
            mybasket = players(DisForm)
            MOUT = False
        Else
        
        If NOEXECUTION Then
                closeAll
                mybasket = players(DisForm)
                PlainBaSket DIS, mybasket, "ESC " & qq$
        Else
        ' look last error
                If Left$(LastErName & " ", 1) <> "?" Then
                        closeAll
                        If basestack1.Owner.Font.charset <> 161 Then
                        wwPlain basestack1, mybasket, " ? " & LastErName, basestack1.Owner.Width, 1000, True
                        If Left$(FK$(13), 4) = "EDIT" Then crNew basestack1, mybasket: wwPlain basestack1, mybasket, "Use SHIFT F1, edit, ESC to return", basestack1.Owner.Width, 1000, True
                        Else
                        wwPlain basestack1, mybasket, " ? " & LastErNameGR, basestack1.Owner.Width, 1000, True
                        If Left$(FK$(13), 4) = "EDIT" Then crNew basestack1, mybasket: wwPlain basestack1, mybasket, "Με το SHIFT F1 διορθώνεις, ESC επιστρέφεις", basestack1.Owner.Width, 1000, True
                        End If
                            
                            LastErName = "?" & LastErName
                            LastErNameGR = "?" & LastErNameGR
                Else
                        mybasket = players(DisForm)
                        wwPlain basestack1, mybasket, " ? " & qq$, basestack1.Owner.Width, 1000, True
                End If
        End If
        crNew basestack1, mybasket
        LastErNum = 0: LastErNum1 = 0
        LastErName = ""
        LastErNameGR = ""
        End If
        players(DisForm) = mybasket
        End If
        mybasket = players(DisForm)
        
        LCTbasketCur DIS, mybasket
         If mybasket.curpos > 0 Then
          crNew basestack1, mybasket
        
        
          
         End If
 mybasket.curpos = 0
MOUT = True
NoAction = False
If ExTarget Then Exit Do
para$ = ""
Loop
If NERR Then
MsgBoxN "ShutDown", vbCritical, "Abnormal Exit"
End If
NoAction = False
DelTemp
Set DisStack.Owner = Nothing
Set basestack1.Owner = Nothing

Set LastGlist = Nothing
Set LastGlist2 = Nothing
Unload Form5
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set MeStack.Owner = Nothing
TEXT1.Dereference
Set Point2Me = Nothing
Exit Sub ' why......................................because form5 be closing the door..
RemoveFont GetCurDir(True) & "TT6492M_.TTF"
MediaPlayer1.closeMovie
  DisableMidi
  TaskMaster.Dispose
  Set TaskMaster = Nothing

Dim x As Form
For Each x In Forms
'MsgBox x.name
If x.name <> Me.name Then Unload x
Next
If App.UnattendedApp Then End
End Sub

Private Sub List1_DblClick()
Dim DUMMY As Boolean
List1.Visible = False
If List1.Tag <> "" Then
If QRY Or GFQRY Then
Else

DUMMY = interpret(basestack1, List1.Tag)
Me.KeyPreview = True
End If
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
Dim DUMMY As Boolean
If KeyAscii = 13 Then
List1.Visible = False
If List1.Tag <> "" Then
If QRY Or GFQRY Then
Else

DUMMY = interpret(basestack1, List1.Tag)
Me.KeyPreview = True
End If
End If
End If
End Sub




Private Sub gList1_KeyDown(KeyCode As Integer, shift As Integer)
Static ctrl As Boolean, noentrance As Boolean, where As Long
Dim aa$, A$, jj As Long, ii As Long
If KeyCode = vbKeyEscape Then
KeyCode = 0
 If Not EditTextWord Then
 ' check if { } is ok...
 If Not blockCheck(TEXT1.Text, DialogLang) Then Exit Sub
 End If
 If TEXT1.UsedAsTextBox Then result = 99
NOEDIT = True: noentrance = False: Exit Sub
End If
If KeyCode = vbKeyPause Then
 KeyCode = 0: NOEDIT = True: noentrance = False
 If Form4.Visible Then Form4.Visible = False
             If TEXT1.Visible Then
                TEXT1.SetFocus
                Form1.SetFocus
            End If
            BreakMe = True
            If ASKINUSE Then
                If BreakMe Then Exit Sub
                Unload NeoMsgBox: ASKINUSE = False: Exit Sub
                End If
            If Form3.ask(basestack1, "Break Key - Hard Reset" & vbCrLf & "Μ2000 - Execution Stop / Τερματισμός Εκτέλεσης") = 1 Then
            
            If AVIRUN Then
            AVI.GETLOST
            End If


On Error Resume Next
noentrance = True
NoAction = True
Close ' we closed all files
QRY = False
escok = True
INK$ = Chr$(27) + Chr$(27)

If MOUT = False Then
NOEXECUTION = True
MOUT = True
Else
MOUT = False
End If
End If
BreakMe = False
 Exit Sub
 
End If
'***************************************
'Exit Sub
If TEXT1.UsedAsTextBox Then
Select Case KeyCode
Case Is = vbKeyTab And (shift Mod 2 = 1), vbKeyUp
result = -1
Case vbKeyReturn, vbKeyTab, vbKeyDown
result = 1
Case Else
Exit Sub
End Select
KeyCode = 0

NOEDIT = True: noentrance = False: Exit Sub

Exit Sub
End If

If noentrance Then
KeyCode = 0
Exit Sub
End If
noentrance = True
Form1mn1Enabled = TEXT1.SelLength > 1
Form1mn2Enabled = Form1mn1Enabled
Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
Select Case KeyCode
Case vbKeyReturn
nochange = True



If TEXT1.AutoIntNewLine Then

KeyCode = 0
Else

End If

'TEXT1.glistN.ShowMe2
nochange = False
'gList1_MarkOut


Case vbKeyControl
ctrl = True
KeyCode = 0
Case vbKeyF1
If (shift And 2) = 2 Then
If TEXT1.SelText <> "" Then
helpmeSub
Else

vHelp
End If
Else
TEXT1.nowrap = Not TEXT1.nowrap
TEXT1.Render
TEXT1.ManualInform
End If

KeyCode = 0
Case vbKeyF2
If shift <> 0 Then
If pagio$ = "GREEK" Then
s$ = InputBoxN("Αναζήτησε προς τα πάνω:", "Συγγραφή Κειμένου", s$)
Else
s$ = InputBoxN("Search to top:", "Text Editor", s$)
End If
If s$ <> "" Then Searchup s$, shift Mod 2 = 1
shift = 0
ElseIf TEXT1.SelText <> "" Or s$ <> "" Then
supsub
End If

KeyCode = 0
Case vbKeyF3
If shift <> 0 Then

If pagio$ = "GREEK" Then
s$ = InputBoxN("Αναζήτησε προς τα κάτω:", "Συγγραφή Κειμένου", s$)
Else
s$ = InputBoxN("Search to down:", "Text Editor", s$)
End If
If s$ <> "" Then SearchDown s$, shift Mod 2 = 1
shift = 0
ElseIf TEXT1.SelText <> "" Or s$ <> "" Then

sdnSub
End If
KeyCode = 0
Case vbKeyF4
If TEXT1.SelText <> "" Then mscatsub
KeyCode = 0
Case vbKeyF5
If TEXT1.SelText <> "" Then rthissub
KeyCode = 0
Case vbKeyF6  ' Set/Show/Reset Para1
MarkSoftButton Para1, PosPara1
KeyCode = 0
Case vbKeyF7  'Set/Show/Reset Para2
MarkSoftButton Para2, PosPara2
KeyCode = 0
Case vbKeyF8  'Set/Show/Reset Para2
MarkSoftButton Para3, PosPara3
KeyCode = 0

Case vbKeyF9  ' Count Words
If TEXT1.glistN.lines > 1 Then
If UserCodePage = 1253 Then
TEXT1.ReplaceTitle = "Λέξεις στο κείμενο:" + CStr(TEXT1.mdoc.WordCount)
Else
TEXT1.ReplaceTitle = "Words in text:" + CStr(TEXT1.mdoc.WordCount)
End If
End If
KeyCode = 0
Case vbKeyF10
TEXT1.showparagraph = Not TEXT1.showparagraph
TEXT1.mdoc.WrapAgainColor
TEXT1.Render
KeyCode = 0

Case vbKeyF11
fState = fState + 1
SetText1
TEXT1.WrapAll
TEXT1.mdoc.WrapAgainColor
TEXT1.ManualInform
KeyCode = 0
Case vbKeyF12
If shift <> 0 Then
mn5sub

Else
showmodules
End If
KeyCode = 0

Case vbKeyTab
nochange = True
gList1.Enabled = False
jj = TEXT1.SelStart
where = jj
ii = 1 + TEXT1.SelStart - TEXT1.ParaSelStart

If TEXT1.SelLength > 0 Then

jj = TEXT1.SelLength + jj - ii
TEXT1.SelStart = ii
TEXT1.SelLength = jj
jj = where
Else
TEXT1.SelStart = ii
End If


If TEXT1.SelText <> "" Then

    A$ = vbCrLf + TEXT1.SelText & "*"
    If shift <> 0 Then  ' βγάλε
        A$ = Replace(A$, vbCrLf + Space$(6), vbCrLf)
        TEXT1.InsertTextNoRender = Mid$(A$, 3, Len(A$) - 3)
         TEXT1.SelStartSilent = ii
         TEXT1.SelLengthSilent = Len(A$) - 3
         TEXT1.mdoc.WrapAgainColor
    Else
        A$ = Replace(A$, vbCrLf, vbCrLf + Space$(6))
        TEXT1.InsertTextNoRender = Mid$(A$, 3, Len(A$) - 3)
        TEXT1.SelStartSilent = where + 6
        TEXT1.SelLengthSilent = Len(A$) - 3 - (where + 6 - ii)
        TEXT1.mdoc.WrapAgainColor
    End If
  
Else
If shift <> 0 Then

    If Mid$(TEXT1.CurrentParagraph, 1, 6) = Space$(6) Then

            TEXT1.SelStartSilent = ii
            TEXT1.SelLengthSilent = 6
            TEXT1.InsertTextNoRender = ""
            TEXT1.SelStartSilent = ii
    End If
    Else
        TEXT1.SelStartSilent = jj
        TEXT1.RemoveUndo Space(6)
        TEXT1.InsertText = Space(6)
        
        TEXT1.SelStartSilent = where + 6
    End If
End If
gList1.Enabled = True
TEXT1.mdoc.WrapAgainColor
TEXT1.Render
nochange = False
'gList1_MarkOut
Case Else

ctrl = False
End Select
noentrance = False
End Sub














Private Sub TEXT1_Inform(tLine As Long, tPos As Long)
If TEXT1.UsedAsTextBox Then

Else
If UserCodePage = 1253 Then
textinformCaption = "Γραμμή(" + CStr(tLine) + ")-Θέση(" + CStr(TEXT1.Charpos) + ")"
TEXT1.ReplaceTitle = "[" + CStr(TEXT1.Charpos) + "-" + CStr(tLine) + "/" + CStr(TEXT1.mdoc.DocLines) + "]  §:" + CStr(TEXT1.mdoc.DocParagraphs) + Mark$ + " " + GetLCIDFromKeyboardLanguage

Else
textinformCaption = "Line(" + CStr(tLine) + ")-Pos(" + CStr(TEXT1.Charpos) + ")"
TEXT1.ReplaceTitle = "[" + CStr(TEXT1.Charpos) + "-" + CStr(tLine) + "/" + CStr(TEXT1.mdoc.DocLines) + "] §:" + CStr(TEXT1.mdoc.DocParagraphs) + Mark$ + " " + GetLCIDFromKeyboardLanguage
End If

End If
End Sub



Private Sub textinform_Click()
Dim s$, k As Long
''If EditTextWord Then

If pagio$ = "GREEK" Then
s$ = InputBoxN("Πήγαινε στη γραμμή (από 1):", "Συγγραφή Κειμένου", s$)
Else
s$ = InputBoxN("Goto line (from 1):", "Text Editor", s$)
End If
If IsNumberA(s$, k) Then
TEXT1.SetRowColumn k, 1
End If
''Else

''End If

End Sub

Private Sub view1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If look1 Then
look1 = False:  lookfirst = False

Cancel = True  ' 2 times
End If

If lookfirst Then look1 = True: view1.Silent = True


End Sub


Private Sub view1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
   Set HTML = view1.Document

End Sub

Private Sub view1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'
On Error Resume Next
If look1 Then
''Set pDisp = view1.Object
view1.SetFocus

End If
End Sub

Private Sub view1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Static prev$
Cancel = True
If HTML Is Nothing Then Exit Sub
If prev$ = HTML.activeElement.toString Then

Else
prev$ = HTML.activeElement.toString
view1.Navigate prev$
Sleep 50
End If
End Sub



Private Sub view1_TitleChange(ByVal Text As String)

If LCase$(Right$(Text, 5)) <> "done/" Then
If InStr(Text, "?") = 1 Then
nnn$ = TClear(Text): Sleep 5
'Beep
view1.Navigate "http://done/"
End If
Else

End If

End Sub
Function TClear(ByVal txt As String) As String
Dim Nb As String, ic As Long
txt = StrConv(txt, vbUnicode)
For ic = 1 To Len(txt) Step 2
Nb = Nb + Mid$(txt, ic, 1)
Next ic
TClear = Nb & "."
End Function
Private Sub view1_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub view1_LostFocus()
'Me.KeyPreview = True
End Sub




Public Sub view1_StatusTextChange11(bstack As basetask, ByVal t1 As String)
On Error Resume Next
exWnd = 0

view1.Visible = False
Sleep 1
PREPARE bstack, t1
Sleep 1
If Form1.Visible Then Form1.refresh
End Sub
Private Sub PREPARE(basestack As basetask, ByVal Nb As String)
Dim g As Long, ic As Long
Dim VP As String, vv As String, cm$, b As String
If needset Then
Nb = StrConv(Nb, vbUnicode)
For ic = 1 To Len(Nb) Step 2
b = b + Mid$(Nb, ic, 1)
Next ic
Else
b = Nb
b = Replace(b, Chr(0), "")
End If
g = InStr(b, "?")
If g > 0 Then b = Mid$(b, g + 1) Else Exit Sub
If b <> "" Then
Do While Parameters(b, VP, vv)
cm$ = VP & "$=" & """" + vv + """"
Execute basestack, cm$, True
Loop
MyDoEvents
''Me.KeyPreview = True
End If
End Sub
Public Sub IEUP(ThisFile As String)
Static Once As Boolean
If Once Then Exit Sub
Once = True
If ThisFile = "" Then

If exWnd <> 0 Then
 Set HTML = Nothing
MyDoEvents
homepage$ = ""
'view1.TabStop = False
    nnn$ = "bye bye"
    exWnd = 0
    
    End If
If view1.Visible Then
MyDoEvents

view1.Navigate "about:blank"
Sleep 50
view1.Visible = False

 End If
 Once = False
    Exit Sub
End If
   


On Error Resume Next
'tf2$ = THISFILE
needset = False
Dim MSD As String
MSD = App.path
AddDirSep MSD
'View1.TabStop = True

If Form1.Visible = False Then Form1.Visible = True: Sleep 400
view1.Visible = True

lookfirst = True
look1 = False



If IESizeX = 0 Or IESizeY < 100 Then
IEX = 0
IEY = 0
IESizeX = Me.ScaleWidth
IESizeY = Me.ScaleHeight
End If
If (IESizeX - IEX) > Me.ScaleWidth Then IESizeX = Me.ScaleWidth - IEX
If (IESizeY - IEY) > Me.ScaleHeight Then IESizeY = Me.ScaleHeight - IEY
view1.Move IEX, IEY, IESizeX, IESizeY
view1.RegisterAsBrowser = True
   If homepage$ = "" Then homepage$ = ThisFile$
   exWnd = 1
view1.Navigate ThisFile$

Do
view1.Visible = True
MyDoEvents2 Me
Sleep 5
Loop Until view1.Visible Or MOUT

If Not MOUT Then
view1.setfoucs
Me.KeyPreview = False
End If

'follow IEX, IEY
cnt = False
Once = False
End Sub
Public Sub follow(ByVal nx As Long, ByVal ny As Long)
Exit Sub
IEX = nx
IEY = ny
If exWnd > 0 Then

End If

End Sub


Private Function Parameters(A As String, b As String, c As String) As Boolean
Dim i, ch As Boolean, vl As Boolean, chs$, all$, many As Long
b = ""
c = ""

'parameters = False
ch = False
vl = False
Do While i < Len(A)
i = i + 1
Select Case Mid$(A, i, 1)
Case "%"
If Mid$(A, i + 1, 1) = "u" Then
i = i + 1
'we have four bytes
many = 6
Else
many = 4
End If
chs$ = "&H"
ch = True
Case ";"
If Not vl Then
' throw it is &amp;
b = ""
End If
Case "+"
If vl = True Then
c = c & " "
Else
b = b & " "
End If
Case "="
vl = True
Case "&", "#"
If b <> "" Then 'skip
vl = False
' here is the end
Exit Do
End If
Case Else
If ch = True Then
chs$ = chs$ & Mid$(A, i, 1)
If Len(chs$) = many Then
If many = 4 Then
chs$ = Chr(Int(chs$))
Else
chs$ = StrConv(Chr(CLng("&h" & Mid$(chs$, 5))) + Chr(CLng(Left$(chs$, 4))), vbFromUnicode)
End If
ch = False
If vl Then
c = c + chs$
Else
b = b + chs$
End If
End If
ElseIf vl = False Then
b = b + Mid$(A, i, 1)
Else
c = c + Mid$(A, i, 1)
End If
End Select
Loop
If c <> "" Then Parameters = True
A = Mid$(A, i + 1)
End Function


Public Sub myBreak(basestack As basetask)
''Dim pagio$

Dim cc As Object
Set cc = New cRegistry

cc.ClassKey = HKEY_CURRENT_USER
 cc.SectionKey = "Software\"
    cc.SectionKey = basickey

   cc.ValueKey = "FONT"
        cc.ValueType = REG_SZ
If Not cc.KeyExists Then
myBold = True
MYFONT = defFontname
            If Not Form1.FontName = MYFONT Then
            MYFONT = "Arial"
            Form1.FontName = MYFONT
            Form1.Font.Italic = False
            Form1.FontName = MYFONT
            
            End If
            MYFONT = Form1.FontName
            
            FFONT = MYFONT
            Err.clear
            DIS.FontName = MYFONT
            DIS.Font.Italic = False
            DIS.FontName = MYFONT
            If Err.Number > 0 Then
            Err.clear
            MYFONT = defFontname
            End If
            If TEXT1.Font.charset <> 161 Then
            
    Font.charset = basestack.myCharSet
    Font.bold = basestack.myBold
    
    DIS.Font.charset = basestack.myCharSet
    DIS.Font.bold = basestack.myBold
    TEXT1.Font.charset = basestack.myCharSet
    TEXT1.Font.bold = basestack.myBold

    List1.Font.charset = basestack.myCharSet
    List1.Font.bold = basestack.myBold

                pagio$ = "LATIN"
                DialogSetupLang 1
                Else
                DIS.Font.charset = basestack.myCharSet
                DialogSetupLang 0
                pagio$ = "GREEK"
            End If
                    SzOne = 14
             PenOne = 15
             PaperOne = 1
             DIS.ForeColor = mycolor(PenOne)
             On Error Resume Next
             cc.Value = Form1.FontName
                 cc.ValueKey = "LINESPACE"
        cc.ValueType = REG_DWORD
        If cc.Value >= 0 And cc.Value <= 120 * dv15 Then
     FeedBasket Form1.DIS, players(0), CLng(cc.Value) \ 2
    Else
   FeedBasket Form1.DIS, players(0), CLng(cc.Value) \ 2
    End If
               cc.ValueKey = "SIZE"
        cc.ValueType = REG_DWORD
             cc.Value = 14
              cc.ValueKey = "BOLD"
        cc.ValueType = REG_DWORD
             cc.Value = 1
              cc.ValueKey = "PEN"
        cc.ValueType = REG_DWORD
        cc.Value = 15
          cc.ValueKey = "PAPER"
        cc.ValueType = REG_DWORD
          cc.Value = 1
                 cc.ValueKey = "COMMAND"
        cc.ValueType = REG_SZ
        cc.Value = pagio$
                cc.ValueKey = "HTML"
        cc.ValueType = REG_SZ
        cc.Value = pagiohtml$
               cc.ValueKey = "CASESENSITIVE"
        cc.ValueType = REG_SZ
        If cc.Value = "" Then
        If casesensitive = True Then
         cc.Value = "YES"
        Else
    
        cc.Value = "NO"
        End If
        End If
Else
' *****************************
        If cc.Value = "" Then
        cc.Value = defFontname
        MYFONT = defFontname
        
        Else
        MYFONT = cc.Value
        On Error Resume Next
        
        Me.Font.name = MYFONT
        Me.Font.Italic = False
        Me.Font.name = MYFONT
        If Me.Font.name <> MYFONT Then
        MYFONT = defFontname
        End If
       
        End If
FFONT = MYFONT
Err.clear
DIS.FontName = MYFONT
DIS.Font.Italic = False
DIS.FontName = MYFONT
If Err.Number > 0 Then
Err.clear
MYFONT = defFontname
End If
    cc.ValueKey = "BOLD"
        cc.ValueType = REG_DWORD
        basestack.myBold = cc.Value <> 0
        Form1.Font.bold = basestack.myBold
    cc.ValueKey = "LINESPACE"
        cc.ValueType = REG_DWORD
        If cc.Value >= 0 And cc.Value <= 120 * dv15 Then
  FeedBasket Form1.DIS, players(0), CLng(cc.Value) \ 2
    Else
  FeedBasket Form1.DIS, players(0), 0
    End If

    cc.ValueKey = "SIZE"
        cc.ValueType = REG_DWORD
        If cc.Value = 0 Then
        cc.Value = 14
        SzOne = 14
        Else
        If cc.Value >= 8 And cc.Value <= 28 Then
        SzOne = cc.Value
        Else
        cc.Value = 14
        SzOne = 14
        End If
        End If
    cc.ValueKey = "PEN"
        cc.ValueType = REG_DWORD
        PenOne = cc.Value
    If Not (PenOne >= 0 And PenOne <= 15) Then PenOne = 15
        
    cc.ValueKey = "PAPER"
        cc.ValueType = REG_DWORD
      
    If cc.Value = PenOne Then cc.Value = 16 - PenOne
   
        DIS.ForeColor = mycolor(PenOne)
    cc.ValueKey = "PAPER"
        cc.ValueType = REG_DWORD
        PaperOne = cc.Value
        cc.ValueKey = "COMMAND"
        cc.ValueType = REG_SZ
        If cc.Value = "" Then
        cc.Value = "GREEK"
        End If
        pagio$ = cc.Value
        cc.ValueKey = "HTML"
        cc.ValueType = REG_SZ
        If cc.Value = "" Then
        cc.Value = "DARK"
        End If
         pagiohtml$ = cc.Value
        
        
        cc.ValueKey = "CASESENSITIVE"
        cc.ValueType = REG_SZ
       
        If cc.Value = "YES" Then
         casesensitive = True
        Else
    
       casesensitive = False
        End If
        
        Set cc = Nothing
        End If
       DIS.ForeColor = mycolor(PenOne) ' NOW PEN IS RGB VALUE
            Font.charset = basestack.myCharSet
    Font.bold = basestack.myBold
    DIS.Font.charset = basestack.myCharSet
    DIS.Font.bold = basestack.myBold
    TEXT1.Font.charset = basestack.myCharSet
    TEXT1.Font.bold = basestack.myBold
    List1.Font.charset = basestack.myCharSet
    List1.Font.bold = basestack.myBold

        Select Case pagio$
        Case "GREEK"
         GREEK basestack1
        Case Else   '"LATIN"
            LATIN basestack1
        
         End Select


End Sub




Public Sub mn1sub()
TEXT1.MarkCut
End Sub

Public Sub mn2sub()
TEXT1.MarkCopy
End Sub

Public Sub mn3sub()
On Error Resume Next
Dim aa$
aa$ = GetTextData(13)
If aa$ = "" Then aa$ = Clipboard.GetText(1)
With TEXT1
If .ParaSelStart = 2 And .glistN.List(.glistN.listindex) = "" Then
.SelStart = .SelStart - 1
End If
.AddUndo ""
.SelText = aa$
.RemoveUndo aa$
.mdoc.WrapAgainColor
End With
End Sub
Public Sub mn4sub()
 If Not EditTextWord Then
 ' check if { } is ok...
 If Not blockCheck(TEXT1.Text, DialogLang) Then Exit Sub
 End If

MyDoEvents
NOEDIT = True

End Sub


Private Sub wdragSub()
TEXT1.glistN.DragEnabled = Not TEXT1.glistN.DragEnabled
End Sub

Public Sub wordwrapsub()
TEXT1.nowrap = Not TEXT1.nowrap
TEXT1.Render
TEXT1.ManualInform
End Sub
Function GetKeY(ascii As Integer) As String
    Dim Buffer As String, Ret As Long
    Buffer = String$(514, 0)
    Dim r&, k&
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      r = CLng(Val("&H" & Right(Hex(r), 4)))
    Ret = GetLocaleInfo(r, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    If Ret > 0 Then
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, CLng(Val("&h" + Left$(Buffer, Ret - 1))))))
    Else
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, 1033)))
    End If
End Function
Public Function GetLCIDFromKeyboard() As Long
    Dim Buffer As String, Ret&, r&
    Buffer = String$(514, 0)
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      r = Val("&H" & Right(Hex(r), 4))
        Ret = GetLocaleInfo(r, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    GetLCIDFromKeyboard = CLng(Val("&h" + Left$(Buffer, Ret - 1)))
End Function
Sub MarkSoftButton(para As Long, pospara As Long)
If TEXT1.glistN.lines = 1 Then Exit Sub
If ShadowMarks Then Exit Sub
If para = 0 Then 'set
    para = TEXT1.mdoc.MarkParagraphID
    pospara = TEXT1.ParaSelStart
    
    If UserCodePage = 1253 Then
        TEXT1.ReplaceTitle = "Ο δείκτης τώρα θα δείχνει αυτή την παράγραφο"
    Else
        TEXT1.ReplaceTitle = "Mark now move to this Paragraph and Position"
    End If
ElseIf para = TEXT1.mdoc.MarkParagraphID And pospara = TEXT1.Charpos Then 'Reset
    para = 0
    
    If UserCodePage = 1253 Then
    TEXT1.ReplaceTitle = "Διαγραφή Δείκτη"
    Else
    TEXT1.ReplaceTitle = "Mark Deleted"
    End If
Else ' goto that paragraph
    If Not TEXT1.mdoc.InvalidPara(para) Then
        TEXT1.SelLengthSilent = 0
        TEXT1.mdoc.MarkParagraphID = para
        TEXT1.glistN.Enabled = False
        TEXT1.ParaSelStart = pospara
        TEXT1.glistN.Enabled = True
        TEXT1.ManualInform
    Else
        para = 0
        If UserCodePage = 1253 Then
            TEXT1.ReplaceTitle = "Δεν βρέθηκε παράγραφος - διαγράφτηκε ο δείκτης"
        Else
            TEXT1.ReplaceTitle = "Paragraph noτ found - marka deleted"
        End If
    End If
End If
End Sub

Function Mark$()
If ShadowMarks Then Mark$ = "": Exit Function
If TEXT1.title = "" Then  'reset all para
Para1 = 0: Para2 = 0: Para3 = 0
ElseIf LastDocTitle$ <> TEXT1.title Then
Para1 = 0: Para2 = 0: Para3 = 0
LastDocTitle$ = TEXT1.title
End If
Dim s$
If Para1 <> 0 Then
If TEXT1.mdoc.InvalidPara(Para1) Then Para1 = 0
If Para1 = TEXT1.mdoc.MarkParagraphID Then
s$ = " [F6] "
Else
s$ = " *F6 "
End If
Else
s$ = " -F6" + ChrW(&H25CA)
End If
If Para2 <> 0 Then
If TEXT1.mdoc.InvalidPara(Para2) Then Para2 = 0
If Para2 = TEXT1.mdoc.MarkParagraphID Then
s$ = s$ + " [F7] "
Else
s$ = s$ + " *F7"
End If
Else
s$ = s$ + " -F7" + ChrW(&H25CA)
End If
If Para3 <> 0 Then
If TEXT1.mdoc.InvalidPara(Para3) Then Para3 = 0
If Para3 = TEXT1.mdoc.MarkParagraphID Then
s$ = s$ + " [F8] "
Else
s$ = s$ + " *F8 "
End If
Else
s$ = s$ + " -F8" + ChrW(&H25CA)
End If
Mark$ = s$

End Function
Public Sub ResetMarks()
Para1 = 0: Para2 = 0: Para3 = 0


End Sub
Public Sub hookme(this As gList)
Set LastGlist = this
End Sub
Public Function mybreak1() As Boolean
Dim i As Long
If Form4.Visible Then Form4.Visible = False
i = MOUT
If ASKINUSE Then
If BreakMe Then Exit Function
Unload NeoMsgBox: ASKINUSE = False: Exit Function
End If
BreakMe = True

INK$ = ""
If MsgBoxN("Break Key - Hard Reset" + vbCrLf + "Μ2000 - Execution Stop / Τερματισμός Εκτέλεσης", vbYesNo, MesTitle$) <> vbNo Then
                
                
                If AVIRUN Then AVI.GETLOST
                On Error Resume Next
                LastErName = ""
                LastErNum = 0
                LastErNum1 = 0
                If Me.Visible Then Me.SetFocus
                closeAll ' we closed all files
                QRY = False
                GFQRY = False
                escok = True
                INK$ = Chr$(27) + Chr$(27)
                If List1.Visible Then
                                List1.Tag = ""
                                List1.Visible = False
                                List1.LeaveonChoose = False
                                INK$ = ""
                End If
                mybreak1 = True
End If
 
 
BreakMe = False
End Function
Public Sub SetText1()
If (600 - hueconv(TEXT1.BackColor)) Mod 360 > 30 And lightconv(TEXT1.BackColor) >= 128 Then TEXT1.ColorSet = 1 Else TEXT1.ColorSet = 0
Select Case fState
Case 0
shortlang = False
TEXT1.NoColor = EditTextWord
Case 1
shortlang = False
TEXT1.NoColor = True
Case 2
shortlang = True
TEXT1.NoColor = EditTextWord
Case 3
shortlang = True
TEXT1.NoColor = True
fState = -1
End Select
End Sub

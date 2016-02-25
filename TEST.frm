VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7860
   Icon            =   "TEST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin M2000.gList gList4 
      Height          =   1920
      Left            =   4050
      TabIndex        =   5
      Top             =   810
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   3387
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
      BorderStyle     =   1
      ShowBar         =   0   'False
      Backcolor       =   3881787
      ForeColor       =   16777215
   End
   Begin M2000.gList gList3 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   1058
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   3881787
      ForeColor       =   16777215
   End
   Begin M2000.gList gList0 
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   4620
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   979
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      Backcolor       =   657930
      ForeColor       =   16777215
   End
   Begin M2000.gList gList1 
      Height          =   1800
      Left            =   90
      TabIndex        =   0
      Top             =   2775
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   3175
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      Backcolor       =   3881787
      ForeColor       =   16777215
   End
   Begin M2000.gList gList3 
      Height          =   615
      Index           =   1
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1455
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   1085
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      Backcolor       =   3881787
      ForeColor       =   16777215
   End
   Begin M2000.gList gList3 
      Height          =   615
      Index           =   2
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2115
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   1085
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      Backcolor       =   3881787
      ForeColor       =   16777215
   End
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   135
      Width           =   7635
      _ExtentX        =   13467
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
      Enabled         =   -1  'True
      Backcolor       =   3881787
      ForeColor       =   16777215
      CapColor        =   16777215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents testpad As TextViewer
Attribute testpad.VB_VarHelpID = -1
Public WithEvents compute As myTextBox
Attribute compute.VB_VarHelpID = -1
Private Label(0 To 2) As New myTextBox

Dim MyBaseTask As New basetask
Dim setupxy As Single
Dim Lx As Long, ly As Long, dr As Boolean, drmove As Boolean
Dim prevx As Long, prevy As Long
Dim A$
Dim bordertop As Long, borderleft As Long
Dim allheight As Long, allwidth As Long, itemWidth As Long, itemwidth3 As Long, itemwidth2 As Long
Dim height1 As Long, width1 As Long
Dim doubleclick As Long

Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Dim EXECUTED As Boolean
Public Property Set Process(mBtask As basetask)
Set MyBaseTask = mBtask
End Property
Private Sub Command1_Click()
trace = True
STq = False
STbyST = True
End Sub
Private Sub Command2_Click()
trace = True
STq = True
STbyST = False
End Sub

Private Sub Command3_Click()
NOEXECUTION = True
trace = True
STEXIT = True
End Sub
Public Sub ComputeNow()
stackshow MyBaseTask
End Sub

Private Sub compute_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 13 Then
KeyCode = 0
gList3(2).BackColor = &H3B3B3B
TestShowCode = False
stackshow MyBaseTask
End If
End Sub
'M2000 [ΕΛΕΓΧΟΣ - CONTROL]

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 27 Then
KeyCode = 0
Unload Me
ElseIf KeyCode = 13 Then
If Not EXECUTED Then
If Not STq Then
STbyST = True
End If
End If
End If
End Sub

Private Sub Form_Load()
Dim i As Long
height1 = 5280 * DYP / 15
width1 = 7860 * DXP / 15
lastfactor = 1
LastWidth = -1
HelpLastWidth = -1
PopUpLastWidth = -1
setupxy = 20
lastfactor = ScaleDialogFix(SizeDialog)
ScaleDialog lastfactor, LastWidth
gList4.NoCaretShow = True
gList4.restrictLines = 3
gList4.CenterText = True
gList2.CapColor = rgb(255, 160, 0)
gList2.HeadLine = ""

gList2.FloatList = True
gList2.FloatLimitTop = ScrY() - players(0).Yt * 2
gList2.FloatLimitLeft = ScrX() - players(0).Xt * 2
gList2.MoveParent = True
gList2.Enabled = True
gList1.DragEnabled = False
gList1.AutoPanPos = True
Set testpad = New TextViewer
gList1.NoWheel = True
Set testpad.Container = gList1
testpad.FileName = ""
testpad.glistN.LeftMarginPixels = 8
testpad.NoMark = True
testpad.NoColor = False
testpad.EditDoc = False
testpad.nowrap = False
testpad.Enabled = True
Set compute = New myTextBox
Set compute.Container = gList0
compute.MaxCharLength = 500 ' as a limit
compute.Locked = False
compute.Enabled = True
compute.Retired
Set Label(0).Container = gList3(0)
Set Label(1).Container = gList3(1)
Set Label(2).Container = gList3(2)
If pagio$ = "GREEK" Then
gList2.HeadLine = "Έλεγχος"
compute.Prompt = "Τυπωσε "
Label(0).Prompt = "Τμήμα: "
Label(1).Prompt = "Εντολή: "
Label(2).Prompt = "Επόμενο: "
' Επόμενο Βήμα/ Next Step
' Αργή Ροή / Slow Flow
' Κράτηση / Stop

gList4.additemFast "Επόμενο Βήμα"
gList4.additemFast "Αργή Ροή"
gList4.additemFast "Διακοπή"
Else
gList2.HeadLine = "Control"
compute.Prompt = "Print "
Label(0).Prompt = "Module: "
Label(1).Prompt = "Id: "
Label(2).Prompt = "Next: "
gList4.additemFast "Next Step"
gList4.additemFast "Slow Flow"
gList4.additemFast "Stop"
End If
gList2.HeadlineHeight = gList2.HeightPixels
gList2.PrepareToShow
gList4.NoPanRight = False
gList4.SingleLineSlide = True
gList4.VerticalCenterText = True
gList4.Enabled = True
gList4.ListindexPrivateUse = 0
gList4.ShowMe
End Sub




Private Sub Form_Unload(Cancel As Integer)
testpad.Dereference
compute.Dereference
Set MyBaseTask = Nothing
trace = False
STq = True
End Sub



Private Sub gList1_CheckGotFocus()
gList1.BackColor = &H606060
gList1.ShowMe2
End Sub

Private Sub gList1_CheckLostFocus()

gList1.BackColor = &H3B3B3B
gList1.ShowMe2
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


Public Property Get Label1(ByVal Index As Long) As String
Label1 = Label(Index)
End Property

Public Property Let Label1(ByVal Index As Long, ByVal rhs As String)
Label(Index) = rhs
End Property
Public Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT, b As Long
b = 2
CopyFromLParamToRect A, thatRect
A.Left = b
A.Right = setupxy - b
A.top = b
A.Bottom = setupxy - b
FillThere thathDC, VarPtr(A), 0
b = 5
A.Left = b
A.Right = setupxy - b
A.top = b
A.Bottom = setupxy - b
FillThere thathDC, VarPtr(A), rgb(255, 160, 0)


End Sub

Private Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT
CopyFromLParamToRect A, thatRect
FillBack thathDC, A, thatbgcolor
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub

Private Sub gList2_LostFocus()
doubleclick = 0
End Sub

Private Sub glist3_CheckGotFocus(Index As Integer)
Dim s$
gList4.SetFocus
If Index < 2 Then
abt = False

vH_title$ = ""
s$ = Label(Index)
Select Case Left$(LTrim(Label(2)) + " ", 1)
Case "?", "!", " ", ".", ":", Is >= "A", Chr$(10), """"
    fHelp MyBaseTask, s$, AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128
End Select
ElseIf Index = 2 Then
TestShowCode = Not TestShowCode
If TestShowCode Then
gList3(2).BackColor = &H606060
Label(2) = Label(2)
Else
gList3(2).BackColor = &H3B3B3B
Label(2) = Label(2)
End If
stackshow MyBaseTask
End If
End Sub

Private Sub gList4_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim A As RECT, b As RECT
CopyFromLParamToRect A, thisrect
CopyFromLParamToRect b, thisrect
A.Left = A.Left + 1 * lastfactor
A.Right = gList4.WidthPixels
b.Right = gList4.WidthPixels
 If item = gList4.listindex Then
   If EXECUTED Then
   FillBack thisHDC, b, &H77FF77
   Else
   
             FillBack thisHDC, b, &H77FFFF
             End If
             'EXECUTED = False
              SetTextColor thisHDC, 0
              b.top = b.Bottom - 1 * lastfactor
       
            FillBack thisHDC, b, &H777777
           
    Else
          
          
    SetTextColor thisHDC, gList4.ForeColor
    b.top = b.Bottom - 1
    FillBack thisHDC, b, 0
    End If
    If item = gList4.listindex Then
  A.Left = A.Left + 1 * lastfactor + gList4.PanPosPixels
  gList4.ForeColor = rgb(128, 0, 128)
  End If
   
   
   PrintItem thisHDC, gList4.List(item), A
    skip = True
End Sub
 
Private Sub gList4_PanLeftRight(Direction As Boolean)
EXECUTED = True
Action
End Sub

Private Sub gList4_selected(item As Long)
EXECUTED = False
gList4.ShowMe

End Sub

Private Sub gList4_Selected2(item As Long)
EXECUTED = True
Action
End Sub
Private Sub Action()
EXECUTED = True ' SO CHANGE THE BACKGROUND COLOR COLOR
Select Case gList4.listindex
Case 0
trace = True
STq = False
STbyST = True
Case 1
trace = True
STq = True
STbyST = False
Case 2
NOEXECUTION = True
trace = True
STq = False
STbyST = True
End Select
gList4.PanPos = 0
gList4.ShowMe2
End Sub

Private Sub gList4_softSelected(item As Long)
EXECUTED = Not EXECUTED
gList4.ShowMe


End Sub
 Private Sub PrintItem(mHdc As Long, c As String, r As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_CENTER Or DT_VCENTER)
    DrawText mHdc, StrPtr(c), -1, r, way
    End Sub
Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)

If Button = 1 Then
    
    If lastfactor = 0 Then lastfactor = 1

    If bordertop < 150 Then
    If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
    dr = True
    mousepointer = vbSizeNWSE
    Lx = x
    ly = y
    End If
    
    Else
    If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then
    dr = True
    mousepointer = vbSizeNWSE
    Lx = x
    ly = y
    End If
    End If

End If
End Sub
Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim addX As Long, addy As Long, factor As Single, Once As Boolean
If Once Then Exit Sub
If Button = 0 Then dr = False: drmove = False
If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then mousepointer = vbSizeNWSE Else If Not (dr Or drmove) Then mousepointer = 0
 Else
 If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then mousepointer = vbSizeNWSE Else If Not (dr Or drmove) Then mousepointer = 0
End If
If dr Then



If bordertop < 150 Then

        If y < (Height - 150) Or y > Height Then addy = (y - ly)
     If x < (Width - 150) Or x > Width Then addX = (x - Lx)
     
Else
    If y < (Height - bordertop) Or y > Height Then addy = (y - ly)
        If x < (Width - borderleft) Or x > Width Then addX = (x - Lx)
    End If
    

    
  If Not ExpandWidth Then addX = 0
        If lastfactor = 0 Then lastfactor = 1
        factor = lastfactor

        
  
        Once = True
        If Height > ScrY() Then addy = -(Height - ScrY()) + addy
        If Width > ScrX() Then addX = -(Width - ScrX()) + addX
        If (addy + Height) / height1 > 0.4 And ((Width + addX) / width1) > 0.4 Then
   
        If addy <> 0 Then SizeDialog = ((addy + Height) / height1)
        lastfactor = ScaleDialogFix(SizeDialog)


        If ((Width * lastfactor / factor + addX) / Height * lastfactor / factor) < (width1 / height1) Then
        addX = -Width * lastfactor / factor - 1
      
           End If

        If addX = 0 Then
        If lastfactor <> factor Then ScaleDialog lastfactor, Width
        Lx = x
        
        Else
        Lx = x * lastfactor / factor
         ScaleDialog lastfactor, (Width + addX) * lastfactor / factor
         End If

        
         
        
        LastWidth = Width
      gList2.HeadlineHeight = gList2.HeightPixels
        gList2.PrepareToShow
        gList1.PrepareToShow
        'testpad.Render
        ly = ly * lastfactor / factor
        End If
        Else
        Lx = x
        ly = y
   
End If
Once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)

If dr Then Me.mousepointer = 0
dr = False
End Sub
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)

lastfactor = factor
setupxy = 20 * factor
bordertop = 10 * dv15 * factor

borderleft = bordertop
allwidth = width1 * factor
allheight = height1 * factor
itemWidth = allwidth - 2 * borderleft
itemwidth3 = (itemWidth - 2 * borderleft) / 3
itemwidth2 = (itemWidth - borderleft) / 2
Move Left, top, allwidth, allheight
FontTransparent = False  ' clear background  or false to write over
gList2.Move borderleft, bordertop, itemWidth, bordertop * 3
gList2.FloatLimitTop = ScrY() - bordertop - bordertop * 3
gList2.FloatLimitLeft = ScrX() - borderleft * 3
gList3(0).Move borderleft, bordertop * 5, itemwidth2, bordertop * 4
gList3(1).Move borderleft, bordertop * 9, itemwidth2, bordertop * 4
gList3(2).Move borderleft, bordertop * 13, itemwidth2, bordertop * 4
gList4.Move borderleft * 2 + itemwidth2, bordertop * 5, itemwidth2, bordertop * 12
gList1.Move borderleft, bordertop * 18, itemWidth, bordertop * 12
gList0.Move borderleft, bordertop * 31, itemWidth, bordertop * 3
End Sub
Function ScaleDialogFix(ByVal factor As Single) As Single
gList2.FontSize = 14.25 * factor
factor = gList2.FontSize / 14.25
gList1.FontSize = 11.25 * factor
gList4.FontSize = 12 * factor
factor = gList1.FontSize / 11.25
gList3(0).FontSize = gList1.FontSize
gList3(1).FontSize = gList1.FontSize
gList3(2).FontSize = gList1.FontSize
gList0.FontSize = gList1.FontSize
ScaleDialogFix = factor
End Function

Public Sub hookme(this As gList)
If Not this Is Nothing Then this.NoWheel = True
End Sub



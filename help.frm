VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   4650
   ClientLeft      =   11925
   ClientTop       =   -6825
   ClientWidth     =   7080
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "help.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MousePointer    =   5  'Size
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin M2000.gList glist1 
      Height          =   3825
      Left            =   330
      TabIndex        =   0
      Top             =   300
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6747
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
      Backcolor       =   -2147483624
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents Label1 As TextViewer
Attribute Label1.VB_VarHelpID = -1
Private l As Long
Private t As Long
Private mt As Integer
Private back$
Private jump As Boolean

 Dim setupxy As Single
 Dim scrTwips As Long

Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Dim Lx As Long, ly As Long, dr As Boolean
Dim bordertop As Long, borderleft As Long
Dim allheight As Long, allwidth As Long, itemWidth As Long


Private Sub Form_Deactivate()
jump = False
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyF12 Then
showmodules
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo there1

If Form1.Visible Then
If Not gList1.EditFlag Then
Form1.SetFocus
INK$ = StrConv(ChrW$(KeyAscii Mod 256), 64, Form1.GetLCIDFromKeyboard)
End If
End If

there1:
End Sub

Private Sub Form_Load()
Set LastGlist2 = Nothing
setupxy = 20 * Helplastfactor
scrTwips = Screen.TwipsPerPixelX
gList1.CapColor = rgb(255, 160, 0)
gList1.LeftMarginPixels = 4
Set Label1 = New TextViewer
Set Label1.Container = gList1
Label1.FileName = ""
Label1.glistN.DropEnabled = False
Label1.glistN.DragEnabled = Not abt
Label1.NoMark = True
Label1.NoColor = True
Label1.EditDoc = False
Label1.nowrap = False
Label1.Enabled = False    '' true before
Label1.glistN.FloatList = True
Label1.glistN.MoveParent = True
With Label1.glistN
If Not abt Then
.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^")
.WordCharRight = ConCat(":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^")
.WordCharRightButIncluded = "(" ' so aaa(sdd) give aaa( as word
Else
.WordCharLeft = "["
.WordCharRight = "]"
.WordCharRightButIncluded = ""
End If
End With
mt = DXP
''Set HelpStack.Owner = Me
''SetTrans Me, 200, &HFFFFFF
If Helplastfactor = 0 Then Helplastfactor = 1
 Helplastfactor = ScaleDialogFix(helpSizeDialog)
If ExpandWidth And False Then
If HelpLastWidth = 0 Then HelpLastWidth = -1
Else
HelpLastWidth = -1
End If
If ExpandWidth Then
If HelpLastWidth = 0 Then HelpLastWidth = -1
Else
HelpLastWidth = -1
End If
''Me.FontSize = Int((ScrY() - 1) / DYP / 70 + 0.5)
''Label1.FontSize = Me.FontSize
''setupxy = Me.FontSize * 20 / 15 * DYP / 15 + 4

End Sub
Public Sub MoveMe()
ScaleDialog Helplastfactor, HelpLastWidth
Hook2 hwnd, gList1
Label1.glistN.SoftEnterFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)

If Button = 1 Then
    
    If Helplastfactor = 0 Then Helplastfactor = 1

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
    

    
  '' If Not ExpandWidth Then addX = 0
        If Helplastfactor = 0 Then Helplastfactor = 1
        factor = Helplastfactor

        
  
        Once = True
        If Height > ScrY() Then addy = -(Height - ScrY()) + addy
        If Width > ScrX() Then addX = -(Width - ScrX()) + addX
        If (addy + Height) / vH_y > 0.4 And ((Width + addX) / vH_x) > 0.4 Then
   
        If addy <> 0 Then helpSizeDialog = ((addy + Height) / vH_y)
        Helplastfactor = ScaleDialogFix(helpSizeDialog)


        If ((Width * Helplastfactor / factor + addX) / Height * Helplastfactor / factor) < (vH_x / vH_y) Then
        addX = -Width * Helplastfactor / factor - 1
      
           End If

        If addX = 0 Then
        
        If Helplastfactor <> factor Then ScaleDialog Helplastfactor, Width

        Lx = x
        
        Else
        Lx = x * Helplastfactor / factor
             ScaleDialog Helplastfactor, (Width + addX) * Helplastfactor / factor
         
   
         End If

        
        HelpLastWidth = Width


''gList1.PrepareToShow
        ly = ly * Helplastfactor / factor
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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UnHook2 hwnd
Set LastGlist2 = Nothing
End Sub

Private Sub Form_Terminate()
''''Set HelpStack.Owner = Nothing
End Sub

Private Sub ffhelp(A$)
If A$ = "цемийа" Then A$ = "ока"
If A$ = "GENERAL" Then A$ = "ALL"
If Left$(A$, 1) < "а" Then
fHelp basestack1, A$, True
Else
fHelp basestack1, A$
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Label1.Dereference  ' to ensure that no reference hold objects..
Set Label1 = Nothing
End Sub

Private Sub gList1_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If item = -1 Then
If gList1.DoubleClickCheck(Button, item, x, y, gList1.WidthPixels - 10 * Helplastfactor, 10 * Helplastfactor, 8 * Helplastfactor, -1) Then
HelpLastWidth = -1
            Unload Me
End If
Else
gList1.mousepointer = 1
End If
End Sub


Private Sub gList1_selected2(item As Long)
Label1.NoMark = False
Label1.EditDoc = True
End Sub

Private Sub glist1_WordMarked(ThisWord As String)
If abt Then
feedback$ = ThisWord
feednow$ = FeedbackExec$
CallGlobal feednow$
Else

ffhelp ThisWord

End If
ThisWord = ""

End Sub
Public Sub FillThereMyVersion2(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT, b As Long
b = setupxy / 3

CopyFromLParamToRect A, thatRect
A.Right = A.Right - b
A.Left = A.Right - setupxy - b
A.top = b
A.Bottom = b + setupxy / 5
FillThere thathDC, VarPtr(A), thatbgcolor
A.top = b + setupxy / 5 + setupxy / 10
A.Bottom = b + setupxy \ 2
FillThere thathDC, VarPtr(A), thatbgcolor
A.top = b + 2 * (setupxy / 5 + setupxy / 10)
A.Bottom = A.top + setupxy / 5
FillThere thathDC, VarPtr(A), thatbgcolor

End Sub
Public Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT, b As Long
b = 2
CopyFromLParamToRect A, thatRect
A.Left = A.Right - b
A.Right = A.Right - setupxy + b
A.top = b
A.Bottom = setupxy - b
FillThere thathDC, VarPtr(A), gList1.dcolor
b = 5
A.Left = A.Left - 3
A.Right = A.Right + 3
A.top = b
A.Bottom = setupxy - b
FillThere thathDC, VarPtr(A), gList1.CapColor


End Sub
Public Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
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

Private Sub Label1_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = -1 Then

FillThereMyVersion thisHDC, thisrect, &HF0F0F0
''skip = True
End If
End Sub
Function ScaleDialogFix(ByVal factor As Single) As Single
gList1.FontSize = 14.25 * factor
factor = gList1.FontSize / 14.25
ScaleDialogFix = factor
End Function
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)
Dim h As Long, i As Long
Helplastfactor = factor
setupxy = 20 * factor
bordertop = 10 * dv15 * factor
borderleft = bordertop
If (NewWidth < 0) Or NewWidth <= vH_x * factor Then
NewWidth = vH_x * factor
End If
allwidth = NewWidth  ''vH_x * factor
allheight = vH_y * factor
itemWidth = allwidth - 2 * borderleft
MyForm Me, Left, top, allwidth, allheight, True, factor

  
gList1.addpixels = 4 * factor
Label1.Move borderleft, bordertop, itemWidth, allheight - bordertop * 2

Label1.NewTitle vH_title$, (4 + UAddPixelsTop) * actor
Label1.Render
gList1.FloatLimitTop = ScrY() - bordertop - bordertop * 3
gList1.FloatLimitLeft = ScrX() - borderleft * 3


End Sub
Public Sub hookme(this As gList)

''Set LastGlist2 = this

End Sub

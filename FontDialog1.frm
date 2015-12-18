VERSION 5.00
Begin VB.Form FontDialog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   8145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FontDialog1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin M2000.gList gList1 
      Height          =   3375
      Left            =   135
      TabIndex        =   0
      Top             =   645
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   5953
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
      dcolor          =   65535
      Backcolor       =   3881787
      ForeColor       =   14737632
      CapColor        =   9797738
   End
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   3420
      _ExtentX        =   6033
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
   Begin M2000.gList glist3 
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   7635
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   661
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
      Backcolor       =   8421504
      ForeColor       =   14737632
      CapColor        =   49344
   End
   Begin M2000.gList gList4 
      Height          =   3060
      Left            =   75
      TabIndex        =   2
      Top             =   4350
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   5398
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
      dcolor          =   65535
      Backcolor       =   3881787
      ForeColor       =   14737632
      CapColor        =   16777215
   End
End
Attribute VB_Name = "FontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Public TEXT1 As myTextBox
Attribute TEXT1.VB_VarHelpID = -1
Dim firstpath As Long
Dim setupxy As Single
Dim Lx As Long, ly As Long, dr As Boolean
Dim scrTwips As Long
Dim bordertop As Long, borderleft As Long
Dim allwidth As Long, itemWidth As Long
Private Sub Form_Load()
loadfileiamloaded = True
scrTwips = Screen.TwipsPerPixelX
' clear data...
setupxy = 20
gList1.Enabled = True
gList2.Enabled = True
gList3.Enabled = True
gList3.LeaveonChoose = True
gList3.VerticalCenterText = True
gList3.restrictLines = 1
gList3.PanPos = 0
gList2.CapColor = rgb(255, 160, 0)
gList2.HeadLine = ""
gList2.FloatList = True
gList2.MoveParent = True
gList3.NoPanRight = False
gList1.NoCaretShow = True
gList1.NoPanLeft = False
gList1.StickBar = True
gList1.ShowBar = True
gList1.VerticalCenterText = True
With gList4
If DialogLang <> 0 Then
.additemFast "Font Type"
.menuEnabled(0) = False
.additemFast "  Bold"
.additemFast "  Italic"
.MenuItem 2, True, False, ReturnBold, "bold"
.MenuItem 3, True, False, ReturnItalic, "italic"
.additemFast "Font Size"
.menuEnabled(3) = False
.additemFast "  12"
.additemFast "Font Charset Number"
.menuEnabled(5) = False
.additemFast "  0"
.additemFast "Font Charset Table"
.menuEnabled(7) = False
.additemFast "  ANSI - 0"
.additemFast "  Default - 1"
.additemFast "  Greek - 161"
.additemFast "  Turkish - 162"
.additemFast "  Hebrew - 177"
.additemFast "  Arabic - 178"
.additemFast "  East Europe - 238"
.additemFast "  Russian - 204"
.additemFast "  Baltic -186"
.additemFast "Font Size Table"
Else
.additemFast "Τύπος Γραμματοσειράς"
.menuEnabled(0) = False
.additemFast "  Έντονη"
.additemFast "  Πλάγια"
.MenuItem 2, True, False, ReturnBold, "bold"
.MenuItem 3, True, False, ReturnItalic, "italic"
.additemFast "Μέγεθος Γραμματοσειράς"
.menuEnabled(3) = False
.additemFast "  12"
.additemFast "Αριθμός Κωδικοσελίδας"
.menuEnabled(5) = False
.additemFast "  0"
.additemFast "Πίνακας Κωδικοσελίδων"
.menuEnabled(7) = False
.additemFast "  ANSI - 0"
.additemFast "  Default - 1"
.additemFast "  Greek - 161"
.additemFast "  Turkish - 162"
.additemFast "  Hebrew - 177"
.additemFast "  Arabic - 178"
.additemFast "  East Europe - 238"
.additemFast "  Russian - 204"
.additemFast "  Baltic -186"
.additemFast "Πίνακας Μεγεθών"
End If
.menuEnabled(17) = False
.additemFast "  8"
.additemFast "  9"
.additemFast "  10"
.additemFast "  11"
.additemFast "  12"
.additemFast "  14"
.additemFast "  16"
.additemFast "  18"
.additemFast "  20"
.additemFast "  22"
.additemFast "  24"
.additemFast "  26"
.additemFast "  28"
.additemFast "  36"
.additemFast "  48"
.additemFast "  72"
.ShowMe
.ShowBar = True
.StickBar = True
.ShowBar = False
.NoCaretShow = True
.ListindexPrivateUse = 1
End With
  
Set TEXT1 = New myTextBox
Set TEXT1.Container = gList3
 lastfactor = ScaleDialogFix(SizeDialog)
If ExpandWidth Then
If LastWidth = 0 Then LastWidth = -1
Else
LastWidth = -1
End If
If ExpandWidth Then
If LastWidth = 0 Then LastWidth = -1
Else
LastWidth = -1
End If
ScaleDialog lastfactor, LastWidth
gList2.HeadLine = FontSelector
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
If selectorLastX = -1 And selectorLastY = -1 Then

Else
Move selectorLastX, selectorLastY
End If
Dim i As Integer
For i = 0 To Screen.FontCount - 1
If Screen.Fonts(i) = ReturnFontName Then
TEXT1 = Screen.Fonts(i)
TEXT1.Locked = False
gList3.Font.charset = ReturnCharset
gList3.Font.bold = ReturnBold
gList3.Font.Italic = ReturnItalic
gList1.ListindexPrivateUse = i
Exit For
End If
Next i
If ReturnSize >= 6 Then gList4.List(4) = "  " & CStr(ReturnSize)
gList4.List(6) = "  " & CStr(ReturnCharset)
gList4.Enabled = True
gList2.TabStop = False
gList1.ShowMe
TEXT1.Locked = False
gList3.listindex = 0
gList3.SoftEnterFocus
End Sub



Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)

If Button = 1 Then

If lastfactor = 0 Then lastfactor = 1

If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
dr = True

Lx = x
ly = y
End If

Else
If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then
dr = True
Lx = x
ly = y
End If

End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim addX As Long, addy As Long, factor As Single, Once As Boolean
If Once Then Exit Sub
If Button = 0 Then dr = False
If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then mousepointer = vbSizeNWSE Else mousepointer = 0
 Else
 If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then mousepointer = vbSizeNWSE Else mousepointer = 0
End If
If dr Then
    If y < (Height - bordertop) Or y > Height Then addy = (y - ly)
    If x < (Width - borderleft) Or x > Width Then addX = (x - Lx)
   If Not ExpandWidth Then addX = 0
        If lastfactor = 0 Then lastfactor = 1
        factor = lastfactor

        
  
        Once = True
        If Height > ScrY() Then addy = -(Height - ScrY()) + addy
        If Width > ScrX() Then addX = -(Width - ScrX()) + addX
        If (addy + Height) / (8145 * DYP / 15) > 0.4 And ((Width + addX) / (3690 * DXP / 15)) > 0.4 Then
   
        If addy <> 0 Then SizeDialog = ((addy + Height) / (8145 * DYP / 15))
        lastfactor = ScaleDialogFix(SizeDialog)


        If ((Width * lastfactor / factor + addX) / Height * lastfactor / factor) < (3690 / 8145 * DXP / DYP) Then
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
        gList2.SoftEnterFocus
      
      
        ly = ly * lastfactor / factor
    
        'End If
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

Private Sub Form_Unload(Cancel As Integer)
DestroyCaret
selectorLastX = Left
selectorLastY = top
Sleep 200
loadfileiamloaded = False
End Sub






Private Sub gList1_ExposeListcount(cListCount As Long)
cListCount = Screen.FontCount
End Sub

Private Sub gList1_GotFocus()
If gList1.listindex = -1 Then gList1.listindex = gList1.ScrollFrom
End Sub


Private Sub gList1_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyEscape Then
 CancelDialog = True
Unload Me
End If
End Sub



Private Sub gList1_ScrollSelected(item As Long, y As Long)
gList3.Font.name = Screen.Fonts(item - 1)
gList3.Font.Italic = gList4.ListSelected(2)
gList3.FontSize = 11.25 * lastfactor
gList3.FontBold = gList4.ListSelected(1)
gList3.Font.charset = Trim$(gList4.List(6))
TEXT1 = Screen.Fonts(item - 1)
End Sub

Private Sub gList1_selected(item As Long)
gList3.Font.name = Screen.Fonts(item - 1)
gList3.Font.Italic = gList4.ListSelected(2)
gList3.FontSize = 11.25 * lastfactor
gList3.FontBold = gList4.ListSelected(1)
gList3.Font.charset = Trim$(gList4.List(6))
TEXT1 = Screen.Fonts(item - 1)
End Sub

Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If gList2.DoubleClickCheck(Button, item, x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor, -1) Then
                      Unload Me
End If

End Sub


Private Sub gList1_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim A As RECT, b As RECT
Dim oldforecolor As Long
oldforecolor = gList1.ForeColor
If item = -1 Then
'FillThere thisHDC, thisrect, gList1.CapColor
'FillThereMyVersion2 thisHDC, thisrect, &HF0F0F0
'skip = True
Else
skip = True
CopyFromLParamToRect A, thisrect
CopyFromLParamToRect b, thisrect
A.top = A.top + 2
If gList1.listindex = item Then
b.Left = 0
FillBack thisHDC, b, 0
gList1.ForeColor = &HFFFFFF
Else
FillBack thisHDC, b, gList1.BackColor

End If

PrintItem thisHDC, Screen.Fonts(item), A
gList1.ForeColor = ForeColor
End If
End Sub
 Private Sub PrintItem(mHdc As Long, c As String, r As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_VCENTER)
    DrawText mHdc, StrPtr(c), -1, r, way
    End Sub
Private Sub gList2_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = -1 Then
FillThere thisHDC, thisrect, gList2.CapColor
FillThereMyVersion thisHDC, thisrect, &H999999
skip = True
End If
End Sub






Private Sub glist3_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If gList3.EditFlag Then Exit Sub
    If gList3.List(0) = "" Then
    gList3.BackColor = &H808080
    gList3.ShowMe2
    Exit Sub
    End If
 
If Button = 1 Then
  gList3.LeftMarginPixels = gList3.WidthPixels - gList3.UserControlTextWidth(gList3.List(0)) / Screen.TwipsPerPixelX
       gList3.BackColor = rgb(0, 160, 0)
    gList3.ShowMe2
Else

    gList3.LeftMarginPixels = lastfactor * 5
  gList3.BackColor = &H808080
   gList3.ShowMe2


End If


End Sub



Private Sub glist3_LostFocus()
'If gList1.listindex > -1 Then Text1 = gList1.List(gList1.listindex)
gList3.BackColor = &H808080
gList3.ShowMe2
End Sub

Private Sub glist3_PanLeftRight(Direction As Boolean)

If TEXT1 = "" Then Exit Sub
If Direction Then
ReturnBold = gList4.ListSelected(1)
ReturnItalic = gList4.ListSelected(2)
ReturnSize = Val(Trim$(gList4.List(4)))
ReturnCharset = Val(Trim$(gList4.List(6)))
If gList1.listindex > -1 Then ReturnFontName = Screen.Fonts(gList1.listindex)
Unload Me
End If
End Sub

Private Sub gList3_Selected2(item As Long)
If item = -2 Then
If gList3.PanPos <> 0 Then
glist3_PanLeftRight (True)
Exit Sub
End If

 gList3.LeftMarginPixels = lastfactor * 5
gList3.BackColor = &H808080
gList3.ForeColor = &HE0E0E0
gList3.EditFlag = False
gList3.NoCaretShow = True
End If
gList3.ShowMe2
End Sub






Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Beep
End If
End Sub
Public Sub FillThereMyVersion2(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT, b As Long
b = CLng(Rnd * 3) + setupxy / 3

CopyFromLParamToRect A, thatRect
A.Left = A.Right - setupxy
A.top = b
A.Bottom = b + setupxy / 5
FillThere thathDC, VarPtr(A), thatbgcolor
A.top = b + setupxy / 5 + setupxy / 10
A.Bottom = b + setupxy \ 2
FillThere thathDC, VarPtr(A), thatbgcolor


End Sub
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

Function ScaleDialogFix(ByVal factor As Single) As Single
gList2.FontSize = 14.25 * factor
factor = gList2.FontSize / 14.25
gList1.FontSize = 11.25 * factor
gList4.FontSize = 11.25 * factor
factor = gList1.FontSize / 11.25

ScaleDialogFix = factor
End Function
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)
lastfactor = factor
gList1.addpixels = 10 * factor
gList3.FontSize = 11.25 * factor
setupxy = 20 * factor
gList1.LeftMarginPixels = 5 * factor
 gList3.LeftMarginPixels = factor * 5


bordertop = 10 * scrTwips * factor
borderleft = bordertop
Dim heightTop As Long, heightSelector As Long, HeightMenu As Long, HeightBottom As Long
Dim shapeHeight As Long
heightTop = 30 * factor * scrTwips
HeightBottom = 30 * factor * scrTwips
' some space here
heightSelector = 240 * factor * scrTwips
HeightMenu = 180 * factor * scrTwips

' some space here
HeightBottom = 30 * factor * scrTwips
If (NewWidth < 0) Or NewWidth <= (246 * scrTwips * factor) Then
NewWidth = 246 * scrTwips * factor
End If
itemWidth = (NewWidth - 2 * borderleft)
allwidth = NewWidth
Dim allheight As Long
gList2.FloatLimitTop = ScrY() - bordertop - heightTop
gList2.FloatLimitLeft = ScrX() - borderleft * 3

allheight = bordertop + heightTop + bordertop + heightSelector + bordertop + HeightMenu + bordertop + HeightBottom + bordertop

Move Left, top, allwidth, allheight
gList2.Move borderleft, bordertop, itemWidth, heightTop
gList1.Move borderleft, 2 * bordertop + heightTop, itemWidth, heightSelector
gList4.Move borderleft, 3 * bordertop + heightTop + heightSelector, itemWidth, HeightMenu
gList3.Move borderleft, allheight - HeightBottom - bordertop, itemWidth, HeightBottom


End Sub

Private Sub gList4_ChangeListItem(item As Long, content As String)
Dim content1 As Single
If item = 4 Then
content1 = Val("0" & Trim$(content))
If content1 > 144 Then
content = gList4.List(item)
Else
content = "  " & CStr(content1)
End If
ElseIf item = 6 Then
content1 = Val("0" & Trim$(content))
If content1 > 255 Then
content = gList4.List(item)
Else
content = "  " & CStr(content1)
End If
End If
End Sub


Private Sub gList4_GotFocus()
If gList4.EditFlag Then gList4.NoCaretShow = False
gList4.ShowMe2
End Sub

Private Sub gList4_LostFocus()
If gList4.EditFlag Then
If Val(Trim$(gList4.List(4))) < 6 Then gList4.List(4) = "  6"
End If
gList4.EditFlag = False
gList4.NoCaretShow = True
gList4.ShowMe2
End Sub

Private Sub gList4_MenuChecked(item As Long)
If item = 2 Then
If gList4.ListSelected(1) Then
gList3.Font.bold = True
Else
gList3.Font.bold = False
End If
ElseIf item = 3 Then
If gList4.ListSelected(2) Then
gList3.Font.Italic = True
Else
gList3.Font.Italic = False
End If
End If
gList3.ShowMe2
End Sub


Private Sub gList4_selected(item As Long)

If item = 5 Or item = 7 Then
If Not gList4.EditFlag Then
 gList4.EditFlag = True
 gList4.NoCaretShow = False
 gList4.ShowMe2
 If Val(Trim$(gList4.List(4))) < 6 Then gList4.List(4) = "  6"
End If
 Else
 gList4.EditFlag = False
 gList4.NoCaretShow = True
 End If
gList3.Font.charset = Trim$(gList4.List(6))

End Sub
Private Sub gList4_Selected2(item As Long)
Dim t$()
If Val(Trim$(gList4.List(4))) < 6 Then gList4.List(4) = "  6"
If item = 4 Or item = 6 Then
If Not gList4.EditFlag Then
 gList4.EditFlag = True
 gList4.NoCaretShow = False
 gList4.ShowMe2
 
End If
Else

 gList4.EditFlag = False
  gList4.NoCaretShow = False
 If item > 7 And item < 17 Then
 t$() = Split(gList4.List(item), " - ")
 gList4.List(6) = "  " + t$(UBound(t$()))
 gList4.ShowMe2
 ElseIf item > 17 Then
 gList4.List(4) = gList4.List(item)
 gList4.ShowMe2
 End If
End If

End Sub
Private Sub gList4_softSelected(item As Long)
gList4_selected item
End Sub

Private Sub glist3_KeyDown(KeyCode As Integer, shift As Integer)

If Not gList3.EditFlag Then


gList1.ShowMe2
gList3.SelStart = 1
     gList3.LeftMarginPixels = lastfactor * 5
  gList3.BackColor = &H808080
  
gList3.EditFlag = True
gList3.NoCaretShow = False
gList3.BackColor = &H0
gList3.ForeColor = &HFFFFFF
gList3.ShowMe2
ElseIf KeyCode = vbKeyReturn Then

DestroyCaret
If TEXT1 <> "" Then
gList3.EditFlag = False
gList3.Enabled = False

glist3_PanLeftRight True
KeyCode = 0
End If
End If

End Sub
Private Sub gList4_SpecialColor(rgbcolor As Long)
If gList4.EditFlag Then

ElseIf gList4.NoCaretShow Then
rgbcolor = rgb(255, 200, 125)
End If
End Sub
Private Sub gList4_ScrollSelected(item As Long, y As Long)
gList4_selected item
End Sub

Private Sub gList1_RegisterGlist(this As gList)
Set LastGlist3 = this
End Sub
Public Sub hookme(this As gList)
Set LastGlist3 = this
End Sub

VERSION 5.00
Begin VB.Form NeoMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   Icon            =   "NeoMsgBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   ScaleHeight     =   4920
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   375
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   3420
      _extentx        =   6033
      _extenty        =   873
      max             =   1
      vertical        =   -1  'True
      font            =   "NeoMsgBox.frx":000C
      enabled         =   -1  'True
      backcolor       =   3881787
      forecolor       =   16777215
      capcolor        =   16777215
   End
   Begin M2000.gList command1 
      Height          =   525
      Index           =   0
      Left            =   4590
      TabIndex        =   1
      Top             =   4245
      Width           =   3225
      _extentx        =   5689
      _extenty        =   926
      max             =   1
      vertical        =   -1  'True
      font            =   "NeoMsgBox.frx":0030
      showbar         =   0   'False
      forecolor       =   16777215
   End
   Begin M2000.gList command1 
      Height          =   525
      Index           =   1
      Left            =   705
      TabIndex        =   2
      Top             =   4305
      Width           =   3330
      _extentx        =   5874
      _extenty        =   926
      max             =   1
      vertical        =   -1  'True
      font            =   "NeoMsgBox.frx":0054
      showbar         =   0   'False
      forecolor       =   16777215
   End
   Begin M2000.gList gList1 
      Height          =   1995
      Left            =   3375
      TabIndex        =   3
      Top             =   960
      Width           =   4755
      _extentx        =   8387
      _extenty        =   3519
      max             =   1
      vertical        =   -1  'True
      font            =   "NeoMsgBox.frx":0078
      showbar         =   0   'False
      backcolor       =   3881787
      forecolor       =   16777215
   End
   Begin M2000.gList gList3 
      Height          =   315
      Left            =   3135
      TabIndex        =   4
      Top             =   3600
      Width           =   4590
      _extentx        =   8096
      _extenty        =   556
      max             =   1
      vertical        =   -1  'True
      font            =   "NeoMsgBox.frx":009C
      showbar         =   0   'False
   End
End
Attribute VB_Name = "NeoMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements InterPress
Public textbox1 As myTextBox
Public WithEvents ListPad As Document
Attribute ListPad.VB_VarHelpID = -1
Private Type myImage
    image As StdPicture
    Height As Long
    Width As Long
    top As Long
    Left As Long
End Type
Dim Image1 As myImage
'This is my new MsgBox
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Dim iTop As Long, iLeft As Long, iwidth As Long, iheight As Long
Dim setupxy As Single
Dim Lx As Long, ly As Long, dr As Boolean, drmove As Boolean
Dim prevx As Long, prevy As Long
Dim A$
Dim bordertop As Long, borderleft As Long
Dim allheight As Long, allwidth As Long, itemWidth As Long, itemwidth3 As Long, itemwidth2 As Long
Dim height1 As Long, width1 As Long
Dim myOk As myButton
Dim myCancel As myButton
Dim all As Long
Dim novisible As Boolean
Private mModalId As Variant


'
Property Get NeverShow() As Boolean
NeverShow = Not novisible
End Property






Private Sub Form_Deactivate()
'Set LastGlist = Nothing

  If ASKINUSE And Not Form2.Visible Then
    If Visible Then
    Me.SetFocus
    End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyPause And Not BreakMe Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Dim photo As Object

novisible = True
''Set LastGlist = Nothing

If AskCancel$ = "" Then command1(1).Visible = False
gList2.enabled = True
command1(0).enabled = True
command1(1).enabled = True
height1 = 2775 * DYP / 15
width1 = 7920 * DXP / 15
lastfactor = 1
LastWidth = -1

Set textbox1 = New myTextBox
Set textbox1.Container = gList3
textbox1.MaxCharLength = 100
If AskInput Then
gList3.Visible = True
textbox1 = AskStrInput$
textbox1.Locked = False
textbox1.enabled = True

Else
gList3.Visible = False  ' new from revision 17 (version 7)
End If
gList1.NoCaretShow = True
gList1.VerticalCenterText = True
gList1.LeftMarginPixels = 8
gList1.enabled = True
Set ListPad = New Document
ListPad = AskText$
If AskDIB$ = "" Then

Set LoadPictureMine = Form3.Icon
Else
    If Left$(AskDIB$, 4) = "cDIB" And Len(AskDIB$) > 12 Then
                Set photo = New cDIBSection
               If cDib(AskDIB$, photo) Then
                   photo.GetDpi 96, 96
                   Set LoadPictureMine = photo.Picture
               Else
                   Set LoadPictureMine = Form3.Icon
               End If
               Set photo = Nothing
       Else
               If CFname(AskDIB$) <> "" Then
                   Set LoadPictureMine = LoadPicture(GetDosPath(CFname(AskDIB$)))
               Else
                   Set LoadPictureMine = Form3.Icon
               End If
    End If
End If
lastfactor = ScaleDialogFix(SizeDialog)
ScaleDialog lastfactor, LastWidth
gList2.enabled = True
gList2.CapColor = rgb(255, 160, 0)
gList2.FloatList = True
gList2.MoveParent = True
gList2.HeadLine = ""
gList2.HeadLine = AskTitle$
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
Set myOk = New myButton
Set myOk.Container = command1(0)

  Set myOk.Callback = Me
  myOk.Index = 1
  myOk.Caption = AskOk$
myOk.enabled = True
Set myCancel = New myButton
Set myCancel.Container = command1(1)
myCancel.Caption = AskCancel$
  Set myCancel.Callback = Me
myCancel.enabled = True
ListPad.WrapAgain

all = ListPad.DocLines

gList1.ShowMe
If AskLastX = -1 And AskLastY = -1 Then

Else
Move AskLastX, AskLastY
End If
If AskInput Then
gList3.TabIndex = 1
End If
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
textbox1.Dereference
myOk.Shutdown
myCancel.Shutdown

gList1.Shutdown
gList2.Shutdown
gList3.Shutdown
command1(0).Shutdown
command1(1).Shutdown
novisible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set myOk = Nothing
Set myCancel = Nothing
AskDIB$ = ""
AskOk$ = ""
AskLastX = Left
AskLastY = top
''Sleep 200
ASKINUSE = False
End Sub

Private Sub gList1_ExposeListcount(cListCount As Long)
cListCount = all
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
                       AskCancel$ = ""
            Unload Me
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim addx As Long, addy As Long, factor As Single, Once As Boolean
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
     If x < (Width - 150) Or x > Width Then addx = (x - Lx)
     
Else
    If y < (Height - bordertop) Or y > Height Then addy = (y - ly)
        If x < (Width - borderleft) Or x > Width Then addx = (x - Lx)
    End If
    

    
   ''If Not ExpandWidth Then
   addx = 0
        If lastfactor = 0 Then lastfactor = 1
        factor = lastfactor

        
  
        Once = True
         If Width > ScrX() Then addx = -(Width - ScrX()) + addx
        If Height > ScrY() Then addy = -(Height - ScrY()) + addy
      
        If (addy + Height) / height1 > 0.4 And ((Width + addx) / width1) > 0.4 Then
   
        If addy <> 0 Then
        If ((addy + Height) / height1) * width1 > ScrX() * 0.9 Then
        addy = 0: addx = 0

        Else
        SizeDialog = ((addy + Height) / height1)
        End If
        End If
        lastfactor = ScaleDialogFix(SizeDialog)


        If ((Width * lastfactor / factor + addx) / Height * lastfactor / factor) < (width1 / height1) Then
        addx = -Width * lastfactor / factor - 1
      
           End If

        If addx = 0 Then
        If lastfactor <> factor Then ScaleDialog lastfactor, Width
        Lx = x
        
        Else
        Lx = x * lastfactor / factor
         ScaleDialog lastfactor, (Width + addx) * lastfactor / factor
         End If

        
         
        
        LastWidth = Width
              gList2.HeadlineHeight = gList2.HeightPixels
        gList2.PrepareToShow
        gList1.PrepareToShow
          ListPad.WrapAgain
        all = ListPad.DocLines
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
Dim aa As Long
On Error Resume Next
lastfactor = factor
setupxy = 16 * factor
bordertop = 8 * dv15 * factor
If AskInput Then aa = 4 Else aa = 5
gList1.StickBar = True
gList1.addpixels = 3 * factor
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

gList1.Width = itemwidth3 * 2 + borderleft
 ListPad.WrapAgain: all = ListPad.DocLines

If AskCancel$ <> "" And ListPad.DocLines < aa Then
gList1.restrictLines = ListPad.DocLines
If AskInput Then
gList1.Move borderleft * 2 + itemwidth3, bordertop * (4 + (aa - ListPad.DocLines) * 2), itemwidth3 * 2 + borderleft, bordertop * ListPad.DocLines * 3
Else
gList1.Move borderleft * 2 + itemwidth3, bordertop * (4 + (5 - ListPad.DocLines) * 2), itemwidth3 * 2 + borderleft, bordertop * ListPad.DocLines * 3
End If
Else
gList1.restrictLines = aa
If AskInput Then
gList1.Move borderleft * 2 + itemwidth3, bordertop * 5, itemwidth3 * 2 + borderleft, bordertop * 9
Else
gList1.Move borderleft * 2 + itemwidth3, bordertop * 5, itemwidth3 * 2 + borderleft, bordertop * 12
End If
End If
If AskInput Then
gList3.Move borderleft * 2 + itemwidth3, bordertop * 15, itemwidth3 * 2 + borderleft, bordertop * 3

End If
If AskCancel$ <> "" Then
command1(1).Move borderleft, bordertop * 19, itemwidth2, bordertop * 3
command1(0).Move borderleft + itemwidth2 + borderleft, bordertop * 19, itemwidth2, bordertop * 3
Else
command1(0).Move borderleft, bordertop * 19, itemWidth, bordertop * 3
End If
If iwidth = 0 Then iwidth = itemwidth3
If iheight = 0 Then iheight = bordertop * 12
Dim curIwidth As Long, curIheight As Long, sc As Single
If Image1.Width > 0 Then
curIwidth = Image1.Width
curIheight = Image1.Height
iLeft = borderleft
iTop = 5 * bordertop
iwidth = itemwidth3
iheight = bordertop * 12
 Line (0, 0)-(ScaleWidth - dv15, ScaleHeight - dv15), Me.BackColor, BF
If (curIwidth / iwidth) < (curIheight / iheight) Then
sc = curIheight / iheight
ImageMove Image1, iLeft + (iwidth - curIwidth / sc) / 2, iTop, curIwidth / sc, iheight
Else
sc = curIwidth / iwidth
ImageMove Image1, iLeft, iTop + (iheight - curIheight / sc) / 2, iwidth, curIheight / sc
End If
End If
End Sub
Function ScaleDialogFix(ByVal factor As Single) As Single
gList2.FontSize = 14.25 * factor
gList1.FontSize = 13.5 * factor
gList3.FontSize = 13.5 * factor

factor = gList2.FontSize / 14.25
command1(0).FontSize = 11.75 * factor
factor = gList1.FontSize / 11.75
command1(1).FontSize = command1(0).FontSize
ScaleDialogFix = factor
End Function

Public Property Set LoadApicture(aImage As StdPicture)
On Error Resume Next
Dim sc As Double
Set Image1.image = Nothing
Image1.Width = 0
If aImage.handle <> 0 Then
Set Image1.image = aImage
If (aImage.Width / iwidth) < (aImage.Height / iheight) Then
sc = aImage.Height / iheight
ImageMove Image1, iLeft + (iwidth - aImage.Width / sc) / 2, iTop, aImage.Width / sc, iheight
Else
sc = aImage.Width / iwidth
ImageMove Image1, iLeft, iTop + (iheight - aImage.Height / sc) / 2, iwidth, aImage.Height / sc
End If
End If


Image1.Height = aImage.Height
Image1.Width = aImage.Width
End Property

Public Property Set LoadPictureMine(aImage As StdPicture)
On Error Resume Next
Dim sc As Double
Set Image1.image = Nothing
Image1.Width = 0
If aImage.handle <> 0 Then
Set Image1.image = aImage
Image1.Height = aImage.Height
Image1.Width = aImage.Width
End If
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
Private Sub ImageMove(A As myImage, neoTop As Long, NeoLeft As Long, NeoWidth As Long, NeoHeight As Long)
If A.image Is Nothing Then Exit Sub
If A.image.Width = 0 Then Exit Sub
If A.image.Type = vbPicTypeIcon Then
Dim aa As New cDIBSection
aa.BackColor = BackColor
aa.CreateFromPicture A.image
aa.ResetBitmapTypeToBITMAP
PaintPicture aa.Picture, neoTop, NeoLeft, NeoWidth, NeoHeight
Else
PaintPicture A.image, neoTop, NeoLeft, NeoWidth, NeoHeight
End If

End Sub



Private Sub gList2_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyEscape Then
                AskCancel$ = ""
            Unload Me

End If
End Sub

Private Sub gList3_Selected2(item As Long)

 command1(0).SetFocus
End Sub

Private Sub InterPress_Press(Index As Long)
If Index = 0 Then
AskResponse$ = AskCancel$
AskCancel$ = ""
Else
If AskInput Then AskStrInput$ = textbox1
AskResponse$ = AskOk$
End If

AskOk$ = ""
Unload Me
End Sub
Private Sub glist1_ReadListItem(item As Long, Content As String)

If item >= 0 Then
Content = ListPad.TextLine(item + 1)
End If
End Sub
Private Sub ListPad_BreakLine(data As String, datanext As String)
    gList1.BreakLine data, datanext
End Sub


Private Sub gList2_RefreshDesktop()
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
End Sub

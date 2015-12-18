VERSION 5.00
Begin VB.Form MyPopUp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "FreeSans"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin M2000.gList gList1 
      Height          =   5475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   9657
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FreeSans"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      dcolor          =   32896
      Backcolor       =   3881787
      ForeColor       =   14737632
      CapColor        =   9797738
   End
End
Attribute VB_Name = "MyPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gokeyboard As Boolean, lastitem As Long, part1 As String, lastgoodnum As Long
Private height1, width1
Dim Lx As Long, ly As Long, dr As Boolean
Dim bordertop As Long, borderleft As Long, lastshift As Integer
Dim allheight As Long, allwidth As Long, itemWidth As Long
Public Sub Up(Optional x As Variant, Optional y As Variant)
If IsMissing(x) Then
x = CSng(MOUSEX())
y = CSng(MOUSEY())
Else
x = x + Form1.Left
y = y + Form1.top
End If

If x + Width > ScrX() Then
If y + Height > ScrY() Then
Move ScrX() - Width, ScrY() - Height
Else
Move ScrX() - Width, y
End If
ElseIf y + Height > ScrY() Then
Move x, ScrY() - Height
Else
Move x, y
End If
Show
MyDoEvents
''gList1.SetFocus
End Sub


Private Sub Form_Load()
part1 = " " + GetStrUntil("(", (textinformCaption)) + "("
With gList1
.NoWheel = True
.restrictLines = 14
.FloatList = True
.MoveParent = True
.SingleLineSlide = True
.NoPanRight = True
.AutoHide = True
End With
height1 = 5475 * DYP / 15
width1 = 4155 * DXP / 15
If pagio$ = "GREEK" Then
With gList1
''
.StickBar = False
''.AddPixels = 4
.VerticalCenterText = True
.additemFast textinformCaption
.addsep
.additemFast "Αποκοπή Ctrl+X"
.menuEnabled(2) = Form1mn1Enabled
.additemFast "Αντιγραφή Ctrl+C"
.menuEnabled(3) = Form1mn2Enabled
.additemFast "Επικόλληση Ctrl+V"
.menuEnabled(4) = Form1mn3Enabled
.addsep
.additemFast "Έξοδος με αλλαγές (ESC)"
.addsep
.additemFast "Έξοδος χωρίς αλλαγές shift F12"
.addsep
.additemFast "Αναζήτησε πάνω F2"
.menuEnabled(10) = Form1supEnabled
.additemFast "Αναζήτησε κάτω F3"
.menuEnabled(11) = Form1sdnEnabled
.additemFast "Κάνε το ίδιο παντού F4"
.menuEnabled(12) = Form1mscatEnabled
.additemFast "Αλλαγή λέξης F5"
.menuEnabled(13) = Form1rthisEnabled
.addsep
.additemFast "Αναδίπλωση λέξεων F1"

.MenuItem 16, True, False, Not Form1.TEXT1.nowrap, "warp"
.additemFast "Μεταφορά Κειμένου"
.MenuItem 17, True, False, Form1.TEXT1.glistN.DragEnabled, "drag"
.additemFast "Χρώμα/Σύμπτυξη Γλώσσας F11"
.MenuItem 18, True, False, shortlang, "short"
.additemFast "Εμφάνιση Παραγράφων F10"
.MenuItem 19, True, False, Form1.TEXT1.showparagraph, "para"
.additemFast "Μέτρηση λέξεων F9"
.addsep
.additemFast "Βοήθεια ctrl+F1"
If Not Form1.EditTextWord Then
.HeadLine = "Μ2000 Συντάκτης"
.addsep
.additemFast "Τμήματα/Συναρτήσεις F12"
.menuEnabled(23) = SubsExist()

End If
End With
Else
With gList1
''gList1.HeadLine = "Μ2000"
.StickBar = False
''''.AddPixels = 4
.VerticalCenterText = True
.additemFast textinformCaption
.addsep
.additemFast "Cut   Ctrl+X"
.menuEnabled(2) = Form1mn1Enabled
.additemFast "Copy  Ctrl+C"
.menuEnabled(3) = Form1mn2Enabled
.additemFast "Paste Ctrl+V"
.menuEnabled(4) = Form1mn3Enabled
.addsep
.additemFast "Save and Exit (ESC)"
.addsep
.additemFast "Discard Changes shift F12"
.addsep
.additemFast "Search up F2"
.menuEnabled(10) = Form1supEnabled
.additemFast "Search down F3"
.menuEnabled(11) = Form1sdnEnabled
.additemFast "Make same all F4"
.menuEnabled(12) = Form1mscatEnabled
.additemFast "Replace word F5"
.menuEnabled(13) = Form1rthisEnabled
.addsep
.additemFast "Word Wrap F1"

.MenuItem 16, True, False, Not Form1.TEXT1.nowrap, "warp"
.additemFast "Drag Enabled"
.MenuItem 17, True, False, Form1.TEXT1.glistN.DragEnabled, "drag"
.additemFast "Color/Short Language F11"
.MenuItem 18, True, False, shortlang, "short"
.additemFast "Paragraph Mark F10"
.MenuItem 19, True, False, Form1.TEXT1.showparagraph, "para"
.additemFast "Word count F9"
.addsep
.additemFast "Help ctrl+F1"
If Not Form1.EditTextWord Then
.HeadLine = "Μ2000 Editor"
.addsep
.additemFast "Modules/Functions F12"
.menuEnabled(23) = SubsExist()
End If

End With
End If
If Pouplastfactor = 0 Then Pouplastfactor = 1
 Pouplastfactor = ScaleDialogFix(helpSizeDialog)
If ExpandWidth And False Then
If PopUpLastWidth = 0 Then PopUpLastWidth = -1
Else
PopUpLastWidth = -1
End If
If ExpandWidth Then
If PopUpLastWidth = 0 Then PopUpLastWidth = -1
Else
PopUpLastWidth = -1
End If
ScaleDialog Pouplastfactor, PopUpLastWidth
gList1.listindex = 0
gList1.ShowBar = True
gList1.ShowBar = False
gList1.NoPanLeft = False
gList1.SoftEnterFocus
End Sub
Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)

If Button = 1 Then
    
    If Pouplastfactor = 0 Then Pouplastfactor = 1

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
        If Pouplastfactor = 0 Then Pouplastfactor = 1
        factor = Pouplastfactor

        
  
        Once = True
        If Height > ScrY() Then addy = -(Height - ScrY()) + addy
        If Width > ScrX() Then addX = -(Width - ScrX()) + addX
        If (addy + Height) / height1 > 0.4 And ((Width + addX) / width1) > 0.4 Then
   
        If addy <> 0 Then helpSizeDialog = ((addy + Height) / height1)
        Pouplastfactor = ScaleDialogFix(helpSizeDialog)


        If ((Width * Pouplastfactor / factor + addX) / Height * Pouplastfactor / factor) < (width1 / height1) Then
        addX = -Width * Pouplastfactor / factor - 1
      
           End If

        If addX = 0 Then
        
        If Pouplastfactor <> factor Then ScaleDialog Pouplastfactor, Width

        Lx = x
        
        Else
        Lx = x * Pouplastfactor / factor
             ScaleDialog Pouplastfactor, (Width + addX) * Pouplastfactor / factor
         
   
         End If

        
        PopUpLastWidth = Width


''gList1.PrepareToShow
        ly = ly * Pouplastfactor / factor
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



Private Sub gList1_ChangeListItem(item As Long, content As String)
Dim content1 As Long
If item = 0 Then
content1 = Int(Val("0" & Trim$(Mid$(content, Len(part1) + 1))))
        If content1 > Form1.TEXT1.mDoc.DocLines Or content1 < 0 Then
        content = gList1.List(item)
              gList1.SelStart = Len(gList1.List(item)) - 1
        Else
        lastgoodnum = content1
        If content1 = 0 Then
        content = part1 & ")"
        gList1.SelStart = 3
        Else
        content = part1 & CStr(content1) & ")"
        End If
        
        End If
End If
End Sub

Private Sub gList1_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
''If X * dv15 > Width / 2 Then

If item = -1 Then

Else
gList1.mousepointer = 1
If gokeyboard Then Exit Sub
gList1.EditFlag = False
''''''''''''''''''''''''''''''
If lastitem = item Then Exit Sub
gList1.ListindexPrivateUse = item
gList1.ShowMe2

lastitem = item
gList1.ListindexPrivateUse = -1
End If
End Sub

Private Sub gList1_KeyDown(KeyCode As Integer, shift As Integer)
gokeyboard = True
If KeyCode = vbKeyEscape Then Unload Me: Exit Sub

If gList1.listindex = -1 Then gList1.ListindexPrivateUse = lastitem

If KeyCode >= vbKey0 And KeyCode <= vbKey9 And gList1.EditFlag = False And gList1.listindex = 0 Then
                        lastitem = 0
                    gList1.PromptLineIdent = Len(part1)
                    gList1.List(0) = ""
                    gList1.SelStart = 3
                    gList1.EditFlag = True

ElseIf gList1.listindex = 0 And gList1.EditFlag = True Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn And gList1.EditFlag = True Then
        gList1.EditFlag = False
        lastitem = 0
        KeyCode = 0
        DoCommand 1
        gList1.ListindexPrivateUse = 0
        gList1.ShowMe2
        
        lastitem = 0
        gList1.ListindexPrivateUse = -1
        End If
End If



End Sub



Private Sub gList1_LostFocus()
Unload Me
End Sub

Private Sub gList1_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)

gokeyboard = False
gList1.PromptLineIdent = 0
If lastitem = item Then Exit Sub
gList1.ListindexPrivateUse = -1
End Sub

Private Sub gList1_ScrollSelected(item As Long, y As Long)
If gokeyboard Then Exit Sub

gList1.EditFlag = False
''''''''''''''''''''''
If lastitem = item - 1 Then Exit Sub
gList1.ListindexPrivateUse = item - 1
gList1.ShowMe2

lastitem = item
gList1.ListindexPrivateUse = -1
End Sub

Private Sub gList1_selected(item As Long)
If Not gokeyboard Then
DoCommand item
Else
If Not gList1.EditFlag Then
gList1.ListindexPrivateUse = item - 1
lastitem = gList1.listindex
gList1.ShowMe2


End If
End If
End Sub

Private Sub gList1_selected2(item As Long)
 DoCommand item + 1

End Sub
Private Sub DoCommand(item As Long)
Select Case item - 1
Case -2
Exit Sub
Case 0
If lastgoodnum > 0 Then
With gList1
.menuEnabled(2) = False
.menuEnabled(3) = False
Form1.TEXT1.SetRowColumn lastgoodnum, 0
.PromptLineIdent = 0
lastitem = 0
.ListindexPrivateUse = -1



Exit Sub
End With
End If
Case 2
Form1.mn1sub
Case 3
Form1.mn2sub
Case 4
Form1.mn3sub
Case 6
Form1.mn4sub
Case 8
Form1.mn5sub
Case 10
Form1.supsub
Case 11
Form1.sdnSub
Case 12
Form1.mscatsub
Case 13
Form1.rthissub
Case 15
gList1.ListSelectedNoRadioCare(17) = Not gList1.ListChecked(17)
Form1.wordwrapsub
Case 16
gList1.ListSelectedNoRadioCare(18) = Not gList1.ListChecked(18)
Form1.TEXT1.glistN.DragEnabled = Not Form1.TEXT1.glistN.DragEnabled
Case 21
Form1.helpmeSub
Case 23
showmodules
Case 17
With Form1.TEXT1
shortlang = Not shortlang
.ManualInform
End With
Case 18
With Form1.TEXT1
.showparagraph = Not .showparagraph
.Render
End With
Case 19
With Form1.TEXT1
If .glistN.lines > 1 Then
If UserCodePage = 1253 Then
.ReplaceTitle = "Λέξεις στο κείμενο:" + CStr(.mDoc.WordCount)
Else
.ReplaceTitle = "Words in text:" + CStr(.mDoc.WordCount)
End If
End If
End With

End Select
Unload Me
End Sub

Function ScaleDialogFix(ByVal factor As Single) As Single
gList1.FontSize = 11.25 * factor
factor = gList1.FontSize / 11.25
ScaleDialogFix = factor
End Function
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)
Dim h As Long, i As Long
Pouplastfactor = factor
gList1.LeftMarginPixels = 30 * factor
setupxy = 20 * factor
bordertop = 10 * dv15 * factor
borderleft = bordertop
If (NewWidth < 0) Or NewWidth <= width1 * factor Then
NewWidth = width1 * factor
End If
allwidth = NewWidth  ''width1 * factor
allheight = height1 * factor
itemWidth = allwidth - 2 * borderleft
''MyForm Me, Left, top, allwidth, allheight, True, factor

Move Left, top, allwidth, allheight
  
gList1.addpixels = 4 * factor

gList1.Move borderleft, bordertop, itemWidth, allheight - bordertop * 2

gList1.CalcAndShowBar
gList1.ShowBar = False
gList1.FloatLimitTop = ScrY() - bordertop - bordertop * 3
gList1.FloatLimitLeft = ScrX() - borderleft * 3

End Sub


Private Sub gList1_SpecialColor(rgbcolor As Long)
rgbcolor = rgb(100, 132, 254)
End Sub
Public Sub hookme(this As gList)
If Not this Is Nothing Then this.NoWheel = True
End Sub


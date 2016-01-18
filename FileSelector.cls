VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim netdisplay As Boolean
Private LastItemCount As Long
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Public WithEvents myDir2 As recDir
Attribute myDir2.VB_VarHelpID = -1
Dim hiddenDir As String
Public WithEvents mDoc1 As Document
Attribute mDoc1.VB_VarHelpID = -1
Public TEXT1 As myTextBox
Attribute TEXT1.VB_VarHelpID = -1
Public Mydir As recDir
Private path$
Event TraceFile(file$)
Event doubleclick(file$)
Event controlmenu(check As Boolean)
Event NewHeadline(newpath$)
Private fakeitemcount As Long
Dim schedule As String
Dim skipselected2 As Boolean
Private mPreserveNpixelsHeaderRight As Long
Private onelineheight As Long
Public recnowchecked As Boolean
Public recnow3checked As Boolean
Public mselChecked  As Boolean
Public NoHeadLine As Boolean
Private mNostateDir As Boolean
Private mselectedFile As String
Dim filetypes As String
Dim ReturnFile As String
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Text1Dir As String
Private topoffsetpixel As Long

Property Let SortType(ByVal rhs As Long)
If Not Mydir.SortType = rhs Then
Mydir.SortType = rhs
If path$ <> "" Then reload
End If
End Property

Property Let FileTypesToDisplay(ByVal rhs As String)
filetypes = rhs
End Property
Property Let FilePath(rhs As String)
Dim first As String
glistN.Enabled = False
If rhs = "" Or Not (path$ = rhs) Then
path$ = rhs
mDoc1.textDoc = rhs
''glistN.AddPixels = 10
If Not NoHeadLine Then
glistN.HeadLine = " " ' give a fake headline
End If
onelineheight = glistN.HeadlineHeight
glistN.HeadlineHeight = onelineheight + 2

fakeitemcount = 0
myDir2.IncludedFolders = True
myDir2.Nofiles = Mydir.Nofiles
myDir2.SortType = Mydir.SortType
myDir2.TopFolder = Mydir.TopFolder
If recnow3checked Then myDir2.LevelStop = 3 Else myDir2.LevelStop = 0
first = myDir2.Dir2(rhs, filetypes, recnowchecked)
glistN.ShowMe
glistN.StickBar = True
glistN.ShowBar = False
End If

End Property
Property Get FilePath() As String
FilePath = path$
End Property
Private Sub Class_Initialize()
Set Mydir = New recDir
Set myDir2 = New recDir
Set mDoc1 = New Document
End Sub

Private Sub glistN_BlockCaret(item As Long, blockme As Boolean, skipme As Boolean)
If Not NostateDir Then
If Mydir.path = "" Then
If item >= 0 Then
If (Mydir.ReadLevel(item) - 1) > 0 Then
blockme = True
skipme = True
End If
End If
End If
End If
End Sub

Private Sub glistN_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If NostateDir Then
' as a setup menu
Else
If Not mselChecked Then Exit Sub
If Mydir.List(item) = "..." Or Mid$(Mydir.List(item), 1) = Chr$(2) Or item < 0 Then Exit Sub
If Button = 1 Then

If x < 20 * lastfactor And y < 20 * lastfactor Then
If Mydir.ReadMark(item) Then
Mydir.unMarkLevel item
Else
Mydir.MarkLevel item
End If
glistN.ShowMe2
Button = 0
End If
End If
End If
End Sub

Private Sub glistN_ExposeListcount(cListCount As Long)
If NostateDir Then Exit Sub
If fakeitemcount = 0 Then fakeitemcount = 1
cListCount = fakeitemcount
End Sub
Private Function mylcasefILE(ByVal A$) As String
A$ = NLtrim$(A$)
If casesensitive Then
mylcasefILE = A$
Else
 mylcasefILE = LCase(A$) '!!!!!!!!!!
 End If

End Function
Private Sub glistN_Find(key As String, where As Long, skip As Boolean)
If NostateDir Then Exit Sub

skip = True
' find file only
Dim i As Long, s$
s$ = mylcasefILE(key)
'Command2.Enabled = False
For i = 0 To Mydir.listcount - 1
If Mydir.abort Then Exit For
If mylcasefILE(Mydir.List(i)) = s$ Then

'If Mydir.List(I) Like "*" & key Then
'********************************************
where = i
''Command2.Enabled = True
Exit For
End If
If Left$(Mydir.List(i), 1) = Chr$(2) Then Exit For
Next i
If Mydir.abort Then where = -1: Exit Sub
'If Not Command2.Enabled Then
'where = -1
'Else
'where = i

'End If


End Sub


Private Sub glistN_ScrollSelected(item As Long, y As Long)
Dim m As Double
If NostateDir Then
' as a setup menu
Else
' item is from 1
If Mydir.List(item - 1) = "..." Then
If Mydir.Nofiles Or (netdisplay And item > LastItemCount) Then
TEXT1 = SelectFolderButton
End If
ElseIf Not glistN.ListSep(item - 1) Then
If Left$(Mydir.List(item - 1), 1) = Chr$(2) Then
'Caption = Mid$(Mydir.List(item - 1), 2)
If Mydir.path = "" Then
If netdisplay And item > LastItemCount Then
If Mid$(Mydir.List(item - 1), 2, 1) = "(" Then
TEXT1 = FindNetworkFolderPath(Mid$(Mydir.List(item - 1), 2))
Else
TEXT1 = Mid$(Mydir.List(item - 1), 2)
End If
Else
m = FreeDiskSpace(Mid$(Mydir.List(item - 1), 2, 1))
If m < 0 Then reload: Exit Sub

If m > 1000 Then
TEXT1 = Format(m / 1000, "##,###,##0.0") & " GB Free"
Else
TEXT1 = Format(m, "##,###,##0.0") & " MB Free"
End If
End If
ElseIf Mydir.Nofiles Then
If Mid$(Mydir.List(item - 1), 2, 1) = "(" Then
TEXT1 = FindNetworkFolderPath(Mid$(Mydir.List(item - 1), 2))
Else
TEXT1 = "\" & Mydir.ExtractName(Mid$(Mydir.List(item - 1), 2))
End If
End If
Else
    If recnowchecked Then
  RaiseEvent TraceFile(Mydir.FindFolder(item - 1) + Mydir.List(item - 1))
    Else
    RaiseEvent TraceFile(Mydir.path + Mydir.List(item - 1))
    End If
End If
End If
End If
End Sub

Private Sub glistN_Selected(item As Long)
Dim m As Double
If NostateDir Then
' as a setup menu
ElseIf Not glistN.ListSep(item - 1) Then
' item is from 1
' new addition locked=false
TEXT1.Locked = False

If Mydir.List(item - 1) = "..." Then
If Mydir.Nofiles Then
TEXT1 = SelectFolderButton
End If
Else
If Left$(Mydir.List(item - 1), 1) = Chr$(2) Then
'Caption = Mid$(Mydir.List(item - 1), 2)
If Mydir.path = "" Then
If netdisplay And item > LastItemCount Then

If Mid$(Mydir.List(item - 1), 2, 1) = "(" Then
TEXT1 = FindNetworkFolderPath(Mid$(Mydir.List(item - 1), 2))
Else
TEXT1 = Mid$(Mydir.List(item - 1), 2)
End If
Else
m = FreeDiskSpace(Mid$(Mydir.List(item - 1), 2, 1))
If m < 0 Then reload: Exit Sub
If m > 1000 Then
TEXT1 = Format(m / 1000, "##,###,##0.0") & " GB Free"
Else
TEXT1 = Format(m, "##,###,##0.0") & " MB Free"
End If
End If
ElseIf Mydir.Nofiles Then
If Mid$(Mydir.List(item - 1), 2, 1) = "(" Then
TEXT1 = FindNetworkFolderPath(Mid$(Mydir.List(item - 1), 2))
Else
TEXT1 = "\" & Mydir.ExtractName(Mid$(Mydir.List(item - 1), 2))
End If
End If
Else
    If recnowchecked Then
    Text1Dir = Mydir.FindFolder(item - 1) + Mydir.List(item - 1)
  RaiseEvent TraceFile(Mydir.FindFolder(item - 1) + Mydir.List(item - 1))
  ' Text1 = Mydir.List(item - 1)
    Else
    Text1Dir = Mydir.path + Mydir.List(item - 1)
    RaiseEvent TraceFile(Mydir.path + Mydir.List(item - 1))
    ' Text1 = Mydir.List(item - 1)
    End If
End If
End If
End If
End Sub
Public Sub AbordAll()
glistN.Enabled = False
TEXT1.glistN.Enabled = False
If Not myDir2 Is Nothing Then
myDir2.abort = True
Set myDir2 = Nothing
skipselected2 = False
disablemenu True
End If
End Sub
Public Function IamBusy() As Boolean
IamBusy = skipselected2
End Function
Private Sub glistN_Selected2(item As Long)
Dim netselect As String
If NostateDir Then
' as a setup menu
Else

If Not myDir2 Is Nothing Then
myDir2.abort = True
Set myDir2 = Nothing
skipselected2 = False
disablemenu True
End If
' this selected file isn't the same as the GET
selectedFile = TEXT1
If netdisplay Then
If Mid$(Mydir.List(item), 2, 1) = "(" Then
' change with path
netselect = FindNetworkFolderPath(Mid$(Mydir.List(item), 2)) ' again because maybe is closed
End If
End If
If item < 0 Then Exit Sub ' -2 is a signal and here is not used
If skipselected2 Then Exit Sub ' like enabled=false...but other events works fine
If glistN.listcount = 0 Then Exit Sub
skipselected2 = True
Dim anystr As String
Set myDir2 = New recDir ' we get a new...the old is reference by myDir
myDir2.IncludedFolders = Mydir.IncludedFolders
myDir2.Nofiles = Mydir.Nofiles
myDir2.SortType = Mydir.SortType
myDir2.TopFolder = Mydir.TopFolder
If recnow3checked Then myDir2.LevelStop = 3
If Mydir.List(item) = "..." Then

If GetPath <> "" Then
If Len(Mydir.path) < 4 Or (ExtractPath(ExtractPath(Mydir.path, False), , True) = "") Then
path$ = ""
Else
            If Right$(GetPath, 1) = "\" Then
            path$ = Mydir.ExtractPath(Left$(GetPath, Len(GetPath) - 1))
            Else
            path$ = Mydir.ExtractPath(GetPath)
            
            End If
End If
anystr = myDir2.Dir2(path$, filetypes, recnowchecked)
End If
ElseIf Left$(Mydir.List(item), 1) = Chr$(2) Then
myDir2.IncludedFolders = True
If netdisplay And netselect <> "" Then
anystr = myDir2.Dir2(netselect, filetypes, recnowchecked)
Else
anystr = myDir2.Dir2(Mid$(Mydir.List(item), 2), filetypes, recnowchecked)
End If
Else

If recnowchecked Then
If Mydir.path = "" Then
TEXT1 = Mid$(Mydir.List(item), 2)
Else
TEXT1 = Mydir.FindFolder(item) + Mydir.List(item)
End If
  RaiseEvent doubleclick(Mydir.FindFolder(item) + Mydir.List(item))
Else
If Mydir.path = "" Then
TEXT1 = Mid$(Mydir.List(item), 2)
Else
TEXT1 = Mydir.List(item)
End If
 RaiseEvent doubleclick(Mydir.path + Mydir.List(item))
End If

skipselected2 = False
End If

If skipselected2 Then glistN.ShowMe2
End If
End Sub

 Private Sub PrintItem(mHdc As Long, c As String, r As RECT, Optional way As Long = DT_NOPREFIX Or DT_NOCLIP)
    DrawText mHdc, StrPtr(c), -1, r, way
    End Sub
Public Property Get HeadLinesNum() As Long
HeadLinesNum = mDoc1.DocLines
End Property

Private Sub gListN_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
' item is from 0
Dim A As RECT, b As RECT
Dim what$, ThatPara As Long, i As Long
If NostateDir Then
' as a setup menu
If item <> -1 Then skip = False
Else
CopyFromLParamToRect A, thisrect
CopyFromLParamToRect b, thisrect
A.Left = A.Left + 1
A.Right = glistN.WidthPixels
b.Right = glistN.WidthPixels
If item < 0 Then  ' this is a headline
If Not skip Then FillBack thisHDC, b, glistN.CapColor
      mDoc1.TextParagraphLine 1, what$, ThatPara
      glistN.HeadlineHeight = mDoc1.DocLines * (onelineheight + 1) + 2
      A.top = A.top + 1
      For i = 1 To mDoc1.DocLines
        PrintItem thisHDC, mDoc1.TextLine(i), A
        A.top = A.top + onelineheight + 1
      Next i
 If topoffsetpixel = 0 Then topoffsetpixel = glistN.LineTopOffsetPixels
     
      skip = True
       Exit Sub
       



ElseIf Left$(Mydir.List(item), 1) = Chr$(2) Then
                'DrawFrameControl thisHDC, a, 0, item
                A.Left = A.Left + 1
                
              A.top = A.top + topoffsetpixel
                
                   If item = glistN.listindex And Not glistN.ListSep(item) Then
            'SetTextColor thisHDC, 0
                 ' b.Top = b.Bottom - 1
                FillBack thisHDC, b, &HBFBFBB
                SetTextColor thisHDC, &HF0F0FF
                   Else
                 FillBack thisHDC, b, &H999999
                SetTextColor thisHDC, &HDDDDDD 'glistN.ForeColor
                End If
       
                 If Mydir.path = "" Then
                 A.Left = A.Left + 1
                 If (Mydir.ReadLevel(item) - 1) > 0 Then SetTextColor thisHDC, &H666666 ' &H999999
                 PrintItem thisHDC, Mid$(Mydir.List(item), 2), A
                 Else
                 A.Left = A.Left + 1 + (Mydir.ReadLevel(item) - 1) * 30 * lastfactor
                PrintItem thisHDC, "\" + Mydir.ExtractName(Mydir.List(item)), A
                End If
Else
    If item = glistN.listindex Then
    
             If Mydir.ReadMark(item) And multifileselection Then
            FillBack thisHDC, b, &HFF '&HFFFF00
            Else
             FillBack thisHDC, b, 0
            End If
            SetTextColor thisHDC, &HF0F0FF ' 0
              b.top = b.Bottom - 1
       
            FillBack thisHDC, b, &H777777
           
    Else
          If Mydir.ReadMark(item) And multifileselection Then
            FillBack thisHDC, b, &HAA
            End If
    SetTextColor thisHDC, glistN.ForeColor
    b.top = b.Bottom - 1
   FillBack thisHDC, b, 0
    End If
    A.Left = A.Left + 1 + (Mydir.ReadLevel(item) - 1) * 30 * lastfactor
     A.top = A.top + topoffsetpixel
    PrintItem thisHDC, Mydir.ExtractName(Mydir.List(item)), A

    
End If
skip = True
End If
End Sub
Private Sub gListN_SpecialColor(rgbcolor As Long)
rgbcolor = &HF0F0FF
End Sub





Private Sub glistN_softSelected(item As Long)
If NostateDir Then

Else

If Mydir.List(item - 1) = "..." Then

Else
If Left$(Mydir.List(item - 1), 1) = Chr$(2) Then

Else
    If recnowchecked Then
  RaiseEvent TraceFile(Mydir.FindFolder(item - 1) + Mydir.List(item - 1))
    Else
    RaiseEvent TraceFile(Mydir.path + Mydir.List(item - 1))
    End If
End If
End If
End If

End Sub

Private Sub glistN_SyncKeyboard(ascii As Integer)
Dim item As Long
If ascii = 32 Then

If Not NostateDir Then
If multifileselection Then

item = glistN.listindex
If item >= 0 Then
ascii = 0
If Mydir.ReadMark(item) Then
Mydir.unMarkLevel item
Else
Mydir.MarkLevel item
End If
End If
End If
End If
End If
End Sub

Private Sub mDoc1_BreakLine(data As String, datanext As String)
If PreserveNpixelsHeaderRight > 0 Then
glistN.BreakLine data, datanext, glistN.Pixels2Twips(PreserveNpixelsHeaderRight)
Else
glistN.BreakLine data, datanext
End If
End Sub
Sub disablemenu(Optional DIS As Boolean = False)
RaiseEvent controlmenu(DIS)
End Sub
Sub reload()
 If skipselected2 Then Exit Sub
 If glistN.listcount = 0 Then Exit Sub
 skipselected2 = True
 disablemenu
Dim anystr As String

schedule = Mydir.List(glistN.listindex)
Set myDir2 = New recDir ' we get a new...the old is reference by myDir
myDir2.IncludedFolders = Mydir.IncludedFolders
myDir2.Nofiles = Mydir.Nofiles
myDir2.SortType = Mydir.SortType
myDir2.TopFolder = Mydir.TopFolder
If recnow3checked Then myDir2.LevelStop = 3 Else myDir2.LevelStop = 0

If Mydir.path = "" Then
anystr = myDir2.Dir2("", filetypes, recnowchecked)
Else
''''If GetPath <> Mydir.path Then Stop
anystr = myDir2.Dir2(GetPath, filetypes, recnowchecked)
End If


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
Private Sub myDir2_DirFinished()
Dim findit As Long, A$(), b$

Set Mydir = myDir2
'glistN.clear  no clear please

fakeitemcount = Mydir.listcount
LastItemCount = fakeitemcount
glistN.ShowMe
netdisplay = False
If Mydir.path = "" Then
SetPath "My Computer"
netdisplay = True
b$ = FindNetworkFoldersNames
While b$ <> ""
Mydir.AddAFile (Chr$(2) + GetStrUntil(vbCrLf, b$, True)), 0
fakeitemcount = fakeitemcount + 1
Wend
Else
SetPath Mydir.path
End If
'If NoHeadLine Then  change mind...is better to inform for the change
RaiseEvent NewHeadline(Mydir.path)

glistN.CalcAndShowBar
glistN.Enabled = True
glistN.ShowMe2
If Mydir.Nofiles Then
If Mydir.path = "" Then
TEXT1 = ""
Else
TEXT1 = SelectFolderButton
End If
Else
TEXT1 = mselectedFile
End If
' so text1 above clear Text1Dir every time we change  the top folder in the list
If schedule <> "" Then
If Left$(schedule, 1) <> Chr$(2) And schedule <> "..." Then

 findit = glistN.find(schedule)
If findit >= 0 Then glistN.listindex = findit Else glistN.listindex = 0
Else
glistN.listindex = 0
End If
schedule = ""
ElseIf TEXT1 <> "" And Not Mydir.Nofiles And Not netdisplay Then
findit = glistN.find(TEXT1)
If findit >= 0 Then
glistN.ShowThis findit + 1
Else
glistN.listindex = 0
If Not SaveDialog Then TEXT1 = ""
End If
Else
glistN.listindex = 0
'Text1 = ""
End If

skipselected2 = False
disablemenu True

glistN.ShowMe
End Sub
Function DriveType(path$) As String
    Select Case GetDriveType(Left(path$, 1))
        Case 0
            DriveType = "Non-Recognized Drive"
        Case 1
            DriveType = "Virtual Drive"
        Case 2
            DriveType = "Removable Drive"
        Case 3
            DriveType = "Fixed Drive"
        Case 4
            DriveType = "Network Drive"
        Case 5
            DriveType = "Cd-Rom Drive"
        Case 6
            DriveType = "Ram-Disk Drive"
        Case Else
            DriveType = "Non-Recognized Drive"
    End Select
End Function
Function DriveSerial(ByVal path$) As Long
' need trailing slash
    Dim Serial As Long, VName As String, FSName As String
    VName = String$(255, Chr$(0))
    FSName = String$(255, Chr$(0))
    GetVolumeInformation path$, VName, 255, Serial, 0, 0, FSName, 255
    DriveSerial = Serial
End Function
Function WeCanWrite(ByVal path$) As Boolean
' need trailing slash
    Dim Serial As Long, VName As String, FSName As String, SystemFlags As Long
    VName = String$(255, Chr$(0))
    FSName = String$(255, Chr$(0))
    GetVolumeInformation path$, VName, 255, 0, 0, SystemFlags, FSName, 255
    WeCanWrite = (SystemFlags And &H80000) = 0
End Function



Public Property Get NostateDir() As Boolean
NostateDir = mNostateDir
End Property

Public Property Let NostateDir(ByVal rhs As Boolean)
glistN.BlockItemcount = Not rhs
mNostateDir = rhs
End Property

Private Sub myDir2_FeedBack(FileName As String)
TEXT1.Info = FileName
End Sub

Public Property Get selectedFile() As String
If Text1Dir <> "" Then ReturnFile = Text1Dir  ' this is with path
If ReturnFile = "" Then ReturnFile = Mydir.path + mselectedFile ' this is without path
selectedFile = ReturnFile

End Property

Public Property Let selectedFile(ByVal rhs As String)
Text1Dir = ""
mselectedFile = rhs
End Property
Public Sub ResetHeightSelector()
Dim oldval As String
topoffsetpixel = glistN.LineTopOffsetPixels
If NostateDir Then
oldval = glistN.HeadLine
glistN.HeadLine = ""
glistN.HeadLine = oldval
Else
glistN.HeadLine = ""
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight
mDoc1.WrapAgain
glistN.HeadlineHeight = mDoc1.DocLines * (onelineheight + 1) + 2
End If
glistN.PrepareToShow
End Sub


Public Property Get PreserveNpixelsHeaderRight() As Long
PreserveNpixelsHeaderRight = mPreserveNpixelsHeaderRight
End Property

Public Property Let PreserveNpixelsHeaderRight(ByVal rhs As Long)
mPreserveNpixelsHeaderRight = rhs
glistN.PreserveNpixelsHeaderRightTwips = rhs * Screen.TwipsPerPixelX
End Property
Public Function GetPath() As String
If hiddenDir <> "" Then
GetPath = hiddenDir
Else
GetPath = mDoc1.TextParagraphOrder(1)
End If
End Function
Sub SetPath(Mdoc1Text As String)
Dim SS$
                 SS$ = UCase(Mdoc1Text)
                    DropLeft "\M2000_USER\", SS$
If SS$ = "" Then
hiddenDir = ""
mDoc1.TextParagraphOrder(1) = Mdoc1Text
Else
hiddenDir = Mdoc1Text
mDoc1.TextParagraphOrder(1) = SS$
End If
End Sub

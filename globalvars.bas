Attribute VB_Name = "globalvars"
' This is for selectors..
Public AskTitle$, AskText$, AskCancel$, AskOk$, AskDIB$, ASKINUSE As Boolean
Public AskInput As Boolean, AskResponse$, AskStrInput$
Public UseAskForMultipleEntry As Boolean
Public BreakMe As Boolean
Public CancelDialog As Boolean
Public SizeDialog As Single, helpSizeDialog As Single
Public textinformCaption As String
Public FileTypesShow As String
Public ReturnFile As String
Public ReturnListOfFiles As String  ' # between
Public Settings As String
Public TopFolder As String
Public AskLastX As Long, AskLastY As Long
Public selectorLastX As Long, selectorLastY As Long
Public FolderOnly As Boolean
Public Const AskCancelGR = "ΑΚΥΡΟ"
Public Const AskOkGR = "ΕΝΤΑΞΕΙ"
Public Const LoadFileCaptionGR = "Φόρτωσε Αρχείο"
Public Const SaveFileCaptionGR = "Σώσε Αρχείο"
Public Const SelectFolderCaptionGR = "Επιλογή Φακέλου"
Public Const SelectFolderButtonGR = "*σύρε δεξιά για επιλογή*"
Public Const FontSelectorGr = "Γραμματοσειρά"
Public Const ColorSelectorGr = "Χρωματολόγιο"
Public Const SetUpGR = "Ρυθμίσεις"
Public Const AskCancelEn = "CANCEL"
Public Const AskOkEn = "OK"
Public Const SetUpEn = "Set Up"
Public Const LoadFileCaptionEn = "Load File"
Public Const SaveFileCaptionEn = "Save File"
Public Const SelectFolderCaptionEn = "Select Folder"
Public Const SelectFolderButtonEn = "*slide right to select*"
Public Const FontSelectorEn = "Font Selector"
Public Const ColorSelectorEn = "Color Selector"
Public SetUp As String
Public LoadFileCaption As String
Public SaveFileCaption As String
Public SelectFolderCaption As String
Public SelectFolderButton As String
Public FontSelector As String
Public ColorSelector As String
Public SaveDialog As Boolean
Public DialogPreview As Boolean, LastWidth As Long, HelpLastWidth As Long, PopUpLastWidth As Long
Public ExpandWidth As Boolean, lastfactor As Single, Helplastfactor As Single, Pouplastfactor As Single
Public NewFolder As Boolean, multifileselection As Boolean
Public FileExist As Boolean
Public UserFileName As String
Private inUse As Boolean
Public ReturnColor As Double
Public ReturnFontName As String
Public ReturnBold As Boolean
Public ReturnItalic As Boolean
Public ReturnCharset As Integer
Public ReturnSize As Single
Public DialogLang As Long
Public Sub DialogSetupLang(lang As Long)
DialogLang = lang
If lang = 0 Then
AskCancel$ = AskCancelGR
AskOk$ = AskOkGR
 LoadFileCaption = LoadFileCaptionGR
 SaveFileCaption = SaveFileCaptionGR
 SelectFolderCaption = SelectFolderCaptionGR
 SelectFolderButton = SelectFolderButtonGR
  FontSelector = FontSelectorGr
ColorSelector = ColorSelectorGr
 SetUp = SetUpGR
Else
AskCancel$ = AskCancelEn
AskOk$ = AskOkEn
 LoadFileCaption = LoadFileCaptionEn
 SaveFileCaption = SaveFileCaptionEn
 SelectFolderCaption = SelectFolderCaptionEn
 SelectFolderButton = SelectFolderButtonEn
  FontSelector = FontSelectorEn
ColorSelector = ColorSelectorEn
 SetUp = SetUpEn
End If
End Sub
Public Function IsSelectorInUse() As Boolean
IsSelectorInUse = inUse
End Function
Public Function OpenColor(bstack As basetask, Thisform As Object, thiscolor As Long) As Boolean
If inUse Then OpenColor = False: Exit Function
inUse = True
ExpandWidth = True
ReturnColor = thiscolor
If Thisform Is Nothing Then
ColorDialog.Show
Else
ColorDialog.Show , Thisform
End If
CancelDialog = False
If Not ColorDialog.Visible Then
    ColorDialog.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
OpenColor = Not CancelDialog
thiscolor = ReturnColor
ExpandWidth = False
inUse = False
End Function
Public Function OpenFont(bstack As basetask, Thisform As Object) As Boolean
If inUse Then OpenFont = False: Exit Function
inUse = True
ExpandWidth = True
If Thisform Is Nothing Then
FontDialog.Show
Else
FontDialog.Show , Thisform
End If
CancelDialog = False
If Not FontDialog.Visible Then
    FontDialog.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
If ReturnFontName <> "" Then OpenFont = Not CancelDialog
ExpandWidth = False
inUse = False
End Function
Public Function OpenImage(bstack As basetask, Thisform As Object, TopDir As String, LastName As String, thattitle As String, TypeList As String) As Boolean
If inUse Then OpenImage = False: Exit Function
inUse = True
' do something with multifiles..
ReturnFile = LastName
If ReturnFile <> "" Then If ExtractPath(LastName) = "" Then ReturnFile = mcd + LastName
SaveDialog = False
FileExist = True
FolderOnly = False
''If TopDir <> "" Then TopFolder = TopDir
If TopDir = "" Then
TopFolder = mcd
ReturnFile = mcd
ElseIf TopDir = "\" Then
TopFolder = ""
ReturnFile = mcd
ElseIf TopDir = "*" Then
TopFolder = ""
ReturnFile = ""

Else
TopFolder = TopDir
End If
ReturnListOfFiles = ""
If TypeList = "" Then FileTypesShow = "BMP|JPG|GIF|WMF|EMF|DIB|ICO|CUR" Else FileTypesShow = TypeList
DialogPreview = True
If thattitle <> "" Then
LoadFileCaption = thattitle
If InStr(Settings, ",expand") = 0 Then
Settings = Settings & ",expand"
End If
End If
If Thisform Is Nothing Then
LoadFile.Show
Else
LoadFile.Show , Thisform
End If
CancelDialog = False
If Not LoadFile.Visible Then
    LoadFile.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
If ReturnListOfFiles <> "" Or ReturnFile <> "" Then OpenImage = Not CancelDialog
inUse = False

' read files
End Function
Public Function OpenDialog(bstack As basetask, Thisform As Object, TopDir As String, LastName As String, thattitle As String, TypeList As String, OpenNew As Boolean, MULTFILES As Boolean) As Boolean
If inUse Then OpenDialog = False: Exit Function
inUse = True
' do something with multifiles..
ReturnFile = LastName
If ReturnFile <> "" Then If ExtractPath(LastName) = "" Then ReturnFile = mcd + LastName
SaveDialog = False
FileExist = OpenNew
FolderOnly = False
' If TopDir <> "" Then TopFolder = TopDir
If TopDir = "" Then
TopFolder = mcd
ReturnFile = mcd
ElseIf TopDir = "\" Then
TopFolder = ""
ReturnFile = mcd
ElseIf TopDir = "*" Then
TopFolder = ""
ReturnFile = ""

Else
TopFolder = TopDir
End If
ReturnListOfFiles = ""
FileTypesShow = TypeList
DialogPreview = False
If thattitle <> "" Then
LoadFileCaption = thattitle
If InStr(Settings, ",expand") = 0 Then
Settings = Settings & ",expand"
End If
End If

If Thisform Is Nothing Then
LoadFile.Show
Else
LoadFile.Show , Thisform
End If
CancelDialog = False
If Not LoadFile.Visible Then
    LoadFile.Visible = True
    MyDoEvents
    End If
Hook3 LoadFile.hwnd, Nothing
WaitDialog bstack

Set LastGlist3 = Nothing
If ReturnListOfFiles <> "" Or ReturnFile <> "" Then OpenDialog = Not CancelDialog
inUse = False
' read files
End Function
Public Function SaveAsDialog(bstack As basetask, Thisform As Object, LastName As String, TopDir As String, thattitle As String, TypeList As String) As Boolean
If inUse Then SaveAsDialog = False: Exit Function
inUse = True
DialogPreview = False
FolderOnly = False
SaveDialog = True
UserFileName = LastName
ReturnFile = ExtractPath(LastName)
FileTypesShow = TypeList
''If TopDir <> "" Then TopFolder = TopDir
If TopDir = "" Then
TopFolder = mcd
ReturnFile = mcd
ElseIf TopDir = "\" Then
TopFolder = ""
ReturnFile = mcd
ElseIf TopDir = "*" Then
TopFolder = ""
ReturnFile = ""

Else
TopFolder = TopDir
End If
If ReturnFile = "" Then ReturnFile = TopDir + LastName
If thattitle <> "" Then
SaveFileCaption = thattitle
If InStr(Settings, ",expand") = 0 Then
Settings = Settings & ",expand"
End If
End If
If Thisform Is Nothing Then
LoadFile.Show
Else
LoadFile.Show , Thisform
End If
 CancelDialog = False
 If Not LoadFile.Visible Then
    LoadFile.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
If ReturnFile <> "" Then SaveAsDialog = Not CancelDialog
inUse = False
End Function
Public Function GetFile(bstack As basetask, thistitle As String, thisfolder As String, onetype As String) As String
Dim thatform As Object
If Form1.Visible Then
Set thatform = Form1
Else
Set thatform = Nothing
End If
    If OpenDialog(bstack, thatform, thisfolder, "", thistitle, onetype, False, False) Then
    GetFile = ReturnFile
    End If

End Function

Public Function FolderSelector(bstack As basetask, Thisform As Object, thatfolder As String, TopDir As String, thattitle As String, newflag As Boolean) As Boolean
If inUse Then FolderSelector = False: Exit Function
inUse = True
DialogPreview = False
ReturnFile = thatfolder
SaveDialog = False
NewFolder = newflag
FolderOnly = True
FileExist = True
If thattitle <> "" Then
SelectFolderCaption = thattitle
If InStr(Settings, ",expand") = 0 Then
Settings = Settings & ",expand"
End If
End If
If NewFolder Then FileExist = False
If TopDir = "" Then
TopFolder = mcd
ReturnFile = mcd
ElseIf TopDir = "\" Then
TopFolder = ""
ReturnFile = mcd
ElseIf TopDir = "*" Then
TopFolder = ""
ReturnFile = ""

Else
TopFolder = TopDir
End If
If Thisform Is Nothing Then
LoadFile.Show
Else
LoadFile.Show , Thisform
End If
CancelDialog = False
If Not LoadFile.Visible Then
    LoadFile.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
If ReturnFile <> "" Then FolderSelector = Not CancelDialog
inUse = False
End Function
Sub ReleaseSelector()
inUse = False
End Sub
Function ConCat(ParamArray aa() As Variant) As String
Dim all$, i As Long
For i = 0 To UBound(aa)
    all$ = all$ & CStr(aa(i))
Next i
ConCat = all$
End Function

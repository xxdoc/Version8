VERSION 5.00
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form5"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetLocaleInfo Lib "KERNEL32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&)
Private Const DWL_ANYTHREAD& = 0
Const LOCALE_ILANGUAGE = 1
Private Declare Function SetErrorMode Lib "KERNEL32" ( _
   ByVal wMode As Long) As Long

Private Const SEM_NOGPFAULTERRORBOX = &H2&
Private Sub Form_Activate()
'If Form1.WindowState <> vbMinimized And Form1.Visible Then Form1.ActiveControl.SetFocus
If Form1.Visible Then Form1.SetFocus
End Sub

Private Sub Form_Load()
Set LastGlist = Nothing
form5iamloaded = True
If Not s_complete Then
Me.Move -10000
Form1.Hide

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set LastGlist = Nothing
Set LastGlist2 = Nothing
'Set DisStak.Owner = Nothing
form5iamloaded = False '
'RemoveFont GetCurDir(True) & "TT6492M_.TTF"
MediaPlayer1.closeMovie
  DisableMidi
  TaskMaster.Dispose
  Set TaskMaster = Nothing
Dim x As Form
For Each x In Forms
''MsgBox X.name
If x.name <> Me.name Then Unload x
Next
Set x = Nothing
If m_bInIDE Then Exit Sub
SetErrorMode SEM_NOGPFAULTERRORBOX
'End
''If App.UnattendedApp Then End
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
INK$ = INK$ & GetKeY(KeyAscii)
End Sub
Public Sub RestoreSizePos()
' calling from form1
Me.Move Form1.Left, Form1.top, Form1.Width, Form1.Height
End Sub
 Function GetKeY(ascii As Integer) As String
    Dim Buffer As String, ret As Long
    Buffer = String$(514, 0)
    Dim r&, k&
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      r = val("&H" & Right(Hex(r), 4))
    ret = GetLocaleInfo(r, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    If ret > 0 Then
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, CLng(val("&h" + Left$(Buffer, ret - 1))))))
    Else
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, 1033)))
    End If
End Function

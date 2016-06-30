Attribute VB_Name = "Module1"
' M2000 starter
' We have to give some stack space
Private Declare Sub DisableProcessWindowsGhosting Lib "user32" ()
Dim m As Object
Private Declare Function GetCommandLineW Lib "KERNEL32" () As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal psString As Long) As Long
Private Const SEM_NOGPFAULTERRORBOX = &H2&
Public m_bInIDE As Boolean
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Function SetErrorMode Lib "KERNEL32" ( _
   ByVal wMode As Long) As Long

Public Function commandW() As String
Static mm$
If mm$ <> "" Then commandW = mm$: Exit Function
If m_bInIDE Then
mm$ = Command
Else
Dim Ptr As Long: Ptr = GetCommandLineW
    If Ptr Then
        PutMem4 VarPtr(commandW), SysAllocStringLen(Ptr, lstrlenW(Ptr))
     If AscW(commandW) = 34 Then
       commandW = Mid$(commandW, InStr(commandW, """ ") + 2)
       Else
            commandW = Mid$(commandW, InStr(commandW, " ") + 1)
        End If
    End If
    End If
    If mm$ = "" And Command <> "" Then commandW = Command Else commandW = mm$
End Function
Sub Main()
DisableProcessWindowsGhosting
Dim a$
On Error Resume Next
With New cFIE
.FEATURE_BROWSER_EMULATION = .InstalledVersion
End With
Set m = CreateObject("M2000.callback")
'Dim m As New M2000.callback
If Err Then
    MsgBox "Install M2000.dll first", vbCritical
Exit Sub
End If
Debug.Assert (InIDECheck = True)
m.Run "start"
a$ = commandW
If Trim$(a$) = "-h" Or Trim$(a$) = "/?" Then frmAbout.Show: Exit Sub
If m.Status = 0 Then
m.Cli a$, ">"
'm.Reset
End If
m.ShowGui = False
Set m = Nothing
If m_bInIDE Then Exit Sub
SetErrorMode SEM_NOGPFAULTERRORBOX
End Sub
Public Function InIDECheck() As Boolean
    m_bInIDE = True
    InIDECheck = True
End Function

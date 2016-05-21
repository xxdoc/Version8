Attribute VB_Name = "Module1"
' M2000 starter
' We have to give some stack space
Dim m As New M2000.callback
Private Declare Function GetCommandLineW Lib "KERNEL32" () As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal psString As Long) As Long

Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long

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
Dim a$
Debug.Assert (InIDECheck = True)
m.Run "start"
a$ = commandW
If Trim$(a$) = "-h" Or Trim$(a$) = "/?" Then frmAbout.Show: Exit Sub
If m.Status = 0 Then
m.Cli a$, ">"
'm.Reset
End If
m.ShowGui = False
Debug.Print "ok"
End Sub
Public Function InIDECheck() As Boolean
    m_bInIDE = True
    InIDECheck = True
End Function

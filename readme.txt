Version 8.1
Revision 19
Now we can use one object in an application


Dim m As New M2000.callback
Sub makeit()
m.Run "start"
End Sub
Sub cli()
If m.Status = 0 Then
    m.cli "c:\note1.gsb"
    m.Reset
End If
Debug.Print "ok"
End Sub
Sub cli2()
If m.Status = 0 Then
    m.cli "", ">"
    m.Reset
End If
m.ShowGui = False
Debug.Print "ok"
End Sub

Sub look()
Debug.Print m.Status
Dim a$
If m.Status < 0 Then Exit Sub
    m.ExecuteStatement "Start"
again:
    If m.Status = 0 Then
        m.Run "show :repeat { clear cmd$ :print $(0), {M2000>}; : line input cmd$ : print" + vbCrLf + " inline cmd$" + vbCrLf + "} always", False
    End If
    If Abs(m.Status) = 1 Then
        a$ = m.ErrorGr
        m.Reset
        m.Run "Print : Print {" + a$ + "}"
        GoTo again
    End If
    Form1.Caption = m.Eval(CStr(Timer))
    m.ShowGui = False
    m.Reset
Debug.Print "ok"
End Sub

Sub look2()
Debug.Print m.Status
If m.Status = 0 Then
    Form1.Caption = m.Eval("100*500")  '
    m.Run "start"
    m.Run "Load c:\note1.gsb"
    m.ShowGui = False
    m.Reset
    Debug.Print "ok"
End If
End Sub
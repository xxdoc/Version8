Attribute VB_Name = "databaseX"
'This is the new version for ADO.
Option Explicit
Dim AABB As Long
Dim conCollection As Collection
Dim Init As Boolean
'  to be changed User and UserPassword
Public JetPrefixUser As String
Public JetPostfixUser As String
Public JetPrefix As String
Public JetPostfix As String
'old Microsoft.Jet.OLEDB.4.0
' Microsoft.ACE.OLEDB.12.0
Public Const JetPrefixHelp = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Public Const JetPostfixHelp = ";Jet OLEDB:Database Password=100101;"
Public DBUser As String ' '= "" ' "admin"  ' or ""
Public DBUserPassword   As String ''= ""
Public extDBUser As String ' '= "" ' "admin"  ' or ""
Public extDBUserPassword   As String ''= ""
Public DBtype As String ' can be mdb or something else
Public Const DBtypeHelp = ".mdb" 'allways help has an mdb as type"
Const DBSecurityOFF = ";Persist Security Info=False"

Private Declare Function MoveFileW Lib "kernel32.dll" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Public Sub KillFile(sFilenName As String)
DeleteFileW StrPtr(sFilenName)
End Sub

Public Function MoveFile(pOldPath As String, pNewPath As String)

    MoveFileW StrPtr(pOldPath), StrPtr(pNewPath)
    
End Function
Public Function isdir(f$) As Boolean
On Error Resume Next
Dim MM As New recDir
Dim lookfirst As Boolean
Dim pad$
If f$ = "" Then Exit Function
If f$ = "." Then f$ = mcd
If InStr(f$, "\..") > 0 Or f$ = ".." Or Left$(f$, 3) = "..\" Then
If Right$(f$, 1) <> "\" Then
pad$ = ExtractPath(f$ & "\", True, True)
Else
pad$ = ExtractPath(f$, True, True)
End If
If pad$ = "" Then
If Right$(f$, 1) <> "\" Then
pad$ = ExtractPath(mcd + f$ & "\", True)
Else
pad$ = ExtractPath(mcd + f$, True)
End If
End If
lookfirst = MM.isdir(pad$)
If lookfirst Then f$ = pad$
Else
f$ = mylcasefILE(f$)
lookfirst = MM.isdir(f$)
If Not lookfirst Then

pad$ = mcd + f$

lookfirst = MM.isdir(pad$)
If lookfirst Then f$ = pad$

End If
End If
isdir = lookfirst
End Function
Public Sub fHelp(bstack As basetask, d$, Optional Eng As Boolean = False)
Dim sql$, b$, p$, c$, gp$, r As Double, bb As Long, I As Long
Dim cd As String, doriginal$
On Error GoTo E5
'ON ERROR GoTo 0
If HelpLastWidth > ScrX() Then HelpLastWidth = -1
doriginal$ = d$

If d$ <> "" Then If Right$(d$, 1) = "(" Then d$ = d$ + ")"
If d$ = "" Or d$ = "F12" Then
d$ = ""
If Right$(d$, 1) = "(" Then d$ = d$ + ")"
p$ = subHash.Show

While ISSTRINGA(p$, c$)
IsLabelA "", c$, b$
If Right$(b$, 1) = "(" Then b$ = b$ + ")"
If gp$ <> "" Then gp$ = b$ + ", " + gp$ Else gp$ = b$
Wend
If vH_title$ <> "" Then b$ = "<| " & vH_title$ & vbCrLf & vbCrLf Else b$ = ""
If Eng Then
        sHelp "User Modules/Functions [F12]", b$ & gp$, (ScrX() - 1) * 3 / 5, (ScrY() - 1) * 4 / 7
Else
        sHelp "Τμήματα/Συναρτήσεις Χρήστη [F12]", b$ & gp$, (ScrX() - 1) * 3 / 5, (ScrY() - 1) * 4 / 7
End If
vHelp Not Form4.Visible
Exit Sub
ElseIf GetSub(d$, I) Or d$ = HERE$ Then
If d$ = HERE$ Then I = bstack.OriginalCode
If vH_title$ <> "" Then
b$ = "<| " & vH_title$ & vbCrLf & vbCrLf
Else
If Eng Then
b$ = "<| " & "User Modules/Functions [F12]" & vbCrLf & vbCrLf
Else
b$ = "<| " & "Τμήματα/Συναρτήσεις Χρήστη [F12]" & vbCrLf & vbCrLf
End If
End If
If Right$(d$, 1) = ")" Then

If Eng Then c$ = "[Function]" Else c$ = "[Συνάρτηση]"
Else
If Eng Then c$ = "[Module]" Else c$ = "[Function]"
End If

Dim ss$
    ss$ = GetNextLine((SBcode(I)))
    If Left$(ss$, 10) = "'11001EDIT" Then
    
    ss$ = Mid$(SBcode(I), Len(ss$) + 3)
    Else
     ss$ = SBcode(I)
     End If
        sHelp d$, c$ + "  " & b$ & ss$, (ScrX() - 1) * 3 / 5, (ScrY() - 1) * 4 / 7
    
        vHelp Not Form4.Visible
Exit Sub
End If




JetPrefix = JetPrefixHelp
JetPostfix = JetPostfixHelp
DBUser = ""
DBUserPassword = ""

cd = App.path
AddDirSep cd

p$ = Chr(34)
c$ = ","
d$ = doriginal$
If Asc(d$) < 128 Then
sql$ = "SELECT * FROM [COMMANDS] WHERE ENGLISH >= '" & UCase(d$) & "'"
Else
sql$ = "SELECT * FROM [COMMANDS] WHERE DESCRIPTION >= '" & myUcase(d$, True) & "'"
End If
b$ = mylcasefILE(cd & "help2000")
getrow bstack, p$ & b$ & p$ & c$ & p$ & sql$ & p$ & ",1," & p$ & p$ & c$ & p$ & p$, False, , , True
sql$ = p$ & b$ & p$ & c$ & p$ & "GROUP" & p$
If bstack.IsNumber(r) Then
If bstack.IsString(gp$) Then
If bstack.IsString(b$) Then
If bstack.IsString(p$) Then
If bstack.IsNumber(r) Then
getrow bstack, sql$ & "," & CStr(1) & "," & Chr(34) & "GROUPNUM" & Chr(34) & "," & Str$(r), False, , , True
If bstack.IsNumber(r) Then
If bstack.IsNumber(r) Then
If bstack.IsString(c$) Then
' nothing
        If Eng Then gp$ = p$
        If vH_title$ <> "" Then
            If vH_title$ = gp$ And Form4.Visible = True Then GoTo E5
        End If
        bb = InStr(b$, "__<ENG>__")
        If bb > 0 Then
            If Eng Then
            c$ = "[" & Trim$(Mid$(c$, InStr(c$, ",") + 1)) & "]"
                b$ = Mid$(b$, bb + 11)
            Else
            c$ = "[" & Mid$(c$, 1, InStr(c$, ",") - 1) & "]"
                b$ = Left$(b$, bb - 1)
            End If
        End If
        If vH_title$ <> "" Then b$ = "<| " & vH_title$ & vbCrLf & vbCrLf & b$ Else b$ = vbCrLf & b$
        
        sHelp gp$, c$ & "  " & b$, (ScrX() - 1) * 3 / 5, (ScrY() - 1) * 4 / 7
    
        vHelp Not Form4.Visible
        End If
    
    End If
End If

End If
End If
End If
End If
End If
E5:
JetPrefix = JetPrefixUser
JetPostfix = JetPostfixUser
DBUser = extDBUser
DBUserPassword = extDBUserPassword
Err.clear
End Sub
Public Function inames(I As Long, lang As Long) As String
If (I And &H3) <> 1 Then
Select Case lang
Case 1

inames = "DESCENDING"
Case Else
inames = "ΦΘΙΝΟΥΣΑ"
End Select
Else
Select Case lang
Case 1
inames = "ASCENDING"
Case Else
inames = "ΑΥΞΟΥΣΑ"
End Select

End If

End Function
Public Function fnames(I As Long, lang As Long) As String
Select Case I
Case 1
    Select Case lang
    Case 1
    fnames = "BOOLEAN"
    Case Else
     fnames = "ΛΟΓΙΚΟΣ"
    End Select
    Exit Function
Case 2
    Select Case lang
    Case 1
    fnames = "BYTE"
    Case Else
     fnames = "ΨΗΦΙΟ"
    End Select
   Exit Function

Case 3
        Select Case lang
    Case 1
    fnames = "INTEGER"
    Case Else
     fnames = "ΑΚΕΡΑΙΟΣ"
    End Select
   Exit Function
Case 4
        Select Case lang
    Case 1
    fnames = "LONG"
    Case Else
     fnames = "ΜΑΚΡΥΣ"
    End Select
   Exit Function
 
Case 5
        Select Case lang
    Case 1
    fnames = "CURRENCY"
    Case Else
     fnames = "ΛΟΓΙΣΤΙΚΟ"
    End Select
   Exit Function

Case 6
    Select Case lang
    Case 1
    fnames = "SINGLE"
    Case Else
     fnames = "ΑΠΛΟΣ"
    End Select
   Exit Function

Case 7
    Select Case lang
    Case 1
    fnames = "DOUBLE"
    Case Else
     fnames = "ΔΙΠΛΟΣ"
    End Select
   Exit Function
Case 8
    Select Case lang
    Case 1
    fnames = "DATEFIELD"
    Case Else
     fnames = "ΗΜΕΡΟΜΗΝΙΑ"
    End Select
   Exit Function
Case 9 '.....................ole 205
    Select Case lang
    Case 1
    fnames = "BINARY"
    Case Else
     fnames = "ΔΥΑΔΙΚΟ"
    End Select
   Exit Function
Case 10 '..........................................202
    Select Case lang
    Case 1
    fnames = "TEXT"
    Case Else
     fnames = "ΚΕΙΜΕΝΟ"
    End Select
   Exit Function
Case 11 '...........205
    fnames = "OLE"
    Exit Function
Case 12 '...........................202
    Select Case lang
    Case 1
    fnames = "MEMO"
    Case Else
     fnames = "ΥΠΟΜΝΗΜΑ"
    End Select
Case Else
fnames = "?"
End Select
End Function

Public Sub NewBase(bstackstr As basetask, r$)
Dim base As String, othersettings As String
 If Not IsStrExp(bstackstr, r$, base) Then Exit Sub ' make it to give error
If FastSymbol(r$, ",") Then
If Not IsStrExp(bstackstr, r$, othersettings) Then Exit Sub  ' make it to give error
End If
 On Error Resume Next
 If Left$(base, 1) = "(" Or JetPostfix = ";" Then Exit Sub ' we can't create in ODBC
If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
If ExtractType(base) = "" Then base = base & ".mdb"

If CFname((base)) <> "" Then
 If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
' check to see if is our
RemoveOneConn base
If CheckMine(base) Then
KillFile base
Err.clear

Else
MyEr "Can 't delete the Base", "Δεν μπορώ να διαγράψω τη βάση"

Exit Sub
End If
End If

 CreateObject("ADOX.Catalog").Create (JetPrefix & base & JetPostfix & othersettings)  'create a new, empty *.mdb-File

End Sub

Public Sub TABLENAMES(bstackstr As basetask, r$, lang As Long)
Dim base As String, tablename As String, scope As Long, cnt As Long, srl As Long, stac1 As New mStiva
Dim myBase  ' variant
scope = 1
If Not IsStrExp(bstackstr, r$, base) Then Exit Sub
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, tablename) Then
scope = 2

End If
End If


    Dim vindx As Boolean

    On Error Resume Next
            If Left$(base, 1) = "(" Or JetPostfix = ";" Then
        'skip this
        Else
            If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
            If ExtractType(base) = "" Then base = base & ".mdb"
            If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
        End If
    If True Then
        On Error Resume Next
        If Not getone(base, myBase) Then
            Set myBase = CreateObject("ADODB.Connection")
            If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                srl = DriveSerial(Left$(base, 3))
                If srl = 0 And Not GetDosPath(base) = "" Then
                    If lang = 0 Then
                        If Not ask("Βάλε το CD/Δισκέτα με το αρχείο " & ExtractName(base)) = vbCancel Then Exit Sub
                    Else
                        If Not ask("Put CD/Disk with file " & ExtractName(base)) = vbCancel Then Exit Sub
                    End If
                End If
                If myBase = "" Then
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.Open JetPrefix & JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Sub
                        End If
                    Else
                        myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF      'open the Connection
                    End If
                End If
                If Err.Number > 0 Then
                    Do While srl <> DriveSerial(Left$(base, 3))
                        If lang = 0 Then
                            If ask("Βάλε το CD/Δισκέτα με αριθμό σειράς " & CStr(srl) & " στον οδηγό " & Left$(base, 1)) = vbCancel Then Exit Do
                        Else
                            If ask("Put CD/Disk with serial number " & CStr(srl) & " in drive " & Left$(base, 1)) = vbCancel Then Exit Do
                        End If
                    Loop
                    If srl = DriveSerial(Left$(base, 3)) Then
                        Err.clear
                        If myBase = "" Then myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBSecurityOFF       'open the Connection
                    End If
                End If
            Else
                If myBase = "" Then
                ' check if we have ODBC
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.Open JetPrefix & JetPostfix
                        If Err.Number Then
                            MyEr Err.Description, Err.Description
                            Exit Sub
                        End If
                    Else
                        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                    End If
                End If
        End If
        If Err.Number > 0 Then GoTo g102
        PushOne base, myBase
    End If
  Dim cat, TBL, rs
     Dim I As Long, j As Long, k As Long, KB As Boolean
  
           Set rs = CreateObject("ADODB.Recordset")
        Set TBL = CreateObject("ADOX.TABLE")
           Set cat = CreateObject("ADOX.Catalog")
           Set cat.activeconnection = myBase
           If cat.activeconnection.errors.Count > 0 Then
           MyEr "Can't connect to Base", "Δεν μπορώ να συνδεθώ με τη βάση"
           Exit Sub
           End If
        If cat.TABLES.Count > 0 Then
        For Each TBL In cat.TABLES
        
        If TBL.Type = "TABLE" Then
        vindx = False
        KB = False
        If scope <> 2 Then
        
        cnt = cnt + 1
                            stac1.DataStr TBL.name
                       If TBL.indexes.Count > 0 Then
                                         For j = 0 To TBL.indexes.Count - 1
                                                   With TBL.indexes(j)
                                                   If (.Unique = False) And (.indexnulls = 0) Then
                                                        KB = True
                                                  Exit For
             '
                                                       End If
                                                   End With
                                                Next j
                                              If KB Then
                    
                                                     stac1.DataVal CDbl(1)
                                                     
                                                Else
                                                    stac1.DataVal CDbl(0)
                                                End If
                                               
                                           
                                            Else
                                            stac1.DataVal CDbl(0)
                                        End If
         ElseIf tablename = TBL.name Then
         cnt = 1
                     rs.Open "Select * From [" & TBL.name & "] ;", myBase, 3, 4 'adOpenStatic, adLockBatchOptimistic
                                         stac1.Flush
                                        stac1.DataVal CDbl(rs.FIELDS.Count)
                                        If TBL.indexes.Count > 0 Then
                                         For j = 0 To TBL.indexes.Count - 1
                                                   With TBL.indexes(j)
                                                   If (.Unique = False) And (.indexnulls = 0) Then
                                                   vindx = True
                                                   Exit For
                                                       End If
                                                   End With
                                                Next j
                                                If vindx Then
                                                
                                                     stac1.DataVal CDbl(1)
                                                Else
                                                    stac1.DataVal CDbl(0)
                                                End If
                                            Else
                                            stac1.DataVal CDbl(0)
                                        End If
                     For I = 0 To rs.FIELDS.Count - 1
                     With rs.FIELDS(I)
                             stac1.DataStr .name
                             If .Type = 203 And .DEFINEDSIZE >= 536870910# Then
                             
                                         If lang = 1 Then
                                        stac1.DataStr "ΜΕΜΟ"
                                        Else
                                        stac1.DataStr "ΥΠΟΜΝΗΜΑ"
                                        End If
                                        
                                        stac1.DataVal CDbl(0)
                            
                             ElseIf .Type = 205 Then
                                       
                                            stac1.DataStr "OLE"
                                       
                                       
                                            stac1.DataVal CDbl(0)
                                     ElseIf .Type = 202 And .DEFINEDSIZE <> 536870910# Then
                                            If lang = 1 Then
                                            stac1.DataStr "TEXT"
                                            Else
                                            stac1.DataStr "ΚΕΙΜΕΝΟ"
                                            End If
                                            stac1.DataVal CDbl(.DEFINEDSIZE)
                                    
                             Else
                                        stac1.DataStr ftype(.Type, lang)
                                        stac1.DataVal CDbl(.DEFINEDSIZE)
                             
                             End If
                     End With
                     Next I
                     rs.Close
                     If vindx Then
                    If TBL.indexes.Count > 0 Then
                             For j = 0 To TBL.indexes.Count - 1
                          With TBL.indexes(j)
                          If (.Unique = False) And (.indexnulls = 0) Then
                          stac1.DataVal CDbl(.Columns.Count)
                          For k = 0 To .Columns.Count - 1
                            stac1.DataStr .Columns(k).name
                             stac1.DataStr inames(.Columns(k).sortorder, lang)
                          Next k
                             Exit For
                             
                             End If
                          End With
                       Next j
                    End If
                     End If
             End If
             End If
            
                                     
                         
               Next TBL
    End If
    If scope = 1 Then
    stac1.PushVal CDbl(cnt)
    Else
    If cnt = 0 Then
     MyEr "No such TABLE in DATABASE", "Δεν υπάρχει τέτοιο αρχείο στη βάση δεδομένων"
    End If
    End If
     bstackstr.soros.MergeTop stac1
     Else
     RemoveOneConn myBase
     MyEr "No such DATABASE", "Δεν υπάρχει τέτοια βάση δεδομένων"
    End If
g102:
End Sub

Public Sub append_table(bstackstr As basetask, base As String, r$, ED As Boolean, Optional lang As Long = -1)
Dim table$, I&, par$, ok As Boolean, t As Double, j&
Dim gindex As Long
ok = False

If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, table$) Then
ok = True
End If
End If
If lang <> -1 Then If IsLabelSymbolNew(r$, "ΣΤΟ", "TO", lang) Then If IsExp(bstackstr, r$, t) Then gindex = CLng(t) Else SyntaxError
Dim Id$
  If InStr(UCase(Trim$(table$)) + " ", "SELECT") = 1 Then
Id$ = table$
Else
Id$ = "SELECT * FROM [" + table$ + "]"
End If


If Not ok Then Exit Sub


If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = "" Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
          On Error Resume Next
          Dim myBase
          
               If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Sub
                Else
                If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                    myBase.Open JetPrefix & JetPostfix
                    If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Sub
                    End If
                Else
                    myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                End If
                End If
                PushOne base, myBase
            End If
           Err.clear
         
         '  If Err.Number > 0 Then GoTo thh
           
           
         '  Set rec = myBase.OpenRecordset(table$, dbOpenDynaset)
          Dim rec, ll$
          
           Set rec = CreateObject("ADODB.Recordset")
            Err.clear
           rec.Open Id$, myBase, 3, 4 'adOpenStatic, adLockBatchOptimistic

 If Err.Number <> 0 Then
ll$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = ll$
 PushOne base, myBase
 Err.clear
rec.Open Id$, myBase, 3, 4
If Err.Number Then
MyEr Err.Description & " " & Id$, Err.Description & " " & Id$
Exit Sub
End If
End If
   
   
If ED Then
If gindex > 0 Then
Err.clear
    rec.MoveLast
    rec.MoveFirst
    rec.AbsolutePosition = gindex '  - 1
    If Err.Number <> 0 Then
    MyEr "Wrong index for table " & table$, "Λάθος δείκτης για αρχείο " & table$
    End If
Else
    rec.MoveLast
End If
' rec.Edit  no need for undo
Else
rec.AddNew
End If
I& = 0
While FastSymbol(r$, ",")
If ED Then
    While FastSymbol(r$, ",")
    I& = I& + 1
    Wend
End If
If IsStrExp(bstackstr, r$, par$) Then
    rec.FIELDS(I&) = par$
ElseIf IsExp(bstackstr, r$, t) Then
    rec.FIELDS(I&) = CStr(t)   '??? convert to a standard format
End If

I& = I& + 1
Wend
Err.clear
rec.UpdateBatch  ' update be an updatebatch
If Err.Number > 0 Then
MyEr "Can't append " & Err.Description, "Αδυναμία προσθήκης:" & Err.Description
End If

End Sub
Public Sub getrow(bstackstr As basetask, r$, Optional ERL As Boolean = True, Optional search$ = " = ", Optional lang As Long = 0, Optional IamHelpFile As Boolean = False)

Dim base As String, table$, from As Long, first$, Second$, ok As Boolean, fr As Double, stac1$, p As Double, I&
ok = False
If IsStrExp(bstackstr, r$, base) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, table$) Then
If FastSymbol(r$, ",") Then
If IsExp(bstackstr, r$, fr) Then
from = CLng(fr)
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, first$) Then
If FastSymbol(r$, ",") Then
If search$ = "" Then
    If IsStrExp(bstackstr, r$, search$) Then
    search$ = " " & search$ & " "
        If FastSymbol(r$, ",") Then
                If IsExp(bstackstr, r$, p) Then
                Second$ = search$ & Str$(p)
                ok = True
            ElseIf IsStrExp(bstackstr, r$, Second$) Then
            If InStr(Second$, "'") > 0 Then
                Second$ = search$ & Chr(34) & Second$ & Chr(34)
            Else
                Second$ = search$ & "'" & Second$ & "'"
                End If
                ok = True
            End If
        End If
 
        End If
    Else
     If IsExp(bstackstr, r$, p) Then
            Second$ = search$ & Str$(p)
            ok = True
            ElseIf IsStrExp(bstackstr, r$, Second$) Then
                      If InStr(Second$, "'") > 0 Then
                Second$ = search$ & Chr(34) & Second$ & Chr(34)
            Else
                Second$ = search$ & "'" & Second$ & "'"
                End If
            ok = True
        End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
'Dim wrkDefault As Workspace,
Dim ii As Long
Dim myBase  ' as variant


Dim rec   '  as variant  too  - As Recordset
Dim srl As Long
On Error Resume Next
' new addition to handle ODBC
' base=""
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this

Else
If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
If ExtractType(base) = "" Then base = base & ".mdb"
If Not IamHelpFile Then If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If

g05:
Err.clear
   On Error Resume Next
Dim Id$
   
      If first$ = "" Then
If InStr(UCase(Trim$(table$)) + " ", "SELECT") = 1 Then
Id$ = table$
Else
Id$ = "SELECT * FROM [" + table$ + "]"
  End If
   Else
Id$ = "SELECT * FROM [" & table$ & "] WHERE [" & first$ & "] " & Second$
 End If

   If Not getone(base, myBase) Then
   
      Set myBase = CreateObject("ADODB.Connection")
   
      
    If DriveType(Left$(base, 3)) = "Cd-Rom" Then
        srl = DriveSerial(Left$(base, 3))
        If srl = 0 And Not GetDosPath(base) = "" Then
                If lang = 0 Then
                    If Not ask("Βάλε το CD/Δισκέτα με το αρχείο " & ExtractName(base)) = vbCancel Then Exit Sub
                Else
                    If Not ask("Put CD/Disk with file " & ExtractName(base)) = vbCancel Then Exit Sub
                End If
         End If

 
 '  If mybase = "" Then ' mybase.Mode = adShareDenyWrite
   If myBase = "" Then myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection

            If Err.Number > 0 Then
            
            Do While srl <> DriveSerial(Left$(base, 3))
                If lang = 0 Then
                If ask("Βάλε το CD/Δισκέτα με αριθμό σειράς " & CStr(srl) & " στον οδηγό " & Left$(base, 1)) = vbCancel Then Exit Do
                Else
                If ask("Put CD/Disk with serial number " & CStr(srl) & " in drive " & Left$(base, 1)) = vbCancel Then Exit Do
                End If
            Loop
            If srl = DriveSerial(Left$(base, 3)) Then
            Err.clear
        If myBase = "" Then myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBSecurityOFF      'open the Connection
        
            End If
        
        End If
    Else
'     myBase.Open JetPrefix & """" & GetDosPath(BASE) & """" & ";Jet OLEDB:Database Password=100101;User Id=" & DBUser  & ";Password=" & DBUserPassword & ";" &  DBSecurityOFF  'open the Connection
 If myBase = "" Then
 If Left$(base, 1) = "(" Or JetPostfix = ";" Then
 myBase.Open JetPrefix & JetPostfix
 If Err.Number Then
 MyEr Err.Description, Err.Description
 Exit Sub
 End If
 Else
 myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
 End If
 End If


    End If

   If Err.Number > 0 Then GoTo g10
   
      PushOne base, myBase
      
      End If

Dim ll$
   Set rec = CreateObject("ADODB.Recordset")
 Err.clear
  rec.Open Id$, myBase, 3, 4
If Err.Number <> 0 Then
ll$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = ll$
 PushOne base, myBase
 Err.clear
rec.Open Id$, myBase, 3, 4
If Err.Number Then
MyEr Err.Description & " " & Id$, Err.Description & " " & Id$
Exit Sub
End If
End If

   

   
  If rec.EOF Then
   ' stack$(BASESTACK) = " 0" & stack$(BASESTACK)
   bstackstr.soros.PushVal CDbl(0)
   rec.Close
  myBase.Close
    
    Exit Sub
  End If
  rec.MoveLast
  ii = rec.RecordCount

If ii <> 0 Then
If from >= 0 Then
  rec.MoveFirst
    If ii >= from Then
  rec.Move from - 1
  End If
End If
    For I& = rec.FIELDS.Count - 1 To 0 Step -1

   Select Case rec.FIELDS(I&).Type
Case 1, 2, 3, 4, 5, 6, 7

 If IsNull(rec.FIELDS(I&)) Then
        bstackstr.soros.PushUndefine          '.PushStr "0"
    Else
        bstackstr.soros.PushVal CDbl(rec.FIELDS(I&))
    
End If


Case 130, 8, 203, 202
If IsNull(rec.FIELDS(I&)) Then
    
     bstackstr.soros.PushStr ""
 Else
  
   bstackstr.soros.PushStr CStr(rec.FIELDS(I&))
  End If
Case 11, 12 ' this is the binary field so we can save unicode there
   Case Else
'
   bstackstr.soros.PushStr "?"
 End Select
   Next I&
   End If
   
   'stack$(BaseSTACK) = " " & Trim$(Str$(II)) + stack$(BaseSTACK)
   bstackstr.soros.PushVal CDbl(ii)


Exit Sub
g10:
If ERL Then
If lang = 0 Then
If ask("Το ερώτημα SQL δεν μπορεί να ολοκληρωθεί" & vbCrLf & table$, True) = vbRetry Then GoTo g05
Else
If ask("SQL can't complete" & vbCrLf & table$) = vbRetry Then GoTo g05
End If
Err.clear
MyErMacro r$, "Can't read a database table :" & table$, "Δεν μπορώ να διαβάσω πίνακα :" & table$
End If
On Error Resume Next


End Sub

Public Sub getnames(bstackstr As basetask, r$, bv As Object, lang)
Dim base As String, table$, from As Long, many As Long, ok As Boolean, fr As Double, stac1$, I&
ok = False
If IsStrExp(bstackstr, r$, base) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, table$) Then
If FastSymbol(r$, ",") Then
If IsExp(bstackstr, r$, fr) Then
from = CLng(fr)
If FastSymbol(r$, ",") Then
If IsExp(bstackstr, r$, fr) Then
many = CLng(fr)

ok = True
End If
End If
End If
End If
End If
End If
End If
Dim ii As Long
Dim myBase ' variant
Dim rec
Dim srl As Long
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = "" Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
Dim Id$
  If InStr(UCase(Trim$(table$)) + " ", "SELECT") = 1 Then
Id$ = table$
Else
Id$ = "SELECT * FROM [" + table$ + "]"
End If

     If Not getone(base, myBase) Then
   
      Set myBase = CreateObject("ADODB.Connection")
   
   
   If DriveType(Left$(base, 3)) = "Cd-Rom" Then
       srl = DriveSerial(Left$(base, 3))
    If srl = 0 And Not GetDosPath(base) = "" Then
    
       If lang = 0 Then
    If Not ask("Βάλε το CD/Δισκέτα με το αρχείο " & ExtractName(base)) = vbCancel Then Exit Sub
    Else
      If Not ask("Put CD/Disk with file " & ExtractName(base)) = vbCancel Then Exit Sub
    End If
     End If

     myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF    'open the Connection

               If Err.Number > 0 Then
        
            Do While srl <> DriveSerial(Left$(base, 3))
            If lang = 0 Then
            If ask("Βάλε το CD/Δισκέτα με αριθμό σειράς " & CStr(srl) & " στον οδηγό " & Left$(base, 1)) = vbCancel Then Exit Do
            Else
            If ask("Put CD/Disk with serial number " & CStr(srl) & " in drive " & Left$(base, 1)) = vbCancel Then Exit Do
            End If
            Loop
            If srl = DriveSerial(Left$(base, 3)) Then
            Err.clear
   myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBSecurityOFF   'open the Connection
                
            End If
        
        End If
   Else
    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
 myBase.Open JetPrefix & JetPostfix
 If Err.Number Then
 MyEr Err.Description, Err.Descnullription
 Exit Sub
 End If
 Else
  myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
End If
End If
On Error GoTo g101
      PushOne base, myBase
      
      End If
 Dim ll$
   Set rec = CreateObject("ADODB.Recordset")
    Err.clear
     rec.Open Id$, myBase, 3, 4
      If Err.Number <> 0 Then
ll$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = ll$
 PushOne base, myBase
 Err.clear
rec.Open Id$, myBase, 3, 4
If Err.Number Then
MyEr Err.Description & " " & Id$, Err.Description & " " & Id$
Exit Sub
End If
End If


 ' DBEngine.Idle dbRefreshCache

  If rec.EOF Then
   ''''''''''''''''' stack$(BASESTACK) = " 0" & stack$(BASESTACK)
bstackstr.soros.PushVal CDbl(0)
  Exit Sub
 
'    wrkDefault.Close
  End If
  rec.MoveLast
  ii = rec.RecordCount

If ii <> 0 Then
If from >= 0 Then
  rec.MoveFirst
    If ii >= from Then
  rec.Move from - 1
  End If
End If
If many + from - 1 > ii Then many = ii - from + 1
bstackstr.soros.PushVal CDbl(ii)
''''''''''''''''' stack$(BASESTACK) = " " & Trim$(Str$(II)) + stack$(BASESTACK)

    For I& = 1 To many
    bv.additemFast CStr(rec.FIELDS(0))   ' USING gList
    
    If I& < many Then rec.MoveNext
    Next
  End If
rec.Close
'myBase.Close

Exit Sub
g101:
MyErMacro r$, "Can't read a table from database", "Δεν μπορώ να διαβάσω ένα πίνακα βάσης δεδομένων"

'myBase.Close
End Sub
Public Sub CommExecAndTimeOut(bstackstr As basetask, r$)
Dim base As String, com2execute As String, comTimeOut As Double
Dim ok As Boolean
comTimeOut = 30
If IsStrExp(bstackstr, r$, base) Then
    If FastSymbol(r$, ",") Then
        If IsStrExp(bstackstr, r$, com2execute) Then
        ok = True
            If FastSymbol(r$, ",") Then
                If Not IsExp(bstackstr, r$, comTimeOut) Then
                ok = False
                End If
            End If
        End If
    End If
End If
If Not ok Then Exit Sub
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = "" Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If

Dim myBase
    
    On Error Resume Next
       If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't execute command in a CD-ROM", "Δεν μπορώ εκτελέσω εντολή στη βάση δεδομένων σε CD-ROM"
                    Exit Sub
                Else
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.Open JetPrefix & JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Sub
                        End If
                    Else
                        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                    End If
                End If
                PushOne base, myBase
    End If
           Err.clear
           If comTimeOut >= 10 Then myBase.CommandTimeout = CLng(comTimeOut)
           If Err.Number > 0 Then Err.clear: myBase.errors.clear
            myBase.Execute com2execute

If myBase.errors.Count <> 0 Then
MyEr "Can't execute command", "Δεν μπορώ εκτελέσω εντολή"
 myBase.errors.clear
End If

' we have response


End Sub



Public Sub MyOrder(bstackstr As basetask, r$)
Dim base As String, tablename As String, fs As String, I&, o As Double, ok As Boolean
ok = False
If IsStrExp(bstackstr, r$, base) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, tablename) Then
ok = True
End If
End If
End If

If Not ok Then Exit Sub
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = "" Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
    
    Dim myBase
    
    On Error Resume Next
       If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Sub
                Else
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.Open JetPrefix & JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Sub
                        End If
                    Else
                        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                    End If
                 
                End If
                PushOne base, myBase
            End If
           Err.clear
           Dim ll$, mcat, pindex, mtable
           Dim okntable As Boolean
          
            Err.clear
            Set mcat = CreateObject("ADOX.Catalog")
            mcat.activeconnection = myBase

            

        If Err.Number <> 0 Then
ll$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = ll$
 PushOne base, myBase
 Err.clear
            Set mcat = CreateObject("ADOX.Catalog")
            mcat.activeconnection = myBase
            

If Err.Number Then
MyEr Err.Description & " " & tablename, Err.Description & " " & tablename
Exit Sub
End If
End If
error.clear
mcat.TABLES(tablename).indexes("ndx").remove
mcat.TABLES(tablename).indexes.refresh

   If mcat.TABLES.Count > 0 Then
   okntable = True
        For Each mtable In mcat.TABLES
        
        If mtable.Type = "TABLE" Then
        If mtable.name = tablename Then
        okntable = False
        Exit For
        End If
        End If
        Next mtable
        If okntable Then GoTo t111
Else
t111:
MyEr "No tables in Database " + ExtractNameOnly(base), "Δεν υπάρχουν αρχεία στη βάση δεδομένων " + ExtractNameOnly(base)
Exit Sub
End If
' now we have mtable from mybase

 mtable.indexes("ndx").remove  ' remove the old index/
 Err.clear
 If mcat.activeconnection.errors.Count > 0 Then
 mcat.activeconnection.errors.clear
 End If
 Err.clear
   Set pindex = CreateObject("ADOX.Index")
    pindex.name = "ndx"  ' standard
    pindex.indexnulls = 0 ' standrard
  
        While FastSymbol(r$, ",")
        If IsStrExp(bstackstr, r$, fs) Then
        If FastSymbol(r$, ",") Then
        If IsExp(bstackstr, r$, o) Then
        
        pindex.Columns.Append fs
        If o = 0 Then
        pindex.Columns(fs).sortorder = CLng(1)
        Else
        pindex.Columns(fs).sortorder = CLng(2)
        End If
        End If
        End If
                 
        End If
        Wend
        If pindex.Columns.Count > 0 Then
        mtable.indexes.Append pindex
             If Err.Number Then
         MyEr Err.Description, Err.Description
         Exit Sub
        End If
mcat.TABLES.Append mtable
mcat.TABLES.refresh
End If
    
End Sub
Public Sub NewTable(bstackstr As basetask, r$)
'BASE As String, tablename As String, ParamArray flds()
Dim base As String, tablename As String, fs As String, I&, n As Double, l As Double, ok As Boolean
ok = False
If IsStrExp(bstackstr, r$, base) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, tablename) Then
ok = True
End If
End If
End If

If Not ok Then Exit Sub
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = "" Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
    Dim okndx As Boolean, okntable As Boolean, one_ok As Boolean
    ' Dim wrkDefault As Workspace
    Dim myBase ' As Database
    Err.clear
    On Error Resume Next
                   If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Sub
                Else
                If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                    myBase.Open JetPrefix & JetPostfix
                    If Err.Number Then
                    MyEr Err.Description, Err.Description
                    Exit Sub
                    End If
                Else
                    myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                End If
                End If
                PushOne base, myBase
            End If
           Err.clear

    On Error Resume Next
   okntable = True
Dim cat, mtable, ll$
  Set cat = CreateObject("ADOX.Catalog")
           Set cat.activeconnection = myBase


If Err.Number <> 0 Then
ll$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = ll$
 PushOne base, myBase
 Err.clear
 Set cat.activeconnection = myBase
If Err.Number Then
MyEr Err.Description & " " & mtable, Err.Description & " " & mtable
Exit Sub
End If
End If

    Set mtable = CreateObject("ADOX.TABLE")
         
' check if table exist

           If cat.TABLES.Count > 0 Then
        For Each mtable In cat.TABLES
          If mtable.Type = "TABLE" Then
        If mtable.name = tablename Then
        okntable = False
        Exit For
        End If
        End If
        Next mtable
       If okntable Then
       Set mtable = CreateObject("ADOX.TABLE")      ' get a fresh one
        mtable.name = tablename
       End If
    
    
 With mtable.Columns

                Do While FastSymbol(r$, ",")
                
                        If IsStrExp(bstackstr, r$, fs) Then
                        one_ok = True
                                If FastSymbol(r$, ",") Then
                                        If IsExp(bstackstr, r$, n) Then
                                
                                            If FastSymbol(r$, ",") Then
                                                If IsExp(bstackstr, r$, l) Then
                                                If n = 10 Then n = 202
                                                If n = 12 Then n = 203: l = 0
                                                    If l <> 0 Then
                                                
                                                     .Append fs, n, l
                                                    Else
                                                     .Append fs, n
                                           
                                                    End If
                                        
                                                End If
                                            End If
                                        End If
                        
                                End If
                
                        End If
                
                Loop
               
End With
        If okntable Then
        cat.TABLES.Append mtable
        If Err.Number Then
         MyEr Err.Description, Err.Description
         Exit Sub
        End If
        cat.TABLES.refresh
        ElseIf Not one_ok Then
        cat.TABLES.Delete tablename
        cat.TABLES.refresh
        End If
        
' may the objects find the creator...


End If



End Sub


Sub BaseCompact(bstackstr As basetask, r$)

Dim base As String, conn, BASE2 As String, realtype$
If Not IsStrExp(bstackstr, r$, base) Then
MissParam r$
Else
If FastSymbol(r$, ",") Then
If Not IsStrExp(bstackstr, r$, BASE2) Then
MissParam r$
Exit Sub
End If
End If
'only for mdb
If Left$(base, 1) = "(" Or JetPostfix = ";" Then Exit Sub ' we can't compact in ODBC use control panel

''If JetPrefix <> JetPrefixHelp Then Exit Sub
  On Error Resume Next
  
If ExtractPath(base) = "" Then
base = mylcasefILE(mcd + base)
Else
  If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
realtype$ = Trim$(ExtractType(base))
If realtype$ <> "" Then
    base = ExtractPath(base, True) + ExtractNameOnly(base)
    If BASE2 = "" Then BASE2 = strTemp & CStr(Timer) & "_0." + realtype$ Else BASE2 = ExtractPath(BASE2) + CStr(Timer) + "_0." + realtype$
    Set conn = CreateObject("JRO.JetEngine")
    base = base & "." + realtype$

   conn.CompactDatabase JetPrefix & base & JetPostfixUser, _
                                GetStrUntil(";", "" + JetPrefix) & _
                                GetStrUntil(":", "" + JetPostfix) & ":Engine Type=5;" & _
                                "Data Source=" & BASE2 & JetPostfixUser
                                

    
    If Err.Number = 0 Then
    If ExtractPath(base) <> ExtractPath(BASE2) Then
       KillFile base
       Sleep 50
        If Err.Number = 0 Then
            MoveFile BASE2, base
            Sleep 50

        Else
            If GetDosPath(BASE2) <> "" Then KillFile BASE2
        End If
    
    Else
        KillFile base
        MoveFile BASE2, base
            Sleep 50
    
    End If
       
    
    
    
    Else
      
      
 
      MyErMacro r$, "Can't compact databese " & ExtractName(base) & "." & " use a back up", "Πρόβλημα με την βάση " & ExtractName(base) & ".mdb χρησιμοποίησε ένα σωσμένο αρχείο"
      End If
      Err.clear
    End If
End If
End Sub

Public Function DELfields(bstackstr As basetask, r$) As Boolean
Dim base$, table$, first$, Second$, ok As Boolean, p As Double
ok = False
If IsStrExp(bstackstr, r$, base$) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, table$) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, first$) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, Second$) Then
ok = True

           If InStr(Second$, "'") > 0 Then
                Second$ = Chr(34) & Second$ & Chr(34)
            Else
                Second$ = "'" & Second$ & "'"
                End If
ElseIf IsExp(bstackstr, r$, p) Then
ok = True
Second$ = Trim$(Str$(p))
Else
MissParam r$
End If
Else
MissParam r$

End If
Else
MissParam r$

End If
Else
MissParam r$

End If
Else
MissParam r$
End If
Else
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this we can 't killfile the base for odbc
Else
    If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = "" Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: DELfields = False: Exit Function
    If CheckMine(base) Then KillFile base
End If

End If
Else
MissParam r$
End If
If Not ok Then DELfields = False: Exit Function
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = "" Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = "" Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: DELfields = False: Exit Function
End If

Dim myBase
   On Error Resume Next
                   If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Function
                Else
                 If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                    myBase.Open JetPrefix & JetPostfix
                    If Err.Number Then
                    MyEr Err.Description, Err.Description
                    DELfields = False: Exit Function
                    End If
                 Else
                    myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                    End If
                 
                End If
                PushOne base, myBase
            End If
           Err.clear

    On Error Resume Next
Dim rec
   
   
   
   If first$ = "" Then
   MyEr "Nothing to delete", "Τίποτα για να σβήσω"
   DELfields = False
   Exit Function
   Else
   myBase.errors.clear
   myBase.Execute "DELETE * FROM [" & table$ & "] WHERE " & first$ & " = " & Second$
   If myBase.errors.Count > 0 Then
   MyEr "Can't delete " & table$, "Δεν μπορώ να διαγράψω"
   Else
    DELfields = True
   End If
   
   End If
   Set rec = Nothing

End Function

Function CheckMine(DBFileName) As Boolean
' M2000 changed to ADO...

Dim Cnn1
 Set Cnn1 = CreateObject("ADODB.Connection")

 On Error Resume Next
 Cnn1.Open JetPrefix & DBFileName & ";Jet OLEDB:Database Password=;User Id=" & DBUser & ";Password=" & DBUserPassword & ";"  ' &  DBSecurityOFF 'open the Connection
 If Err Then
 Err.clear
 Cnn1.Open JetPrefix & DBFileName & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF    'open the Connection
 If Err Then
 Else
 CheckMine = True
 End If
 Cnn1.Close
 Else
 End If
End Function


Private Sub PushOne(conname As String, v As Variant)
On Error Resume Next
conCollection.Add v, conname
Set v = conCollection(conname)
End Sub
Sub CloseAllConnections()
Dim v As Variant, bb As Boolean
On Error Resume Next
If Not Init Then Exit Sub
If conCollection.Count > 0 Then
Dim I As Long
Err.clear
For I = conCollection.Count To 1 Step -1
On Error Resume Next
bb = conCollection(I).connectionstring <> ""
If Err.Number = 0 Then

  If conCollection(I).activeconnection <> "" Then conCollection(I).Close
     ''   If conCollection(I) <> "" Then conCollection(I).Close
End If
conCollection.remove I
Err.clear
Next I
Set conCollection = New Collection
End If
Err.clear
End Sub
Public Sub RemoveOneConn(conname)
On Error Resume Next
Dim VV
VV = conCollection(conname)
If Not Err.Number <> 0 Then
Err.clear
If VV.connectionstring <> "" Then
If Err.Number = 0 Then If VV.activeconnection <> "" Then VV.Close
Err.clear
End If
conCollection.remove conname
Err.clear
End If
End Sub
Private Function getone(conname As String, this As Variant) As Boolean
On Error Resume Next
Dim v As Variant
InitMe
Err.clear
Set v = conCollection(conname)
If Err.Number = 0 Then getone = True: Set this = v
Err.clear
End Function

Private Sub InitMe()
If Init Then Exit Sub
Set conCollection = New Collection
Init = True
End Sub
Function ftype(ByVal a As Long, lang As Long) As String
Select Case lang
Case 0
Select Case a
    Case 0
ftype = "ΑΔΕΙΟ"
    Case 2
ftype = "ΨΗΦΙΟ"
    Case 3
ftype = "ΑΚΕΡΑΙΟΣ"
    Case 4
ftype = "ΑΠΛΟΣ"
    Case 5
ftype = "ΔΙΠΛΟΣ"
    Case 6
ftype = "ΛΟΓΙΣΤΙΚΟ"
    Case 7
ftype = "ΗΜΕΡΟΜΗΝΙΑ"
    Case 8
ftype = "BSTR"
    Case 9
ftype = "IDISPATCH"
    Case 10
ftype = "ERROR"
    Case 11
ftype = "ΛΟΓΙΚΟΣ"
    Case 12
ftype = "VARIANT"
    Case 13
ftype = "IUNKNOWN"
    Case 14
ftype = "DECIMAL"
    Case 16
ftype = "TINYINT"
    Case 17
ftype = "UNSIGNEDTINYINT"
    Case 18
ftype = "UNSIGNEDSMALLINT"
    Case 19
ftype = "UNSIGNEDINT"
    Case 20
ftype = "ΜΑΚΡΥΣ"   'LONG
    Case 21
ftype = "UNSIGNEDBIGINT"
    Case 64
ftype = "FILETIME"
    Case 72
ftype = "GUID"
    Case 128
ftype = "BINARY"
    Case 129
ftype = "CHAR"
    Case 130
ftype = "WCHAR"
    Case 131
ftype = "NUMERIC"
    Case 132
ftype = "USERDEFINED"
    Case 133
ftype = "DBDATE"
    Case 134
ftype = "DBTIME"
    Case 135
ftype = "ΗΜΕΡΟΜΗΝΙΑ" 'DBTIMESTAMP
    Case 136
ftype = "CHAPTER"
    Case 138
ftype = "PROPVARIANT"
    Case 139
ftype = "VARNUMERIC"
    Case 200
ftype = "VARCHAR"
    Case 201
ftype = "LONGVARCHAR"
    Case 202
ftype = "ΚΕΙΜΕΝΟ" '"VARWCHAR"
    Case 203
ftype = "LONGVARWCHAR"
    Case 204
ftype = "ΔΥΑΔΙΚΟ"  ' "VARBINARY"
    Case 205
ftype = "OLE" '"LONGVARBINARY"
    Case 8192
ftype = "ARRAY"
Case Else
ftype = "????"


End Select

Case Else  ' this is for 1
Select Case a
    Case 0
ftype = "EMPTY"
    Case 2
ftype = "BYTE"  'SMALLINT
    Case 3
ftype = "INTEGER"
    Case 4
ftype = "SINGLE"
    Case 5
ftype = "DOUBLE"
    Case 6
ftype = "CURRENCY"
    Case 7
ftype = "DATE"
    Case 8
ftype = "BSTR"
    Case 9
ftype = "IDISPATCH"
    Case 10
ftype = "ERROR"
    Case 11
ftype = "BOOLEAN"
    Case 12
ftype = "VARIANT"
    Case 13
ftype = "IUNKNOWN"
    Case 14
ftype = "DECIMAL"
    Case 16
ftype = "TINYINT"
    Case 17
ftype = "UNSIGNEDTINYINT"
    Case 18
ftype = "UNSIGNEDSMALLINT"
    Case 19
ftype = "UNSIGNEDINT"
    Case 20
ftype = "BIGINT"
    Case 21
ftype = "UNSIGNEDBIGINT"
    Case 64
ftype = "FILETIME"
    Case 72
ftype = "GUID"
    Case 128
ftype = "BINARY"
    Case 129
ftype = "CHAR"
    Case 130
ftype = "WCHAR"
    Case 131
ftype = "NUMERIC"
    Case 132
ftype = "USERDEFINED"
    Case 133
ftype = "DBDATE"
    Case 134
ftype = "DBTIME"
    Case 135
ftype = "DBTIMESTAMP"
    Case 136
ftype = "CHAPTER"
    Case 138
ftype = "PROPVARIANT"
    Case 139
ftype = "VARNUMERIC"
    Case 200
ftype = "VARCHAR"
    Case 201
ftype = "LONGVARCHAR"
    Case 202
ftype = "VARWCHAR"
    Case 203
ftype = "LONGVARWCHAR"
    Case 204
ftype = "VARBINARY"
    Case 205
ftype = "OLE"
    Case 8192
ftype = "ARRAY"


Case Else
ftype = "????"
End Select
End Select
End Function
Sub GeneralErrorReport(aBasBase As Variant)
Dim errorObject

 For Each errorObject In aBasBase.activeconnection.errors
 Debug.Print "Description :"; errorObject.Description
 Debug.Print "Number:"; Hex(errorObject.Number)
 Next
End Sub

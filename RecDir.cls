VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "recDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private LastPath As String
Private Const MAX_PATH As Long = 260
Private Const MAX_PATH_UNICODE As Long = 260 * 2 - 1
Private onetime As Boolean
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type
Private Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName(MAX_PATH * 2 - 1) As Byte
cAlternate(14 * 2 - 1) As Byte
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long

Dim myfile$(), myLevel() As Integer
Public Event DirFinished()
Public Event feedback(FileName$) ' dir or filenames...
Dim Filetypes2find As String  ' input  "jpg|bmp"  or ""
Dim mrec As Boolean
Public abort As Boolean
Public SortType As Long ' 0 is the default, 1 by name, 2 by type
Public IncludedFolders As Boolean
Public Nofiles As Boolean
Public Pattern As String
Private mTopFolder As String
Dim InUseItems As Long
Public LevelStop As Long

Private Declare Function GetDiskFreeSpace Lib "kernel32" _
 Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
 lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
 lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) _
 As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
  Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
  ByVal lpBuffer As String) As Long


Private Declare Function GetLongPathName Lib "kernel32" _
   Alias "GetLongPathNameW" _
  (ByVal lpszShortPath As Long, _
   ByVal lpszLongPath As Long, _
   ByVal cchBuffer As Long) As Long
Public lastTimeStamp As Date
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Function FileTimeToDate(File_Time As FILETIME) As Date
Dim System_Time As SYSTEMTIME
 
    ' post from JasonC in vbforums
    ' Convert the FILETIME structure into a Date.
    '
    
    ' Convert the FILETIME into a SYSTEMTIME.
    Call FileTimeToSystemTime(File_Time, System_Time)
 
    ' Convert the SYSTEMTIME into a Date.
    FileTimeToDate = SystemTimeToDate(System_Time)
    
End Function
 
Private Function SystemTimeToDate(System_Time As SYSTEMTIME) As Date
    
    '
    ' Convert a SYSTEMTIME into a Date.
    '
    
    With System_Time
        SystemTimeToDate = CDate( _
            Format$(.wMonth) & "/" & _
            Format$(.wDay) & "/" & _
            Format$(.wYear) & " " & _
            Format$(.wHour) & ":" & _
            Format$(.wMinute, "00") & ":" & _
            Format$(.wSecond, "00"))
    End With
    
End Function

Public Function GetLongName(strTest As String) As String
   Dim sLongPath As String
   Dim buff As String
   Dim cbbuff As Long
   Dim result As Long
 
   buff = Space$(MAX_PATH_UNICODE)
   cbbuff = Len(buff)
 
   result = GetLongPathName(StrPtr(strTest), StrPtr(buff), cbbuff)
 
   If result > 0 Then
      sLongPath = Left$(buff, result)
   End If
 
   GetLongName = sLongPath
 
End Function
  Function TotalDiskSpace(DriveLetter As String) As Double
' Returns the total storage capacity for a drive
Dim Dletter As String, x As Long

    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberofFreeClusters As Long
    Dim TotalClusters As Long

    Dletter = Left(DriveLetter, 1) & ":\"
    x = GetDiskFreeSpace(Dletter, SectorsPerCluster, _
      BytesPerSector, NumberofFreeClusters, TotalClusters)
    
    If x = 0 Then 'Error occurred
        TotalDiskSpace = -99 'Assign an arbitrary error value
        Exit Function
    End If
    TotalDiskSpace = _
    Int(CDbl(SectorsPerCluster * BytesPerSector) * TotalClusters / 1024 / 1024)
End Function
Function NumberofDrives() As Integer
' Returns the number of drives
    Dim i As Integer
    Dim Buffer As String * 255
    Dim BuffLen As Long
    Dim DriveCount As Integer
   
    BuffLen = GetLogicalDriveStrings(Len(Buffer), Buffer)
    DriveCount = 0
' Search for a null -- which separates the drives
    For i = 1 To BuffLen
        If Asc(Mid(Buffer, i, 1)) = 0 Then _
          DriveCount = DriveCount + 1
    Next i
    NumberofDrives = DriveCount
End Function

Function DriveName(Index As Integer) As String
' Returns the drive letter using an index
' Returns an empty string if index > number of drives
    Dim i As Integer
    Dim Buffer As String * 255
    Dim BuffLen As Long
    Dim TheDrive As String
    Dim DriveCount As Integer
   
    BuffLen = GetLogicalDriveStrings(Len(Buffer), Buffer)

    TheDrive = ""
    DriveCount = 0
    For i = 1 To BuffLen
        If Asc(Mid(Buffer, i, 1)) <> 0 Then _
          TheDrive = TheDrive & Mid(Buffer, i, 1)
        If Asc(Mid(Buffer, i, 1)) = 0 Then 'null separates drives
            DriveCount = DriveCount + 1
            If DriveCount = Index Then
                DriveName = UCase(Left(TheDrive, 1))
                Exit Function
            End If
            TheDrive = ""
        End If
    Next i
End Function
    

Function FreeDiskSpace(DriveLetter As String) As Double
' Returns the number of free bytes for a drive
Dim Dletter As String, x As Double
    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberofFreeClusters As Long
    Dim TotalClusters As Long

    Dletter = Left(DriveLetter, 1) & ":\"
    x = GetDiskFreeSpace(Dletter, SectorsPerCluster, _
      BytesPerSector, NumberofFreeClusters, TotalClusters)
    
    If x = 0 Then 'Error occurred
        FreeDiskSpace = -99 'Assign an arbitrary error value
        Exit Function
    End If
    FreeDiskSpace = _
      SectorsPerCluster * BytesPerSector * NumberofFreeClusters
End Function




Public Property Get path() As String
If Len(LastPath) = 0 Then
Else
If Right$(LastPath, 1) <> "\" Then
path = LastPath & "\"
Else
path = LastPath
End If
End If
End Property

Public Property Get IsEmpty() As Boolean

IsEmpty = InUseItems = 0

End Property
Public Property Get listcount() As Long
If IsEmpty Then
listcount = 0
Else
listcount = InUseItems  ' UBound(myfile$())
End If
End Property
Public Sub MarkLevel(item As Long)
myLevel(item) = -Abs(myLevel(item))
End Sub
Public Sub unMarkLevel(item As Long)  ' this mark can be used for multiple choosing or for hide/show.
myLevel(item) = Abs(myLevel(item))    'prefered the first one because glist cannot hide item...yet
End Sub
Public Function ReadLevel(item As Long) As Integer
ReadLevel = Abs(myLevel(item))
End Function
Public Function ReadMark(item As Long) As Boolean
ReadMark = Sgn(myLevel(item)) = -1
End Function
Public Function FindFolder(item As Long) As String
Dim i As Long, lev As Integer
lev = Abs(myLevel(item))
If lev = 1 Then
FindFolder = path
Else
For i = item - 1 To 0 Step -1
If Abs(myLevel(i)) < lev Then Exit For
Next i
 FindFolder = Mid$(myfile$(i), 2) & "\"
End If
End Function

Public Sub AddAFile(file$, level As Integer)
If Not onetime Then onetime = True: Once
If InUseItems >= UBound(myfile$()) Then
ReDim Preserve myfile$(UBound(myfile$()) * 2)
ReDim Preserve myLevel(UBound(myLevel()) * 2) As Integer
End If
 myfile$(InUseItems) = file
myLevel(InUseItems) = level  ' level must be not 0, mark = false
InUseItems = InUseItems + 1


End Sub
Public Sub ClearList()
ReDim myfile$(100)
ReDim myLevel(100) As Integer
InUseItems = 0
End Sub
Public Property Get List(Index As Long) As String
If Index < 0 Then
List = ""
Else
If Not onetime Then onetime = True: Once
List = myfile$(Index)
End If
End Property
Public Property Let List(Index As Long, rhs As String)
myfile$(Index) = rhs
End Property

Private Function CheckFile(ByVal expan$) As Boolean
Dim ee As Integer
If Filetypes2find = "||" Then
CheckFile = True
Else
CheckFile = InStr(Filetypes2find, "|" & UCase(Mid$(expan$, InStrRev(expan$, ".") + 1)) & "|") > 0
End If
End Function

Public Function Dir2(Optional ByVal ff As String = "", Optional FileType$ = "", Optional recursive As Boolean = True, Optional Append As Boolean = False) As String
Dim i As Integer
If LastPath <> "" And ff = "" Then

Else
LastPath = ff

Static ndx As Long
If ff = "" Then
ClearList

ndx = 0
For i = 1 To NumberofDrives
 'Me.additem DriveName(i) & ":\" & CStr(TotalDiskSpace(DriveName(i)))
 AddAFile Chr$(2) + DriveName(i) & ":", 1 + Abs(TotalDiskSpace(DriveName(i)) < 0) 'use level as not enabled
Next i
If listcount > 0 Then Dir2 = List(0)
RaiseEvent DirFinished
Exit Function
End If
'If Right$(ff, 1) <> "\" Then ff = ff & "\"
'' filetype$ input  "jpg|bmp"  or ""

End If
Filetypes2find = "|" & UCase(FileType$) & "|"
mrec = recursive
Dim s As String


If ff <> "" Then
If Append Then
ndx = listcount
Else
            ClearList ' CHANGE THIS TO APPEND...
           ndx = 0
    End If
                abort = False
' ff must be a path
' so strip \
            If Right$(ff, 1) = "\" Then ff = Left$(ff, Len(ff) - 1)
        
                If look(1, ff, (ndx), IncludedFolders, Nofiles) > 0 Then
               
                    Dir2 = List(ndx)
                    ndx = ndx + 1
     
     
                    End If
           
Else
            If listcount > ndx Then
            Dir2 = List(ndx)
            ndx = ndx + 1   ' read with auto increment
            Else
            ndx = 0  ' reset
                Dir2 = "" ' this is the end
            End If
End If
RaiseEvent DirFinished
End Function
Public Function FileAttr(A$, attr As Long) As Boolean
Dim wfd As WIN32_FIND_DATA
On Error GoTo there1
Dim fhandle As Long

fhandle = FindFirstFile(StrPtr(A$), wfd)

If (fhandle > 0) Then
attr = wfd.dwFileAttributes
lastTimeStamp = FileTimeToDate(wfd.ftLastWriteTime)
FileAttr = True
FindClose fhandle
Else
Dim b$
If Right$(A$, 1) = "\" Then
b$ = A$ & "."
Else
b$ = A$ & "\."
End If
fhandle = FindFirstFile(StrPtr(b$), wfd)

If (fhandle > 0) Then
attr = wfd.dwFileAttributes
lastTimeStamp = FileTimeToDate(wfd.ftLastWriteTime)
FileAttr = True
FindClose fhandle
Else
If GetLongName(A$) <> "" Then
attr = CLng(GetAttr(A$))
lastTimeStamp = 0
FileAttr = True
End If
End If
End If
Exit Function
there1:
End Function

Public Function ExistFile(A$) As Boolean
Dim wfd As WIN32_FIND_DATA
On Error GoTo there2
Dim fhandle As Long
fhandle = FindFirstFile(StrPtr(A$), wfd)
ExistFile = (fhandle > 0)
If ExistFile Then FindClose fhandle: lastTimeStamp = FileTimeToDate(wfd.ftLastWriteTime)
Exit Function
there2:
End Function
Public Function ExistFolder(A$) As Boolean
Dim wfd As WIN32_FIND_DATA
On Error GoTo there3
Dim fhandle As Long
fhandle = FindFirstFile(StrPtr(A$ & "\*.*"), wfd)
ExistFolder = (fhandle > 0)
If ExistFolder Then FindClose fhandle: lastTimeStamp = FileTimeToDate(wfd.ftLastWriteTime): Exit Function
ExistFolder = ExistFolder And (wfd.dwFileAttributes And 16) > 0

Exit Function
there3:
End Function

Public Function Dir1(ByVal ThisFile As String, curdir As String) As String

    If curdir = "" Or ExtractPath(ThisFile) = ThisFile Then
    
      If ExistFolder(ThisFile) Then
      Dir1 = ThisFile
      End If
      
      Exit Function
    End If

    If ExtractPath(ThisFile) = "" Then ThisFile = curdir & ThisFile
    If ExistFile(ThisFile) Then Dir1 = ThisFile
End Function
Private Function look(ByVal level As Integer, ByVal folder As String, Optional offset As Long = 0, Optional foldersincluded As Boolean = False, Optional foldersonly As Boolean = False) As Long
Dim lookpath As String, GUARD As String, mf$(), maxf As Long, k As Long
If LevelStop > 0 Then If level > LevelStop Then Exit Function
Dim fold$(), fff As Long, ADDTHIS As Long, mlen As Long
Dim Filter As Long
  Dim wfd As WIN32_FIND_DATA, fhandle As Long
On Error GoTo 1000:
Dim thatmax As Long, thatmaxtoo As Long
thatmax = 20
thatmaxtoo = 20
ReDim mf$(1 To thatmax)
ReDim fold$(1 To thatmaxtoo)
If (Right$("\" & folder, 2) = "\." Or Right$("\" & folder, 3) = "\..") Then Exit Function
'  The paths need to have the correct unicode lookpath prefix "\\?\" for local HDs or "\\?\UNC" for network shares.
If Pattern <> "" Then
If Left$(folder, 2) = "\\" Then
fhandle = FindFirstFile(StrPtr(folder & "\" & Pattern), wfd)
Else
fhandle = FindFirstFile(StrPtr("\\?\" & folder & "\" & Pattern), wfd)
End If

Else
If Left$(folder, 2) = "\\" Then
fhandle = FindFirstFile(StrPtr(folder & "\*.*"), wfd)
Else
fhandle = FindFirstFile(StrPtr("\\?\" & folder & "\*.*"), wfd)
End If
End If
' if fhandle = -1 then we have a failure but this isn't the end of search
' for now we exit from such situation.
' maybe the better is to wait for a period of 100ms before we quit...reading again.
If fhandle > 0 Then
lookpath = StripNulls(wfd.cFileName)
End If


'' abort = True

If level Mod 3 = 0 Then mywait basestack1, 5

While lookpath <> ""
 
    If isdirP(wfd) Or lookpath = ".." Then
       '' If Right$("\" & lookpath, 2) = "\." Or Right$("\" & lookpath, 3) = "\.." Then
         If lookpath = "." Or lookpath = ".." Then
        ' skip
        Else
        fff = fff + 1
        If fff > thatmaxtoo Then thatmaxtoo = thatmaxtoo * 2: ReDim Preserve fold$(1 To thatmaxtoo)
        '' If foldersonly And level > 1 Then RaiseEvent feedback("\" & lookpath)
        If foldersonly Then RaiseEvent feedback("\" & lookpath)
        Select Case SortType
        Case 1, 2
        fold$(fff) = lookpath  ' no type for folder
        Case Else
        fold$(fff) = myFILEDATeTIME(wfd.ftLastWriteTime) + Format$(fff, "00000") + lookpath
        End Select
        End If
    ' do nothing
    ElseIf Not foldersonly Then
    If CheckFile(lookpath) Then
        maxf = maxf + 1
         If maxf > thatmax Then thatmax = thatmax * 2: ReDim Preserve mf$(1 To thatmax)
         Select Case SortType  ' case sensitive
         Case 1
         mf$(maxf) = lookpath
         Case 2
         mf$(maxf) = LCase(ExtractType(lookpath)) & "#" & lookpath
         Case Else
         ' maybe here filedatatime can be converted to a unicode reader
         ' I am sure that this isn't ok...for unicode
         mf$(maxf) = myFILEDATeTIME(wfd.ftLastWriteTime) + Format$(maxf, "00000") + lookpath
        End Select
        End If
    End If
  If FindNextFile(fhandle, wfd) = 0& Then
  lookpath = ""
  Else
    lookpath = StripNulls(wfd.cFileName)
    End If
Wend
FindClose fhandle
' #1
' maybe we can hide that witha switch
 If level = 1 And folder <> "" And ((LCase(folder & "\") <> LCase(TopFolder)) Or (TopFolder = "\")) Then AddAFile "...", level: ADDTHIS = ADDTHIS + 1  '  may we needaswitch here to exclude that
If foldersonly Or foldersincluded Then
     ' #2
    If fff > 0 Then QuickSortExtended fold$(), 1, fff      ' we can sort the folders here
End If
    ' 3#  for each folder we  want to put  files  first and folders + files later...(this is myway)
If maxf > 0 Then
     QuickSortExtended mf$(), 1, maxf
    Select Case SortType
    Case 1
    For k = 1 To maxf
        AddAFile mf$(k), level
    Next k
    Case 2
    For k = 1 To maxf
        AddAFile Mid$(mf$(k), InStr(mf$(k), "#") + 1), level
    Next k
    Case Else
    For k = 1 To maxf
        AddAFile Mid$(mf$(k), 20), level
    Next k
    End Select
  
    RaiseEvent feedback(folder & " (" & CStr(maxf) & ") files")
End If
If Not mrec Then
If foldersonly Or foldersincluded Then
    If fff > 0 Then
    Select Case SortType
    Case 1, 2
        For k = 1 To fff
            AddAFile Chr$(2) + folder & "\" & fold$(k), level
    Next k
    Case Else
        For k = 1 To fff
            AddAFile Chr$(2) + folder & "\" & Mid$(fold$(k), 20), level
        Next k
   End Select
        maxf = maxf + fff  'IF FOLDERONLY THEN MAXF=0
   
    End If
End If
Else
If fff > 0 Then
Dim i As Long
For i = 1 To fff
If foldersincluded Or foldersonly Then
Select Case SortType
Case 1, 2
AddAFile Chr$(2) + folder & "\" & fold$(i), level: maxf = maxf + 1
Case Else
AddAFile Chr$(2) + folder & "\" & Mid$(fold$(i), 20), level: maxf = maxf + 1
End Select
End If
Select Case SortType
Case 1, 2
maxf = maxf + look(level + 1, folder & "\" & fold$(i), maxf + offset, IncludedFolders, Nofiles)

Case Else
maxf = maxf + look(level + 1, folder & "\" & Mid$(fold$(i), 20), maxf + offset, IncludedFolders, Nofiles)
End Select
If abort Then Exit For
Next i
End If
End If


look = maxf + ADDTHIS
Exit Function
1000:
look = 0
End Function


Public Sub QuickSortExtended(Arr() As String, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
Dim Piv As String, Tmp As String '<- adjust types here, when switching to something different than Long
     If UB - LB = 1 Then
     M1 = LB
      If Arr(M1) > Arr(UB) Then Tmp = Arr(M1): Arr(M1) = Arr(UB): Arr(UB) = Tmp
      Exit Sub
     Else
       M1 = (LB + UB) \ 2
             If Arr(M1) = Arr(LB) Then
                M2 = UB - 1
                M1 = LB
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If Arr(UB) < Arr(LB) Then Tmp = Arr(LB): Arr(LB) = Arr(UB): Arr(UB) = Tmp
                        Exit Sub
                    End If
                Loop Until Arr(M1) <> Arr(LB)
                Piv = Arr(M1)
                If M1 > LB Then If Arr(LB) > Piv Then Arr(M1) = Arr(LB): Arr(LB) = Piv: Piv = Arr(M1)
            Else
                Piv = Arr(M1)
                M1 = LB
                Do While (Arr(M1) < Piv): M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While (Arr(M2) > Piv): M2 = M2 - 1: Loop
      If M1 <= M2 Then
        Tmp = Arr(M1): Arr(M1) = Arr(M2): Arr(M2) = Tmp 'swap
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While (Arr(M1) < Piv): M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortExtended Arr, LB, M2
    If M1 < UB Then QuickSortExtended Arr, M1, UB
End Sub
Private Function IsSymbol3(A$, c$) As Boolean
Dim b$
b$ = LTrim$(A$)
  If b$ <> "" Then
    If Left$(b$, Len(c$)) = c$ Then
        A$ = Mid$(b$, Len(c$) + 1)
        IsSymbol3 = True
End If
End If
End Function
Private Function PointPos(f$) As Long
Dim er As Long, er2 As Long
While IsSymbol3(f$, """")
f$ = GetStrUntil(Chr(34), f$)
Wend
Dim i As Long, j As Long, oj As Long
If f$ = "" Then
PointPos = 1
Else
er = 3
er2 = 3
For i = 1 To Len(f$)
er = er + 1
er2 = er2 + 1
Select Case Mid$(f$, i, 1)
Case "."
oj = j: j = i
Case "\", "/", ":", Is = """"
If er = 2 Then oj = 0: j = i - 2: Exit For
er2 = 1
oj = j: j = 0
If oj = 0 Then oj = i - 1: If oj < 0 Then oj = 0
Case " "
If j > 0 Then Exit For
If er2 = 2 Then oj = 0: j = i - 1: Exit For
er = 1
Case "|", "'"
j = i - 1
Exit For
Case Is > " "
If j > 0 Then oj = j Else oj = 0
Case Else
If oj <> 0 Then j = oj Else j = i
Exit For
End Select
Next i
If j = 0 Then
If oj = 0 Then
j = Len(f$) + 1
Else
j = oj
End If
End If
While Mid$(f$, j, i) = " "
j = j - 1
Wend
PointPos = j
End If
End Function
Public Function ExtractType(f$, Optional jj As Long = 0) As String
Dim i As Long, j As Long, d$
If IsSymbol3(f$, """") Then f$ = GetStrUntil(Chr(34), f$)
If f$ = "" Then ExtractType = "": Exit Function
If jj > 0 Then
j = jj
Else
j = PointPos(f$)
End If
d$ = f$ & " "
If j < Len(d$) Then
For i = j To Len(d$)
Select Case Mid$(d$, i, 1)
Case "/", "|", "'", " ", Is = """"
i = i + 1
Exit For
End Select
Next i
If (i - j - 2) < 1 Then
ExtractType = ""
Else
ExtractType = Mid$(d$, j + 1, i - j - 2)
End If
Else
ExtractType = ""
End If
End Function
Public Function ExtractPath(f$, Optional IsAfile As Boolean = True) As String
If f$ = ".." Then ExtractPath = f$: Exit Function
Dim i As Long, j As Long
j = PointPos(f$)
If Mid$(f$, j, 1) = "." Then
'
Else
j = Len(f$)
End If
For i = j To 1 Step -1
Select Case Mid$(f$, i, 1)
Case Is < " ", "\", "/", ":"
Exit For
End Select
Next i
If Not IsAfile Then
    For i = i - 1 To 1 Step -1
    Select Case Mid$(f$, i, 1)
    Case Is < " ", "\", "/", ":"
    Exit For
    End Select
    Next i
End If
If i = 0 Then
ExtractPath = ""
Else
ExtractPath = Mid$(f$, 1, i)
End If
End Function
Public Function ExtractName(f$) As String
Dim i As Long, j As Long, k$
If f$ = "" Then Exit Function
j = PointPos(f$)
If Mid$(f$, j, 1) = "." Then
k$ = ExtractType(f$, j)
Else
j = Len(f$)
End If
For i = j To 1 Step -1
Select Case Mid$(f$, i, 1)
Case Is < " ", "\", "/", ":"
Exit For
End Select
Next i
If k$ = "" Then
If Mid$(f$, i + j - i, 1) = "." Then
ExtractName = Mid$(f$, i + 1, j - i - 1)
Else
ExtractName = Mid$(f$, i + 1, j - i)

End If
Else
ExtractName = Mid$(f$, i + 1, j - i) + k$
End If

End Function
Public Function isReadOnly(ByVal f$) As Boolean
Dim V&

Err.clear
On Error Resume Next
If Right$(f$, 1) = "\" Then f$ = Mid$(f$, 1, Len(f$) - 1)
If f$ <> "" Then
If FileAttr(f$, V&) Then isReadOnly = (V& And 1) > 0 Else isReadOnly = True

Else
isReadOnly = True

End If
End Function
Private Function isdirP(wfd As WIN32_FIND_DATA) As Boolean
isdirP = (wfd.dwFileAttributes And 16) > 0
End Function
Public Function isdir(f$) As Boolean
Dim V&
Err.clear
If f$ = "" Then Exit Function
On Error Resume Next
If FileAttr(f$, V&) Then
isdir = (V& And 16) > 0
End If
End Function
Private Sub Class_Initialize()
''
End Sub
Sub Once()
ReDim myfile$(100)
ReDim myLevel(100) As Integer
End Sub
Public Function GetStrUntil(ByVal sStr As String, fromStr As String, Optional RemoveSstr As Boolean = True) As String
Dim i As Long
If fromStr = "" Then GetStrUntil = "": Exit Function
i = InStr(fromStr, sStr)
If (i < 2) And Not ((i > 0) And RemoveSstr) Then
GetStrUntil = ""
fromStr = ""
Else
GetStrUntil = Left$(fromStr, i - 1)
If RemoveSstr Then
fromStr = Mid$(fromStr, Len(sStr) + i)
Else
fromStr = Mid$(fromStr, i)
End If
End If
End Function
Public Sub DropLeft(ByVal uStr As String, fromStr As String)
Dim i As Long
i = InStr(fromStr, uStr)
If i > 0 Then
fromStr = Mid$(fromStr, i + Len(uStr))
Else
fromStr = ""
End If
End Sub
Private Function StripNulls(ByVal OriginalStr As String) As String
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function


Public Property Get TopFolder() As String
TopFolder = mTopFolder
End Property

Public Property Let TopFolder(ByVal rhs As String)
If Right$(rhs, 1) <> "\" Then
mTopFolder = rhs & "\"
Else
mTopFolder = rhs
End If
End Property
Function CleanName(ByVal name$) As String
name$ = Replace(name$, """", "-")
name$ = Replace(name$, "\", "-")
name$ = Replace(name$, "<", "-")
name$ = Replace(name$, ">", "-")
name$ = Replace(name$, ":", "-")
name$ = Replace(name$, "/", "-")
CleanName = Replace(name$, "|", "-")
End Function

Private Sub Class_Terminate()
abort = True
End Sub
Private Function myFILEDATeTIME(aa As FILETIME) As String
' i need 14 bytes,
myFILEDATeTIME = Right$("00000000000000" & Hex$(aa.dwHighDateTime) + Left$("0000" + Hex$(HighWord((aa.dwLowDateTime))), 4), 14)




End Function


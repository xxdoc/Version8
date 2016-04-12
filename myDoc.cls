VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event BreakLine(data$, datanext$)
Public Event Lastline()
Public Event ColorDataEvent(thispara As Long, data$, col() As Variant, feedback As Long)
Private lastlinelastchar As Long
Private lastcolordatacur As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameW" (ByVal lpszLongPath As Long, _
ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar& Lib "kernel32" (ByVal codepage&, ByVal dwFlags&, MultiBytes As Any, ByVal cBytes&, ByVal pWideChars&, ByVal cWideChars&)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public final$
Dim mPosId As Long
Dim mCurrentParagraph As Long
' each paragraph has a number
Private Type Paragraph
  Content() As String
  colordata() As Variant
  startfrom As Long
lines As Long  ' computed by the width of ruler...
End Type
Public LfLead As Boolean
Public CrTail As Boolean
Dim DocParaBack() As Long, DocParaNext() As Long, para() As Paragraph
Dim neopara As Long, paramax As Long, docmax As Long, docpara As Long
Dim delDocPara As Long
Public MarkParagraphID As Long  ' this is the direct paranum non indirect
Dim mKeystart As String
Public BackMove As Boolean
Public LastSelStart As Long
Public ListLoadedType As Long
Private Declare Function CompareString Lib "kernel32" Alias "CompareStringW" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long) As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private words$, Unique As Long
Private p3 As Long
Dim mlcid As Long
Private pos() As Long, SLen() As Long, hit() As Long
Dim Buffer() As Byte, bufOffset As Long, bufsize As Long, SeekBuf As Long
Function IsLast(ByVal Paranum As Long) As Boolean
' for MarkParagraphID
IsLast = (DocParaNext(Paranum) = 0)
End Function
Function IsEmpty() As Boolean
If DocParaNext(0) = 0 And DocParaBack(0) = 0 Then IsEmpty = True
End Function
Function FindNewPara() As Long
If delDocPara > 0 Then
FindNewPara = delDocPara
delDocPara = DocParaNext(delDocPara)
Else
docpara = docpara + 1
If docpara > docmax Then
docmax = docmax * 2
ReDim Preserve DocParaBack(0 To docmax)
ReDim Preserve DocParaNext(0 To docmax)
ReDim Preserve para(1 To docmax)
End If
FindNewPara = docpara
If Not para(docpara).startfrom = 0 Then Stop  ' major error
End If
End Function
Sub InsertDocPara(ByVal Paranum As Long, ByVal Previus As Long, ByVal NextPara As Long)
'ok
    DocParaNext(Paranum) = NextPara
    DocParaBack(Paranum) = Previus
    DocParaNext(Previus) = Paranum
    DocParaBack(NextPara) = Paranum
End Sub
Sub UnlinkDocPara(Paranum As Long)
   DocParaNext(DocParaBack(Paranum)) = DocParaNext(Paranum)
    DocParaBack(DocParaNext(Paranum)) = DocParaBack(Paranum)
    
    para(Paranum).startfrom = 0
End Sub
Sub LinkDocParaAfter(Paranum As Long, thisparanum As Long)
    DocParaBack(DocParaNext(thisparanum)) = Paranum
    DocParaNext(Paranum) = DocParaNext(thisparanum)
    DocParaBack(Paranum) = thisparanum
    DocParaNext(thisparanum) = Paranum
    If thisparanum > 0 Then para(Paranum).startfrom = para(thisparanum).startfrom
End Sub

Sub RemoveDocParaIndex(Paranum As Long)
    DocParaNext(DocParaBack(Paranum)) = DocParaNext(Paranum)
    DocParaBack(DocParaNext(Paranum)) = DocParaBack(Paranum)
    para(Paranum).lines = 0   ' clear' also can be a flag to show that this is deleted
    ReDim para(Paranum).Content(1 To 1)
    ReDim para(Paranum).colordata(1)
    DocParaNext(Paranum) = delDocPara
    para(Paranum).startfrom = 0
    delDocPara = Paranum
End Sub
Function InvalidPara(Paranum) As Boolean
If Paranum < 1 Or Paranum > docpara Then
InvalidPara = True
Else
InvalidPara = para(Paranum).lines = 0
End If
End Function
Sub WritePara(Paranum As Long, ByVal data$)
    Dim neodata$, neolines As Long, thatmax As Long, lookthis As Long
    Do
    neolines = para(Paranum).lines + 1
    If neolines = 1 Then
    thatmax = 10
    ReDim Preserve para(Paranum).Content(1 To thatmax)
        ReDim para(Paranum).colordata(1)
       '' RaiseEvent ColorDataEvent(Paranum, data$, para(Paranum).ColorData(), LookThis)
          
    End If
    thatmax = UBound(para(Paranum).Content())
    If thatmax < neolines Then
    thatmax = thatmax * 2
    ReDim Preserve para(Paranum).Content(1 To thatmax)
    ReDim para(Paranum).colordata(1)
    
    End If
    If False Then
    
   End If
    
    RaiseEvent BreakLine(data$, neodata$)  ' no method inside this class
    If neodata$ = "" Then neodata$ = data$: data$ = ""
    para(Paranum).Content(neolines) = neodata$
    para(Paranum).lines = neolines
    neodata$ = ""
    Loop Until data$ = ""
End Sub
Sub WriteParaOneline(Paranum As Long, ByVal neodata$)
    ReDim para(Paranum).Content(1 To 1)
    para(Paranum).Content(1) = neodata$
    para(Paranum).lines = 1
    ReDim para(Paranum).colordata(1)
        para(Paranum).startfrom = 0
        
End Sub
Public Sub ColorThis(mNum As Long)

Dim docback As Long, lookthis As Long, data$, oldstart As Long, lenfinal As Long
lenfinal = Len(final$)
If lenfinal = 0 Then final$ = ";": lenfinal = 1
    If mNum = 0 Then Exit Sub
    data$ = TextParagraph(mNum)
    If Trim(data$) = "" Then
        lookthis = DocParaBack(mNum)
        If lookthis > 0 Then
                    If para(lookthis).startfrom > 0 And UBound(para(lookthis).colordata()) = 2 Then
                                    ReDim para(mNum).colordata(2)
                                    para(mNum).colordata(0) = 15
                                    para(mNum).colordata(1) = 1
                                    If para(mNum).startfrom <> mNum Then
                                            para(mNum).startfrom = para(lookthis).startfrom
                                    End If
                    Else
                                para(mNum).startfrom = 0
                    End If
        End If
        GoTo finish
    End If
    docback = DocParaBack(mNum)
    If docback = 0 Then
            docback = mNum
again:
            oldstart = para(mNum).startfrom
            data$ = TextParagraph(mNum)
            RaiseEvent ColorDataEvent(mNum, data$, para(mNum).colordata(), lookthis)
            If lookthis = 15 Then
                    
                        para(mNum).startfrom = mNum
                        docback = mNum
            Else
                     para(mNum).startfrom = 0
                        If oldstart <> 0 Then
                                 lookthis = DocParaNext(mNum)
                                Do While lookthis > 0
                                        If para(lookthis).startfrom <> oldstart Then mNum = lookthis: GoTo finish
                                        para(lookthis).startfrom = 0
                                        ColorThis lookthis
                                        
                                Loop
                        Else
                        lookthis = mNum
                        End If
                        mNum = lookthis
                        GoTo finish
            End If
again2:
            lookthis = DocParaNext(mNum)
            Do While lookthis > 0
                        data$ = TextParagraph(lookthis)
                        If Left$(Trim$(data$), lenfinal) = final$ Or Right$(RTrim(data$), lenfinal) = final$ Then
                                        ReDim para(lookthis).colordata(4)
                                        para(lookthis).colordata(0) = 15
                                        para(lookthis).colordata(1) = 1
                                        para(lookthis).colordata(2) = 14
                                        para(lookthis).colordata(3) = rinstr(data$, final$) + lenfinal
                                        para(lookthis).startfrom = para(docback).startfrom
                                        ' restore break
                                        mNum = lookthis
again3:
                                        lookthis = DocParaNext(mNum)
                                        Do While lookthis > 0
                                            If para(lookthis).startfrom <> para(docback).startfrom Then GoTo finish
                                            para(lookthis).startfrom = 0
                                            ColorThis lookthis
                                        Loop
                                        mNum = lookthis
                                        GoTo finish
                    Else
                    If data$ <> "" Then
                            ReDim para(lookthis).colordata(3)
                            para(lookthis).colordata(0) = 15
                            para(lookthis).colordata(1) = 1
                            
                               para(lookthis).colordata(2) = 15
                            para(lookthis).colordata(3) = Len(data$) + 1
                            Else
                            ReDim para(lookthis).colordata(1)
                            para(lookthis).colordata(0) = 15
                            para(lookthis).colordata(1) = 1
                            End If
                            para(lookthis).startfrom = para(docback).startfrom
                    End If
                    lookthis = DocParaNext(lookthis)
            Loop
                    mNum = lookthis
                    GoTo finish
    ElseIf para(docback).startfrom = 0 Then
    lookthis = 0
      GoTo again
    Else   ' no is  a part
    If UBound(para(docback).colordata()) = 4 Then
        lookthis = 0
    GoTo again
    Else
        mNum = docback
        If para(docback).startfrom = docback Then
        lookthis = 0: GoTo again
        End If
         GoTo again2
    End If
    End If
finish:
If mNum = 0 Then Exit Sub
    mNum = DocParaNext(mNum)
End Sub
 

Sub PlaceColorData(Paranum As Long, thatcolordata() As Variant)
    para(Paranum).colordata = thatcolordata()
End Sub
Sub ReWritePara(Paranum As Long, ByVal data$, Optional skipcolor As Boolean = False)
    Dim neodata$, neolines As Long, thatmax As Long, lookthis As Long
    If Paranum = 0 Then Exit Sub
    para(Paranum).lines = 0  ' clear
    ReDim para(Paranum).Content(1 To 1)
    Do
    neolines = para(Paranum).lines + 1
    If neolines = 1 Then
    thatmax = 10
    ReDim Preserve para(Paranum).Content(1 To thatmax)
    End If
    thatmax = UBound(para(Paranum).Content())
    If thatmax < neolines Then
    thatmax = thatmax * 2
    ReDim Preserve para(Paranum).Content(1 To thatmax)
    End If
    RaiseEvent BreakLine(data$, neodata$)  ' no method inside this class
    If neodata$ = "" Then neodata$ = data$: data$ = ""
    para(Paranum).Content(neolines) = neodata$
    para(Paranum).lines = neolines
    neodata$ = ""
    Loop Until data$ = ""
     If Not skipcolor Then
            ColorThis (Paranum)
        
    End If
End Sub
Sub AppendParagraphOneLine(myparagraph$)
    Dim getAparaNum As Long
    getAparaNum = FindNewPara
    WriteParaOneline getAparaNum, myparagraph$
    If docpara = 0 Then
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
End Sub
Function AppendParagraph(myparagraph$) As Long
    Dim getAparaNum As Long
    getAparaNum = FindNewPara
    
    ' getAparaNum is paragraph absolute address
    If docpara = 0 Then
    WritePara getAparaNum, myparagraph$
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
    WritePara getAparaNum, myparagraph$
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    AppendParagraph = getAparaNum
End Function
Function AppendParagraphPlusColor(myparagraph$, colordata() As Variant, Optional offset As Long = 0) As Long
    Dim getAparaNum As Long, mcolor() As Variant
    mcolor() = colordata()
    getAparaNum = FindNewPara
    
    ' getAparaNum is paragraph absolute address
    If docpara = 0 Then
    WritePara getAparaNum, myparagraph$
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
    WritePara getAparaNum, myparagraph$
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    If offset > 0 Then mcolor() = Relocate(mcolor(), offset)
    para(getAparaNum).colordata() = mcolor()
    AppendParagraphPlusColor = getAparaNum
End Function

Function InsertAfterParagraph(Paranum As Long, myparagraph$) As Long
    Dim getAparaNum As Long
    getAparaNum = FindNewPara
    WritePara getAparaNum, myparagraph$
    If Paranum = -1 Then
     If docpara = 0 Then
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    Else
    InsertDocPara getAparaNum, Paranum, DocParaNext(Paranum)
    End If
    InsertAfterParagraph = getAparaNum
End Function
Function InsertAfterParagraphPlusColor(Paranum As Long, myparagraph$, colordata() As Variant, Optional offset As Long = 0) As Long
    Dim getAparaNum As Long, mcolor() As Variant
    mcolor() = colordata()
    getAparaNum = FindNewPara
    WritePara getAparaNum, myparagraph$
    If Paranum = -1 Then
     If docpara = 0 Then
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    Else
    InsertDocPara getAparaNum, Paranum, DocParaNext(Paranum)
    End If
        If offset > 0 Then mcolor() = Relocate(mcolor(), offset)
    para(getAparaNum).colordata() = mcolor()
    
    InsertAfterParagraphPlusColor = getAparaNum
End Function

Function InsertAfterParagraph2(Paranum As Long, myparagraph$) As Long
    Dim getAparaNum As Long
    getAparaNum = FindNewPara

    WritePara getAparaNum, myparagraph$

    If Paranum = -1 Then
     If docpara = 0 Then
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    Else
    InsertDocPara getAparaNum, Paranum, DocParaNext(Paranum)
    End If
    
  If Paranum > 0 Then para(DocParaNext(Paranum)).startfrom = para(Paranum).startfrom
  ColorThis DocParaNext(Paranum)
    InsertAfterParagraph2 = getAparaNum
End Function
 
Function InsertBeforeParagraph(Paranum As Long, myparagraph$) As Long
    Dim getAparaNum As Long
    getAparaNum = FindNewPara
    WritePara getAparaNum, myparagraph$
       InsertDocPara getAparaNum, DocParaBack(Paranum), Paranum
    InsertBeforeParagraph = getAparaNum
End Function
Property Let textDocFast(norawtext As String)
    Dim where As Long ', before As Long, t_text As String
    Dim a() As String
    where = 1
    a() = Split(norawtext, vbCrLf)
        
        If UBound(a()) = LBound(a()) Then
        AppendParagraphOneLine norawtext
        Else
      For where = LBound(a()) To UBound(a())
      AppendParagraphOneLine a(where)
      Next where
    End If
End Property
Property Let textDoc(norawtext As String)
    Dim where As Long
    Dim a() As String
    where = 1
    If InStr(norawtext, vbLf) = 0 Then
    a() = Split(norawtext, vbCr)
    Else
    a() = Split(norawtext, vbCrLf)
    End If
        norawtext = Join(a(), vbLf)
        
        If InStr(norawtext, vbLf) = 0 Then
        AppendParagraph norawtext
        Else
      a() = Split(norawtext, vbLf)
      For where = LBound(a()) To UBound(a())
      AppendParagraph a(where)
      Next where
    End If
    ' colorize at a second stage...(new)
WrapAgain
End Property
Property Let textDocColor(norawtext As String)
    Dim where As Long
    Dim a() As String
    where = 1
    If InStr(norawtext, vbLf) = 0 Then
    a() = Split(norawtext, vbCr)
    Else
    a() = Split(norawtext, vbCrLf)
    End If
        norawtext = Join(a(), vbLf)
        
        If InStr(norawtext, vbLf) = 0 Then
        ColorThis AppendParagraph(norawtext)
        Else
      a() = Split(norawtext, vbLf)
      For where = LBound(a()) To UBound(a())
     ColorThis AppendParagraph(a(where))
      Next where
    End If
    ' colorize at a second stage...(new)
WrapAgain
End Property

Sub textParagraphPlusColor(norawtext As String, colordata() As Variant, Optional offset As Long = 0)
Dim where As Long, copyColor() As Variant
where = AppendParagraph(norawtext)
copyColor() = colordata()
If offset > 0 Then copyColor() = Relocate(copyColor(), offset)
para(where).colordata = copyColor()
End Sub
Public Function textFormat(ByVal linebreak2use$) As String
'
Dim mlines As Long, neodata As String
 Dim mycur As Long, thiscur, l As Long
mycur = DocParaNext(0)
textFormat = Space$(500)
Do While mycur <> 0
    If thiscur <> 0 Then neodata = linebreak2use$ Else thiscur = 1
    For mlines = 1 To para(mycur).lines
        neodata = neodata + para(mycur).Content(mlines)
    Next mlines
l = Len(neodata)
If Len(textFormat) < thiscur + l Then textFormat = textFormat + Space$((thiscur + l) + 100)
Mid$(textFormat, thiscur, l) = neodata
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
textFormat = Left$(textFormat, thiscur - 1)
End Function
Private Function textArray(ByVal linebreak2use$, mycur As Long, minimumbytes2exit As Long) As Byte()
'
Dim mlines As Long, neodata As String
 Dim thiscur, l As Long, textbuf As String
If minimumbytes2exit < 500 Then minimumbytes2exit = 500
textbuf = Space$(500)
Do While mycur <> 0 And Not Len(textbuf) > minimumbytes2exit
    If thiscur <> 0 Then neodata = linebreak2use$ Else thiscur = 1
    For mlines = 1 To para(mycur).lines
        neodata = neodata + para(mycur).Content(mlines)
    Next mlines
l = Len(neodata)
If Len(textbuf) < thiscur + l Then textbuf = textbuf + Space$((thiscur + l) + 100)
Mid$(textbuf, thiscur, l) = neodata
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
textArray = Left$(textbuf, thiscur - 1)
End Function
Private Function textString(ByVal linebreak2use$, mycur As Long, minimumbytes2exit As Long) As String
'
Dim mlines As Long, neodata As String
 Dim thiscur, l As Long, textbuf As String
If minimumbytes2exit < 500 Then minimumbytes2exit = 500
textbuf = Space$(500)
Do While mycur <> 0 And Not Len(textbuf) > minimumbytes2exit
    If thiscur <> 0 Then neodata = linebreak2use$ Else thiscur = 1
    For mlines = 1 To para(mycur).lines
        neodata = neodata + para(mycur).Content(mlines)
    Next mlines
l = Len(neodata)
If Len(textbuf) < thiscur + l Then textbuf = textbuf + Space$((thiscur + l) + 100)
Mid$(textbuf, thiscur, l) = neodata
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
textString = Left$(textbuf, thiscur - 1)
End Function
Property Get textDoc() As String
Attribute textDoc.VB_UserMemId = 0
Dim mlines As Long, neodata As String
 Dim mycur As Long, thiscur, l As Long
mycur = DocParaNext(0)
textDoc = Space$(500)
Do While mycur <> 0
    If thiscur <> 0 Then neodata = vbCrLf Else thiscur = 1
    For mlines = 1 To para(mycur).lines
        neodata = neodata + para(mycur).Content(mlines)
    Next mlines
l = Len(neodata)
If Len(textDoc) < thiscur + l Then textDoc = textDoc + Space$((thiscur + l) + 100)
Mid$(textDoc, thiscur, l) = neodata
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
If Typename(thiscur) = "Empty" Then

Else
textDoc = Left$(textDoc, thiscur - 1)
End If
End Sub
Function textHtmlDoc(colortable() As Variant) As String
Dim mlines As Long, neodata As String
 Dim mycur As Long, thiscur, l As Long, a$, i As Long, mcolor() As Variant, qu$
 qu$ = """"
mycur = DocParaNext(0)
textHtmlDoc = Space$(500)
Do While mycur <> 0
    neodata = ""
    For mlines = 1 To para(mycur).lines
        neodata = neodata + para(mycur).Content(mlines)
    Next mlines
If neodata = "" Then
If thiscur <> 0 Then neodata = "<p>" Else thiscur = 1
Else
' put color here
a$ = ReplaceStr("</", Chr$(1) + Chr$(2), neodata)
a$ = ReplaceStr("<", Chr$(3), a$)
a$ = ReplaceStr(">", Chr$(4), a$)
a$ = ReplaceStr("  ", Chr$(7) + Chr$(7), a$)
a$ = ReplaceStr(Chr$(7) + " ", Chr$(7) + Chr$(7), a$)
If UBound(para(mycur).colordata()) > 1 Then
mcolor() = para(mycur).colordata()
neodata = ""
l = 1
If mcolor(l) > 1 Then mcolor(l) = 1
Do
If l > 3 Then
If mcolor(l - 3) <> 15 And mcolor(l - 3) <> 7 Then
If mcolor(l) - mcolor(l - 2) < 1 Then
If l > 1 Then neodata = neodata + Replace(Mid$(a$, mcolor(l - 2)), Chr$(7), " ") + "</font>"
Else
If l > 1 Then neodata = neodata + Replace(Mid$(a$, mcolor(l - 2), mcolor(l) - mcolor(l - 2)), Chr$(7), " ") + "</font>"
End If
Else
If l > 1 Then
    If mcolor(l) - mcolor(l - 2) < 1 Then
    neodata = neodata + Mid$(a$, mcolor(l - 2)) + "</font>"
    Else
neodata = neodata + Mid$(a$, mcolor(l - 2), mcolor(l) - mcolor(l - 2)) + "</font>"
End If
End If
End If
Else
If l > 1 Then
    If mcolor(l) - mcolor(l - 2) < 1 Then
    neodata = neodata + Mid$(a$, mcolor(l - 2)) + "</font>"
    Else
    neodata = neodata + Mid$(a$, mcolor(l - 2), mcolor(l) - mcolor(l - 2)) + "</font>"
    End If
End If
End If
neodata = neodata + "<font color=" + qu$ + "#" + colortable(mcolor(l - 1)) + qu$ + ">"

  l = l + 2
Loop Until l > UBound(mcolor())
If mcolor(l - 2) > 0 Then
If mcolor(l - 3) <> 15 And mcolor(l - 3) <> 7 Then
a$ = neodata + Replace(Mid$(a$, mcolor(l - 2)), Chr$(7), " ") + "</font>"
Else
a$ = neodata + Mid$(a$, mcolor(l - 2)) + "</font>"
End If
Else
a$ = neodata + "</font>"
End If

End If

a$ = ReplaceStr(Chr$(1) + Chr$(2), "&lt;/", a$)
a$ = ReplaceStr(Chr$(3), "&lt;", a$)
a$ = ReplaceStr(Chr$(4), "&gt;", a$)
a$ = ReplaceStr(Chr$(7), "&nbsp;", a$)
If thiscur <> 0 Then neodata = "<br>" + a$ Else thiscur = 1: neodata = a$

End If
l = Len(neodata)
If Len(textHtmlDoc) < thiscur + l Then textHtmlDoc = textHtmlDoc + Space$((thiscur + l) + 100)
Mid$(textHtmlDoc, thiscur, l) = neodata
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
If Typename(thiscur) = "Empty" Then

Else
textHtmlDoc = Left$(textHtmlDoc, thiscur - 1) + vbCrLf
End If
End Function
Property Get textDoc1LINE() As String  ' MORE SPEED IF WE KNOW THAT WE DON'T USE LINE BREAK
Dim mlines As Long, neodata As String
 Dim mycur As Long, thiscur, l As Long
mycur = DocParaNext(0)
textDoc1LINE = Space$(1000)
Do While mycur <> 0
    If thiscur <> 0 Then neodata = vbCrLf Else thiscur = 1
        neodata = neodata + para(mycur).Content(1)
   
l = Len(neodata)
If Len(textDoc1LINE) < thiscur + l Then textDoc1LINE = textDoc1LINE + Space$((thiscur + l) + 100)
Mid$(textDoc1LINE, thiscur, l) = neodata
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
textDoc1LINE = Left$(textDoc1LINE, thiscur - 1)
End Property
Private Sub InsertLine(b$, Paranum, cur)
    Dim mlines As Long, neodata As String
    If Paranum = 0 Then Exit Sub
    If cur <> 0 Then neodata = vbCrLf Else cur = 1
    For mlines = 1 To para(Paranum).lines
        neodata = neodata + para(Paranum).Content(mlines)
    Next mlines

If Len(b$) < cur + Len(neodata) Then b$ = b$ + Space$(Len(b$) * 2)
Mid$(b$, cur, Len(neodata)) = neodata
cur = cur + Len(neodata)
End Sub

Public Sub InsertDoc(ByVal Paranum As Long, fromPos As Long, norawtext As String)
    Dim t_text As String, Once As Boolean
    Dim a() As String
    Dim b$
    Once = True
      a() = Split(norawtext, vbCrLf)
        norawtext = Join(a(), vbLf)
     a() = Split(norawtext + vbLf, vbLf)
Dim i As Long
    i = LBound(a())
    Do

    t_text = a(i)
    i = i + 1
    If t_text = "" And i > UBound(a()) Then
    If b$ <> "" Then
        ReWritePara Paranum, TextParagraph(Paranum) + b$, True
        b$ = ""
    End If
    Exit Do
    End If
    If Once Then
    Once = False
        b$ = TextParagraph(Paranum)
        If fromPos = 1 Then
          ReWritePara Paranum, t_text, True
          fromPos = Len(t_text) + 1
       
        ElseIf fromPos > Len(b$) Then
           
           ReWritePara Paranum, b$ & t_text, True
            fromPos = Len(b$ & t_text) + 1
            
            b$ = ""
        Else
        If fromPos <= 0 Then fromPos = 1
        
        ReWritePara Paranum, Left$(b$, fromPos - 1) + t_text, True
        b$ = Mid$(b$, fromPos)
        fromPos = fromPos + Len(t_text)
        
        End If
    Else
    ColorThis (Paranum)
    Paranum = InsertAfterParagraph(Paranum, t_text)
         fromPos = Len(t_text) + 1
       
    End If
    Loop
    MarkParagraphID = Paranum
   ColorThis (Paranum)
End Sub
Property Get DocLines() As Long
    Dim mycur As Long, mlines As Long
    mycur = DocParaNext(0)
    While mycur <> 0
        mlines = para(mycur).lines + mlines
        mycur = DocParaNext(mycur)
    Wend
    DocLines = mlines
End Property
Property Get DocParagraphs() As Long
    Dim mycur As Long, dPara As Long
    mycur = DocParaNext(0)
    While mycur <> 0
        dPara = dPara + 1
        mycur = DocParaNext(mycur)
    Wend
    DocParagraphs = dPara
End Property

Property Get FirstParagraphLine(Paranum As Long) As Long
    Dim mycur As Long, mlines As Long, mlinesold As Long, valid As Boolean
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        If mycur = Paranum Then valid = True: Exit Do
        mlines = para(mycur).lines + mlines
        mycur = DocParaNext(mycur)
    Loop
    If valid Then FirstParagraphLine = mlinesold
End Property

'
Public Sub FindPos(ByVal fromLineNo As Long, ByVal fromPos As Long, CharNumber As Long, toLineNo As Long, toPos As Long, toParagraph As Long, toParaPos As Long)
Dim probelen As Long, i As Long, target As Long, sum As Long
Dim mycur As Long, mlines As Long, mlinesold As Long, sumtohere As Long
toLineNo = fromLineNo
toPos = fromPos
toParagraph = DocParaNext(0)
toParaPos = 1
If fromLineNo < 1 Then Exit Sub
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        If fromLineNo <= mlines Then Exit Do
         sumtohere = sumtohere + TextParagraphLen(mycur)
        mycur = DocParaNext(mycur)
        If Not mycur = 0 Then sumtohere = sumtohere + 2         ' now we put the vbcrlf length (these are not bytes)
    Loop
    toParagraph = mycur
    If CharNumber = 0 Then Exit Sub
    If mycur <> 0 Then
        target = fromLineNo - mlinesold
        probelen = 0
        For i = 1 To target - 1
             toPos = Len(para(mycur).Content(i))
             probelen = probelen + toPos
             sumtohere = sumtohere + toPos
        Next i
        toLineNo = fromLineNo
        toPos = Len(para(mycur).Content(target))
        If CharNumber < 0 Then
            If fromPos = 0 Then
                CharNumber = sumtohere
            Else
                CharNumber = sumtohere + fromPos - 1
            End If
            Exit Sub
        Else
           sum = toPos - fromPos
        End If
        If target = para(mycur).lines Then
           
                        If (CharNumber - sum) <= 2 Then
                            If (CharNumber - sum) = 2 Then
                                 If DocParaNext(mycur) = 0 Then
                              toPos = CharNumber + fromPos
                              toParaPos = toPos + probelen - 2
                              
                              Else
                                toPos = 0
                                toParagraph = DocParaNext(mycur)
                                toLineNo = toLineNo + 1
                                toParaPos = 0
                                End If
                            Else
                                toPos = CharNumber + fromPos
                                toParaPos = toPos + probelen
                            End If
                            Exit Sub
                        End If
               
        Else
                If CharNumber <= sum Then
                    toPos = CharNumber + fromPos
                    toParaPos = toPos + probelen
                    Exit Sub
                End If
        End If
        sum = toPos - fromPos
        probelen = probelen + toPos
        Do
                For i = target + 1 To para(mycur).lines - 1
                    toLineNo = i + mlinesold  ' set the right lineNo
                    toPos = Len(para(mycur).Content(i))
                    probelen = probelen + toPos
                    sum = sum + toPos
                    If sum >= CharNumber Then
                        probelen = probelen - toPos  ' rewind
                        toParaPos = probelen + CharNumber - (sum - toPos)
                        toPos = CharNumber - (sum - toPos)
                        Exit Sub
                    End If
                Next i
    
        If para(mycur).lines < i Then
            toPos = 0
        Else
                toPos = Len(para(mycur).Content(i))
        End If
             toLineNo = i + mlinesold
               

                
                If DocParaNext(mycur) = 0 Then
                        If (CharNumber - sum) <= toPos Then
                            toPos = CharNumber - sum - 1
                        End If
                        toParaPos = toPos + probelen
                        Exit Do
                Else
                        If (CharNumber - sum - toPos) <= 2 Then
                               If (CharNumber - sum - toPos) = 2 Then
                                    toPos = 0
                                    toParagraph = DocParaNext(mycur)
                                    toLineNo = toLineNo + 1
                                    toParaPos = 0
                                Else
                                    toPos = CharNumber - sum
                                    toParaPos = toPos + probelen
                                End If
                                Exit Sub
                        End If
                End If
   
                probelen = 0
                sum = sum + toPos + 2
                mycur = DocParaNext(mycur)
                toParagraph = mycur
                target = 0
                mlinesold = mlines
                mlines = para(mycur).lines + mlines
        Loop
    End If
    toParaPos = toParaPos + 1
    toPos = toPos + 1
End Sub
Public Function RemoveEmptyParagraph(Paranum As Long) As Boolean
' but not last, only erased
Dim oldone As Long
If TextParagraphLen(Paranum) = 0 Then
RemoveEmptyParagraph = True
oldone = DocParaNext(Paranum)
If oldone = 0 Then oldone = DocParaBack(Paranum)
If oldone = 0 Then
    para(Paranum).lines = 1  ' clear
    ReDim para(Paranum).Content(1 To 1)
    ReDim para(Paranum).colordata(1)
Else
RemoveDocParaIndex Paranum
Paranum = oldone
End If
End If
End Function

Property Get LastCurLineChar() As Long
LastCurLineChar = lastlinelastchar
End Property
Property Get lastcolordata() As Variant()
If lastcolordatacur <> 0 Then
lastcolordata = Relocate(para(lastcolordatacur).colordata(), lastlinelastchar)
Else
Dim dummy(1)
lastcolordata = dummy()
End If
End Property
Function Relocate(aa() As Variant, ByVal offset As Long)
Dim i As Long, bbb(), j As Long
If offset = 0 Or UBound(aa()) <= 1 Then
Relocate = aa()
Exit Function
End If
For i = 1 To UBound(aa()) - 1 Step 2
If aa(i) > offset Then Exit For
Next i
If i > UBound(aa()) Then GoTo there:
    If aa(i) < offset Then
there:
        ReDim bbb(2)
        bbb(0) = aa(UBound(aa()) - 1)
        bbb(1) = 1
    Else
    If aa(i) - offset <> 1 Then
        ReDim bbb(UBound(aa()) - i + 4)
        bbb(0) = aa(i - 3) ' preserve color
        bbb(1) = 1
        j = 2
    Else
        ReDim bbb(UBound(aa()) - i + 2)
    End If
        For i = i - 1 To UBound(aa()) - 1 Step 2
        bbb(j) = aa(i) 'color
        bbb(j + 1) = aa(i + 1) - offset
        j = j + 2
        Next i
    End If

Relocate = bbb()
End Function
Property Get TextLine(lineno As Long) As String
If lineno < 1 Then Exit Property
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
    lastcolordatacur = mycur
    lastlinelastchar = 0
    For mlines = 1 To lineno - mlinesold - 1
    lastlinelastchar = lastlinelastchar + Len(para(mycur).Content(mlines))
    Next mlines
        TextLine = para(mycur).Content(lineno - mlinesold)
    End If
End Property
Property Get TextLine2(lineno As Long, showspaces As Boolean) As String
If lineno < 1 Then Exit Property
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
    lastcolordatacur = mycur
    lastlinelastchar = 0
    For mlines = 1 To lineno - mlinesold - 1
    lastlinelastchar = lastlinelastchar + Len(para(mycur).Content(mlines))
    Next mlines
    If showspaces Then
        TextLine2 = Replace(Replace(para(mycur).Content(lineno - mlinesold), " ", ChrW(&HB7)), ChrW(&H2007), ChrW(&H2423))
    
    Else
    TextLine2 = para(mycur).Content(lineno - mlinesold)
        End If
    End If
End Property
Property Let TextLine(lineno As Long, that As String)
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
        para(mycur).Content(lineno - mlinesold) = that
    End If
End Property
Property Get ParagraphFromOrder(ByVal order As Long) As Long
    ParagraphFromOrder = -1
    Dim mycur As Long
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order - 1
        If order = 0 Then ParagraphFromOrder = mycur: Exit Do
        mycur = DocParaNext(mycur)
    Loop
End Property
Property Get Paracolordata(Paranum As Long) As Variant()
' it is valid
Paracolordata = para(Paranum).colordata()
End Property
Property Get ParagraphOrder(Paranum As Long) As Long
    ParagraphOrder = 0
    Dim mycur As Long, order As Long
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order + 1
        If mycur = Paranum Then ParagraphOrder = order: Exit Do
        mycur = DocParaNext(mycur)
    Loop
End Property
Property Get LastParagraph() As Long
LastParagraph = DocParaBack(0)
End Property
Property Get SizeCRLF() As Long
Dim mycur As Long, suma As Long, k As Long
If IsEmpty Then
Else
 mycur = DocParaNext(0)
  k = -1
    Do While mycur > 0
    k = k + 1
        suma = suma + TextParagraphLen(mycur)
        mycur = DocParaNext(mycur)
    Loop
    End If
    SizeCRLF = suma + k * 2
End Property
Property Get TextParagraphOrder(ByVal order As Long) As String
    ' 1 is the first
    ' -1 is the last
    Dim mycur As Long
    If order = -1 Then
        mycur = DocParaBack(0)
        TextParagraphOrder = TextParagraph(mycur)
        Exit Property
    ElseIf order = 0 Then
        order = 1
    End If
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order - 1
        If order = 0 Then
            TextParagraphOrder = TextParagraph(mycur)
            Exit Property
        End If
        mycur = DocParaNext(mycur)
    Loop
End Property
Property Get TextParagraphOrderRemove(ByVal order As Long) As String
    ' 1 is the first
    ' -1 is the last
    Dim mycur As Long
    If order = -1 Then
        mycur = DocParaBack(0)
        TextParagraphOrderRemove = TextParagraph(mycur)
         RemoveDocParaIndex mycur
        Exit Property
    ElseIf order = 0 Then
        order = 1
    End If
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order - 1
        If order = 0 Then
            TextParagraphOrderRemove = TextParagraph(mycur)
            RemoveDocParaIndex mycur
            Exit Property
        End If
        mycur = DocParaNext(mycur)
    Loop
End Property
Public Function FindStrUp(s$, FirstPara As Long, FirstPosPara As Long) As Boolean
Dim a As Long, b As Long, mC As Long
a = FirstPara
mC = FirstPosPara - 1
If FirstPosPara <= 0 Then
mC = -1
a = DocParaBack(a)
If a = 0 Then Exit Function
Else
mC = mC - Len(s$)
End If
If mC > 0 Then b = InstrRevTextCompareAnyLocale(Me.TextParagraph(a), s$, mC, mlcid) 'using default locale????
If b = 0 Then
a = DocParaBack(a)
Do While b = 0 And a > 0
b = InstrRevTextCompareAnyLocale(Me.TextParagraph(a), s$, , mlcid)   'using default locale????
If b Then Exit Do
a = DocParaBack(a)
Loop
End If
If b > 0 Then
FirstPara = a
FirstPosPara = b
FindStrUp = True
End If
End Function
Public Function FindStrDown(s$, FirstPara As Long, FirstPosPara As Long) As Boolean
'return true if found
Dim a As Long, b As Long, mC As Long
If FirstPosPara <= 0 Then mC = 1 Else mC = FirstPosPara + 1
a = FirstPara
b = InstrTextCompareAnyLocale(mC, Me.TextParagraph(a), s$, mlcid)
If b = 0 Then
a = DocParaNext(a)
Do While b = 0 And a > 0
b = InstrTextCompareAnyLocale(1, TextParagraph(a), s$, mlcid)
If b Then Exit Do
a = DocParaNext(a)
Loop
End If
If b > 0 Then
FirstPara = a
FirstPosPara = b
FindStrDown = True
End If
End Function

Public Function FindStr(s$, ByVal st&, FindParaOrder As Long, FindPosPara As Long) As Long
Dim mycur As Long, ac&, ac1&, acp&, mlines As Long, wh&, order As Long
mycur = DocParaNext(0)
order = 1
    Do While mycur > 0
 order = order + 1
    ac& = acp&
    If st& <= 0 Then st& = 1
    For mlines = 1 To para(mycur).lines
        ac1& = ac& + Len(para(mycur).Content(mlines))
        If ac1& > st& Then
        If st& - ac& <= 0 Then st& = ac& + 1
        wh& = InstrTextCompareAnyLocale(st& - ac&, para(mycur).Content(mlines), s$, mlcid)
        If wh& > 0 Then
            FindParaOrder = order - 1
            FindPosPara = wh& + ac& - acp&
            FindStr = wh& + ac&
        Exit Function
        
        End If
        Exit Do
        
        End If
    Next mlines
      acp& = ac1& + 2  'for linefeed
      mycur = DocParaNext(mycur)
      
    Loop
    If mycur > 0 Then
        If mlines < para(mycur).lines Then
            For mlines = mlines To para(mycur).lines
            ac1& = ac& + Len(para(mycur).Content(mlines))
            wh& = InstrTextCompareAnyLocale(1, para(mycur).Content(mlines), s$, mlcid)
                If wh& > 0 Then
                        FindParaOrder = order
                        FindPosPara = wh& + ac& - acp&
                        FindStr = wh& + ac&
                        Exit Function
                End If
        Next mlines
        Else
            order = order - 1
        End If
          acp& = ac1& + 2   'for linefeed
          mycur = DocParaNext(mycur)
       
    End If
    
        Do While mycur > 0
        order = order + 1
        ac& = acp&
    For mlines = 1 To para(mycur).lines
        ac1& = ac& + Len(para(mycur).Content(mlines))
        wh& = InstrTextCompareAnyLocale(1, para(mycur).Content(mlines), s$, mlcid)
        If wh& > 0 Then
                 FindParaOrder = order
                 FindPosPara = wh& + ac& - acp&
                 FindStr = wh& + ac&
                 Exit Function
                 Exit Do
         End If
     Next mlines
      acp& = ac1& + 2  'for linefeed
      mycur = DocParaNext(mycur)
    Loop
    

    
End Function
Property Let TextParagraphOrder(ByVal order As Long, data$)
    Dim mycur As Long
    If order = -1 Then
        mycur = DocParaBack(0)
        ReWritePara mycur, data$
        Exit Property
    ElseIf order = 0 Then
        order = 1
    End If
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order - 1
        If order = 0 Then
            ReWritePara mycur, data$
            Exit Property
        End If
        mycur = DocParaNext(mycur)
    Loop
End Property
Property Get TextParagraph(ByVal Paranum As Long) As String
If InvalidPara(Paranum) Then Exit Property
    Dim mlines As Long, neodata As String
    If Paranum = 0 Then TextParagraph = "": Exit Property
    For mlines = 1 To para(Paranum).lines
        neodata = neodata + para(Paranum).Content(mlines)
    Next mlines
    TextParagraph = neodata
End Property
Property Get TextParagraphKey(ByVal Paranum As Long) As String
    Dim mlines As Long, neodata As String
    If Paranum = 0 Then TextParagraphKey = "": Exit Property
    For mlines = 1 To para(Paranum).lines
        neodata = neodata + para(Paranum).Content(mlines)
    Next mlines
    TextParagraphKey = Mid$(neodata, mKeystart)
End Property

Property Let TextParagraph(ByVal Paranum As Long, data$)
    If Paranum = 0 Then Exit Property
    para(Paranum).lines = 1   ' clear' also can be a flag to show that this is deleted
    ReDim para(Paranum).Content(1 To 1)
    ReDim para(Paranum).colordata(1)   ' erase old color  --
    para(Paranum).Content(1) = data$
End Property

Property Get TextParagraphLen(ByVal Paranum As Long) As Long
' without addition of 2 chars vbcrlf
    Dim mlines As Long, neodatalen As Long
    If Paranum = 0 Then TextParagraphLen = "": Exit Property
    For mlines = 1 To para(Paranum).lines
        neodatalen = neodatalen + Len(para(Paranum).Content(mlines))
    Next mlines
    TextParagraphLen = neodatalen
End Property
Function TextNextParagraph(Paranum As Long) As String
Dim mycur As Long
mycur = DocParaNext(Paranum)
If mycur > 0 Then
TextNextParagraph = TextParagraph(mycur)
Paranum = mycur
End If

End Function
Function DeleteNextParagraph(ByVal Paranum As Long) As String
    Dim mynextcur As Long
    mynextcur = DocParaNext(Paranum)
    If mynextcur <> 0 Then
        DeleteNextParagraph = TextParagraph(mynextcur)
        RemoveDocParaIndex mynextcur
    End If
End Function
Function DeletePreviusParagraph(ByVal Paranum As Long) As String
    Dim mybackcur As Long
    mybackcur = DocParaBack(Paranum)
    If mybackcur <> 0 Then
        DeletePreviusParagraph = TextParagraph(mybackcur)
        RemoveDocParaIndex mybackcur
    End If
End Function

Property Get TextParagraph2Line(ByVal Paranum As Long, ByVal Charpos As Long, newselstart) As Long
Dim mlines As Long, counter As Long
    If Paranum = 0 Then Exit Property
    If para(Paranum).lines = 0 Then Exit Sub
     For mlines = 1 To para(Paranum).lines
        counter = counter + Len(para(Paranum).Content(mlines))
        If Charpos <= counter Then
            newselstart = Charpos - (counter - Len(para(Paranum).Content(mlines))) '  - 1 to write from right to left
            TextParagraph2Line = mlines - 1
            Exit Property
        End If
    Next mlines
    If Charpos > counter Then newselstart = Len(para(Paranum).Content(para(Paranum).lines)) + 1
    TextParagraph2Line = mlines - 2
End Property
Function TextParagraphLine2(ByVal lineno As Long, thistext As String, Paranum As Long, showpar As Boolean) As Long
    ' for lineno get line and paragraph id..
    ' we want to mark the current paragraph
    Dim mycur As Long, mlines As Long, mlinesold As Long, i As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
    If lineno = mlines Then RaiseEvent Lastline
      If (lineno - mlinesold) > 0 Then
      If showpar Then
      thistext = Replace(para(mycur).Content(lineno - mlinesold), " ", ChrW(&HB7))
      thistext = Replace(thistext, ChrW(&H2007), ChrW(&H2423))
      Else
      thistext = para(mycur).Content(lineno - mlinesold)
      End If
      End If
        Paranum = mycur
        If (lineno - mlinesold) = 1 Then
            TextParagraphLine2 = 0
        Else
            For i = 1 To lineno - mlinesold - 1
                TextParagraphLine2 = TextParagraphLine2 + Len(para(mycur).Content(i))
            Next i
        End If
        
    Else
        thistext = ""
        Paranum = 0
    End If
End Function
Function TextParagraphLine(ByVal lineno As Long, thistext As String, Paranum As Long) As Long
    ' for lineno get line and paragraph id..
    ' we want to mark the current paragraph
    Dim mycur As Long, mlines As Long, mlinesold As Long, i As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
    If lineno = mlines Then RaiseEvent Lastline
      If (lineno - mlinesold) > 0 Then thistext = para(mycur).Content(lineno - mlinesold)
        Paranum = mycur
        If (lineno - mlinesold) = 1 Then
            TextParagraphLine = 0
        Else
            For i = 1 To lineno - mlinesold - 1
                TextParagraphLine = TextParagraphLine + Len(para(mycur).Content(i))
            Next i
        End If
        
    Else
        thistext = ""
        Paranum = 0
    End If
End Function

Sub GoTopLinePreviusParagraph(lineno As Long, thistext As String, Paranum As Long)
' we give Paranum and we get lineno and thistext as first line, and new paranum
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    If mycur > 0 Then
        If mycur = Paranum Then
            thistext = para(Paranum).Content(1)  ' first line
            lineno = 1
            Exit Sub
        End If
    End If
    Do While mycur <> 0
        mlinesold = mlines
        If DocParaNext(mycur) = Paranum Then Exit Do
        mlines = para(mycur).lines + mlines
        mycur = DocParaNext(mycur)
    Loop
    If DocParaNext(mycur) = Paranum Then
        Paranum = mycur
        thistext = para(Paranum).Content(1)  ' first line
        lineno = mlinesold + 1
    Else
        Paranum = DocParaNext(0): mlinesold = 0
        thistext = para(Paranum).Content(1)  ' first line
        lineno = mlinesold + 1
    End If
End Sub
Sub GoTopLineNextParagraph(lineno As Long, thistext As String, Paranum As Long)
' we give Paranum and we get lineno and thistext as first line at the next Paranum..
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        If mycur = Paranum Then Exit Do  'lineno=mlinesold+1
        mlines = para(mycur).lines + mlines
        mycur = DocParaNext(mycur)
    Loop
    If mycur = Paranum Then
        If DocParaNext(mycur) <> 0 Then
            mlines = para(mycur).lines + mlines
            mycur = DocParaNext(mycur)
        End If
        Paranum = mycur ' new paragraph id
        thistext = para(Paranum).Content(1)  ' first line
        lineno = mlinesold + 1
    Else
        mycur = DocParaNext(0): mlinesold = 0  ' go to start again if not found..
        Paranum = mycur ' new paragraph id
        thistext = para(Paranum).Content(1)  ' first line
        lineno = mlinesold + 1
    End If
End Sub
Public Sub WrapAgain()
Dim mycur As Long
mycur = DocParaNext(0)
Do While mycur <> 0
 ReWritePara mycur, TextParagraph(mycur), True
mycur = DocParaNext(mycur)
Loop

End Sub
Public Sub WrapAgainBlock(startblock As Long, endblock As Long)
Dim mycur As Long
mycur = startblock
Do While mycur <> 0
 ReWritePara mycur, TextParagraph(mycur), True
 If mycur = endblock Then Exit Do
 mycur = DocParaNext(mycur)
Loop
End Sub

Public Sub WrapAgainColor()
Dim mycur As Long
mycur = DocParaNext(0)

Do While mycur <> 0


ColorThis mycur
Loop

End Sub
Public Sub WrapAgainColorBlock(startblock As Long, endblock As Long)
Dim mycur As Long
mycur = startblock
If startblock = endblock Then
ColorThis mycur
Else
Do While mycur <> 0
ColorThis mycur
If mycur = endblock Then Exit Do
Loop

If mycur = endblock Then ColorThis mycur
End If
End Sub

Private Sub Class_Initialize()
' from this version paramax= docmax
'paramax = 20 ' 20 parapraphs ...content
docmax = 20  ' 20 paragraphs..links only
ReDim para(1 To docmax)
ReDim DocParaBack(0 To docmax)
ReDim DocParaNext(0 To docmax)
ReDim Buffer(1 To 1) As Byte
bufOffset = 0
bufsize = 0
SeekBuf = 1
mlcid = 1033
final$ = "}"
End Sub

Sub BackSpaceNchars(fromPARA As Long, fromPos As Long, ByVal Nchars As Long)
' we return the current frompara and from pos
Dim ww As Long, ThatPara
ww = TextParagraphLen(fromPARA)
'If fromPos > ww Then Nchars = Nchars - (fromPos - ww): fromPos = ww

Do While Nchars > 0
If fromPos = 0 Then
fromPARA = DocParaBack(fromPARA)
fromPos = TextParagraphLen(fromPARA) + 1
TextParagraph(fromPARA) = TextParagraph(fromPARA) + Chr$(10) + DeleteNextParagraph(fromPARA)
Nchars = Nchars - 1
ElseIf fromPos = 1 Then
If Nchars > 2 Then
Nchars = Nchars - 2
fromPARA = DocParaBack(fromPARA)
fromPos = TextParagraphLen(fromPARA) + 1
TextParagraph(fromPARA) = TextParagraph(fromPARA) + DeleteNextParagraph(fromPARA)

' CONTINUE
Else
Nchars = Nchars - 1
fromPARA = DocParaBack(fromPARA)
fromPos = TextParagraphLen(fromPARA) + 2
TextParagraph(fromPARA) = TextParagraph(fromPARA) + Chr$(13) + DeleteNextParagraph(fromPARA)
End If
ElseIf fromPos > Nchars Then

TextParagraph(fromPARA) = Left(TextParagraph(fromPARA), fromPos - Nchars - 1) + Mid$(TextParagraph(fromPARA), fromPos)
fromPos = fromPos - Nchars
Exit Do
Else
Nchars = Nchars - fromPos + 1
      ThatPara = DocParaBack(fromPARA)
      
        If fromPos = TextParagraphLen(fromPARA) Then
        RemoveDocParaIndex fromPARA
        
        Else
        TextParagraph(fromPARA) = Mid$(TextParagraph(fromPARA), fromPos)
        End If
        
        If ThatPara = 0 Then ThatPara = DocParaNext(0)
        fromPARA = ThatPara
        If fromPARA = 0 Then
        ' make only one empty paragraph
        ''Stop
        Exit Sub
        End If
      fromPos = TextParagraphLen(fromPARA) + 1
      If Nchars >= 2 Then
      Nchars = Nchars - 2
      
      TextParagraph(fromPARA) = TextParagraph(fromPARA) + DeleteNextParagraph(fromPARA)
      ElseIf Nchars = 1 Then
      TextParagraph(fromPARA) = TextParagraph(fromPARA) + Chr$(13) + DeleteNextParagraph(fromPARA)
      fromPos = fromPos + 1
      Nchars = 0
      Else
      
      Exit Sub
      End If
     
    
    
End If

Loop


End Sub

Private Sub QuickSortExtended(Arr() As Long, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
On Error GoTo abc1
Dim Piv As String, Tmp As String '<- adjust types here, when switching to something different than Long
     If UB - LB = 1 Then
     M1 = LB
      If StrComp(TextParagraphKey(Arr(M1)), TextParagraphKey(Arr(UB)), vbTextCompare) = 1 Then SwapPara M1, UB, Arr
      Exit Sub
     Else
       M1 = (LB + UB) \ 2
             If StrComp(TextParagraphKey(Arr(M1)), TextParagraphKey(Arr(LB)), vbTextCompare) = 0 Then
                M2 = UB - 1
                M1 = LB
                Piv = TextParagraphKey(Arr(LB))
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If StrComp(TextParagraphKey(Arr(UB)), Piv, vbTextCompare) = -1 Then SwapPara LB, UB, Arr
                        Exit Sub
                    End If
                Loop Until StrComp(TextParagraphKey(Arr(M1)), Piv, vbTextCompare) <> 0
                Piv = TextParagraphKey(Arr(M1))
                If M1 > LB Then If StrComp(TextParagraphKey(Arr(LB)), Piv, vbTextCompare) = 1 Then SwapPara M1, LB, Arr: Piv = TextParagraphKey(Arr(M1))
            Else
                Piv = TextParagraphKey(Arr(M1))
                M1 = LB
                Do While StrComp(TextParagraphKey(Arr(M1)), Piv, vbTextCompare) = -1: M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While StrComp(TextParagraphKey(Arr(M2)), Piv, vbTextCompare) = 1: M2 = M2 - 1: Loop
      If M1 <= M2 Then
       If M1 <> M2 Then SwapPara M1, M2, Arr
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While StrComp(TextParagraphKey(Arr(M1)), Piv, vbTextCompare) = -1: M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortExtended Arr, LB, M2
    If M1 < UB Then QuickSortExtended Arr, M1, UB
    Exit Sub
abc1:
    
    
End Sub
Sub SwapPara(ByVal order1 As Long, ByVal order2 As Long, Arr() As Long)
Dim paranum1 As Long, paranum2 As Long
If order1 = order2 Then Exit Sub
''If order1 > order2 Then paranum1 = order1: order1 = order2: order2 = paranum1
paranum1 = Arr(order1)
paranum2 = Arr(order2)
Arr(order2) = paranum1  ' now we fix the order in the temporary array.
Arr(order1) = paranum2



UnlinkDocPara paranum2
LinkDocParaAfter paranum2, Arr(order1 - 1)
UnlinkDocPara paranum1
LinkDocParaAfter paranum1, Arr(order2 - 1)



End Sub


Sub SortDoc(KeyStart As Long, order1 As Long, order2 As Long)
Dim NDX() As Long, i As Long
ReDim NDX(DocParagraphs) As Long
If KeyStart < 0 Then
mKeystart = 1
Else
mKeystart = KeyStart
End If
NDX(0) = 0
For i = 1 To UBound(NDX())
NDX(i) = DocParaNext(NDX(i - 1))
Next i
If order1 <= 0 Then order1 = 1
If order2 <= 0 Then order2 = i - 1
QuickSortExtended NDX, order1, order2

End Sub
Public Function Advance(cursor As Long)
If cursor < 0 Then cursor = 0
If cursor > docpara Then cursor = 0: Exit Function
cursor = DocParaNext(cursor)
Advance = cursor
End Function
Public Function BackStep(cursor As Long)
If cursor < 0 Then cursor = 0
If cursor > docpara Then cursor = 0: Exit Function
cursor = DocParaBack(cursor)
BackStep = cursor
End Function

Private Function GetDosPath(LongPath As String) As String

Dim s As String
Dim i As Long
Dim PathLength As Long

        i = Len(LongPath) * 2 + 2

        s = String(1024, 0)

        PathLength = GetShortPathName(StrPtr(LongPath), StrPtr(s), i)

        GetDosPath = Left$(s, PathLength)

End Function
Private Function NeoUnicodeFile(FileName$) As Boolean
Dim hFile
Dim f$
Const FILE_ATTRIBUTE_NORMAL = &H80
Const GENERIC_WRITE = &H40000000
If Left$(FileName$, 2) <> "\\" Then

f$ = "\\?\" + FileName$
Else
f$ = FileName$
End If
On Error Resume Next
hFile = CreateFile(StrPtr(f$), GENERIC_WRITE, ByVal 0, ByVal 0, 2, FILE_ATTRIBUTE_NORMAL, ByVal 0)
''Sleep 1
CloseHandle hFile
NeoUnicodeFile = (GetDosPath(f$) <> "")
End Function
Public Function SaveUnicodeOrAnsi(ByVal FileName As String, Optional mode2save As Long = 0, Optional Append As Boolean = False) As Boolean
' using doc as extension you can read it from word...with automatic conversion to unicode
' OVERWRITE ALWAYS
Dim w As Long, a() As Byte, f$, i As Long, bb As Byte, yesswap As Boolean, lf$
Const Utf8CodePage As Long = 65001
On Error GoTo t12345
'' 0 vbcrlf, 1 vblf, 2 vbcr
lf$ = choose(ListLoadedType \ 10 + 1, vbCrLf, vbLf, vbCr)
If mode2save \ 10 > 0 Then
' overide loaded status
lf$ = choose(mode2save \ 10 + 1, vbCrLf, vbLf, vbCr)
End If
mode2save = mode2save Mod 10
If Not Append Then
If Not NeoUnicodeFile(FileName) Then Exit Function
End If
f$ = GetDosPath(FileName)
If Err.Number > 0 Or f$ = "" Then Exit Function
w = FreeFile
Open f$ For Binary As w
' mode2save
' 0 is utf-le
If Append Then Seek #w, LOF(w) + 1
mode2save = mode2save Mod 10
Dim maxmw As Long, iPos As Long
 Dim bbBuf() As Byte, iLen As Long



If mode2save = 0 Then
 a() = ChrW(&HFEFF)
If Not Append Then Put #w, , a()

ElseIf mode2save = 1 Then
a() = ChrW(&HFFFE) ' big endian...need swap
If Not Append Then Put #w, , a()
yesswap = True
ElseIf mode2save = 2 Then  'utf8
If Not Append Then
        Put #w, , CByte(&HEF)
        Put #w, , CByte(&HBB)
        Put #w, , CByte(&HBF)
        End If
     ''   Put #W, , Utf16toUtf8(buf)
            iPos = DocParaNext(0)
            maxmw = 327680
        While iPos <> 0
            a() = textArray(lf$, iPos, maxmw)
            iLen = WideCharToMultiByte(Utf8CodePage, 0, VarPtr(a(0)), CLng(UBound(a) - LBound(a) + 1) \ 2, 0, 0, 0, 0)
            ReDim bbBuf(0 To iLen - 1)
            iLen = WideCharToMultiByte(Utf8CodePage, 0, VarPtr(a(0)), CLng(UBound(a) - LBound(a) + 1) \ 2, VarPtr(bbBuf(0)), iLen, 0, 0)
            Put #w, , bbBuf()
        Wend
     
     
        Close w
    SaveUnicodeOrAnsi = True
        Exit Function
ElseIf mode2save = 3 Then ' ascii
            iPos = DocParaNext(0)
            maxmw = 327680
            While iPos <> 0
                Put #w, , textString(lf$, iPos, maxmw)
            Wend
            Close w
            SaveUnicodeOrAnsi = True
            Exit Function
End If


iPos = DocParaNext(0)
maxmw = 327680
If yesswap Then
While iPos <> 0
a() = textArray(lf$, iPos, maxmw)
For i = 0 To UBound(a()) - 1 Step 2
bb = a(i): a(i) = a(i + 1): a(i + 1) = bb
Next i
Put #w, , a()
Wend
Else
While iPos <> 0
a() = textArray(lf$, iPos, maxmw)
Put #w, , a()
Wend
End If
Close w
SaveUnicodeOrAnsi = True
t12345:
End Function
Public Sub ReadUnicodeOrANSI(FileName As String, Optional Merge As Boolean = False, Optional feedback As Long)
Dim i&, FNr&, BLen&, WChars&, BOM As Integer, BTmp As Byte, b() As Byte, lbt As Long
Dim mLof As Long
' code from Schmidt, member of vbforums
' this code changed for Document Class
' Change to read line by line, so no two copies needed
If FileName = "" Then Exit Sub
If Not Merge Then EmptyDoc
ReDim Buffer(1 To 1) As Byte
bufOffset = 0
bufsize = 0
SeekBuf = 1
On Error GoTo ErrHandler
  BLen = FileLen(GetDosPath(FileName))
  If BLen = 0 Then Exit Sub
  
  FNr = FreeFile
  Open GetDosPath(FileName) For Binary Access Read As FNr
  
    Get FNr, , BOM
    Select Case BOM
      Case &HFEFF, &HFFFE 'one of the two possible 16 Bit BOMs
        If BLen >= 3 Then
         NewSeek FNr, 3
          ' now we can read line by line
           mLof = LOF(FNr)
           While Not (mLof < SeekBuf)
           Me.AppendParagraphOneLine getUniStringlINE(FNr, mLof, BOM = &HFEFF, lbt)
          Wend
          feedback = lbt
        End If
      Case &HBBEF 'the start of a potential UTF8-BOM
        Get FNr, , BTmp
        If BTmp = &HBF Then 'it's indeed the UTF8-BOM
        
          If BLen >= 4 Then
           NewSeek FNr, 4
          ' now we can read line by line
           mLof = LOF(FNr)
           While Not (mLof < SeekBuf)
          AppendParagraphOneLine getUniStringLineUtF8(FNr, mLof, lbt)
          Wend
          End If
         feedback = 2 + lbt
        Else
       NewSeek FNr, 1
       mLof = LOF(FNr)
        While Not (mLof < SeekBuf)
           AppendParagraphOneLine getAnsiStringlINE(FNr, mLof, lbt)
          Wend
          feedback = 3 + lbt
        End If
        
      Case Else
      
       NewSeek FNr, 1
       mLof = LOF(FNr)
        While Not (mLof < SeekBuf)
           AppendParagraphOneLine getAnsiStringlINE(FNr, mLof, lbt)
          Wend
          feedback = 3 + lbt
    End Select
    
    
ErrHandler:
If FNr Then Close FNr
ReDim Buffer(1 To 1) As Byte
bufOffset = 0
bufsize = 0
SeekBuf = 1
ListLoadedType = lbt \ 10
If Err Then MyEr Err.Description, Err.Description
End Sub


Public Property Get WordCount() As Long
 Dim mycur As Long
mycur = DocParaNext(0)
Do While mycur <> 0
WordCount = WordCount + mWordCount(TextParagraph(mycur))
mycur = DocParaNext(mycur)
Loop
End Property
Private Function mWordCount(ByVal s$) As Long
Dim i&, l As Long
Static p2 As Long, P1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 - 2
  For i = p2 To p2 + l * 2 Step 2
  GetMem2 i, P1
  Select Case P1
    Case 65 To 90, 97 To 122, Is > 127
    If i - p4 = 2 Then mWordCount = mWordCount + 1
    Case 39
    If i - p4 = 2 Then p4 = i
    Case Else
    p4 = i
  End Select
  Next i
End Function
Private Function mWordMap(s$) As String
Dim i&, l As Long
Static p2 As Long, P1 As Integer, p3 As Long, p4 As Long
Dim ss$, pos As Long
pos = -2
  l = Len(s): If l = 0 Then mWordMap = " ": Exit Function

  p2 = StrPtr(s$): l = l - 1
  p3 = p2
  p4 = p2 - 2
  For i = p2 To p2 + l * 2 Step 2
  GetMem2 i, P1
  Select Case P1
    Case 65 To 90, 97 To 122, Is > 127
    If i - p4 = 2 Then p2 = i: pos = pos + 2
    Case 39
    If i - p4 = 2 Then p4 = i
    Case Else
    If i - p4 > 2 Then
            CopyMemory ByVal p3 + pos, ByVal p2, i - p2
            pos = pos + i - p2
            PutMem2 p3 + pos, 32   ' place a space...
            
    End If
    p4 = i
  End Select
  Next i
  If i - p4 > 2 Then
        CopyMemory ByVal p3 + pos, ByVal p2, i - p2
        pos = pos + i - p2
  End If
  pos = pos / 2
  If pos <= 0 Then
  mWordMap = " "
  Else
  mWordMap = Left$(s$, pos)
  End If

End Function

Public Sub CompactLines(Optional addtimes As Boolean = True)
Dim mycur As Long, nextCur As Long, s$, many As Long
mycur = DocParaNext(0)
Do While mycur <> 0
many = 1
s$ = TextParagraph(mycur)
nextCur = DocParaNext(mycur)
Do While nextCur <> 0
If StrComp(s$, TextParagraph(nextCur), vbTextCompare) = 0 Then
many = many + 1
RemoveDocParaIndex nextCur
nextCur = DocParaNext(mycur)
Else
Exit Do
End If
Loop
If addtimes Then TextParagraph(mycur) = TextParagraph(mycur) & " - " & CStr(many)
mycur = nextCur
Loop

End Sub
Public Function UniqueWords(Optional uLCID As Long = -1, Optional nowrap As Boolean = True) As Long
If uLCID = -1 Then uLCID = mlcid
 ResetWords
 If nowrap Then
 words$ = textDoc1LINE
 Else
words$ = textDoc
End If
Unique = CountWords2(words$, uLCID)
UniqueWords = Unique
End Function
Private Function CountWords2(a$, uLCID As Long) As Long
Dim l As Long
Const NORM_IGNORECASE = &H1
Dim top As Long
l = Len(a$)
If l = 1 Then
CountWords2 = -(Trim(a$) <> "")
top = CountWords2
If top > 0 Then
ReDim pos(1 To 1) As Long, SLen(1 To 1) As Long, hit(1 To 1) As Long
pos(1) = StrPtr(a$)
SLen(1) = 1
End If
Else
Dim p5 As Long
p5 = 512 '' L \ 32 + 1
ReDim pos(1 To p5) As Long, SLen(1 To p5) As Long, hit(1 To p5) As Long
Dim i As Long, k As Long
Dim p2 As Long, p4 As Long, p6 As Long, bb As Integer
p2 = StrPtr(a$)
p3 = p2
Dim md As Long, mda As Long, mtop As Long
l = p2 + 2 * l
p4 = p2
Do
         If p4 >= l Then
         If p4 <> p2 Then GoTo there1
        CountWords2 = top: Exit Function
        End If
            GetMem2 p4, bb
            Select Case bb
            Case 65 To 90, 97 To 122, Is > 127
            Case 39
            If p4 = p2 Then p2 = p4 + 2
            Case Else
            If p4 <> p2 Then
there1:
                k = (p4 - p2) \ 2
                mtop = top
                mda = 0
                md = (mtop - mda + 1) \ 2
                p6 = mda + md
                Do While md > 0
                    Select Case CompareString(uLCID, NORM_IGNORECASE, p2, k, pos(p6), SLen(p6))
                        Case 1
                            mtop = mtop - md
                            md = (mtop - mda + 1) \ 2
                        Case 3
                            mda = mda + md
                            md = (mtop - mda + 1) \ 2
                        Case 2
                            hit(p6) = hit(p6) + 1
                            Exit Do
                    End Select
                    p6 = mda + md
                Loop
                If md = 0 Then
                If top = p5 Then
                p5 = p5 * 2
                ReDim Preserve pos(1 To p5) As Long, SLen(1 To p5) As Long, hit(1 To p5) As Long
                End If
                    top = top + 1
                    
                    If mda = top - 1 Then
                        pos(top) = p2
                        SLen(top) = k
                    Else
                        CopyMemory ByVal VarPtr(pos(1)) + mda * 4 + 4, ByVal VarPtr(pos(1)) + mda * 4, (top - mda - 1) * 4
                        CopyMemory ByVal VarPtr(SLen(1)) + mda * 4 + 4, ByVal VarPtr(SLen(1)) + mda * 4, (top - mda - 1) * 4
                        CopyMemory ByVal VarPtr(hit(1)) + mda * 4 + 4, ByVal VarPtr(hit(1)) + mda * 4, (top - mda - 1) * 4
                        pos(mda + 1) = p2
                        SLen(mda + 1) = k
                        hit(mda + 1) = 0
                    End If
                End If
            End If
            p2 = p4 + 2
            End Select
            p4 = p4 + 2
  Loop
CountWords2 = top
End If
End Function
Private Function MakeString(copythat As Long, ByVal lSize As Long) As String
MakeString = Space$(lSize)
CopyMemory ByVal StrPtr(MakeString), ByVal copythat, lSize * 2
End Function
Function ExportWords(Optional usenumbers As Boolean = True) As String()
Dim a() As String
If Unique = 0 Then
ExportWords = Split(mWordMap(textDoc))
Exit Function
Else
ReDim a(1 To Unique)
Dim i As Long
For i = 1 To Unique
a(i) = MakeString(pos(i), SLen(i)) & " (" & CStr(hit(i) + 1) & ")"
Next i
ExportWords = a()
End If
End Function
Sub ExportWords2ObjectAddItem(ThisObject As Object, Optional usenumbers As Boolean = True)
Dim a() As String, i As Long

If Unique = 0 Then
a() = Split(mWordMap(textDoc))
QuickSortString a(), LBound(a()), UBound(a())
For i = LBound(a()) To UBound(a())
ThisObject.additem a(i)
Next i
Exit Sub
Else
If usenumbers Then
For i = 1 To Unique
ThisObject.additem MakeString(pos(i), SLen(i)) & " (" & CStr(hit(i) + 1) & ")"
Next i
Else
For i = 1 To Unique
ThisObject.additem MakeString(pos(i), SLen(i))
Next i

End If

End If
End Sub
Public Sub ResetWords()
words$ = ""
p3 = 0
ReDim pos(1 To 1) As Long, SLen(1 To 1) As Long, hit(1 To 1) As Long
Unique = 0
End Sub
Public Sub QuickSortString(Arr() As String, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
Dim Piv As String, Tmp As String '<- adjust types here, when switching to something different than Long
  If UB - LB = 1 Then
     M1 = LB

     If StrComp(Arr(M1), Arr(UB), vbTextCompare) = 1 Then Tmp = Arr(M1): Arr(M1) = Arr(UB): Arr(UB) = Tmp
     Exit Sub
  Else
    M1 = (LB + UB) \ 2
 
    If StrComp(Arr(M1), Arr(LB), vbTextCompare) = 0 Then
      M2 = UB - 1
      M1 = LB
      Do
        M1 = M1 + 1
        If M1 > M2 Then
          If StrComp(Arr(UB), Arr(LB), vbTextCompare) = -1 Then Tmp = Arr(LB): Arr(LB) = Arr(UB): Arr(UB) = Tmp
          Exit Sub
        End If

      Loop Until StrComp(Arr(M1), Arr(LB), vbTextCompare) <> 0
      Piv = Arr(M1)
      If M1 > LB Then
        If StrComp(Arr(LB), Piv, vbTextCompare) = 1 Then Arr(M1) = Arr(LB): Arr(LB) = Piv: Piv = Arr(M1)
      End If
    Else
      Piv = Arr(M1)
      M1 = LB
      Do While StrComp(Arr(M1), Piv, vbTextCompare) = -1: M1 = M1 + 1: Loop
    End If
  End If
  
  M2 = UB
  Do
    Do While StrComp(Arr(M2), Piv, vbTextCompare) = 1: M2 = M2 - 1: Loop

    If M1 <= M2 Then
      Tmp = Arr(M1): Arr(M1) = Arr(M2): Arr(M2) = Tmp 'swap
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While StrComp(Arr(M1), Piv, vbTextCompare) = -1: M1 = M1 + 1: Loop

  Loop
  If LB < M2 Then QuickSortString Arr, LB, M2
  If M1 < UB Then QuickSortString Arr, M1, UB
End Sub
Private Function getAnsiStringlINE(f As Long, thislof As Long, Optional linebreaktype As Long) As String
Dim a() As Byte, s1 As String, ss As Long, lbreak As Byte, check As Byte
Dim se As Long, st1 As Long
On Error GoTo a11
Do While Not (thislof < SeekBuf)
    GetByte f, SeekBuf, check
    If check <> 10 And check <> 13 Then
     getAnsiStringlINE = getAnsiStringlINE + Chr(check)
    Else
        If Not (thislof < SeekBuf) Then
            ss = SeekBuf
            lbreak = check
            GetByte f, SeekBuf, check
            If check <> 10 And check <> 13 Or lbreak = check Then
                NewSeek f, ss  ' restore it
            If check = 10 Then linebreaktype = 10 Else linebreaktype = 20
            End If
            
        End If
        Exit Do
    End If
Loop
''Debug.Print getAnsiStringlINE
Exit Function
a11:
On Error GoTo 0
SeekBuf = thislof + 1
End Function
Private Function getUniStringlINE(f As Long, thislof As Long, Optional LEndian As Boolean = True, Optional linebreaktype As Long) As String
Dim a() As Byte, s1 As String, ss As Long, lbreak As String, Tmp As Byte
a = " "
Dim p As Integer
On Error GoTo a12
Do While Not (thislof < SeekBuf)
    ''Get #f, , a()
    GetWord f, SeekBuf, p
    s1 = ChrW(p)
    If Not LEndian Then
    a() = s1
    Tmp = a(0): a(0) = a(1): a(1) = Tmp
    s1 = a()
    End If
    
    If s1 <> vbCr And s1 <> vbLf Then
        getUniStringlINE = getUniStringlINE + s1
     
    Else
    ' linebreal mybe is a CR or LF or CRLF (or LFCR???)
        If Not (thislof < SeekBuf) Then
            ss = SeekBuf
            lbreak = s1
                GetWord f, SeekBuf, p
                s1 = ChrW(p)
                If Not LEndian Then
                a() = s1
                Tmp = a(0): a(0) = a(1): a(1) = Tmp
                s1 = a()
                End If
            If s1 <> vbCr And s1 <> vbLf Or lbreak = s1 Then
                  NewSeek f, ss
            If s1 = vbLf Then linebreaktype = 10 Else linebreaktype = 20
            End If
        End If
        Exit Do
    End If
Loop
Exit Function
a12:
On Error GoTo 0
SeekBuf = thislof + 1
End Function
Private Function getUniStringLineUtF8(f As Long, thislof As Long, Optional linebreaktype As Long) As String

Dim ss As Long, lbreak As Byte, check As Byte
Dim st As Long, WChars&
Dim shortbuf() As Byte
ReDim shortbuf(1 To 200) As Byte
On Error GoTo a13
Do While Not (thislof < SeekBuf)
tryagain:
    If thislof < SeekBuf Then Exit Do
    GetByte f, SeekBuf, check

    If (check And &H80) <> 0 Then
        st = st + 1
    If st > UBound(shortbuf) Then ReDim Preserve shortbuf(1 To UBound(shortbuf) * 2) As Byte
    shortbuf(st) = check
    GoTo tryagain
    End If
    If check <> 10 And check <> 13 Then
        st = st + 1
    If st > UBound(shortbuf) Then ReDim Preserve shortbuf(1 To UBound(shortbuf) * 2) As Byte
    shortbuf(st) = check
    Else
        If Not (thislof < SeekBuf) Then
            ss = SeekBuf
            lbreak = check
              GetByte f, SeekBuf, check
            If (check And &H80) <> 0 Then GoTo notryother
            
            If (check <> 10) And (check <> 13) Or (lbreak = check) Then
notryother:

                NewSeek f, ss  ' restore it
                If check = 10 Then linebreaktype = 10 Else linebreaktype = 20
     
            End If
            
        End If
        Exit Do
    End If
Loop
If st > 0 Then
            WChars = MultiByteToWideChar(65001, 0, shortbuf(1), st, 0, 0)
            getUniStringLineUtF8 = Space$(WChars)
            MultiByteToWideChar 65001, 0, shortbuf(1), st, StrPtr(getUniStringLineUtF8), WChars

End If
Exit Function
a13:
On Error GoTo 0
SeekBuf = thislof + 1
End Function
Sub EmptyDoc()
delDocPara = 0
docpara = 0
docmax = 20  '
ReDim para(1 To docmax)
ReDim DocParaBack(0 To docmax)
ReDim DocParaNext(0 To docmax)
End Sub
Sub FillBuffer(f As Long)
Dim c As Long
If bufsize = 0 Then
bufsize = LOF(f) Mod 327680
If bufsize = 0 Then bufsize = LOF(f)
End If
c = LOF(f) - SeekBuf + 1
If UBound(Buffer()) > 1 Then bufOffset = bufOffset + UBound(Buffer())
If c > 327680 Then
        ReDim Buffer(1 To 327680) As Byte
        Get f, , Buffer
ElseIf c > 0 Then
        ReDim Buffer(1 To c) As Byte
        Get f, , Buffer
End If
End Sub
Sub NewSeek(f As Long, ByVal location As Long)
If bufsize = 0 Then
bufsize = LOF(f) Mod 327680
If bufsize = 0 Then bufsize = LOF(f) Else bufsize = 327680
GoTo fillmeplease
End If
If location - bufOffset > 0 And location <= bufsize + bufOffset Then
Else
fillmeplease:
Seek #f, (location \ bufsize) * bufsize + 1
FillBuffer f
End If
SeekBuf = location
End Sub
Sub GetByte(f As Long, ByVal location As Long, b As Byte)
If location > UBound(Buffer) + bufOffset Then
Seek #f, (location \ bufsize) * bufsize + 1
FillBuffer f
End If
b = Buffer(location - bufOffset)
SeekBuf = location + 1
End Sub
Sub GetWord(f As Long, ByVal location As Long, b As Integer)
If location > UBound(Buffer) + bufOffset Then
Seek #f, (location \ bufsize) * bufsize + 1
FillBuffer f
End If
GetMem2 VarPtr(Buffer(location - bufOffset)), b
SeekBuf = location + 2
End Sub
Function InstrTextCompareAnyLocale(position As Long, from$, this$, Optional uLCID As Long = 1033) As Long

Const NORM_IGNORECASE = &H1
Dim P1 As Long, p2 As Long, l1 As Long, l2 As Long, cu As Long
l1 = (Len(from$) - Len(this$)) * 2 - (position - 1) * 2
If l1 < 0 Then Exit Function
P1 = StrPtr(from$) + (position - 1) * 2
p2 = StrPtr(this$)
l2 = Len(this$)
For cu = P1 To P1 + l1 Step 2
If CompareString(uLCID, NORM_IGNORECASE, cu, l2, p2, l2) = 2 Then
InstrTextCompareAnyLocale = (cu - StrPtr(from$)) \ 2 + 1
Exit Function
End If
Next cu
End Function
Function InstrRevTextCompareAnyLocale(from$, this$, Optional position As Long = -1, Optional uLCID As Long = 1033) As Long

Const NORM_IGNORECASE = &H1
Dim P1 As Long, p2 As Long, l1 As Long, l2 As Long, cu As Long
If Len(from$) < position Then position = -1
If position > 0 Then
l1 = (Len(from$) - Len(this$)) * 2 - (Len(from$) - position) * 2
Else
l1 = (Len(from$) - Len(this$)) * 2
End If
If l1 < 0 Then Exit Function
P1 = StrPtr(from$)
p2 = StrPtr(this$)
l2 = Len(this$)
For cu = P1 + l1 To P1 Step -2
If CompareString(uLCID, NORM_IGNORECASE, cu, l2, p2, l2) = 2 Then
InstrRevTextCompareAnyLocale = (cu - P1) \ 2 + 1
Exit Function
End If
Next cu
End Function

Public Property Get LCID() As Long
LCID = mlcid
End Property

Public Property Let LCID(ByVal rhs As Long)
mlcid = rhs
End Property
Function FindWord(Find$, Direction As Boolean, there As Long, cur As Long) As Boolean
Dim oncemore As Boolean, result As Boolean, res$
again:
If Direction Then
 result = FindStrDown(Find$, there, cur)
 Else
 result = FindStrUp(Find$, there, cur)
 End If
If result Then
res$ = TextParagraph(there)
If cur > 1 Then
    Select Case AscW(Mid$(res$, cur - 1, 1))
    Case 39, 65 To 90, 97 To 122, Is > 127
    If Direction Then cur = cur + 1 Else cur = cur - 1
    GoTo again
    Case Else
    End Select
End If
If (Len(res$) - cur + 1) > Len(Find$) Then
    Select Case AscW(Mid$(res$, cur + Len(Find$), 1))
    Case 39, 65 To 90, 97 To 122, Is > 127
    If Direction Then cur = cur + 1 Else cur = cur - 1
    GoTo again
    Case Else
    End Select
End If
''
FindWord = True
Exit Function
Else
If oncemore Then Exit Function
If Direction Then
there = 1
cur = 1
Else
there = DocParagraphs
cur = TextParagraphLen(ParagraphFromOrder(there))
End If
oncemore = True
End If
End Function
Function FindIdentifier(Find$, Direction As Boolean, there As Long, cur As Long) As Boolean
' there is paragraph number not paragraph order
Dim oncemore As Boolean, result As Boolean, res$
If Len(Find$) > 1 And Left$(Find$, 1) = "&" Then Find$ = Mid$(Find$, 2): cur = cur + 1
If InStr(Find$, Chr$(13)) > 0 Then Exit Function
If InStr(Find$, Chr$(10)) > 0 Then Exit Function
again:
If Direction Then
 result = FindStrDown(Find$, there, cur)
 Else
 result = FindStrUp(Find$, there, cur)
 End If
If result Then
res$ = TextParagraph(there)

If cur > 1 Then
    Select Case AscW(Mid$(res$, cur - 1, 1))
    Case 46, 65 To 90, 95, 97 To 122, Is > 127
    If Direction Then cur = cur + 1 Else cur = cur - 1
    GoTo again
    Case Else
    End Select
End If
Select Case AscW(Right$(Find$, 1))
Case 36, 37, 40
Case Else
If (Len(res$) - cur + 1) > Len(Find$) Then
    Select Case AscW(Mid$(res$, cur + Len(Find$), 1))
    Case 40, 46, 48 To 57, 65 To 90, 95, 97 To 122, Is > 127
    If Direction Then cur = cur + 1 Else cur = cur - 1
    GoTo again
    Case Else
    End Select
End If
End Select
''
FindIdentifier = True
Exit Function
Else
If oncemore Then Exit Function
If Direction Then
there = 1
cur = 1
Else
there = DocParagraphs
If there <> 0 Then
cur = TextParagraphLen(ParagraphFromOrder(there))
End If
End If
oncemore = True
End If
End Function
Public Function NextTextParagraph(ByVal Paranum As Long, nextparanum As Long) As String
If Not IsLast(Paranum) Then
nextparanum = DocParaNext(Paranum)
NextTextParagraph = TextParagraph(nextparanum)
End If
End Function





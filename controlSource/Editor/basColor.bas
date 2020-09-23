Attribute VB_Name = "basColor"
Option Explicit

Rem ==========APIs=============
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long

Rem ==========CONSTANTs=============

Private Const COL_RED = 2426302
Private Const COL_WATERMILENRED = 3540170
Private Const COL_BLUE = 11158056
Private Const COL_GREEN = 3651889
Private Const COL_ORANGE = 1945598
Private Const COL_VIOLET = 10435956
Private Const COL_TABLEGREEN = 8639048
Private Const COL_PINK = 16012788
Private Const COL_AQUA = 10661419
Private Const COL_GRAY = 11451322
Private Const COL_LIGHTPINK = 16360695


Private Const HTMLMAIN = " IMG BASE TITLE HEAD META LINK OBJECT PARAM A HR DIV BR BODY HTML SPAN LAYER NOLAYER BGSOUND "
Private Const HTMLTABLE = " TABLE TBODY TD TH THEAD TFOOT TR TT OL UL LI DL DT DD COL COLGROUP CAPTION MULTICOL "
Private Const HTMLTEXT = " BASEFONT CENTER SUB SUP FONT H1 H2 H3 H4 H5 H6 MARQUEE SMALL SPACER STRIKE STRONG PLAINTEXT PRE P U I B BIG BLOCKQUOTE BLING EM "
Private Const HTMLMISC = " COMMENT MENU MOBR ISINDEX VAR WBR XMP MAP KEYGEN EMBED NOEMBED DEL DFN DIR INS BDO ABBR AREA APPLET ADDRESS ACRONYM CITE CODE LISTING KBD Q S SAMP LEGEND "
Private Const HTMLFORM = " NOFRAMES STYLE BUTTON FORM LABEL TEXTAREA INPUT SELECT OPTION IFRAME ILAYER FRAME FRAMESET FIELDSET OPTGROUP "
Private Const HTMLSCRIPT = " NOSCRIPT SCRIPT SERVER "
Private Const ASPWORDS = " OPTION EXPLICIT SET AS IS END DIM REDIM PUBLIC SUB BYVAL IF THEN ELSE PRIVATE FOR NEXT TO EXIT DO LOOP WHILE UNTIL ENDIF NOTHING NULL SELECT CASE LCASE FUNCTION AND UCASE LBOUND UBOUND XOR OR EMPTY REM "
Private Const ASPOBJECTS = " RESPONSE ADDHEADER APPENDTOLOG BINARYWRITE BUFFER CACHECONTROL CHARSET CLEAR CONTENTTYPE COOKIES END EXPIRES EXPIRESABSOLUTE FLUSH ISCLIENTCONNECTED PICS REDIRECT STATUS WRITE REQUEST BINARYREAD CLIENTCERTIFICATE COOKIES FORM ITEM QUERYSTRING SERVERVARIABLES TOTALBYTES SESSION ABANDON CODEPAGE CONTENTS LCID SESSIONID STATICOBJECTS TIMEOUT VALUE APPLICATION CONTENTS LOCK STATICOBJECTS UNLOCK VALUE SERVER CREATEOBJECT HTMLENCODE MAPPATH SCRIPTTIMEOUT URLENCODE URLPATHENCODE SCRIPTINGCONTEXT APPLICATION REQUEST RESPONSE SERVER SESSION COUNT ITEMKEY "
Private Const ASPMISC = " LANGUAGE CODEPAGE "

Public Const WM_SETREDRAW = &HB
Public Const WM_USER = &H400
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_HIDESELECTION = (WM_USER + 63)
Public Const EM_GETLINECOUNT = &HBA

Const CHAR_COMMENT = "'"

Type WORD_TYPE
    Text As String  ' word to be colored
    Color As Long   ' color to color the word :)
End Type

Type LETTER_TYPE
    Start As Long   ' first time the letter appears in the list
    Finish As Long  ' last time the letter appears in the list
End Type

Type POINTAPI
    X As Long
    y As Long
End Type


'#--------------------------------------------------------------------------#
'#  variables
'#--------------------------------------------------------------------------#
Dim Words() As WORD_TYPE
Dim htmlWords() As WORD_TYPE
Dim Letters() As LETTER_TYPE
Dim htmlLetters() As LETTER_TYPE
Dim Strings() As String
Dim ColorColl As Collection
Public sText As String
Public Declare Function GetCaretPos Lib "User32" (lpPoint As POINTAPI) As Long
Public MDoColor As Boolean



'//--[DoColor]--------------------------------------------------------------//
'
'  Here it is - the beast itself. This routine colors
'  a single line of text within the RTB. It will
'  split each line up into words using the custom
'  split function (SplitWords), then match each word
'  against the list of keywords.
'
Public Sub DoColor(RTB As Richtextbox, ByVal lStart As Long, ByVal lFinish As Long, Optional ByVal pDoColor As Boolean = True)
Dim sWords()    As String
Dim sLine       As String
Dim sChar       As String
Dim lCurPos     As Long
Dim lIndex      As Long
Dim lColor      As Long
Dim lPos        As Long
Dim Lpos2       As Long
Dim lCom        As Long
Dim i           As Long
Dim html        As String
    
    If Not MDoColor Then Exit Sub
    ' grab the line
    sLine = Trim$(Mid$(sText, lStart, lFinish - lStart))
    ' remove the EOL
    sLine = RemoveEOL(sLine)
    ' remove the quotes so they're not colored
    sLine = RemoveStrings(sLine)
    
    ' split the line into words using our custom function
    sWords = SplitWords(sText, sLine, lFinish)
    
    ' check each word against the list
    lCurPos = 1
    html = ""
    ' search for each word in the array
    For i = LBound(sWords) To UBound(sWords)
        sWords(i) = Trim(sWords(i))
        sWords(i) = Replace(sWords(i), Chr(10), "")
        sWords(i) = Replace(sWords(i), Chr(13), "")
        sWords(i) = Replace(sWords(i), Chr(9), "")
        If Trim$(sWords(i)) <> "" Then
          If Left$(sWords(i), 1) <> "'" Then sWords(i) = Replace(sWords(i), Chr(9), "")
          If IsWithinASP(sText, sWords(i), InStr(lStart, sText, sWords(i))) Or Trim(sWords(i)) = "<%" Or Trim(sWords(i)) = "%>" Then
            ' check for comment in the middle of a line
            If Left$(sWords(i), 1) = CHAR_COMMENT Then
              
                ' color the rest of the line
                If InStr(lStart, sText, sWords(i)) > 0 Then
                  
                    RTB.SelStart = InStr(lStart, sText, sWords(i)) - 1
                    RTB.SelLength = Len(sWords(i))
                    RTB.SelColor = COL_GREEN
                End If
            
            ElseIf sWords(i) = "<%" Or sWords(i) = "%>" Then
                lColor = COL_AQUA
                RTB.SelStart = InStr(lStart + lCurPos - 1, sText, sWords(i)) - 1
                RTB.SelLength = Len(sWords(i))
                RTB.SelColor = lColor
                'HighLightWord RTB, sWords(i), vbYellow
            Else
                If sWords(i) = "ByVal" Or sWords(i) = "Byte" Then
                    DoEvents: End If
                sChar = Left$(LCase$(sWords(i)), 1)
                If sChar = "=" Then
                  sWords(i) = Mid(sWords(i), 2)
                  sChar = Left$(LCase$(sWords(i)), 1)
                End If
                lColor = GetColor(sWords(i), True)
                If lColor Then
                    RTB.SelStart = InStr(lStart + lCurPos - 1, sText, sWords(i)) - 1
                    RTB.SelLength = Len(sWords(i))
                    RTB.SelColor = lColor
                Else
                    RTB.SelStart = InStr(lStart + lCurPos - 1, sText, sWords(i)) - 1
                    RTB.SelLength = Len(sWords(i))
                    RTB.SelColor = vbBlack
                End If
            End If
          Else
            html = html & sWords(i) & " "
          End If
        End If
        ' move the current position within the line on
        lCurPos = lCurPos + Len(sWords(i))
        
    Next i
    DoColorHTML RTB, sText, html, lStart, lFinish
End Sub

'//--[DoClipBoardPaste]-----------------------------------------------------//
'
'  Call this when text has been pasted into the
'  RTB. It will grab the text, split it into lines
'  and color it.
'
Public Sub DoClipBoardPaste(RTB As Richtextbox)
Dim lCursor As Long
Dim lStart As Long
Dim lFinish As Long
Dim sText As String
Dim lFinish1 As Long
    On Error Resume Next
    ' store the cursor position
    lCursor = RTB.SelStart
    
    ' add the text and color it
    LockWindowUpdate RTB.hWnd
    
    ' get the text to be pasted from the clipboard
    sText = Clipboard.GetText
    
    ' get the start point - this should be the previous
    ' vbCrLf to where the text was inserted, to make
    ' sure that if it's inserted mid-line, the whole
    ' line is colored
    lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart + 1) + 1
    If lStart = 0 Then lStart = RTB.SelStart
    ' also store the finish point
    lFinish1 = RTB.SelStart + Len(sText)
    
    
    ' now add the text to the box
    RTB.SelText = sText
    RTB.SelColor = vbBlack
    basColor.sText = RTB.Text
    
    Do While lStart < lFinish1 + 1
        ' find the end of this line
        lFinish = InStr(lStart + 1, RTB.Text, vbCrLf)
        If lFinish = 0 Then lFinish = lStart + Len(sText)
            
        ' color it
        DoColor RTB, lStart, lFinish
        
        ' move up to get the next line
        lStart = lFinish + 2
    Loop
    
    ' restore the original values
    RTB.SelStart = lCursor + Len(sText)
    RTB.SelColor = vbBlack
    
    ' null the keypress (to avoid the text pasting twice)
    LockWindowUpdate 0&

End Sub



'//--[SplitWords]---------------------------------------------------------//
'
'  Since splitting a line into words by a single
'  character is not acceptable because we have to
'  take several end of word characters into account,
'  this routine was written.
'  It searches through the string from left to right
'  and locates the nearest word break char from a list
'  then splits at that word.
'
Private Function SplitWords(ByVal pContent As String, ByVal sText As String, ByVal pEnd As Long) As String()
Dim i As Long, lPos As Long
Dim j As Long
Dim sWords() As String
Dim sWordBreaks(0 To 4) As String
Dim lBreakPoints() As Long
Dim sSubWords() As String 'Collection of <%
Dim sSubEWords() As String 'Collection of %>
Dim lBreak As Long
    
    ' list of word break characters
    sWordBreaks(0) = " "
    sWordBreaks(1) = "("
    sWordBreaks(2) = CHAR_COMMENT
    sWordBreaks(3) = "."
    sWordBreaks(4) = "="
    
     ' comments
    ReDim lBreakPoints(UBound(sWordBreaks))

    ' get them words!
    ReDim sWords(0)
    lPos = 1
    Do
    
        ' locate the word break points
        For i = 0 To UBound(sWordBreaks)
            lBreakPoints(i) = InStr(lPos, sText, sWordBreaks(i))
        Next i
        
        ' now work out which is closest
        lBreak = Len(sText) + 1
        For i = 0 To UBound(lBreakPoints)
            If lBreakPoints(i) <> 0 Then
                If lBreakPoints(i) < lBreak Then lBreak = lBreakPoints(i)
            End If
        Next i
    
        ' now split out the word
        ' if no break point was found, then we've
        ' hit the end of the line, so add all the rest
        If lBreak = Len(sText) + 1 Then
            'sWords(UBound(sWords)) = Mid$(sText, lPos) 'Changed
            sSubWords = SplitASPTags(Mid$(sText, lPos))
            For i = LBound(sSubWords) To UBound(sSubWords)
              sWords(UBound(sWords)) = sSubWords(i)
              If i <> UBound(sSubWords) Then
                ReDim Preserve sWords(UBound(sWords) + 1)
              End If
            Next
        Else
            ' add this word - first check for a comment
            'IsWithinASP for comment in asp
            If Mid$(sText, lBreak, 1) = CHAR_COMMENT And IsWithinASP(pContent, Mid$(sText, lPos, lBreak - lPos), InStrRev(Mid(pContent, 1, pEnd), Mid$(sText, lPos, lBreak - lPos))) Then
                ' first add the word
                sWords(UBound(sWords)) = Mid$(sText, lPos, lBreak - lPos)
                ' then add the rest as a comment
                ReDim Preserve sWords(UBound(sWords) + 1)
                sWords(UBound(sWords)) = Mid$(sText, lBreak)
                ' now return and exit
                SplitWords = sWords
                Exit Function
            Else
                'sWords(UBound(sWords)) = Mid$(sText, lPos, lBreak - lPos) 'Changed
                sSubWords = SplitASPTags(Mid$(sText, lPos, lBreak - lPos))
                For i = LBound(sSubWords) To UBound(sSubWords)
                  sWords(UBound(sWords)) = sSubWords(i)
                  If i <> UBound(sSubWords) Then
                    ReDim Preserve sWords(UBound(sWords) + 1)
                  End If
                Next
            End If
        End If
        sWords(UBound(sWords)) = Replace(sWords(UBound(sWords)), Chr(9), "")
        ReDim Preserve sWords(UBound(sWords) + 1)
    
        ' move the pointer on a bit
        lPos = lBreak + 1
        
        ' setup the exit condition
        If lPos >= Len(sText) Then Exit Do
    
    Loop

    ' return the array
    SplitWords = sWords

End Function

'//--[RemoveEOL]------------------------------------------------------------//
'
'  Removes leading and trailing vbCrLf from strings
'
Private Function RemoveEOL(ByVal sText As String) As String
Dim sTmp As String
    ' remove leading or trailing vbCrLf from the string
    sTmp = sText
    If Left$(sTmp, 2) = vbCrLf Then
        sTmp = Right$(sTmp, Len(sTmp) - 2)
    End If
    If Right$(sTmp, 2) = vbCrLf Then
        sTmp = Left$(sTmp, Len(sTmp) - 2)
    End If
    RemoveEOL = sTmp
End Function

'//--[RemoveStrings]-------------------------------------------------------//
'
'  Removes any quoted strings from the text, but only
'  those that aren't within comments of course.
'
Private Function RemoveStrings(ByVal sText As String) As String
Dim lCom As Long
Dim lPos As Long
Dim Lpos2 As Long

    lCom = InStr(1, sText, CHAR_COMMENT)
    lPos = InStr(1, sText, Chr$(34))
    If lPos < lCom Or lCom = 0 Then
        Do While lPos <> 0
            ' find the end " char to make a pair
            Lpos2 = InStr(lPos + 1, sText, Chr$(34))
            If Lpos2 <> 0 Then
                ' we've found a pair, so remove it
                sText = Mid$(sText, 1, lPos - 1) & Mid$(sText, Lpos2 + 1)
                ' find the next starting " avoiding
                ' comments within strings
                lCom = InStr(Lpos2 + 1, sText, CHAR_COMMENT)
                lPos = InStr(Lpos2 + 1, sText, Chr$(34))
                If lPos > lCom Then Exit Do
            Else
                Exit Do
            End If
        Loop
    End If
    
    ' return
    RemoveStrings = sText
    
End Function

'//--[CombSort]------------------------------------------------------------//
'
'  This is a standard comb sort - you could replace
'  this with any other sorting algorithm, I just prefer
'  this one because a) i wrote it :), and b) it performs
'  well across all ranges of input arrays - it makes
'  no assumptions about how sorted the array already
'  is, because it doesn't matter.
'  The comb sort is a slight variation on the bubblesort,
'  and i know what you're thinking - ewwww, bubble sorts -
'  but you'd be wrong, the comb is only fractionally
'  slower than a quicksort... so enjoy!
'  for more on the combsort, read here:
'  http://yagni.com/combsort/index.php
'  http://cs.clackamas.cc.or.us/molatore/cs260Spr01/combsort.htm
'
Private Sub CombSort(Arr() As WORD_TYPE)
Dim i As Long, j As Long, t As WORD_TYPE
Dim swapped As Boolean
Dim gap As Long
   
    gap = UBound(Arr)
    
    Do
        gap = (gap * 10) \ 13
        If gap = 9 Or gap = 10 Then gap = 11
        If gap < 1 Then gap = 1
        
        swapped = False
        For i = 0 To UBound(Arr) - gap
            j = i + gap
            If Arr(i).Text > Arr(j).Text Then
                LSet t = Arr(j)
                LSet Arr(j) = Arr(i)
                LSet Arr(i) = t
                swapped = True
            End If
        Next i
        
        If (gap = 1) And (Not swapped) Then Exit Do
    Loop
    
End Sub

Public Function IsWithinASP(ByVal pContent As String, ByVal pText As String, ByVal pEnd As Long) As Boolean
'
' whether the given word(the word which is going to color) is within the asp code
'
Dim Ltmp As String
Dim Lpos1 As Long
Dim Lpos2 As Long
  On Error Resume Next
  Ltmp = Mid(pContent, 1, pEnd + 1)
  If Ltmp <> "" Then
    Lpos1 = InStrRev(Ltmp, "<%")
    Lpos2 = InStrRev(Ltmp, "%>")
    If Lpos1 > 0 Then
      If Lpos1 > Lpos2 Then
        IsWithinASP = True
      End If
    End If
  End If
End Function

Private Function SplitASPTags(ByVal pText As String) As String()
Dim j As Long
Dim i As Long
Dim sSubWords() As String
Dim sSubEWords() As String
Dim sWords() As String
  ReDim sWords(0)
  'Search for <%
  sSubWords = Split(pText, "<%")
  If LBound(sSubWords) >= 0 Then
    For i = LBound(sSubWords) To UBound(sSubWords)
      'Search For %>
      If sSubWords(i) <> "" Then
        sSubEWords = Split(sSubWords(i), "%>")
        If LBound(sSubEWords) >= 0 Then
          For j = LBound(sSubEWords) To UBound(sSubEWords)
            sWords(UBound(sWords)) = sSubEWords(j)
            'splitted %> should be added for coloring
            If j <> UBound(sSubEWords) Then
              ReDim Preserve sWords(UBound(sWords) + 1)
              sWords(UBound(sWords)) = "%>"
            End If
            ReDim Preserve sWords(UBound(sWords) + 1)
          Next
        End If
      End If
      'splitted <% should be added for coloring
      If i <> UBound(sSubWords) Then
        sWords(UBound(sWords)) = "<%"
        ReDim Preserve sWords(UBound(sWords) + 1)
      End If
    Next
  End If
  SplitASPTags = sWords
End Function

Private Sub DoColorHTML(ByVal RTB As Richtextbox, ByVal pText As String, ByVal pHTML As String, ByVal pStart As Long, pEnd As Long)
'
' Coloring the html tags
'
Dim i As Long
Dim tags() As String
Dim sChar As String
Dim lIndex As Integer
Dim lColor As Long
Dim lPos As Integer
Dim lTag As String
Dim sText As String
  On Error Resume Next
  sText = Mid(pText, 1, pEnd)
  
  If pHTML <> "" Then
    tags = Split(pHTML, ">")
    For i = LBound(tags) To UBound(tags)
      tags(i) = Trim(tags(i))
      If InStr(tags(i), "<") > 0 Then tags(i) = Mid(tags(i), InStr(tags(i), "<"))
      If Left(tags(i), 4) = "<!--" Then
        pEnd = InStr(pStart + 1, pText, "-->")
        RTB.SelStart = pStart - 1
        RTB.SelLength = pEnd - pStart + 3
        RTB.SelColor = COL_AQUA
        pStart = pEnd + 3
      ElseIf Left(tags(i), 1) = "<" Then
        lPos = InStr(tags(i), " ")
        If lPos = 0 Then lPos = Len(tags(i))
        sChar = Mid(LCase(tags(i)), 2, 1)
        lTag = IIf(sChar = "/", Trim(Mid(tags(i), 3, lPos - 1)), Trim(Mid(tags(i), 2, lPos - 1)))
        If sChar = "/" Then sChar = Mid(LCase(tags(i)), 3, 1)
        lColor = GetColor(lTag)
        pEnd = GetCloseTagPos(sText, pStart + 1)
        If lColor Then
            pStart = InStr(pStart, sText, Split(tags(i), " ")(0))
            RTB.SelStart = pStart - 1
            RTB.SelLength = pEnd - pStart + 1
            RTB.SelColor = lColor
        End If
        pStart = pEnd + 1
      End If
    Next
  End If
End Sub

Private Function GetCloseTagPos(ByVal pText As String, ByVal pStart As Long) As Long
'
'Get the last close tag(>); suppose any tag contains asp tags within it.
'ex: <p id="<%=lid%>">
'
Dim lPos As Long
Dim lChar As String
  On Error Resume Next
  lPos = InStr(pStart, pText, ">")
  lChar = Mid(pText, lPos - 1, 1)
  Do Until Trim(lChar) <> "%"
    lPos = InStr(lPos + 1, pText, ">")
    lChar = Mid(pText, lPos - 1, 1)
  Loop
  GetCloseTagPos = lPos
End Function

Private Function GetColor(ByVal Word As String, Optional ByVal ASP As Boolean) As Long
  If ASP Then
    If InStr(1, ASPWORDS, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
      GetColor = COL_BLUE
    ElseIf InStr(1, ASPOBJECTS, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
      GetColor = COL_PINK
    ElseIf InStr(1, ASPMISC, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
      GetColor = COL_WATERMILENRED
    End If
  Else
    If InStr(1, HTMLMAIN, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
      GetColor = COL_BLUE
    ElseIf InStr(1, HTMLTABLE, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
      GetColor = COL_TABLEGREEN
    ElseIf InStr(1, HTMLTEXT, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
      GetColor = COL_ORANGE
    ElseIf InStr(1, HTMLFORM, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
      GetColor = COL_VIOLET
    ElseIf InStr(1, HTMLMISC, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
      GetColor = COL_PINK
    ElseIf InStr(1, HTMLSCRIPT, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
      GetColor = COL_RED
    Else
      GetColor = COL_LIGHTPINK
    End If
  End If
End Function


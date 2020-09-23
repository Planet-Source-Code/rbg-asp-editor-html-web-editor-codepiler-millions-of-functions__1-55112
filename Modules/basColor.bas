Attribute VB_Name = "basColor"
Option Explicit

'#--------------------------------------------------------------------------#
'#  apis, enums, consts, declares
'#--------------------------------------------------------------------------#
' api to stop the window refreshing
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Const COL_KEYWORD = vbBlue '&H800000    ' dark blue
Const COL_TAG = vbBlue '11753728    ' dark blue
Const COL_TABLE = 38400 'GREEN
Const COL_HTML = 150 'RED
Const COL_HTML_PPR = 150 'RED 'html properties
Const COL_SCRIPT = 150 'RED
Const COL_NOCOLOR = 16750335 'PINK
Const COL_COMMENT_HTML = 8092539 'HTML COMMENT
Const COL_COMMENT = &H8000&     ' middle green
Const HTMLTAGS = " HTML TITLE HEAD !DOCTYPE META LINK SCRIPT BODY TABLE TD TR THEAD TBODY B U I STRONG SUB SUP FORM INPUT STYLE IMG EMPED A BR HR H1 H2 H3 H4 H5 H6 "
Const HTMLPROPERTIES = " NAME WIDTH HEIGHT STYLE BACKGROUND BGCOLOR SRC LOOP COLSPAN ROWSPAN ID VALUE ALIGN TYPE CLASS VALIGN READONLY CHECKED SELECTED BORDER CELLSPACING CELLPADDING NOWRAP "
Const HTMLDELIMITER = vbTab & " =<>/!"
Const CHAR_COMMENT = "'"        ' comment line char
Public Const WM_USER = &H400
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Private Const WM_SETREDRAW = &HB
Public Const EM_HIDESELECTION = (WM_USER + 63)

Const RGB_COMMENT As String = "0,128,0"
Const RGB_STRING As String = "0,0,250"
Const RGB_RESERVED As String = "0,0,255"
Const RGB_FUNC_OBJ As String = "255,0,0"
Const RGB_DELIMITER As String = "0,0,0"
Const RGB_TAGDELIMITER As String = "0,0,250"
Const RGB_TAG As String = "150,0,0"
Const RGB_PROPERTY As String = "250,0,0"
Const RGB_NORMAL As String = "0,0,0"

Enum SyntaxTypes
    ColorComment = 0
    ColorString = 1
    ColorReserved = 2
    ColorFuncObj = 3
    ColorDelimiter = 4
    Colornormal = 5
    ColorTagdelimiter = 6
    ColorTag = 7
    ColorProperty = 8
End Enum

' RGB values derived from constants
Private mrgbComment As Long
Private mrgbString As Long
Private mrgbReserved As Long
Private mrgbFuncObj As Long
Private mrgbDelimiter As Long
Private mrgbNormal As Long
Private mrgbTagdelimiter As Long
Private mrgbTag As Long
Private mrgbProperty As Long

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

Public Enum ModifyTypes
    AddText = 0
    DeleteText = 1
    ReplaceText = 2
    CutText = 3
    PasteText = 4
End Enum

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
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public MDoColor As Boolean

'#--------------------------------------------------------------------------#
'#  methods
'#--------------------------------------------------------------------------#

'//--[InitKeyWords]-----------------------------------------------------------//
'
'  Builds the arrays of keywords, then builds
'  an alphabetical index of the array to aid
'  searching later on.
'
Public Sub InitKeyWords()
    ' initialize the array of words
    ReDim Words(0 To 39)
    Words(0).Text = "Option"
    Words(0).Color = COL_KEYWORD
    Words(1).Text = "Explicit"
    Words(1).Color = COL_KEYWORD
    Words(2).Text = "Set"
    Words(2).Color = COL_KEYWORD
    Words(3).Text = "As"
    Words(3).Color = COL_KEYWORD
    Words(4).Text = "Is"
    Words(4).Color = COL_KEYWORD
    Words(5).Text = "End"
    Words(5).Color = COL_KEYWORD
    Words(6).Text = "Dim"
    Words(6).Color = COL_KEYWORD
    Words(7).Text = "ReDim"
    Words(7).Color = COL_KEYWORD
    Words(8).Text = "Public"
    Words(8).Color = COL_KEYWORD
    Words(9).Text = "Sub"
    Words(9).Color = COL_KEYWORD
    Words(10).Text = "ByVal"
    Words(10).Color = COL_KEYWORD
    Words(11).Text = "If"
    Words(11).Color = COL_KEYWORD
    Words(12).Text = "Then"
    Words(12).Color = COL_KEYWORD
    Words(13).Text = "Else"
    Words(13).Color = COL_KEYWORD
    Words(14).Text = "Private"
    Words(14).Color = COL_KEYWORD
    Words(15).Text = "For"
    Words(15).Color = COL_KEYWORD
    Words(16).Text = "Next"
    Words(16).Color = COL_KEYWORD
    Words(17).Text = "To"
    Words(17).Color = COL_KEYWORD
    Words(18).Text = "Exit"
    Words(18).Color = COL_KEYWORD
    Words(19).Text = "Do"
    Words(19).Color = COL_KEYWORD
    Words(20).Text = "Loop"
    Words(20).Color = COL_KEYWORD
    Words(21).Text = "While"
    Words(21).Color = COL_KEYWORD
    Words(22).Text = "Until"
    Words(22).Color = COL_KEYWORD
    Words(23).Text = "EndIf"
    Words(23).Color = COL_KEYWORD
    Words(24).Text = "Nothing"
    Words(24).Color = COL_KEYWORD
    Words(25).Text = "Null"
    Words(25).Color = COL_KEYWORD
    Words(26).Text = "Select"
    Words(26).Color = COL_KEYWORD
    Words(27).Text = "Case"
    Words(27).Color = COL_KEYWORD
    Words(28).Text = "Lcase"
    Words(28).Color = COL_KEYWORD
    Words(29).Text = "Function"
    Words(29).Color = COL_KEYWORD
    Words(30).Text = "And"
    Words(30).Color = COL_KEYWORD
    Words(31).Text = "Ucase"
    Words(31).Color = COL_KEYWORD
    Words(32).Text = "LBound"
    Words(32).Color = COL_KEYWORD
    Words(33).Text = "UBound"
    Words(33).Color = COL_KEYWORD
    Words(34).Text = "Xor"
    Words(34).Color = COL_KEYWORD
    Words(35).Text = "Or"
    Words(35).Color = COL_KEYWORD
    Words(36).Text = "Empty"
    Words(36).Color = COL_KEYWORD
    Words(37).Text = "<%"
    Words(37).Color = COL_KEYWORD
    Words(38).Text = "%>"
    Words(38).Color = COL_KEYWORD
    Words(39).Text = "Rem"
    Words(39).Color = COL_COMMENT
    ' sort the array
    CombSort Words
    
    ' build the index of letter positions
    BuildIndex
End Sub

'//--[LoadFile]-------------------------------------------------------------//
'
'  Loads and colors a file in the RTB
'
Public Sub LoadFile(RTB As RichTextBox, Optional ByVal sFilePath As String)
Dim FF As Long
Dim lStart As Long
Dim lFinish As Long
Dim Text As String
Dim lSelstart As Integer
Dim lExt As String
    DoEvents
    
    If sFilePath <> "" Then
      lExt = Mid(sFilePath, InStrRev(sFilePath, ".") + 1)
      If (UCase(lExt) = "ASP" Or UCase(lExt) = "HTM" Or UCase(lExt) = "HTML") Then
        MDoColor = True
      Else
        MDoColor = False
      End If
      ' load the file
      FF = FreeFile
      Open sFilePath For Input As FF
        RTB.Text = Input(LOF(FF), FF)
      Close FF
      lSelstart = 0
    Else
      lSelstart = Len(RTB.Text) - 20
    End If

    ' split the text into lines and color them one by one
    LockWindowUpdate RTB.hWnd
    RTB.Visible = False
    Text = RTB.Text
    basColor.sText = RTB.Text
    lStart = 1
    Do While lStart <> 2 And lStart < Len(Text)
        ' find the end of this line
        lFinish = InStr(lStart + 1, Text, vbCrLf)
        If lFinish = 0 Then lFinish = Len(Text)
            
        ' color it
        DoColor RTB, lStart, lFinish
        
        ' move up to get the next line
        lStart = lFinish + 2
    Loop
    
    ' reset the cursor
    On Error Resume Next
    RTB.SelStart = lSelstart
    RTB.Visible = True
    LockWindowUpdate 0&
End Sub

'//--[DoColor]--------------------------------------------------------------//
'
'  Here it is - the beast itself. This routine colors
'  a single line of text within the RTB. It will
'  split each line up into words using the custom
'  split function (SplitWords), then match each word
'  against the list of keywords.
'
Public Sub DoColor(RTB As RichTextBox, ByVal lStart As Long, ByVal lFinish As Long, Optional ByVal pDoColor As Boolean = True)
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
    
        If Trim$(sWords(i)) <> "" Then
          If Left$(sWords(i), 1) <> "'" Then sWords(i) = Replace(sWords(i), Chr(9), "")
          If IsWithinASP(sText, sWords(i), InStr(lStart, sText, sWords(i))) Or Trim(sWords(i)) = "<%" Or Trim(sWords(i)) = "%>" Then
            ' check for comment in the middle of a line
            If Left$(sWords(i), 1) = CHAR_COMMENT Then
              
                ' color the rest of the line
                If InStr(lStart, sText, sWords(i)) > 0 Then
                  
                    RTB.SelStart = InStr(lStart, sText, sWords(i)) - 1
                    RTB.SelLength = Len(sWords(i))
                    RTB.SelColor = COL_COMMENT
                End If
            
            ElseIf sWords(i) = "<%" Or sWords(i) = "%>" Then
                lColor = COL_COMMENT_HTML
                RTB.SelStart = InStr(lStart + lCurPos - 1, sText, sWords(i)) - 1
                RTB.SelLength = Len(sWords(i))
                RTB.SelColor = lColor
                'HighLightWord RTB, sWords(i), vbYellow
            Else
                ' its a normal keyword - so color it
                ' first get the array positions from
                ' the index
                If sWords(i) = "ByVal" Or sWords(i) = "Byte" Then
                    DoEvents: End If
                sChar = Left$(LCase$(sWords(i)), 1)
                If sChar = "=" Then
                  sWords(i) = Mid(sWords(i), 2)
                  sChar = Left$(LCase$(sWords(i)), 1)
                End If
                ' if we've got a valid alphabetic char
                If sChar <> "" Then
                    ' convert this char to an index in the letters array
                    lIndex = Asc(sChar) - 97
                    ' if the index is a valid one - this
                    ' means that the text is a word, so
                    ' we should try to color it
                    If lIndex >= 0 And lIndex < UBound(Letters) Then
                        ' color the word, passing the index parameters
                        lColor = GetColor(sWords(i), _
                                    Letters(lIndex).Start, _
                                    Letters(lIndex).Finish)
                        ' if a color was returned - color the word
                        If lColor Then
                            ' locate the word in the line
                            RTB.SelStart = InStr(lStart + lCurPos - 1, sText, sWords(i)) - 1
                            RTB.SelLength = Len(sWords(i))
                            RTB.SelColor = lColor
                        End If
                    End If
                End If ' sChar <> ""
            End If ' CHAR_COMMENT
          Else 'if not asp, color for html
            html = html & sWords(i) & " "
          End If 'within asp
        End If ' sWords(i) <> ""
        
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
Public Sub DoClipBoardPaste(RTB As RichTextBox)
Dim lCursor As Long
Dim lStart As Long
Dim lFinish As Long
Dim sText As String
Dim p1 As Long, p2 As Long
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
    lFinish = RTB.SelStart + Len(sText)
    
    ' now add the text to the box
    RTB.SelText = sText
    basColor.sText = RTB.Text
    
    ' now color each line individually starting
    ' from lStart since this is the position of
    ' the first changed line
    p1 = lStart
    Do
        ' find the next EOL character, this combined
        ' with lStart gives us the line to color
        p2 = InStr(p1, RTB.Text, vbCrLf)
        If p2 = 0 Then p2 = lFinish
                    
        ' now strip out this line and color it
        ' color it black first to remove any
        ' previous coloring..
        RTB.SelStart = p1 - 1
        RTB.SelLength = p2 - p1
        RTB.SelColor = vbBlack
        DoColor RTB, p1, p2
        
        ' move the start pointer on to just after
        ' the last EOL character - essentially onto
        ' the next actual line of text
        p1 = p2 + 2
              
        ' exit condition - keep going until we can't
        ' find any more vbCrLf (<>2) and while
        ' p1 (the start of line pointer) is lower
        ' that lFinish (the end of the text we're
        ' coloring)... easy enough
        If p1 = 2 Or p1 >= lFinish + 2 Then Exit Do
        DoEvents
    Loop
    
    ' restore the original values
    RTB.SelStart = lCursor + Len(sText)
    RTB.SelColor = vbBlack
    
    ' null the keypress (to avoid the text pasting twice)
    LockWindowUpdate 0&

End Sub

'#--------------------------------------------------------------------------#
'#  private internals
'#--------------------------------------------------------------------------#

'//--[BuildIndex]----------------------------------------------------------//
'
'  Takes the Words array and constructs an alphabetical
'  index which it puts into the Letters array.
'  Each item in the letters array accounts for a letter
'  in the alphabet - Letters(0) = "a".
'  The .Start property is the Index in the Words array
'  at which that letter starts, and the finish is the
'  same. The purpose of this is to get Hi and Lo params
'  for the GetColor (a standard binary search algorithm).
'  This saves several loops round the algorithm.
'
Private Sub BuildIndex()
Dim i As Long, j As Long
Dim sChar As String
Dim bStart As Boolean

    ' go through each letter in the alphabet
    ReDim Letters(25)
    For i = 0 To 25
        ' get the current char
        sChar = Chr$(i + 97)
        ' find the first and last instances of the letter
        For j = LBound(Words) To UBound(Words)
            If Left$(LCase$(Words(j).Text), 1) = sChar Then
                If Not bStart Then
                    ' found the start
                    bStart = True
                    Letters(i).Start = j
                End If
                ' if we've hit the end of the list
                If j = UBound(Words) Then
                    Letters(i).Finish = j
                    Exit Sub
                End If
            Else
                ' its a different char
                If bStart Then
                    ' we've found the end
                    Letters(i).Finish = j - 1
                    bStart = False
                    Exit For
                End If
                ' see if we've gone too far -
                ' there are no words beginning with
                ' this letter in the list
                If Left$(LCase$(Words(j).Text), 1) > sChar Then
                    Exit For
                End If
            End If
        Next j
    Next i

End Sub

'//--[GetColor]--------------------------------------------------------------//
'
'  Searches the Words array for a match using a standard
'  binary search algorithm, using the Lo and Hi params
'  as starting points.
'
Private Function GetColor(ByVal sWord As String, _
                          ByVal Lo As Long, _
                          ByVal Hi As Long) As Long
Dim lHi As Long
Dim lLo As Long
Dim lMid As Long
    
    ' standard binary search the words array
    ' return the color if a match is found
    lLo = Lo
    lHi = Hi
    Do While lHi >= lLo
        lMid = (lLo + lHi) \ 2
        If LCase$(Words(lMid).Text) = LCase$(sWord) Then
            GetColor = Words(lMid).Color
            Exit Do
        End If
        If LCase$(Words(lMid).Text) > LCase$(sWord) Then
            lHi = lMid - 1
        Else
            lLo = lMid + 1
        End If
    Loop
    
End Function

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
Dim sWordBreaks(0 To 2) As String
Dim lBreakPoints() As Long
Dim sSubWords() As String 'Collection of <%
Dim sSubEWords() As String 'Collection of %>
Dim lBreak As Long
    
    ' list of word break characters
    sWordBreaks(0) = " "
    'sWordBreaks(1) = "("
    'sWordBreaks(2) = ")"
    'sWordBreaks(3) = "."
'    sWordBreaks(3) = "="
    sWordBreaks(1) = "'"
    sWordBreaks(2) = CHAR_COMMENT ' comments
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

Private Function IsWithinASP(ByVal pContent As String, ByVal pText As String, ByVal pEnd As Long) As Boolean
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

'Private Sub DoColorHTML(ByVal RTB As RichTextBox, ByVal pText As String, ByVal pHTML As String, ByVal pStart As Long, pEnd As Long)
''
'' Coloring the html tags
''
'
'Dim i As Long
'Dim tags() As String
'Dim lStr As String
'Dim lIndex As Integer
'Dim lColor As Long
'Dim lPos As Integer
'Dim lTag As String
'Dim sText As String
'  On Error Resume Next
'  sText = Mid(pText, 1, pEnd)
'  If pHTML <> "" Then
'    tags = Split(pHTML, ">")
'    For i = LBound(tags) To UBound(tags)
'      tags(i) = Trim(tags(i))
'      If InStr(tags(i), "<") > 0 Then tags(i) = Mid(tags(i), InStr(tags(i), "<"))
'      If Left(tags(i), 4) = "<!--" Then
'        pEnd = InStr(pStart + 1, pText, "-->")
'        RTB.SelStart = pStart - 1
'        RTB.SelLength = pEnd - pStart + 3
'        RTB.SelColor = COL_COMMENT_HTML
'        pStart = pEnd + 3
'      ElseIf Left(tags(i), 1) = "<" Then
'        lPos = InStr(tags(i), " ")
'        If lPos = 0 Then lPos = Len(tags(i))
'        pEnd = GetCloseTagPos(sText, pStart + 1) 'InStr(pStart + 1, sText, ">")
'        lStr = Mid(RTB.Text, pStart, pEnd)
'        'ParseLine lStr, RTB, pStart
'        pStart = pEnd + 1
'      End If
'    Next
'  End If
'End Sub

Private Sub DoColorHTML(ByVal RTB As RichTextBox, ByVal pText As String, ByVal pHTML As String, ByVal pStart As Long, pEnd As Long)
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
        RTB.SelColor = COL_COMMENT_HTML
        pStart = pEnd + 3
      ElseIf Left(tags(i), 1) = "<" Then
        lPos = InStr(tags(i), " ")
        If lPos = 0 Then lPos = Len(tags(i))
        'If InStr(tags(i), " ") > 0 Then
          
          sChar = Mid(LCase(tags(i)), 2, 1)
          lTag = IIf(sChar = "/", Trim(Mid(tags(i), 3, lPos - 1)), Trim(Mid(tags(i), 2, lPos - 1)))
          If sChar = "/" Then sChar = Mid(LCase(tags(i)), 3, 1)
          
          ' if we've got a valid alphabetic char
          If sChar <> "" Then
              ' convert this char to an index in the letters array
              lIndex = Asc(sChar) - 97
              ' if the index is a valid one - this
              ' means that the text is a word, so
              ' we should try to color it
              If lIndex >= 0 And lIndex < UBound(Letters) Then
                  ' color the word, passing the index parameters
                  lColor = GetColorHTML(lTag, _
                              htmlLetters(lIndex).Start, _
                              htmlLetters(lIndex).Finish)
                  pEnd = GetCloseTagPos(sText, pStart + 1) 'InStr(pStart + 1, sText, ">")
                  ' if a color was returned - color the word
                  If lColor Then
                      pStart = InStr(pStart, sText, Split(tags(i), " ")(0))
                      ' locate the word in the line
                      RTB.SelStart = pStart - 1
                      RTB.SelLength = pEnd - pStart + 1
                      RTB.SelColor = lColor
                  End If
                  pStart = pEnd + 1
              End If
          End If ' sChar <> ""
        'End If
      End If
    Next
  End If
End Sub

Public Sub InitKeyhtmlWordsHTML()
    ' initialize the array of htmlWords for html
    ReDim htmlWords(0 To 18)
    htmlWords(0).Text = "HTML"
    htmlWords(0).Color = COL_TAG
    htmlWords(1).Text = "BODY"
    htmlWords(1).Color = COL_TAG
    htmlWords(2).Text = "P"
    htmlWords(2).Color = COL_TAG
    htmlWords(3).Text = "B"
    htmlWords(3).Color = COL_TAG
    htmlWords(4).Text = "TABLE"
    htmlWords(4).Color = COL_TABLE
    htmlWords(5).Text = "TD"
    htmlWords(5).Color = COL_TABLE
    htmlWords(6).Text = "TR"
    htmlWords(6).Color = COL_TABLE
    htmlWords(7).Text = "TH"
    htmlWords(7).Color = COL_TABLE
    htmlWords(8).Text = "TBODY"
    htmlWords(8).Color = COL_TABLE
    htmlWords(9).Text = "CAPTION"
    htmlWords(9).Color = COL_TABLE
    htmlWords(10).Text = "SCRIPT"
    htmlWords(10).Color = COL_SCRIPT
    htmlWords(11).Text = "U"
    htmlWords(11).Color = COL_TAG
    htmlWords(12).Text = "A"
    htmlWords(12).Color = COL_TAG
    htmlWords(13).Text = "IMG"
    htmlWords(13).Color = COL_TAG
    htmlWords(14).Text = "INPUT"
    htmlWords(14).Color = COL_TAG
    htmlWords(15).Text = "FORM"
    htmlWords(15).Color = COL_TAG
    htmlWords(16).Text = "BR"
    htmlWords(16).Color = COL_TAG
    htmlWords(17).Text = "HR"
    htmlWords(17).Color = COL_TAG
    htmlWords(18).Text = "STRONG"
    htmlWords(18).Color = COL_TAG
    ' sort the array
    CombSort htmlWords
    
    ' build the index of letter positions
    BuildIndexHTML
End Sub

Private Sub BuildIndexHTML()
Dim i As Long, j As Long
Dim sChar As String
Dim bStart As Boolean

    ' go through each letter in the alphabet
    ReDim htmlLetters(25)
    For i = 0 To 25
        ' get the current char
        sChar = Chr$(i + 97)
        ' find the first and last instances of the letter
        For j = LBound(htmlWords) To UBound(htmlWords)
            If Left$(LCase$(htmlWords(j).Text), 1) = sChar Then
                If Not bStart Then
                    ' found the start
                    bStart = True
                    htmlLetters(i).Start = j
                End If
                ' if we've hit the end of the list
                If j = UBound(htmlWords) Then
                    htmlLetters(i).Finish = j
                    Exit Sub
                End If
            Else
                ' its a different char
                If bStart Then
                    ' we've found the end
                    htmlLetters(i).Finish = j - 1
                    bStart = False
                    Exit For
                End If
                ' see if we've gone too far -
                ' there are no htmlWords beginning with
                ' this letter in the list
                If Left$(LCase$(htmlWords(j).Text), 1) > sChar Then
                    Exit For
                End If
            End If
        Next j
    Next i

End Sub

Private Function GetColorHTML(ByVal sWord As String, _
                          ByVal Lo As Long, _
                          ByVal Hi As Long) As Long
Dim lHi As Long
Dim lLo As Long
Dim lMid As Long
    
    ' standard binary search the htmlWords array
    ' return the color if a match is found
    lLo = Lo
    lHi = Hi
    Do While lHi >= lLo
        lMid = (lLo + lHi) \ 2
        If LCase$(htmlWords(lMid).Text) = LCase$(sWord) Then
            GetColorHTML = htmlWords(lMid).Color
            Exit Do
        End If
        If LCase$(htmlWords(lMid).Text) > LCase$(sWord) Then
            lHi = lMid - 1
        Else
            lLo = lMid + 1
        End If
    Loop
    If GetColorHTML = 0 Then GetColorHTML = COL_NOCOLOR
End Function

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

Public Sub InitParser()
    Dim vArr
    
    vArr = Split(RGB_COMMENT, ",")
    mrgbComment = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_STRING, ",")
    mrgbString = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_RESERVED, ",")
    mrgbReserved = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_FUNC_OBJ, ",")
    mrgbFuncObj = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_DELIMITER, ",")
    mrgbDelimiter = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_NORMAL, ",")
    mrgbNormal = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_TAGDELIMITER, ",")
    mrgbTagdelimiter = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_TAG, ",")
    mrgbTag = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_PROPERTY, ",")
    mrgbProperty = RGB(vArr(0), vArr(1), vArr(2))
    
End Sub

'
' Sub Highlight
' Color this range in the RichTextBox. Note that you could also apply bold,
' italic, etc. to the selection at the same time.
'
Private Sub Highlight(RTB As RichTextBox, SyntaxType As SyntaxTypes, StartPos As Long, Length As Long)
        RTB.SelStart = StartPos - 1
        RTB.SelLength = Length
    
    Select Case SyntaxType
        Case SyntaxTypes.ColorComment
            RTB.SelColor = mrgbComment
        Case SyntaxTypes.ColorString
            RTB.SelColor = mrgbString
        Case SyntaxTypes.ColorReserved
            RTB.SelColor = mrgbReserved
        Case SyntaxTypes.ColorFuncObj
            RTB.SelColor = mrgbFuncObj
        Case SyntaxTypes.ColorDelimiter
            RTB.SelColor = mrgbDelimiter
        Case SyntaxTypes.ColorTag
            RTB.SelColor = mrgbTag
        Case SyntaxTypes.ColorTagdelimiter
            RTB.SelColor = mrgbTagdelimiter
        Case SyntaxTypes.ColorProperty
            RTB.SelColor = mrgbProperty
        Case Else
            RTB.SelColor = mrgbNormal
    End Select

End Sub

'
' Sub ParseLine
' Lines are treated independently. Parseline is the main parsing code. Scan
' line from left to right, emitting text to be colored.
'
Private Sub ParseLine(ByVal s As String, RTB As RichTextBox, ByVal RTBPos As Long)
    'Debug.Print s
    
    Dim bInString As Boolean    ' are we in a quoted string?
    bInString = False
    
    Dim bInWord As Boolean      ' are we in a word? (not a string, comment,
                                ' or delimiter)
    bInWord = False
    
    Dim sCurString As String        ' the current set of characters
    Dim lCurStringStart As Long     '   - where it starts
    Dim sCurChar As String          ' the current character
    
    Dim i As Long
    
    For i = 1 To Len(s)
        sCurChar = Mid(s, i, 1)

        If sCurChar = """" Then
            ' if not already in a string, then this quote begins a string
            ' otherwise, we are in a string, and this quote ends it
            If bInString Then
                sCurString = sCurString & sCurChar
                Highlight RTB, ColorString, lCurStringStart + RTBPos - 1, i - lCurStringStart + 1
                sCurString = ""
                bInString = False
            Else
                If bInWord Then
                    ' before we encounterd the string we were processing a word
                    Highlight RTB, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - lCurStringStart
                    sCurString = ""
                    bInWord = False
                End If
                
                bInString = True
                sCurString = sCurChar
                lCurStringStart = i
            End If
            
            GoTo Next_i ' get next character
        End If
                
        If InStr(1, HTMLDELIMITER, sCurChar) > 0 Then
            If bInWord Then
                ' before we encounterd the delimiter we were processing a word
                Highlight RTB, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - lCurStringStart
                sCurString = ""
                bInWord = False
                If sCurChar = "=" Then
                  sCurString = "="
                  bInWord = True
                End If
            End If
            
            Highlight RTB, ColorTagdelimiter, i + RTBPos - 1, 1
            GoTo Next_i
        End If
            
        If (Not bInWord) And (Not bInString) Then
            bInWord = True
            sCurString = sCurChar
            lCurStringStart = i
            
            GoTo Next_i ' get next character
        End If
            
        ' add current character to the "word" we are in the middle of
        sCurString = sCurString & sCurChar
Next_i:     ' VB style continue
    Next
    
    If bInString Then
        ' before we encounterd the end of the line we were processing a string
        Highlight RTB, ColorString, lCurStringStart + RTBPos - 1, i - lCurStringStart
    ElseIf bInWord Then
        ' before we encounterd the end of the line we were processing a word
        Highlight RTB, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - lCurStringStart
    End If

ExitSub:
    Exit Sub
End Sub

'
' Sub ParseLines
' Feed text, line by line, to the parser.
'
Private Sub ParseLines(ByVal s As String, RTB As RichTextBox, ByVal RTBPos As Long)
    Dim lStartPos As Long
    Dim lEndPos As Long
    
    lStartPos = 1
    
    s = s & vbCrLf
    lEndPos = InStr(lStartPos, s, vbCrLf)
    Do While lEndPos > 0
        ParseLine Mid(s, lStartPos, lEndPos - lStartPos), RTB, RTBPos + lStartPos - 1
        lStartPos = lEndPos + Len(vbCrLf)
        lEndPos = InStr(lStartPos, s, vbCrLf)
    Loop
           
        
End Sub

'
' Function ParseWord
' Determine color for this word by checking for its existence in the keyword
' lists. The word being checked it padded with spaces to prevent matches
' with substrings of keywords.
'
Private Function ParseWord(ByVal Word As String) As SyntaxTypes
    If InStr(1, HTMLTAGS, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
        ParseWord = ColorTag
    ElseIf InStr(1, HTMLPROPERTIES, " " & UCase(Word) & " ", vbTextCompare) > 0 Then
        ParseWord = ColorProperty
    ElseIf Left(Word, 1) = "=" Then
        ParseWord = ColorString
    Else
        ParseWord = Colornormal
    End If
End Function


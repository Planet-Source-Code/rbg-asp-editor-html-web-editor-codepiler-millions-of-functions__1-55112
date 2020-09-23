VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C23E2FE5-54DA-4A88-86B5-6D60D0C5A456}#1.0#0"; "INTELL~1.OCX"
Begin VB.Form frmDocument 
   Appearance      =   0  'Flat
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   6090
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPreview 
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   -20000
      ScaleHeight     =   1995
      ScaleWidth      =   2520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1095
      Width           =   2520
      Begin SHDocVwCtl.WebBrowser wbPreview 
         Height          =   1770
         Left            =   180
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   105
         Width           =   2130
         ExtentX         =   3757
         ExtentY         =   3122
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.PictureBox picSource 
      BorderStyle     =   0  'None
      Height          =   10695
      Left            =   45
      ScaleHeight     =   10695
      ScaleWidth      =   15150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   -30
      Width           =   15150
      Begin IntellProj.Intellisense IntellBox 
         Height          =   690
         Left            =   720
         TabIndex        =   6
         Top             =   495
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   1217
      End
      Begin VB.PictureBox picSidebar 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   180
         ScaleHeight     =   1455
         ScaleWidth      =   255
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   105
         Width           =   255
         Begin VB.Shape spSidebar 
            BorderColor     =   &H00C0C0C0&
            Height          =   360
            Left            =   135
            Top             =   180
            Width           =   15
         End
      End
      Begin RichTextLib.RichTextBox RTB 
         Height          =   10575
         Left            =   15
         TabIndex        =   0
         Top             =   75
         Width           =   15090
         _ExtentX        =   26617
         _ExtentY        =   18653
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmDocument.frx":058A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imlIbox 
         Left            =   1485
         Top             =   2025
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocument.frx":060A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocument.frx":095E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocument.frx":0CB2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6090
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4605
      Width           =   6090
      Begin MSComctlLib.ImageList imlTabs 
         Left            =   4770
         Top             =   -60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   71
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocument.frx":0D3F
               Key             =   "SOURCECLICK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocument.frx":0E17
               Key             =   "SOURCEHIDE"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocument.frx":0F18
               Key             =   "VIEWCLICK"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocument.frx":131F
               Key             =   "VIEWHIDE"
            EndProperty
         EndProperty
      End
      Begin VB.Shape spBottom 
         Height          =   15
         Left            =   2460
         Top             =   15
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   270
         Left            =   0
         Picture         =   "frmDocument.frx":1404
         Top             =   0
         Width           =   1065
      End
      Begin VB.Image Image2 
         Height          =   270
         Left            =   945
         Picture         =   "frmDocument.frx":14CC
         Top             =   0
         Width           =   1365
      End
   End
   Begin VB.Shape spSide 
      BorderColor     =   &H00C0C0C0&
      Height          =   360
      Left            =   3300
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bDirty As Boolean
Private mExt As String
Private mView As Integer '0-Single,1-Both

Public mChange As Boolean
Public Mfilename As String
Public Mkey As String


Private Sub Form_Load()
Dim Lstr As String
  Me.Hide
  'RTB.RightMargin = Screen.Width * 6
  Image2.Tag = "VIEWHIDE"
  Image1.Tag = "SOURCEHIDE"
  'load keywords
  IntellBox.SmallIcons = imlIbox
  Lstr = StringForIntellbox("HTML")
  IntellBox.PopulateListFromString Lstr, True
  ' setup and load!!
  wbPreview.Navigate2 ""
  InitKeyWords
  InitKeyhtmlWordsHTML
  RTB.SelIndent = 400
  Image1_Click
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  spBottom.Left = 0
  spBottom.Top = 0
  spBottom.Width = Me.ScaleWidth
  
  picSource.Width = Me.ScaleWidth - spSide.Width
  picSource.Top = 0
  If mView = 0 Then
    picSource.Height = Me.ScaleHeight - picFooter.Height
  Else
    picSource.Height = Me.ScaleHeight / 2
  End If
  RTB.Left = 0
  RTB.Top = 0
  RTB.Width = picSource.Width
  RTB.Height = picSource.Height - Screen.TwipsPerPixelY
  
  picSidebar.Left = 0
  picSidebar.Top = 0
  picSidebar.Height = RTB.Height - Screen.TwipsPerPixelY * 17
  
  spSidebar.Left = picSidebar.Width - spSidebar.Width
  spSidebar.Top = 0
  spSidebar.Height = picSidebar.Height
  
  picPreview.Width = picSource.Width - spSide.Width
  If mView = 0 Then
    picPreview.Top = 0
  Else
    picPreview.Top = picSource.Height + Screen.TwipsPerPixelY * 1
    picPreview.Left = picSource.Left
  End If
  picPreview.Height = picSource.Height
  
  wbPreview.Left = 0
  wbPreview.Top = 0
  wbPreview.Width = picPreview.Width
  wbPreview.Height = picPreview.Height - Screen.TwipsPerPixelY
  
  spSide.Top = 0
  spSide.Left = 0
  spSide.Height = Me.ScaleHeight
  
  IntellBox.Width = Screen.TwipsPerPixelX * 150
  IntellBox.Height = Screen.TwipsPerPixelY * 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  If mChange Then
    If MsgBox("Document has been changed. Do you want to save changes", vbQuestion + vbYesNo) = vbYes Then
      frmMain.mnuFileSave_Click
    End If
  End If
  frmMain.frmCollection.Remove Mkey
End Sub

Private Sub Image1_Click()
  If Image1.Tag = "SOURCEHIDE" Then
    Image1.Picture = imlTabs.ListImages("SOURCECLICK").Picture
    Image1.Tag = "SOURCECLICK"
    Image2.Picture = imlTabs.ListImages("VIEWHIDE").Picture
    Image2.Tag = "VIEWHIDE"
    Image1.Top = Image2.Top
    Image1.ZOrder vbBringToFront
    picPreview.Left = -20000
    picSource.Left = spSide.Width
  End If
End Sub

Private Sub Image2_Click()
  Screen.MousePointer = vbHourglass
  If Image2.Tag = "VIEWHIDE" Then
    Image2.Picture = imlTabs.ListImages("VIEWCLICK").Picture
    Image2.Tag = "VIEWCLICK"
    Image1.Picture = imlTabs.ListImages("SOURCEHIDE").Picture
    Image1.Tag = "SOURCEHIDE"
    Image2.Top = Image1.Top
    Image2.ZOrder vbBringToFront
    picSource.Left = -20000
    picPreview.Left = spSide.Width
    ShowPreview
  End If
  Screen.MousePointer = vbDefault
End Sub

Public Sub RTB_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lCursor As Long
Dim lSelectLen As Long
Dim lStart As Long
Dim lFinish As Long
Dim sText As String
    mChange = True
    ' ------------------------------
    ' here's the on the fly coloring
    ' ------------------------------
    On Error Resume Next
    ' check for Ctrl+C
    If KeyCode = vbKeyC And Shift = 2 Then Exit Sub
    
    ' check for text being pasted into the box
    If KeyCode = vbKeyV And Shift = 2 Then
        
        Screen.MousePointer = vbHourglass
        DoClipBoardPaste RTB
        KeyCode = 0
        Screen.MousePointer = vbNormal
        Exit Sub
        
    End If
    
    ' if the cursor is moving to a different
    ' line then process the orginal line
    If KeyCode = 13 Or _
         KeyCode = vbKeyUp Or _
            KeyCode = vbKeyDown Then
    
        ' only color this line if it's been changed
        If bDirty Or KeyCode = 13 Then
                        
            ' lock the window to cancel out flickering
            LockWindowUpdate RTB.hWnd
            
            ' store the current cursor pos
            ' and current selection if there is any
            lCursor = RTB.SelStart
            lSelectLen = RTB.SelLength
            
            ' get the line start and end
            If lCursor <> 0 Then
                lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart) + 2
                If lStart = 2 Then lStart = 1
            Else
                lStart = 1
            End If
            lFinish = InStr(lCursor + 1, RTB.Text, vbCrLf)
            If lFinish = 0 Then lFinish = Len(RTB.Text)
            
            ' do the coloring
            basColor.sText = RTB.Text
            DoColor RTB, lStart, lFinish
            
            ' if ENTER was pressed, we should color the next line
            ' as well, so that if a line is broken by the ENTER
            ' the new line and the old line are colored properly
            If KeyCode = 13 Then
                lStart = lCursor + 1
                lFinish = InStr(lStart, RTB.Text, vbCrLf)
                If lFinish = 0 Then lFinish = Len(RTB.Text)
                ' only color if another line exists
                If lStart - 1 <> lFinish Then
                  RTB.SelStart = lStart - 1
                  RTB.SelLength = lFinish - lStart
                  RTB.SelColor = vbBlack
                  DoColor RTB, lStart, lFinish
               End If
            End If
            
            ' reset the properties
            RTB.SelStart = lCursor
            RTB.SelLength = lSelectLen
            RTB.SelColor = vbBlack
            
            ' reset the flag and release the window
            bDirty = False
            LockWindowUpdate 0&
            
        End If
        
    ElseIf Not IsControlKey(KeyCode) Then
                
        ' a different key was pressed - and
        ' this will alter the line so it
        ' needs recoloring when we move off it
        If Not bDirty Then
            
            LockWindowUpdate RTB.hWnd
            
            ' get the line start and end
            lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart + 1) + 1
            lFinish = InStr(RTB.SelStart + 1, RTB.Text, vbCrLf)
            If lFinish = 0 Then lFinish = Len(RTB.Text)
            
            ' color the line (remembering the cursor position)
            lCursor = RTB.SelStart
            lSelectLen = RTB.SelLength
            RTB.SelStart = lStart
            RTB.SelLength = lFinish - lStart
            RTB.SelColor = vbBlack
            RTB.SelStart = lCursor
            RTB.SelLength = lSelectLen
            bDirty = True
            
            LockWindowUpdate 0&
            
        End If
        
    End If
    If IntellBox.Visible = True Then
    Select Case KeyCode
    Case vbKeyUp
        IntellBox.MoveListUp
        RTB.SetFocus
        KeyCode = vbNull
    Case vbKeyDown
        IntellBox.MoveListDown
        RTB.SetFocus
        KeyCode = vbNull
    Case vbKeyPageUp
        IntellBox.MoveToTop
        RTB.SetFocus
        KeyCode = vbNull
    Case vbKeyPageDown
        IntellBox.MoveToBottom
        RTB.SetFocus
        KeyCode = vbNull
    Case vbKeyDelete
        IntellBox.Visible = False
        IntellBox.Clear
    Case vbKeyHome
        IntellBox.MoveToTop
        RTB.SetFocus
        KeyCode = vbNull
    Case vbKeyEnd
        IntellBox.MoveToBottom
        RTB.SetFocus
        KeyCode = vbNull
    Case vbKeyLeft
        IntellBox.Visible = False
        IntellBox.Clear
    Case vbKeyRight
        IntellBox.Visible = False
        IntellBox.Clear
    End Select
  End If
End Sub

Private Function IsControlKey(ByVal KeyCode As Long) As Boolean
    ' check if the key is a control key
    Select Case KeyCode
        Case vbKeyLeft, vbKeyRight, vbKeyHome, _
             vbKeyEnd, vbKeyPageUp, vbKeyPageDown, _
             vbKeyShift, vbKeyControl
            IsControlKey = True
        Case Else
            IsControlKey = False
    End Select
End Function

Public Function ShowPreview()
'
'Show the preview
'
Dim FileNum As Integer
  On Error Resume Next
  FileNum = FreeFile
  Open App.Path & "\CodePiler.html" For Output As #FileNum
  Print #FileNum, RTB.Text
  Close #FileNum
  wbPreview.Navigate2 App.Path & "\CodePiler.html"
End Function

Private Function StringForIntellbox(Optional ByVal pObject As String) As String
Dim Lstr As String
  Select Case UCase(pObject)
  Case "RESPONSE"
    Lstr = ""
    Lstr = Lstr & ",2,AddHeader,AddHeader,9,0" & vbCrLf
    Lstr = Lstr & ",2,AppendToLog,AppendToLog,11,0" & vbCrLf
    Lstr = Lstr & ",2,BinaryWrite,BinaryWrite,11,0" & vbCrLf
    Lstr = Lstr & ",3,Buffer,Buffer,6,0" & vbCrLf
    Lstr = Lstr & ",3,CacheControl,CacheControl,12,0" & vbCrLf
    Lstr = Lstr & ",3,CharSet,CharSet,7,0" & vbCrLf
    Lstr = Lstr & ",2,Clear,Clear,5,0" & vbCrLf
    Lstr = Lstr & ",3,ContentType,ContentType,11,0" & vbCrLf
    Lstr = Lstr & ",3,Cookies,Cookies,7,0" & vbCrLf
    Lstr = Lstr & ",2,End,End,3,0" & vbCrLf
    Lstr = Lstr & ",3,Expires,Expires,7,0" & vbCrLf
    Lstr = Lstr & ",3,ExpiresAbsolute,ExpiresAbsolute,15,0" & vbCrLf
    Lstr = Lstr & ",2,Flush,Flush,5,0" & vbCrLf
    Lstr = Lstr & ",2,IsClientConnected,IsClientConnected,17,0" & vbCrLf
    Lstr = Lstr & ",2,Pics,Pics,4,0" & vbCrLf
    Lstr = Lstr & ",2,Redirect,Redirect,8,0" & vbCrLf
    Lstr = Lstr & ",3,Status,Status,6,0" & vbCrLf
    Lstr = Lstr & ",2,Write,Write,5,0"
    IntellBox.Height = Me.TextHeight("Z") * 6
  Case "REQUEST"
    Lstr = ""
    Lstr = Lstr & ",2,BinaryRead,BinaryRead,10,0" & vbCrLf
    Lstr = Lstr & ",3,ClientCertificate,ClientCertificate,17,0" & vbCrLf
    Lstr = Lstr & ",3,Cookies,Cookies,7,0" & vbCrLf
    Lstr = Lstr & ",3,Form,Form,4,0" & vbCrLf
    Lstr = Lstr & ",3,Item,Item,4,0" & vbCrLf
    Lstr = Lstr & ",3,QueryString,QueryString,11,0" & vbCrLf
    Lstr = Lstr & ",3,ServerVariables,ServerVariables,15,0" & vbCrLf
    Lstr = Lstr & ",3,TotalBytes,TotalBytes,9,0"
    IntellBox.Height = Me.TextHeight("Z") * 6
  Case "SESSION"
    Lstr = ""
    Lstr = Lstr & ",2,Abandon,Abandon,7,0" & vbCrLf
    Lstr = Lstr & ",3,CodePage,CodePage,8,0" & vbCrLf
    Lstr = Lstr & ",3,Contents,Contents,8,0" & vbCrLf
    Lstr = Lstr & ",3,LCID,LCID,4,0" & vbCrLf
    Lstr = Lstr & ",3,SessionID,SessionID,9,0" & vbCrLf
    Lstr = Lstr & ",3,StaticObjects,StaticObjects,13,0" & vbCrLf
    Lstr = Lstr & ",3,Timeout,Timeout,7,0" & vbCrLf
    Lstr = Lstr & ",3,Value,Value,7,0"
    IntellBox.Height = Me.TextHeight("Z") * 6
  Case "APPLICATION"
    Lstr = ""
    Lstr = Lstr & ",3,Contents,Contents,8,0" & vbCrLf
    Lstr = Lstr & ",2,Lock,Lock,4,0" & vbCrLf
    Lstr = Lstr & ",3,StaticObjects,StaticObjects,13,0" & vbCrLf
    Lstr = Lstr & ",2,Unlock,Unlock,6,0" & vbCrLf
    Lstr = Lstr & ",3,Value,Value,5,0"
    IntellBox.Height = Me.TextHeight("Z") * 4
  Case "SERVER"
    Lstr = ""
    Lstr = Lstr & ",2,CreateObject,CreateObject,12,0" & vbCrLf
    Lstr = Lstr & ",2,HTMLEncode,HTMLEncode,10,0" & vbCrLf
    Lstr = Lstr & ",2,MapPath,MapPath,7,0" & vbCrLf
    Lstr = Lstr & ",3,ScriptTimeout,ScriptTimeout,13,0" & vbCrLf
    Lstr = Lstr & ",2,URLEncode,URLEncode,9,0" & vbCrLf
    Lstr = Lstr & ",2,URLPathEncode,URLPathEncode,13,0"
    IntellBox.Height = Me.TextHeight("Z") * 6
  Case "SCRIPTINGCONTEXT"
    Lstr = ""
    Lstr = Lstr & ",3,Application,Application,11,0" & vbCrLf
    Lstr = Lstr & ",3,Request,Request,7,0" & vbCrLf
    Lstr = Lstr & ",3,Response,Response,8,0" & vbCrLf
    Lstr = Lstr & ",3,Server,Server,6,0" & vbCrLf
    Lstr = Lstr & ",3,Session,Session,7,0"
    IntellBox.Height = Me.TextHeight("Z") * 5
  Case "CONTENTS", "STATICOBJECTS", "QUERYSTRING", "FORM", "COOKIES"
    Lstr = ""
    Lstr = Lstr & ",3,Count,Count,5,0" & vbCrLf
    Lstr = Lstr & ",3,Item,Item,4,0" & vbCrLf
    Lstr = Lstr & ",3,Key,Key,3,0"
    IntellBox.Height = Me.TextHeight("Z") * 3
  Case "HTML"
    Lstr = ""
    Lstr = Lstr & ",1,!DOCTYPE,!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">,@,0" & vbCrLf
    Lstr = Lstr & ",1,A HREF,A HREF=""></A>,8,0" & vbCrLf
    Lstr = Lstr & ",1,HTML,HTML></HTML>,5,0" & vbCrLf
    Lstr = Lstr & ",2,Non-Breaking Space;,&nbsp;,@,1" & vbCrLf
    Lstr = Lstr & ",1,HEAD,head></head>,5,0" & vbCrLf
    Lstr = Lstr & ",1,BODY,body></body>,5,0" & vbCrLf
    Lstr = Lstr & ",1,TITLE,title></title>,6,0" & vbCrLf
    Lstr = Lstr & ",1,PRE,pre></pre>,4,0" & vbCrLf
    Lstr = Lstr & ",1,H1,hl></hl>,3,0" & vbCrLf
    Lstr = Lstr & ",1,H6,h6></h6>,3,0" & vbCrLf
    Lstr = Lstr & ",1,B,b></b>,2,0" & vbCrLf
    Lstr = Lstr & ",1,I,i></i>,2,0" & vbCrLf
    Lstr = Lstr & ",1,TT,tt></tt>,3,0" & vbCrLf
    Lstr = Lstr & ",1,CITE,cite></cite>,5,0" & vbCrLf
    Lstr = Lstr & ",1,EM,em></em>,3,0" & vbCrLf
    Lstr = Lstr & ",1,STRONG,strong></strong>,7,0" & vbCrLf
    Lstr = Lstr & ",1,FONT SIZE,font size=></font>,10,0" & vbCrLf
    Lstr = Lstr & ",1,TABLE,table></table>,6,0" & vbCrLf
    Lstr = Lstr & ",1,TR,tr></tr>,3,0" & vbCrLf
    Lstr = Lstr & ",1,TD,td></td>,3,0" & vbCrLf
    Lstr = Lstr & ",1,TH,th></th>,3,0" & vbCrLf
    Lstr = Lstr & ",1,LINK,link></link>,3,0" & vbCrLf
    Lstr = Lstr & ",2,@,&#64;,@,1"
    IntellBox.Height = Me.TextHeight("Z") * 6
  End Select
  StringForIntellbox = Lstr
End Function

Private Sub ProcessIntell(KeyAscii As Integer)
Dim Pt As POINTAPI
Dim Lword As String
Dim Llist As String
  Select Case KeyAscii
  Case 60  '<
    If mChange Then
      IntellBox.PopulateListFromString StringForIntellbox("HTML")
    End If
    ' Get the position of the cursor
    GetCaretPos Pt
    ' Move the popup window to the caret
    IntellBox.Move (Pt.x * Screen.TwipsPerPixelX) + Me.TextWidth("Z") + RTB.Left + 50 - 315, _
                   (Pt.y * Screen.TwipsPerPixelY) + Me.TextHeight("Z") + RTB.Top + 50
    ' Check if the popup window is within the form
    If IntellBox.Left + IntellBox.Width > ScaleWidth Then IntellBox.Move ScaleWidth - IntellBox.Width - 300
    If IntellBox.Top + IntellBox.Height > ScaleHeight Then IntellBox.Move IntellBox.Left, (Pt.y * Screen.TwipsPerPixelX) - IntellBox.Height
    ' Show the popup window
    IntellBox.Visible = True
    mChange = False
  Case 62 '>
    IntellBox.Visible = False
    IntellBox.Clear
  Case vbKeyReturn
    If IntellBox.Visible = True Then
      RTB.SelStart = RTB.SelStart - IntellBox.InputLen - IntellBox.RemovePrev
      RTB.SelLength = IntellBox.InputLen + IntellBox.RemovePrev
      'rtb.SelColor = varColorTag
      RTB.SelText = IntellBox.Value
      RTB.SelStart = RTB.SelStart - IntellBox.CursorAdjust
      IntellBox.Visible = False
      IntellBox.Clear
      KeyAscii = vbNull
    End If
  Case vbKeyUp
    If IntellBox.Visible = True Then
      IntellBox.MoveListUp
      KeyAscii = vbNull
    End If
  Case vbKeyDown
    If IntellBox.Visible = True Then
      IntellBox.MoveListDown
      KeyAscii = vbNull
    End If
  Case vbKeyBack
    If IntellBox.Visible = True Then
      If IntellBox.InputLen = 0 Then
        IntellBox.Visible = False
        IntellBox.Clear
      Else
        IntellBox.RemoveChar
      End If
    End If
  Case vbKeyEscape
    IntellBox.Visible = False
    IntellBox.Clear
  Case vbKeyLeft
    IntellBox.Visible = False
    IntellBox.Clear
  Case 46 'Decimal point
    If IntellBox.Visible Then ProcessIntell vbKeyReturn
    Lword = ""
    Llist = ""
    Lword = GetWord
    Llist = StringForIntellbox(Lword)
    If Llist <> "" Then
      IntellBox.PopulateListFromString Llist, False
      'Move the intellbox to cursor position
      GetCaretPos Pt
      IntellBox.Move (Pt.x * Screen.TwipsPerPixelX) + Me.TextWidth("Z") + RTB.Left + 20 - 315, _
                     (Pt.y * Screen.TwipsPerPixelY) + Me.TextHeight("Z") + RTB.Top + 20
      If IntellBox.Left + IntellBox.Width > ScaleWidth Then IntellBox.Move ScaleWidth - IntellBox.Width - 300
      If IntellBox.Top + IntellBox.Height > ScaleHeight Then IntellBox.Move IntellBox.Left, (Pt.y * Screen.TwipsPerPixelX) - IntellBox.Height
      IntellBox.Visible = True
      mChange = True
    End If
  Case 32 'Space
    If IntellBox.Visible Then ProcessIntell vbKeyReturn
  Case Else:
    If IntellBox.Visible = True Then IntellBox.AddChar Chr(KeyAscii)
  End Select
End Sub

Private Function GetWord() As String
Dim lPos As Long
Dim Lpos1 As Long
Dim Lpos2 As Long
Dim Ltmp As String
Dim Lstr As String
  Lstr = UCase(Mid(RTB.Text, 1, RTB.SelStart))
  Ltmp = ""
  For lPos = Len(Lstr) To 1 Step -1
    If (Asc(Mid(Lstr, lPos, 1)) > vbKeyA - 1 And Asc(Mid(Lstr, lPos, 1)) < vbKeyZ + 1) Then
      Ltmp = Ltmp & Mid(Lstr, lPos, 1)
    Else
      Exit For
    End If
  Next
  Lpos1 = InStrRev(Lstr, "<%")
  Lpos2 = InStrRev(Lstr, "%>")
  If Lpos1 > 0 Then
    If Lpos1 > Lpos2 Then
      GetWord = StrReverse(Ltmp)
    End If
  End If
End Function

Private Sub RTB_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  ProcessIntell KeyAscii
End Sub

Private Sub RTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'  RTB_KeyDown vbKeyReturn, 0
End Sub

Private Sub RTB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    PopupMenu frmMain.mmupopup, , RTB.Left + x, RTB.Top + y
  End If
End Sub

Private Sub wbPreview_StatusTextChange(ByVal Text As String)
  frmMain.stBar.Panels("P1").Text = Text
End Sub

Public Sub LoadNew(Optional ByVal pASP As Integer)
Dim Lstr As String
  Lstr = ""
  If pASP > -1 Then
    If pASP = 1 Then Lstr = Lstr & "<%@ Laguage=VBScript%>" & vbCrLf & vbCrLf
    Lstr = Lstr & "<HTML>" & vbCrLf
    Lstr = Lstr & "<head>" & vbCrLf
    Lstr = Lstr & "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">" & vbCrLf
    Lstr = Lstr & "<meta name=""GENERATOR"" content=""CodePiler"">" & vbCrLf
    Lstr = Lstr & "<meta name=""ProgId"" content=""CodePiler.Document"">" & vbCrLf
    Lstr = Lstr & "<title>New Document " & frmMain.mBlankPage & "</title>" & vbCrLf
    Lstr = Lstr & "</head>" & vbCrLf
    Lstr = Lstr & "<body>" & vbCrLf & vbCrLf
    Lstr = Lstr & "</body>" & vbCrLf
    Lstr = Lstr & "</html>" & vbCrLf
  End If
  RTB.Visible = False
  RTB.Text = Lstr
  If frmMain.ctbHeader.ButtonChecked("Split") Then ShowBoth
End Sub

Public Sub ShowBoth()
'Show the both source and preview
  mView = 1
  picFooter.Visible = False
  ShowPreview
  Form_Resize
End Sub

Public Sub ShowSource()
'Show only source
  mView = 0
  picFooter.Visible = True
  picPreview.Left = -20000
  Form_Resize
End Sub

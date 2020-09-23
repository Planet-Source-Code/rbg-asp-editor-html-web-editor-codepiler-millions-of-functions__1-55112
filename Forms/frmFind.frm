VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFind 
   Caption         =   "Find and Replace"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgResult 
      Left            =   270
      Top             =   4305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":084A
            Key             =   "RESULT"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFindAll 
      Caption         =   "Find All"
      Height          =   360
      Left            =   6525
      TabIndex        =   20
      Top             =   1245
      Width           =   1230
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   5910
      Picture         =   "frmFind.frx":090C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Width           =   330
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   3945
      TabIndex        =   18
      Top             =   840
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   105
      Top             =   3015
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbFind 
      Height          =   705
      Left            =   6210
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3330
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   1244
      _Version        =   393217
      TextRTF         =   $"frmFind.frx":0B3D
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   135
      Top             =   885
   End
   Begin VB.ComboBox cboFindWhat 
      Height          =   315
      ItemData        =   "frmFind.frx":0BC1
      Left            =   1950
      List            =   "frmFind.frx":0BC8
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1335
      Width           =   1380
   End
   Begin VB.ComboBox cboFindIn 
      Height          =   315
      ItemData        =   "frmFind.frx":0BD9
      Left            =   1950
      List            =   "frmFind.frx":0BE6
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   825
      Width           =   1950
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Height          =   360
      Left            =   6525
      TabIndex        =   4
      Top             =   840
      Width           =   1230
   End
   Begin VB.CommandButton ReplaceButton 
      Caption         =   "Replace"
      Height          =   360
      Left            =   6525
      TabIndex        =   5
      Top             =   1830
      Width           =   1230
   End
   Begin VB.CommandButton ReplaceAllButton 
      Caption         =   "Replace All"
      Height          =   360
      Left            =   6525
      TabIndex        =   6
      Top             =   2220
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   6525
      TabIndex        =   7
      Top             =   2730
      Width           =   1230
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Case sensitive"
      Height          =   240
      Left            =   1950
      TabIndex        =   2
      Top             =   3300
      Value           =   1  'Checked
      Width           =   1380
   End
   Begin VB.CheckBox chkWhole 
      Caption         =   "Whole word only"
      Height          =   240
      Left            =   1950
      TabIndex        =   3
      Top             =   3660
      Width           =   1545
   End
   Begin VB.TextBox txtFindwhat 
      Height          =   795
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1335
      Width           =   2880
   End
   Begin VB.TextBox txtReplace 
      Height          =   750
      Left            =   1950
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2325
      Width           =   4290
   End
   Begin MSComctlLib.ListView lsvResults 
      Height          =   1470
      Left            =   975
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4500
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   2593
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgResult"
      SmallIcons      =   "imgResult"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Matched Text"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Results:"
      Height          =   195
      Left            =   990
      TabIndex        =   17
      Top             =   4200
      Width           =   585
   End
   Begin VB.Line Line3 
      X1              =   660
      X2              =   8330
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   615
      X2              =   8330
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   600
      Picture         =   "frmFind.frx":0C21
      Top             =   195
      Width           =   240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find and Replace"
      Height          =   195
      Left            =   930
      TabIndex        =   14
      Top             =   210
      Width           =   1230
   End
   Begin VB.Line Line1 
      X1              =   495
      X2              =   8165
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   450
      X2              =   8165
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options:"
      Height          =   195
      Left            =   1200
      TabIndex        =   13
      Top             =   3300
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find in:"
      Height          =   195
      Left            =   1290
      TabIndex        =   12
      Top             =   885
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find what:"
      Height          =   195
      Left            =   1050
      TabIndex        =   11
      Top             =   1335
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace with:"
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   2325
      Width           =   975
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Msearched() As String 'List of completed documents
Private Mpos As Long 'Position of find
Private Mflags As Integer 'Flags of case sensitive/wholeword
Private Meditor As Editor 'Global editor
Private Mreplacement As Long 'No of replacements
Private Mrtb As RichTextBox 'For hidden replacement

Private Sub cboFindIn_Click()
Dim lFolder As String
Dim lFolders As Folders
Dim lFso As New FileSystemObject
  If cboFindIn.ListIndex = 0 Then 'Current opened document
    If Not frmEditor.tabMain.SelectedTab Is Nothing Then
      cmdBrowse.Visible = False
      txtFind.Tag = ""
      txtFind.Locked = True
      txtFind.Text = Replace(frmEditor.tabMain.SelectedTab.Caption, "*", "")
    End If
  ElseIf cboFindIn.ListIndex = 1 Then 'Current local site,not work for remote site
    If frmEditor.tvSiteFiles.Nodes.Count > 0 Then
      If frmEditor.tvSiteFiles.Nodes("F0").Tag <> "R" Then
        Set Meditor = Nothing
        cmdBrowse.Visible = False
        txtFind.Locked = True
        txtFind.Text = frmEditor.tvSiteFiles.Nodes("F0").Text
        txtFind.Tag = txtFind.Text
      Else
        cboFindIn.ListIndex = 0
      End If
    End If
  ElseIf cboFindIn.ListIndex = 2 Then 'Selected folder
    cmdBrowse.Visible = True
    txtFind.Text = ""
    txtFind.Tag = ""
    txtFind.Locked = False
    If frmEditor.tvSiteFiles.Nodes.Count > 0 Then
      If frmEditor.tvSiteFiles.Nodes("F0").Tag = "R" Then
        cboFindIn.ListIndex = 0
        Exit Sub
      End If
    End If
    
  End If
End Sub

Private Sub chkCase_Click()
  Mflags = chkCase.Value * 4 + chkWhole.Value * 2
End Sub

Private Sub chkWhole_Click()
  Mflags = chkCase.Value * 4 + chkWhole.Value * 2
End Sub

Private Sub cmdBrowse_Click()
Dim lFolder As String
  lFolder = BrowseForFolder(Me.hwnd, "Select Folder")
  If lFolder <> "" Then
    Set Meditor = Nothing
    txtFind.Text = lFolder
  End If
End Sub

Private Sub cmdFindAll_Click()
  lsvResults.ListItems.Clear
  Do Until FindWords(txtFindwhat.Text, cboFindIn.ListIndex, True, True) = -1
    'loop through all folders/files
  Loop
  If cboFindIn.ListIndex > 0 Then
    Call FindWords(txtFindwhat.Text, cboFindIn.ListIndex, True, True) 'For complete
  End If
End Sub

Public Sub FindNextButton_Click()
  lsvResults.ListItems.Clear
  FindWords txtFindwhat.Text, cboFindIn.ListIndex
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  cboFindIn.ListIndex = Mfindin
  cboFindWhat.ListIndex = 0
  txtFindwhat.Text = Mfindwhat
  txtReplace.Text = Mreplace
  chkCase.Value = Mcasesensitive
  chkWhole.Value = Mwholeword
  Mpos = 1
  Mflags = chkCase.Value * 4 + chkWhole.Value * 2
  tmr.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Meditor = Nothing
  Mfindin = cboFindIn.ListIndex
  Mfindwhat = txtFindwhat.Text
  Mreplace = txtReplace.Text
  Mcasesensitive = chkCase.Value
  Mwholeword = chkWhole.Value
End Sub

Private Sub lsvResults_DblClick()
Dim lFile As String
Dim lSite As clsSite
Dim lVirtual As String
Dim lURL As String
  If Not lsvResults.SelectedItem Is Nothing Then
    'Get the file
    lFile = IIf(txtFind.Text = lsvResults.SelectedItem.Text, txtFind.Text, txtFind.Text & lsvResults.SelectedItem.Text)
    If S102_File_Exists(lFile) Or (txtFind.Text = lsvResults.SelectedItem.Text) Then
      'site files, fill the url/virtual path
      If cboFindIn.ListIndex = 1 Then
        Set lSite = Msitedetails.Item(txtFind.Tag)
        If Not lSite Is Nothing Then
          lVirtual = lSite.LocalPath
          lURL = lSite.URL
        End If
      End If
      'open the documents
      If txtFind.Text <> lsvResults.SelectedItem.Text Then
        frmEditor.LoadDocument lFile, , lVirtual, , lURL
      End If
      'highlight the searched word
      'Selected item has the tag value as selstart
      If Not frmEditor.tabMain.SelectedTab Is Nothing Then
        If lFile = frmEditor.tabMain.SelectedTab.Key Then
          If lsvResults.SelectedItem.Tag <> "" Then
            frmEditor.RTB(val(frmEditor.tabMain.SelectedTab.Tag)).SelStart = val(lsvResults.SelectedItem.Tag)
            frmEditor.RTB(val(frmEditor.tabMain.SelectedTab.Tag)).SelLength = Len(lsvResults.SelectedItem.ListSubItems(1))
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub ReplaceButton_Click()
  ReplaceWord txtFindwhat.Text, txtReplace.Text
End Sub

Private Sub ReplaceAllButton_Click()
  lsvResults.ListItems.Clear
  ReplaceAll txtFindwhat.Text, txtReplace.Text, cboFindIn.ListIndex
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub

Private Sub tmr_Timer()
  On Error Resume Next
  txtFindwhat.SetFocus
  tmr.Enabled = False
End Sub

Private Function FindWords(ByVal pWord As String, Optional ByVal pSearchFor As Integer, Optional ByVal pMsg As Boolean = True, Optional ByVal pAddToList As Boolean) As Long
'
'Search for: current doc/local site
'
Dim lPos As Long
Dim lNode As Object
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  If Meditor Is Nothing Then GetSearchDocument pWord, txtFind.Text, pSearchFor
  If Not Meditor Is Nothing Then
    lPos = FindWord(Meditor, pWord)
    If pSearchFor = 0 Then
      If lPos = -1 Then
        If pMsg Then MsgBox "The specified document has been searched. " & IIf(Mreplacement > 0, Mreplacement & " replacement(s) was made.", ""), vbOKOnly + vbInformation, Mtitle
        FindWords = -1
      Else
        If pAddToList Then
          Set lNode = lsvResults.ListItems.Add(, , IIf(Meditor.FileName <> "", Replace(Meditor.FileName, txtFind.Text, "", , , vbTextCompare), Meditor.Key), , "RESULT")
          lNode.ListSubItems.Add , , Meditor.SelText
          lNode.Tag = lPos
          lNode.EnsureVisible
          lsvResults.Refresh
        End If
      End If
      FindWords = lPos
    Else
      If lPos = -1 Then
        If Not Meditor Is Nothing Then
          AddSearched Meditor.FileName
          Set Meditor = Nothing
        End If
        GetSearchDocument pWord, txtFind.Text, pSearchFor
        lPos = FindWord(Meditor, pWord)
      End If
      If lPos > -1 Then
        If pAddToList Then
          Set lNode = lsvResults.ListItems.Add(, , Replace(Meditor.FileName, txtFind.Text, "", , , vbTextCompare), , "RESULT")
          lNode.ListSubItems.Add , , Meditor.SelText
          lNode.Tag = lPos
          lNode.EnsureVisible
          lsvResults.Refresh
        End If
      End If
      FindWords = lPos
    End If
  Else
    If pMsg Then
      If pAddToList Then
        MsgBox "The specified local site has been searched. " & IIf(lsvResults.ListItems.Count > 0, lsvResults.ListItems.Count & " item(s) found.", ""), vbInformation + vbOKOnly
        If lsvResults.ListItems.Count > 0 Then lsvResults.ListItems(1).EnsureVisible
      Else
        MsgBox "The specified local site has been searched.", vbInformation + vbOKOnly, Mtitle
      End If
    End If
    Mpos = -1
    ReDim Msearched(0)
    FindWords = -1
  End If
  Screen.MousePointer = vbDefault
End Function

Private Sub AddSearched(ByVal pFilename As String)
'
'Collect the completed document
'
Dim lCount As Integer
Dim li As Integer
  On Error Resume Next
  lCount = UBound(Msearched) + 1
  For li = 0 To lCount - 1
    If Msearched(li) = pFilename Then
      Exit Sub
    End If
  Next
  ReDim Preserve Msearched(lCount)
  Msearched(lCount) = pFilename
End Sub

Private Sub RemoveSearched(ByVal pFilename As String)
'
'Remove the completed document from collection when close the docuemnt
'
Dim lCount As Integer
Dim lRemove As Boolean
Dim li As Integer
  On Error Resume Next
  lCount = UBound(Msearched) + 1
  For li = 0 To lCount - 1
    If Msearched(li) = pFilename Then
      lRemove = True
    Else
      If lRemove Then
        Msearched(li - 1) = Msearched(li)
      End If
    End If
  Next
  ReDim Preserve Msearched(lCount - 1)
End Sub

Private Function FindWord(ByRef RTB As Editor, ByVal pWord As String) As Long
'
'Find the word as given
'
  If Mpos = -1 Then Mpos = 0
  Mpos = RTB.FindWord(pWord, Mpos + 1, Mflags)
  FindWord = Mpos
End Function

Private Function FileInSite(ByVal pFile As String) As Boolean
'
'Check for files in site
'
Dim lNode As Object
  On Error GoTo Cerr
  Set lNode = frmEditor.tvSiteFiles.Nodes(pFile) 'Ucase Changed
  FileInSite = True
  Exit Function
Cerr:
  FileInSite = False
End Function

Private Function FileInCompleted(ByVal pFilename As String) As Boolean
'
'Check for document is completed
'
Dim lCount As Integer
Dim li As Integer
  On Error GoTo Cerr
  lCount = UBound(Msearched) + 1
  For li = 0 To lCount - 1
    If Msearched(li) = pFilename Then
      FileInCompleted = True
      Exit Function
    End If
  Next
Cerr:
  FileInCompleted = False
End Function

Private Function WordInDocument(ByVal pDocument As String, ByVal pWord As String) As Boolean
'
'Test for word in document
'
Dim lPos As Long
  On Error GoTo Cerr
  rtbFind.LoadFile pDocument
  lPos = rtbFind.Find(pWord, 1, , Mflags)
  If lPos > -1 Then
    WordInDocument = True
  Else
    WordInDocument = False
  End If
  Exit Function
Cerr:
  WordInDocument = False
End Function

Private Function ReplaceWord(ByVal pWord As String, ByVal pReplace As String, Optional ByVal pMsg As Boolean = True, Optional ByVal pAddToList As Boolean) As Long
'
'Replace word
'
Dim lPos As Long
Dim lNode As Object
  lPos = -1
  If Not Meditor Is Nothing Then
    'If pWord <> Meditor.SelText Then
      lPos = FindWords(pWord, , pMsg)
    'End If
    If lPos > -1 Then
      If pAddToList Then
        Set lNode = lsvResults.ListItems.Add(, , IIf(Meditor.FileName <> "", Replace(Meditor.FileName, txtFind.Text, "", , , vbTextCompare), Meditor.Key), , "RESULT")
        lNode.ListSubItems.Add , , pWord
        lNode.Tag = lPos
        lNode.EnsureVisible
        lsvResults.Refresh
      End If
      Meditor.ReplaceWord pReplace
      Meditor.Changed = True
      Mpos = Mpos + Len(pReplace) 'mpos is used to find next
      'lPos = FindWords(pWord, , pMsg)
    End If
    ReplaceWord = Mpos
  End If
End Function

Private Sub ReplaceAll(ByVal pWord As String, ByVal pReplace As String, Optional ByVal pSearchFor As Integer)
'
'Replace all words
'
Dim lPos As Long
Dim lFile As String
Dim lExt As String
Dim lSite As String
Dim lNode As Object
Dim lSearchedWord As String
Dim li As Long
  If pSearchFor = 1 Or pSearchFor = 2 Then
    ReDim Msearched(0)
    
    If MsgBox("Operation cannot undo the replacements of closed documents" & vbCrLf & "Do you want to proceed?", vbQuestion + vbYesNo, Mtitle) = vbYes Then
      Screen.MousePointer = vbHourglass
      
      'From Tab(opened documents)
      For li = 1 To frmEditor.tabMain.Tabs.Count
        lFile = frmEditor.tabMain.Tabs.Item(li).Key
        If FileInSite(lFile) And FileInCompleted(lFile) = False Then
          frmEditor.tabMain.Tabs.Item(li).Selected = True
          Set Meditor = frmEditor.RTB(val(frmEditor.tabMain.Tabs.Item(li).Tag))
          lPos = 0
          Do Until lPos = -1
            lPos = ReplaceWord(pWord, pReplace, False, True)
          Loop
          AddSearched Meditor.FileName
        End If
      Next
      'From path
      If pSearchFor > 0 Then
        ReplaceAllFolder pWord, pReplace, txtFind.Text
      End If
      Screen.MousePointer = vbDefault
    End If
    MsgBox "Finished! " & lsvResults.ListItems.Count & " items found. " & lsvResults.ListItems.Count & " items replaced. ", vbInformation, Mtitle
    If lsvResults.ListItems.Count > 0 Then lsvResults.ListItems(1).EnsureVisible
  Else
    Screen.MousePointer = vbHourglass
    If Meditor Is Nothing Then Set Meditor = frmEditor.RTB(frmEditor.GetActiveRTB)
    Do Until lPos = -1
      lPos = ReplaceWord(pWord, pReplace, True, True)
      Mreplacement = Mreplacement + 1
    Loop
    Screen.MousePointer = vbDefault
  End If
  Mreplacement = 0
End Sub

Private Function ReplaceAndSave(ByVal pWord As String, ByVal pReplace As String) As Long
'
'Replace the word and save the file hidden (returns the noof replacements)
'
Dim lPos As Long
Dim Lpos1 As Long
Dim Lreplacement As Long
Dim lSearchedWord As String
Dim lNode As Object
  If Not Mrtb Is Nothing Then
    Do Until lPos = -1
      lPos = Mrtb.Find(pWord, lPos + 1, , Mflags)
      If lPos > -1 Then
        lSearchedWord = Mrtb.SelText
        Mrtb.SelText = pReplace
        Lpos1 = lPos
        lPos = lPos + Len(pReplace)
        Lreplacement = Lreplacement + 1
        Set lNode = lsvResults.ListItems.Add(, , Replace(Mrtb.FileName, txtFind.Text, "", , , vbTextCompare), , "RESULT")
        lNode.ListSubItems.Add , , lSearchedWord
        lNode.Tag = Lpos1
        lNode.EnsureVisible
        lsvResults.Refresh
      End If
    Loop
    Mrtb.SaveFile Mrtb.FileName, rtfText
    ReplaceAndSave = Lreplacement
  End If
End Function

Private Function GetSearchDocument(ByVal pWord As String, Optional ByVal pPath As String, Optional ByVal pSearchFor As Integer) As Boolean
'
'Get document for search; it it is nt opened,open and search
'
Dim lFso As New FileSystemObject
Dim lFolders As Folders
Dim lFolder As Folder
Dim lFile As File
Dim lName As String
Dim lExt As String
Dim lSite As clsSite
  If S107_Folder_Exist(pPath) Then
    'First get all files
    For Each lFile In lFso.GetFolder(pPath).Files
      'Not allow all files, only supported files
      lName = lFile.Path
      lExt = Mid(lName, InStrRev(lName, ".") + 1)
      If UCase(lExt) = "ASP" Or UCase(lExt) = "HTML" Or UCase(lExt) = "HTM" Or UCase(lExt) = "JS" Or UCase(lExt) = "INI" Or UCase(lExt) = "TXT" Or UCase(lExt) = "XML" Then
        'Check already searched
        If FileInCompleted(lName) = False Then
          'Check for search word is present
          If WordInDocument(lName, pWord) Then
            'If file is in site/not
            If pSearchFor = 1 Then
              Set lSite = Msitedetails.Item(frmEditor.cboSites.Text)
              If Not lSite Is Nothing Then
                frmEditor.LoadDocument lName, , lSite.LocalPath, , lSite.URL
              End If
            ElseIf pSearchFor = 2 Then
              frmEditor.LoadDocument lName
            End If
            Set Meditor = frmEditor.RTB(frmEditor.tabMain.SelectedTab.Tag)
            GetSearchDocument = True
            Exit For
          End If
        End If
      End If
    Next
    'Next subfolders
    If GetSearchDocument = False Then
      Set lFolders = lFso.GetFolder(pPath).SubFolders
      For Each lFolder In lFolders
        GetSearchDocument = GetSearchDocument(pWord, lFolder.Path, pSearchFor)
        If GetSearchDocument Then Exit For
      Next
    End If
  Else 'if it is searched for current doucment
    If Not frmEditor.tabMain.SelectedTab Is Nothing Then
      Set Meditor = frmEditor.RTB(frmEditor.tabMain.SelectedTab.Tag)
      GetSearchDocument = True
    End If
  End If
End Function

Private Sub ReplaceAllFolder(ByVal pWord As String, ByVal pReplace As String, Optional ByVal pPath As String)
'
'Replace all search words hiddenly
'
Dim lFso As New FileSystemObject
Dim lFolders As Folders
Dim lFolder As Folder
Dim lFile As File
Dim lName As String
Dim lExt As String
Dim lSite As String
  If S107_Folder_Exist(pPath) Then
    If frmEditor.tvSiteFiles.Nodes.Count > 0 Then lSite = frmEditor.tvSiteFiles.Nodes(1).Text
    'First get all files
    For Each lFile In lFso.GetFolder(pPath).Files
      'Not allow all files, only supported files
      lName = lFile.Path
      lExt = Mid(lName, InStrRev(lName, ".") + 1)
      If UCase(lExt) = "ASP" Or UCase(lExt) = "HTML" Or UCase(lExt) = "HTM" Or UCase(lExt) = "JS" Or UCase(lExt) = "INI" Or UCase(lExt) = "TXT" Or UCase(lExt) = "XML" Then
        'Check already searched
        If FileInCompleted(lName) = False Then
          'Check for search word is present
          If WordInDocument(lName, pWord) Then
            'Replace the words
            Set Mrtb = rtbFind
            Mrtb.LoadFile lName, rtfText
            Call ReplaceAndSave(pWord, pReplace)
            AddSearched lName
          End If
        End If
      End If
    Next
    'Next subfolders
    Set lFolders = lFso.GetFolder(pPath).SubFolders
    For Each lFolder In lFolders
      ReplaceAllFolder pWord, pReplace, lFolder.Path
    Next
  End If
End Sub

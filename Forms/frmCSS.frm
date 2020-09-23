VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCSS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CSS Editor"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "frmCSS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Edit code"
      TabPicture(0)   =   "frmCSS.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvProps"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tvCSS"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imlIcons"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pp"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmNew"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmRem"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmNewProp"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmOK"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmNo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Preview"
      TabPicture(1)   =   "frmCSS.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "IE1"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmNo 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6255
         TabIndex        =   12
         Top             =   4185
         Width           =   825
      End
      Begin VB.CommandButton cmOK 
         Caption         =   "S&ave"
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   4185
         Width           =   825
      End
      Begin VB.CommandButton Command1 
         Caption         =   "R&emove"
         Height          =   375
         Left            =   3915
         TabIndex        =   10
         Top             =   4185
         Width           =   1005
      End
      Begin VB.CommandButton cmNewProp 
         Caption         =   "New &Property..."
         Height          =   375
         Left            =   2430
         TabIndex        =   9
         Top             =   4185
         Width           =   1455
      End
      Begin VB.CommandButton cmRem 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   3690
         Width           =   870
      End
      Begin VB.CommandButton cmNew 
         Caption         =   "Add &New ID..."
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   3690
         Width           =   1320
      End
      Begin VB.PictureBox pp 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   90
         ScaleHeight     =   255
         ScaleWidth      =   2190
         TabIndex        =   2
         Top             =   4275
         Width           =   2250
         Begin VB.TextBox txData 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   270
            TabIndex        =   3
            Top             =   30
            Width           =   1905
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   0
            Picture         =   "frmCSS.frx":107A
            Top             =   15
            Width           =   240
         End
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   2925
         Top             =   1755
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
               Picture         =   "frmCSS.frx":1404
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCSS.frx":19A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCSS.frx":1F3C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvCSS 
         Height          =   3255
         Left            =   90
         TabIndex        =   1
         Top             =   375
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   5741
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   317
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "imlIcons"
         Appearance      =   1
      End
      Begin SHDocVwCtl.WebBrowser IE1 
         Height          =   4275
         Left            =   -74955
         TabIndex        =   4
         Top             =   360
         Width           =   7110
         ExtentX         =   12541
         ExtentY         =   7541
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
      Begin MSComctlLib.ListView lvProps 
         Height          =   3660
         Left            =   2385
         TabIndex        =   8
         Top             =   375
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   6456
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Property ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   5106
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item Data:"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   4095
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmCSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim globalPos As Long, globalPos2 As Long

Private Sub cmNew_Click()
On Error Resume Next
If tvCSS.SelectedItem Is Nothing Then tvCSS.SelectedItem = tvCSS.Nodes(1)
If tvCSS.SelectedItem.Image = 2 Then tvCSS.SelectedItem = tvCSS.SelectedItem.parent
tvCSS.Nodes.Add tvCSS.SelectedItem.Key, tvwChild, "tmp", , 2
tvCSS.SelectedItem = tvCSS.Nodes("tmp")
tvCSS.Nodes("tmp").Key = ""
If tvCSS.SelectedItem.parent.Key = "tags" Then tvCSS.SelectedItem.Text = "New1"
If tvCSS.SelectedItem.parent.Key = "classes" Then tvCSS.SelectedItem.Text = ".class1"
If tvCSS.SelectedItem.parent.Key = "uids" Then tvCSS.SelectedItem.Text = "#UID1"
tvCSS.SetFocus
tvCSS.StartLabelEdit
End Sub

Private Sub cmNewProp_Click()
On Error Resume Next
If tvCSS.SelectedItem.Image = 1 Then Exit Sub
lvProps.ListItems.Add lvProps.ListItems.Count + 1, , "property-name", , 3
lvProps.SelectedItem = lvProps.ListItems(lvProps.ListItems.Count)
lvProps.SelectedItem.ListSubItems.Add 1
lvProps.SetFocus
lvProps.StartLabelEdit
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
On Error Resume Next
Dim i As Long, tmp As String
frmEditor.RTB(frmEditor.GetCurrentIndex).SelStart globalPos - 1
frmEditor.RTB(frmEditor.GetCurrentIndex).SelLength globalPos2 - globalPos
tmp = vbCrLf
For i = 1 To tvCSS.Nodes.Count
If tvCSS.Nodes(i).Image = 2 Then tmp = tmp & tvCSS.Nodes(i).Key & vbCrLf
Next i
frmEditor.RTB(frmEditor.GetCurrentIndex).Paste tmp
Unload Me
End Sub

Private Sub cmRem_Click()
On Error Resume Next
If tvCSS.SelectedItem.Image = 2 Then tvCSS.Nodes.Remove tvCSS.SelectedItem.Key Else MsgBox "You cannot 'remove' a category. Select the tag" & vbCrLf & "you want to remove, and then click 'Remove'.", vbInformation
tvCSS.SetFocus
End Sub

Private Sub Command1_Click()
On Error Resume Next
lvProps.ListItems.Remove lvProps.SelectedItem.Index
lvProps.SetFocus
End Sub

Private Sub Form_Load()
tvCSS.Nodes.Add , , "tags", "HTML Tags", 1
tvCSS.Nodes("tags").Expanded = True
tvCSS.Nodes.Add , , "uids", "Unique IDs", 1
tvCSS.Nodes.Add , , "classes", "Classes", 1
LoadCSSData
End Sub

Private Sub lvProps_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error Resume Next
txData.Text = lvProps.SelectedItem.ListSubItems(1).Text
txData.SetFocus
End Sub

Private Sub lvProps_DblClick()
On Error Resume Next
txData.SetFocus
End Sub

Private Sub lvProps_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
txData.Text = Item.ListSubItems(1).Text
End Sub

Private Sub lvProps_LostFocus()
On Error Resume Next
Dim whole As String, i As Long
whole = tvCSS.Nodes(CLng(lvProps.Tag)).Text & Space(4) & "{"
For i = 1 To lvProps.ListItems.Count
whole = whole & lvProps.ListItems(i).Text & ": " & lvProps.ListItems(i).ListSubItems(1).Text & "; "
Next i
whole = whole & "}"
tvCSS.Nodes(CLng(lvProps.Tag)).Key = whole
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Dim sText As String
If SSTab1.Tab = 1 Then
  If Ext(Caption) = "css" Then
    Open FullPath(App.Path, "cssview.html") For Binary As #1
    sText = Space$(LOF(1))
    Get #1, , sText
    sText = Replace(sText, "foobar.css", Caption)
    Close #1
    
    Open FullPath(App.Path, "tmpcss.html") For Output As #1
    Print #1, sText
    Close #1
    IE1.Navigate FullPath(App.Path, "tmpcss.html")
  Else
    IE1.Navigate Caption
  End If
IE1.SetFocus
Else
lvProps.SetFocus
End If
End Sub

Private Sub tvCSS_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error Resume Next
If tvCSS.SelectedItem.Image = 1 Then Cancel = True: Exit Sub
End Sub

Private Sub tvCSS_DblClick()
On Error Resume Next
lvProps.SetFocus
End Sub

Private Sub tvCSS_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
lvProps.ListItems.Clear
lvProps.Tag = tvCSS.SelectedItem.Index
Dim props() As String, whole As String, vals() As String
Dim pos As Long, pos2 As Long, i As Long, i2 As Long
pos = InStr(Node.Key, "{")
If pos = 0 Then Exit Sub
pos2 = InStr(pos + 1, NoStrings(Node.Key), "}")
If pos2 = 0 Then Exit Sub
pos = pos + 1
whole = Mid$(Node.Key, pos, pos2 - pos)
whole = Trim(whole)
props = Split(whole, ";")
For i = 0 To UBound(props)
props(i) = Trim(props(i))
vals = Split(props(i), ":")
vals(0) = Trim(vals(0))
vals(1) = Trim(vals(1))
lvProps.ListItems.Add , vals(0), vals(0), , 3
lvProps.ListItems(vals(0)).ListSubItems.Add 1
lvProps.ListItems(vals(0)).ListSubItems(1).Text = vals(1)
Next i
txData.Text = ""
End Sub

Private Sub txData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
lvProps.SetFocus
lvProps.SelectedItem.ListSubItems(1).Text = txData.Text
End If
End Sub

Sub LoadCSSData()
On Error Resume Next
Dim pos As Long, pos2 As Long
Dim lines() As String, whole As String
pos = InStr(1, frmEditor.RTB(frmEditor.GetCurrentIndex).Text, "<STYLE", vbTextCompare)
pos2 = InStr(pos + 1, frmEditor.RTB(frmEditor.GetCurrentIndex).Text, "</STYLE>", vbTextCompare)
If pos2 = 0 And pos > 0 Then MsgBox "Invalid file format.", vbCritical:  Exit Sub
If pos = 0 Then
pos = InStr(1, frmEditor.RTB(frmEditor.GetCurrentIndex).Text, "<!--")
pos2 = InStr(pos + 1, frmEditor.RTB(frmEditor.GetCurrentIndex).Text, "-->", vbTextCompare)
Else
pos = pos + InStr(pos + 1, frmEditor.RTB(frmEditor.GetCurrentIndex).Text, ">") + 1
GoTo existing
End If
If pos = 0 Or pos2 = 0 Then MsgBox "Invalid file format.", vbCritical:  Exit Sub
pos = pos + 4
pos2 = pos2 - 3
If pos <= 0 Or pos2 <= 0 Then MsgBox "Invalid file format.", vbCritical:  Exit Sub
existing:
whole = Mid$(frmEditor.RTB(frmEditor.GetCurrentIndex).Text, pos, pos2 - pos)
whole = Replace(whole, Chr(10), "")
lines() = Split(whole, Chr(13))
globalPos = pos
globalPos2 = pos2
For pos = 0 To UBound(lines)
If Trim(lines(pos)) = "" Then GoTo nxt
Dim parent As String
parent = "tags"
If Left(lines(pos), 1) = "#" Then parent = "uids"
If InStr(lines(pos), ".") > 0 Then parent = "classes"
tvCSS.Nodes.Add parent, tvwChild, lines(pos), GetCSSElementName(lines(pos)), 2
nxt:
Next pos
End Sub

Function Ext(File As String) As String
'extension only
On Error Resume Next
Dim i As Long
i = InStr(StrReverse(File), ".")
If i = 0 Then Ext = File: Exit Function
Ext = Right(File, i - 1)
End Function

Function FullPath(lpPath As String, lpFile As String) As String
If Right(lpPath, 1) <> "\" Then lpPath = lpPath & "\"
FullPath = lpPath & lpFile
'fullpath, after resolving the "\" problems
End Function

Function NoStrings(Text As String) As String
'no strings
On Error Resume Next
Dim temp As String, tmps() As String, i As Long
temp = Text
tmps = Split(temp, Chr(34))
temp = ""
If UBound(tmps) = 0 Then GoTo nxt
For i = 0 To UBound(tmps)
If i And 1 Then tmps(i) = Space(Len(tmps(i)) + 2)
temp = temp & tmps(i)
Next i
tmps = Split(temp, "'")
nxt:
If UBound(tmps) = 0 Then GoTo finish
For i = 0 To UBound(tmps)
If i And 1 Then tmps(i) = Space(Len(tmps(i)) + 2)
temp = temp & tmps(i)
Next i
finish:
If temp = "" Then temp = Text
NoStrings = temp
End Function
Function GetCSSElementName(ByVal StrCSS As String) As String
Dim pos As Long
pos = InStr(StrCSS, "{")
If pos = 0 Then GetCSSElementName = StrCSS: Exit Function
StrCSS = Left(StrCSS, pos - 1)
StrCSS = Replace(Trim(StrCSS), vbTab, "")
GetCSSElementName = IIf(Left(StrCSS, 1) <> "#" And InStr(StrCSS, ".") = 0, UCase(StrCSS), StrCSS)
End Function


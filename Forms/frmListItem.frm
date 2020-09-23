VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmListItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List items"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   225
      Picture         =   "frmListItem.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdRemove 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   585
      Picture         =   "frmListItem.frx":00F1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   300
      Left            =   945
      Picture         =   "frmListItem.frx":0259
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3495
      TabIndex        =   1
      Top             =   2625
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   2370
      TabIndex        =   0
      Top             =   2625
      Width           =   975
   End
   Begin MSComctlLib.ListView lsvItems 
      Height          =   1965
      Left            =   225
      TabIndex        =   3
      Top             =   435
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   3466
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5767
      EndProperty
   End
End
Attribute VB_Name = "frmListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  lsvItems.ListItems.Add , , "Item " & lsvItems.ListItems.Count + 1
  lsvItems.ListItems(lsvItems.ListItems.Count).Selected = True
  lsvItems.SetFocus
  lsvItems.StartLabelEdit
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdEdit_Click()
  If Not lsvItems.SelectedItem Is Nothing Then
    lsvItems.SetFocus
    lsvItems.StartLabelEdit
  End If
End Sub

Private Sub cmdOk_Click()
Dim li As Integer
  Screen.MousePointer = vbHourglass
  frmList.lsvItems.Clear
  For li = 1 To lsvItems.ListItems.Count
    frmList.lsvItems.AddItem lsvItems.ListItems(li).Text
  Next
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  If Not lsvItems.SelectedItem Is Nothing Then
    lsvItems.ListItems.Remove lsvItems.SelectedItem.Index
  End If
End Sub

Private Sub Form_Load()
Dim li As Integer
  lsvItems.ColumnHeaders(1).Width = lsvItems.Width - Screen.TwipsPerPixelX * 22
  lsvItems.ListItems.Clear
  For li = 0 To frmList.lsvItems.ListCount - 1
    lsvItems.ListItems.Add , , frmList.lsvItems.List(li)
  Next
End Sub

Private Sub lsvItems_AfterLabelEdit(Cancel As Integer, NewString As String)
  If Trim(NewString) = "" Then Cancel = 1
End Sub

'
'Private Sub cmdOk_Click()
'Dim lTag As String
'Dim li As Integer
'  lTag = ""
'  Screen.MousePointer = vbHourglass
'  For li = 1 To lsvItems.ListItems.Count
'    lTag = lTag & vbTab & "<option value=" & li - 1 & ">" & lsvItems.ListItems(li).Text & "</option>" & vbCrLf
'  Next
'  If lTag = "" Then lTag = Tag & vbTab & "<option>" & "</option>" & vbCrLf
'  lTag = vbCrLf & vbTab & "<select name=" & IIf(Trim(txtName.Text) <> "", txtName.Text, "list1") & ">" & vbCrLf & lTag & vbTab & "</select>" & vbCrLf
'  frmEditor.RTB(frmEditor.GetActiveRTB).Paste lTag
'  Screen.MousePointer = vbDefault
'  Unload Me
'End Sub

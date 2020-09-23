VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRecordset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recordset"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecordset.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsvRecords 
      Height          =   5310
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9366
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MConnectionString As String
Public MConnectionName As String
Public MTable As String

Private Sub Form_Load()
  Me.Caption = Me.Caption & " [" & MTable & "]"
  LoadRecordset
End Sub

'
'User Function
'
Private Function LoadRecordset()
'
'Load records
'
Dim lCon As Object
Dim lRecordset As Object
Dim Litem As ListItem
Dim li As Integer
Dim lRecNo As Long
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  Set lCon = CreateObject("ADODB.Connection")
  lCon.CursorLocation = 3 'Useclient
  Call lCon.Open(MConnectionString)
  If lCon.State = 1 Then ' adStateOpen
    Set lRecordset = CreateObject("ADODB.Recordset") ' CreateObject("ADODB.Recordset") ' New ADODB.Recordset
    lRecordset.CursorType = 2 ' adOpenDynamic
    Set lRecordset = lCon.Execute("Select * from " & MTable)
    If lRecordset.Fields.Count > 0 Then
      'Build grid
      lsvRecords.ListItems.Clear
      lsvRecords.ColumnHeaders.Clear
      Call lsvRecords.ColumnHeaders.Add(, , "Record No") ', 1500, vbAlignLeft)
      For li = 0 To lRecordset.Fields.Count - 1
        Call lsvRecords.ColumnHeaders.Add(, , lRecordset.Fields(li).Name) ', 1500, vbAlignLeft)
      Next
    End If
    If lRecordset.RecordCount > 0 Then
      'Move values
      lRecNo = 1
      Do While lRecordset.EOF = False
        Set Litem = lsvRecords.ListItems.Add(, , lRecNo)
        For li = 0 To lRecordset.Fields.Count - 1
          Litem.ListSubItems.Add , , lRecordset.Fields(li).Value
        Next
        lRecordset.MoveNext
        lRecNo = lRecNo + 1
      Loop
      If lsvRecords.ListItems.Count > 0 Then
        lsvRecords.ListItems(1).Selected = True
        lsvRecords.ListItems(1).EnsureVisible
      End If
    End If
  End If
  Screen.MousePointer = vbDefault
End Function



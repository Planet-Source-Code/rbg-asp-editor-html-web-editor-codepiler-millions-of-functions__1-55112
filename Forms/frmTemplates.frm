VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTemplates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Document"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTemplates.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlTem 
      Left            =   3690
      Top             =   3075
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":000C
            Key             =   "ASP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":050B
            Key             =   "HTML"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkOpenDialog 
      Caption         =   "&Don't show this dialog in future"
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   6150
      Width           =   2670
   End
   Begin MSComctlLib.ListView lsvRecent 
      Height          =   735
      Left            =   660
      TabIndex        =   5
      Top             =   3915
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1296
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   6879
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   6255
      TabIndex        =   1
      Top             =   4995
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   6255
      TabIndex        =   2
      Top             =   5415
      Width           =   1110
   End
   Begin MSComctlLib.ListView lsvTemplates 
      Height          =   2850
      Left            =   240
      TabIndex        =   0
      Top             =   1995
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   5027
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   0
   End
   Begin VB.PictureBox picTop 
      Height          =   1305
      Left            =   90
      ScaleHeight     =   1245
      ScaleWidth      =   7440
      TabIndex        =   3
      Top             =   60
      Width           =   7500
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   -30
         Picture         =   "frmTemplates.frx":0A47
         Top             =   0
         Width           =   7485
      End
      Begin VB.Image imgTest 
         Height          =   210
         Left            =   1755
         Top             =   825
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   4440
      Left            =   75
      TabIndex        =   4
      Top             =   1500
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   7832
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Recent"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlFiles 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":6E2A
            Key             =   "MYCOMPUTER"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":725B
            Key             =   "FOLDERCLOSE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":73D6
            Key             =   "FOLDEROPEN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":7617
            Key             =   "DES"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":78AE
            Key             =   "HISTORY"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":7CF2
            Key             =   "HTODAY"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":7F82
            Key             =   "HPAST"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":8214
            Key             =   "EMPTY"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":848C
            Key             =   "ASP"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":86E9
            Key             =   "JS"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":8957
            Key             =   "HTM"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":8BB4
            Key             =   "HTML"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":8E11
            Key             =   "TXT"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":9084
            Key             =   "DOC"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":92F7
            Key             =   "INI"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":956A
            Key             =   "BAT"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":97DD
            Key             =   "DAT"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":9A50
            Key             =   "WAV"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":9BDC
            Key             =   "MP3"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":9D68
            Key             =   "MPG"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":9EF4
            Key             =   "AVI"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":A080
            Key             =   "MPEG"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":A20C
            Key             =   "CDA"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":A398
            Key             =   "DEFAULT"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":A624
            Key             =   "CSS"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":A7AB
            Key             =   "GIF"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":AA1F
            Key             =   "JPG"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":AC93
            Key             =   "BMP"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":AF07
            Key             =   "PNG"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":B17B
            Key             =   "ICO"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":B3EF
            Key             =   "TIF"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":B663
            Key             =   "JPEG"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":B8D7
            Key             =   "TIFF"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":BB4B
            Key             =   "PSD"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":BDBF
            Key             =   "MYN"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":BF5F
            Key             =   "MYD"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":C381
            Key             =   "FLOPPY"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":C7A5
            Key             =   "DISK"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":C9E9
            Key             =   "CDROM"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mLeft As Single

Private Sub cmdCancel_Click()
  frmEditor.Mdocumentype = -1
  lsvTemplates.ListItems(1).Selected = True
  tabMain.Tabs(1).Selected = True
  chkOpenDialog.Value = 0
  'Unload Me
  Me.Hide
End Sub

Public Sub cmdOk_Click()
  If tabMain.SelectedItem.Index = 1 Then
    If Not lsvTemplates.SelectedItem Is Nothing Then
      frmEditor.Mdocumentype = val(lsvTemplates.SelectedItem.Tag)
    End If
  Else
    If Not lsvRecent.SelectedItem Is Nothing Then
      frmEditor.Mrecentfile = lsvRecent.SelectedItem.Key
    End If
  End If
  lsvTemplates.ListItems(1).Selected = True
  SaveSetting App.Title, "OpenDialog", "OpenDialog", chkOpenDialog.Value
  frmEditor.Mopendialog = chkOpenDialog.Value
  tabMain.Tabs(1).Selected = True
  Me.Hide
  'Unload Me
End Sub

Private Sub Form_Load()
  lsvTemplates.Icons = imlTem
  mLeft = lsvTemplates.Left
  With lsvTemplates
    .ListItems.Add(, , "ASP Page", "ASP").Tag = 1
    .ListItems.Add(, , "HTML Page", "HTML").Tag = 0
  End With
  LoadRecent
  tabMain.Tabs(1).Selected = True
  lsvTemplates.ListItems(1).Selected = True
  chkOpenDialog.Value = IIf(frmEditor.Mopendialog, 1, 0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    frmEditor.Mdocumentype = -1
    Cancel = 1
    chkOpenDialog.Value = 0
    lsvTemplates.ListItems(1).Selected = True
    tabMain.Tabs(1).Selected = True
    Me.Hide
  End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lsvRecent.Width = lsvTemplates.Width
  lsvRecent.Height = lsvTemplates.Height
  lsvRecent.Top = lsvTemplates.Top
End Sub

Private Sub lsvRecent_DblClick()
  cmdOk_Click
End Sub

Private Sub lsvTemplates_DblClick()
  cmdOk_Click
End Sub

Private Function GetFileImg(ByVal pFilename As String) As String
Dim lTmp As String
  On Error GoTo Cerr
  lTmp = Mid(pFilename, InStrRev(pFilename, ".") + 1)
  imgTest.Picture = imlFiles.ListImages(UCase(lTmp)).Picture
  GetFileImg = UCase(lTmp)
  Exit Function
Cerr:
  GetFileImg = "DEFAULT"
End Function

Private Sub lsvTemplates_ItemClick(ByVal Item As MSComctlLib.ListItem)
'  cmdOk.SetFocus
End Sub

Private Sub tabMain_Click()
  If tabMain.SelectedItem.Index = 1 Then 'New
    lsvRecent.Left = -20000
    lsvTemplates.Left = mLeft
  Else 'Recent
    lsvTemplates.Left = -20000
    lsvRecent.Left = mLeft
  End If
End Sub

Public Function LoadRecent()
Dim li As Long
Dim lFile As String
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  lsvRecent.ListItems.Clear
  lsvRecent.SmallIcons = imlFiles
  For li = 1 To frmEditor.tvHistory.Nodes.Count
    If frmEditor.tvHistory.Nodes(li).Tag = "1" Then
      lFile = Split(frmEditor.tvHistory.Nodes(li).Key, "^")(1)
      If lFile <> "" Then
        lsvRecent.ListItems.Add(, lFile, Mid(lFile, InStrRev(lFile, "\") + 1), , GetFileImg(lFile)).ListSubItems.Add , , Mid(lFile, 1, InStrRev(lFile, "\") - 1)
      End If
    End If
  Next
  Screen.MousePointer = vbDefault
End Function

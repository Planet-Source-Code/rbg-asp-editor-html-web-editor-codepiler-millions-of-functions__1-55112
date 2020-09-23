VERSION 5.00
Begin VB.Form frmSites 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sites"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
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
   Icon            =   "frmSites.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3825
      Top             =   2100
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   3510
      TabIndex        =   4
      Top             =   1485
      Width           =   990
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   345
      Left            =   3510
      TabIndex        =   3
      Top             =   1035
      Width           =   990
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   345
      Left            =   3510
      TabIndex        =   2
      Top             =   585
      Width           =   990
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Default         =   -1  'True
      Height          =   345
      Left            =   3510
      TabIndex        =   1
      Top             =   135
      Width           =   990
   End
   Begin VB.ListBox lstSites 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   3255
   End
End
Attribute VB_Name = "frmSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mChangeList As Boolean
Public mSite As String
Private WithEvents mSiteFrm As frmSiteDetails
Attribute mSiteFrm.VB_VarHelpID = -1
Public Event SiteSaved(ByVal pInfo As String, ByVal pNew As Boolean)


Private Sub cmdEdit_Click()
  If lstSites.ListIndex >= 0 Then
    Set mSiteFrm = New frmSiteDetails
    mSiteFrm.LoadDetails lstSites.Text
    mSiteFrm.Show vbModal
  End If
End Sub

Private Sub cmdNew_Click()
  Set mSiteFrm = New frmSiteDetails
  mSiteFrm.LoadDetails ""
  mSiteFrm.Show vbModal
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  If lstSites.ListIndex > -1 Then
    If MsgBox("Are you sure to delete the Site '" & lstSites.List(lstSites.ListIndex) & "'?", vbYesNo + vbQuestion, Mtitle) = vbYes Then
      DeleteSite
    End If
  End If
End Sub

Private Sub Form_Load()
  LoadSites
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mSite = ""
  If mChangeList Then
    frmEditor.LoadSites
  End If
End Sub

Private Sub lstSites_DblClick()
  Call cmdEdit_Click
End Sub

Private Sub mSiteFrm_SiteSaved(ByVal pSite As String, ByVal pNew As Boolean)
  If pNew Then
    lstSites.AddItem pSite
    mChangeList = True
  Else
    mChangeList = True
    lstSites.List(lstSites.ListIndex) = pSite
  End If
End Sub

Private Sub tmr_Timer()
  tmr.Enabled = False
  Unload Me
End Sub

Private Function LoadSites()
Dim li As Long
   Screen.MousePointer = vbHourglass
   lstSites.Clear
   For li = 0 To frmEditor.cboSites.ListCount - 1
      If frmEditor.cboSites.List(li) <> "---------------" And frmEditor.cboSites.List(li) <> "Define sites..." Then
        lstSites.AddItem frmEditor.cboSites.List(li)
      End If
   Next
   Screen.MousePointer = vbDefault
End Function

Private Function DeleteSite()
  If lstSites.ListIndex >= 0 Then
    Msitedetails.Remove lstSites.Text
    lstSites.RemoveItem lstSites.ListIndex
    mChangeList = True
    RaiseEvent SiteSaved("", True)
  End If
End Function


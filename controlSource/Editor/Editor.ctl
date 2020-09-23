VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{202B977B-50CE-45BD-BAAE-BAFD6C4B3B19}#1.1#0"; "INTELL~1.OCX"
Begin VB.UserControl Editor 
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   9075
   Begin VB.PictureBox picResize 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   735
      ScaleHeight     =   75
      ScaleWidth      =   540
      TabIndex        =   12
      Top             =   3690
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   105
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9075
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4350
      Width           =   9075
      Begin VB.CheckBox chkApply 
         Caption         =   "Apply previous values"
         Height          =   225
         Left            =   6780
         TabIndex        =   9
         Top             =   30
         Visible         =   0   'False
         Width           =   2205
      End
      Begin MSComctlLib.ImageList imlTabs 
         Left            =   4155
         Top             =   -135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   93
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":0000
               Key             =   "DESIGNCLICK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":010C
               Key             =   "DESIGNHIDE"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":0250
               Key             =   "SPLITCLICK"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":0389
               Key             =   "SPLITHIDE"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":04F7
               Key             =   "CODEHIDE"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":05E6
               Key             =   "CODECLICK"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picTool 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5160
         ScaleHeight     =   300
         ScaleWidth      =   1455
         TabIndex        =   11
         Top             =   30
         Width           =   1455
         Begin VB.Shape shPointer 
            BorderColor     =   &H00808080&
            Height          =   255
            Left            =   1065
            Top             =   0
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Image imgBack 
            Height          =   240
            Left            =   30
            MouseIcon       =   "Editor.ctx":06D2
            MousePointer    =   99  'Custom
            Picture         =   "Editor.ctx":0F9C
            ToolTipText     =   "Back"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgForward 
            Height          =   240
            Left            =   330
            MouseIcon       =   "Editor.ctx":1079
            MousePointer    =   99  'Custom
            Picture         =   "Editor.ctx":1943
            ToolTipText     =   "Forward"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgStop 
            Height          =   240
            Left            =   705
            MouseIcon       =   "Editor.ctx":1A1F
            MousePointer    =   99  'Custom
            Picture         =   "Editor.ctx":22E9
            ToolTipText     =   "Stop"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgRefresh 
            Height          =   240
            Left            =   1080
            MouseIcon       =   "Editor.ctx":26EB
            MousePointer    =   99  'Custom
            Picture         =   "Editor.ctx":2FB5
            ToolTipText     =   "Refresh"
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Image imgSource 
         Height          =   270
         Left            =   0
         MouseIcon       =   "Editor.ctx":309D
         MousePointer    =   99  'Custom
         Picture         =   "Editor.ctx":3967
         Top             =   0
         Width           =   1395
      End
      Begin VB.Image imgView 
         Height          =   270
         Left            =   1290
         MouseIcon       =   "Editor.ctx":3A43
         MousePointer    =   99  'Custom
         Picture         =   "Editor.ctx":430D
         Top             =   0
         Width           =   1365
      End
      Begin VB.Shape spBottom 
         Height          =   15
         Left            =   3960
         Top             =   135
         Width           =   360
      End
      Begin VB.Image imgSplit 
         Height          =   270
         Left            =   2565
         MouseIcon       =   "Editor.ctx":4441
         MousePointer    =   99  'Custom
         Picture         =   "Editor.ctx":4D0B
         Top             =   0
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VB.PictureBox picSource 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   1095
      ScaleHeight     =   3675
      ScaleWidth      =   6765
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -150
      Width           =   6765
      Begin IntellProj.Intellisense IntellBox 
         Height          =   1095
         Left            =   555
         TabIndex        =   10
         Top             =   465
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   1931
      End
      Begin MSComctlLib.ImageList imlIbox 
         Left            =   1200
         Top             =   1920
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   11
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":4E69
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":4F08
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":5046
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlIbox_backup 
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
               Picture         =   "Editor.ctx":5184
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":5216
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Editor.ctx":529A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picLinesBack 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4680
         Left            =   810
         ScaleHeight     =   4680
         ScaleWidth      =   510
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   510
         Begin VB.PictureBox picLines 
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4680
            Left            =   0
            ScaleHeight     =   4680
            ScaleWidth      =   495
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   15
            Width           =   495
            Begin VB.Label txtNo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000003&
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000004&
               Height          =   1785
               Left            =   75
               TabIndex        =   8
               Top             =   30
               Width           =   375
            End
         End
      End
      Begin VB.PictureBox picSidebar 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4680
         Left            =   180
         ScaleHeight     =   4680
         ScaleWidth      =   405
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   105
         Width           =   405
         Begin VB.Shape spSidebar 
            BorderColor     =   &H00000000&
            Height          =   360
            Left            =   135
            Top             =   180
            Visible         =   0   'False
            Width           =   15
         End
      End
      Begin RichTextLib.RichTextBox RTB 
         Height          =   3420
         Left            =   750
         TabIndex        =   0
         Top             =   -15
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   6033
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"Editor.ctx":532C
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
      Begin VB.Shape spBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   255
         Left            =   1020
         Top             =   3555
         Width           =   600
      End
   End
   Begin VB.PictureBox picPreview 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   100
      ScaleHeight     =   1995
      ScaleWidth      =   2520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1125
      Width           =   2520
      Begin SHDocVwCtl.WebBrowser wbPreview 
         Height          =   930
         Left            =   135
         TabIndex        =   1
         Top             =   615
         Visible         =   0   'False
         Width           =   945
         ExtentX         =   1667
         ExtentY         =   1640
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
         Location        =   ""
      End
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   7860
      Picture         =   "Editor.ctx":53AC
      Top             =   3705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgResize 
      Height          =   75
      Left            =   720
      MousePointer    =   7  'Size N S
      Top             =   3810
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Shape spSide 
      BorderColor     =   &H00C0C0C0&
      Height          =   360
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Menu mmupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mmucut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mmucopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mmupaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mmusepo 
         Caption         =   "-"
      End
      Begin VB.Menu mmuselect 
         Caption         =   "&Select All"
      End
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Rem ===================================
Rem ===    A S P   E D I T O R   ======
Rem ===================================

Rem -----------------------------------
Rem        Private Declarations
Rem -----------------------------------
Private WithEvents mFrmConfirm As frmConfirm
Attribute mFrmConfirm.VB_VarHelpID = -1
Private bDirty As Boolean
Private mView As Integer '0-Single,1-Both
Private mRedoing As Boolean
Private mFindWord As String
Private mFindFlags As Integer
Private mPosition As Long
Private mPreviewed As Boolean
Private mLinesPerPage As Integer
Private mTmpFile As String
Private mFormValues As New Collection
Private mCSSFiles As String
Private mCSSTags As String
Private mTagCompleted As Boolean
Private mUndos As New clsChanges
Private mProceedRemoteChanges As Boolean 'Save the changes in remote before show preview
Private mDontShowConfirm As Boolean 'Dont show the confirm

Rem -----------------------------------
Rem         Properties Buffer
Rem -----------------------------------
Private mClasses As Long
Private mChange As Boolean
Private mCSSList As String
Private mFilename As String
Private mAppPath As String 'Opened file path
Private mCpPath As String 'Codepiler exe path
Private mLocalhost As String
Private mKey As String
Private mBlankpage As Integer
Private mLineno As Boolean
Private mVirtualpath As String
Private mWordwrap As Boolean
Private mAutoComplete As Boolean
Private mIntelisense As Boolean
Private mIsRemote As Boolean
Private mIsHistory As Boolean
Private mHeight As Single
Private mTitle As String
Private mFullmode As Boolean
Private mMode As Integer

Rem -----------------------------------
Rem             Subclassing
Rem -----------------------------------
Implements ISubclass
Private m_emr As EMsgResponse
Private bSubclassing As Boolean

Rem -----------------------------------
Rem             Constants
Rem -----------------------------------
Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114
Private Const WM_MOUSEWHEEL = &H20A
Private Const EM_GETTHUMB = &HBE
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETFIRSTVISIBLELINE = &HCE

Rem -----------------------------------
Rem               Events
Rem -----------------------------------
Public Event BrowserStatusChanged(ByVal pText As String)
Public Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event HScroll()
Public Event VScroll()
Public Event MouseWheel()
Public Event Position(ByVal X As Long, ByVal y As Long)
Public Event FileSaved(ByVal pFilename As String, ByVal pNew As Boolean)
Public Event StateChanged()
Public Event ProgressStatus(ByVal pValue As Integer, ByVal pMax As Integer, ByVal pStatus As String)
Public Event LockControl()
Public Event ReleaseControl()
Public Event Changed(ByVal pChange As Boolean)
Public Event DocumentOpened(ByVal pNew As Boolean)
Public Event ModeChanged(ByVal Mode As Integer)



Rem -----------------------------------
Rem           Actions
Rem -----------------------------------

Private Sub chkApply_Click()
Dim objDoc 'As HTMLDocument
Dim objItem ' As HTMLFormElement
Dim objCtl As Object
  If chkApply.value = 1 Then
    Set objDoc = wbPreview.Document
    For Each objItem In objDoc.Forms
      For Each objCtl In objItem.elements
        If ControlExists(objCtl.Name) = True Then
          If mFormValues(objCtl.Name) <> "" Then
            objCtl.value = mFormValues(objCtl.Name)
          End If
        End If
      Next
    Next
  End If
End Sub

Private Sub imgBack_Click()
  On Error GoTo Cerr
  wbPreview.GoBack
Cerr:
End Sub

Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If shPointer.Visible = False Then
    shPointer.Visible = True
  End If
  If shPointer.Left <> imgBack.Left - Screen.TwipsPerPixelX Then
    shPointer.Left = imgBack.Left - Screen.TwipsPerPixelX
  End If
End Sub

Private Sub imgForward_Click()
  On Error GoTo Cerr
  wbPreview.GoForward
Cerr:
End Sub

Private Sub imgForward_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If shPointer.Visible = False Then
    shPointer.Visible = True
  End If
  If shPointer.Left <> imgForward.Left - Screen.TwipsPerPixelX Then
    shPointer.Left = imgForward.Left - Screen.TwipsPerPixelX
  End If
End Sub

Private Sub imgRefresh_Click()
  On Error GoTo Cerr
  wbPreview.Refresh2
Cerr:
End Sub

Private Sub imgRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If shPointer.Visible = False Then
    shPointer.Visible = True
  End If
  If shPointer.Left <> imgRefresh.Left - Screen.TwipsPerPixelX Then
    shPointer.Left = imgRefresh.Left - Screen.TwipsPerPixelX
  End If
End Sub

Private Sub imgResize_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  picResize.Visible = True
  picResize.Left = imgResize.Left
  picResize.Width = imgResize.Width
  picResize.Top = imgResize.Top + y - (picResize.Height / 2)
End Sub

Private Sub imgResize_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim lRange As Single
Dim lTop As Single
  lRange = Screen.TwipsPerPixelY * 100 'Just range
  lTop = imgResize.Top + y - (picResize.Height / 2)
  If (lTop >= picSource.Top + lRange) And (lTop <= picFooter.Top - lRange) Then
    picResize.Top = lTop
  End If
End Sub

Private Sub imgResize_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  picResize.Visible = False
  mHeight = picResize.Top - picSource.Top
  UserControl_Resize
End Sub

Private Sub imgSplit_Click()
  Screen.MousePointer = vbHourglass
  If imgSplit.Tag = "SPLITHIDE" Then
    imgSplit.Picture = imlTabs.ListImages("SPLITCLICK").Picture
    imgSplit.Tag = "SPLITCLICK"
    imgSource.Picture = imlTabs.ListImages("CODEHIDE").Picture
    imgSource.Tag = "CODEHIDE"
    imgView.Picture = imlTabs.ListImages("DESIGNHIDE").Picture
    imgView.Tag = "DESIGNHIDE"
    imgSplit.Top = imgSource.Top
    imgSplit.ZOrder vbBringToFront
    ChangeMousepointer 3
    picSource.Left = spSide.Width
    picPreview.Left = spSide.Width
    mView = 1
    chkApply.Visible = True
    picTool.Visible = True
    UserControl_Resize
    mMode = 3
    RaiseEvent ModeChanged(3)
    ShowPreview
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub imgStop_Click()
  On Error GoTo Cerr
  wbPreview.Stop
Cerr:
End Sub

Private Sub imgStop_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If shPointer.Visible = False Then
    shPointer.Visible = True
  End If
  If shPointer.Left <> imgStop.Left - Screen.TwipsPerPixelX Then
    shPointer.Left = imgStop.Left - Screen.TwipsPerPixelX
  End If
End Sub

Private Sub IntellBox_Click()
  On Error Resume Next
  ProcessIntell vbKeyReturn
  RTB.SetFocus
End Sub

Private Sub IntellBox_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  ProcessIntell vbKeyReturn
  RTB.SetFocus
End Sub

Private Sub IntellBox_WordComleted(ByVal Completed As Boolean)
  mTagCompleted = Completed
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
'
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
  ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'// scroll events
    Select Case iMsg
    Case WM_VSCROLL
        Scrolltxtno
        RaiseEvent VScroll
    Case WM_MOUSEWHEEL
        Scrolltxtno
        RaiseEvent MouseWheel
    Case WM_HSCROLL
        RaiseEvent HScroll
    End Select
End Function

Private Sub mFrmConfirm_Confirm(ByVal Result As Boolean, ByVal DontShow As Boolean)
  mProceedRemoteChanges = Result
  mDontShowConfirm = DontShow
End Sub

Private Sub mmucopy_Click()
  Copy
End Sub

Private Sub mmucut_Click()
  Cut
End Sub

Private Sub mmupaste_Click()
  Paste
End Sub

Private Sub mmuselect_Click()
  SelectAll
End Sub

Private Sub picFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If shPointer.Visible Then shPointer.Visible = False
End Sub

Private Sub RTB_Click()
  Scrolltxtno
  RaisePosition
End Sub

Private Sub RTB_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lStr As String
  If KeyCode = vbKeyReturn Then KeyCode = 0
  RaisePosition
End Sub

Private Sub RTB_SelChange()
  'mUndos.CanUndo = True
  RaiseEvent StateChanged
End Sub

Private Sub tmr_Timer()
Dim lScroll As Long
Dim lVisibleRow As Long
Dim lAmount As Long
Dim li As Integer
Dim lStr As String
On Error Resume Next
  lScroll = SendMessage(RTB.hWnd, EM_GETTHUMB, 0&, 0&)
  lVisibleRow = SendMessage(RTB.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
  lAmount = ((picLines.TextHeight("A") / Screen.TwipsPerPixelY) * lVisibleRow) + (lVisibleRow * 3)
  lStr = ""
  For li = lVisibleRow + 1 To lVisibleRow + mLinesPerPage + 1
    lStr = lStr & li & vbCrLf
  Next
  txtNo.Caption = lStr

  picLines.Top = 0 - ((lScroll - lAmount) * Screen.TwipsPerPixelY)
  picLines.Height = picLinesBack.Height - picLines.Top
  tmr.Enabled = False
End Sub

Private Sub UserControl_ExitFocus()
  IntellBox.Visible = False
End Sub

Private Sub UserControl_Hide()
  IntellBox.Visible = False
End Sub

Private Sub UserControl_Initialize()
Dim lStr As String
  mLinesPerPage = picLines.Height / picLines.TextHeight("A") + 2
  imgView.Tag = "DESIGNHIDE"
  imgSource.Tag = "CODEHIDE"
  mUndos.Richtextbox = RTB
  wbPreview.Offline = False
  'load keywords
  IntellBox.SmallIcons = imlIbox
  lStr = StringForIntellbox("HTML")
  IntellBox.PopulateListFromString lStr, True
  'implements Subclass
  pAttachMessages
  ' setup and load!!
  RTB.SelIndent = 500
  RTB.HideSelection = False
  imgSource_Click
  MDoColor = True
  'trapUndo = True
  'testing
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  
  If imgResize.Visible Then imgResize.Visible = False
  If picResize.Visible Then picResize.Visible = False
  
  chkApply.Left = UserControl.Width - chkApply.Width - Screen.TwipsPerPixelX * 5
  picTool.Left = chkApply.Left - picTool.Width - Screen.TwipsPerPixelX * 10
  
  spBottom.Left = 0
  spBottom.Top = 0
  spBottom.Width = UserControl.ScaleWidth
  
  picSource.Width = UserControl.ScaleWidth - spSide.Width
  picSource.Top = 0
  If mView = 0 Then
    picSource.Height = UserControl.ScaleHeight - picFooter.Height
  Else
    picSource.Height = IIf(mHeight, mHeight, (UserControl.ScaleHeight / 2) - Screen.TwipsPerPixelY * 11)
  End If
  
  RTB.Left = Screen.TwipsPerPixelX ' picLinesBack.Width + Screen.TwipsPerPixelX
  RTB.Top = Screen.TwipsPerPixelY
  RTB.Width = picSource.Width - RTB.Left
  RTB.Height = picSource.Height - Screen.TwipsPerPixelY
  
  spBorder.Left = 0 'picLinesBack.Width
  spBorder.Top = 0
  spBorder.Width = RTB.Width
  spBorder.Height = RTB.Height
  
  picSidebar.Left = 0
  picSidebar.Top = 0
  picSidebar.Height = RTB.Height - Screen.TwipsPerPixelY * 17
  
  spSidebar.Left = picSidebar.Width - spSidebar.Width
  spSidebar.Top = 0
  spSidebar.Height = picSidebar.Height
  
  picLinesBack.Left = picSidebar.Left
  picLinesBack.Top = 0
  picLinesBack.Height = picSidebar.Height
  
  mLinesPerPage = picLinesBack.Height / picLinesBack.TextHeight("A") + 2
  
  picLines.Left = 0
  picLines.Height = picLinesBack.Height - picLines.Top
  
  txtNo.Top = 0
  txtNo.Left = 0 - Screen.TwipsPerPixelX * 2
  txtNo.Width = picLines.Width
  txtNo.Height = picLinesBack.Height + Screen.TwipsPerPixelY * 50
  
  picPreview.Width = picSource.Width - spSide.Width
  If mView = 0 Then
    picPreview.Top = 0
  Else
    If imgResize.Visible = False Then imgResize.Visible = True
    imgResize.Left = picPreview.Left
    imgResize.Width = picPreview.Width
    imgResize.Top = picSource.Height
    picPreview.Top = imgResize.Top + imgResize.Height
    picPreview.Left = picSource.Left
  End If
  picPreview.Height = (UserControl.ScaleHeight - picFooter.Height) - picPreview.Top
  
  wbPreview.Left = 0
  wbPreview.Top = 0
  wbPreview.Width = picPreview.Width
  wbPreview.Height = picPreview.Height - Screen.TwipsPerPixelY
  
  spSide.Top = 0
  spSide.Left = 0
  spSide.Height = UserControl.ScaleHeight
  
  IntellBox.Width = Screen.TwipsPerPixelX * 150
  IntellBox.Height = Screen.TwipsPerPixelY * 150
  
  Scrolltxtno
End Sub

Private Sub UserControl_Terminate()
  If FileExists(mAppPath & mTmpFile) Then
    Kill mAppPath & mTmpFile
  End If
  pDetachMessages
  'Set mUndoStack = Nothing
  'Set mRedoStack = Nothing
  mFilename = Empty
  mChange = Empty
End Sub

Private Sub imgSource_Click()
  If imgSource.Tag = "CODEHIDE" Then
    imgSource.Picture = imlTabs.ListImages("CODECLICK").Picture
    imgSource.Tag = "CODECLICK"
    imgView.Picture = imlTabs.ListImages("DESIGNHIDE").Picture
    imgView.Tag = "DESIGNHIDE"
    imgSplit.Picture = imlTabs.ListImages("SPLITHIDE").Picture
    imgSplit.Tag = "SPLITHIDE"
    imgSource.Top = imgView.Top
    imgSource.ZOrder vbBringToFront
    ChangeMousepointer 1
    picPreview.Left = -20000
    picSource.Left = spSide.Width
    chkApply.Visible = False
    picTool.Visible = False
    mView = 0
    mMode = 1
    UserControl_Resize
    RaiseEvent ModeChanged(1)
  End If
End Sub

Private Sub imgView_Click()
  Screen.MousePointer = vbHourglass
  If imgView.Tag = "DESIGNHIDE" Then
    imgView.Picture = imlTabs.ListImages("DESIGNCLICK").Picture
    imgView.Tag = "DESIGNCLICK"
    imgSource.Picture = imlTabs.ListImages("CODEHIDE").Picture
    imgSource.Tag = "CODEHIDE"
    imgSplit.Picture = imlTabs.ListImages("SPLITHIDE").Picture
    imgSplit.Tag = "SPLITHIDE"
    imgView.Top = imgSource.Top
    imgView.ZOrder vbBringToFront
    ChangeMousepointer 2
    picSource.Left = -20000
    picPreview.Left = spSide.Width
    mView = 0
    chkApply.Visible = True
    picTool.Visible = True
    If mFullmode = False Then UserControl_Resize
    mMode = 2
    RaiseEvent ModeChanged(2)
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
Dim lStr As String
    If mChange = False And KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight And KeyCode <> vbKeyHome And KeyCode <> vbKeyEnd And KeyCode <> vbKeyPageDown And KeyCode <> vbKeyPageUp And KeyCode <> vbKeyInsert And Shift = 0 Then
      mChange = True
      RaiseEvent Changed(True)
    End If
    mPreviewed = False
    ' ------------------------------
    ' here's the on the fly coloring
    ' ------------------------------
    On Error Resume Next
    If KeyCode = Asc(vbTab) Then
      If RTB.SelText <> "" Then
        If Shift = 0 Then
          Call Indent
        ElseIf Shift = 1 Then
          Call Outdent
        End If
      Else
        If IntellBox.Visible = True Then ProcessIntell vbKeyReturn: Exit Sub
        RTB.SelText = vbTab
      End If
      GoTo Goto_Ext
    End If
    
    If Shift = 1 And KeyCode = 57 Then
      If IntellBox.Visible = True Then ProcessIntell vbKeyReturn
    End If
    
    ' check for numbering
    If (KeyCode = vbKeyEnd And Shift = 2) Or _
        (KeyCode = vbKeyHome And Shift = 2) Or _
          KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Or _
            KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Then
      Scrolltxtno
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
      'PutLineNos
      Scrolltxtno
    End If
    
    
  
    ' check for search
    If KeyCode = vbKeyF3 Then
      'mFindWord = "df" 'for checking
      If mFindWord <> "" Then
        If mPosition < 0 Then
          MsgBox "The specified document has been searched", vbInformation, mTitle
          mPosition = mPosition + 1
        Else
          Findword mFindWord, mPosition + 1, mFindFlags
        End If
      End If
      GoTo Goto_Ext
    End If
    
    ' check for Ctrl+C
    If KeyCode = vbKeyC And Shift = 2 Then GoTo Goto_Ext
    
    ' check for text being pasted into the box
    If KeyCode = vbKeyV And Shift = 2 Then
        
        Screen.MousePointer = vbHourglass
        DoClipBoardPaste RTB
        Scrolltxtno
        Screen.MousePointer = vbNormal
        GoTo Goto_Ext
        
    End If
    
    ' check for preview
    If KeyCode = vbKeyF7 Then
      lStart = RTB.SelStart
      If mView = 1 And mPreviewed = False Then ShowPreview
      RTB.SelStart = lStart
      RTB.SetFocus
      GoTo Goto_Ext
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
            lStart = InStrRev(RTB.Text, ">", RTB.SelStart)
            lFinish = InStrRev(RTB.Text, "<", RTB.SelStart)
            If lStart > lFinish Then
              RTB.SelColor = vbBlack
            Else
              lFinish = InStrRev(RTB.Text, "<%", RTB.SelStart)
              If lStart < lFinish Then RTB.SelColor = vbBlack
            End If
            bDirty = True
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
        Case vbKeySpace
          If IntellBox.Visible Then
            If mTagCompleted Then
              IntellBox.Visible = False
            Else
              ProcessIntell vbKeyReturn
            End If
            Exit Sub
          End If
        End Select
    End If

    'Auto indent
    If KeyCode = vbKeyReturn And IntellBox.Visible = False Then
      lStr = GetIndent
      RTB.SelText = vbCrLf & lStr
    End If
  
Goto_Ext:
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub RTB_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  ProcessIntell KeyAscii
End Sub

Private Sub RTB_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'  RTB_KeyDown vbKeyReturn, 0
End Sub

Private Sub RTB_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = vbRightButton Then
    PopupMenu mmupopup, , RTB.Left + X, RTB.Top + y
  End If
End Sub

Private Sub wbPreview_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Dim objDoc 'As HTMLDocument
Dim objItem ' As HTMLFormElement
Dim objCtl As Object
Dim lOk As Boolean
  wbPreview.Visible = False
  lOk = False
  If CStr(PostData) <> "" Then
    Set objDoc = wbPreview.Document
    For Each objItem In objDoc.Forms
      For Each objCtl In objItem.elements
        If ControlExists(objCtl.Name) = False Then
          mFormValues.Add objCtl.value, objCtl.Name
        End If
      Next
      lOk = True
    Next
  End If
'  chkApply.Visible = lOk
End Sub

Private Sub wbPreview_GotFocus()
  If mPreviewed = False Then ShowPreview
End Sub

Private Sub wbPreview_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
  On Error Resume Next
  If LCase(URL) = "http:///" Then
    wbPreview.Visible = False
  Else
    wbPreview.Visible = True
  End If
  Screen.MousePointer = vbDefault
  RTB.SetFocus
  If mTmpFile <> "" Then Kill mAppPath & "\" & mTmpFile
End Sub

Private Sub wbPreview_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
  'wbPreview.Document.All.tags("UL").innerHTML = "<a href=''>Click here to go line no</a>"
End Sub

Private Sub wbPreview_StatusTextChange(ByVal Text As String)
  RaiseEvent BrowserStatusChanged(Text)
End Sub

Rem -----------------------------------
Rem             Properties
Rem -----------------------------------

Public Property Let AppPath(ByVal value As String)
'
'
'
  mCpPath = value
End Property

Public Property Get AppPath() As String
'
'
'
  AppPath = mCpPath
End Property

Public Property Let AutoComplete(ByVal value As Boolean)
'
'
'
  mAutoComplete = value
  IntellBox.AutoComplete = value
End Property

Public Property Get AutoComplete() As Boolean
'
'
'
  AutoComplete = mAutoComplete
End Property


Public Property Get CanCopy()
'
'is ready to copy
'
  CanCopy = RTB.SelLength
End Property

Public Property Get CanCut()
'
'is ready to cut
'
  CanCut = RTB.SelLength
End Property

Public Property Get CanPaste()
'
'is ready to paste
'
  CanPaste = (Clipboard.GetFormat(vbCFText) = True)
End Property

Public Property Get CanRedo() As Boolean
'
'is ready to redo
'
  CanRedo = mUndos.CanRedo
   'CanRedo = RedoStack.Count > 0 'Changed
End Property

Public Property Get CanUndo() As Boolean
'
'is ready to undo
'
  CanUndo = mUndos.CanUndo
   'CanUndo = UndoStack.Count > 1 'Changed
End Property

Public Property Get Changed() As Boolean
'
'Get the text changed status
'
  Changed = mChange
End Property

Public Property Let Changed(ByVal Change As Boolean)
'
'Set the text is changed
'
  mChange = Change
End Property

Public Property Get Filename() As String
'
'
'
  Filename = mFilename
End Property

Public Property Let Filename(ByVal value As String)
'
'
'
  value = Replace(value, "\\", "\")
  If Left(value, 1) = "\" Then value = "\" & value
  mFilename = value
End Property

Public Property Let Fullmode(ByVal value As Boolean)
'
'
'
  mFullmode = value
End Property

Public Property Get Fullmode() As Boolean
'
'
'
  Fullmode = mFullmode
End Property

Public Property Let Intelisense(ByVal value As Boolean)
'
'
'
  mIntelisense = value
  If Not value Then IntellBox.Visible = False
End Property

Public Property Get Intelisense() As Boolean
'
'
'
  Intelisense = mIntelisense
End Property

Public Property Let IsHistory(ByVal value As Boolean)
'
'
'
  mIsHistory = value
End Property

Public Property Get IsHistory() As Boolean
'
'
'
  IsHistory = mIsHistory
End Property

Public Property Let IsRemote(ByVal value As Boolean)
'
'
'
  mIsRemote = value
End Property

Public Property Get IsRemote() As Boolean
'
'
'
  IsRemote = mIsRemote
End Property


Public Property Let Key(ByVal value As String)
'
'
'
  mKey = value
End Property

Public Property Get Key() As String
'
'
'
  Key = mKey
End Property

Public Property Let Lineno(ByVal value As Boolean)
'
'
'
  mLineno = value
  On Error GoTo Cerr
  If value Then
    ShowLineno True
  Else
    ShowLineno False
  End If
  RTB.SetFocus
Cerr:
End Property

Public Property Get Lineno() As Boolean
'
'
'
  Lineno = mLineno
End Property

Public Property Let Localhost(ByVal value As String)
'
'
'
  mLocalhost = value
End Property


Public Property Get Localhost() As String
'
'
'
  Localhost = mLocalhost
End Property

Public Property Let Blankpage(ByVal value As Integer)
'
'
'
  mBlankpage = value
End Property

Public Property Get Mode() As Integer
'
'
'
  Mode = mMode
End Property

Public Property Get Blankpage() As Integer
'
'
'
  Blankpage = mBlankpage
End Property

Public Property Let Path(ByVal value As String)
'
'
'
  mAppPath = value
End Property

Public Property Get Path() As String
'
'
'
  Path = mAppPath
End Property

Public Property Let SelColor(ByVal Color As Long)
'
'
'
  RTB.SelColor = Color
End Property

Public Property Get SelColor() As Long
'
'
'
  SelColor = RTB.SelColor
End Property

Public Property Let SelLength(ByVal pLength As Long)
'
'Set the start position
'
  RTB.SelLength = pLength
End Property

Public Property Get SelLength() As Long
'
'Set the start position
'
  SelLength = RTB.SelLength
End Property

Public Property Let SelStart(ByVal pPos As Long)
'
'Set the start position
'
  RTB.SelStart = pPos
End Property

Public Property Get SelStart() As Long
'
'Set the start position
'
  SelStart = RTB.SelStart
End Property

Public Property Let SelText(ByVal pText As String)
'
'
'
  RTB.SelText = ""
  RTB.SelText = pText
End Property

Public Property Get SelText() As String
'
'
'
  SelText = RTB.SelText
End Property

Public Property Let Styles(ByVal value As Long)
'
'
'
  mClasses = value
End Property

Public Property Get Styles() As Long
'
'
'
  Styles = mClasses
End Property

Public Property Let StylesList(ByVal value As Long)
'
'
'
  mCSSList = value
End Property

Public Property Get StylesList() As Long
'
'
'
  StylesList = mCSSList
End Property

Public Property Let SyntaxHighlighting(ByVal value As Boolean)
'
'
'
Dim lStart As Long
  MDoColor = value
  If value Then
    DoColorString 1, Len(RTB.Text)
  Else
    Screen.MousePointer = vbHourglass
    LockWindowUpdate RTB.hWnd
    lStart = RTB.SelStart
    RTB.SelStart = 0
    RTB.SelLength = Len(RTB.Text)
    RTB.SelColor = vbBlack
    RTB.SelLength = 0
    RTB.SelStart = lStart
    LockWindowUpdate 0&
    Screen.MousePointer = vbDefault
  End If
End Property

Public Property Get SyntaxHighlighting() As Boolean
'
'
'
  SyntaxHighlighting = MDoColor
End Property

Public Property Let Text(ByVal value As String)
'
'Set the text
'
  RTB.Text = value
End Property

Public Property Get Text() As String
'
'Get the string of richtextbox
'
  Text = RTB.Text
End Property

Public Property Let Title(ByVal value As String)
'
'Set the text
'
  mTitle = value
End Property

Public Property Get Title() As String
'
'Get the string of richtextbox
'
  Title = mTitle
End Property

Public Property Let VirtualPath(ByVal value As String)
'
'
'
  mVirtualpath = value
End Property

Public Property Get VirtualPath() As String
'
'
'
  VirtualPath = mVirtualpath
End Property

Public Property Let Wordwrap(ByVal value As Boolean)
'
'
'
  mWordwrap = value
  SendMessageLong CLng(RTB.hWnd), EM_SETTARGETDEVICE, &O0, IIf(value, &O1, &O0)
End Property

Public Property Get Wordwrap() As Boolean
'
'
'
  Wordwrap = mWordwrap
End Property


Rem -----------------------------------
Rem           User Functions
Rem -----------------------------------

Private Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
Dim tempParam$
Dim d&
  On Error Resume Next
  If Len(lParam1) > Len(lParam2) Then 'swap
    tempParam$ = lParam1
    lParam1 = lParam2
    lParam2 = tempParam$
  End If
  d& = Len(lParam2) - Len(lParam1)
  Change = Mid(lParam2, startSearch - d&, d&)
End Function

Private Function ChangeMousepointer(ByVal Mode As Integer)
'
'Change the mousepointer
'
  imgSource.MousePointer = vbCustom
  imgSource.MouseIcon = imgHand.Picture
  imgView.MousePointer = vbCustom
  imgView.MouseIcon = imgHand.Picture
  imgSplit.MousePointer = vbCustom
  imgSplit.MouseIcon = imgHand.Picture
  Select Case Mode
  Case 1
    imgSource.MousePointer = vbDefault
  Case 2
    imgView.MousePointer = vbDefault
  Case 3
    imgSplit.MousePointer = vbDefault
  End Select
End Function

Private Function ControlExists(ByVal pKey As String) As Boolean
'
'Chack for the control already added
'
Dim lValue As String
  On Error GoTo Cerr
  lValue = mFormValues(pKey)
  ControlExists = True
  Exit Function
Cerr:
  ControlExists = False
End Function

Public Function Copy() As String
'
'Copy the string
'
  Clipboard.Clear
  Clipboard.SetText RTB.SelText
  Copy = RTB.SelText
End Function

Public Function Cut() As String
'
'Cut the string
'
  Clipboard.Clear
  Clipboard.SetText RTB.SelText
  Cut = RTB.SelText
  mUndos.RTB_KeyDown vbKeyX, 2
  RTB.SelText = vbNullString
  mChange = True
  RaiseEvent Changed(True)
End Function

Public Function Delete() As String
'
'Delete the string
'
  Delete = RTB.SelText
  RTB.SelText = vbNullString
  mChange = True
  RaiseEvent Changed(True)
End Function

Public Function DoColorString(ByVal lStart As Long, ByVal lFinish As Long) As Boolean
'
'Color the string
'
Dim lCursor As Long
Dim sText As String
Dim lLines As Long
Dim lFinish1 As Long
Dim li As Long

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    ' store the cursor position
    lCursor = RTB.SelStart
    
    ' add the text and color it
    LockWindowUpdate RTB.hWnd
    
    ' now add the text to the box
    RTB.SelStart = lStart
    RTB.SelLength = lFinish - lStart + 1
    RTB.SelColor = vbBlack
    basColor.sText = RTB.Text
    lFinish1 = lFinish
    
    'Raise the event
    If InStr(RTB.Text, vbCrLf) > 0 Then
      lLines = UBound(Split(RTB.Text, vbCrLf))
    End If
    RaiseEvent ProgressStatus(1, lLines, "Highlighting...")
    RaiseEvent LockControl
    Do While lStart < lFinish1 + 1
        ' find the end of this line
        lFinish = InStr(lStart + 1, RTB.Text, vbCrLf)
        If lFinish = 0 Then lFinish = lStart + Len(sText)
            
        ' color it
        DoColor RTB, lStart, lFinish
        
        ' move up to get the next line
        lStart = lFinish + 2
        RaiseEvent ProgressStatus(li, lLines, "")
        li = li + 1
    Loop
    RaiseEvent ReleaseControl
    RaiseEvent ProgressStatus(lLines + 1, lLines, "Done")
    ' restore the original values
    RTB.SelStart = lCursor + Len(sText)
    RTB.SelColor = vbBlack
    
    ' null the keypress (to avoid the text pasting twice)
    LockWindowUpdate 0&
    Screen.MousePointer = vbDefault
End Function

Public Function FileExists(ByVal pFilename As String) As Boolean
'
'Check for file existence
'
On Error GoTo S102_Err
  If FileLen(pFilename) > 0 Then
    FileExists = True
  Else
    FileExists = False
  End If
  GoTo S102_Out
S102_Err:
  FileExists = False
S102_Out:
End Function

Public Function Findword(ByVal pText As String, ByVal pPosition As Long, ByVal pFlags As Integer) As Long
'
'Find the word and highliight
'
  mFindWord = pText
  mFindFlags = pFlags
  Findword = RTB.Find(pText, pPosition, , pFlags)
  mPosition = Findword
  If mPosition > -1 Then RTB.SelLength = Len(Trim(pText))
End Function

'Public Sub Focus()
''
''Set the focus
''
'  On Error Resume Next
'  RTB.SetFocus
'End Sub

Private Function GetIndent() As String
Dim lStr As String
Dim i As Long
    lStr = GetLineText2
    For i = 1 To Len(lStr)
        Select Case Mid$(lStr, i, 1)
        Case " ", Chr(vbKeyTab)
            GetIndent = GetIndent & Mid$(lStr, i, 1)
        Case Else
            '// first letter
            Exit Function
        End Select
    Next
End Function

Private Function GetLineText2() As String
Dim intCurrLine As Integer
Dim lStart As Integer
Dim lLen As Integer
Dim strSearch As String
    '// Get current line
    intCurrLine = SendMessage(RTB.hWnd, EM_LINEFROMCHAR, RTB.SelStart, 0&)
    '// Set the start pos at the beginning of the line
    lStart = SendMessage(RTB.hWnd, EM_LINEINDEX, intCurrLine, 0&) + 1
    If Err Then
        '// Line does not exist
        GetLineText2 = ""
        Err.Clear
        Exit Function
    End If
    '// Get the length of the line
    lLen = SendMessage(RTB.hWnd, EM_LINELENGTH, lStart, 0&) + 1
    '// Select the line
    GetLineText2 = Mid$(RTB.Text, lStart, lLen)
End Function

Private Function GetLinkStyles() As String
'
'Get the included css files
'
Dim Lpos1 As Long
Dim Lpos2 As Long
Dim lPath As String
Dim lStyle As String
Dim lList As String
Dim lVirtualpath As String
Dim lAppPath As String
Dim li As Integer
Dim fn As Integer
Dim lPathSplit As Variant
  Lpos1 = 1
  lList = ""
  Do
    Lpos1 = InStr(Lpos1, RTB.Text, "<LINK", vbTextCompare)
    If Lpos1 > 0 Then
      Lpos2 = InStr(Lpos1, RTB.Text, ">")
      If Lpos2 > 0 And (InStr(Lpos1 + 2, RTB.Text, "<") > Lpos2 Or InStr(Lpos1 + 2, RTB.Text, "<") <= 0) Then
        Lpos1 = InStr(Lpos1, RTB.Text, "HREF=", vbTextCompare)
        If Lpos1 > 0 Then
          Lpos2 = InStr(Lpos1, RTB.Text, ".CSS", vbTextCompare)
          If Lpos2 > 0 Then
            lPath = Mid(RTB.Text, Lpos1 + 5, Lpos2 - Lpos1)
            lPath = Replace(lPath, """", "")
            lPath = Replace(LCase(lPath), "file:///", "")
            lPath = Replace(lPath, "|", ":")
            lPath = Replace(lPath, "/", "\")
            mVirtualpath = Replace(mVirtualpath, "/", "\")
            lAppPath = mVirtualpath
            If InStr(lPath, "\") > 0 Then
              If InStr(lPath, "..") > 0 Then
                lPathSplit = Split(lPath, "..")
                For li = 0 To UBound(lPathSplit) - 1
                  lAppPath = Left(lAppPath, InStrRev(lAppPath, "\") - 1)
                Next
                lPath = lAppPath & Mid(lPath, InStrRev(lPath, "..") + 2)
              End If
            Else
              lPath = lAppPath & "\" & lPath
            End If
            If FileExists(lPath) Then
              'If InStr(mCSSFiles, LCase(lPath)) = 0 Then
                fn = FreeFile
                Open lPath For Input As fn
                  lStyle = Input(LOF(fn), fn)
                Close fn
                lList = lList & SplitStylesClasses(lStyle)
                mCSSFiles = mCSSFiles & vbCrLf & LCase(lPath)
              'End If
            End If
          End If
        End If
      End If
    End If
  Loop While InStr(Lpos2 + 1, RTB.Text, "<LINK", vbTextCompare) > 0
  GetLinkStyles = lList
End Function

Private Function GetStylesString() As String
'
'Get the style tag classes
'
Dim Lpos1 As Long
Dim Lpos2 As Long
Dim lStyle As String
Dim lList As String
  Lpos1 = 1
  lList = ""
  Do
    Lpos1 = InStr(Lpos1, RTB.Text, "<STYLE", vbTextCompare)
    If Lpos1 > 0 Then
      Lpos2 = InStr(Lpos1, RTB.Text, "</STYLE>", vbTextCompare)
      If Lpos2 > 0 And Lpos2 > Lpos1 Then
        lStyle = Mid(RTB.Text, Lpos1 + 6, Lpos2 - Lpos1 - 7)
        lList = lList & SplitStylesClasses(lStyle)
      End If
    End If
  Loop While InStr(Lpos2 + 1, RTB.Text, "</STYLE>", vbTextCompare) > 0
  GetStylesString = lList
End Function

Private Function GetWord(Optional pAsp As Boolean = True) As String
Dim Lpos As Long
Dim Lpos1 As Long
Dim Lpos2 As Long
Dim Ltmp As String
Dim lStr As String
  lStr = UCase(Mid(RTB.Text, 1, RTB.SelStart))
  Ltmp = ""
  For Lpos = Len(lStr) To 1 Step -1
    If (Asc(Mid(lStr, Lpos, 1)) > vbKeyA - 1 And Asc(Mid(lStr, Lpos, 1)) < vbKeyZ + 1) Then
      Ltmp = Ltmp & Mid(lStr, Lpos, 1)
    Else
      Exit For
    End If
  Next
  If pAsp Then
    Lpos1 = InStrRev(lStr, "<%")
    Lpos2 = InStrRev(lStr, "%>")
    'If Lpos1 > 0 Then
      'If Lpos1 > Lpos2 Then
        GetWord = StrReverse(Ltmp)
      'End If
    'End If
  Else
    Lpos1 = InStrRev(lStr, "<")
    Lpos2 = InStrRev(lStr, ">")
    If Lpos1 > 0 Then
      If Lpos1 > Lpos2 Then
        GetWord = StrReverse(Ltmp)
      End If
    End If
  End If
End Function

Public Function GotoLine(ByVal pLine As Long) As Boolean
'
'Goto the line here specified
'
Dim lCount As Long
Dim li As Long
Dim lLen As Long
Dim lStart As Long
  lCount = SendMessage(RTB.hWnd, EM_GETLINECOUNT, 0&, 0&)
  If pLine > lCount Or pLine = 0 Then
    MsgBox "Line number out of range.", vbInformation, mTitle
    GotoLine = False
    Exit Function
  End If
  li = 1
  lStart = 1
  Do Until (li > pLine)
    If li = pLine Then
      If InStr(lStart, RTB.Text, vbCrLf) > 0 Then
        lLen = Len(Mid(RTB.Text, lStart, InStr(lStart, RTB.Text, vbCrLf) - lStart))
      Else
        lLen = Len(Mid(RTB.Text, lStart))
      End If
      RTB.SelStart = lStart - 1
      RTB.SelLength = lLen
    End If
    lStart = InStr(lStart, RTB.Text, vbCrLf) + 2
    li = li + 1
  Loop
  GotoLine = True
End Function

Public Sub Indent()
Dim Lpos As Long
Dim Lpos1 As Long
Dim lStart As Long
Dim lText As String
  'If RTB.SelText <> "" Then
    LockWindowUpdate RTB.hWnd
    lStart = RTB.SelStart + 1
    Lpos = InStrRev(RTB.Text, vbCrLf, lStart)
    Lpos1 = InStr(lStart + IIf(Right(RTB.SelText, 2) = CStr(vbCrLf), Len(RTB.SelText) - 2, Len(RTB.SelText)), RTB.Text, vbCrLf) - 1
    If Lpos1 <= 0 Then Lpos1 = Len(RTB.Text)
    If Lpos <= 0 Then
      Lpos = 1
    Else
      Lpos = Lpos + 2
    End If
    lText = Mid(RTB.Text, Lpos, Lpos1 - Lpos + 1)
    RTB.SelStart = Lpos - 1
    RTB.SelLength = Len(lText)
    lText = vbTab & lText
    lText = Replace(lText, vbCrLf, vbCrLf & vbTab)
    If InStr(lText, vbCrLf) = 1 Then lText = Mid(lText, 3)
    RTB.SelText = lText
'    DoColor RTB, lPos - 1, lPos + Len(lText)
    RTB.SelStart = Lpos - 1
    RTB.SelLength = Len(lText)
    LockWindowUpdate 0&
  'End If
End Sub

Private Function IsControlKey(ByVal KeyCode As Long) As Boolean
    ' check if the key is a control key
    Select Case KeyCode
        Case vbKeyLeft, vbKeyRight, vbKeyHome, _
             vbKeyEnd, vbKeyPageUp, vbKeyPageDown, _
             vbKeyShift, vbKeyControl, vbKeyF1, _
             vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, _
             vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, _
             vbKeyF10, vbKeyF11, vbKeyF12
            IsControlKey = True
        Case Else
            IsControlKey = False
    End Select
End Function

Private Sub LoadFile(RTB As Richtextbox, Optional ByVal sFilePath As String)
Dim FF As Long
Dim lStart As Long
Dim lFinish As Long
Dim Text As String
Dim lSelstart As Integer
Dim lExt As String
Dim lLines As Integer
Dim li As Integer
    DoEvents
    
    If sFilePath <> "" Then
      lExt = Mid(sFilePath, InStrRev(sFilePath, ".") + 1)
'      If (UCase(lExt) = "ASP" Or UCase(lExt) = "HTM" Or UCase(lExt) = "HTML") Then
'        MDoColor = True
'      Else
'        MDoColor = False
'      End If
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
    If InStr(RTB.Text, vbCrLf) > 0 Then
      lLines = UBound(Split(RTB.Text, vbCrLf))
      'lLines = SendMessage(RTB.hWnd, EM_GETLINECOUNT, 0&, 0&)
    End If
    RaiseEvent ProgressStatus(1, lLines, "Loading..." & Mid(sFilePath, InStrRev(sFilePath, "\") + 1))
    RaiseEvent LockControl
    li = 1
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
        RaiseEvent ProgressStatus(li, lLines, "")
        li = li + 1
    Loop
    
    RaiseEvent ReleaseControl
    RaiseEvent ProgressStatus(lLines + 1, lLines, "Done")
    ' reset the cursor
    On Error Resume Next
    RTB.SelStart = lSelstart
    RTB.Visible = True
    LockWindowUpdate 0&
End Sub

Public Sub LoadNew(Optional ByVal pAsp As Integer)
Dim lStr As String
Dim li
  lStr = ""
  If pAsp > -1 Then
    If pAsp = 1 Then lStr = lStr & "<%@ Language=VBScript%>" & vbCrLf & vbCrLf
    lStr = lStr & "<HTML>" & vbCrLf
    lStr = lStr & "<head>" & vbCrLf
    lStr = lStr & vbTab & "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">" & vbCrLf
    lStr = lStr & vbTab & "<meta name=""GENERATOR"" content=""Codepiler"">" & vbCrLf
    lStr = lStr & vbTab & "<meta name=""ProgId"" content=""Codepiler.Document"">" & vbCrLf
    lStr = lStr & vbTab & "<title>New Document " & mBlankpage & "</title>" & vbCrLf
    lStr = lStr & "</head>" & vbCrLf
    lStr = lStr & "<body>" & vbCrLf & vbCrLf
    lStr = lStr & "</body>" & vbCrLf
    lStr = lStr & "</html>" & vbCrLf
    MDoColor = True
  End If
  RTB.Visible = False
  RTB.Text = lStr
  LoadFile RTB
  'PutLineNos
  Scrolltxtno
  RaiseEvent DocumentOpened(True)
End Sub

Public Function OpenFile(ByVal pFilename As String)
'
'Load the file
'
Dim lExt As String
  imgSplit.Visible = False
  If mLocalhost <> "" And mVirtualpath <> "" Then
    imgSplit.Visible = True
  End If
  lExt = Mid(pFilename, InStrRev(pFilename, ".") + 1)
  If (UCase(lExt) <> "ASP" And UCase(lExt) <> "HTM" And UCase(lExt) <> "HTML") Then
    imgView.Visible = False
    imgSplit.Visible = False
  End If
  LoadFile RTB, pFilename
  'PutLineNos
  Scrolltxtno
  RaiseEvent DocumentOpened(False)
End Function

Public Sub Outdent()
Dim Lpos As Long
Dim Lpos1 As Long
Dim lStart As Long
Dim lText As String
'  If RTB.SelText <> "" Then
    LockWindowUpdate RTB.hWnd
    lStart = RTB.SelStart + 1
    Lpos = InStrRev(RTB.Text, vbCrLf, lStart)
    Lpos1 = InStr(lStart + IIf(Right(RTB.SelText, 2) = CStr(vbCrLf), Len(RTB.SelText) - 2, Len(RTB.SelText)), RTB.Text, vbCrLf) - 1
    If Lpos1 <= 0 Then Lpos1 = Len(RTB.Text)
    If Lpos <= 0 Then
      Lpos = 1
    Else
      Lpos = Lpos + 2
    End If
    lText = Mid(RTB.Text, Lpos, Lpos1 - Lpos + 1)
    RTB.SelStart = Lpos - 1
    RTB.SelLength = Len(lText)
    lText = Replace(lText, vbCrLf & vbTab, vbCrLf)
    If InStr(lText, vbTab) = 1 Then lText = Mid(lText, 2)
    RTB.SelText = lText
    RTB.SelStart = Lpos - 1
    RTB.SelLength = Len(lText)
    LockWindowUpdate 0&
'  End If
End Sub

Public Function Paste(Optional ByVal pString As String)
'
'Paste the string
'
  If pString <> "" Then
    Clipboard.Clear
    Clipboard.SetText pString
  End If
  DoEvents
  RTB_KeyDown vbKeyV, 2
  mChange = True
  RaiseEvent Changed(True)
End Function

Private Sub pAttachMessages()
    On Error Resume Next
    AttachMessage Me, RTB.hWnd, WM_VSCROLL
    AttachMessage Me, RTB.hWnd, WM_HSCROLL
End Sub

Private Sub pDetachMessages()
    On Error Resume Next
    DetachMessage Me, RTB.hWnd, WM_VSCROLL
    AttachMessage Me, RTB.hWnd, WM_HSCROLL
End Sub

Public Function PrintText()
  On Error GoTo Cerr
   cdMain.Flags = cdlPDReturnDC + cdlPDNoPageNums
   If RTB.SelLength = 0 Then
      cdMain.Flags = cdMain.Flags + cdlPDAllPages
   Else
      cdMain.Flags = cdMain.Flags + cdlPDSelection
   End If
   cdMain.ShowPrinter
   Printer.Print ""
   LockWindowUpdate RTB.hWnd
   RTB.SelStart = 0
   RTB.SelLength = Len(RTB.Text)
   RTB.SelPrint cdMain.hDC
   LockWindowUpdate 0&
Cerr:
End Function

Private Sub ProcessIntell(KeyAscii As Integer)
Dim Pt As POINTAPI
Dim Lword As String
Dim lList As String
Dim Lpos1 As Long
Dim Lpos2 As Long
  On Error Resume Next
  If mIntelisense = False Then Exit Sub
  If IntellBox.Visible Then
    GetCaretPos Pt
    ' Move the popup window to the caret
    IntellBox.Move (Pt.X * Screen.TwipsPerPixelX) + picSource.TextWidth("Z") + RTB.Left, _
                   (Pt.y * Screen.TwipsPerPixelY) + picSource.TextHeight("Z") + RTB.Top
    ' Check if the popup window is within the form
    If IntellBox.Left + IntellBox.Width > ScaleWidth Then IntellBox.Move ScaleWidth - IntellBox.Width - 300
    If IntellBox.Top + IntellBox.Height > ScaleHeight Then IntellBox.Move IntellBox.Left, (Pt.y * Screen.TwipsPerPixelX) - IntellBox.Height
  End If
  Select Case KeyAscii
  Case 60  '<
    Lpos1 = InStrRev(RTB.Text, "<%", RTB.SelStart)
    Lpos2 = InStrRev(RTB.Text, "%>", RTB.SelStart)
    If Lpos1 < Lpos2 Or Lpos1 = 0 Then
      If mChange Then
        IntellBox.PopulateListFromString StringForIntellbox("HTML")
      End If
      ' Get the position of the cursor
      GetCaretPos Pt
      ' Move the popup window to the caret
      IntellBox.Move (Pt.X * Screen.TwipsPerPixelX) + picSource.TextWidth("Z") + RTB.Left, _
                     (Pt.y * Screen.TwipsPerPixelY) + picSource.TextHeight("Z") + RTB.Top
      ' Check if the popup window is within the form
      If IntellBox.Left + IntellBox.Width > ScaleWidth Then IntellBox.Move ScaleWidth - IntellBox.Width - 300
      If IntellBox.Top + IntellBox.Height > ScaleHeight Then IntellBox.Move IntellBox.Left, (Pt.y * Screen.TwipsPerPixelX) - IntellBox.Height
      ' Show the popup window
      IntellBox.Visible = True
      mChange = False
    End If
  Case 62 '>
    IntellBox.Visible = False
    IntellBox.Clear
  Case vbKeyReturn
    If IntellBox.Visible = True Then
      If IntellBox.value <> "" Then
        LockWindowUpdate RTB.hWnd
        RTB.SelStart = RTB.SelStart - IntellBox.InputLen - IntellBox.RemovePrev
        RTB.SelLength = IntellBox.InputLen + IntellBox.RemovePrev
        RTB.SelText = IntellBox.value
        RTB.SelStart = RTB.SelStart - IntellBox.CursorAdjust
        LockWindowUpdate 0&
        bDirty = True
      End If
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
  Case 61 '=
    If IntellBox.Visible Then ProcessIntell vbKeyReturn
    Lword = GetWord(False)
    lList = ""
    If UCase(Lword) = "CLASS" Then
      SplitClasses
      If mCSSList <> "" Then
        If Len(mCSSList) >= 2 Then mCSSList = Left(mCSSList, Len(mCSSList) - 2)
        IntellBox.PopulateListFromString mCSSList, False
        'Move the intellbox to cursor position
        GetCaretPos Pt
        IntellBox.Move (Pt.X * Screen.TwipsPerPixelX) + picSource.TextWidth("Z") + RTB.Left + 20 - 315, _
                       (Pt.y * Screen.TwipsPerPixelY) + picSource.TextHeight("Z") + RTB.Top + 20
        If IntellBox.Left + IntellBox.Width > ScaleWidth Then IntellBox.Move ScaleWidth - IntellBox.Width - 300
        If IntellBox.Top + IntellBox.Height > ScaleHeight Then IntellBox.Move IntellBox.Left, (Pt.y * Screen.TwipsPerPixelX) - IntellBox.Height
        If UBound(Split(mCSSList, vbCrLf)) < 10 Then
          IntellBox.Height = UBound(Split(mCSSList, vbCrLf)) * Screen.TwipsPerPixelY * 12.75 '((Screen.TwipsPerPixelY * 16) * 3) - ((Screen.TwipsPerPixelX) * 3)
        Else
          IntellBox.Height = 9 * Screen.TwipsPerPixelY * 12.75
        End If
        IntellBox.Visible = True
        
      End If
    End If
  Case 46 'Decimal point
    If IntellBox.Visible Then ProcessIntell vbKeyReturn
    Lword = ""
    lList = ""
    Lword = GetWord
    lList = StringForIntellbox(Lword)
    If lList <> "" Then
      IntellBox.PopulateListFromString lList, False
      'Move the intellbox to cursor position
      GetCaretPos Pt
      IntellBox.Move (Pt.X * Screen.TwipsPerPixelX) + picSource.TextWidth("Z") + RTB.Left + 20 - 315, _
                     (Pt.y * Screen.TwipsPerPixelY) + picSource.TextHeight("Z") + RTB.Top + 20
      If IntellBox.Left + IntellBox.Width > ScaleWidth Then IntellBox.Move ScaleWidth - IntellBox.Width - 300
      If IntellBox.Top + IntellBox.Height > ScaleHeight Then IntellBox.Move IntellBox.Left, (Pt.y * Screen.TwipsPerPixelX) - IntellBox.Height
      IntellBox.Visible = True
      
    End If
  Case 32 'Space
    'If IntellBox.Visible = True Then IntellBox.Visible = False
    If IntellBox.Visible = True Then ProcessIntell vbKeyReturn
  Case 40 '(
    If IntellBox.Visible = True Then ProcessIntell vbKeyReturn
  Case vbKeyTab
    If IntellBox.Visible = True Then ProcessIntell vbKeyReturn
  
  Case Else:
    If IntellBox.Visible = True Then IntellBox.AddChar Chr(KeyAscii)
  End Select
End Sub

Private Sub RaisePosition()
Dim lCurrentLine As Long
Dim lCurrentCur As Long
  lCurrentLine = SendMessage(RTB.hWnd, EM_LINEFROMCHAR, -1, 0&) + 1
  lCurrentCur = SendMessage(RTB.hWnd, EM_LINEINDEX, lCurrentLine - 1, 0&)
  RaiseEvent Position(RTB.SelStart - lCurrentCur + 1, lCurrentLine)
End Sub

Public Function Redo() As Boolean
'
'Redo the undo text
'
  mUndos.Redo
  Scrolltxtno
  RaiseEvent StateChanged
  RTB.SetFocus
End Function

Public Sub ReplaceWord(ByVal pText As String)
'
'Replace the selected word
'
  RTB.SelText = pText
End Sub

Public Function SaveAsFile() As String
'
'Save as file
'
Dim sFile As String
Dim lExt As String
  On Error GoTo Cerr
  If mFilename <> "" Then lExt = Mid(mFilename, InStrRev(mFilename, ".") + 1)
  With cdMain
    If mFilename <> "" Then .Filename = mFilename
    If (UCase(lExt) <> "ASP" And UCase(lExt) <> "HTM" And UCase(lExt) <> "HTML") Then
      .DialogTitle = "Save file"
      .Filter = "All Files(*.*)|*.*"
    Else
      .DialogTitle = "Save Web Page"
      .Filter = "All Web Pages(*.html *.htm *.asp *.shtml)|*.html;*.htm;*.asp;*.shtml|Asp(*.asp)|*.asp|Htm(*.htm)|*.htm|All Files(*.*)|*.*"
    End If
    If mVirtualpath <> "" Then .InitDir = mVirtualpath
    .CancelError = True
    .ShowSave
    If Len(.Filename) = 0 Then
      Exit Function
    End If
    sFile = .Filename
  End With
  RTB.SaveFile sFile, 1
  mFilename = sFile
  mChange = False
  SaveAsFile = sFile
  RaiseEvent FileSaved(sFile, True)
  RaiseEvent Changed(False)
Cerr:
End Function

Public Sub SaveFile(ByVal pFilename As String, Optional ByVal Force As Boolean)
'
'Save the file name
'
  If FileExists(pFilename) Or Force Then
    RTB.SaveFile pFilename, 1
    mChange = False
    RaiseEvent FileSaved(pFilename, False)
    RaiseEvent Changed(False)
  Else
    MsgBox "Working file doesnot exist!", vbCritical, mTitle
  End If
End Sub

Public Function Scrolltxtno()
  tmr.Enabled = True
End Function

Public Sub SelectAll()
'
'Select all
'
  RTB.SetFocus
  RTB.SelStart = 0
  RTB.SelLength = Len(RTB.Text)
End Sub

Private Sub ShowLineno(Optional ByVal pShow As Boolean = True)
'
'Show the line no
'
  picLinesBack.Visible = pShow
  If pShow Then Scrolltxtno
End Sub

Public Function ShowPreview()
'
'Show the preview
'
Dim FileNum As Integer
Dim lFilename As String
Dim lLocalhost As String

  On Error Resume Next
  Screen.MousePointer = vbHourglass
  If mLocalhost <> "" And mVirtualpath <> "" And mFilename <> "" Then
    If mChange Then 'if any change, then save the changs before view
      If IsRemote Then 'Show the confirm for remote changes
        If mDontShowConfirm = False Then
          Set mFrmConfirm = New frmConfirm
          mFrmConfirm.Show vbModal
        End If
        If mProceedRemoteChanges Then
          SaveFile mFilename
        End If
      Else 'if local site
        SaveFile mFilename
      End If
    End If
    lFilename = mFilename 'Ucase Changed
    lLocalhost = mLocalhost 'Ucase Changed
    lFilename = Replace(lFilename, mVirtualpath, "", , , vbTextCompare) 'Ucase Changed
    lFilename = Replace(lFilename, "\", "/")
    If Left(lFilename, 1) = "/" Then lFilename = Mid(lFilename, 2)
    If Right(lLocalhost, 1) <> "/" Then lLocalhost = lLocalhost & "/"
    wbPreview.Navigate2 lLocalhost & lFilename & "?" & Format(Now, "hhmmss")
  Else 'new file or work space files
    If Right(mAppPath, 1) <> "\" Then mAppPath = mAppPath & "\"
    If mChange Then
      If mTmpFile <> "" Then Kill mAppPath & mTmpFile
      FileNum = FreeFile
      mTmpFile = "rnd" & Rnd(1000) & "fnm.htm"
      Open mAppPath & mTmpFile For Output As #FileNum
      Print #FileNum, RTB.Text
      Close #FileNum
    End If
    wbPreview.Navigate2 mAppPath & mTmpFile & "?" & Format(Now, "hhmmss")
  End If
  mPreviewed = True
End Function

Public Function SplitClasses()
'
'Get the classes from stylesheet files
'
  mClasses = 0
  If InStr(mCSSList, "vbcrlf") > 0 Then
    mClasses = UBound(Split(mCSSList, vbCrLf)) + 2
  End If
  mCSSList = GetLinkStyles
  mCSSList = mCSSList & GetStylesString()
End Function

Private Function SplitStylesClasses(ByVal pStyle As String) As String
'
'Split all classes
'
Dim lLines As Variant
Dim lClasses As String
Dim lStr As String
Dim Lpos As Long
Dim li As Long
  Lpos = 1
  pStyle = Replace(pStyle, vbCrLf, "")
  lLines = Split(pStyle, "}")
  For li = LBound(lLines) To UBound(lLines)
    Lpos = InStr(1, lLines(li), "{")
    If Lpos > 0 Then
      lStr = Mid(lLines(li), 1, Lpos - 1)
      lStr = Trim(lStr)
      lStr = Replace(lStr, vbTab, "")
      If Left(lStr, 1) = "." Then
        lClasses = lClasses & ",2," & Right(lStr, Len(lStr) - 1) & "," & Right(lStr, Len(lStr) - 1) & "," & Len(lStr) - 1 & ",0" & vbCrLf
      End If
    End If
  Next
  SplitStylesClasses = lClasses
End Function

Private Function StringForIntellbox(Optional ByVal pObjName As String) As String
Dim lStr As String
  Select Case UCase(pObjName)
  Case "RESPONSE"
    lStr = ""
    lStr = lStr & ",2,AddHeader,AddHeader,9,0" & vbCrLf
    lStr = lStr & ",2,AppendToLog,AppendToLog,11,0" & vbCrLf
    lStr = lStr & ",2,BinaryWrite,BinaryWrite,11,0" & vbCrLf
    lStr = lStr & ",3,Buffer,Buffer,6,0" & vbCrLf
    lStr = lStr & ",3,CacheControl,CacheControl,12,0" & vbCrLf
    lStr = lStr & ",3,CharSet,CharSet,7,0" & vbCrLf
    lStr = lStr & ",2,Clear,Clear,5,0" & vbCrLf
    lStr = lStr & ",3,ContentType,ContentType,11,0" & vbCrLf
    lStr = lStr & ",3,Cookies,Cookies,7,0" & vbCrLf
    lStr = lStr & ",2,End,End,3,0" & vbCrLf
    lStr = lStr & ",3,Expires,Expires,7,0" & vbCrLf
    lStr = lStr & ",3,ExpiresAbsolute,ExpiresAbsolute,15,0" & vbCrLf
    lStr = lStr & ",2,Flush,Flush,5,0" & vbCrLf
    lStr = lStr & ",2,IsClientConnected,IsClientConnected,17,0" & vbCrLf
    lStr = lStr & ",2,Pics,Pics,4,0" & vbCrLf
    lStr = lStr & ",2,Redirect,Redirect,8,0" & vbCrLf
    lStr = lStr & ",3,Status,Status,6,0" & vbCrLf
    lStr = lStr & ",2,Write,Write,5,0"
    IntellBox.Height = 10 * Screen.TwipsPerPixelY * 13
  Case "REQUEST"
    lStr = ""
    lStr = lStr & ",2,BinaryRead,BinaryRead,10,0" & vbCrLf
    lStr = lStr & ",3,ClientCertificate,ClientCertificate,17,0" & vbCrLf
    lStr = lStr & ",3,Cookies,Cookies,7,0" & vbCrLf
    lStr = lStr & ",3,Form,Form,4,0" & vbCrLf
    lStr = lStr & ",3,Item,Item,4,0" & vbCrLf
    lStr = lStr & ",3,QueryString,QueryString,11,0" & vbCrLf
    lStr = lStr & ",3,ServerVariables,ServerVariables,15,0" & vbCrLf
    lStr = lStr & ",3,TotalBytes,TotalBytes,9,0"
    IntellBox.Height = 9 * Screen.TwipsPerPixelY * 12.75
  Case "SESSION"
    lStr = ""
    lStr = lStr & ",2,Abandon,Abandon,7,0" & vbCrLf
    lStr = lStr & ",3,CodePage,CodePage,8,0" & vbCrLf
    lStr = lStr & ",3,Contents,Contents,8,0" & vbCrLf
    lStr = lStr & ",3,LCID,LCID,4,0" & vbCrLf
    lStr = lStr & ",3,SessionID,SessionID,9,0" & vbCrLf
    lStr = lStr & ",3,StaticObjects,StaticObjects,13,0" & vbCrLf
    lStr = lStr & ",3,Timeout,Timeout,7,0" & vbCrLf
    lStr = lStr & ",3,Value,Value,7,0"
    IntellBox.Height = 9 * Screen.TwipsPerPixelY * 12.75
  Case "ASPERROR"
    lStr = ""
    lStr = lStr & ",2,ASPCode(),ASPCode(),9,0" & vbCrLf
    lStr = lStr & ",3,Number(),Number(),8,0" & vbCrLf
    lStr = lStr & ",3,Source(),Source(),8,0" & vbCrLf
    lStr = lStr & ",3,Category(),Category(),9,0" & vbCrLf
    lStr = lStr & ",3,File(),File(),6,0" & vbCrLf
    lStr = lStr & ",3,Line(),Line(),6,0" & vbCrLf
    lStr = lStr & ",3,Column(),Column(),8,0" & vbCrLf
    lStr = lStr & ",3,Description(),Description(),13,0" & vbCrLf
    lStr = lStr & ",3,ASPDescription(),ASPDescription(),16,0"
    IntellBox.Height = 9 * Screen.TwipsPerPixelY * 12.75
  Case "APPLICATION"
    lStr = ""
    lStr = lStr & ",3,Contents,Contents,8,0" & vbCrLf
    lStr = lStr & ",2,Lock,Lock,4,0" & vbCrLf
    lStr = lStr & ",3,StaticObjects,StaticObjects,13,0" & vbCrLf
    lStr = lStr & ",2,Unlock,Unlock,6,0" & vbCrLf
    lStr = lStr & ",3,Value,Value,5,0"
    IntellBox.Height = 6 * Screen.TwipsPerPixelY * 12.25
  Case "SERVER"
    lStr = ""
    lStr = lStr & ",2,CreateObject,CreateObject,12,0" & vbCrLf
    lStr = lStr & ",2,Execute,Execute,7,0" & vbCrLf
    lStr = lStr & ",2,GetLastError(),GetLastError(),14,0" & vbCrLf
    lStr = lStr & ",2,HTMLEncode,HTMLEncode,10,0" & vbCrLf
    lStr = lStr & ",2,MapPath,MapPath,7,0" & vbCrLf
    lStr = lStr & ",3,ScriptTimeout,ScriptTimeout,13,0" & vbCrLf
    lStr = lStr & ",2,Transfer,Transfer,8,0" & vbCrLf
    lStr = lStr & ",2,URLEncode,URLEncode,9,0" & vbCrLf
    lStr = lStr & ",2,URLPathEncode,URLPathEncode,13,0"
    IntellBox.Height = 7 * Screen.TwipsPerPixelY * 12.5
  Case "OBJECTCONTEXT"
    lStr = ""
    lStr = lStr & ",2,SetAbort(),SetAbort(),10,0" & vbCrLf
    lStr = lStr & ",2,SetComplete(),SetComplete(),13,0"
'    lStr = lStr & ",3,Application,Application,11,0" & vbCrLf
'    lStr = lStr & ",3,Request,Request,7,0" & vbCrLf
'    lStr = lStr & ",3,Response,Response,8,0" & vbCrLf
'    lStr = lStr & ",3,Server,Server,6,0" & vbCrLf
'    lStr = lStr & ",3,Session,Session,7,0"
    IntellBox.Height = 3 * Screen.TwipsPerPixelY * 10.75
  Case "CONTENTS", "STATICOBJECTS", "QUERYSTRING", "FORM", "COOKIES"
    lStr = ""
    lStr = lStr & ",3,Count,Count,5,0" & vbCrLf
    lStr = lStr & ",3,Item,Item,4,0" & vbCrLf
    lStr = lStr & ",3,Key,Key,3,0" & vbCrLf
    lStr = lStr & ",3,Remove,Remove,6,0" & vbCrLf
    lStr = lStr & ",3,RemoveAll,RemoveAll,9,0"
    IntellBox.Height = 6 * Screen.TwipsPerPixelY * 12.25
  Case "HTML"
    lStr = ""
    lStr = lStr & ",1,a,a ></a>,2,0" & vbCrLf
    lStr = lStr & ",1,abbr,abbr></abbr>,5,0" & vbCrLf
    lStr = lStr & ",1,acronym,acronym></acronym>,8,0" & vbCrLf
    lStr = lStr & ",1,address,address></address>,8,0" & vbCrLf
    lStr = lStr & ",1,applet,applet></applet>,7,0" & vbCrLf
    lStr = lStr & ",1,area,area></area>,5,0" & vbCrLf
    lStr = lStr & ",1,b,b></b>,2,0" & vbCrLf
    lStr = lStr & ",1,base,base></base>,5,0" & vbCrLf
    lStr = lStr & ",1,basefont,basefont></basefont>,9,0" & vbCrLf
    lStr = lStr & ",1,bdo,bdo></bdo>,4,0" & vbCrLf
    lStr = lStr & ",1,bgsound,bgsound></bgsound>,8,0" & vbCrLf
    lStr = lStr & ",1,big,big></big>,4,0" & vbCrLf
    lStr = lStr & ",1,bling,blink></blink>,6,0" & vbCrLf
    lStr = lStr & ",1,blockquote,blockquote></blockquote>,11,0" & vbCrLf
    lStr = lStr & ",1,body,body></body>,5,0" & vbCrLf
    lStr = lStr & ",1,br,br>,3,0" & vbCrLf
    lStr = lStr & ",1,button,button></button>,7,0" & vbCrLf
    lStr = lStr & ",1,caption,caption></caption>,8,0" & vbCrLf
    lStr = lStr & ",1,center,center></center>,7,0" & vbCrLf
    lStr = lStr & ",1,cite,cite></cite>,5,0" & vbCrLf
    lStr = lStr & ",1,code,code></code>,5,0" & vbCrLf
    lStr = lStr & ",1,col,col>,4,0" & vbCrLf
    lStr = lStr & ",1,colgroup,colgroup></colgroup>,9,0" & vbCrLf
    lStr = lStr & ",1,comment,comment></comment>,8,0" & vbCrLf
    lStr = lStr & ",1,dd,dd></dd>,3,0" & vbCrLf
    lStr = lStr & ",1,del,del></del>,4,0" & vbCrLf
    lStr = lStr & ",1,dfn,dfn></dfn>,4,0" & vbCrLf
    lStr = lStr & ",1,dir,dir></dir>,4,0" & vbCrLf
    lStr = lStr & ",1,div,div></div>,4,0" & vbCrLf
    lStr = lStr & ",1,dl,dl></dl>,3,0" & vbCrLf
    lStr = lStr & ",1,dt,dt></dt>,3,0" & vbCrLf
    lStr = lStr & ",1,em,em></em>,3,0" & vbCrLf
    lStr = lStr & ",1,embed,embed></embed>,6,0" & vbCrLf
    lStr = lStr & ",1,fieldset,fieldset></fieldset>,9,0" & vbCrLf
    lStr = lStr & ",1,font,font></font>,5,0" & vbCrLf
    lStr = lStr & ",1,form,form></form>,5,0" & vbCrLf
    lStr = lStr & ",1,frame,frame>,6,0" & vbCrLf
    lStr = lStr & ",1,frameset,frameset></frameset>,9,0" & vbCrLf
    lStr = lStr & ",1,h1,h1></h1>,3,0" & vbCrLf
    lStr = lStr & ",1,h2,h2></h2>,3,0" & vbCrLf
    lStr = lStr & ",1,h3,h3></h3>,3,0" & vbCrLf
    lStr = lStr & ",1,h4,h4></h4>,3,0" & vbCrLf
    lStr = lStr & ",1,h5,h5></h5>,3,0" & vbCrLf
    lStr = lStr & ",1,h6,h6></h6>,3,0" & vbCrLf
    lStr = lStr & ",1,head,head></head>,5,0" & vbCrLf
    lStr = lStr & ",1,hr,hr>,3,0" & vbCrLf
    lStr = lStr & ",1,html,html></html>,5,0" & vbCrLf
    lStr = lStr & ",1,i,i></i>,2,0" & vbCrLf
    lStr = lStr & ",1,iframe,iframe></iframe>,7,0" & vbCrLf
    lStr = lStr & ",1,ilayer,ilayer></ilayer>,7,0" & vbCrLf
    lStr = lStr & ",1,img,img >,4,0" & vbCrLf
    lStr = lStr & ",1,input,input >,6,0" & vbCrLf
    lStr = lStr & ",1,ins,ins></ins>,4,0" & vbCrLf
    lStr = lStr & ",1,isindex,isindex></isindex>,8,0" & vbCrLf
    lStr = lStr & ",1,kbd,kbd></kbd>,4,0" & vbCrLf
    lStr = lStr & ",1,keygen,keygen>,7,0" & vbCrLf
    lStr = lStr & ",1,label,label></label>,6,0" & vbCrLf
    lStr = lStr & ",1,layer,layer></layer>,6,0" & vbCrLf
    lStr = lStr & ",1,legend,legent></legent>,7,0" & vbCrLf
    lStr = lStr & ",1,li,li></li>,3,0" & vbCrLf
    lStr = lStr & ",1,link,link></link>,5,0" & vbCrLf
    lStr = lStr & ",1,listing,listing></listing>,8,0" & vbCrLf
    lStr = lStr & ",1,map,map></map>,4,0" & vbCrLf
    lStr = lStr & ",1,marquee,marquee></marquee>,8,0" & vbCrLf
    lStr = lStr & ",1,menu,menu></menu>,5,0" & vbCrLf
    lStr = lStr & ",1,meta,meta>,5,0" & vbCrLf
    lStr = lStr & ",1,multicol,multicol></multicol>,9,0" & vbCrLf
    lStr = lStr & ",1,mobr,nobr></nobr>,5,0" & vbCrLf
    lStr = lStr & ",1,noembed,noembed></noembed>,8,0" & vbCrLf
    lStr = lStr & ",1,noframes,noframes></noframes>,9,0" & vbCrLf
    lStr = lStr & ",1,nolayer,nolayer></nolayer>,8,0" & vbCrLf
    lStr = lStr & ",1,noscript,noscript></noscript>,9,0" & vbCrLf
    lStr = lStr & ",1,object,object></object>,7,0" & vbCrLf
    lStr = lStr & ",1,ol,ol></ol>,3,0" & vbCrLf
    lStr = lStr & ",1,optgroup,optgroup></optgroup>,9,0" & vbCrLf
    lStr = lStr & ",1,option,option></option>,7,0" & vbCrLf
    lStr = lStr & ",1,p,p></p>,2,0" & vbCrLf
    lStr = lStr & ",1,param,param>,6,0" & vbCrLf
    lStr = lStr & ",1,plaintext,palintext></plaintext>,10,0" & vbCrLf
    lStr = lStr & ",1,pre,pre></pre>,4,0" & vbCrLf
    lStr = lStr & ",1,q,q></q>,2,0" & vbCrLf
    lStr = lStr & ",1,s,s></s>,2,0" & vbCrLf
    lStr = lStr & ",1,samp,samp></samp>,5,0" & vbCrLf
    lStr = lStr & ",1,script,script></script>,7,0" & vbCrLf
    lStr = lStr & ",1,select,select></select>,7,0" & vbCrLf
    lStr = lStr & ",1,server,server></server>,7,0" & vbCrLf
    lStr = lStr & ",1,small,small></small>,6,0" & vbCrLf
    lStr = lStr & ",1,spacer,spacer>,7,0" & vbCrLf
    lStr = lStr & ",1,span,span></span>,5,0" & vbCrLf
    lStr = lStr & ",1,strike,strike></strike>,7,0" & vbCrLf
    lStr = lStr & ",1,strong,strong></strong>,7,0" & vbCrLf
    lStr = lStr & ",1,style,style></style>,6,0" & vbCrLf
    lStr = lStr & ",1,sub,sub></sub>,4,0" & vbCrLf
    lStr = lStr & ",1,sup,sup></sup>,4,0" & vbCrLf
    lStr = lStr & ",1,table,table></table>,6,0" & vbCrLf
    lStr = lStr & ",1,tbody,tbody></tbody>,6,0" & vbCrLf
    lStr = lStr & ",1,td,td></td>,3,0" & vbCrLf
    lStr = lStr & ",1,textarea,textarea></textarea>,9,0" & vbCrLf
    lStr = lStr & ",1,tfoot,tfoot></tfoot>,6,0" & vbCrLf
    lStr = lStr & ",1,th,th></th>,3,0" & vbCrLf
    lStr = lStr & ",1,thead,thead></thead>,6,0" & vbCrLf
    lStr = lStr & ",1,title,title></title>,6,0" & vbCrLf
    lStr = lStr & ",1,tr,tr></tr>,3,0" & vbCrLf
    lStr = lStr & ",1,tt,tt></tt>,3,0" & vbCrLf
    lStr = lStr & ",1,u,u></u>,2,0" & vbCrLf
    lStr = lStr & ",1,ul,ul></ul>,3,0" & vbCrLf
    lStr = lStr & ",1,var,var></var>,4,0" & vbCrLf
    lStr = lStr & ",1,wbr,wbr>,4,0" & vbCrLf
    lStr = lStr & ",1,xmp,xmp></xmp>,4,0"

    IntellBox.Height = 9 * Screen.TwipsPerPixelY * 12.75
  End Select
  StringForIntellbox = lStr
End Function

Public Function Undo() As Boolean
'
'Undo the text
'
  mUndos.Undo
  Scrolltxtno
  RaiseEvent StateChanged
  RTB.SetFocus
End Function

Public Sub View(ByVal Mode As Integer)
'
'View the editor in Code(1),View(2), and Code/View(3)
'
  If Mode = 1 Then
    imgSource_Click
  ElseIf Mode = 2 Then
    imgView_Click
  ElseIf Mode = 3 Then
    If imgSplit.Visible Then imgSplit_Click
  End If
End Sub


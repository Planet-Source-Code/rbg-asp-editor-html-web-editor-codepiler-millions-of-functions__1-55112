VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Image"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
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
   Icon            =   "frmImage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   5115
      Picture         =   "frmImage.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   825
      Width           =   330
   End
   Begin VB.ComboBox cboTarget 
      Height          =   315
      ItemData        =   "frmImage.frx":07BB
      Left            =   1620
      List            =   "frmImage.frx":07CE
      TabIndex        =   8
      Text            =   "_Default"
      Top             =   4740
      Width           =   1335
   End
   Begin VB.ComboBox cboLink 
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Text            =   "http://"
      Top             =   3690
      Width           =   1095
   End
   Begin VB.CheckBox chkMakeLink 
      Caption         =   "Make into link"
      Height          =   195
      Left            =   705
      TabIndex        =   5
      Top             =   3315
      Width           =   1395
   End
   Begin VB.TextBox txtActualLink 
      Height          =   315
      Left            =   1605
      TabIndex        =   7
      Top             =   4200
      Width           =   3225
   End
   Begin VB.TextBox txtBorder 
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Text            =   "0"
      Top             =   2415
      Width           =   480
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   825
      Width           =   3495
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Top             =   1342
      Width           =   465
   End
   Begin VB.TextBox txtHeight 
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   1860
      Width           =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5865
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   3555
      TabIndex        =   10
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4575
      TabIndex        =   11
      Top             =   5295
      Width           =   870
   End
   Begin VB.CheckBox chkAspCode 
      Caption         =   "Make as ASP Code"
      Height          =   195
      Left            =   3765
      TabIndex        =   9
      Top             =   2535
      Width           =   1680
   End
   Begin MSComCtl2.UpDown UpDown3 
      Height          =   315
      Left            =   2101
      TabIndex        =   12
      Top             =   1860
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      OrigLeft        =   3180
      OrigTop         =   720
      OrigRight       =   3420
      OrigBottom      =   960
      Max             =   99999
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   2085
      TabIndex        =   13
      Top             =   1342
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      OrigLeft        =   1530
      OrigTop         =   735
      OrigRight       =   1770
      OrigBottom      =   975
      Max             =   99999
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2085
      TabIndex        =   16
      Top             =   2415
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      OrigLeft        =   1486
      OrigTop         =   1065
      OrigRight       =   1726
      OrigBottom      =   1350
      Max             =   99999
      Enabled         =   -1  'True
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   690
      X2              =   8405
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   735
      X2              =   8405
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   375
      Picture         =   "frmImage.frx":07FA
      Top             =   105
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   705
      TabIndex        =   22
      Top             =   885
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image"
      Height          =   195
      Left            =   825
      TabIndex        =   21
      Top             =   210
      Width           =   450
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7910
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   195
      X2              =   7910
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Link:"
      Height          =   195
      Left            =   705
      TabIndex        =   20
      Top             =   4260
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Target:"
      Height          =   195
      Left            =   705
      TabIndex        =   19
      Top             =   4800
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link Type:"
      Height          =   195
      Left            =   705
      TabIndex        =   18
      Top             =   3750
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Border:"
      Height          =   195
      Left            =   705
      TabIndex        =   17
      Top             =   2475
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      Height          =   195
      Left            =   705
      TabIndex        =   15
      Top             =   1402
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      Height          =   195
      Left            =   705
      TabIndex        =   14
      Top             =   1920
      Width           =   525
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private m_Picture As Picture
Private m_bm As BITMAP

Public Function ImageReadOK(FileName As String) As Boolean
  On Error Resume Next
  Set m_Picture = LoadPicture(FileName)
  If Err Then
    ImageReadOK = False
    Exit Function
  End If
  ImageReadOK = (GetObjectAPI(m_Picture.Handle, Len(m_bm), m_bm) = Len(m_bm))
End Function

Public Property Get WidthPixels() As Long
  WidthPixels = m_bm.bmWidth
End Property

Public Property Get HeightPixels() As Long
  HeightPixels = m_bm.bmHeight
End Property

Public Property Get WidthHiMetric() As Long
  WidthHiMetric = m_Picture.Width
End Property

Public Property Get HeightHiMetric() As Long
  HeightHiMetric = m_Picture.Height
End Property

Private Sub chkMakeLink_Click()
  Enable chkMakeLink.Value
End Sub

Private Sub Command1_Click()
Dim Image As String
Dim lPath As String
Dim Link As String
  If frmEditor.RTB(frmEditor.GetActiveRTB).Tag <> "" Then
    lPath = GetVirtualPath(Split(frmEditor.RTB(frmEditor.GetActiveRTB).Tag, "^")(0), txtName.Text, frmEditor.tabMain.SelectedTab.Key)
  End If
  If lPath = "" Then lPath = "file:///" & Replace(txtName.Text, "\", "/")
  If chkAspCode.Value = False Then
    Image = "<IMG SRC=""" & lPath & """ border=""" & txtBorder.Text & """ width=""" & txtWidth.Text & """ height=""" & txtHeight.Text & """>"
    If chkMakeLink.Value = False Then
      Link = Image
    Else
      Link = "<a href=""" & cboLink.Text & txtActualLink.Text & """ target=""" & cboTarget.Text & """ >" & Image & "</a>"
    End If
  Else
    If chkMakeLink.Value = False Then
      Image = "Responce.write" & Chr(34) & "<IMG SRC=""" & Chr(34) & lPath & Chr(34) & """ border=""" & Chr(34) & txtBorder.Text & Chr(34) & """ width=""" & Chr(34) & txtWidth.Text & Chr(34) & """ height=""" & Chr(34) & txtHeight.Text & Chr(34) & """>" & Chr(34)
      Link = Image
    Else
      Image = "<IMG SRC=""" & Chr(34) & lPath & Chr(34) & """ border=""" & Chr(34) & txtBorder.Text & Chr(34) & """ width=""" & Chr(34) & txtWidth.Text & Chr(34) & """ height=""" & Chr(34) & txtHeight.Text & Chr(34) & """>"
      Link = "Response.Write" & Chr(34) & "<a href=""" & Chr(34) & cboLink.Text & txtActualLink.Text & Chr(34) & """ target=""" & Chr(34) & cboTarget.Text & Chr(34) & """ >" & Image & "</a>" & Chr(34)
    End If
  End If
  'Clipboard.SetText Link
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste Link
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub cmdBrowse_Click()
Dim sFile As String
  With CommonDialog1
    .DialogTitle = "Open"
    .CancelError = False
    .Filter = "All Images(*.jpg *.gif)|*.jpg;*.gif|Jpeg(*.jpg)|*.jpg|Gif(*.gif)|*.gif|All Files(*.*)|*.*"
    .DefaultExt = "JPEG"
    .ShowOpen
    If Len(.FileName) = 0 Then
      Exit Sub
    End If
    sFile = .FileName
    txtName.Text = sFile
    ImageReadOK (CommonDialog1.FileName)
    txtWidth.Text = WidthPixels
    txtHeight.Text = HeightPixels
  End With
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub

Private Sub Enable(ByVal pEnable As Boolean)
  cboLink.Enabled = pEnable
  txtActualLink.Enabled = pEnable
  cboTarget.Enabled = pEnable
End Sub

Private Sub Form_Load()
  chkMakeLink_Click
End Sub

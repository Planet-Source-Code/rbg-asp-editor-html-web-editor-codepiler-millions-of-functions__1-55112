VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStylesheet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stylesheet"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
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
   Icon            =   "frmStylesheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4575
      Picture         =   "frmStylesheet.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   330
   End
   Begin VB.CheckBox chkImport 
      Caption         =   "Import"
      Height          =   285
      Left            =   1365
      TabIndex        =   2
      Top             =   1215
      Width           =   780
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1350
      TabIndex        =   0
      Top             =   735
      Width           =   3195
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4995
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   2805
      TabIndex        =   3
      Top             =   1770
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3945
      TabIndex        =   4
      Top             =   1770
      Width           =   960
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Make image responce write"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   3765
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image Image1 
      Height          =   165
      Left            =   225
      Picture         =   "frmStylesheet.frx":07BB
      Top             =   180
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File/URL:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   510
      TabIndex        =   7
      Top             =   795
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stylesheet Include"
      Height          =   195
      Left            =   525
      TabIndex        =   6
      Top             =   150
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   300
      X2              =   7970
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   255
      X2              =   7970
      Y1              =   525
      Y2              =   525
   End
End
Attribute VB_Name = "frmStylesheet"
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

Private Sub Command1_Click()
Dim Image As String
Dim lPath As String
Dim lStyleOpen As String
Dim lStyleClose As String
Dim Link As String
  If frmEditor.RTB(frmEditor.GetActiveRTB).Tag <> "" Then
    lPath = GetVirtualPath(Split(frmEditor.RTB(frmEditor.GetActiveRTB).Tag, "^")(0), Text3.Text, frmEditor.tabMain.SelectedTab.Key)
  End If
  lStyleOpen = "<style type=""text/css"">" & vbCrLf
  lStyleClose = vbCrLf & "</style>"
  If lPath = "" Then lPath = "file:///" & Replace(Text3.Text, "\", "/")
  If chkImport.Value = 1 Then
    Link = vbCrLf & lStyleOpen & "<!-- @import url(""" & lPath & """) -->" & lStyleClose
  Else
    Link = vbCrLf & "<Link href=""" & lPath & """ rel=""stylesheet"" type=""text/css"">"
  End If
  If InStr(frmEditor.RTB(frmEditor.GetActiveRTB).Text, "<head>") > 0 Then
    frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(frmEditor.RTB(frmEditor.GetActiveRTB).Text, "<head>") + 6
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
    .Filter = "Stylesheet(*.css)|*.css"
    .DefaultExt = "CSS"
    .ShowOpen
    If Len(.FileName) = 0 Then
      Exit Sub
    End If
    sFile = .FileName
    Text3.Text = sFile
    ImageReadOK (CommonDialog1.FileName)
  End With
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub

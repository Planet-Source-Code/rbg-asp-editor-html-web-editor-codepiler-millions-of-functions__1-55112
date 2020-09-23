VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4650
      Picture         =   "frmBase.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   915
      Width           =   330
   End
   Begin MSComDlg.CommonDialog cdlBrowse 
      Left            =   3210
      Top             =   -75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4035
      TabIndex        =   4
      Top             =   1935
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   2940
      TabIndex        =   3
      Top             =   1935
      Width           =   945
   End
   Begin VB.TextBox txtValue 
      Height          =   330
      Left            =   1485
      TabIndex        =   0
      Top             =   900
      Width           =   3150
   End
   Begin VB.ComboBox cboTarget 
      Height          =   315
      ItemData        =   "frmBase.frx":0231
      Left            =   1485
      List            =   "frmBase.frx":0233
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1395
      Width           =   2460
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Href:"
      Height          =   195
      Left            =   810
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Target:"
      Height          =   195
      Left            =   810
      TabIndex        =   6
      Top             =   1455
      Width           =   540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   270
      X2              =   7985
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line1 
      X1              =   315
      X2              =   7985
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Base"
      Height          =   195
      Left            =   780
      TabIndex        =   5
      Top             =   255
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   375
      Picture         =   "frmBase.frx":0235
      Top             =   255
      Width           =   360
   End
End
Attribute VB_Name = "frmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Dim strMeta As String
  strMeta = "<base href=""" & txtValue & """"
  If cboTarget.ListIndex >= 0 Then strMeta = strMeta & " target=""" & cboTarget.Text & """"
  strMeta = strMeta & " >"
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste strMeta
  Unload Me
End Sub

Private Sub cmdBrowse_Click()
  With cdlBrowse
        .FileName = ""
        .CancelError = False
        .Filter = "HTML files (*.htm)|*.htm|ASP Files(*.asp)|*.asp"
        .ShowOpen
  End With
  txtValue = cdlBrowse.FileTitle
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  With cboTarget
    .AddItem ""
    .AddItem "_blank"
    .AddItem "_parent"
    .AddItem "_self"
    .AddItem "_top"
  End With
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub


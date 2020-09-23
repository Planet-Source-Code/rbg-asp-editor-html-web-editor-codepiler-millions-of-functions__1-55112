VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRefresh 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refresh"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   5040
      Picture         =   "frmRefresh.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      Width           =   330
   End
   Begin MSComDlg.CommonDialog cdlBrowse 
      Left            =   5250
      Top             =   -105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtUrl 
      Height          =   300
      Left            =   2820
      TabIndex        =   2
      Top             =   1275
      Width           =   2220
   End
   Begin VB.OptionButton optCh2 
      Caption         =   "Refresh this Document"
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   1695
      Width           =   2160
   End
   Begin VB.OptionButton optCh1 
      Caption         =   "Go To URL :"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1305
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4425
      TabIndex        =   6
      Top             =   2265
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   3315
      TabIndex        =   5
      Top             =   2265
      Width           =   945
   End
   Begin VB.TextBox txtValue 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   810
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      Height          =   195
      Left            =   2145
      TabIndex        =   10
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Action :"
      Height          =   195
      Left            =   735
      TabIndex        =   9
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delay:"
      Height          =   195
      Left            =   735
      TabIndex        =   8
      Top             =   855
      Width           =   465
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   300
      X2              =   8015
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line1 
      X1              =   345
      X2              =   8015
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      Height          =   195
      Left            =   735
      TabIndex        =   7
      Top             =   210
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   435
      Picture         =   "frmRefresh.frx":0231
      Top             =   195
      Width           =   240
   End
End
Attribute VB_Name = "frmRefresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Dim strMeta As String
  strMeta = "<meta http-equiv=""refresh"" content=""" & IIf(IsNumeric(txtValue), txtValue, 0)
  If optCh1.Value = True Then
     strMeta = strMeta & ";URL=" & txtUrl.Text & """>"
  Else
     strMeta = strMeta & """>"
  End If
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
  txtUrl = cdlBrowse.FileTitle
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub

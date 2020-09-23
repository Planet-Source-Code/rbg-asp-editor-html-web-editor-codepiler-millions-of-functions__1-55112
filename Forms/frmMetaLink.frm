VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMetaLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meta Link"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
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
   ScaleHeight     =   3165
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   5745
      Picture         =   "frmMetaLink.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   885
      Width           =   330
   End
   Begin VB.TextBox txtTitle 
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
      Left            =   4065
      TabIndex        =   3
      Top             =   1470
      Width           =   2010
   End
   Begin VB.TextBox txtRel 
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
      Left            =   1575
      TabIndex        =   4
      Top             =   1980
      Width           =   2010
   End
   Begin VB.TextBox txtRev 
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
      Left            =   4065
      TabIndex        =   5
      Top             =   2010
      Width           =   2010
   End
   Begin VB.TextBox txtId 
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
      Left            =   1575
      TabIndex        =   2
      Top             =   1440
      Width           =   2010
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3975
      TabIndex        =   6
      Top             =   2595
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5130
      TabIndex        =   7
      Top             =   2595
      Width           =   945
   End
   Begin VB.TextBox txtUrl 
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
      Left            =   1575
      TabIndex        =   0
      Top             =   900
      Width           =   4140
   End
   Begin MSComDlg.CommonDialog cdlBrowse 
      Left            =   5370
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3660
      TabIndex        =   13
      Top             =   1485
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1185
      TabIndex        =   12
      Top             =   2010
      Width           =   285
   End
   Begin VB.Label lblone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rev:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3660
      TabIndex        =   11
      Top             =   2055
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   540
      Picture         =   "frmMetaLink.frx":0231
      Top             =   165
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meta Link"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1005
      TabIndex        =   10
      Top             =   210
      Width           =   705
   End
   Begin VB.Line Line1 
      X1              =   435
      X2              =   8105
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   390
      X2              =   8105
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1230
      TabIndex        =   9
      Top             =   1455
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Href:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   945
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   645
      Picture         =   "frmMetaLink.frx":0475
      Top             =   180
      Width           =   240
   End
End
Attribute VB_Name = "frmMetaLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
 Dim strMeta As String
 strMeta = "<link href=""" & txtUrl & """"
 If txtId.Text <> "" Then strMeta = strMeta & " id=""" & txtId & """"
 If txtRel.Text <> "" Then strMeta = strMeta & " rel=""" & txtRel & """"
 If txtRev.Text <> "" Then strMeta = strMeta & " rev=""" & txtRev & """"
 If txtTitle.Text <> "" Then strMeta = strMeta & " title=""" & txtTitle & """"
 strMeta = strMeta & ">"
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


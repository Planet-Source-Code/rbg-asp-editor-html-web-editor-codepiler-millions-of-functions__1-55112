VERSION 5.00
Begin VB.Form frmMeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meta"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5565
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
   ScaleHeight     =   3675
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4140
      TabIndex        =   4
      Top             =   3075
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   3045
      TabIndex        =   3
      Top             =   3075
      Width           =   945
   End
   Begin VB.TextBox txtContent 
      Height          =   1395
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1380
      Width           =   3645
   End
   Begin VB.TextBox txtValue 
      Height          =   330
      Left            =   3780
      TabIndex        =   1
      Top             =   795
      Width           =   1305
   End
   Begin VB.ComboBox cboAtt 
      Height          =   315
      ItemData        =   "frmMeta.frx":0000
      Left            =   1455
      List            =   "frmMeta.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   780
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Content:"
      Height          =   195
      Left            =   660
      TabIndex        =   8
      Top             =   1380
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value:"
      Height          =   195
      Left            =   3225
      TabIndex        =   7
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attribute:"
      Height          =   195
      Left            =   660
      TabIndex        =   6
      Top             =   840
      Width           =   705
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   225
      X2              =   7940
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line1 
      X1              =   270
      X2              =   7940
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meta"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   180
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   345
      Picture         =   "frmMeta.frx":0025
      Top             =   150
      Width           =   240
   End
End
Attribute VB_Name = "frmMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
 Dim strMeta As String
 If cboAtt.ListIndex = 0 Then
   strMeta = "<meta name=""" & txtValue & """ content=""" & txtContent & """>"
 ElseIf cboAtt.ListIndex = 1 Then
    strMeta = "<meta http-equiv=""" & txtValue & """ content=""" & txtContent & """>"
 Else
    strMeta = "<meta content=""" & txtContent & """>"
 End If
 frmEditor.RTB(frmEditor.GetActiveRTB).Paste strMeta
 Unload Me
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub


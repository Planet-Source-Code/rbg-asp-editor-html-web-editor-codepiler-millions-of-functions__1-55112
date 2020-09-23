VERSION 5.00
Begin VB.Form frmKeyword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyword"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
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
   ScaleHeight     =   3000
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      Top             =   2430
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   3270
      TabIndex        =   1
      Top             =   2430
      Width           =   945
   End
   Begin VB.TextBox txtContent 
      Height          =   1395
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   765
      Width           =   3645
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword:"
      Height          =   195
      Left            =   750
      TabIndex        =   4
      Top             =   765
      Width           =   690
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   225
      X2              =   7940
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Line Line1 
      X1              =   270
      X2              =   7940
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label lblHead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword"
      Height          =   195
      Left            =   645
      TabIndex        =   3
      Top             =   150
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   300
      Picture         =   "frmKeyword.frx":0000
      Top             =   135
      Width           =   240
   End
End
Attribute VB_Name = "frmKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mType As Integer

Private Sub cmdAdd_Click()
 Dim strMeta As String
  If mType = 1 Then
    strMeta = "<meta name=""keywords"""
  Else
    strMeta = "<meta name=""description"""
  End If
  strMeta = strMeta & " content=""" & txtContent & """>"
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste strMeta
  Unload Me
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Rem================
Rem mType=1  for keyword
Rem mType=2 for description
Rem================
Private Sub Form_Load()

 If mType = 1 Then
   lblHead.Caption = "Keywords"
   Me.Caption = "Keywords"
   lblCaption.Caption = "Keywords"
 Else
   lblHead.Caption = "Description"
   Me.Caption = "Description"
   lblCaption.Caption = "Description"
 End If
End Sub

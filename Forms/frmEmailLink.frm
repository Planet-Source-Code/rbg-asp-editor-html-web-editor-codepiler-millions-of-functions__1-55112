VERSION 5.00
Begin VB.Form frmEmailLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Link"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
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
   Icon            =   "frmEmailLink.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1230
      TabIndex        =   1
      Top             =   1320
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2985
      TabIndex        =   3
      Top             =   1830
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   330
      Left            =   1905
      TabIndex        =   2
      Top             =   1830
      Width           =   945
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   1245
      TabIndex        =   0
      Top             =   855
      Width           =   2700
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      Height          =   195
      Left            =   585
      TabIndex        =   6
      Top             =   1380
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Link"
      Height          =   195
      Left            =   765
      TabIndex        =   5
      Top             =   315
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   345
      X2              =   5385
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   300
      X2              =   5340
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   915
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   420
      Picture         =   "frmEmailLink.frx":000C
      Top             =   285
      Width           =   270
   End
End
Attribute VB_Name = "frmEmailLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
Dim lEmail
  lEmail = "<a href=""mailto:" & txtEmail.Text & """>" & txtText.Text & "</a>"
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste lEmail
  Unload Me
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub

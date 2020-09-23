VERSION 5.00
Begin VB.Form frmGotoline 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Go To Line"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGoto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2850
      TabIndex        =   3
      Top             =   660
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   2865
      TabIndex        =   2
      Top             =   165
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   195
      TabIndex        =   1
      Text            =   "1"
      Top             =   705
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Line Number:"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   390
      Width           =   945
   End
End
Attribute VB_Name = "frmGotoline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  If frmEditor.RTB(frmEditor.GetActiveRTB).GotoLine(Int(Text1.Text)) Then
    Unload Me
  Else
    Text1.Text = "1"
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
  End If
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If Chr(KeyAscii) = "-" Then
      KeyAscii = 0
  Else
      KeyAscii = fnNumKey(KeyAscii)
  End If
End Sub

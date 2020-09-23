VERSION 5.00
Begin VB.Form frmBmark 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anchor"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
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
   Icon            =   "frmBookmark.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1170
      TabIndex        =   0
      Top             =   675
      Width           =   2700
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   330
      Left            =   1845
      TabIndex        =   1
      Top             =   1155
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2925
      TabIndex        =   2
      Top             =   1155
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   255
      Picture         =   "frmBookmark.frx":000C
      Top             =   60
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   525
      TabIndex        =   4
      Top             =   735
      Width           =   465
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   225
      X2              =   5265
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line1 
      X1              =   270
      X2              =   5310
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anchor"
      Height          =   195
      Left            =   660
      TabIndex        =   3
      Top             =   135
      Width           =   510
   End
End
Attribute VB_Name = "frmBmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
Dim Bookmark
  Bookmark = "<a name=""" & Text1.Text & """></a>"
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste Bookmark
  frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = frmEditor.RTB(frmEditor.GetActiveRTB).SelStart - 4
  Unload Me
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub


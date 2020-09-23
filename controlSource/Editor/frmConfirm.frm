VERSION 5.00
Begin VB.Form frmConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Piler"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
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
   ScaleHeight     =   2070
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUpload 
      Height          =   195
      Left            =   742
      TabIndex        =   4
      Top             =   1485
      Width           =   195
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2310
      TabIndex        =   3
      Top             =   930
      Width           =   975
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Default         =   -1  'True
      Height          =   315
      Left            =   1110
      TabIndex        =   2
      Top             =   945
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't show this message in future. Always update the changes"
      Height          =   405
      Left            =   1027
      TabIndex        =   5
      Top             =   1455
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   180
      Picture         =   "frmConfirm.frx":0000
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Do you want to proceed?"
      Height          =   195
      Left            =   1155
      TabIndex        =   1
      Top             =   495
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview needs to upload the changes."
      Height          =   195
      Left            =   1155
      TabIndex        =   0
      Top             =   255
      Width           =   2760
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Confirm(ByVal Result As Boolean, ByVal DontShow As Boolean)

Private Sub cmdCancel_Click()
  RaiseEvent Confirm(True, chkUpload.value)
  Unload Me
End Sub

Private Sub cmdYes_Click()
  RaiseEvent Confirm(True, chkUpload.value)
  Unload Me
End Sub

VERSION 5.00
Begin VB.Form frmLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Link"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ClipControls    =   0   'False
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
   Icon            =   "frmLink.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1650
      TabIndex        =   3
      Top             =   1590
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Text            =   "http://"
      Top             =   675
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmLink.frx":058A
      Left            =   3570
      List            =   "frmLink.frx":059D
      TabIndex        =   1
      Text            =   "_Default"
      Top             =   675
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1650
      TabIndex        =   2
      Top             =   1140
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   2970
      TabIndex        =   5
      Top             =   2475
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4005
      TabIndex        =   6
      Top             =   2475
      Width           =   900
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Make as ASP Code"
      Height          =   270
      Left            =   1650
      TabIndex        =   4
      Top             =   2025
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   360
      Picture         =   "frmLink.frx":05C9
      Top             =   90
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link type"
      Height          =   195
      Left            =   765
      TabIndex        =   11
      Top             =   705
      Width           =   645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Target"
      Height          =   255
      Left            =   2910
      TabIndex        =   10
      Top             =   705
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actual link"
      Height          =   195
      Left            =   660
      TabIndex        =   9
      Top             =   1155
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrption"
      Height          =   195
      Left            =   630
      TabIndex        =   8
      Top             =   1605
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link"
      Height          =   195
      Left            =   690
      TabIndex        =   7
      Top             =   105
      Width           =   270
   End
   Begin VB.Line Line1 
      X1              =   330
      X2              =   6670
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   285
      X2              =   6625
      Y1              =   485
      Y2              =   485
   End
End
Attribute VB_Name = "frmLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Link
  If Check1.Value = False Then
    Link = "<a href=""" & Combo1.Text & Text1.Text & """ target=""" & Combo2.Text & """ >" & Text2.Text & "</a>"
  Else
    Link = "Responce.write(" & Chr(34) & "<a href=""" & Chr(34) & Combo1.Text & Text1.Text & Chr(34) & """ target=""" & Chr(34) & Combo2.Text & Chr(34) & """ >" & Text2.Text & "</a>)" & Chr(34)
  End If
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste Link
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub


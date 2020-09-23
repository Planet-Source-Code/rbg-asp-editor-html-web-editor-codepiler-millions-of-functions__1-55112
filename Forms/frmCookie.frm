VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCookie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cookie Wizard"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
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
   Icon            =   "frmCookie.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1530
      TabIndex        =   0
      Top             =   705
      Width           =   2475
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1515
      TabIndex        =   2
      Top             =   1545
      Width           =   2475
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3090
      TabIndex        =   4
      Top             =   2055
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   2055
      TabIndex        =   3
      Top             =   2055
      Width           =   900
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1530
      TabIndex        =   1
      Top             =   1125
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   19791873
      CurrentDate     =   37555
   End
   Begin VB.Image Image1 
      Height          =   195
      Left            =   270
      Picture         =   "frmCookie.frx":000C
      Top             =   150
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   525
      TabIndex        =   8
      Top             =   750
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry date:"
      Height          =   195
      Left            =   525
      TabIndex        =   7
      Top             =   1170
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
      Height          =   195
      Left            =   525
      TabIndex        =   6
      Top             =   1590
      Width           =   390
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   210
      X2              =   5550
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   255
      X2              =   5595
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cookie"
      Height          =   195
      Left            =   615
      TabIndex        =   5
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmCookie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Link
  Link = "<%" & vbCrLf & vbTab & "Response.Cookies(""Type"") = """ & Text1.Text & """" & vbCrLf & vbTab & "Response.Cookies(""Type"").Expires = """ & DTPicker1.Value & """" & vbCrLf & vbTab & "Response.Cookies(""Type"").Domain =""""" & vbCrLf & vbTab & "Response.Cookies(""Type"").Path = """ & Text3.Text & """" & vbCrLf & vbTab & "Response.Cookies(""Type"").Secure = FALSE " & vbCrLf & "%>" & vbCrLf
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste Link
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub

Private Sub Form_Load()
  DTPicker1.Value = Now
End Sub

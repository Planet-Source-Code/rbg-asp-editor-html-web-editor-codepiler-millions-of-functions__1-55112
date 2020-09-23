VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConnectionString 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connection String Wizard"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
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
   ScaleHeight     =   4305
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDetails 
      BorderStyle     =   0  'None
      Height          =   2430
      Left            =   630
      ScaleHeight     =   2430
      ScaleWidth      =   5250
      TabIndex        =   14
      Top             =   1125
      Visible         =   0   'False
      Width           =   5250
      Begin VB.CommandButton cmdOpen 
         Height          =   315
         Left            =   4485
         Picture         =   "frmConnectionString.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   570
         Width           =   375
      End
      Begin VB.TextBox txtServer 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   1485
         Width           =   2355
      End
      Begin VB.TextBox txtUsername 
         Height          =   315
         Left            =   1545
         TabIndex        =   6
         Top             =   1035
         Width           =   2355
      End
      Begin VB.TextBox txtDBName 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   570
         Width           =   2895
      End
      Begin VB.ComboBox cboDriver 
         Height          =   315
         ItemData        =   "frmConnectionString.frx":0231
         Left            =   1575
         List            =   "frmConnectionString.frx":023B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1950
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   2385
         Left            =   0
         Top             =   0
         Width           =   5205
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   195
         Left            =   375
         TabIndex        =   20
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblDriver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver:"
         Height          =   195
         Left            =   375
         TabIndex        =   18
         Top             =   2010
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   375
         TabIndex        =   17
         Top             =   1545
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   195
         Left            =   375
         TabIndex        =   16
         Top             =   1095
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database:"
         Height          =   195
         Left            =   375
         TabIndex        =   15
         Top             =   630
         Width           =   750
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3795
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   630
      TabIndex        =   11
      Top             =   3750
      Width           =   1065
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< Back"
      Height          =   345
      Left            =   3585
      TabIndex        =   10
      Top             =   3750
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      Default         =   -1  'True
      Height          =   345
      Left            =   4815
      TabIndex        =   9
      Top             =   3750
      Width           =   1065
   End
   Begin VB.PictureBox picDatabase 
      BorderStyle     =   0  'None
      Height          =   2430
      Left            =   630
      ScaleHeight     =   2430
      ScaleWidth      =   5250
      TabIndex        =   13
      Top             =   1125
      Width           =   5250
      Begin VB.OptionButton optMySql 
         Caption         =   "My SQL"
         Height          =   345
         Left            =   1710
         TabIndex        =   2
         Top             =   1560
         Width           =   1605
      End
      Begin VB.OptionButton optSQL 
         Caption         =   "SQL Server"
         Height          =   345
         Left            =   1710
         TabIndex        =   1
         Top             =   1050
         Width           =   1605
      End
      Begin VB.OptionButton optMSAccess 
         Caption         =   "Microsoft Access"
         Height          =   345
         Left            =   1710
         TabIndex        =   0
         Top             =   555
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   2385
         Left            =   0
         Top             =   0
         Width           =   5205
      End
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the database provider"
      Height          =   195
      Left            =   630
      TabIndex        =   19
      Tag             =   "Select the database provider"
      Top             =   795
      Width           =   2085
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection String"
      Height          =   195
      Left            =   1155
      TabIndex        =   12
      Top             =   240
      Width           =   1275
   End
   Begin VB.Line Line1 
      X1              =   420
      X2              =   6620
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   375
      X2              =   6715
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   555
      Picture         =   "frmConnectionString.frx":026F
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmConnectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ConnectionStringOk(ByVal ConnectionString As String)
Public Event CancelWizard()

Private Sub cmdBack_Click()
  picDatabase.Visible = True
  picDetails.Visible = False
  lblCaption.Caption = lblCaption.Tag
  cmdNext.Caption = "Next >>"
  cmdBack.Visible = False
End Sub

Private Sub cmdCancel_Click()
  RaiseEvent CancelWizard
  Unload Me
End Sub

Private Sub cmdNext_Click()
Dim lStr As String
  If cmdNext.Caption = "Next >>" Then
    lblCaption.Caption = "Enter the database details"
    picDatabase.Visible = False
    picDetails.Visible = True
    If optMSAccess.Value Then
      txtServer.Visible = False
      lblServer.Visible = False
      cboDriver.Visible = False
      lblDriver.Visible = False
      cmdOpen.Visible = True
    ElseIf optSQL.Value Then
      txtServer.Visible = True
      lblServer.Visible = True
      cboDriver.Visible = False
      lblDriver.Visible = False
      cmdOpen.Visible = False
    ElseIf optMySql.Value Then
      txtServer.Visible = True
      lblServer.Visible = True
      cboDriver.Visible = True
      lblDriver.Visible = True
      cmdOpen.Visible = False
    End If
    cmdBack.Visible = True
    cmdNext.Caption = "Finish"
  Else
    If optMSAccess.Value Then
      lStr = "Driver={Microsoft Access Driver (*.mdb)};"
      lStr = lStr & "DBQ=" & txtDBName.Text & ";"
      lStr = lStr & "UID=" & txtUsername.Text & ";"
      lStr = lStr & "PWD=" & txtPassword.Text & ";"
    ElseIf optSQL.Value Then
      lStr = "Driver={SQL Server};"
      lStr = lStr & "SERVER=" & txtServer.Text & ";"
      lStr = lStr & "DATABASE=" & txtDBName.Text & ";"
      lStr = lStr & "UID=" & txtUsername.Text & ";"
      lStr = lStr & "PWD=" & txtPassword.Text & ";"
    ElseIf optMySql.Value Then
      lStr = "Driver={" & cboDriver.Text & "};"
      lStr = lStr & "SERVER=" & txtServer.Text & ";"
      lStr = lStr & "DATABASE=" & txtDBName.Text & ";"
      lStr = lStr & "UID=" & txtUsername.Text & ";"
      lStr = lStr & "PWD=" & txtPassword.Text & ";"
    End If
    
    RaiseEvent ConnectionStringOk(lStr)
    Unload Me
  End If
End Sub

Private Sub cmdOpen_Click()
  CD.FileName = ""
  CD.Filter = "Access Files (*.mdb)|*.mdb"
  CD.CancelError = False
  CD.ShowOpen
  If CD.FileName <> "" Then
    txtDBName.Text = CD.FileName
  End If
End Sub

Private Sub Form_Load()
  cboDriver.ListIndex = 0
End Sub

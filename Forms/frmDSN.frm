VERSION 5.00
Begin VB.Form frmDSN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DSN Connection"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
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
   Icon            =   "frmDSN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4620
      TabIndex        =   7
      Top             =   3630
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   3645
      TabIndex        =   6
      Top             =   3630
      Width           =   810
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   2100
      TabIndex        =   5
      Top             =   3060
      Width           =   3330
   End
   Begin VB.TextBox txtRecordset 
      Height          =   330
      Left            =   2100
      TabIndex        =   4
      Top             =   2619
      Width           =   3330
   End
   Begin VB.TextBox txtSql 
      Height          =   315
      Left            =   2100
      TabIndex        =   3
      Top             =   2193
      Width           =   3330
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   2100
      TabIndex        =   0
      Top             =   870
      Width           =   3330
   End
   Begin VB.TextBox txtPassword 
      Height          =   330
      Left            =   2100
      TabIndex        =   1
      Top             =   1311
      Width           =   3330
   End
   Begin VB.TextBox txtUser 
      Height          =   330
      Left            =   2100
      TabIndex        =   2
      Top             =   1752
      Width           =   3330
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   285
      Picture         =   "frmDSN.frx":058A
      Top             =   210
      Width           =   240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DSN Connection"
      Height          =   195
      Left            =   615
      TabIndex        =   14
      Top             =   240
      Width           =   1155
   End
   Begin VB.Line Line1 
      X1              =   315
      X2              =   7985
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   270
      X2              =   7985
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Query:"
      Height          =   195
      Left            =   675
      TabIndex        =   13
      Top             =   3060
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recordset Name:"
      Height          =   195
      Left            =   675
      TabIndex        =   12
      Top             =   2685
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sql Name:"
      Height          =   195
      Left            =   675
      TabIndex        =   11
      Top             =   2250
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      Height          =   195
      Left            =   705
      TabIndex        =   10
      Top             =   1380
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   705
      TabIndex        =   9
      Top             =   945
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Left            =   690
      TabIndex        =   8
      Top             =   1815
      Width           =   750
   End
End
Attribute VB_Name = "frmDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim lStr As String
  lStr = "<%" & vbCrLf & _
                vbTab & "'" & vbCrLf & _
                vbTab & "' Open the database connection with ADO Object using DSN" & vbCrLf & _
                vbTab & "'" & vbCrLf & _
                vbTab & "Set lConnection = Server.CreateObject(""ADODB.Connection"")" & vbCrLf & _
                vbTab & "Set " & txtRecordset.Text & "= Server.CreateObject(""ADODB.RecordSet"")" & vbCrLf & _
                vbTab & "lConnection.Open ""DSN=" & txtName.Text & ";UID=" & txtUser.Text & ";PWD=" & txtPassword.Text & """" & vbCrLf & _
                vbTab & txtSql.Text & " = """ & txtQuery.Text & """" & vbCrLf & _
                vbTab & "Set " & txtRecordset.Text & " = lConnection.Execute(" & txtSql.Text & ")" & vbCrLf & _
                "%>" & vbCrLf & _
                "<%" & vbCrLf & _
                vbTab & "'" & vbCrLf & _
                vbTab & "' Close the connections" & vbCrLf & _
                vbTab & "'" & vbCrLf & _
                vbTab & txtRecordset.Text & ".Close" & vbCrLf & _
                vbTab & "lConnection.Close" & vbCrLf & "%>" & vbCrLf
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste lStr
  Unload Me
End Sub

Private Sub Command5_Click()
  Unload Me
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub

Private Sub Picture2_Click()

End Sub

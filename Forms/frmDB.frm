VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database connection"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
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
   Icon            =   "frmDB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   5130
      Picture         =   "frmDB.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   825
      Width           =   330
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   1995
      TabIndex        =   3
      Top             =   2250
      Width           =   3075
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   1995
      TabIndex        =   2
      Top             =   1770
      Width           =   3075
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1995
      TabIndex        =   1
      Top             =   1290
      Width           =   3075
   End
   Begin VB.TextBox Text5 
      Height          =   330
      Left            =   1995
      TabIndex        =   4
      Top             =   2730
      Width           =   3075
   End
   Begin VB.TextBox txtDB 
      Height          =   330
      Left            =   1995
      TabIndex        =   0
      Top             =   810
      Width           =   3075
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4215
      TabIndex        =   6
      Top             =   3300
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   3195
      TabIndex        =   5
      Top             =   3300
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   195
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   390
      Picture         =   "frmDB.frx":07BB
      Top             =   195
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Query Name:"
      Height          =   195
      Left            =   510
      TabIndex        =   12
      Top             =   1845
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Name:"
      Height          =   195
      Left            =   510
      TabIndex        =   11
      Top             =   1365
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database Name:"
      Height          =   195
      Left            =   510
      TabIndex        =   10
      Top             =   885
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Query String:"
      Height          =   195
      Left            =   510
      TabIndex        =   9
      Top             =   2805
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recordset Name:"
      Height          =   195
      Left            =   510
      TabIndex        =   8
      Top             =   2325
      Width           =   1245
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   255
      X2              =   6595
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line1 
      X1              =   300
      X2              =   6500
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database Connection"
      Height          =   195
      Left            =   750
      TabIndex        =   7
      Top             =   225
      Width           =   1545
   End
End
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim lStr As String
  lStr = "<%" & vbCrLf & _
                vbTab & "'" & vbCrLf & _
                vbTab & "' Open the database connection with ADO Object" & vbCrLf & _
                vbTab & "'" & vbCrLf & _
                vbTab & "Set lConnection = Server.CreateObject(""ADODB.Connection"")" & vbCrLf & _
                vbTab & "Set " & Text4.Text & "= Server.CreateObject(""ADODB.RecordSet"")" & vbCrLf & _
                vbTab & Text2.Text & " = ""DRIVER={Microsoft Access Driver (*.mdb)} DBQ="" & Server.MapPath(""" & txtDB.Text & """)" & vbCrLf & _
                vbTab & "lConnection.Open " & Text2.Text & vbCrLf & _
                vbTab & Text3.Text & " = " & Text5.Text & vbCrLf & _
                vbTab & "Set " & Text4.Text & " = lConnection.Execute(" & Text3.Text & ")" & vbCrLf & _
                "%>" & vbCrLf & _
                "<%" & vbCrLf & _
                vbTab & "'" & vbCrLf & _
                vbTab & "' Close the connections" & vbCrLf & _
                vbTab & "'" & vbCrLf & _
                vbTab & Text4.Text & ".Close" & vbCrLf & _
                vbTab & "lConnection.Close" & vbCrLf & "%>" & vbCrLf
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste lStr
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub cmdBrowse_Click()
Dim sFile As String
  With CommonDialog1
    .DialogTitle = "Open"
    .CancelError = False
    .Filter = "Access Database (*.mdb)|*.mdb"
    .DefaultExt = "Access Database"
    .ShowOpen
    If Len(.FileName) = 0 Then
      Exit Sub
    End If
    sFile = .FileName
    txtDB.Text = sFile
  End With
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub


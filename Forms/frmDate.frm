VERSION 5.00
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
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
   Icon            =   "frmDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3495
      TabIndex        =   4
      Top             =   3615
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   330
      Left            =   2385
      TabIndex        =   3
      Top             =   3615
      Width           =   945
   End
   Begin VB.ComboBox cboTime 
      Height          =   315
      ItemData        =   "frmDate.frx":000C
      Left            =   2085
      List            =   "frmDate.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2985
      Width           =   2010
   End
   Begin VB.ComboBox cboDay 
      Height          =   315
      ItemData        =   "frmDate.frx":0036
      Left            =   2085
      List            =   "frmDate.frx":004F
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2460
      Width           =   2010
   End
   Begin VB.ListBox lsDate 
      Height          =   1425
      ItemData        =   "frmDate.frx":0082
      Left            =   2085
      List            =   "frmDate.frx":00A7
      TabIndex        =   0
      Top             =   825
      Width           =   2355
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time Format:"
      Height          =   195
      Left            =   915
      TabIndex        =   8
      Top             =   3045
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day Format:"
      Height          =   195
      Left            =   915
      TabIndex        =   7
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Format:"
      Height          =   195
      Left            =   915
      TabIndex        =   6
      Top             =   825
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   915
      TabIndex        =   5
      Top             =   255
      Width           =   345
   End
   Begin VB.Line Line1 
      X1              =   510
      X2              =   7550
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   465
      X2              =   7505
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   510
      Picture         =   "frmDate.frx":0131
      Stretch         =   -1  'True
      Top             =   210
      Width           =   315
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  'Load date format
  lsDate.ListIndex = 0
  cboDay.ListIndex = 0
  cboTime.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

'Date formats
'22/04/1985
'22-04-1985
'22.04.1985
'22/4/85
'4/22/85
'04/22/1985
'1985-04-22
'April 22, 1985
'22 Apr 1985
'22-apr-85
'22 April , 1985
'day
'Monday,
'Monday
'Mon,
'Mon
'mon,
'mon
Private Sub cmdAdd_Click()
Dim lDay As String
Dim lTime As String
Dim lDate As String
Dim lCode As String
  lCode = ""
  'Day
  lDay = ""
  Select Case cboDay.ListIndex
  Case 1
    lDay = Format(Now, "dddd,")
  Case 2
    lDay = Format(Now, "dddd")
  Case 3
    lDay = Format(Now, "ddd,")
  Case 4
    lDay = Format(Now, "ddd")
  Case 5
    lDay = LCase(Format(Now, "ddd,"))
  Case 6
    lDay = LCase(Format(Now, "ddd"))
  End Select
  If lDay <> "" Then lCode = lCode & lDay & " "
  'Date
  lDate = ""
  Select Case lsDate.ListIndex
  Case 0
    lDate = Format(Now, "dd/MM/yyyy")
  Case 1
    lDate = Format(Now, "dd-MM-yyyy")
  Case 2
    lDate = Format(Now, "dd.MM.yyyy")
  Case 3
    lDate = Format(Now, "d/M/yy")
  Case 4
    lDate = LCase(Format(Now, "M/d/yy"))
  Case 5
    lDate = LCase(Format(Now, "MM/dd/yyyy"))
  Case 6
    lDate = Format(Now, "yyyy-MM-dd")
  Case 7
    lDate = Format(Now, "MMMM dd, yyyy")
  Case 8
    lDate = Format(Now, "dd MMM yyyy")
  Case 9
    lDate = LCase(Format(Now, "dd-MMM-yy"))
  Case 10
    lDate = LCase(Format(Now, "dd MMMM, yyyy"))
  End Select
  If lDate <> "" Then lCode = lCode & lDate & " "
  'time
  lTime = ""
  If cboTime.ListIndex = 1 Then
    lTime = Format(Now, "hh:mm AM/PM")
  ElseIf cboTime.ListIndex = 2 Then
    lTime = Format(Now, "hh:mm")
  End If
  If lTime <> "" Then lCode = lCode & lTime
  
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste lCode
  Unload Me
End Sub



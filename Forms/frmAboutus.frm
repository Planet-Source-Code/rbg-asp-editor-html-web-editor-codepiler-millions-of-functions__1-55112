VERSION 5.00
Begin VB.Form frmAboutus 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAboutus.frx":0000
   ScaleHeight     =   5190
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -2000
      TabIndex        =   3
      Top             =   165
      Width           =   405
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.codepiler.com"
      Height          =   195
      Left            =   5805
      MouseIcon       =   "frmAboutus.frx":B7A9
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4020
      Width           =   1890
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4185
      TabIndex        =   6
      Top             =   450
      Width           =   1335
   End
   Begin VB.Label lblRegisteredTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "user"
      Height          =   195
      Left            =   5130
      TabIndex        =   5
      Top             =   1575
      Width           =   315
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registered to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   4
      Top             =   1245
      Width           =   1200
   End
   Begin VB.Label lblRemaining 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   2
      Top             =   3090
      Width           =   120
   End
   Begin VB.Label lblKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Key: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   1
      Top             =   2085
      Width           =   1500
   End
   Begin VB.Label lblRegistrationKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Key"
      Height          =   195
      Left            =   4515
      TabIndex        =   0
      Top             =   2460
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   5190
      Left            =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmAboutus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'RGB value of lable
' 118,151,196 'status
' 174,173,179 'Remaining days,key lable
' 59,87,135 'Registration key
'

Rem -------------------------------
Rem      Private Declarations
Rem -------------------------------
Private Gboltrl As Boolean 'Trail
Private Guser As String
Private Gbolexpired As Boolean 'Expired
Private Gintdays As Integer 'No of days

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim lDatereg As Date
Dim lDuration As Integer
Dim lDays As Integer
  'On Error GoTo Cerr
  lblRemaining.Visible = False
  lblURL.ForeColor = RGB(59, 87, 135)
  lblUser.ForeColor = RGB(174, 173, 179)
  lblRegisteredTo.ForeColor = RGB(118, 151, 196)
  lblKey.ForeColor = RGB(174, 173, 179)
  lblRegistrationKey.ForeColor = RGB(59, 87, 135)
  lblVersion.ForeColor = RGB(174, 173, 179)
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  If GetSetting(App.Title, "regkey", "regkey") = "" Then
    lblKey.Visible = False
    lblRegistrationKey.Visible = False
    lblRegisteredTo.ForeColor = vbRed
    lblRegisteredTo.Caption = "Evaluation version. Try it..."
  Else
    lblRemaining.Visible = True
    lblKey.Visible = True
    lblRegistrationKey.Visible = True
    lblRemaining.ForeColor = RGB(174, 173, 179)
    lblRegisteredTo.Caption = GetSetting(App.Title, "reguser", "reguser")
    lblRegistrationKey.Caption = GetSetting(App.Title, "regkey", "regkey")
    'harddiskkey verification
    
        'Expire
        If Not IsExpired Then
          lDatereg = CDate(GetSetting(App.Title, "regdate", "regdate"))
          lDuration = val(GetSetting(App.Title, "regvalid", "regvalid"))
          lDays = DateDiff("d", lDatereg, DateAdd("m", lDuration, lDatereg))
          If DateDiff("d", lDatereg, Date) >= 0 Then
            lblRemaining.Caption = "Remaining days... " & lDays - DateDiff("d", lDatereg, Date)
          End If
        Else
          Me.MousePointer = vbDefault
          lblRemaining.ForeColor = vbRed
          lblRemaining.Caption = "Date is expired!"
        End If
  End If
  Exit Sub
'Cerr:
'  MsgBox Err.Description,vbInformation,mtitle
End Sub

Private Function IsExpired() As Boolean
'
'StrRgd - Registered date
'LstOpn - Last time it is opened
'
Dim strDate                 As String
Dim oldDate                 As String
Dim strSplash               As String
    On Error GoTo Cerr
    
    strDate = GetMySetting("StrRgd")
    oldDate = GetMySetting("LstOpn")
    If strDate = "" Then 'Fresh User
        SaveMySetting "StrRgd", Encode(CDate(GetSetting(App.Title, "regdate", "regdate")))
        SaveMySetting "LstOpn", Encode(Now)
        strDate = CDate(GetSetting(App.Title, "regdate", "regdate"))
        oldDate = Now
    Else
        strDate = Decode(strDate)
        If oldDate = "" Then    'For Checking Installed trial version date
            Gbolexpired = True
        Else
            oldDate = Decode(oldDate)
        End If
        If Not Gbolexpired Then         'For Malpractice attempting to Modify system date
            If CDate(oldDate) < CDate(strDate) Then Gbolexpired = True
            If CDate(oldDate) > Now Then Gbolexpired = True
        End If
    End If
    If Not Gbolexpired Then             'For Checking Expiry
        SaveMySetting "LstOpn", Encode(Now)
        Gintdays = DateDiff("d", Now, CDate(strDate) + 365)
        If val(Gintdays) <= 0 Then Gbolexpired = True
    End If
    IsExpired = Gbolexpired
    Exit Function
Cerr:
    IsExpired = False
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If lblURL.ForeColor <> RGB(59, 87, 135) Then lblURL.ForeColor = RGB(59, 87, 135)
End Sub

Private Sub lblURL_Click()
  OpenWebsite lblURL.Caption
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblURL.ForeColor = vbBlue
End Sub

Private Sub OpenWebsite(strWebsite As String)
    If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, 1) < 33 Then
    End If
    DoEvents
End Sub


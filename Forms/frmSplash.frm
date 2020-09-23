VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
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
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   5190
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register Now"
      Height          =   345
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6720
      Top             =   2205
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   6825
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblURL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.codepiler.com"
      Height          =   195
      Left            =   5985
      MouseIcon       =   "frmSplash.frx":B7A9
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3990
      Width           =   1890
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Left            =   4230
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   2610
   End
   Begin VB.Image imgProgress 
      Height          =   165
      Left            =   4275
      Picture         =   "frmSplash.frx":C073
      Tag             =   "2520"
      Top             =   3405
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Label lblRegistrationKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Key"
      Height          =   195
      Left            =   4545
      TabIndex        =   7
      Top             =   2325
      Width           =   1185
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
      Left            =   4230
      TabIndex        =   6
      Top             =   1950
      Width           =   1500
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
      Left            =   4230
      TabIndex        =   5
      Top             =   2955
      Width           =   120
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
      Left            =   4230
      TabIndex        =   4
      Top             =   1110
      Width           =   1200
   End
   Begin VB.Label lblRegisteredTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "user"
      Height          =   195
      Left            =   5160
      TabIndex        =   3
      Top             =   1440
      Width           =   315
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
      Left            =   4215
      TabIndex        =   2
      Top             =   315
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3840
      TabIndex        =   0
      Top             =   4470
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   5190
      Left            =   0
      Top             =   0
      Width           =   8205
   End
End
Attribute VB_Name = "frmSplash"
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

Rem -----------------------------
Rem     Private Declarations
Rem -----------------------------
Private mStop As Boolean
Private mRegistered As Boolean

Private Sub cmdClose_Click()
  mStop = True
  If Mreuse Then Unload Me
End Sub

Private Sub cmdRegister_Click()
  frmRegistration.Show vbModal
  If Not Mreuse Then
    If val(GetSetting(App.Title, "regvalid", "regvalid")) > 0 Then
      mRegistered = True
    End If
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If lblURL.ForeColor <> RGB(59, 87, 135) Then lblURL.ForeColor = RGB(59, 87, 135)
End Sub

Private Sub lblURL_Click()
  OpenWebsite lblURL.Caption
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblURL.ForeColor = vbBlue
End Sub

Private Sub Form_Click()
  If Mreuse Then Unload Me
End Sub

Private Sub Form_Load()
  'Title
  Mtitle = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  'Color
  lblStatus.ForeColor = RGB(118, 151, 196)
  lblUser.ForeColor = RGB(174, 173, 179)
  lblRegisteredTo.ForeColor = RGB(118, 151, 196)
  lblKey.ForeColor = RGB(174, 173, 179)
  lblRegistrationKey.ForeColor = RGB(59, 87, 135)
  lblVersion.ForeColor = RGB(174, 173, 179)
  lblRemaining.ForeColor = RGB(174, 173, 179)
  lblURL.ForeColor = RGB(59, 87, 135)
  cmdClose.BackColor = RGB(118, 151, 196)
  cmdRegister.BackColor = RGB(118, 151, 196)
  'Visible
  lblStatus.Visible = False
  lblRemaining.Visible = False
  lblUser.Visible = False
  lblRegisteredTo.Visible = False
  lblKey.Visible = False
  lblRegistrationKey.Visible = False
  lblRemaining.Visible = False
  lblURL.Visible = False
  cmdClose.Visible = False
  imgProgress.Visible = True
  cmdRegister.Visible = True
  'load for reuse
  If Mreuse Then
    CheckRegistration
    Me.MousePointer = vbDefault
    lblStatus.Visible = False
    lblURL.Visible = True
    cmdClose.Visible = True
    cmdClose.Left = -20000
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Mreuse = True
End Sub

Public Function CheckRegistration() As Boolean
'
'Check the registration
'
Dim lKey As String
Dim lDatereg As Date
Dim lDuration As Integer
Dim lDays As Integer
  On Error GoTo Cerr
  Show
  If Mreuse = False Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
  'Checking
  If GetSetting(App.Title, "regkey", "regkey") = "" Then
    'Check trial version
    If CheckTrailExpire Then
      FillTrailExpiry
      CheckRegistration = False
    Else
      FillTrailDetails
      CheckRegistration = True
    End If
  Else
    'For registered version
    'harddiskkey verification
    lKey = GetSetting(App.Title, "regkey", "regkey")
    lKey = Mid(lKey, 1, InStrRev(lKey, "-") - 1)
    lKey = Replace(lKey, "-", "")
    If GetDecodedKey(GetHardDiskKey(), 20) = lKey Then
      'Check registered version
      If CheckRealExpire Then
        FillRealExpiry
        CheckRegistration = False
      Else
        FillRealDetails
        CheckRegistration = True
      End If
    Else
      FillNotRegistered
      CheckRegistration = False
    End If
  End If
  'waiting for close from close button
  If CheckRegistration = False And Mreuse = False Then
    Do Until mStop Or mRegistered
      DoEvents
    Loop
    If mRegistered Then CheckRegistration = True
  End If
  Exit Function
Cerr:
  S110_WriteLog "Check registration... " & Err.Description
  Err.Clear
End Function

Private Function CheckRealExpire() As Boolean
'
'StrRgd - Registered date
'LstOpn - Last time it is opened
'
Dim strDate As String
Dim oldDate As String
Dim lDuration As Integer
Dim lExpire As Boolean
    On Error GoTo Cerr
    
    strDate = GetMySetting("StrRgd")
    oldDate = GetMySetting("LstOpn")
    If strDate = "" Then 'Fresh User
        SaveMySetting "StrRgd", Encode(CDate(GetSetting(App.Title, "regdate", "regdate")))
        strDate = CDate(GetSetting(App.Title, "regdate", "regdate"))
        SaveMySetting "LstOpn", Encode(Now)
        oldDate = Now
    Else
        strDate = Decode(strDate)
        If oldDate = "" Then    'For Checking Installed trial version date
            lExpire = True
        Else
            oldDate = Decode(oldDate) 'For Malpractice attempting to Modify system date
            If CDate(oldDate) < CDate(strDate) Then lExpire = True
            If CDate(oldDate) > Now Then lExpire = True
        End If
    End If
    If Not lExpire Then             'For Checking Expiry
        SaveMySetting "LstOpn", Encode(Now)
        lDuration = val(GetSetting(App.Title, "regvalid", "regvalid"))
        If val(DateDiff("d", Now, DateAdd("M", lDuration, CDate(strDate)))) <= 0 Then lExpire = True
    End If
    If lExpire Then Call SaveSetting(App.Title, "regvalid", "regvalid", "0")
    CheckRealExpire = lExpire
    Exit Function
Cerr:
    lExpire = False
    S110_WriteLog "Checkrealexpire... " & Err.Description
    Err.Clear
End Function

Private Function CheckTrailExpire() As Boolean
'
'StrRgd - Registered date
'LstOpn - Last time it is opened
'
Dim strDate As String
Dim oldDate As String
Dim lExpire As Boolean
    On Error GoTo Cerr
    
    strDate = GetMySetting("StrRgd")
    oldDate = GetMySetting("LstOpn")
    If strDate = "" Then 'Fresh User
        SaveMySetting "StrRgd", Encode(Date)
        SaveMySetting "LstOpn", Encode(Now)
        strDate = Date
        oldDate = Now
    Else
        strDate = Decode(strDate)
        If oldDate = "" Then    'For Checking Installed trial version date
            lExpire = True
        Else
            oldDate = Decode(oldDate) 'For Malpractice attempting to Modify system date
            If CDate(oldDate) < CDate(strDate) Then lExpire = True
            If CDate(oldDate) > Now Then lExpire = True
        End If
    End If
    If Not lExpire Then             'For Checking Expiry
        SaveMySetting "LstOpn", Encode(Now)
        If val(DateDiff("d", Now, DateAdd("d", 15, CDate(strDate)))) <= 0 Then lExpire = True
    End If
    CheckTrailExpire = lExpire
    Exit Function
Cerr:
    lExpire = False
    S110_WriteLog "Checktrailexpire... " & Err.Description
    Err.Clear
End Function


Private Sub FillTrailDetails()
'
'Fill the trail details
'
Dim lDatereg As Date
Dim lRemaining As Integer
Dim lWidth As Integer
Dim lDays As Integer
  On Error GoTo Cerr
  Me.MousePointer = vbHourglass
  lDatereg = Decode(GetMySetting("StrRgd"))
  lDays = 15

  lblStatus.Visible = True
  lblRemaining.Visible = True
  lblRegisteredTo.Visible = True
  lblRegisteredTo.Caption = "Evaluation version. Try it..."
  If DateDiff("d", lDatereg, Date) >= 0 Then
    imgProgress.Visible = True
    lRemaining = lDays - DateDiff("d", lDatereg, Date)
    lWidth = CInt(imgProgress.Tag) - CInt((CInt(imgProgress.Tag) / lDays) * lRemaining)
    lblRemaining.Caption = "Remaining days... " & lRemaining
    imgProgress.Width = lWidth
  End If
  Exit Sub
Cerr:
  S110_WriteLog "Filltraildetails... " & Err.Description
  Err.Clear
End Sub

Private Sub FillTrailExpiry()
'
'Fill the trail expiry details
'
  cmdClose.Visible = True
  lblRegisteredTo.Visible = True
  lblRemaining.Visible = True
  lblURL.Visible = True
  lblRegisteredTo.ForeColor = vbRed
  lblRegisteredTo.Caption = "Evaluation version. Expired!"
  lblRemaining.Caption = "15 day(s) evaluation completed!"
  lblURL.Caption = "Purchase Codepiler at www.codepiler.com"
End Sub

Private Sub FillRealDetails()
'
'Fill the trail expiry details
'
Dim lDatereg As Date
Dim lDuration As Integer
Dim lRemaining As Integer
Dim lWidth As Integer
Dim lDays As Integer
  On Error GoTo Cerr
  lDatereg = Decode(GetMySetting("StrRgd"))
  lDuration = val(GetSetting(App.Title, "regvalid", "regvalid"))
  lDays = DateDiff("d", lDatereg, DateAdd("m", lDuration, lDatereg))
  
  Me.MousePointer = vbHourglass
  cmdRegister.Visible = False
  lblStatus.Visible = True
  lblRegisteredTo.Visible = True
  lblRemaining.Visible = True
  lblRegistrationKey.Visible = True
  lblUser.Visible = True
  lblKey.Visible = True
  lblRegisteredTo.Caption = GetSetting(App.Title, "reguser", "reguser")
  lblRegistrationKey.Caption = GetSetting(App.Title, "regkey", "regkey")
  If DateDiff("d", lDatereg, Date) >= 0 Then
    imgProgress.Visible = True
    lRemaining = lDays - DateDiff("d", lDatereg, Date)
    lWidth = CInt(imgProgress.Tag) - (CInt(imgProgress.Tag) / lDays) * lRemaining
    lblRemaining.Caption = "Remaining days... " & lRemaining
    imgProgress.Width = lWidth
  End If
  Exit Sub
Cerr:
  S110_WriteLog "Fillrealdetails... " & Err.Description
  Err.Clear
End Sub

Private Sub FillRealExpiry()
'
'Fill the real expiry details
'
Dim lDuration As Integer
  lDuration = val(GetSetting(App.Title, "regvalid", "regvalid"))
  lblRemaining.Visible = True
  lblRegisteredTo.Visible = True
  lblRegistrationKey.Visible = True
  lblUser.Visible = True
  lblKey.Visible = True
  cmdClose.Visible = True
  lblURL.Visible = True
  lblRegisteredTo.Caption = GetSetting(App.Title, "reguser", "reguser")
  lblRegistrationKey.Caption = GetSetting(App.Title, "regkey", "regkey")
  lblRemaining.ForeColor = vbRed
  lblRemaining.Caption = "Date is expired!"
  lblRemaining.Caption = (lDuration * 12) & " day(s) evaluation completed!"
  lblURL.Caption = "Purchase Codepiler at www.codepiler.com"
End Sub

Private Sub FillNotRegistered()
'
'Fill the invalid registration
'
  cmdRegister.Visible = False
  cmdClose.Visible = True
  lblURL.Visible = True
  lblURL.Caption = "Purchase Codepiler at www.codepiler.com"
  lblRegisteredTo.ForeColor = vbRed
  lblRegisteredTo.Caption = "Not registered in this computer!"
End Sub

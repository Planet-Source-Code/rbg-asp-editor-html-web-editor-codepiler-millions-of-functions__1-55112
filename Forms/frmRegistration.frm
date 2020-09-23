VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmRegistration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Codepiler Registration"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet iNet 
      Left            =   300
      Top             =   5055
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   7665
      TabIndex        =   3
      Top             =   4980
      Width           =   1065
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6435
      TabIndex        =   2
      Top             =   4980
      Width           =   1065
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   345
      Left            =   5220
      TabIndex        =   1
      Top             =   4980
      Width           =   1065
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3990
      TabIndex        =   0
      Top             =   4980
      Width           =   1065
   End
   Begin VB.PictureBox picScreens 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Index           =   2
      Left            =   4170
      ScaleHeight     =   4665
      ScaleWidth      =   4845
      TabIndex        =   16
      Top             =   0
      Width           =   4845
      Begin VB.TextBox txtHarddiskkey 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   210
         TabIndex        =   30
         Top             =   2595
         Width           =   2925
      End
      Begin VB.TextBox txtKey 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   3855
         TabIndex        =   23
         Top             =   3450
         Width           =   810
      End
      Begin VB.TextBox txtKey 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   2943
         TabIndex        =   22
         Top             =   3450
         Width           =   810
      End
      Begin VB.TextBox txtKey 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   2032
         TabIndex        =   21
         Top             =   3450
         Width           =   810
      End
      Begin VB.TextBox txtKey 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1121
         TabIndex        =   20
         Top             =   3450
         Width           =   810
      End
      Begin VB.TextBox txtKey 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   3450
         Width           =   810
      End
      Begin VB.Image imgError 
         Height          =   720
         Left            =   3960
         Picture         =   "frmRegistration.frx":000C
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblGetKey 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "admin@codepiler.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3120
         MouseIcon       =   "frmRegistration.frx":07BA
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   2160
         Width           =   1560
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial no"
         Height          =   195
         Left            =   210
         TabIndex        =   29
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Key"
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   3150
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   3780
         TabIndex        =   27
         Top             =   3518
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   2865
         TabIndex        =   26
         Top             =   3518
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   1950
         TabIndex        =   25
         Top             =   3518
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   1050
         TabIndex        =   24
         Top             =   3518
         Width           =   60
      End
      Begin VB.Label lblMainStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration is completed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   18
         Top             =   1545
         Width           =   4515
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration is completed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   17
         Top             =   945
         Width           =   3690
      End
   End
   Begin VB.PictureBox picScreens 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Index           =   1
      Left            =   4260
      ScaleHeight     =   4665
      ScaleWidth      =   4845
      TabIndex        =   9
      Top             =   0
      Width           =   4845
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1395
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   2760
         Width           =   2505
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   1395
         TabIndex        =   14
         Top             =   2280
         Width           =   2505
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   2820
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   2340
         Width           =   420
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Email address and password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   315
         TabIndex        =   11
         Top             =   945
         Width           =   4815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Give your email address and password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   345
         TabIndex        =   10
         Top             =   1665
         Width           =   3090
      End
   End
   Begin VB.PictureBox picScreens 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Index           =   0
      Left            =   4185
      ScaleHeight     =   4665
      ScaleWidth      =   4845
      TabIndex        =   4
      Top             =   0
      Width           =   4845
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Online"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1500
         TabIndex        =   8
         Top             =   2655
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1500
         TabIndex        =   7
         Top             =   3075
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This will register Codepiler on your computer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   345
         TabIndex        =   6
         Top             =   2010
         Width           =   3660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Codepiler Registration Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   315
         TabIndex        =   5
         Top             =   945
         Width           =   4815
      End
   End
   Begin VB.Image imgScreen 
      Height          =   4680
      Left            =   0
      Picture         =   "frmRegistration.frx":1084
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9000
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9000
      Y1              =   4680
      Y2              =   4680
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIndex As Integer 'To maintain the screen show

Private Sub cmdBack_Click()
  mIndex = mIndex - 1
  ShowScreen mIndex
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFinish_Click()
'
Dim lDate As String
Dim lDuration As Integer
Dim lKey As String
Dim lSno As String
Dim li As Integer
    lKey = ""
    For li = 0 To 4
      If Trim(txtKey(li).Text) <> "" Then
        lKey = lKey & txtKey(li).Text & "-"
      Else
        lKey = ""
        Exit For
      End If
    Next
    If lKey <> "" Then
      lSno = txtKey(0) & txtKey(1) & txtKey(2) & txtKey(3)
      If GetDecodedKey(GetHardDiskKey(), 20) = lSno Then
        SaveSetting App.Title, "regkey", "regkey", Left(lKey, Len(lKey) - 1)
        'Initiate the Settings
        InitiateSettings
        GetDateSettings txtKey(4).Text, lDate, lDuration
        SaveMySetting "StrRgd", ""
        SaveMySetting "LstOpn", ""
        SaveSetting App.Title, "regdate", "regdate", lDate
        SaveSetting App.Title, "regvalid", "regvalid", lDuration
        SaveSetting App.Title, "reguser", "reguser", txtEmail.Text
        'disable the registration menu
        Unload Me
      Else
        MsgBox "Invalid registration key.", vbInformation, Mtitle
      End If
    Else
      MsgBox "Invalid registration key.", vbInformation, Mtitle
    End If
End Sub

Private Sub cmdNext_Click()
'
  mIndex = mIndex + 1
  ShowScreen mIndex
End Sub

Private Sub Form_Load()
  mIndex = 0
  ShowScreen mIndex
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

'
'User functions
'
Private Function ShowScreen(ByVal pIndex As Integer)
'
'Show the screen depenceupon the next/back clicked
'
'
'Status(Errors) when submit harddiskkey for registration:
'1 - Invalid user
'2 - Not Purchased
'3 - User Deleted (invalid user)
'4 - On hold (User is locked)
'5 - User is not verified
'keyok - successfully finished
'
'
Dim li As Integer
Dim lQuery As String
Dim lRegistration As String
  For li = picScreens.LBound To picScreens.UBound
    picScreens(li).Left = -20000
  Next
  picScreens(pIndex).Left = imgScreen.Left + imgScreen.Width
  'Button settings
  cmdFinish.Enabled = False
  cmdBack.Enabled = False
  cmdNext.Enabled = True
  If pIndex > 0 Then
    cmdBack.Enabled = True
  End If
  If pIndex = picScreens.UBound Then
    cmdNext.Enabled = False
    cmdFinish.Enabled = True
    cmdFinish.SetFocus
  End If
  On Error Resume Next
  Select Case pIndex
  Case 0
    'optType(1).SetFocus
  Case 1
    lQuery = IIf(optType(1).Value, "1", "0")
    txtEmail.SetFocus
  Case 2
    txtHarddiskkey.Text = GetHardDiskKey
    cmdFinish.Enabled = False
    lQuery = "email=" & txtEmail.Text & "&pass=" & txtPassword.Text & "&harddiskkey=" & txtHarddiskkey.Text & "&computername=" & GetComputerName
    If optType(1).Value Then
      For li = 0 To 4
        txtKey(li).Enabled = False
        txtKey(li).Text = ""
      Next
      lblGetKey.Visible = False
      lblMain.Caption = "Registration on progress"
      lblMainStatus.Caption = "Registration contacts with admin for key"
      'lRegistration = iNet.OpenURL("http://raga/rbg/aspeditor/web/onlinereg/regline.asp?" & lQuery) 'Raga
      lRegistration = iNet.OpenURL("http://www.codepiler.com/onlinereg/regline.asp?" & lQuery)
      Select Case lRegistration
      Case "1", "3"
        lblMain.Caption = "Error on Registration"
        lblMainStatus.Caption = "Invalid Details." & vbCrLf & "Please input the registeration details of Codepiler.com"
        lblMainStatus.ForeColor = vbRed
      Case "2"
        lblMain.Caption = "Error on Registration"
        lblMainStatus.Caption = "Please purchase Codepiler before registering!"
        lblGetKey.Caption = "www.codepiler.com"
        lblGetKey.Visible = True
        lblMainStatus.ForeColor = vbRed
      Case "4"
        lblMain.Caption = "Error on Registration"
        lblMainStatus.Caption = "You are currently locked. Contact admin@codepiler.com"
        lblMainStatus.ForeColor = vbRed
      Case "5"
        lblMain.Caption = "Error on Registration"
        lblMainStatus.Caption = "Invalid User."
        lblMainStatus.ForeColor = vbRed
      Case Else
        If InStr(lRegistration, "keyok") > 0 Then
          cmdFinish.Enabled = True
          lRegistration = Replace(lRegistration, "keyok", "")
          lblMain.Caption = "Registration completed"
          lblMainStatus.Caption = "Successfully completed. Click finish to register."
          lblMainStatus.ForeColor = vbBlack
          For li = 0 To 4
            txtKey(li).Enabled = True
            txtKey(li).Text = Split(lRegistration, "-")(li)
          Next
        Else
          lblMain.Caption = "Error on Registration"
          lblMainStatus.Caption = "Unable to connect CodePiler.com!"
          lblMainStatus.ForeColor = vbRed
        End If
      End Select
    Else
      lblGetKey.Caption = "admin@codepiler.com"
      lblGetKey.Visible = True
      lblMain.Caption = "Registration on progress"
      lblMainStatus.Caption = "Email the below serialno to admin@codepiler.com and" & vbCrLf & "get your Registeration Key in Email"
      lblMainStatus.ForeColor = vbBlack
      For li = 0 To 4
        txtKey(li).Enabled = True
        txtKey(li).Text = ""
      Next
      txtKey(0).SetFocus
    End If
  End Select
End Function

Private Sub InitiateSettings()
'
'Initialise some settings
'
    GMonth(1) = "L"
    GMonth(2) = "N"
    GMonth(3) = "I"
    GMonth(4) = "O"
    GMonth(5) = "F"
    GMonth(6) = "T"
    GMonth(7) = "Q"
    GMonth(8) = "X"
    GMonth(9) = "A"
    GMonth(10) = "B"
    GMonth(11) = "R"
    GMonth(12) = "K"
        
    GYear(1) = "E"
    GYear(2) = "D"
    GYear(3) = "S"
    GYear(4) = "V"
    GYear(5) = "B"
    GYear(6) = "T"
    GYear(7) = "G"
    GYear(8) = "X"
    GYear(9) = "M"
    GYear(10) = "W"
    GYear(11) = "Z"
    GYear(12) = "Q"
    GYear(13) = "U"
    GYear(14) = "F"
    GYear(15) = "N"
    GYear(16) = "P"
    GYear(17) = "R"
    GYear(18) = "I"
    GYear(19) = "H"
    GYear(20) = "A"
    GYear(21) = "O"
    GYear(22) = "L"
    
    'Duration
    GDuration(1) = "U"
    GDuration(2) = "D"
    GDuration(3) = "S"
    GDuration(4) = "R"
    GDuration(5) = "B"
    GDuration(6) = "N"
    GDuration(7) = "G"
    GDuration(8) = "Z"
    GDuration(9) = "M"
    GDuration(10) = "W"
    GDuration(11) = "X"
    GDuration(12) = "O"
    GDuration(13) = "E"
    GDuration(14) = "F"
    GDuration(15) = "C"
    GDuration(16) = "P"
    GDuration(17) = "H"
    GDuration(18) = "I"
    GDuration(19) = "V"
    GDuration(20) = "A"
    GDuration(21) = "Q"
    GDuration(22) = "T"
    GDuration(23) = "K"
    GDuration(24) = "J"
End Sub

Private Sub lblGetKey_Click()
  If optType(1) Then
    OpenWebsite "www.codepiler.com"
  Else
    OpenWebsite "mailto:admin@yahoo.com?subject=Register codepiler&body=Email: " & txtEmail.Text & "; Password: " & txtPassword.Text & "; Harddisk serialno: " & txtHarddiskkey.Text & ". Hereby i send the details and i request you to send the registration key."
  End If
End Sub

Private Sub txtKey_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If cmdFinish.Enabled = False Then cmdFinish.Enabled = True
End Sub

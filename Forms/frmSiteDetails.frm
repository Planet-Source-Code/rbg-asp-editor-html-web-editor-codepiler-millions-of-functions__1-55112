VERSION 5.00
Begin VB.Form frmSiteDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Site Details"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
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
   Icon            =   "frmSiteDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      Height          =   315
      Left            =   5415
      Picture         =   "frmSiteDetails.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Test connection"
      Top             =   4290
      Width           =   330
   End
   Begin VB.CommandButton cmdConStr 
      Height          =   315
      Left            =   5055
      Picture         =   "frmSiteDetails.frx":027F
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Open connection string wizard"
      Top             =   4290
      Width           =   330
   End
   Begin VB.Frame frmLocal 
      Height          =   2340
      Left            =   705
      TabIndex        =   17
      Top             =   1185
      Width           =   5040
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   180
         ScaleHeight     =   255
         ScaleWidth      =   810
         TabIndex        =   24
         Top             =   1140
         Width           =   810
         Begin VB.OptionButton optRemote 
            Caption         =   "Remote"
            Height          =   195
            Left            =   -30
            TabIndex        =   25
            Top             =   60
            Width           =   870
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   180
         ScaleHeight     =   225
         ScaleWidth      =   750
         TabIndex        =   22
         Top             =   270
         Width           =   750
         Begin VB.OptionButton optLocal 
            Caption         =   "Local"
            Height          =   195
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.TextBox txtPort 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4425
         TabIndex        =   6
         Text            =   "21"
         Top             =   1800
         Width           =   390
      End
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3015
         TabIndex        =   5
         Top             =   1785
         Width           =   1125
      End
      Begin VB.TextBox txtUsername 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   1800
         Width           =   1260
      End
      Begin VB.TextBox txtServer 
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   1785
         Width           =   1095
      End
      Begin VB.TextBox txtFolder 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   615
         Width           =   4275
      End
      Begin VB.CommandButton cmdBrowse 
         Height          =   315
         Left            =   4485
         Picture         =   "frmSiteDetails.frx":04D5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   615
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Left            =   4425
         TabIndex        =   21
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   3015
         TabIndex        =   20
         Top             =   1545
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   195
         Left            =   1500
         TabIndex        =   19
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1515
         Width           =   540
      End
   End
   Begin VB.TextBox txtConnectionString 
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   4290
      Width           =   3345
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4830
      TabIndex        =   12
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   3765
      TabIndex        =   11
      Top             =   4875
      Width           =   915
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Text            =   "http://"
      Top             =   3786
      Width           =   4035
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1695
      TabIndex        =   0
      Top             =   720
      Width           =   2430
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection:"
      Height          =   195
      Left            =   705
      TabIndex        =   16
      Top             =   4350
      Width           =   870
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   285
      Picture         =   "frmSiteDetails.frx":0706
      Top             =   105
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URL:"
      Height          =   195
      Left            =   705
      TabIndex        =   15
      Top             =   3840
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   705
      TabIndex        =   14
      Top             =   780
      Width           =   465
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   8130
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   270
      X2              =   8190
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site info."
      Height          =   195
      Left            =   660
      TabIndex        =   13
      Top             =   135
      Width           =   645
   End
End
Attribute VB_Name = "frmSiteDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSitename As String
Private mError As String
Private WithEvents mCSFrm As frmConnectionString
Attribute mCSFrm.VB_VarHelpID = -1
Public Event SiteSaved(ByVal pInfo As String, ByVal pNew As Boolean)

Private Sub cmdBrowse_Click()
Dim lFolder As String
  lFolder = BrowseForFolder(Me.hwnd, "Select local path")
  If lFolder <> "" Then
    txtFolder.Text = lFolder
  End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdConStr_Click()
  Set mCSFrm = New frmConnectionString
  mCSFrm.Show vbModal
End Sub

Private Sub cmdOk_Click()
  If Trim(txtName.Text) <> "" Then
    SaveSite
    Unload Me
  Else
    MsgBox "Sitename should be wanted!", vbInformation, Mtitle
    txtName.SetFocus
  End If
End Sub

Rem ============================
Rem User Functions
Rem ============================
Private Sub SaveSite()
'
'Save the site details
'format: Name^path^url^connectionstring
'
Dim fn As Integer
Dim lContent As String
Dim lTmp As String
Dim lSite As clsSite
  If S102_File_Exists(App.Path & "\Sites.dat") Then
    fn = FreeFile
    Open App.Path & "\Sites.dat" For Input As fn
      lContent = Input(LOF(fn), fn)
    Close fn
  End If
  lTmp = txtName.Text & "^" & txtFolder.Text & "^" & txtUrl.Text & IIf(txtConnectionString <> "", "^" & txtName.Text & "~" & txtConnectionString.Text, "")
  If mSitename = "" Then 'New site
    Set lSite = Msitedetails.Item(txtName.Text)
    If Not lSite Is Nothing Then
      If MsgBox("Site already exists. Do you want to replace?", vbQuestion + vbOKCancel, Mtitle) = vbCancel Then
        Exit Sub
      End If
    End If
    Set lSite = Msitedetails.Newsite
    With lSite
      .Name = txtName.Text
      .LocalPath = txtFolder.Text
      .URL = txtUrl.Text
      .ConnectionString.Add txtName.Text & "~" & txtConnectionString.Text
      .UseFTP = optRemote.Value
      .Server = txtServer.Text
      .Username = txtUsername.Text
      .Password = txtPassword.Text
      .Port = val(txtPort.Text)
    End With
    Msitedetails.Add lSite
    Msitedetails.Save
    RaiseEvent SiteSaved(txtName.Text, True)
  Else 'Edit site
    Set lSite = Msitedetails.Item(mSitename)
    If lSite Is Nothing Then
      Set lSite = Msitedetails.Newsite
    Else
      Msitedetails.Remove mSitename
    End If
    With lSite
      .Name = txtName.Text
      .LocalPath = txtFolder.Text
      .URL = txtUrl.Text
      .ConnectionString.Add txtName.Text & "~" & txtConnectionString.Text
      .UseFTP = optRemote.Value
      .Server = txtServer.Text
      .Username = txtUsername.Text
      .Password = txtPassword.Text
      .Port = val(txtPort.Text)
    End With
    Msitedetails.Add lSite
    Msitedetails.Save
    RaiseEvent SiteSaved(txtName.Text, False)
    '---updating RTB MlocalHost property---------
    If MoldLocalHost <> txtUrl Then
      frmEditor.UpdateSiteLocalHost MoldLocalHost, txtUrl
    End If
  End If
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  mSitename = ""
End Sub

Private Sub mCSFrm_ConnectionStringOk(ByVal ConnectionString As String)
  txtConnectionString.Text = ConnectionString
End Sub

Private Sub cmdTest_Click()
  Screen.MousePointer = vbHourglass
  If TestConnection(txtConnectionString.Text) = False Then
    MsgBox mError, vbCritical, Mtitle
  Else
    MsgBox "Connection Succeeded.", vbInformation, Mtitle
  End If
  Screen.MousePointer = vbDefault
End Sub


'
'User function
'
Private Function TestConnection(ByVal pConnectionString As String) As Boolean
'
'Test the connection for valid
'
Dim lCon As Object
  Err.Clear
  mError = ""
  On Error GoTo Cerr
  Set lCon = CreateObject("ADODB.Connection")
  TestConnection = True
  Call lCon.Open(pConnectionString)
  If lCon.State = 1 Then 'Open state
    TestConnection = True
  Else
    TestConnection = False
  End If
  Exit Function
Cerr:
  mError = Err.Description
  TestConnection = False
End Function

Public Sub LoadDetails(Optional ByVal Sitename As String)
Dim lSite As clsSite
Dim lConn As String
  mSitename = Sitename
  Set lSite = Msitedetails.Item(Sitename)
  If Not lSite Is Nothing Then
    With lSite
      txtName.Text = .Name
      txtFolder.Text = .LocalPath
      txtUrl.Text = .URL
      MoldLocalHost = .URL
      txtServer.Text = .Server
      txtUsername.Text = .Username
      txtPassword.Text = .Password
      txtPort.Text = IIf(.Port = 0, 21, .Port)
      If .UseFTP Then
        optRemote.Value = True
        optLocal.Value = False
      Else
        optLocal.Value = True
        optRemote.Value = False
      End If
      If Not lSite.ConnectionString Is Nothing Then
        If lSite.ConnectionString.Count > 0 Then
          lConn = lSite.ConnectionString.Item(1)
          txtConnectionString.Text = Split(lConn, "~")(1)
          txtConnectionString.Tag = Split(lConn, "~")(0)
        Else
          txtConnectionString.Text = ""
          txtConnectionString.Tag = ""
        End If
      Else
        txtConnectionString.Text = ""
        txtConnectionString.Tag = ""
      End If
    End With
    'SetRemoteSetting
  Else
    txtName.Text = "New Site 1"
    optLocal.Value = True
    
  End If
End Sub

Private Sub optLocal_Click()
  On Error GoTo Cerr
  'Enable
  txtFolder.Enabled = True
  cmdBrowse.Enabled = True
  txtServer.Enabled = False
  txtUsername.Enabled = False
  txtPassword.Enabled = False
  txtPort.Enabled = False
  optRemote.Value = False
  'Coloring
  txtFolder.BackColor = vbWhite
  txtServer.BackColor = vbButtonFace
  txtUsername.BackColor = vbButtonFace
  txtPassword.BackColor = vbButtonFace
  txtPort.BackColor = vbButtonFace
  txtFolder.SetFocus
Cerr:
End Sub

Private Sub optRemote_Click()
  On Error GoTo Cerr
  'Enable
  txtFolder.Enabled = False
  cmdBrowse.Enabled = False
  txtServer.Enabled = True
  txtUsername.Enabled = True
  txtPassword.Enabled = True
  txtPort.Enabled = True
  optLocal.Value = False
  'Coloring
  txtFolder.BackColor = vbButtonFace
  txtServer.BackColor = vbWhite
  txtUsername.BackColor = vbWhite
  txtPassword.BackColor = vbWhite
  txtPort.BackColor = vbWhite
  txtServer.SetFocus
Cerr:
End Sub

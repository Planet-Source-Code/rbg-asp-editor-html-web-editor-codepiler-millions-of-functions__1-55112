VERSION 5.00
Begin VB.Form frmConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connection"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
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
   ScaleHeight     =   2400
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTestConnection 
      Height          =   315
      Left            =   5730
      Picture         =   "frmConnection.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Test connection"
      Top             =   1335
      Width           =   330
   End
   Begin VB.CommandButton cmdConStr 
      Height          =   315
      Left            =   5355
      Picture         =   "frmConnection.frx":0273
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1335
      Width           =   330
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   5055
      TabIndex        =   4
      Top             =   1860
      Width           =   1005
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   3915
      TabIndex        =   3
      Top             =   1860
      Width           =   1005
   End
   Begin VB.TextBox txtConnectionString 
      Height          =   300
      Left            =   2310
      TabIndex        =   1
      Top             =   1350
      Width           =   3000
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   2310
      TabIndex        =   0
      Top             =   855
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection String:"
      Height          =   195
      Left            =   855
      TabIndex        =   7
      Top             =   1410
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   855
      TabIndex        =   6
      Top             =   915
      Width           =   465
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   300
      X2              =   8015
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   345
      X2              =   8015
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection"
      Height          =   195
      Left            =   780
      TabIndex        =   5
      Top             =   210
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   450
      Picture         =   "frmConnection.frx":04C9
      Top             =   165
      Width           =   240
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Provider=MSDASQL.1;Extended Properties="DBQ=Z:\Aruna\Data\NursingHome.mdb;DefaultDir=Z:\Aruna\Data;Driver={Microsoft Access Driver (*.mdb)};DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=srimedimax;UID=admin;"
Private mError As String
Public mSitename As String
Public mEdit As Boolean
Private WithEvents mCSFrm As frmConnectionString
Attribute mCSFrm.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdConStr_Click()
  Set mCSFrm = New frmConnectionString
  mCSFrm.Show vbModal
End Sub

Private Sub cmdOk_Click()
Dim lNode As Object
Dim lKey As String
  On Error Resume Next
  If txtName.Text = "" Then
    MsgBox "Connection name should be wanted!", vbInformation, Mtitle
    txtName.SetFocus
    Exit Sub
  End If

  If ConnectionExistence(txtName.Text) And mEdit = False Then
    MsgBox "Connection already exists! Rename the connection.", vbCritical, Mtitle
    txtName.SetFocus
    Exit Sub
  End If
  lKey = txtName.Tag
  If mEdit = False Then
    frmEditor.tvApplication.Nodes.Add(, , txtName.Text, txtName.Text, "CONNECTION").Tag = txtConnectionString.Text  'Ucase Changed
  Else
    frmEditor.tvApplication.Nodes(lKey).Tag = txtConnectionString.Text 'Ucase Changed'Ucase Changed
    frmEditor.tvApplication.Nodes(lKey).Text = txtName.Text 'Ucase Changed
    frmEditor.tvApplication.Nodes(lKey).Key = txtName.Text ''Ucase Changed
    frmEditor.tvApplication.Nodes.Remove "T" & lKey ''Ucase Changed
    frmEditor.tvApplication.Nodes.Remove "V" & lKey ''Ucase Changed
  End If
  SaveDBDetails
  frmEditor.LoadConnection txtName.Text, txtConnectionString.Text
  Unload Me
End Sub

Private Sub cmdTestConnection_Click()
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

Private Function ConnectionExistence(ByVal pConnectionName As String) As Boolean
'
'Is connection exist
'
Dim lNode As Object
  On Error GoTo Cerr
  'pConnectionName = pConnectionName 'Ucase Changed
  Set lNode = frmEditor.tvApplication.Nodes(pConnectionName)
  ConnectionExistence = Not (lNode Is Nothing)
  Exit Function
Cerr:
  ConnectionExistence = False
End Function

Private Sub Form_Unload(Cancel As Integer)
  mEdit = False
End Sub

Private Sub mCSFrm_ConnectionStringOk(ByVal ConnectionString As String)
  txtConnectionString.Text = ConnectionString
End Sub

Rem ============================
Rem User Functions
Rem ============================

Private Sub SaveDBDetails()
'
'Save the connection details
'
Dim lSite As clsSite
Dim lConn As String
  Set lSite = Msitedetails.Item(mSitename)
  If Not lSite Is Nothing Then
    lConn = txtName.Text & "~" & txtConnectionString.Text
    lSite.ConnectionString.Add lConn
    Msitedetails.Remove lSite.Name
    Msitedetails.Add lSite
    Msitedetails.Save
  End If
End Sub


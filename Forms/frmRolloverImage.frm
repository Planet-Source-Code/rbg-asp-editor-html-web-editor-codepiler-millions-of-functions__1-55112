VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRolloverImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rollover Image"
   ClientHeight    =   3270
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
   Icon            =   "frmRolloverImage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowseU 
      Height          =   315
      Left            =   5535
      Picture         =   "frmRolloverImage.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2085
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowseR 
      Height          =   315
      Left            =   5535
      Picture         =   "frmRolloverImage.frx":023D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1635
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowseO 
      Height          =   315
      Left            =   5535
      Picture         =   "frmRolloverImage.frx":046E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1170
      Width           =   330
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4965
      TabIndex        =   8
      Top             =   2685
      Width           =   900
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   3915
      TabIndex        =   7
      Top             =   2685
      Width           =   900
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4020
      Top             =   645
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   1965
      TabIndex        =   5
      Top             =   2085
      Width           =   3555
   End
   Begin VB.TextBox txtRollover 
      Height          =   315
      Left            =   1965
      TabIndex        =   3
      Top             =   1635
      Width           =   3555
   End
   Begin VB.TextBox txtOriginal 
      Height          =   315
      Left            =   1965
      TabIndex        =   1
      Top             =   1170
      Width           =   3555
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1965
      TabIndex        =   0
      Top             =   720
      Width           =   1605
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   330
      Picture         =   "frmRolloverImage.frx":069F
      Top             =   195
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rollover Image"
      Height          =   195
      Left            =   675
      TabIndex        =   13
      Top             =   165
      Width           =   1080
   End
   Begin VB.Line Line1 
      X1              =   330
      X2              =   8000
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   285
      X2              =   8000
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Href URL:"
      Height          =   195
      Left            =   675
      TabIndex        =   12
      Top             =   2145
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rollover Image:"
      Height          =   195
      Left            =   675
      TabIndex        =   11
      Top             =   1695
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Original Image:"
      Height          =   195
      Left            =   675
      TabIndex        =   10
      Top             =   1230
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image Name:"
      Height          =   195
      Left            =   675
      TabIndex        =   9
      Top             =   780
      Width           =   960
   End
End
Attribute VB_Name = "frmRolloverImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim lSwapimage As String
Dim lSwapRestore As String
Dim lImage As String
Dim lStr As String
Dim lScript As String
Dim lCloseScript As String
Dim lPos As Long
  lSwapimage = vbTab & "function CP_imageSwap() { " & vbCrLf & _
                vbTab & vbTab & "var i,j=0,imgT,a=CP_imageSwap.arguments; document.CP_img=new Array; for(i=0;i<(a.length-2);i+=3)" & vbCrLf & _
                vbTab & vbTab & "if ((imgT=document.all(a[i]))!=null){document.CP_img[j++]=imgT; if(!imgT.oSrc) imgT.oSrc=imgT.src; imgT.src=a[i+2];}" & vbCrLf & _
                vbTab & "}" & vbCrLf & vbCrLf
  lSwapRestore = vbTab & "function CP_imageRestore() {" & vbCrLf & _
                  vbTab & vbTab & "var i,imgT,a=document.CP_img; for(i=0;a&&i<a.length&&(imgT=a[i])&&imgT.oSrc;i++) imgT.src=imgT.oSrc;" & vbCrLf & _
                  vbTab & "}" & vbCrLf
  lScript = vbCrLf & "<script language=""JavaScript"">" & vbCrLf & vbCrLf & vbTab & "<!-- Code Piler Generated" & vbCrLf
  lCloseScript = vbTab & "-->" & vbCrLf & "</script>" & vbCrLf
  If Trim(txtName.Text) <> "" Then
    If Trim(txtOriginal.Text) <> "" And Trim(txtRollover.Text) <> "" Then
      lImage = vbCrLf & vbTab & "<a href=""" & IIf(Trim(txtUrl.Text) <> "", Trim(txtUrl.Text), "#") & """ onMouseOut=""CP_imageRestore()"" onMouseOver=""CP_imageSwap('" & txtName.Text & "','','" & txtRollover.Text & "',1)""><img name=""" & txtName.Text & """ border=""0"" src=""" & txtOriginal.Text & """></a> "
      lStr = frmEditor.RTB(frmEditor.GetActiveRTB).Text
      'Insert the image tag
      lPos = InStr(1, LCase(lStr), "<body") + 1
      If lPos > 0 Then
        frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(lPos, LCase(lStr), ">") + 1
        frmEditor.RTB(frmEditor.GetActiveRTB).Paste lImage
      ElseIf InStr(1, LCase(lStr), "<html>") > 0 Then
        frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, LCase(lStr), "<html>") + 6
        frmEditor.RTB(frmEditor.GetActiveRTB).Paste lImage
      Else
        frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = Len(lStr)
        frmEditor.RTB(frmEditor.GetActiveRTB).Paste lImage
      End If
      'find the position to insert the script
      If InStr(lStr, "function cp_imageswap()") = 0 Then
        If InStr(1, LCase(lStr), "<!-- code piler generated") > 0 Then
          frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, LCase(lStr), "<!-- code piler generated") + 25
          frmEditor.RTB(frmEditor.GetActiveRTB).Paste lSwapimage & lSwapRestore
        ElseIf InStr(1, LCase(lStr), "<head>") > 0 Then
          frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, LCase(lStr), "<head>") + 6
          frmEditor.RTB(frmEditor.GetActiveRTB).Paste lScript & lSwapimage & lSwapRestore & lCloseScript
        ElseIf InStr(1, LCase(lStr), "<html>") > 0 Then
          frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, LCase(lStr), "<html>") + 6
          frmEditor.RTB(frmEditor.GetActiveRTB).Paste lScript & lSwapimage & lSwapRestore & lCloseScript
        Else
          frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = 1
          frmEditor.RTB(frmEditor.GetActiveRTB).Paste lScript & lSwapimage & lSwapRestore & lCloseScript
        End If
      End If
      
      Unload Me
    Else
      If Trim(txtOriginal.Text) = "" Then
        MsgBox "Original image must required.", vbInformation + vbOKOnly, Mtitle
        txtOriginal.SetFocus
      ElseIf Trim(txtRollover.Text) = "" Then
        MsgBox "Rollover image must required.", vbInformation + vbOKOnly, Mtitle
        txtRollover.SetFocus
      End If
    End If
  Else
    MsgBox "Name of the image must required.", vbInformation + vbOKOnly, Mtitle
    txtName.SetFocus
  End If
End Sub

Private Sub cmdBrowseO_Click()
Dim lName As String
  On Error GoTo Cerr
  CD.FileName = ""
  CD.Filter = "All Picture Files(*.gif,*.jpg)|*.gif;*.jpg|All files(*.*)|*.*"
  CD.CancelError = True
  CD.ShowOpen
  lName = CD.FileName
  lName = "file:///" & Replace(lName, "\", "/")
  txtOriginal.Text = lName
Cerr:
End Sub

Private Sub cmdBrowseR_Click()
Dim lName As String
  On Error GoTo Cerr
  CD.FileName = ""
  CD.Filter = "All Picture Files(*.gif,*.jpg)|*.gif;*.jpg|All files(*.*)|*.*"
  CD.CancelError = True
  CD.ShowOpen
  lName = CD.FileName
  lName = "file:///" & Replace(lName, "\", "/")
  txtRollover.Text = lName
Cerr:
End Sub

Private Sub cmdBrowseU_Click()
  On Error GoTo Cerr
  CD.FileName = ""
  CD.Filter = "All Picture Files(*.gif,*.jpg)|*.gif;*.jpg|All files(*.*)|*.*"
  CD.CancelError = True
  CD.ShowOpen
  txtUrl.Text = CD.FileName
Cerr:
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{55E69722-DD29-4623-A059-5B96E8A9018D}#1.2#0"; "ColorPicker.ocx"
Begin VB.Form frmTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
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
   Icon            =   "frmTable.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkASPCode 
      Caption         =   "Make as ASP Code"
      Height          =   270
      Left            =   4155
      TabIndex        =   13
      Top             =   2010
      Width           =   1665
   End
   Begin VB.PictureBox picPaleteB 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5550
      ScaleHeight     =   255
      ScaleWidth      =   270
      TabIndex        =   33
      Top             =   1560
      Width           =   270
      Begin VB.Image imgPaleteB 
         Height          =   255
         Left            =   0
         Picture         =   "frmTable.frx":058A
         Top             =   0
         Width           =   270
      End
   End
   Begin prjColorPicker.ColorPicker ColorPicker 
      Height          =   2190
      Left            =   6060
      TabIndex        =   16
      Top             =   1800
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   3863
   End
   Begin VB.OptionButton optSize 
      Caption         =   "In pixels"
      Height          =   255
      Index           =   0
      Left            =   3150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3420
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optSize 
      Caption         =   "In percent"
      Height          =   255
      Index           =   1
      Left            =   4230
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3420
      Width           =   1050
   End
   Begin VB.PictureBox picPaleteF 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5565
      ScaleHeight     =   255
      ScaleWidth      =   270
      TabIndex        =   32
      Top             =   1110
      Width           =   270
      Begin VB.Image imgPaleteF 
         Height          =   255
         Left            =   0
         Picture         =   "frmTable.frx":0612
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.ComboBox cboAlignment 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "center"
      Top             =   1515
      Width           =   1215
   End
   Begin VB.TextBox txtBorderSize 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Text            =   "0"
      Top             =   1965
      Width           =   960
   End
   Begin VB.TextBox txtPadding 
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Text            =   "1"
      Top             =   2385
      Width           =   960
   End
   Begin VB.TextBox txtSpacing 
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Text            =   "1"
      Top             =   2820
      Width           =   960
   End
   Begin VB.TextBox txtCol 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Text            =   "1"
      Top             =   1065
      Width           =   960
   End
   Begin VB.TextBox txtRow 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Text            =   "0"
      Top             =   630
      Width           =   960
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2115
      TabIndex        =   7
      Text            =   "100"
      Top             =   3405
      Width           =   975
   End
   Begin VB.CheckBox chkSpaceWidth 
      Caption         =   "Space Width"
      Height          =   255
      Left            =   735
      TabIndex        =   6
      Top             =   3405
      Width           =   1200
   End
   Begin VB.TextBox txtColor1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4110
      TabIndex        =   11
      Text            =   "#000000"
      Top             =   1065
      Width           =   1425
   End
   Begin VB.CheckBox chkAlternative 
      Caption         =   "Alternative Color"
      Height          =   255
      Left            =   3510
      TabIndex        =   10
      Top             =   630
      Width           =   1830
   End
   Begin VB.TextBox txtColor2 
      Height          =   300
      Left            =   4095
      TabIndex        =   12
      Text            =   "#000000"
      Top             =   1515
      Width           =   1425
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6360
      TabIndex        =   15
      Top             =   1050
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   6360
      TabIndex        =   14
      Top             =   630
      Width           =   945
   End
   Begin MSComCtl2.UpDown udSpacing 
      Height          =   300
      Left            =   2761
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2820
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtSpacing"
      BuddyDispid     =   196617
      OrigLeft        =   3015
      OrigTop         =   2820
      OrigRight       =   3270
      OrigBottom      =   3120
      Max             =   99
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udPadding 
      Height          =   300
      Left            =   2761
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2385
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtPadding"
      BuddyDispid     =   196616
      OrigLeft        =   3015
      OrigTop         =   2385
      OrigRight       =   3270
      OrigBottom      =   2685
      Max             =   99
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udBorder 
      Height          =   300
      Left            =   2745
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1965
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtBorderSize"
      BuddyDispid     =   196615
      OrigLeft        =   3015
      OrigTop         =   1965
      OrigRight       =   3270
      OrigBottom      =   2265
      Max             =   99
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udRow 
      Height          =   300
      Left            =   2761
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   630
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtRow"
      BuddyDispid     =   196619
      OrigLeft        =   3015
      OrigTop         =   630
      OrigRight       =   3270
      OrigBottom      =   930
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udCol 
      Height          =   300
      Left            =   2761
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1065
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtCol"
      BuddyDispid     =   196618
      OrigLeft        =   3015
      OrigTop         =   1065
      OrigRight       =   3270
      OrigBottom      =   1365
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   255
      Picture         =   "frmTable.frx":069A
      Top             =   135
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alignment:"
      Height          =   195
      Left            =   735
      TabIndex        =   31
      Top             =   1575
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Border Size:"
      Height          =   195
      Left            =   735
      TabIndex        =   30
      Top             =   2010
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cell Padding:"
      Height          =   195
      Left            =   735
      TabIndex        =   29
      Top             =   2445
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cell Spacing:"
      Height          =   195
      Left            =   735
      TabIndex        =   28
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Columns:"
      Height          =   195
      Left            =   735
      TabIndex        =   27
      Top             =   1125
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rows:"
      Height          =   195
      Left            =   735
      TabIndex        =   26
      Top             =   690
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color1:"
      Height          =   195
      Left            =   3510
      TabIndex        =   25
      Top             =   1110
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color2:"
      Height          =   195
      Left            =   3510
      TabIndex        =   24
      Top             =   1575
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table"
      Height          =   195
      Left            =   600
      TabIndex        =   18
      Top             =   135
      Width           =   390
   End
   Begin VB.Line Line1 
      X1              =   270
      X2              =   7940
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   225
      X2              =   7940
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Label lblTable 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   17
      Top             =   6900
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtcolor1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lColor As Long
  lColor = ColorPicker.GetLongRGB(txtColor1.Text)
  picPaleteF.BackColor = IIf(lColor <> -1, lColor, vbButtonFace)
End Sub

Private Sub chkAlternative_Click()
  If chkAlternative.Value = 1 Then
    txtPadding.Text = "0"
    txtSpacing.Text = "0"
  End If
  txtColor1.Enabled = chkAlternative.Value
  txtColor2.Enabled = chkAlternative.Value
  picPaleteB.Enabled = chkAlternative.Value
  picPaleteF.Enabled = chkAlternative.Value
End Sub

Private Sub imgPaleteF_Click()
  On Error Resume Next
  ColorPicker.Tag = 1
  ColorPicker.Top = picPaleteF.Top + picPaleteF.Height
  ColorPicker.Left = picPaleteF.Left - (ColorPicker.Width - picPaleteF.Width)
  If Trim(txtColor1.Text) <> "" Then ColorPicker.ColorWebRGB = txtColor1.Text
  ColorPicker.ShowColor = Not ColorPicker.ShowColor
End Sub

Private Sub txtcolor2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lColor As Long
  lColor = ColorPicker.GetLongRGB(txtColor2.Text)
  picPaleteB.BackColor = IIf(lColor <> -1, lColor, vbButtonFace)
End Sub

Private Sub imgPaleteB_Click()
  On Error Resume Next
  ColorPicker.Tag = 2
  ColorPicker.Top = picPaleteB.Top + picPaleteB.Height
  ColorPicker.Left = picPaleteB.Left - (ColorPicker.Width - picPaleteF.Width)
  If Trim(txtColor2.Text) <> "" Then ColorPicker.ColorWebRGB = txtColor2.Text
  ColorPicker.ShowColor = Not ColorPicker.ShowColor
End Sub

Private Sub ColorPicker_ColorSelect(ByVal Color As Long, ByVal WebRGBFormat As String)
  If ColorPicker.Tag = 1 Then
    picPaleteF.BackColor = Color
    txtColor1.Text = WebRGBFormat
  Else
    picPaleteB.BackColor = Color
    txtColor2.Text = WebRGBFormat
  End If
End Sub

Private Sub cmdAdd_Click()
Dim wsizeasp As String
Dim wsize As String
  wsize = ""
  If chkSpaceWidth.Value = True Then
    If optSize(0).Value = True Then
      wsize = " width=" & Text6.Text
      wsizeasp = " width=" & Chr(34) & Chr(34) & Text6.Text & Chr(34) & Chr(34)
    Else
      wsize = " width=" & Text6.Text & "%"
      wsizeasp = " width=" & Chr(34) & Chr(34) & Text6.Text & "%" & Chr(34) & Chr(34)
    End If
  End If
  If chkAspCode.Value = False Then
    Call AddTable(txtRow.Text, txtCol.Text, wsize)
  Else
    Call AddTableAsp(txtRow.Text, txtCol.Text, wsizeasp)
  End If
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  On Error Resume Next
  cboAlignment.AddItem "Left"
  cboAlignment.AddItem "Center"
  cboAlignment.AddItem "Right"
  ColorPicker.ShowColor = False
  chkAlternative_Click
End Sub

Rem ==============================
Rem USER FUNCTIONS
Rem ==============================

Function AddTable(ColumnCount As Long, RowCount As Long, swth As String) As String
Dim tmp
Dim j As Long
Dim k As Long
Dim lColor1 As String
Dim lColor2 As String
Dim lColor As String
Dim quote$
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  quote$ = Chr$(34)
  lColor1 = IIf(chkAlternative.Value = 1 And txtColor1.Text <> "", "bgcolor='" & txtColor1.Text & "'", "")
  lColor2 = IIf(chkAlternative.Value = 1 And txtColor2.Text <> "", "bgcolor='" & txtColor2.Text & "'", "")
  tmp = "<table align=" & Chr(34) & cboAlignment.Text & Chr(34) & " border=" & Chr(34) & txtBorderSize.Text & Chr(34) & " cellpadding=" & Chr(34) & txtPadding.Text & Chr(34) & swth & " cellspacing=" & Chr(34) & txtSpacing.Text & Chr(34) & ">" & vbCrLf
  For j = 1 To RowCount
    tmp = tmp & vbTab & "<tr>" & vbCrLf
    lColor = IIf(j Mod 2 = 0, lColor2, lColor1)
    For k = 1 To ColumnCount
      If lColor = "" Then
        tmp = tmp & vbTab & vbTab & "<td>&nbsp;</td>" & vbCrLf
      Else
        tmp = tmp & vbTab & vbTab & "<td " & lColor & ">&nbsp;</td>" & vbCrLf
      End If
    Next k
    tmp = tmp & vbTab & "</tr>" & vbCrLf
  Next j
  tmp = tmp & "</table>" & vbCrLf
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste tmp
  Screen.MousePointer = vbDefault
End Function

Function AddTableAsp(ColumnCount As Long, RowCount As Long, swth As String) As String
Dim tmp
Dim j As Long
Dim k As Long
Dim lColor1 As String
Dim lColor2 As String
Dim lColor As String
Dim quote$
  On Error Resume Next
  quote$ = Chr$(34)
  lColor1 = IIf(chkAlternative.Value = 1 And txtColor1.Text <> "", "bgcolor='" & txtColor1.Text & "'", "")
  lColor2 = IIf(chkAlternative.Value = 1 And txtColor2.Text <> "", "bgcolor='" & txtColor2.Text & "'", "")
  tmp = "Response.Write" & Chr(34) & "<table align=" & Chr(34) & Chr(34) & cboAlignment.Text & Chr(34) & Chr(34) & " border=" & Chr(34) & Chr(34) & txtBorderSize.Text & Chr(34) & Chr(34) & " cellpadding=" & Chr(34) & Chr(34) & txtPadding.Text & Chr(34) & Chr(34) & swth & " cellspacing=" & Chr(34) & Chr(34) & txtSpacing.Text & Chr(34) & Chr(34) & ">" & Chr(34) & vbCrLf
  For j = 1 To RowCount
    lColor = IIf(j Mod 2 = 0, lColor2, lColor1)
    For k = 1 To ColumnCount
      If lColor = "" Then
        tmp = tmp & "Response.Write" & Chr(34) & "<td>&nbsp;</td>" & Chr(34) & vbCrLf
      Else
        tmp = tmp & "Response.Write" & Chr(34) & "<td " & lColor & ">&nbsp;</td>" & Chr(34) & vbCrLf
      End If
    Next k
    tmp = tmp & "Response.Write" & Chr(34) & "</tr>" & Chr(34)
  Next j
  tmp = tmp & vbCrLf & "Response.Write" & Chr(34) & "</table>" & Chr(34) & vbCrLf
  Clipboard.SetText tmp
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste
End Function

Private Sub Form_Initialize()
  InitXP
End Sub

VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List/Menu"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
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
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optList 
      Caption         =   "List"
      Height          =   300
      Left            =   2295
      TabIndex        =   5
      Top             =   3090
      Width           =   660
   End
   Begin VB.OptionButton optMenu 
      Caption         =   "Menu"
      Height          =   300
      Left            =   1350
      TabIndex        =   4
      Top             =   3090
      Width           =   825
   End
   Begin VB.ListBox lsvItems 
      Height          =   1035
      Left            =   1350
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   1935
      Width           =   2745
   End
   Begin VB.ComboBox cboDirection 
      Height          =   315
      ItemData        =   "frmList.frx":000C
      Left            =   1350
      List            =   "frmList.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1500
      Width           =   1800
   End
   Begin VB.TextBox txtSize 
      Height          =   300
      Left            =   1350
      TabIndex        =   1
      Top             =   1080
      Width           =   630
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1350
      TabIndex        =   0
      Top             =   660
      Width           =   3105
   End
   Begin VB.CommandButton cmdItems 
      Caption         =   "List items"
      Height          =   330
      Left            =   5250
      TabIndex        =   10
      Top             =   1785
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   5250
      TabIndex        =   9
      Top             =   1065
      Width           =   990
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   5250
      TabIndex        =   8
      Top             =   645
      Width           =   990
   End
   Begin VB.CheckBox chkDisabled 
      Caption         =   "Disabled"
      Height          =   270
      Left            =   1350
      TabIndex        =   7
      Top             =   3960
      Width           =   960
   End
   Begin VB.CheckBox chkAllowMultile 
      Caption         =   "Allow multiple selections"
      Height          =   285
      Left            =   1350
      TabIndex        =   6
      Top             =   3555
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   285
      Picture         =   "frmList.frx":0038
      Top             =   105
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   195
      Left            =   540
      TabIndex        =   16
      Top             =   3143
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected:"
      Height          =   195
      Left            =   540
      TabIndex        =   15
      Top             =   1935
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direction:"
      Height          =   195
      Left            =   540
      TabIndex        =   14
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      Height          =   195
      Left            =   540
      TabIndex        =   13
      Top             =   1133
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   540
      TabIndex        =   12
      Top             =   713
      Width           =   465
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   225
      X2              =   6420
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6435
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select"
      Height          =   195
      Left            =   585
      TabIndex        =   11
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdItems_Click()
  frmListItem.Show vbModal
End Sub

Private Sub cmdOk_Click()
Dim lName As String
Dim lDir As String
Dim lSize As String
Dim lDisabled As String
Dim lMultiple As String
Dim lTag As String
Dim lSelect As String
Dim li As Integer
  Screen.MousePointer = vbHourglass
  lTag = ""
  lName = IIf(txtName.Text <> "", " name=""" & txtName.Text & """ ", "")
  lDir = " dir=""" & IIf(cboDirection.ListIndex = 1, "rtl", "ltr") & """ "
  lSize = IIf(optList.Value = True, " size=" & txtSize.Text & " ", "")
  lDisabled = IIf(chkDisabled.Value = 1, " disabled=""disabled"" ", "")
  lMultiple = IIf(chkAllowMultile.Value = 1 And optList.Value = True, " multiple ", "")
  For li = 0 To lsvItems.ListCount - 1
    lTag = lTag & vbTab & "<option value=" & li + 1 & IIf(lsvItems.Selected(li) = True, " selected ", "") & ">" & lsvItems.List(li) & "</option>" & vbCrLf
  Next
  lSelect = vbCrLf & vbTab & "<select" & lName & lSize & lMultiple & lDisabled & lDir & ">" & vbCrLf & lTag & vbTab & "</select>" & vbCrLf
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste lSelect
  Screen.MousePointer = vbDefault
  Unload Me
End Sub


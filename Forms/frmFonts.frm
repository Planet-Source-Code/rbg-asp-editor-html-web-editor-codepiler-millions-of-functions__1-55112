VERSION 5.00
Begin VB.Form frmFonts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fonts List"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
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
   Icon            =   "frmFonts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lsFontslist_b 
      Height          =   645
      ItemData        =   "frmFonts.frx":000C
      Left            =   -1350
      List            =   "frmFonts.frx":0025
      TabIndex        =   12
      Top             =   3435
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5475
      TabIndex        =   6
      Top             =   1260
      Width           =   960
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   5490
      TabIndex        =   5
      Top             =   825
      Width           =   960
   End
   Begin VB.CommandButton cmdRemoveChoose 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2925
      TabIndex        =   4
      Top             =   3525
      Width           =   375
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2925
      TabIndex        =   3
      Top             =   3210
      Width           =   375
   End
   Begin VB.ListBox lsChosen 
      Height          =   1230
      Left            =   3390
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1785
   End
   Begin VB.ListBox lsFonts 
      Height          =   1230
      Left            =   1065
      Sorted          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1785
   End
   Begin VB.ListBox lsFontsList 
      Height          =   1425
      ItemData        =   "frmFonts.frx":0109
      Left            =   1065
      List            =   "frmFonts.frx":0122
      TabIndex        =   0
      Top             =   1125
      Width           =   4110
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4815
      TabIndex        =   2
      Top             =   825
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4425
      TabIndex        =   1
      Top             =   825
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chosen Fonts:"
      Height          =   195
      Left            =   3390
      TabIndex        =   11
      Top             =   2700
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fonts:"
      Height          =   195
      Left            =   1065
      TabIndex        =   10
      Top             =   2700
      Width           =   465
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   420
      X2              =   8135
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   465
      X2              =   8135
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fonts List"
      Height          =   195
      Left            =   975
      TabIndex        =   7
      Top             =   210
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   540
      Picture         =   "frmFonts.frx":0206
      Top             =   135
      Width           =   315
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  lsFontsList.AddItem "(New Fonts List)"
  lsFontsList.ListIndex = lsFontsList.NewIndex
End Sub

Private Sub cmdCancel_Click()
Dim li As Integer
  lsFontsList.Clear
  For li = 0 To lsFontslist_b.ListCount - 1
    lsFontsList.AddItem lsFontslist_b.List(li)
  Next
  lsFonts.ListIndex = 0
  lsFontsList.ListIndex = 0
  lsFontsList_Click
  Me.Hide
End Sub

Private Sub cmdChoose_Click()
Dim li As Integer
Dim lFonts As String
  If lsFonts.ListIndex > -1 Then
    For li = 0 To lsChosen.ListCount - 1
      If lsChosen.List(li) = lsFonts.List(lsFonts.ListIndex) Then Exit Sub
    Next
    lsChosen.AddItem lsFonts.List(lsFonts.ListIndex)
    lFonts = ""
    For li = 0 To lsChosen.ListCount - 1
      lFonts = lFonts & lsChosen.List(li) & ", "
    Next
    lFonts = Trim(lFonts)
    If Right(lFonts, 1) = "," Then lFonts = Left(lFonts, Len(lFonts) - 1)
    lsFontsList.List(lsFontsList.ListIndex) = lFonts
  End If
End Sub

Private Sub cmdOk_Click()
Dim li As Integer
  Screen.MousePointer = vbHourglass
  lsFontslist_b.Tag = lsFontslist_b.ListCount
  lsFontslist_b.Clear
  For li = 0 To lsFontsList.ListCount - 1
    lsFontslist_b.AddItem lsFontsList.List(li)
  Next
  On Error GoTo Cnext
  If lsFontsList.ListIndex > -1 Then
    frmCSSEditor.cFontFamily.Text = lsFontsList.List(lsFontsList.ListIndex)
  End If
Cnext:
  lsFonts.ListIndex = 0
  If lsFontsList.ListCount > 0 Then
    lsFontsList.ListIndex = 0
    lsFontsList_Click
  End If
  frmEditor.LoadFontsMenu
  Screen.MousePointer = vbDefault
  Me.Hide
End Sub

Private Sub cmdRemove_Click()
  If lsFontsList.ListIndex > -1 Then
    lsFontsList.RemoveItem lsFontsList.ListIndex
  End If
End Sub

Private Sub cmdRemoveChoose_Click()
Dim lFonts As String
Dim li As Integer
  If lsChosen.ListIndex > -1 Then
    lsChosen.RemoveItem lsChosen.ListIndex
    lFonts = ""
    For li = 0 To lsChosen.ListCount - 1
      lFonts = lFonts & lsChosen.List(li) & ", "
    Next
    lFonts = Trim(lFonts)
    If Right(lFonts, 1) = "," Then lFonts = Left(lFonts, Len(lFonts) - 1)
    lsFontsList.List(lsFontsList.ListIndex) = lFonts
  End If
End Sub

Private Sub Form_Load()
Dim li As Integer
  lsFonts.Clear
  For li = 1 To Screen.FontCount
    If Trim(Screen.Fonts(li)) <> "" Then
      lsFonts.AddItem Trim(Screen.Fonts(li))
    End If
  Next
  lsFontslist_b.Tag = "7"
  lsFontsList.Tag = "-1"
  lsFontsList.ListIndex = 0
  lsFontsList_Click
End Sub

Private Sub lsFontsList_Click()
Dim lFonts As Variant
Dim li As Integer
  If val(lsFontsList.Tag) <> lsFontsList.ListIndex Then
    lsChosen.Clear
    If lsFontsList.ListIndex > -1 Then
      If InStr(lsFontsList.List(lsFontsList.ListIndex), ",") > 0 Then
        lFonts = Split(lsFontsList.List(lsFontsList.ListIndex), ",")
        For li = 0 To UBound(lFonts)
          lsChosen.AddItem Trim(lFonts(li))
        Next
      End If
    End If
    lsFontsList.Tag = lsFontsList.ListIndex
  End If
End Sub

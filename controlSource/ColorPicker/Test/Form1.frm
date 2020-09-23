VERSION 5.00
Object = "*\A..\prjColorPicker.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin prjColorPicker.ColorPicker ColorPicker1 
      Height          =   2205
      Left            =   1185
      TabIndex        =   1
      Top             =   1095
      Width           =   3165
      _extentx        =   5583
      _extenty        =   3889
   End
   Begin VB.CommandButton Command1 
      Caption         =   "color"
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   795
      Width           =   870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ColorPicker1_ColorCancel()
  MsgBox "cancel"
End Sub

Private Sub ColorPicker1_ColorSelect(ByVal Color As Long, ByVal WebRGBColor As String)
  MsgBox "ok"
End Sub

Private Sub Command1_Click()
  ColorPicker1.ShowColor = True
End Sub

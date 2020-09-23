VERSION 5.00
Begin VB.Form frmSpecialCharacters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Special Characters"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
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
   Icon            =   "frmSpecialCharacters.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   6225
      TabIndex        =   1
      Top             =   885
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Top             =   1305
      Width           =   945
   End
   Begin VB.TextBox txtValue 
      Height          =   300
      Left            =   1455
      MaxLength       =   28
      TabIndex        =   0
      Top             =   885
      Width           =   1095
   End
   Begin VB.PictureBox picCharacters 
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
      Height          =   4020
      Left            =   915
      Picture         =   "frmSpecialCharacters.frx":000C
      ScaleHeight     =   4020
      ScaleWidth      =   4980
      TabIndex        =   3
      Top             =   1365
      Width           =   4980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insert:"
      Height          =   195
      Left            =   915
      TabIndex        =   5
      Top             =   915
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Characters"
      Height          =   195
      Left            =   780
      TabIndex        =   4
      Top             =   330
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   300
      X2              =   8220
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   270
      X2              =   8160
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   420
      Picture         =   "frmSpecialCharacters.frx":1452
      Stretch         =   -1  'True
      Top             =   270
      Width           =   315
   End
End
Attribute VB_Name = "frmSpecialCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCharacters(9, 12) As String

Private Sub cmdAdd_Click()
  frmEditor.RTB(frmEditor.GetActiveRTB).Paste txtValue
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Initialize()
  InitXP
End Sub

Private Sub Form_Load()
  Call InitializeSpecialCharacters
End Sub

Private Sub picCharacters_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim posx As Integer
  Dim posy As Integer
  If x > (13 * 26) * Screen.TwipsPerPixelX Then Exit Sub
  posx = (x \ Screen.TwipsPerPixelX) \ 30
  If y > (10 * 28) * Screen.TwipsPerPixelY Then Exit Sub
  posy = (y \ Screen.TwipsPerPixelY) \ 30
  txtValue.Text = GetInsertValue(posy, posx)
End Sub


Private Sub InitializeSpecialCharacters()
 mCharacters(0, 0) = "&nbsp;"
 mCharacters(0, 1) = "&iexcl;"
 mCharacters(0, 2) = "&cent;"
 mCharacters(0, 3) = "&pound;"
 mCharacters(0, 4) = "&yen;"
 mCharacters(0, 5) = "&sect;"
 mCharacters(0, 6) = "&uml;"
 mCharacters(0, 7) = "&copy;"
 mCharacters(0, 8) = "&laquo;"
 mCharacters(0, 9) = "&not;"
 mCharacters(0, 10) = "&reg;"

 mCharacters(1, 0) = "&deg;"
 mCharacters(1, 1) = "&plusmn;"
 mCharacters(1, 2) = "&acute;"
 mCharacters(1, 3) = "&micro;"
 mCharacters(1, 4) = "&para;"
 mCharacters(1, 5) = "&middot;"
 mCharacters(1, 6) = "&cedil;"
 mCharacters(1, 7) = "&raquo;"
 mCharacters(1, 8) = "&iquest;"
 mCharacters(1, 9) = "&Agrave;"
 mCharacters(1, 10) = "&Aacute;"
 

 mCharacters(2, 0) = "&Acirc;"
 mCharacters(2, 1) = "&Atilde;"
 mCharacters(2, 2) = "&Auml;"
 mCharacters(2, 3) = "&Aring;"
 mCharacters(2, 4) = "&AElig;"
 mCharacters(2, 5) = "&Ccedil;"
 mCharacters(2, 6) = "&Egrave;"
 mCharacters(2, 7) = "&Eacute;"
 mCharacters(2, 8) = "&Ecirc;"
 mCharacters(2, 9) = "&Euml;"
 mCharacters(2, 10) = "&Igrave;"

 mCharacters(3, 0) = "&Iacute;"
 mCharacters(3, 1) = "&Icirc;"
 mCharacters(3, 2) = "&Iuml;"
 mCharacters(3, 3) = "&Ntilde;"
 mCharacters(3, 4) = "&Ograve;"
 mCharacters(3, 5) = "&Oacute;"
 mCharacters(3, 6) = "&Ocirc;"
 mCharacters(3, 7) = "&Otilde;"
 mCharacters(3, 8) = "&Ouml;"
 mCharacters(3, 9) = "&Oslash;"
 mCharacters(3, 10) = "&Ugrave;"
 

 mCharacters(4, 0) = "&Uacute;"
 mCharacters(4, 1) = "&Ucirc;"
 mCharacters(4, 2) = "&Uuml;"
 mCharacters(4, 3) = "&szlig;"
 mCharacters(4, 4) = "&agrave;"
 mCharacters(4, 5) = "&aacute;"
 mCharacters(4, 6) = "&acirc;"
 mCharacters(4, 7) = "&atilde;"
 mCharacters(4, 8) = "&auml;"
 mCharacters(4, 9) = "&aring;"
 mCharacters(4, 10) = "&aelig;"
  
 mCharacters(5, 0) = "&ccedil;"
 mCharacters(5, 1) = "&egrave;"
 mCharacters(5, 2) = "&eacute;"
 mCharacters(5, 3) = "&ecirc;"
 mCharacters(5, 4) = "&euml;"
 mCharacters(5, 5) = "&igrave;"
 mCharacters(5, 6) = "&iacute;"
 mCharacters(5, 7) = "&icirc;"
 mCharacters(5, 8) = "&iuml;"
 mCharacters(5, 9) = "&ntilde;"
 mCharacters(5, 10) = "&ograve;"
 
 mCharacters(6, 0) = "&oacute;"
 mCharacters(6, 1) = "&ocirc;"
 mCharacters(6, 2) = "&otilde;"
 mCharacters(6, 3) = "&ouml;"
 mCharacters(6, 4) = "&divide;"
 mCharacters(6, 5) = "&oslash;"
 mCharacters(6, 6) = "&ugrave;"
 mCharacters(6, 7) = "&uacute;"
 mCharacters(6, 8) = "&ucirc;"
 mCharacters(6, 9) = "&uuml;"
 mCharacters(6, 10) = "&yuml;"
 
 mCharacters(7, 0) = "&#8218;"
 mCharacters(7, 1) = "&#402;"
 mCharacters(7, 2) = "&#8222;"
 mCharacters(7, 3) = "&#8230;"
 mCharacters(7, 4) = "&#8224;"
 mCharacters(7, 5) = "&#8225;"
 mCharacters(7, 6) = "&#710;"
 mCharacters(7, 7) = "&#8240;"
 mCharacters(7, 8) = "&#8249;"
 mCharacters(7, 9) = "&#338;"
 mCharacters(7, 10) = "&#8216;"
 
 mCharacters(8, 0) = "&#8217;"
 mCharacters(8, 1) = "&#8220;"
 mCharacters(8, 2) = "&#8221;"
 mCharacters(8, 3) = "&#8226;"
 mCharacters(8, 4) = "&#8211;"
 mCharacters(8, 5) = "&#8212;"
 mCharacters(8, 6) = "&#732;"
 mCharacters(8, 7) = "&#8482;"
 mCharacters(8, 8) = "&#8250;"
 mCharacters(8, 9) = "&#339;"
 mCharacters(8, 10) = "&#376;"
End Sub


Private Function GetInsertValue(Lrow As Integer, lcol As Integer) As String
On Error Resume Next
  GetInsertValue = mCharacters(Lrow, lcol)
End Function

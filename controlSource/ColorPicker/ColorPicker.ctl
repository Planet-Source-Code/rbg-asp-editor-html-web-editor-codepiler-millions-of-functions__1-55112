VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl ColorPicker 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   3165
   ToolboxBitmap   =   "ColorPicker.ctx":0000
   Begin MSComctlLib.ImageList imgColorTables 
      Left            =   1140
      Top             =   3345
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   211
      ImageHeight     =   121
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorPicker.ctx":0312
            Key             =   "COLORCUBES"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorPicker.ctx":225B
            Key             =   "CONTINUOUSTONES"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorPicker.ctx":41A2
            Key             =   "WINDOWSOS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorPicker.ctx":5E22
            Key             =   "MACOS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorPicker.ctx":7DD1
            Key             =   "GRAYSCALE"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   525
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picColorBack 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2190
      Left            =   0
      ScaleHeight     =   2160
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   0
      Width           =   3165
      Begin VB.PictureBox lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   30
         ScaleHeight     =   270
         ScaleWidth      =   510
         TabIndex        =   6
         Top             =   30
         Width           =   540
      End
      Begin VB.CommandButton cmdDefault 
         Height          =   285
         Left            =   2160
         Picture         =   "ColorPicker.ctx":84F1
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Default Color"
         Top             =   45
         Width           =   315
      End
      Begin VB.CommandButton cmdCustomColor 
         Height          =   285
         Left            =   2505
         Picture         =   "ColorPicker.ctx":8870
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Custom Color"
         Top             =   45
         Width           =   315
      End
      Begin VB.CommandButton cmdTables 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2895
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Color Tables"
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox picColor 
         AutoSize        =   -1  'True
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
         Height          =   1815
         Left            =   -15
         MouseIcon       =   "ColorPicker.ctx":8925
         MousePointer    =   99  'Custom
         Picture         =   "ColorPicker.ctx":91EF
         ScaleHeight     =   1815
         ScaleWidth      =   3165
         TabIndex        =   1
         Top             =   360
         Width           =   3165
         Begin VB.Image spPointer 
            Height          =   165
            Left            =   1500
            Picture         =   "ColorPicker.ctx":B18E
            Top             =   1200
            Width           =   165
         End
      End
      Begin VB.Shape sp 
         Height          =   360
         Left            =   2880
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblHexCode 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   195
         Left            =   1365
         TabIndex        =   2
         Top             =   90
         Width           =   120
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuColorCubes 
         Caption         =   "Color Cubes"
      End
      Begin VB.Menu mnuContinuousTones 
         Caption         =   "Continuous Tones"
      End
      Begin VB.Menu mnuWindowsOS 
         Caption         =   "Windows OS"
      End
      Begin VB.Menu mnuMacOS 
         Caption         =   "Mac OS"
      End
      Begin VB.Menu mnuGrayScale 
         Caption         =   "Gray Scale"
      End
   End
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mShowColor As Boolean
Private mPoiterHeight As Single
Private mPoiterWidth As Single

Public Event ColorSelect(ByVal Color As Long, ByVal WebRGBFormat As String)
Public Event ColorCancel()
Public Event KeyPress(ByVal KeyAscii As Integer)

Private Const ColorSet = " ALICEBLUE(#F0F8FF) ANTIQUEWHITE(#FAEBD7) AQUA(#00FFFF) AQUAMARINE(#7FFFD4) AZURE(#F0FFFF) BEIGE(#F5F5DC) BISQUE(#FFE4C4) BLACK(#000000) BLANCHEDALMOND(#FFEBCD) BLUE(#0000FF) BLUEVIOLET(#8A2BE2) BROWN(#A52A2A) BURLYWOOD(#DEB887) CADETBLUE(#5F9EA0) CHARTREUSE(#7FFF00) CHOCOLATE(#D2691E) CORAL(#FF7F50) CORNFLOWER(#6495ED) CORNSILK(#FFF8DC) CRIMSON(#DC143C)" & _
                         " CYAN(#00FFFF) DARKBLUE(#00008B) DARKCYAN(#008B8B) DARKGOLDENROD(#B8860B) DARKGRAY(#A9A9A9) DARKGREEN(#006400) DARKKHAKI(#BDB76B) DARKMAGENTA(#8B008B) DARKOLIVEGREEN(#556B2F) DARKORANGE(#FF8C00) DARKORCHID(#9932CC) DARKRED(#8B0000) DARKSALMON(#E9967A) DARKSEAGREEN(#8FBC8B) DARKSLATEBLUE(#483D8B) DARKSLATEGRAY(#2F4F4F) DARKTURQUOISE(#00CED1) DARKVIOLET(#9400D3)" & _
                         " DEEPPINK(#FF1493) DEEPSKYBLUE(#00BFFF) DIMGRAY(#696969) DODGERBLUE(#1E90FF) FIREBRICK(#B22222) FLORALWHITE(#FFFAF0) FORESTGREEN(#228B22) FUCHIA(#FF00FF) GAINSBORO(#DCDCDC) GHOSTWHITE(#F8F8FF)GOLD(#FFD700) GOLDENROD(#DAA520) GRAY(#808080) GREEN(#008000)GREENYELLOW(#ADFF2F) HONEYDEW(#F0FFF0) HOTPINK(#FF69B4) INDIANRED(#CD5C5C) INDIGO(#4B0082) IVORY(#FFFFF0)" & _
                         " KHAKI(#F0E68C) LAVENDER(#E6E6FA) LAVENDERBLUSH(#FFF0F5) LAWNGREEN(#7CFC00) LEMONCHIFFON(#FFFACD) LIGHTBLUE(#ADD8E6) LIGHTCORAL(#F08080) LIGHTCYAN(#E0FFFF) LIGHTGOLDENRODYELLOW(#FAFAD2) LIGHTGREEN(#90EE90)LIGHTGREY(#D3D3D3) LIGHTPINK(#FFB6C1) LIGHTSALMON(#FFA07A) LIGHTSEAGREEN(#20B2AA) LIGHTSKYBLUE(#87CEFA) LIGHTSLATEGRAY(#778899) LIGHTSTEELBLUE(#B0C4DE)" & _
                         " LIGHTYELLOW(#FFFFE0) LIME(#00FF00) LIMEGREEN(#32CD32) LINEN(#FAF0E6) MAGENTA(#FF00FF) MAROON(#800000) MEDIUMAQUAMARINE(#66CDAA) MEDIUMBLUE(#0000CD) MEDIUMORCHID(#BA55D3) MEDIUMPURPLE(#9370DB) MEDIUMSEAGREEN(#3CB371) MEDIUMSLATEBLUE(#7B68EE) MEDIUMSPRINGGREEN(#00FA9A) MEDIUMTURQUOISE(#48D1CC) MEDIUMVIOLETRED(#C71585) MIDNIGHTBLUE(#191970) MINTCREAM(#F5FFFA)" & _
                         " MISTYROSE(#FFE4E1) MOCCASIN(#FFE4B5) NAVAJOWHITE(#FFDEAD) NAVY(#000080) OLDLACE(#FDF5E6) OLIVE(#808000) OLIVEDRAB(#6B8E23) ORANGE(#FFA500) ORANGERED(#FF4500) ORCHID(#DA70D6) PALEGOLDENROD(#EEE8AA) PALEGREEN(#98FB98) PALETURQUOISE(#AFEEEE) PALEVIOLETRED(#DB7093) PAPAYAWHIP(#FFEFD5) PEACHPUFF(#FFDAB9) PERU(#CD853F) PINK(#FFC0CB) PLUM(#DDA0DD) POWDERBLUE(#B0E0E6)" & _
                         " PURPLE(#800080) RED(#FF0000) ROSYBROWN(#BC8F8F) ROYALBLUE(#4169E1) SADDLEBROWN(#8B4513) SALMON(#FA8072) SANDYBROWN(#F4A460) SEAGREEN(#2E8B57) SEASHELL(#FFF5EE) SIENNA(#A0522D) SILVER(#C0C0C0) SKYBLUE(#87CEEB) SLATEBLUE(#6A5ACD) SLATEGRAY(#708090) SNOW(#FFFAFA) SPRINGGREEN(#00FF7F) STEELBLUE(#4682B4) TAN(#D2B48C) TEAL(#008080) THISTLE(#D8BFD8) TOMATO(#FF6347)" & _
                         " TURQUOISE(#40E0D0) VIOLET(#EE82EE) WHEAT(#F5DEB3) WHITE(#FFFFFF) WHITESMOKE(#F5F5F5) YELLOW(#FFFF00) YELLOWGREEN(#9ACD32)"


Private Sub cmdCustomColor_Click()
Dim lColor As Long
  On Error GoTo Cerr
  picColorBack.Visible = False
  CD.CancelError = True
  CD.ShowColor
  lColor = CD.Color
  RaiseEvent ColorSelect(lColor, GetWebRGB(lColor))
  Exit Sub
Cerr:
  RaiseEvent ColorCancel
End Sub

Private Sub cmdCustomColor_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    Me.ShowColor = False
  End If
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cmdCustomColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If lblColor.Tag <> "" And lblColor.Tag <> "-1" Then
    lblColor.BackColor = CLng(lblColor.Tag)
  End If
  lblHexCode.Caption = "System Color Picker"
End Sub

Private Sub cmdDefault_Click()
  lblColor.BackColor = picColorBack.BackColor
  lblHexCode.Caption = GetWebRGB(lblColor.BackColor)
End Sub

Private Sub cmdDefault_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    Me.ShowColor = False
  End If
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cmdDefault_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblColor.BackColor = picColorBack.BackColor
  lblHexCode.Caption = "Default Color"
End Sub

Private Sub cmdTables_Click()
  PopupMenu mnuFile, , cmdTables.Left, cmdTables.Top + cmdTables.Height
End Sub

Private Sub cmdTables_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    Me.ShowColor = False
  End If
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub mnuColorCubes_Click()
  On Error Resume Next
  CheckMenu 1
  picColor.Picture = imgColorTables.ListImages("COLORCUBES").Picture
  cmdDefault.SetFocus
End Sub

Private Sub mnuContinuousTones_Click()
  On Error Resume Next
  CheckMenu 2
  picColor.Picture = imgColorTables.ListImages("CONTINUOUSTONES").Picture
  cmdDefault.SetFocus
End Sub

Private Sub mnuGrayScale_Click()
  On Error Resume Next
  CheckMenu 5
  picColor.Picture = imgColorTables.ListImages("GRAYSCALE").Picture
  cmdDefault.SetFocus
End Sub

Private Sub mnuMacOS_Click()
  On Error Resume Next
  CheckMenu 4
  picColor.Picture = imgColorTables.ListImages("MACOS").Picture
  cmdDefault.SetFocus
End Sub

Private Sub mnuWindowsOS_Click()
  On Error Resume Next
  CheckMenu 3
  picColor.Picture = imgColorTables.ListImages("WINDOWSOS").Picture
  cmdDefault.SetFocus
End Sub

Private Sub picColor_Click()
  picColorBack.Visible = False
  RaiseEvent ColorSelect(lblColor.BackColor, GetWebRGB(lblColor.BackColor))
End Sub

Private Sub picColor_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    Me.ShowColor = False
  End If
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lLeft As Single
Dim lTop As Single
  On Error Resume Next
  If picColor.Point(X, Y) <> lblColor.BackColor Then
    lblColor.BackColor = picColor.Point(X, Y)
    lblHexCode.Caption = GetWebRGB(lblColor.BackColor)
  End If
  If spPointer.Visible = False Then spPointer.Visible = True
  lTop = Y - (Y Mod mPoiterHeight)
  lLeft = X - (X Mod mPoiterWidth)
  spPointer.Left = lLeft
  spPointer.Top = lTop
End Sub

Private Sub picColorBack_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    Me.ShowColor = False
  End If
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picColorBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If lblColor.Tag <> "" Then
    If lblColor.BackColor <> CLng(lblColor.Tag) Then
      lblColor.BackColor = CLng(lblColor.Tag)
      lblHexCode.Caption = GetWebRGB(lblColor.BackColor)
      lblHexCode.Refresh
    End If
  Else
    If lblHexCode.Caption <> GetWebRGB(lblColor.BackColor) Then
      lblHexCode.Caption = GetWebRGB(lblColor.BackColor)
      lblHexCode.Refresh
    End If
  End If
End Sub

Private Sub spPointer_Click()
  picColor_Click
End Sub

Private Sub UserControl_Initialize()
  mPoiterHeight = spPointer.Height - Screen.TwipsPerPixelY
  mPoiterWidth = spPointer.Width - Screen.TwipsPerPixelX
  mnuContinuousTones_Click
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  UserControl.Height = picColorBack.Height
  UserControl.Width = picColorBack.Width
End Sub

'
'Properties
'
Public Property Get ShowColor() As Boolean
  ShowColor = mShowColor
End Property

Public Property Let ShowColor(ByVal Value As Boolean)
  mShowColor = Value
  spPointer.Visible = False
  picColorBack.Visible = mShowColor
End Property

Public Property Get Color() As Long
  Color = lblColor.BackColor
End Property

Public Property Let Color(ByVal Value As Long)
  lblColor.Tag = Value
  lblColor.BackColor = Value
  lblHexCode.Caption = GetWebRGB(lblColor.BackColor)
End Property

Public Property Get ColorWebRGB() As String
  ColorWebRGB = GetWebRGB(lblColor.BackColor)
End Property

Public Property Let ColorWebRGB(ByVal Value As String)
Dim lColor As Long
  lColor = GetLongRGB(Value)
  lblColor.Tag = lColor
  lblColor.BackColor = lColor
  lblHexCode.Caption = GetWebRGB(lblColor.BackColor)
End Property


'
'User Function
'
Private Function CheckMenu(ByVal pIndex As Integer)
'
'Check the menu which is clicked
'
  mnuColorCubes.Checked = False
  mnuContinuousTones.Checked = False
  mnuWindowsOS.Checked = False
  mnuMacOS.Checked = False
  mnuGrayScale.Checked = False
  Select Case pIndex
  Case 1
    mnuColorCubes.Checked = True
  Case 2
    mnuContinuousTones.Checked = True
  Case 3
    mnuWindowsOS.Checked = True
  Case 4
    mnuMacOS.Checked = True
  Case 5
    mnuGrayScale.Checked = True
  End Select
End Function

Public Function GetWebRGB(ByVal pColor As Long) As String
'
'Return the long color value into rgb hex form for web only
'
Dim lHex As String
Dim lZeros As String
  On Error GoTo Cerr
  lHex = CStr(Hex(pColor))
  If Len(lHex) < 6 Then
    lZeros = String(6 - Len(lHex), "0")
    lHex = lZeros & lHex
  End If
  GetWebRGB = "#" & Mid(lHex, 5, 2) & Mid(lHex, 3, 2) & Mid(lHex, 1, 2)
  Exit Function
Cerr:
  GetWebRGB = "#"
End Function

Public Function GetLongRGB(ByVal pWebColor As String) As Long
'
'Get the long color for web color
'
  On Error GoTo Cerr
  If Mid(pWebColor, 1, 1) <> "#" Then pWebColor = GetColor(pWebColor)
  If Mid(pWebColor, 1, 1) = "#" Then pWebColor = Mid(pWebColor, 2)
  pWebColor = Mid(pWebColor, 5, 2) & Mid(pWebColor, 3, 2) & Mid(pWebColor, 1, 2)
  GetLongRGB = CLng("&H" & UCase(pWebColor))
  Exit Function
Cerr:
  GetLongRGB = -1
End Function

Private Function GetColor(ByVal pColor As String) As String
'
'Get the color from the Color set
'
Dim lPos As Long
  On Error Resume Next
  lPos = InStr(1, ColorSet, " " & pColor & "(#", vbTextCompare)
  If lPos > 0 Then
    lPos = lPos + Len(" " & pColor & "(#")
    If lPos > 0 Then
      GetColor = Mid(ColorSet, lPos, 6)
    End If
  End If
End Function

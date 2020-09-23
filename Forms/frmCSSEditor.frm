VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{55E69722-DD29-4623-A059-5B96E8A9018D}#1.2#0"; "ColorPicker.ocx"
Begin VB.Form frmCSSEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CSS Editor"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSSEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2820
      Top             =   150
   End
   Begin prjColorPicker.ColorPicker ColorPicker 
      Height          =   2190
      Left            =   1275
      TabIndex        =   135
      Top             =   2265
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   3863
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   1965
      Top             =   975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   64
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":005F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":00B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0105
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0158
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":01AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":01FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0251
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":02F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":034A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":039D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":03F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0443
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0496
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":04E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":053C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":058F
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":05E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0635
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0688
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":06DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":072E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0781
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":07D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0827
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":087A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":08CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0920
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0973
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":09C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0A19
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0ABF
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0B65
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0C0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0CB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0D57
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0DFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0EA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0F49
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":0FEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":1095
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":10E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":113B
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":118E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":11E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":1234
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":1287
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":12DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":132D
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":1380
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":13D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":1426
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":1479
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Imgs 
      Left            =   1230
      Top             =   2850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSSEditor.frx":14CC
            Key             =   "CLASS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvClasses 
      Height          =   6045
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   10663
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4763
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2655
      Top             =   5625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   6045
      Left            =   3120
      TabIndex        =   1
      Top             =   90
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   10663
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Font and Text"
      TabPicture(0)   =   "frmCSSEditor.frx":1546
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label10"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "imgFont"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cFontFamily"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cFontStyle"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cFontVariant"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cFontWeight"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cFontSize"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cTextDecoration"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cTextTransform"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cTextAlign"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cVerticalAlign"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cColor"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "picPaleteF"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Background"
      TabPicture(1)   =   "frmCSSEditor.frx":1562
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line3"
      Tab(1).Control(1)=   "Line4"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(5)=   "Label14"
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(7)=   "Label16"
      Tab(1).Control(8)=   "cBackgroundRepeat"
      Tab(1).Control(9)=   "cBackgroundAttachment"
      Tab(1).Control(10)=   "cBackgroundPosition"
      Tab(1).Control(11)=   "cBackgroundImage"
      Tab(1).Control(12)=   "cBackground"
      Tab(1).Control(13)=   "cmdBrowseB"
      Tab(1).Control(14)=   "cBackgroundColor"
      Tab(1).Control(15)=   "picPaleteB"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Layout"
      TabPicture(2)   =   "frmCSSEditor.frx":157E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cPadding"
      Tab(2).Control(1)=   "cPaddingLeft"
      Tab(2).Control(2)=   "cPaddingBottom"
      Tab(2).Control(3)=   "cPaddingRight"
      Tab(2).Control(4)=   "cPaddingTop"
      Tab(2).Control(5)=   "cMargin"
      Tab(2).Control(6)=   "cMarginLeft"
      Tab(2).Control(7)=   "cMarginBottom"
      Tab(2).Control(8)=   "cMarginRight"
      Tab(2).Control(9)=   "cMarginTop"
      Tab(2).Control(10)=   "Label26"
      Tab(2).Control(11)=   "Label25"
      Tab(2).Control(12)=   "Label24"
      Tab(2).Control(13)=   "Label23"
      Tab(2).Control(14)=   "Label22"
      Tab(2).Control(15)=   "Label21"
      Tab(2).Control(16)=   "Label20"
      Tab(2).Control(17)=   "Label19"
      Tab(2).Control(18)=   "Label18"
      Tab(2).Control(19)=   "Label17"
      Tab(2).Control(20)=   "Line6"
      Tab(2).Control(21)=   "Line5"
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "Border"
      TabPicture(3)   =   "frmCSSEditor.frx":159A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cBorderWidth"
      Tab(3).Control(1)=   "cBorderLeftWidth"
      Tab(3).Control(2)=   "cBorderBottomWidth"
      Tab(3).Control(3)=   "cBorderRightWidth"
      Tab(3).Control(4)=   "cBorderTopWidth"
      Tab(3).Control(5)=   "cBorderStyle"
      Tab(3).Control(6)=   "cBorderLeftStyle"
      Tab(3).Control(7)=   "cBorderBottomStyle"
      Tab(3).Control(8)=   "cBorderRightStyle"
      Tab(3).Control(9)=   "cBorderTopStyle"
      Tab(3).Control(10)=   "cBorderTopColor"
      Tab(3).Control(11)=   "cBorderRightColor"
      Tab(3).Control(12)=   "cBorderBottomColor"
      Tab(3).Control(13)=   "cBorderLeftColor"
      Tab(3).Control(14)=   "cBorderColor"
      Tab(3).Control(15)=   "cBorderTop"
      Tab(3).Control(16)=   "cBorderRight"
      Tab(3).Control(17)=   "cBorderBottom"
      Tab(3).Control(18)=   "cBorderLeft"
      Tab(3).Control(19)=   "cBorder"
      Tab(3).Control(20)=   "Label50"
      Tab(3).Control(21)=   "Label49"
      Tab(3).Control(22)=   "Label48"
      Tab(3).Control(23)=   "Label47"
      Tab(3).Control(24)=   "Label46"
      Tab(3).Control(25)=   "Label45"
      Tab(3).Control(26)=   "Label44"
      Tab(3).Control(27)=   "Label43"
      Tab(3).Control(28)=   "Label42"
      Tab(3).Control(29)=   "Label41"
      Tab(3).Control(30)=   "Label40"
      Tab(3).Control(31)=   "Label39"
      Tab(3).Control(32)=   "Label38"
      Tab(3).Control(33)=   "Label37"
      Tab(3).Control(34)=   "Label36"
      Tab(3).Control(35)=   "Label35"
      Tab(3).Control(36)=   "Label34"
      Tab(3).Control(37)=   "Label33"
      Tab(3).Control(38)=   "Label32"
      Tab(3).Control(39)=   "Label31"
      Tab(3).Control(40)=   "Label30"
      Tab(3).Control(41)=   "Label29"
      Tab(3).Control(42)=   "Label28"
      Tab(3).Control(43)=   "Label27"
      Tab(3).Control(44)=   "Line8"
      Tab(3).Control(45)=   "Line7"
      Tab(3).ControlCount=   46
      TabCaption(4)   =   "Positioning"
      TabPicture(4)   =   "frmCSSEditor.frx":15B6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cZIndex"
      Tab(4).Control(1)=   "cWidth"
      Tab(4).Control(2)=   "cTop"
      Tab(4).Control(3)=   "cLeft"
      Tab(4).Control(4)=   "cHeight"
      Tab(4).Control(5)=   "cClip"
      Tab(4).Control(6)=   "cOverflow"
      Tab(4).Control(7)=   "cPosition"
      Tab(4).Control(8)=   "cVisibility"
      Tab(4).Control(9)=   "Label59"
      Tab(4).Control(10)=   "Label58"
      Tab(4).Control(11)=   "Label57"
      Tab(4).Control(12)=   "Label56"
      Tab(4).Control(13)=   "Label55"
      Tab(4).Control(14)=   "Label54"
      Tab(4).Control(15)=   "Label53"
      Tab(4).Control(16)=   "Label52"
      Tab(4).Control(17)=   "Label51"
      Tab(4).Control(18)=   "Line10"
      Tab(4).Control(19)=   "Line9"
      Tab(4).ControlCount=   20
      TabCaption(5)   =   "Classification"
      TabPicture(5)   =   "frmCSSEditor.frx":15D2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdBorwseC"
      Tab(5).Control(1)=   "cFloat"
      Tab(5).Control(2)=   "cClear"
      Tab(5).Control(3)=   "cLineStyleImage"
      Tab(5).Control(4)=   "cLineStyle"
      Tab(5).Control(5)=   "cDisplay"
      Tab(5).Control(6)=   "cLineStyleType"
      Tab(5).Control(7)=   "cLineStylePosition"
      Tab(5).Control(8)=   "Label67"
      Tab(5).Control(9)=   "Label66"
      Tab(5).Control(10)=   "Label65"
      Tab(5).Control(11)=   "Label64"
      Tab(5).Control(12)=   "Label63"
      Tab(5).Control(13)=   "Label62"
      Tab(5).Control(14)=   "Label61"
      Tab(5).Control(15)=   "Label60"
      Tab(5).Control(16)=   "Line12"
      Tab(5).Control(17)=   "Line11"
      Tab(5).ControlCount=   18
      Begin VB.PictureBox picPaleteB 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -72120
         ScaleHeight     =   255
         ScaleWidth      =   270
         TabIndex        =   134
         Top             =   960
         Width           =   270
         Begin VB.Image imgPaleteB 
            Height          =   255
            Left            =   0
            MouseIcon       =   "frmCSSEditor.frx":15EE
            MousePointer    =   99  'Custom
            Picture         =   "frmCSSEditor.frx":1EB8
            Top             =   0
            Width           =   270
         End
      End
      Begin VB.PictureBox picPaleteF 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5670
         ScaleHeight     =   255
         ScaleWidth      =   270
         TabIndex        =   133
         Top             =   1515
         Width           =   270
         Begin VB.Image imgPaleteF 
            Height          =   255
            Left            =   0
            MouseIcon       =   "frmCSSEditor.frx":1F40
            MousePointer    =   99  'Custom
            Picture         =   "frmCSSEditor.frx":280A
            Top             =   0
            Width           =   270
         End
      End
      Begin VB.TextBox cBackgroundColor 
         Height          =   300
         Left            =   -73575
         TabIndex        =   132
         Text            =   "#000000"
         Top             =   922
         Width           =   1425
      End
      Begin VB.CommandButton cmdBorwseC 
         Caption         =   "..."
         Height          =   300
         Left            =   -69690
         TabIndex        =   131
         Top             =   2370
         Width           =   360
      End
      Begin VB.CommandButton cmdBrowseB 
         Caption         =   "..."
         Height          =   300
         Left            =   -69645
         TabIndex        =   130
         Top             =   1455
         Width           =   360
      End
      Begin VB.ComboBox cFloat 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2892
         Left            =   -73920
         List            =   "frmCSSEditor.frx":289F
         TabIndex        =   61
         Top             =   3870
         Width           =   1470
      End
      Begin VB.ComboBox cClear 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":28B6
         Left            =   -73920
         List            =   "frmCSSEditor.frx":28C6
         TabIndex        =   62
         Top             =   4350
         Width           =   1470
      End
      Begin VB.TextBox cLineStyleImage 
         Height          =   300
         Left            =   -73245
         TabIndex        =   59
         Top             =   2370
         Width           =   3540
      End
      Begin VB.TextBox cLineStyle 
         Height          =   300
         Left            =   -73245
         TabIndex        =   60
         Top             =   2865
         Width           =   1155
      End
      Begin VB.ComboBox cDisplay 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":28E3
         Left            =   -73245
         List            =   "frmCSSEditor.frx":28F3
         TabIndex        =   56
         Top             =   840
         Width           =   1470
      End
      Begin VB.ComboBox cLineStyleType 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2917
         Left            =   -73245
         List            =   "frmCSSEditor.frx":2933
         TabIndex        =   57
         Top             =   1350
         Width           =   1470
      End
      Begin VB.ComboBox cLineStylePosition 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2988
         Left            =   -73245
         List            =   "frmCSSEditor.frx":2992
         TabIndex        =   58
         Top             =   1860
         Width           =   1470
      End
      Begin VB.TextBox cZIndex 
         Height          =   300
         Left            =   -73830
         TabIndex        =   51
         Top             =   2767
         Width           =   1155
      End
      Begin VB.TextBox cWidth 
         Height          =   300
         Left            =   -73830
         TabIndex        =   55
         Top             =   4740
         Width           =   1155
      End
      Begin VB.TextBox cTop 
         Height          =   300
         Left            =   -73830
         TabIndex        =   54
         Top             =   4245
         Width           =   1155
      End
      Begin VB.TextBox cLeft 
         Height          =   300
         Left            =   -73830
         TabIndex        =   53
         Top             =   3765
         Width           =   1155
      End
      Begin VB.TextBox cHeight 
         Height          =   300
         Left            =   -73830
         TabIndex        =   52
         Top             =   3285
         Width           =   1155
      End
      Begin VB.TextBox cClip 
         Height          =   300
         Left            =   -73830
         TabIndex        =   47
         Top             =   832
         Width           =   1155
      End
      Begin VB.ComboBox cOverflow 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":29A7
         Left            =   -73830
         List            =   "frmCSSEditor.frx":29B7
         TabIndex        =   48
         Top             =   1290
         Width           =   1470
      End
      Begin VB.ComboBox cPosition 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":29DA
         Left            =   -73830
         List            =   "frmCSSEditor.frx":29E7
         TabIndex        =   49
         Top             =   1785
         Width           =   1470
      End
      Begin VB.ComboBox cVisibility 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2A07
         Left            =   -73830
         List            =   "frmCSSEditor.frx":2A14
         TabIndex        =   50
         Top             =   2280
         Width           =   1470
      End
      Begin VB.ComboBox cBorderWidth 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2A32
         Left            =   -71070
         List            =   "frmCSSEditor.frx":2A3F
         TabIndex        =   41
         Top             =   2895
         Width           =   1470
      End
      Begin VB.ComboBox cBorderLeftWidth 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2A58
         Left            =   -71070
         List            =   "frmCSSEditor.frx":2A65
         TabIndex        =   40
         Top             =   2409
         Width           =   1470
      End
      Begin VB.ComboBox cBorderBottomWidth 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2A7E
         Left            =   -71070
         List            =   "frmCSSEditor.frx":2A8B
         TabIndex        =   39
         Top             =   1926
         Width           =   1470
      End
      Begin VB.ComboBox cBorderRightWidth 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2AA4
         Left            =   -71070
         List            =   "frmCSSEditor.frx":2AB1
         TabIndex        =   38
         Top             =   1443
         Width           =   1470
      End
      Begin VB.ComboBox cBorderTopWidth 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2ACA
         Left            =   -71070
         List            =   "frmCSSEditor.frx":2AD7
         TabIndex        =   37
         Top             =   960
         Width           =   1470
      End
      Begin VB.ComboBox cBorderStyle 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2AF0
         Left            =   -73995
         List            =   "frmCSSEditor.frx":2B0F
         TabIndex        =   36
         Top             =   5610
         Width           =   1470
      End
      Begin VB.ComboBox cBorderLeftStyle 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2B56
         Left            =   -73995
         List            =   "frmCSSEditor.frx":2B75
         TabIndex        =   35
         Top             =   5118
         Width           =   1470
      End
      Begin VB.ComboBox cBorderBottomStyle 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2BBC
         Left            =   -73995
         List            =   "frmCSSEditor.frx":2BDB
         TabIndex        =   34
         Top             =   4627
         Width           =   1470
      End
      Begin VB.ComboBox cBorderRightStyle 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2C22
         Left            =   -73995
         List            =   "frmCSSEditor.frx":2C41
         TabIndex        =   33
         Top             =   4136
         Width           =   1470
      End
      Begin VB.ComboBox cBorderTopStyle 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2C88
         Left            =   -73995
         List            =   "frmCSSEditor.frx":2CA7
         TabIndex        =   32
         Top             =   3645
         Width           =   1470
      End
      Begin VB.TextBox cBorderTopColor 
         Height          =   300
         Left            =   -73995
         TabIndex        =   27
         Top             =   960
         Width           =   1155
      End
      Begin VB.TextBox cBorderRightColor 
         Height          =   300
         Left            =   -73995
         TabIndex        =   28
         Top             =   1440
         Width           =   1155
      End
      Begin VB.TextBox cBorderBottomColor 
         Height          =   300
         Left            =   -73995
         TabIndex        =   29
         Top             =   1920
         Width           =   1155
      End
      Begin VB.TextBox cBorderLeftColor 
         Height          =   300
         Left            =   -73995
         TabIndex        =   30
         Top             =   2415
         Width           =   1155
      End
      Begin VB.TextBox cBorderColor 
         Height          =   300
         Left            =   -73995
         TabIndex        =   31
         Top             =   2895
         Width           =   1155
      End
      Begin VB.TextBox cBorderTop 
         Height          =   300
         Left            =   -71070
         TabIndex        =   42
         Top             =   3645
         Width           =   2310
      End
      Begin VB.TextBox cBorderRight 
         Height          =   300
         Left            =   -71085
         TabIndex        =   43
         Top             =   4140
         Width           =   2310
      End
      Begin VB.TextBox cBorderBottom 
         Height          =   300
         Left            =   -71085
         TabIndex        =   44
         Top             =   4620
         Width           =   2310
      End
      Begin VB.TextBox cBorderLeft 
         Height          =   300
         Left            =   -71085
         TabIndex        =   45
         Top             =   5115
         Width           =   2310
      End
      Begin VB.TextBox cBorder 
         Height          =   300
         Left            =   -71085
         TabIndex        =   46
         Top             =   5610
         Width           =   2310
      End
      Begin VB.TextBox cPadding 
         Height          =   300
         Left            =   -73110
         TabIndex        =   26
         Top             =   5475
         Width           =   3975
      End
      Begin VB.TextBox cPaddingLeft 
         Height          =   300
         Left            =   -73110
         TabIndex        =   25
         Top             =   4954
         Width           =   1155
      End
      Begin VB.TextBox cPaddingBottom 
         Height          =   300
         Left            =   -73110
         TabIndex        =   24
         Top             =   4436
         Width           =   1155
      End
      Begin VB.TextBox cPaddingRight 
         Height          =   300
         Left            =   -73110
         TabIndex        =   23
         Top             =   3918
         Width           =   1155
      End
      Begin VB.TextBox cPaddingTop 
         Height          =   300
         Left            =   -73110
         TabIndex        =   22
         Top             =   3400
         Width           =   1155
      End
      Begin VB.TextBox cMargin 
         Height          =   300
         Left            =   -73110
         TabIndex        =   21
         Top             =   2882
         Width           =   3975
      End
      Begin VB.TextBox cMarginLeft 
         Height          =   300
         Left            =   -73110
         TabIndex        =   20
         Top             =   2364
         Width           =   1155
      End
      Begin VB.TextBox cMarginBottom 
         Height          =   300
         Left            =   -73110
         TabIndex        =   19
         Top             =   1846
         Width           =   1155
      End
      Begin VB.TextBox cMarginRight 
         Height          =   300
         Left            =   -73110
         TabIndex        =   18
         Top             =   1328
         Width           =   1155
      End
      Begin VB.TextBox cMarginTop 
         Height          =   300
         Left            =   -73110
         TabIndex        =   17
         Top             =   810
         Width           =   1155
      End
      Begin VB.TextBox cBackground 
         Height          =   1575
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   3615
         Width           =   3900
      End
      Begin VB.TextBox cBackgroundImage 
         Height          =   300
         Left            =   -73575
         TabIndex        =   12
         Top             =   1458
         Width           =   3900
      End
      Begin VB.ComboBox cBackgroundPosition 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2CEE
         Left            =   -73575
         List            =   "frmCSSEditor.frx":2D07
         TabIndex        =   15
         Top             =   3075
         Width           =   1800
      End
      Begin VB.ComboBox cBackgroundAttachment 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2D48
         Left            =   -73575
         List            =   "frmCSSEditor.frx":2D52
         TabIndex        =   14
         Top             =   2529
         Width           =   1290
      End
      Begin VB.ComboBox cBackgroundRepeat 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2D65
         Left            =   -73575
         List            =   "frmCSSEditor.frx":2D75
         TabIndex        =   13
         Top             =   1986
         Width           =   1290
      End
      Begin VB.TextBox cColor 
         Height          =   300
         Left            =   4230
         TabIndex        =   4
         Text            =   "#000000"
         Top             =   1470
         Width           =   1425
      End
      Begin VB.ComboBox cVerticalAlign 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2D9E
         Left            =   1770
         List            =   "frmCSSEditor.frx":2DA8
         TabIndex        =   11
         Top             =   5370
         Width           =   1290
      End
      Begin VB.ComboBox cTextAlign 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2DB8
         Left            =   1770
         List            =   "frmCSSEditor.frx":2DC8
         TabIndex        =   10
         Top             =   4800
         Width           =   1290
      End
      Begin VB.ComboBox cTextTransform 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2DEA
         Left            =   1770
         List            =   "frmCSSEditor.frx":2DFA
         TabIndex        =   9
         Top             =   4245
         Width           =   1290
      End
      Begin VB.ComboBox cTextDecoration 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2E26
         Left            =   1770
         List            =   "frmCSSEditor.frx":2E39
         TabIndex        =   8
         Top             =   3690
         Width           =   1290
      End
      Begin VB.TextBox cFontSize 
         Height          =   300
         Left            =   1770
         TabIndex        =   7
         Top             =   3150
         Width           =   510
      End
      Begin VB.ComboBox cFontWeight 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2E6D
         Left            =   1770
         List            =   "frmCSSEditor.frx":2E98
         TabIndex        =   6
         Top             =   2580
         Width           =   1290
      End
      Begin VB.ComboBox cFontVariant 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2EE8
         Left            =   1770
         List            =   "frmCSSEditor.frx":2EF2
         TabIndex        =   5
         Top             =   2025
         Width           =   1290
      End
      Begin VB.ComboBox cFontStyle 
         Height          =   315
         ItemData        =   "frmCSSEditor.frx":2F0A
         Left            =   1770
         List            =   "frmCSSEditor.frx":2F17
         TabIndex        =   3
         Top             =   1470
         Width           =   1290
      End
      Begin VB.TextBox cFontFamily 
         Height          =   300
         Left            =   1770
         TabIndex        =   2
         Top             =   915
         Width           =   3900
      End
      Begin VB.Image imgFont 
         Height          =   240
         Left            =   5700
         MouseIcon       =   "frmCSSEditor.frx":2F34
         MousePointer    =   99  'Custom
         Picture         =   "frmCSSEditor.frx":37FE
         Top             =   975
         Width           =   225
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clear:"
         Height          =   195
         Left            =   -74805
         TabIndex        =   129
         Top             =   4410
         Width           =   435
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Float:"
         Height          =   195
         Left            =   -74805
         TabIndex        =   128
         Top             =   3930
         Width           =   420
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Positioning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -74805
         TabIndex        =   127
         Top             =   3450
         Width           =   1035
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line style:"
         Height          =   195
         Left            =   -74790
         TabIndex        =   126
         Top             =   2925
         Width           =   735
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line style image:"
         Height          =   195
         Left            =   -74790
         TabIndex        =   125
         Top             =   2418
         Width           =   1200
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line style position:"
         Height          =   195
         Left            =   -74790
         TabIndex        =   124
         Top             =   1912
         Width           =   1335
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line style type:"
         Height          =   195
         Left            =   -74790
         TabIndex        =   123
         Top             =   1406
         Width           =   1110
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display:"
         Height          =   195
         Left            =   -74790
         TabIndex        =   122
         Top             =   900
         Width           =   570
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00808080&
         X1              =   -74835
         X2              =   -68790
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74835
         X2              =   -68790
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   195
         Left            =   -74715
         TabIndex        =   121
         Top             =   4785
         Width           =   480
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
         Height          =   195
         Left            =   -74715
         TabIndex        =   120
         Top             =   4305
         Width           =   330
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   195
         Left            =   -74715
         TabIndex        =   119
         Top             =   3825
         Width           =   345
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         Height          =   195
         Left            =   -74715
         TabIndex        =   118
         Top             =   3345
         Width           =   525
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Z-index"
         Height          =   195
         Left            =   -74715
         TabIndex        =   117
         Top             =   2820
         Width           =   540
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visibility:"
         Height          =   195
         Left            =   -74715
         TabIndex        =   116
         Top             =   2340
         Width           =   615
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         Height          =   195
         Left            =   -74715
         TabIndex        =   115
         Top             =   1845
         Width           =   615
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Overflow:"
         Height          =   195
         Left            =   -74715
         TabIndex        =   114
         Top             =   1365
         Width           =   720
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clip:"
         Height          =   195
         Left            =   -74715
         TabIndex        =   113
         Top             =   885
         Width           =   315
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -71955
         TabIndex        =   112
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   111
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   110
         Top             =   1503
         Width           =   435
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   109
         Top             =   1986
         Width           =   570
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   108
         Top             =   2469
         Width           =   345
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   107
         Top             =   2955
         Width           =   480
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -74835
         TabIndex        =   106
         Top             =   3315
         Width           =   465
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   105
         Top             =   3705
         Width           =   330
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   104
         Top             =   4196
         Width           =   435
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   103
         Top             =   4687
         Width           =   570
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   102
         Top             =   5178
         Width           =   345
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Style:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   101
         Top             =   5670
         Width           =   420
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -74835
         TabIndex        =   100
         Top             =   675
         Width           =   480
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -71955
         TabIndex        =   99
         Top             =   3315
         Width           =   615
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   98
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   97
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   96
         Top             =   1980
         Width           =   570
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   95
         Top             =   2460
         Width           =   345
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   94
         Top             =   2955
         Width           =   435
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   93
         Top             =   3705
         Width           =   330
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   92
         Top             =   4185
         Width           =   435
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   91
         Top             =   4680
         Width           =   570
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   90
         Top             =   5175
         Width           =   345
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border:"
         Height          =   195
         Left            =   -71955
         TabIndex        =   89
         Top             =   5670
         Width           =   540
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Padding:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   88
         Top             =   5528
         Width           =   630
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Padding left:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   87
         Top             =   5007
         Width           =   915
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Padding bottom:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   86
         Top             =   4489
         Width           =   1185
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Padding right:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   85
         Top             =   3971
         Width           =   1005
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Padding top:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   84
         Top             =   3453
         Width           =   915
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   83
         Top             =   2935
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin left:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   82
         Top             =   2417
         Width           =   825
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin bottom:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   81
         Top             =   1899
         Width           =   1095
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin right:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   80
         Top             =   1381
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin top:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   79
         Top             =   863
         Width           =   825
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74835
         X2              =   -68790
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   -74835
         X2              =   -68790
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74835
         X2              =   -68790
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   -74835
         X2              =   -68790
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Background:"
         Height          =   195
         Left            =   -74745
         TabIndex        =   78
         Top             =   3615
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         Height          =   195
         Left            =   -74730
         TabIndex        =   77
         Top             =   3135
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachment:"
         Height          =   195
         Left            =   -74730
         TabIndex        =   76
         Top             =   2595
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repeat:"
         Height          =   195
         Left            =   -74730
         TabIndex        =   75
         Top             =   2055
         Width           =   585
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image:"
         Height          =   195
         Left            =   -74730
         TabIndex        =   74
         Top             =   1515
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   195
         Left            =   -74730
         TabIndex        =   73
         Top             =   975
         Width           =   435
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   -74835
         X2              =   -68790
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74835
         X2              =   -68790
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74835
         X2              =   -68880
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   -74835
         X2              =   -68790
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   195
         Left            =   3645
         TabIndex        =   72
         Top             =   1515
         Width           =   435
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   165
         X2              =   6210
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   165
         X2              =   6210
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Vertical Align:"
         Height          =   195
         Left            =   270
         TabIndex        =   71
         Top             =   5430
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Align:"
         Height          =   195
         Left            =   270
         TabIndex        =   70
         Top             =   4860
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Transform:"
         Height          =   195
         Left            =   270
         TabIndex        =   69
         Top             =   4305
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Decoration:"
         Height          =   195
         Left            =   270
         TabIndex        =   68
         Top             =   3750
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         Height          =   195
         Left            =   270
         TabIndex        =   67
         Top             =   3195
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight:"
         Height          =   195
         Left            =   270
         TabIndex        =   66
         Top             =   2640
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Variant:"
         Height          =   195
         Left            =   270
         TabIndex        =   65
         Top             =   2085
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Style:"
         Height          =   195
         Left            =   270
         TabIndex        =   64
         Top             =   1530
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font:"
         Height          =   195
         Left            =   270
         TabIndex        =   63
         Top             =   975
         Width           =   390
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu myhy1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu myhy2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuClasses 
      Caption         =   "Classes"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu myhy3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmCSSEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Mfilename As String

Private mChange As Boolean
Private mFilechange As Boolean
Private mProperties() As String
Private mClass As Integer

Private Sub cBackground_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBackgroundAttachment_Click()
  mChange = True
End Sub

Private Sub cBackgroundAttachment_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBackgroundColor_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBackgroundColor_Click()
  mChange = True
End Sub

Private Sub cBackgroundColor_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lColor As Long
  lColor = ColorPicker.GetLongRGB(cBackgroundColor.Text)
  picPaleteB.BackColor = IIf(lColor <> -1, lColor, vbButtonFace)
End Sub

Private Sub cBackgroundImage_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBackgroundImage_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lColor As Long
  lColor = ColorPicker.GetLongRGB(cBackgroundColor.Text)
  picPaleteB.BackColor = IIf(lColor <> -1, lColor, vbButtonFace)
End Sub

Private Sub cBackgroundPosition_Click()
  mChange = True
End Sub

Private Sub cBackgroundPosition_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBackgroundRepeat_Click()
  mChange = True
End Sub

Private Sub cBackgroundRepeat_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cClip_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cColor_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lColor As Long
  lColor = ColorPicker.GetLongRGB(cColor.Text)
  picPaleteF.BackColor = IIf(lColor <> -1, lColor, vbButtonFace)
End Sub

Private Sub cHeight_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cmdBorwseC_Click()
  On Error GoTo Cerr
  CD.FileName = ""
  CD.Filter = "All Picture Files(*.gif *.jpg)|*.gif;*.jpg|All files(*.*)|*.*"
  CD.CancelError = True
  CD.ShowOpen
  cLineStyleImage.Text = "url(" & CD.FileName & ")"
  mChange = True
Cerr:
End Sub

Private Sub cmdBrowseB_Click()
  On Error GoTo Cerr
  CD.FileName = ""
  CD.Filter = "All Picture Files(*.gif,*.jpg)|*.gif;*.jpg|All files(*.*)|*.*"
  CD.CancelError = True
  CD.ShowOpen
  cBackgroundImage.Text = "url(" & CD.FileName & ")"
  mChange = True
Cerr:
End Sub

Private Sub ColorPicker_ColorSelect(ByVal Color As Long, ByVal WebRGBFormat As String)
  If ColorPicker.Tag = 1 Then
    picPaleteF.BackColor = Color
    cColor.Text = WebRGBFormat
  Else
    picPaleteB.BackColor = Color
    cBackgroundColor.Text = WebRGBFormat
  End If
  mChange = True
End Sub

Private Sub cTop_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cLeft_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cWidth_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cZIndex_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cLineStyleImage_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cLineStyle_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cFontStyle_Click()
  mChange = True
End Sub

Private Sub cFontStyle_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cFontVariant_Click()
  mChange = True
End Sub

Private Sub cFontVariant_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cFontWeight_Click()
  mChange = True
End Sub

Private Sub cFontWeight_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cMarginTop_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cMarginRight_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cMarginBottom_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cMarginLeft_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cMargin_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cOverflow_Click()
  mChange = True
End Sub

Private Sub cOverflow_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cVisibility_Click()
  mChange = True
End Sub

Private Sub cVisibility_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cPosition_Click()
  mChange = True
End Sub

Private Sub cPosition_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cDisplay_Click()
  mChange = True
End Sub

Private Sub cDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cLineStyleType_Click()
  mChange = True
End Sub

Private Sub cLineStyleType_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cLineStylePosition_Click()
  mChange = True
End Sub

Private Sub cLineStylePosition_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cFloat_Click()
  mChange = True
End Sub

Private Sub cFloat_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cClear_Click()
  mChange = True
End Sub

Private Sub cClear_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cPaddingTop_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cPaddingRight_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cPaddingBottom_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cPaddingLeft_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cPadding_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderTopColor_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderRightColor_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderBottomColor_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderLeftColor_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderColor_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderTopWidth_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderRightWidth_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderBottomWidth_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderLeftWidth_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderWidth_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderTopWidth_Click()
  mChange = True
End Sub

Private Sub cBorderRightWidth_Click()
  mChange = True
End Sub

Private Sub cBorderBottomWidth_Click()
  mChange = True
End Sub

Private Sub cBorderLeftWidth_Click()
  mChange = True
End Sub

Private Sub cBorderWidth_Click()
  mChange = True
End Sub

Private Sub cBorderTopStyle_Click()
  mChange = True
End Sub

Private Sub cBorderRightStyle_Click()
  mChange = True
End Sub

Private Sub cBorderBottomStyle_Click()
  mChange = True
End Sub

Private Sub cBorderLeftStyle_Click()
  mChange = True
End Sub

Private Sub cBorderStyle_Click()
  mChange = True
End Sub

Private Sub cBorderTopStyle_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderRightStyle_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderBottomStyle_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderLeftStyle_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderStyle_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderTop_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderRight_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderBottom_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorderLeft_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cBorder_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cTextAlign_Click()
  mChange = True
End Sub

Private Sub cTextAlign_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cTextDecoration_Click()
  mChange = True
End Sub

Private Sub cTextDecoration_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cVerticalAlign_Click()
  mChange = True
End Sub

Private Sub cVerticalAlign_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cTextTransform_Click()
  mChange = True
End Sub

Private Sub cTextTransform_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    If ColorPicker.ShowColor Then
      ColorPicker.ShowColor = False
      KeyCode = 0
    End If
  End If
End Sub

Private Sub imgFont_Click()
  frmFonts.Show vbModal
End Sub

Private Sub imgPaleteF_Click()
  On Error Resume Next
  ColorPicker.Tag = 1
  ColorPicker.Top = tabMain.Top + picPaleteF.Top + picPaleteF.Height
  ColorPicker.Left = tabMain.Left + picPaleteF.Left - (ColorPicker.Width - picPaleteF.Width)
  If Trim(cColor.Text) <> "" Then ColorPicker.ColorWebRGB = cColor.Text
  ColorPicker.ShowColor = Not ColorPicker.ShowColor
End Sub

Private Sub imgPaleteB_Click()
  On Error Resume Next
  ColorPicker.Tag = 2
  ColorPicker.Top = tabMain.Top + picPaleteB.Top + picPaleteB.Height
  ColorPicker.Left = tabMain.Left + picPaleteB.Left - (ColorPicker.Width - picPaleteF.Width)
  If Trim(cBackgroundColor.Text) <> "" Then ColorPicker.ColorWebRGB = cBackgroundColor.Text
  ColorPicker.ShowColor = Not ColorPicker.ShowColor
End Sub

Private Sub lsvClasses_DblClick()
  tmr.Enabled = True
End Sub

Private Sub lsvClasses_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    mnuRemove_Click
  ElseIf KeyCode = vbKeyF2 Then
    mnuRename_Click
  End If
End Sub

Private Sub mnuClose_Click()
  mnuNew_Click
End Sub

Private Sub mnuRemove_Click()
  If Not lsvClasses.SelectedItem Is Nothing Then
    If MsgBox("Are you sure remove the class '" & lsvClasses.SelectedItem.Text & "'?", vbQuestion + vbYesNo, Mtitle) = vbYes Then
      lsvClasses.ListItems.Remove lsvClasses.SelectedItem.Key
      mChange = False
    End If
  End If
  If lsvClasses.ListItems.Count = 0 Then tabMain.Enabled = False
End Sub

Private Sub mnuRename_Click()
  If Not lsvClasses.SelectedItem Is Nothing Then
    lsvClasses.StartLabelEdit
  End If
End Sub

Private Sub mnuSave_Click()
  If mChange Then SaveProperties lsvClasses.Tag
  SaveCSS Mfilename
End Sub

Private Sub mnuSaveAs_Click()
  If mChange Then SaveProperties lsvClasses.Tag
  SaveCSS ""
End Sub

Private Sub Form_Load()
  lsvClasses.SmallIcons = Imgs
  mClass = 1
  'LoadColors
  ColorPicker.ShowColor = False
  ClearProperties
  InitProperties
  tabMain.Tab = 0
  If Mfilename <> "" Then
    LoadCSS Mfilename
  Else
    mnuNew_Click
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim lResult As VbMsgBoxResult
  If mFilechange Or mChange Then
    lResult = MsgBox(IIf(Mfilename <> "", mID(Mfilename, InStrRev(Mfilename, "\") + 1), "New Stylesheet") & vbCrLf & vbCrLf & "has been changed. Do you want to save the changes?", vbQuestion + vbYesNoCancel, Mtitle)
    If lResult = vbCancel Then
      Cancel = 1
      Exit Sub
    ElseIf lResult = vbYes Then
      SaveCSS Mfilename
    End If
  End If
  Set frmCSSEditor = Nothing
End Sub

Private Sub lsvClasses_AfterLabelEdit(Cancel As Integer, NewString As String)
  On Error GoTo Cerr
  If lsvClasses.SelectedItem.Text <> NewString Then
    lsvClasses.SelectedItem.Key = NewString 'Ucase Changed
    lsvClasses.Tag = NewString
  End If
  Exit Sub
Cerr:
  MsgBox "Class already exist.", vbInformation, Mtitle
  Cancel = 1
  lsvClasses.StartLabelEdit
End Sub

Private Sub lsvClasses_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Not Item Is Nothing Then
    If lsvClasses.Tag <> Item.Key Then SaveProperties lsvClasses.Tag
    lsvClasses.Tag = Item.Key
    SetProperty Item.Tag
  End If
End Sub

Private Sub lsvClasses_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    PopupMenu mnuClasses, , lsvClasses.Left + x, lsvClasses.Top + y
  End If
End Sub

Private Sub mnuAdd_Click()
  On Error Resume Next
  'Save the previous changes
  If Not lsvClasses.SelectedItem Is Nothing Then
    SaveProperties lsvClasses.SelectedItem.Key
  End If
  'Clear the properties
  ClearProperties
  mChange = True
  tabMain.Enabled = True
  'Add new class
  If lsvClasses.ListItems("class " & mClass) Is Nothing Then
    lsvClasses.ListItems.Add , "class " & mClass, ".Class " & mClass, , "CLASS"
  End If
  lsvClasses.ListItems("class " & mClass).Selected = True
  mClass = mClass + 1
  lsvClasses.StartLabelEdit
End Sub

Private Sub mnuNew_Click()
Dim lResult As VbMsgBoxResult
  If mFilechange Or mChange Then
    lResult = MsgBox(IIf(Mfilename <> "", mID(Mfilename, InStrRev(Mfilename, "\") + 1), "New Stylesheet") & vbCrLf & vbCrLf & "has been changed. Do you want to save the changes?", vbQuestion + vbYesNoCancel, Mtitle)
    If lResult = vbCancel Then
      Exit Sub
    ElseIf lResult = vbYes Then
      SaveCSS Mfilename
    End If
  End If
  mClass = 1
  ClearProperties
  lsvClasses.ListItems.Clear
  tabMain.Enabled = False
  Mfilename = ""
  Me.Caption = "New Stylesheet [CSS Editor]"
End Sub

Private Sub mnuOpen_Click()
Dim lResult As VbMsgBoxResult
  If mFilechange Then
    lResult = MsgBox(IIf(Mfilename <> "", mID(Mfilename, InStrRev(Mfilename, "\") + 1), "New Stylesheet") & vbCrLf & vbCrLf & "has been changed. Do you want to save the changes?", vbQuestion + vbYesNoCancel, Mtitle)
    If lResult = vbCancel Then
      Exit Sub
    ElseIf lResult = vbYes Then
      SaveCSS Mfilename
    End If
  End If
  CD.FileName = ""
  CD.CancelError = False
  CD.Filter = "Stylesheets (*.css)|*.css"
  CD.DialogTitle = "Open Stylesheet"
  CD.ShowOpen
  If CD.FileName <> "" Then
    mClass = 1
    LoadCSS CD.FileName
    Me.Caption = mID(CD.FileName, InStrRev(CD.FileName, "\") + 1) & " [CSS Editor]"
  End If
End Sub

Private Sub cColor_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cFontFamily_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub cFontSize_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub picPaleteB_Click()
  imgPaleteB_Click
End Sub

Private Sub picPaleteB_KeyDown(KeyCode As Integer, Shift As Integer)
  mChange = True
End Sub

Private Sub picPaleteF_Click()
  imgPaleteF_Click
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
  ColorPicker.ShowColor = False
End Sub

'User Functions
'
Private Function LoadCSS(ByVal pFilename As String)
'
'Load the css file
'
Dim fn As Integer
Dim lContent As String
Dim lClasses As Variant
Dim lClass As String
Dim lName As Variant
Dim lProperty As String
Dim li As Integer
Dim lj As Integer
Dim lResult As VbMsgBoxResult
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  tabMain.Enabled = False
  lsvClasses.ListItems.Clear
  fn = FreeFile
  Open pFilename For Input As fn
    lContent = Input(LOF(fn), fn)
  Close fn
  lContent = Replace(lContent, vbCrLf, " ")
  lClasses = Split(lContent, "}")
  For li = LBound(lClasses) To UBound(lClasses)
    lClasses(li) = Trim(lClasses(li)) 'Replace(lClasses(li), vbTab, "")
    If lClasses(li) <> "" Then
      If InStr(lClasses(li), "{") > 0 Then
        lClass = mID(lClasses(li), 1, InStr(lClasses(li), "{") - 1)
        lProperty = mID(lClasses(li), InStr(lClasses(li), "{") + 1)
      Else
        lClass = lClasses(li)
        lProperty = ""
      End If
      lClass = Replace(lClass, vbTab, "")
      lProperty = Replace(lProperty, vbTab, "")
      If InStr(lClass, ",") > 0 Then
        lName = Split(lClass, ",")
        For lj = LBound(lName) To UBound(lName)
          lsvClasses.ListItems.Add(, Trim(lName(lj)), Trim(lName(lj)), , "CLASS").Tag = Trim(lProperty) 'Ucase Changed
        Next
      Else
        lsvClasses.ListItems.Add(, Trim(lClass), Trim(lClass), , "CLASS").Tag = Trim(lProperty) 'Ucase Changed
      End If
    End If
  Next
  If lsvClasses.ListItems.Count > 0 Then
    lsvClasses.Refresh
    lsvClasses.Tag = lsvClasses.ListItems(1).Key
    lsvClasses.ListItems(1).Selected = True
    lsvClasses.ListItems(1).EnsureVisible
    tabMain.Enabled = True
  End If
  Mfilename = pFilename
  Screen.MousePointer = vbDefault
End Function

Private Function SetProperty(ByVal pProperty As String)
'
'Set the property to edit
'
Dim lProperties As Variant
Dim lName As String
Dim lValue As String
Dim li As Integer
Dim lColor As Long
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  ClearProperties
  lProperties = Split(pProperty, ";")
  For li = LBound(lProperties) To UBound(lProperties)
    If InStr(lProperties(li), ":") > 0 Then
      lName = Trim(Split(lProperties(li), ":")(0))
      lName = Replace(lName, "-", "")
      lValue = Trim(Split(lProperties(li), ":")(1))
      If lValue <> "" Then
        Me.Controls("c" & lName).Text = lValue
      End If
    End If
  Next
  'set the color
  lColor = ColorPicker.GetLongRGB(cColor.Text)
  picPaleteF.BackColor = IIf(lColor <> -1, lColor, vbButtonFace)
  lColor = ColorPicker.GetLongRGB(cBackgroundColor.Text)
  picPaleteB.BackColor = IIf(lColor <> -1, lColor, vbButtonFace)
  Screen.MousePointer = vbDefault
End Function

Private Function ClearProperties()
'
'Clear the properties before assign
'
Dim lColor As Long
  cFontFamily.Text = ""
  'Color clear
  cColor.Text = ""
  lColor = ColorPicker.GetLongRGB(cColor.Text)
  picPaleteF.BackColor = IIf(lColor <> -1, lColor, vbButtonFace)
  '
  cFontStyle.Text = ""
  cFontVariant.Text = ""
  cFontWeight.Text = ""
  cFontSize.Text = ""
  cTextDecoration.Text = ""
  cTextTransform.Text = ""
  cTextAlign.Text = ""
  cVerticalAlign.Text = ""
  'Background color
  cBackgroundColor.Text = ""
  lColor = ColorPicker.GetLongRGB(cBackgroundColor.Text)
  picPaleteB.BackColor = IIf(lColor <> -1, lColor, vbButtonFace)
  '
  cBackgroundImage.Text = ""
  cBackgroundRepeat.Text = ""
  cBackgroundAttachment.Text = ""
  cBackgroundPosition.Text = ""
  cMarginTop.Text = ""
  cMarginRight.Text = ""
  cMarginBottom.Text = ""
  cMarginLeft.Text = ""
  cMargin.Text = ""
  cPaddingTop.Text = ""
  cPaddingRight.Text = ""
  cPaddingBottom.Text = ""
  cPaddingLeft.Text = ""
  cPadding.Text = ""
  cBorderTopColor.Text = ""
  cBorderRightColor.Text = ""
  cBorderBottomColor.Text = ""
  cBorderLeftColor.Text = ""
  cBorderColor.Text = ""
  cBorderTopWidth.Text = ""
  cBorderRightWidth.Text = ""
  cBorderBottomWidth.Text = ""
  cBorderLeftWidth.Text = ""
  cBorderWidth.Text = ""
  cBorderTopStyle.Text = ""
  cBorderRightStyle.Text = ""
  cBorderBottomStyle.Text = ""
  cBorderLeftStyle.Text = ""
  cBorderStyle.Text = ""
  cBorderTop.Text = ""
  cBorderRight.Text = ""
  cBorderBottom.Text = ""
  cBorderLeft.Text = ""
  cBorder.Text = ""
  cClip.Text = ""
  cHeight.Text = ""
  cLeft.Text = ""
  cOverflow.Text = ""
  cPosition.Text = ""
  cTop.Text = ""
  cVisibility.Text = ""
  cWidth.Text = ""
  cZIndex.Text = ""
  cFloat.Text = ""
  cClear.Text = ""
  cDisplay.Text = ""
  cLineStyleType.Text = ""
  cLineStyleImage.Text = ""
  cLineStylePosition.Text = ""
  cLineStyle.Text = ""
  'mChange = False
  'mFilechange = False
  'mClass = 1
End Function

Private Function SaveProperties(ByVal pKey As String)
'
'Save the properties of changed class
'
Dim Litem As ListItem
Dim lProperties As Variant
Dim lTag As String
Dim li As Integer
Dim lName As String
Dim lControlname As String
Dim lValue As String
  On Error Resume Next
  If mChange Then
    Set Litem = lsvClasses.ListItems(pKey)
    If Not Litem Is Nothing Then
      lTag = ""
      'save the properties not in previous
      For li = LBound(mProperties) To UBound(mProperties)
        If InStr(1, Litem.Tag, mProperties(li), vbTextCompare) = 0 Then
          lControlname = mProperties(li)
          lControlname = Replace(lControlname, "-", "")
          If Trim(Me.Controls("c" & lControlname).Text) <> "" Then
            lTag = lTag & mProperties(li) & ": " & Me.Controls("c" & lControlname).Text & "; "
          End If
        End If
      Next
      'update the propeties if exists
      lProperties = Split(Litem.Tag, ";")
      For li = LBound(lProperties) To UBound(lProperties)
        If InStr(lProperties(li), ":") > 0 Then
          lName = Trim(Split(lProperties(li), ":")(0))
          lControlname = lName
          lControlname = Replace(lControlname, "-", "")
          lValue = ""
          If Trim(Me.Controls("c" & lControlname).Text) <> "" Then
            lValue = Me.Controls("c" & lControlname).Text
          End If
          If lValue <> "" Then
            lTag = lTag & LCase(lName) & ": " & lValue & "; "
          End If
        End If
      Next
      Litem.Tag = lTag
    End If
    mFilechange = True
    mChange = False
  End If
End Function

Private Function SaveCSS(ByVal pFilename As String)
'
'Save the CSS file
'
Dim li As Integer
Dim lContent As String
Dim lClass As String
Dim fn As Integer
  Screen.MousePointer = vbHourglass
  lContent = ""
  For li = 1 To lsvClasses.ListItems.Count
    lClass = lsvClasses.ListItems(li).Text & vbCrLf
    lClass = lClass & vbTab & "{" & vbCrLf & vbTab & vbTab & lsvClasses.ListItems(li).Tag & vbCrLf & vbTab & "}" & vbCrLf
    lContent = lContent & lClass
  Next
  fn = FreeFile
  If Trim(pFilename) = "" Then
    CD.FileName = ""
    CD.Filter = "CSS files (*.css)|*.css"
    CD.DefaultExt = "css"
    CD.CancelError = False
    CD.ShowSave
    If CD.FileName = "" Then Exit Function
    pFilename = CD.FileName
  End If
  Open pFilename For Output As #fn
  Print #fn, lContent
  Close #fn
  mChange = False
  mFilechange = False
  Screen.MousePointer = vbDefault
End Function

Private Function InitProperties()
'
'Load the properties into array
'
  ReDim mProperties(61)
  mProperties(0) = "font-family"
  mProperties(1) = "color"
  mProperties(2) = "font-style"
  mProperties(3) = "font-weight"
  mProperties(4) = "font-variant"
  mProperties(5) = "font-size"
  mProperties(6) = "text-decoration"
  mProperties(7) = "text-transform"
  mProperties(8) = "text-align"
  mProperties(9) = "vertical-align"
  mProperties(10) = "background-color"
  mProperties(11) = "background-image"
  mProperties(12) = "background-repeat"
  mProperties(13) = "background-attachment"
  mProperties(14) = "background-position"
  mProperties(15) = "background"
  mProperties(16) = "margin-top"
  mProperties(17) = "margin-right"
  mProperties(18) = "margin-bottom"
  mProperties(19) = "margin-left"
  mProperties(20) = "margin"
  mProperties(21) = "padding-top"
  mProperties(22) = "padding-right"
  mProperties(23) = "padding-bottom"
  mProperties(24) = "padding-left"
  mProperties(25) = "padding"
  mProperties(26) = "border-top-color"
  mProperties(27) = "border-right-color"
  mProperties(28) = "border-bottom-color"
  mProperties(29) = "border-left-color"
  mProperties(30) = "border-color"
  mProperties(31) = "border-top-width"
  mProperties(32) = "border-right-width"
  mProperties(33) = "border-bottom-width"
  mProperties(34) = "border-left-width"
  mProperties(35) = "border-width"
  mProperties(36) = "border-top-style"
  mProperties(37) = "border-right-style"
  mProperties(38) = "border-bottom-style"
  mProperties(39) = "border-left-style"
  mProperties(40) = "border-style"
  mProperties(41) = "border-top"
  mProperties(42) = "border-right"
  mProperties(43) = "border-bottom"
  mProperties(44) = "border-left"
  mProperties(45) = "border"
  mProperties(46) = "clip"
  mProperties(47) = "height"
  mProperties(48) = "left"
  mProperties(49) = "overflow"
  mProperties(50) = "position"
  mProperties(51) = "top"
  mProperties(52) = "visibility"
  mProperties(53) = "width"
  mProperties(54) = "z-index"
  mProperties(55) = "float"
  mProperties(56) = "clear"
  mProperties(57) = "display"
  mProperties(58) = "line-style-type"
  mProperties(59) = "line-style-image"
  mProperties(60) = "line-style-position"
  mProperties(61) = "line-style"
End Function

Private Sub tmr_Timer()
  mnuRename_Click
  tmr.Enabled = False
End Sub

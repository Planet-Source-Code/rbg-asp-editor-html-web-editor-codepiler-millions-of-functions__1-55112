VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "VBALTB~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Code Piler"
   ClientHeight    =   8760
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10665
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   8385
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18283
            Key             =   "P1"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTools 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7635
      Left            =   0
      ScaleHeight     =   7635
      ScaleWidth      =   3600
      TabIndex        =   2
      Top             =   750
      Width           =   3600
      Begin MSComctlLib.ListView lvPaths 
         Height          =   1125
         Left            =   600
         TabIndex        =   9
         Top             =   5565
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1984
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.DriveListBox drvBox 
         Height          =   315
         Left            =   2700
         TabIndex        =   8
         Top             =   5355
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.DirListBox drBox 
         Height          =   315
         Left            =   2685
         TabIndex        =   7
         Top             =   5700
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.FileListBox flBox 
         Height          =   285
         Left            =   2700
         TabIndex        =   6
         Top             =   6045
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.DirListBox subDr 
         Height          =   315
         Left            =   2700
         TabIndex        =   5
         Top             =   6330
         Visible         =   0   'False
         Width           =   630
      End
      Begin MSComctlLib.TreeView tvFiles 
         Height          =   4170
         Left            =   255
         TabIndex        =   4
         Top             =   405
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   7355
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imlFiles 
         Left            =   2805
         Top             =   3615
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":058A
               Key             =   "ASP"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0664
               Key             =   "DEFAULT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0761
               Key             =   "HTM"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0B87
               Key             =   "HTML"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0FAD
               Key             =   "JS"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1054
               Key             =   "MYCOMPUTER"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":146F
               Key             =   "DISK"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16BC
               Key             =   "SHAREDDISK"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1AC1
               Key             =   "CDROM"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1EEC
               Key             =   "SHAREDCDROM"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2312
               Key             =   "FLOPPY"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2741
               Key             =   "SHAREDFLOPPY"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2B5D
               Key             =   "FOLDERCLOSE"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2CF6
               Key             =   "FOLDEROPEN"
            EndProperty
         EndProperty
      End
      Begin VB.Shape spBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   420
         Left            =   360
         Top             =   4695
         Width           =   795
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   10665
      TabIndex        =   0
      Top             =   0
      Width           =   10665
      Begin VB.ComboBox cboFind 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8235
         TabIndex        =   10
         Top             =   45
         Width           =   1950
      End
      Begin MSComDlg.CommonDialog cdMain 
         Left            =   7290
         Top             =   165
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer tmr 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8835
         Top             =   165
      End
      Begin vbalTBar6.cReBar crbHeader 
         Left            =   3900
         Top             =   0
         _ExtentX        =   5239
         _ExtentY        =   1270
      End
      Begin vbalTBar6.cToolbar tbrMenu 
         Height          =   375
         Left            =   0
         Top             =   405
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   661
      End
      Begin vbalTBar6.cToolbarHost tbhMenu 
         Height          =   345
         Left            =   600
         TabIndex        =   1
         Top             =   75
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   609
         BorderStyle     =   0
      End
      Begin vbalTBar6.cToolbar ctbHeader 
         DragMode        =   1  'Automatic
         Height          =   345
         Left            =   15
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   609
      End
      Begin MSComctlLib.ImageList imlMenu_ 
         Left            =   6750
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   34
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2E7D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2EF3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F7F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3001
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":308B
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3162
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":31F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3261
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":32CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3341
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":33C5
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3467
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":34E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3572
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":35FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3694
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":36F9
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":375F
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":37CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":382D
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":388B
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":38E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":39AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3A27
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3AA1
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3B10
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3B9D
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3C07
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3CF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3D5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3DD4
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3E48
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3EC0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlMenu 
         Left            =   7725
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   36
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3F4B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":405A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":41D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4367
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":43F1
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":457F
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4713
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4810
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":490F
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4A7B
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4C03
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4DA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4EB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":503A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":519D
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5241
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":52A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":530C
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":537B
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":53DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5438
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5496
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5639
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":57C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5856
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58F7
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5984
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5A0B
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5A9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5B36
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5BAB
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5D23
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5E94
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6014
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6124
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":623A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mmupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mmucut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mmucopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mmupaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mmusepo 
         Caption         =   "-"
      End
      Begin VB.Menu mmuselect 
         Caption         =   "&Select All"
      End
   End
   Begin VB.Menu mmupopup2 
      Caption         =   "popup2"
      Visible         =   0   'False
      Begin VB.Menu mmupop2cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mmupop2copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mmupop2paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu pop2sep 
         Caption         =   "-"
      End
      Begin VB.Menu mmupop2select 
         Caption         =   "&Select All"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public frmCollection As Collection
Public mBlankPage As Integer
Public Mdocumentype As Integer '1-Asp,0-HTML
Public WithEvents mStandardMenu As cPopupMenu
Attribute mStandardMenu.VB_VarHelpID = -1
Private WithEvents mTopMenu As cPopupMenu
Attribute mTopMenu.VB_VarHelpID = -1
Private UndoStack As New Collection
Private RedoStack As New Collection
Private bRedoing As Boolean
Private lUndoCount As Long
Private Mpropertyindex As Integer

Private Const cFilters As String = "All Web Pages(*.html *.htm *.asp *.shtml)|*.html;*.htm;*.asp;*.shtml|Asp(*.asp)|*.asp|Htm(*.htm)|*.htm|All Files(*.*)|*.*"

Private Sub ctbHeader_ButtonClick(ByVal lButton As Long)
Dim Lkey As String
  On Error Resume Next
  Lkey = ctbHeader.ButtonKey(lButton)
  Select Case UCase(Lkey)
  Case "NEW"
    mnuFileNew_Click
  Case "SAVE"
    mnuFileSave_Click
  Case "OPEN"
    mnuFileOpen_Click
  Case "CUT"
    mnuEditCut_Click
  Case "COPY"
    mnuEditCopy_Click
  Case "PASTE"
    mnuEditPaste_Click
  Case "REDO"
  Case "UNDO"
  Case "FIND"
    mmufind_Click
  Case "BOLD"
    mnuBold_Click
  Case "ITALICS"
    mnuItalics_Click
  Case "UNDERLINE"
    mnuUnderline_Click
  Case "LEFT"
    mnuLeft_Click
  Case "CENTER"
    mnuCenter_Click
  Case "RIGHT"
    mnuRight_Click
  Case "TABLE"
    mmutable_Click
  Case "IMAGE"
    mmuimage_Click
  Case "LINK"
    mmuaddlink_Click
  Case "REFRESH"
    Me.ActiveForm.ShowPreview
  Case "SPLIT"
    ctbHeader.ButtonVisible("Refresh") = ctbHeader.ButtonChecked("Split")
    Split ctbHeader.ButtonChecked("Split")
  End Select
End Sub

Private Sub MDIForm_Load()
  picTools.Visible = False
  BuildToolBar
  Set frmCollection = New Collection
  LoadDrives
  MDIForm_Resize
  Show
  LoadDocument
  Mpropertyindex = 1
End Sub

Private Sub MDIForm_Resize()
  On Error Resume Next
  tmr.Enabled = True
End Sub

Private Sub mmucopy_Click()
  mnuEditCopy_Click
End Sub

Private Sub mmucut_Click()
  mnuEditCut_Click
End Sub

Private Sub mmupaste_Click()
  mnuEditPaste_Click
End Sub

Private Sub mmuselect_Click()
  mnuEditDSelectAll_Click
End Sub

Private Sub tmr_Timer()
  On Error Resume Next
  crbHeader.RebarSize
  
  spBorder.Left = 0
  spBorder.Top = 0
  spBorder.Width = picTools.Width
  spBorder.Height = picTools.Height
  
  tvFiles.Left = Screen.TwipsPerPixelX * spBorder.BorderWidth
  tvFiles.Width = picTools.Width - (Screen.TwipsPerPixelX * spBorder.BorderWidth * 2)
  tvFiles.Height = picTools.Height - (Screen.TwipsPerPixelY * spBorder.BorderWidth * 2)
  tvFiles.Top = Screen.TwipsPerPixelY * spBorder.BorderWidth
  
  If picTools.Visible = False Then picTools.Visible = True
  tmr.Enabled = False
End Sub

Private Sub mStandardMenu_Click(ItemNumber As Long)
Dim Lkey As String
  Lkey = mStandardMenu.ItemKey(ItemNumber)
  Select Case UCase(Lkey)
  'FILE MENU
  Case "MNUNEW"
    mnuFileNew_Click
  Case "MNUOPEN"
    mnuFileOpen_Click
  Case "MNUSAVE"
    mnuFileSave_Click
  Case "MNUSAVEAS"
    ShowSave
  Case "MNUCLOSE"
    CloseDocument
  Case "MNUCLOSEALL"
    CloseAllDocuments
  Case "MNUPRINT"
    mnuFilePrint_Click
  Case "MNUEXIT"
    mnuFileExit_Click
  'EDIT MENU
  Case "MNUREDO"
    'mmuredo_CLICK
  Case "MNUUNDO"
    'mnuEditUndo_CLICK
  Case "MNUCUT"
    mnuEditCut_Click
  Case "MNUCOPY"
    mnuEditCopy_Click
  Case "MNUPASTE"
    mnuEditPaste_Click
  Case "MNUDELETE"
    mnuDeleteText_Click
  Case "MNUSELECTALL"
    mnuEditDSelectAll_Click
  Case "MNUFIND"
    mmufind_Click
  Case "MNUGOTOLINE"
    mmugoto_Click
  'VIEW
  Case "MNUWEBPREVIEW"
    mnuFilePrintPreview_Click
  Case "MNUWORKSPACE"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    picTools.Visible = mStandardMenu.Checked(ItemNumber)
  'OPTIONS
  Case "MNUWORDWRAP"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    mnuWordwrap_Click mStandardMenu.Checked(ItemNumber)
  'HTML
  Case "MNULINK"
    mmuinslink_Click
  Case "MNUIMAGE"
    mmuimage_Click
  Case "MNUBOOKMARK"
    mmuinsbookmark_Click
  Case "MNUTABLES"
    mmutable_Click
  Case "MNUMARQUEE"
    Clipboard.SetText "<Marquee>&nbsp;</Marquee>"
    mnuEditPaste_Click
  Case "MNUSPAN"
    Clipboard.SetText "<Span></Span>"
    mnuEditPaste_Click
  Case "MNUDIV"
    Clipboard.SetText "<Div></Div>"
    mnuEditPaste_Click
  Case "MNUCLIENT"
    Clipboard.SetText "<Script LANGUAGE=javascript>" & vbCrLf & "<!--" & vbCrLf & vbCrLf & "-->" & vbCrLf & "</Script>"
    mnuEditPaste_Click
  Case "MNUSERVER"
    Clipboard.SetText "<Script LANGUAGE=vbscript RUNAT=Server>" & vbCrLf & vbCrLf & "</Script>"
    mnuEditPaste_Click
  Case "MNUFORM"
    Clipboard.SetText "<Form id=FORM1 name=FORM1 action="""" method=POST>" & vbCrLf & "<P>&nbsp;</P>" & vbCrLf & "</Form>"
    mnuEditPaste_Click
  Case "MNUTEXTBOX"
    Clipboard.SetText "<Input type=TEXT name=TEXT1 value=""TEXT1"">"
    mnuEditPaste_Click
  Case "MNUTEXTAREA"
    Clipboard.SetText "<Inputarea name=TEXTAREA1>Textarea1</Textarea>"
    mnuEditPaste_Click
  Case "MNUOPTIONBUTTON"
    Clipboard.SetText "<Input type=RADIO name=RADIO1 value=""RADIO1"">"
    mnuEditPaste_Click
  Case "MNUPUSHBUTTON"
    Clipboard.SetText "<Input type=BUTTON name=BUTTON1 value=""BUTTON1"">"
    mnuEditPaste_Click
  Case "MNUCHECKBOX"
    Clipboard.SetText "<Input type=CHECKBOX name=CHECKBOX1 value=""CHECKBOX1"" CHECKED>"
    mnuEditPaste_Click
  Case "MNULABEL"
    Clipboard.SetText "<LABEL>test</LABEL>"
    mnuEditPaste_Click
  'FORMAT
  Case "MNUBOLD"
    mnuBold_Click
  Case "MNUITALIC"
    mnuItalics_Click
  Case "MNUUNDERLINE"
    mnuUnderline_Click
  Case "MNUSUPERSCRIPT"
    Clipboard.SetText "<SUP></SUP>"
    mnuEditPaste_Click
  Case "MNUSUBSCRIPT"
    Clipboard.SetText "<SUB></SUB>"
    mnuEditPaste_Click
  Case "MNULEFT"
    mnuLeft_Click
  Case "MNUCENTER"
    mnuCenter_Click
  Case "MNURIGHT"
    mnuRight_Click
  Case "MNUBACKGROUNDCOLOR"
    'mnuBackcolor_CLICK
  'WIZARDS
  Case "MNUDSNCONNECTION"
    mmudbcon_Click
  Case "MNUDBCONNECTION"
    mmuadodatabase_Click
  Case "MNUCOOKIE"
    mmucookiewiz_Click
  
  Case "MNUCASCADE"
    Me.Arrange ChildArrange.Cascade
  Case "MNUTILEVERTICAL"
    Me.Arrange ChildArrange.TileVertical
  Case "MNUTILEHORIZONTAL"
    Me.Arrange ChildArrange.TileHorizontal
  Case "MNUICONS"
    Me.Arrange ChildArrange.Icons
  End Select
End Sub

Private Sub mTopMenu_Click(ItemNumber As Long)
Dim Lkey As String
  Lkey = mTopMenu.ItemKey(ItemNumber)
  mTopMenu.Checked(ItemNumber) = Not mTopMenu.Checked(ItemNumber)
  Select Case UCase(Lkey)
  Case "MNUWORDWRAP"
  Case "MNUSYNTAXCOLORING"
  End Select
End Sub

Private Sub tvFiles_Click()
Dim Lnode As Node
  On Error Resume Next
  If Not tvFiles.SelectedItem Is Nothing Then
    Set Lnode = tvFiles.SelectedItem
    If Lnode.Image <> "ASP" And Lnode.Image <> "HTML" And Lnode.Image <> "HTM" And Lnode.Image <> "JS" And Lnode.Image <> "DEFAULT" Then
      If Lnode.Image = "FOLDERCLOSE" Then Lnode.Image = "FOLDEROPEN"
      LoadFiles Lnode.Key, Lnode.Key
      If Lnode.Image = "DISK" Or Lnode.Image = "CDROM" Or Lnode.Image = "FLOPPY" Then tvFiles.Nodes.Remove Lnode.Key & " "
    End If
  End If
End Sub

Private Sub tvFiles_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Image = "FOLDEROPEN" Then Node.Image = "FOLDERCLOSE"
End Sub

Private Sub tvFiles_DblClick()
Dim Lnode As Node
  If Not tvFiles.SelectedItem Is Nothing Then
    Set Lnode = tvFiles.SelectedItem
    If Lnode.Tag <> "" Then
      If S102_File_Exists(Lnode.Key) = True Then
        If UCase(Mid(Lnode.Key, InStrRev(Lnode.Key, ".") + 1)) = "ASP" Or UCase(Mid(Lnode.Key, InStrRev(Lnode.Key, ".") + 1)) = "HTM" Or UCase(Mid(Lnode.Key, InStrRev(Lnode.Key, ".") + 1)) = "HTML" Or UCase(Mid(Lnode.Key, InStrRev(Lnode.Key, ".") + 1)) = "JS" Then
          LoadDocument Lnode.Key
        End If
      End If
    End If
  End If
End Sub

Private Sub tvFiles_Expand(ByVal Node As MSComctlLib.Node)
  On Error Resume Next
  If Node.Image <> "ASP" And Node.Image <> "HTML" And Node.Image <> "HTM" Or Node.Image <> "JS" Or Node.Image <> "DEFAULT" Then
    If Node.Image = "FOLDERCLOSE" Then Node.Image = "FOLDEROPEN"
    LoadFiles Node.Key, Node.Key
    If Node.Image = "DISK" Or Node.Image = "CDROM" Or Node.Image = "FLOPPY" Then tvFiles.Nodes.Remove Node.Key & " "
  End If
End Sub

Rem ==============================
Rem User functions
Rem ==============================

Private Sub BuildToolBar()
    
    With ctbHeader
        .ImageSource = CTBExternalImageList
        .SetImageList imlMenu, CTBImageListNormal
        .ImageStandardBitmapType = CTBHistorySmallColor
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True, True
        .AddButton "New", 0, , , "", CTBAutoSize, "New"
        .AddButton "Open", 1, , , "", CTBAutoSize, "Open"
        .AddButton "Save", 2, , , "", CTBAutoSize, "Save"
        .AddButton , , , , , CTBSeparator
        .AddButton "Cut", 8, , , "", CTBAutoSize, "Cut"
        .AddButton "Copy", 9, , , "", CTBAutoSize, "Copy"
        .AddButton "Paste", 10, , , "", CTBAutoSize, "Paste"
        .AddButton , , , , , CTBSeparator
        .AddButton "Undo", 7, , , "", CTBAutoSize, "Undo"
        .AddButton "Redo", 6, , , "", CTBAutoSize, "Redo"
        .AddButton , , , , , CTBSeparator
        .AddButton "Find", 11, , , "", CTBAutoSize, "Find"
        .AddControl cboFind.hWnd
        .AddButton , , , , , CTBSeparator
        .AddButton "Bold", 15, , , "", CTBAutoSize, "Bold"
        .AddButton "Italics", 16, , , "", CTBAutoSize, "Italics"
        .AddButton "Underline", 17, , , "", CTBAutoSize, "Underline"
        .AddButton , , , , , CTBSeparator
        .AddButton "Left", 18, , , "", CTBAutoSize, "Left"
        .AddButton "Center", 19, , , "", CTBAutoSize, "Center"
        .AddButton "Right", 20, , , "", CTBAutoSize, "Right"
        .AddButton , , , , , CTBSeparator
        .AddButton "Table", 33, , , "", CTBAutoSize, "Table"
        .AddButton , , , , , CTBSeparator
        .AddButton "Link", 13, , , "", CTBAutoSize, "Link"
        .AddButton "Image", 14, , , "", CTBAutoSize, "Image"
        .AddButton , , , , , CTBSeparator
        .AddButton "Split", 34, , , "", CTBCheck + CTBAutoSize, "Split"
        .AddButton "Refresh", 35, , , "", CTBAutoSize, "Refresh"
        .ButtonVisible("Refresh") = False
        .Visible = True
        .Wrappable = True
    End With
    
    Call pCreateMenu
    
    With tbrMenu
      .CreateFromMenu mStandardMenu
      .Wrappable = True
      .DrawStyle = CTBDrawOfficeXPStyle
      ' the menu toolbar doesn't look good if XP themes are switched on,
      ' so turn them off:
      On Error Resume Next
      SetWindowTheme tbrMenu.hWnd, StrPtr(" "), StrPtr(" ")
      On Error GoTo 0
   End With
   With tbhMenu
      '.ImageSource =  ilstHeader
      .MDIToolbar = True
      .Capture tbrMenu
      .Width = tbhMenu.MDIToolbarMinWidth * Screen.TwipsPerPixelX
   End With
        
    With crbHeader
        'To make the Toolbar visible
        .CreateRebar picTop.hWnd
          
        'Top Menu
        .AddBandByHwnd tbhMenu.hWnd, , , , "MenuBar"
        .BandChildMinWidth(0) = 64
        
        'Toolbar
        .AddBandByHwnd ctbHeader.hWnd, , , , "Toolbar"
        .BandChildMinWidth(crbHeader.BandCount - 1) = 24
        
        .Visible = True
    End With
    
End Sub

Private Sub pCreateMenu()
Dim iP As Long
Dim iP2 As Long
  Set mStandardMenu = New cPopupMenu
  mStandardMenu.ImageList = imlMenu
  mStandardMenu.hWndOwner = picTop.hWnd
  mStandardMenu.OfficeXpStyle = True
    With mStandardMenu
        'Creation of Form Menus
        iP = .AddItem("&File", , , , , , , "mnuFile")
            .AddItem "&New", , , iP, 0, , , "mnuNew"
            .AddItem "&Open", , , iP, 1, , , "mnuOpen"
            .AddItem "-", , , iP
            .AddItem "&Save", , , iP, 2, , , "mnuSave"
            .AddItem "Save As...", , , iP, , , , "mnuSaveAs"
            .AddItem "-", , , iP
            .AddItem "&Close", , , iP, , , , "mnuClose"
            .AddItem "Close All", , , iP, , , , "mnuCloseAll"
            .AddItem "-", , , iP
            .AddItem "&Print...", , , iP, 5, , , "mnuPrint"
            .AddItem "-", , , iP
            .AddItem "E&xit", , , iP, , , , "mnuExit"
        iP = .AddItem("&Edit", , , , , , , "mnuEdit")
            .AddItem "&Redo", , , iP, 6, , , "mnuRedo"
            .AddItem "&Undo", , , iP, 7, , , "mnuUndo"
            .AddItem "-", , , iP
            .AddItem "Cu&t", , , iP, 8, , , "mnuCut"
            .AddItem "&Copy", , , iP, 9, , , "mnuCopy"
            .AddItem "&Paste", , , iP, 10, , , "mnuPaste"
            .AddItem "&Delete", , , iP, , , , "mnuDelete"
            .AddItem "-", , , iP
            .AddItem "Select All", , , iP, , , , "mnuSelectAll"
            .AddItem "-", , , iP
            .AddItem "&Find and Replace...", , , iP, 11, , , "mnuFind"
            .AddItem "&Goto Line...", , , iP, , , , "mnuGotoline"
            .AddItem "-", , , iP
            .AddItem "Word Wrap", , , iP, , True, , "mnuWordwrap"
        iP = .AddItem("&View", , , , , , , "mnuView")
            .AddItem "Preview in Browser", , , iP, , , , "mnuWebPreview"
            .AddItem "Workspace", , , iP, , True, , "mnuWorkspace"
        iP = .AddItem("&HTML", , , , , , , "mnuHTML")
            .AddItem "&Link...", , , iP, 13, , , "mnuLink"
            .AddItem "&Image...", , , iP, 14, , , "mnuImage"
            .AddItem "&Bookmark...", , , iP, , , , "mnuBookMark"
            .AddItem "&Table...", , , iP, 33, , , "mnuTables"
            .AddItem "&Marquee", , , iP, , , , "mnuMarquee"
            .AddItem "&Div", , , iP, , , , "mnuDiv"
            .AddItem "&Span", , , iP, , , , "mnuSpan"
            .AddItem "&Form", , , iP, , , , "mnuForm"
            iP2 = .AddItem("&Input", , , iP, , , , "mnuInput")
                .AddItem "&Text Box", , , iP2, 23, , , "mnuTextbox"
                .AddItem "&Text Area", , , iP2, 24, , , "mnuTextarea"
                .AddItem "&Option Button", , , iP2, 25, , , "mnuOptionbutton"
                .AddItem "&Push Button", , , iP2, 26, , , "mnuPushbutton"
                .AddItem "&Check Box", , , iP2, 27, , , "mnuCheckbox"
                .AddItem "&Label", , , iP2, 29, , , "mnuLabel"
            iP2 = .AddItem("Script Block", , , iP, , , , "mnuScripBlock")
                .AddItem "&Client", , , iP2, , , , "mnuClient"
                .AddItem "&Server", , , iP2, , , , "mnuServer"
            .AddItem "-", , , iP
            iP2 = .AddItem("&Wizards", , , iP, , , , "mnuWizards")
                .AddItem "&DSN Connection", , , iP2, , , , "mnuDSNConnection"
                .AddItem "&Database Connection", , , iP2, , , , "mnuDBConnection"
                .AddItem "&Cookie", , , iP2, , , , "mnuCookie"
        iP = .AddItem("&Format", , , , , , , "mnuFormat")
            .AddItem "&Bold", , , iP, 15, , , "mnuBold"
            .AddItem "&Italic", , , iP, 16, , , "mnuItalic"
            .AddItem "&Underline", , , iP, 17, , , "mnuUnderline"
            .AddItem "&Superscript", , , iP, , , , "mnuSuperscript"
            .AddItem "&Subscript", , , iP, , , , "mnuSubscript"
            .AddItem "-", , , iP
            iP2 = .AddItem("&Align", , , iP, , , , "mnuAlign")
                .AddItem "&Left", , , iP2, 18, , , "mnuLeft"
                .AddItem "&Center", , , iP2, 19, , , "mnuCenter"
                .AddItem "&Right", , , iP2, 20, , , "mnuRight"
            .AddItem "-", , , iP
            .AddItem "&Background Color", , , iP, 21, , , "mnuBackgroundcolor"
        
        iP = .AddItem("Wi&ndows", , , , , , , "mnuWindows")
            .AddItem "&Cascade", , , iP, 30, , , "mnuCascade"
            .AddItem "&Tile Horizontal", , , iP, 31, , , "mnuTileHorizontal"
            .AddItem "Tile &Vertical", , , iP, 32, , , "mnuTileVertical"
            .AddItem "&Icons", , , iP, , , , "mnuIcons"
    End With
    Set mTopMenu = New cPopupMenu
    mTopMenu.ImageList = imlMenu
    mTopMenu.hWndOwner = picTop.hWnd
    mTopMenu.OfficeXpStyle = True
    With mTopMenu
        'creation of PopUpMenus
        'For All PopUp Menus
        
    End With
End Sub

Private Sub mnuWordwrap_Click(ByVal pVal As Boolean)
Dim Li As Integer
  If frmCollection.Count > 0 Then
    For Li = 1 To frmCollection.Count
      If Not frmCollection(Li) Is Nothing Then
        SendMessageLong CLng(frmCollection(Li).RTB.hWnd), EM_SETTARGETDEVICE, &O0, IIf(pVal, &O1, &O0)
      End If
    Next
  End If
End Sub

Private Sub ChangeSyntaxColor(ByVal pValue As Boolean)
Dim Li As Integer
  If frmCollection.Count > 0 Then
    For Li = 1 To frmCollection.Count
      If Not frmCollection(Li) Is Nothing Then
        frmCollection(Li).SyntaxColoring = pValue
      End If
    Next
  End If
End Sub

Private Sub mnuFileNew_Click()
  LoadDocument
End Sub

Private Sub mnuFileOpen_Click()
  cdMain.FileName = ""
  cdMain.CancelError = False
  cdMain.Filter = "HTML Files|*.htm;*.html|ASP files|*.asp|All files|*.*"
  cdMain.ShowOpen
  If cdMain.FileName <> "" Then
    If S102_File_Exists(cdMain.FileName) = True Then
      LoadDocument cdMain.FileName
    End If
  End If
End Sub

Public Sub mnuFileSave_Click()
Dim Lfrmdoc As frmDocument
  Screen.MousePointer = vbHourglass
  Set Lfrmdoc = frmCollection.Item(CStr(Me.ActiveForm.Mkey))
  If Not Lfrmdoc Is Nothing Then
    If Lfrmdoc.Mfilename = "" Then 'New doc
      ShowSave
    Else 'Opened doc
      Lfrmdoc.RTB.SaveFile Lfrmdoc.Mfilename, 1
    End If
    Lfrmdoc.mChange = False
  End If
  Screen.MousePointer = vbDefault
End Sub

Public Sub ShowSave()
Dim sFile As String
Dim Lfrmdoc As frmDocument
  With cdMain
    .DialogTitle = "Save Web Page"
    .Filter = cFilters
    .CancelError = False
    .ShowSave
    If Len(.FileName) = 0 Then
      Exit Sub
    End If
    sFile = .FileName
  End With
  Set Lfrmdoc = frmCollection.Item(Me.ActiveForm.Mkey)
  If Not Lfrmdoc Is Nothing Then
    Lfrmdoc.RTB.SaveFile sFile, 1
    Lfrmdoc.Caption = Mid(sFile, InStrRev(sFile, "\") + 1)
    Lfrmdoc.Mfilename = sFile
    frmCollection.Remove CStr(Me.ActiveForm.Mkey)
    Me.ActiveForm.Mkey = sFile
    frmCollection.Add Me.ActiveForm, sFile
  End If
End Sub

Private Sub mnuFilePrintPreview_Click()
Dim Ret As Long
Dim FileNum As Integer
  On Error Resume Next
  FileNum = FreeFile
  Open App.Path & "\ASPEdit.htm" For Output As #FileNum
  Print #FileNum, frmCollection(Me.ActiveForm.Mkey).RTB.Text
  Close #FileNum
  Ret = ShellExecute(Me.hWnd, "Open", App.Path & "\aspedit.htm", vbNullString, vbNullString, SW_SHOWMAXIMIZED)
End Sub

Private Sub mnuFilePrint_Click()
  cdMain.ShowPrinter
End Sub

Private Sub mnuFileExit_Click()
  'unload the form
  Unload Me
End Sub

Private Sub mnuEditCut_Click()
Dim Lfrmdoc As frmDocument
  On Error Resume Next
  Set Lfrmdoc = frmCollection.Item(CStr(Me.ActiveForm.Mkey))
  If Not Lfrmdoc Is Nothing Then
    Clipboard.SetText Lfrmdoc.RTB.SelText
    Lfrmdoc.RTB.SelText = vbNullString
    Lfrmdoc.mChange = True
  End If
End Sub

Public Sub mnuEditCopy_Click()
Dim Lfrmdoc As frmDocument
  On Error Resume Next
  Set Lfrmdoc = frmCollection.Item(CStr(Me.ActiveForm.Mkey))
  If Not Lfrmdoc Is Nothing Then
    Clipboard.SetText Lfrmdoc.RTB.SelText
  End If
End Sub

Public Sub mnuEditPaste_Click()
Dim Lfrmdoc As frmDocument
  On Error Resume Next
  Set Lfrmdoc = frmCollection.Item(CStr(Me.ActiveForm.Mkey))
  If Not Lfrmdoc Is Nothing Then
    Lfrmdoc.RTB_KeyDown vbKeyV, 2
    Lfrmdoc.mChange = True
  End If
End Sub

Public Sub mnuDeleteText_Click()
Dim Lfrmdoc As frmDocument
  On Error Resume Next
  Set Lfrmdoc = frmCollection.Item(CStr(Me.ActiveForm.Mkey))
  If Not Lfrmdoc Is Nothing Then
    Lfrmdoc.RTB.SelText = vbNullString
    Lfrmdoc.mChange = True
  End If
End Sub

Public Sub mnuEditDSelectAll_Click()
Dim Lfrmdoc As frmDocument
  On Error Resume Next
  Set Lfrmdoc = frmCollection.Item(CStr(Me.ActiveForm.Mkey))
  If Not Lfrmdoc Is Nothing Then
    Lfrmdoc.RTB.SetFocus
    Lfrmdoc.RTB.SelStart = 0
    Lfrmdoc.RTB.SelLength = Len(Lfrmdoc.RTB.Text)
  End If
End Sub

Private Sub mmuaddlink_Click()
  frmLink.Show
End Sub

Private Sub mmuadodatabase_Click()
  frmDB.Show
End Sub

Private Sub mmudbcon_Click()
  frmDSN.Show
End Sub

Private Sub mmufind_Click()
  frmFind.Show
End Sub

Private Sub mmuimage_Click()
  frmImage.Show
End Sub

Private Sub mmuinslink_Click()
  mmuaddlink_Click
End Sub

Private Sub mmuinsbookmark_Click()
  On Error Resume Next
  frmBmark.Show
End Sub

Private Sub mnuBold_Click()
  frmCollection(Me.ActiveForm.Mkey).RTB.SelText = "<B></B>"
End Sub

Private Sub mnuCenter_Click()
  frmCollection(Me.ActiveForm.Mkey).RTB.SelText = "<p align=""center"">  </p>"
End Sub

Private Sub mnuItalics_Click()
  frmCollection(Me.ActiveForm.Mkey).RTB.SelText = "<I></I>"
End Sub

Private Sub mnuLeft_Click()
  frmCollection(Me.ActiveForm.Mkey).RTB.SelText = "<p align=""left"">  </p>"
End Sub

Private Sub mnuRight_Click()
  frmCollection(Me.ActiveForm.Mkey).RTB.SelText = "<p align=""right"">  </p>"
End Sub

Private Sub mnuUnderline_Click()
  frmCollection(Me.ActiveForm.Mkey).RTB.SelText = "<U></U>"
End Sub

Private Sub mmucookiewiz_Click()
  frmCookie.Show
End Sub

Private Sub mmugoto_Click()
  frmGotoline.Show vbModal
End Sub

Private Sub mmutable_Click()
  frmTable.Show
End Sub

Private Function LoadDocument(Optional ByVal pFilename As String)
'
'Load the file; if file is blank then open blank document
'
Dim Lfrmdoc As frmDocument
  On Error Resume Next
  
  Set Lfrmdoc = frmCollection.Item(pFilename)
  If Lfrmdoc Is Nothing Then
    If pFilename <> "" Then
      Screen.MousePointer = vbHourglass
      Set Lfrmdoc = New frmDocument
      If S102_File_Exists(pFilename) Then
        stBar.Panels("P1").Text = "Loading " & Mid(pFilename, InStrRev(pFilename, "\") + 1) & "..."
        Lfrmdoc.Caption = Mid(pFilename, InStrRev(pFilename, "\") + 1)
        Lfrmdoc.Mfilename = pFilename
        Lfrmdoc.Mkey = pFilename
        LoadFile Lfrmdoc.RTB, pFilename
        stBar.Panels("P1").Text = ""
      Else
        MsgBox "File not found!", vbCritical
        Set Lfrmdoc = Nothing
        Exit Function
      End If
    Else
      frmTemplates.Show vbModal
      Screen.MousePointer = vbHourglass
      Set Lfrmdoc = New frmDocument
      mBlankPage = mBlankPage + 1
      Lfrmdoc.Caption = "New Document " & mBlankPage
      Lfrmdoc.Mfilename = ""
      Lfrmdoc.Mkey = Lfrmdoc.hWnd
      Lfrmdoc.LoadNew Mdocumentype
      LoadFile Lfrmdoc.RTB
      pFilename = Lfrmdoc.hWnd
    End If
    If Not Lfrmdoc Is Nothing Then
      frmCollection.Add Lfrmdoc, CStr(pFilename)
    End If
  Else
    Lfrmdoc.ZOrder vbBringToFront
  End If
  Screen.MousePointer = vbDefault
End Function

Private Sub Undo()
Dim cUndo As New clsUndo
  AddToUndoStack cUndo
End Sub

Private Sub AddToUndoStack(cUndo As clsUndo)
  On Error Resume Next
  If UndoStack.Count > 0 Then
    If UndoStack.Count = lUndoCount Then
      UndoStack.Remove (1)
    End If
  End If
  UndoStack.Add cUndo
End Sub

Private Function LoadSettings()
  Mpropnamecolor = GetSetting(App.Title, "Color", "Propname")
  Mpropvalcolor = GetSetting(App.Title, "Color", "Propval")
  Mtagcolor = GetSetting(App.Title, "Color", "Tag")
  Mentitycolor = GetSetting(App.Title, "Color", "Entity")
  Mcommentcolor = GetSetting(App.Title, "Color", "Comment")
  Mpropnamebold = GetSetting(App.Title, "Bold", "Propname")
  Mpropvalbold = GetSetting(App.Title, "Bold", "Propval")
  Mtagbold = GetSetting(App.Title, "Bold", "Tag")
  Mentitybold = GetSetting(App.Title, "Bold", "Entity")
  Mcommentbold = GetSetting(App.Title, "Bold", "Comment")
  Mpropnameitalic = GetSetting(App.Title, "Italic", "Propname")
  Mpropvalitalic = GetSetting(App.Title, "Italic", "Propval")
  Mtagitalic = GetSetting(App.Title, "Italic", "Tag")
  Mentityitalic = GetSetting(App.Title, "Italic", "Entity")
  Mcommentitalic = GetSetting(App.Title, "Italic", "Comment")
  Mfaasp = GetSetting(App.Title, "FileAssociation", "ASP")
  Mfahtm = GetSetting(App.Title, "FileAssociation", "HTM")
  Mfahtml = GetSetting(App.Title, "FileAssociation", "HTML")
  Masppreviewpath = GetSetting(App.Title, "PreviewPath", "ASPPreview", "C:\Inetpath\wwwroot")
  Mlocalhosturl = GetSetting(App.Title, "PreviewPath", "LocalhostURL", "http://localhost")
  Mwebbrowserpath = GetSetting(App.Title, "PreviewPath", "WebBrowserPath")
End Function

Public Sub GotoLine(LineNum As Long, Highlight As Boolean)
Dim temp As Integer
Dim Num As Integer
Dim Pos  As Integer
Dim LastPos As Integer
Dim Cut As Integer
  On Error GoTo Done:
  If LineNum = 0 Then Exit Sub
  Pos = 1
  Num = 1
  temp = 0
  Do
    LastPos = temp
    temp = InStr(Pos, Me.ActiveForm.RTB.Text, vbLf)
    If temp = 0 Then GoTo aredo:
    If temp >= 1 Then
      Num = Num + 1
      Pos = temp + 2
    End If
  Loop Until Num >= LineNum
  Cut = 1
aredo:
  If temp = 0 Then
    LastPos = 0
    temp = Len(Me.ActiveForm.RTB.Text)
    Cut = 0
  End If
  If LineNum = 1 Then
    temp = 0
    LastPos = InStr(1, Me.ActiveForm.RTB.Text, vbLf)
    If LastPos = 0 Then
      LastPos = Len(Me.ActiveForm.RTB.Text)
    End If
    Cut = 0
  End If
  Me.ActiveForm.RTB.SelStart = temp
  If Highlight = True Then Me.ActiveForm.RTB.SelLength = LastPos - Cut
  Me.ActiveForm.RTB.SetFocus
Done:
End Sub

Private Sub LoadDrives()
Dim Li As Integer
  drvBox.Refresh
  tvFiles.ImageList = imlFiles
  tvFiles.Nodes.Clear
  tvFiles.Nodes.Add , , "D0", "My Computer", "MYCOMPUTER"
  For Li = 0 To drvBox.ListCount - 1
    tvFiles.Nodes.Add "D0", tvwChild, drvBox.List(Li), drvBox.List(Li), GetDriveImg(drvBox.List(Li))
    tvFiles.Nodes.Add drvBox.List(Li), tvwChild, drvBox.List(Li) & " ", " ", "FOLDERCLOSE"
  Next
  tvFiles.Nodes("D0").Expanded = True
End Sub

Private Sub LoadFiles(ByVal pPath As String, ByVal pParent As String)
Dim Li As Integer
Dim Lj As Integer
Dim Lparent As String
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  If PathExists(pPath) = False Then
    If InStr(pPath, ":") > 0 And InStr(pPath, "[") > 0 Then
      drBox.Path = Trim(Mid(pPath, 1, InStr(pPath, "[") - 1)) & "\"
    Else
      drBox.Path = pPath & "\"
    End If
    If Err.Number = 0 Then
      For Li = 0 To drBox.ListCount - 1
        tvFiles.Nodes.Add pParent, tvwChild, drBox.List(Li), Mid(drBox.List(Li), InStrRev(drBox.List(Li), "\") + 1), "FOLDERCLOSE"
        Err.Clear
        subDr.Path = drBox.List(Li)
        If Err.Number = 0 Then
          For Lj = 0 To subDr.ListCount - 1
            tvFiles.Nodes.Add drBox.List(Li), tvwChild, subDr.List(Lj), Mid(subDr.List(Lj), InStrRev(subDr.List(Lj), "\") + 1), "FOLDERCLOSE"
          Next
        End If
        Err.Clear
        flBox.Path = drBox.List(Li)
        If Err.Number = 0 Then
          For Lj = 0 To flBox.ListCount - 1
            tvFiles.Nodes.Add(drBox.List(Li), tvwChild, drBox.List(Li) & "\" & flBox.List(Lj), flBox.List(Lj), GetExt(flBox.List(Lj))).Tag = "1"
          Next
        End If
        Err.Clear
      Next
      flBox.Path = drBox.Path
      If Err.Number = 0 Then
        For Li = 0 To flBox.ListCount - 1
          tvFiles.Nodes.Add(pParent, tvwChild, drBox.Path & "\" & flBox.List(Li), flBox.List(Li), GetExt(flBox.List(Li))).Tag = "1"
        Next
      End If
    End If
    tvFiles.Nodes(pParent).EnsureVisible
    lvPaths.ListItems.Add , drBox.Path, drBox.Path
    Err.Clear
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Function GetExt(ByVal pFilename As String) As String
Dim Ltmp As String
  On Error Resume Next
  Ltmp = Mid(pFilename, InStrRev(pFilename, ".") + 1)
  If UCase(Ltmp) = "ASP" Or UCase(Ltmp) = "HTML" Or UCase(Ltmp) = "HTM" Or UCase(Ltmp) = "JS" Then
    GetExt = UCase(Ltmp)
  Else
    GetExt = "DEFAULT"
  End If
End Function

Private Function GetDriveImg(ByVal pDrive As String) As String
Dim Ltmp As String
Dim Lfso As New FileSystemObject
Dim Ldrive As Drive
  On Error Resume Next
  If InStr(pDrive, ":") > 0 And InStr(pDrive, "[") > 0 Then
    pDrive = Trim(Mid(pDrive, 1, InStr(pDrive, "[") - 1)) & "\"
  Else
    pDrive = pDrive & "\"
  End If
  Set Ldrive = Lfso.GetDrive(pDrive)
  If Not Ldrive Is Nothing Then
    If Ldrive.DriveType = CDRom Then
      Ltmp = "CDROM"
    ElseIf Ldrive.DriveType = Removable Then
      Ltmp = "FLOPPY"
    Else
      Ltmp = "DISK"
    End If
  Else
    Ltmp = "DISK"
  End If
  GetDriveImg = Ltmp
End Function

Private Function PathExists(ByVal pPath As String) As Boolean
Dim Litem As ListItem
  On Error Resume Next
  Set Litem = lvPaths.ListItems(pPath)
  PathExists = Not Litem Is Nothing
  Err.Clear
End Function

Private Function CloseAllDocuments()
Dim Li As Integer
  On Error Resume Next
  For Li = frmCollection.Count To 1 Step -1
    Unload frmCollection.Item(Li)
    frmCollection.Remove Li
  Next
End Function

Private Function CloseDocument()
  On Error Resume Next
  frmCollection.Remove Me.ActiveForm.Mkey
  Unload Me.ActiveForm
End Function

Private Function AddFileToMenu(ByVal pFile As String)
Dim Li As Integer
Dim lIndex As Long
Dim Lkey As String
Dim Lfrmdoc As frmDocument
Dim lResult As Boolean
  On Error Resume Next
  For Li = 1 To frmCollection.Count
    Set Lfrmdoc = frmCollection(Li)
    If Not Lfrmdoc Is Nothing Then
      Lkey = Lfrmdoc.Mkey
      mStandardMenu.Checked(mStandardMenu.IndexForKey(Lkey)) = False
      lResult = IIf(Lkey = pFile, True, False)
    End If
  Next
End Function

Private Function Split(Optional ByVal pSplit As Boolean = True)
Dim Li As Integer
  For Li = 1 To frmCollection.Count
    If pSplit Then
      frmCollection.Item(Li).ShowBoth
    Else
      frmCollection.Item(Li).ShowSource
    End If
  Next
End Function

'Private Sub pCreateMenu()
'Dim iP As Long
'Dim iP2 As Long
'  Set mStandardMenu = New cPopupMenu
'  mStandardMenu.ImageList = imlMenu
'  mStandardMenu.hWndOwner = picTop.hWnd
'  mStandardMenu.OfficeXpStyle = True
'    With mStandardMenu
'        'Creation of Form Menus
'        iP = .AddItem("&File", , , , , , , "mnuFile")
'            .AddItem "&New" & vbTab & "Ctrl+N", , , iP, 0, , , "mnuNew"
'            .AddItem "&Open" & vbTab & "Ctrl+O", , , iP, 1, , , "mnuOpen"
'            .AddItem "-", , , iP
'            .AddItem "&Save" & vbTab & "Ctrl+S", , , iP, 2, , , "mnuSave"
'            .AddItem "Save As..." & vbTab & "Ctrl+Shift+S", , , iP, , , , "mnuSaveAs"
'            .AddItem "-", , , iP
'            .AddItem "&Close", , , iP, , , , "mnuClose"
'            .AddItem "Close All", , , iP, , , , "mnuCloseAll"
'            .AddItem "-", , , iP
'            .AddItem "&Print..." & vbTab & "Ctrl+P", , , iP, 5, , , "mnuPrint"
'            .AddItem "-", , , iP
'            .AddItem "E&xit", , , iP, , , , "mnuExit"
'        iP = .AddItem("&Edit", , , , , , , "mnuEdit")
'            .AddItem "&Redo", , , iP, 6, , , "mnuRedo"
'            .AddItem "&Undo", , , iP, 7, , , "mnuUndo"
'            .AddItem "-", , , iP
'            .AddItem "Cu&t" & vbTab & "Ctrl+X", , , iP, 8, , , "mnuCut"
'            .AddItem "&Copy" & vbTab & "Ctrl+C", , , iP, 9, , , "mnuCopy"
'            .AddItem "&Paste" & vbTab & "Ctrl+V", , , iP, 10, , , "mnuPaste"
'            .AddItem "&Delete", , , iP, , , , "mnuDelete"
'            .AddItem "-", , , iP
'            .AddItem "Select All" & vbTab & "Ctrl+A", , , iP, , , , "mnuSelectAll"
'            .AddItem "-", , , iP
'            .AddItem "&Find and Replace..." & vbTab & "Ctrl+F", , , iP, 11, , , "mnuFind"
'            .AddItem "&Goto Line..." & vbTab & "Ctrl+G", , , iP, , , , "mnuGotoline"
'            .AddItem "-", , , iP
'            .AddItem "Word Wrap", , , iP, , True, , "mnuWordwrap"
'        iP = .AddItem("&View", , , , , , , "mnuView")
'            .AddItem "Preview in Browser", , , iP, , , , "mnuWebPreview"
'            .AddItem "Workspace", , , iP, , True, , "mnuWorkspace"
'        iP = .AddItem("&HTML", , , , , , , "mnuHTML")
'            .AddItem "&Link..." & vbTab & "Ctrl+Shift+L", , , iP, 13, , , "mnuLink"
'            .AddItem "&Image..." & vbTab & "Ctrl+M", , , iP, 14, , , "mnuImage"
'            .AddItem "&Bookmark..." & vbTab & "Ctrl+Shift+B", , , iP, , , , "mnuBookMark"
'            .AddItem "&Table...", , , iP, 33, , , "mnuTables"
'            .AddItem "&Marquee", , , iP, , , , "mnuMarquee"
'            .AddItem "&Div", , , iP, , , , "mnuDiv"
'            .AddItem "&Span", , , iP, , , , "mnuSpan"
'            .AddItem "&Form", , , iP, , , , "mnuForm"
'            iP2 = .AddItem("&Input", , , iP, , , , "mnuInput")
'                .AddItem "&Text Box", , , iP2, 23, , , "mnuTextbox"
'                .AddItem "&Text Area", , , iP2, 24, , , "mnuTextarea"
'                .AddItem "&Option Button", , , iP2, 25, , , "mnuOptionbutton"
'                .AddItem "&Push Button", , , iP2, 26, , , "mnuPushbutton"
'                .AddItem "&Check Box", , , iP2, 27, , , "mnuCheckbox"
'                .AddItem "&Label", , , iP2, 29, , , "mnuLabel"
'            iP2 = .AddItem("Script Block", , , iP, , , , "mnuScripBlock")
'                .AddItem "&Client", , , iP2, , , , "mnuClient"
'                .AddItem "&Server", , , iP2, , , , "mnuServer"
'            .AddItem "-", , , iP
'            iP2 = .AddItem("&Wizards", , , iP, , , , "mnuWizards")
'                .AddItem "&ADO Database", , , iP2, , , , "mnuDSNConnection"
'                .AddItem "&DB Connection", , , iP2, , , , "mnuDBConnection"
'                .AddItem "&Cookie", , , iP2, , , , "mnuCookie"
'                .AddItem "&Include File", , , iP2, , , , "mnuIncludeFile"
'        iP = .AddItem("&Format", , , , , , , "mnuFormat")
'            .AddItem "&Bold" & vbTab & "Ctrl+B", , , iP, 15, , , "mnuBold"
'            .AddItem "&Italic" & vbTab & "Ctrl+I", , , iP, 16, , , "mnuItalic"
'            .AddItem "&Underline" & vbTab & "Ctrl+U", , , iP, 17, , , "mnuUnderline"
'            .AddItem "-", , , iP
'            iP2 = .AddItem("&Align", , , iP, , , , "mnuAlign")
'                .AddItem "&Left", , , iP2, 18, , , "mnuLeft"
'                .AddItem "&Center", , , iP2, 19, , , "mnuCenter"
'                .AddItem "&Right", , , iP2, 20, , , "mnuRight"
'            .AddItem "-", , , iP
'            .AddItem "&Background Color", , , iP, 21, , , "mnuBackgroundcolor"
'
'        iP = .AddItem("Wi&ndows", , , , , , , "mnuWindows")
'            .AddItem "&Cascade", , , iP, 30, , , "mnuCascade"
'            .AddItem "&Tile Horizontal", , , iP, 31, , , "mnuTileHorizontal"
'            .AddItem "Tile &Vertical", , , iP, 32, , , "mnuTileVertical"
'            .AddItem "&Icons", , , iP, , , , "mnuIcons"
'    End With
'    Set mTopMenu = New cPopupMenu
'    mTopMenu.ImageList = imlMenu
'    mTopMenu.hWndOwner = picTop.hWnd
'    mTopMenu.OfficeXpStyle = True
'    With mTopMenu
'        'creation of PopUpMenus
'        'For All PopUp Menus
'
'    End With
'End Sub

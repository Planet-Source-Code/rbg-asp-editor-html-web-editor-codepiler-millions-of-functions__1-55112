VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "VBALTB~1.OCX"
Object = "{CA4DCCF5-7118-43AF-B8F9-4A32885010B6}#10.0#0"; "Edt10.ocx"
Begin VB.Form frmEditor 
   Caption         =   "Code Piler"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawMode        =   2  'Blackness
      ForeColor       =   &H80000008&
      Height          =   6720
      Left            =   3315
      ScaleHeight     =   6720
      ScaleWidth      =   60
      TabIndex        =   33
      Top             =   810
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   4560
      Top             =   5490
   End
   Begin MSComctlLib.ProgressBar prgbarMain 
      Height          =   285
      Left            =   5520
      TabIndex        =   32
      Top             =   6735
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   4275
      Top             =   4905
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   90
      ScaleHeight     =   6975
      ScaleWidth      =   3090
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "3090"
      Top             =   735
      Width           =   3090
      Begin VB.PictureBox picApplication 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   45
         ScaleHeight     =   390
         ScaleWidth      =   2910
         TabIndex        =   28
         Top             =   1665
         Width           =   2910
         Begin VB.Timer tmrOpen 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   2295
            Top             =   360
         End
         Begin MSComctlLib.ImageList imlApplication 
            Left            =   2460
            Top             =   510
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   6
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":084A
                  Key             =   "CONNECTION"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":0C8C
                  Key             =   "FIELD"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":0D71
                  Key             =   "TABLES"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":0FEC
                  Key             =   "TABLE"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":1153
                  Key             =   "VIEW"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":11D0
                  Key             =   "TABLEV"
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox picAProperty 
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   3225
            TabIndex        =   29
            Top             =   0
            Width           =   3225
            Begin VB.Label lblApplication 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data Explorer"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   345
               TabIndex        =   30
               Top             =   45
               Width           =   1155
            End
            Begin VB.Shape spBorderA 
               BorderColor     =   &H00C0C0C0&
               Height          =   165
               Left            =   1350
               Top             =   45
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Image imgApplication 
               Height          =   255
               Left            =   30
               Picture         =   "frmEditor.frx":1347
               Top             =   30
               Width           =   240
            End
         End
         Begin MSComctlLib.TreeView tvApplication 
            Height          =   1095
            Left            =   30
            TabIndex        =   31
            Top             =   495
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1931
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   176
            LabelEdit       =   1
            LineStyle       =   1
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
         Begin VB.Shape spBorderAT 
            BorderColor     =   &H00C0C0C0&
            Height          =   510
            Left            =   405
            Top             =   1245
            Width           =   1635
         End
      End
      Begin VB.PictureBox picSite 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   45
         ScaleHeight     =   360
         ScaleWidth      =   2910
         TabIndex        =   22
         Top             =   1275
         Width           =   2910
         Begin MSComctlLib.ListView lsvSitePath 
            Height          =   390
            Left            =   195
            TabIndex        =   0
            Top             =   1290
            Visible         =   0   'False
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   688
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
         Begin VB.ComboBox cboSites 
            Height          =   315
            ItemData        =   "frmEditor.frx":1420
            Left            =   30
            List            =   "frmEditor.frx":1422
            Style           =   2  'Dropdown List
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   375
            Width           =   2415
         End
         Begin VB.PictureBox picSProperty 
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   3225
            TabIndex        =   23
            Top             =   0
            Width           =   3225
            Begin VB.Label lblSite 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Site files"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   345
               TabIndex        =   24
               Top             =   45
               Width           =   720
            End
            Begin VB.Shape spBorderS 
               BorderColor     =   &H00C0C0C0&
               Height          =   165
               Left            =   1350
               Top             =   45
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Image imgSites 
               Height          =   240
               Left            =   30
               Picture         =   "frmEditor.frx":1424
               Top             =   30
               Width           =   240
            End
         End
         Begin MSComctlLib.TreeView tvSiteFiles 
            Height          =   1095
            Left            =   30
            TabIndex        =   25
            Top             =   870
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1931
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   176
            LabelEdit       =   1
            LineStyle       =   1
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
         Begin VB.Shape spBorderST 
            BorderColor     =   &H00C0C0C0&
            Height          =   510
            Left            =   1320
            Top             =   1245
            Width           =   1635
         End
      End
      Begin VB.PictureBox picHistory 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   2910
         TabIndex        =   18
         Top             =   810
         Width           =   2910
         Begin VB.PictureBox picHProperty 
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   3225
            TabIndex        =   19
            Top             =   0
            Width           =   3225
            Begin VB.Image imgHistory 
               Height          =   240
               Left            =   30
               Picture         =   "frmEditor.frx":15AD
               Top             =   30
               Width           =   240
            End
            Begin VB.Shape spBorderH 
               BorderColor     =   &H00C0C0C0&
               Height          =   165
               Left            =   1350
               Top             =   45
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblHistory 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "History"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   345
               TabIndex        =   20
               Top             =   45
               Width           =   615
            End
         End
         Begin MSComctlLib.TreeView tvHistory 
            Height          =   1095
            Left            =   30
            TabIndex        =   21
            Top             =   495
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1931
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   176
            LabelEdit       =   1
            LineStyle       =   1
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
         Begin VB.Shape spBorderHT 
            BorderColor     =   &H00C0C0C0&
            Height          =   510
            Left            =   405
            Top             =   1245
            Width           =   1635
         End
      End
      Begin VB.PictureBox picToolbox 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   15
         ScaleHeight     =   270
         ScaleWidth      =   3180
         TabIndex        =   15
         Top             =   0
         Width           =   3180
         Begin VB.Label lblTool 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Toolbox"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   105
            TabIndex        =   16
            Top             =   30
            Width           =   675
         End
      End
      Begin VB.PictureBox picWorkSpace 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   0
         ScaleHeight     =   390
         ScaleWidth      =   3285
         TabIndex        =   7
         Top             =   405
         Width           =   3285
         Begin VB.PictureBox picWSProperty 
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   3225
            TabIndex        =   14
            Top             =   0
            Width           =   3225
            Begin VB.Image imgWorkspace 
               Height          =   240
               Left            =   60
               Picture         =   "frmEditor.frx":19E1
               Top             =   45
               Width           =   240
            End
            Begin VB.Label lblWorkspace 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Workspace"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   375
               TabIndex        =   17
               Top             =   45
               Width           =   945
            End
            Begin VB.Shape spBorderWS 
               BorderColor     =   &H00C0C0C0&
               Height          =   165
               Left            =   1350
               Top             =   45
               Visible         =   0   'False
               Width           =   675
            End
         End
         Begin VB.DriveListBox drvBox 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            TabIndex        =   11
            Top             =   1005
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.DirListBox drBox 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2385
            TabIndex        =   10
            Top             =   1350
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.FileListBox flBox 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   9
            Top             =   1695
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.DirListBox subDr 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1740
            TabIndex        =   8
            Top             =   1935
            Visible         =   0   'False
            Width           =   630
         End
         Begin MSComctlLib.ListView lvPaths 
            Height          =   390
            Left            =   30
            TabIndex        =   12
            Top             =   1515
            Visible         =   0   'False
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   688
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
         Begin MSComctlLib.TreeView tvFiles 
            Height          =   1095
            Left            =   45
            TabIndex        =   13
            Top             =   375
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1931
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   176
            LabelEdit       =   1
            LineStyle       =   1
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
            Left            =   2385
            Top             =   405
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   43
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":1E02
                  Key             =   "MYCOMPUTER"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":2233
                  Key             =   "FOLDERCLOSE"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":23AE
                  Key             =   "FOLDERCLOSEONLINE"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":27B6
                  Key             =   "FOLDEROPEN"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":29F7
                  Key             =   "FOLDEROPENONLINE"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":2E07
                  Key             =   "DES"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":309E
                  Key             =   "HISTORY"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":34E2
                  Key             =   "HTODAY"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":3772
                  Key             =   "HPAST"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":3A04
                  Key             =   "EMPTY"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":3C7C
                  Key             =   "ASP"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":3ED9
                  Key             =   "JS"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":4147
                  Key             =   "HTM"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":43A4
                  Key             =   "HTML"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":4601
                  Key             =   "TXT"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":4874
                  Key             =   "DOC"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":4AE7
                  Key             =   "INI"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":4D5A
                  Key             =   "BAT"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":4FCD
                  Key             =   "DAT"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":5240
                  Key             =   "WAV"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":53CC
                  Key             =   "MP3"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":5558
                  Key             =   "MPG"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":56E4
                  Key             =   "AVI"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":5870
                  Key             =   "MPEG"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":59FC
                  Key             =   "CDA"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":5B88
                  Key             =   "DEFAULT"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":5E14
                  Key             =   "CSS"
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":5F9B
                  Key             =   "GIF"
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":620F
                  Key             =   "JPG"
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":6483
                  Key             =   "BMP"
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":66F7
                  Key             =   "PNG"
               EndProperty
               BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":696B
                  Key             =   "ICO"
               EndProperty
               BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":6BDF
                  Key             =   "TIF"
               EndProperty
               BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":6E53
                  Key             =   "JPEG"
               EndProperty
               BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":70C7
                  Key             =   "TIFF"
               EndProperty
               BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":733B
                  Key             =   "PSD"
               EndProperty
               BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":75AF
                  Key             =   "MYN"
               EndProperty
               BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":774F
                  Key             =   "MYD"
               EndProperty
               BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":7B71
                  Key             =   "FLOPPY"
               EndProperty
               BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":7F95
                  Key             =   "DISK"
               EndProperty
               BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":81D9
                  Key             =   "CDROM"
               EndProperty
               BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":846C
                  Key             =   "DRAG"
               EndProperty
               BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditor.frx":85CE
                  Key             =   "DRAGCOPY"
               EndProperty
            EndProperty
         End
         Begin VB.Shape spBorder 
            BorderColor     =   &H00C0C0C0&
            Height          =   315
            Left            =   45
            Top             =   1935
            Width           =   1635
         End
      End
   End
   Begin VB.PictureBox picEditor 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   3885
      ScaleHeight     =   3255
      ScaleWidth      =   4485
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1110
      Width           =   4545
      Begin vbalDTab6.vbalDTabControl tabMain 
         Height          =   1500
         Left            =   255
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1125
         Visible         =   0   'False
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   2646
         TabAlign        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Begin Editors.Editor RTB 
            Height          =   930
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   375
            Visible         =   0   'False
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   1640
         End
      End
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8835
      Top             =   165
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   8235
      TabIndex        =   3
      Top             =   45
      Width           =   1950
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7665
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6668
            Key             =   "P1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "P2"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   "P3"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Key             =   "PRow"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Key             =   "PCol"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "INS"
            Key             =   "PIns"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "CAPS"
            Key             =   "PCaps"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   7290
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin vbalTBar6.cReBar crbHeader 
      Left            =   3900
      Top             =   15
      _ExtentX        =   5239
      _ExtentY        =   1270
   End
   Begin vbalTBar6.cToolbar tbrMenu 
      Height          =   375
      Left            =   -15
      Top             =   375
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   661
   End
   Begin vbalTBar6.cToolbarHost tbhMenu 
      Height          =   345
      Left            =   600
      TabIndex        =   4
      Top             =   75
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   609
      BorderStyle     =   0
   End
   Begin vbalTBar6.cToolbar ctbHeader 
      DragMode        =   1  'Automatic
      Height          =   345
      Left            =   30
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
            Picture         =   "frmEditor.frx":8730
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":87A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8832
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":88B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":893E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8A15
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8B7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8D1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8D95
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8E25
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8EAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8F47
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":8FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9012
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9081
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":90E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":913E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":919C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":925E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":92DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9354
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":93C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9450
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":94BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9533
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":95A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":960F
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9687
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":96FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9773
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMenu 
      Left            =   7695
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   48
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":97FE
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9983
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9BC4
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":9E37
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":A098
            Key             =   "UNDO"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":A20C
            Key             =   "REDO"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":A379
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":A4E5
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":A66D
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":A80A
            Key             =   "SEARCH"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":AA71
            Key             =   "LINK"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":ACC5
            Key             =   "IMAGE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":B0DC
            Key             =   "BOLD"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":B278
            Key             =   "ITALIC"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":B4E4
            Key             =   "UNDERLINE"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":B754
            Key             =   "LEFT"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":B8E3
            Key             =   "CENTER"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":BA76
            Key             =   "RIGHT"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":BC07
            Key             =   "TEXTBOX"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":BC97
            Key             =   "TEXTAREA"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":BD38
            Key             =   "OPTION"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":BDC5
            Key             =   "BUTTON"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":BE4C
            Key             =   "CHECK"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":BEDF
            Key             =   "COMBO"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":BF77
            Key             =   "LABEL"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":BFEC
            Key             =   "TABLE"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":C153
            Key             =   "SQLCON"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":C2DA
            Key             =   "DBCON"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":C451
            Key             =   "TOOLBOX"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":C6B0
            Key             =   "BOOKMARK"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":CB73
            Key             =   "JUSTIFY"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":CD08
            Key             =   "INDENT"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":CF76
            Key             =   "OUTDENT"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":D1E6
            Key             =   "HISTORY"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":D62A
            Key             =   "DATAEXPLORER"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":D713
            Key             =   "CLEARHISTORY"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":DB58
            Key             =   "SITE"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":DCF1
            Key             =   "SUBMIT"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":DDAA
            Key             =   "RESET"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":DE44
            Key             =   "HIDDEN"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":DEC3
            Key             =   "REFRESHTOOLBOX"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":DFD7
            Key             =   "FULLMODE"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":E0C6
            Key             =   "BROWSERPREVIEW"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":E340
            Key             =   "CODE"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":E4C0
            Key             =   "VIEW"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":E70E
            Key             =   "CODEVIEW"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":E980
            Key             =   "FONT"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":EC15
            Key             =   "COOKIES"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlFile_Backup 
      Left            =   9180
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":ED12
            Key             =   "ASP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":EDEC
            Key             =   "DEFAULT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":EEE9
            Key             =   "HTM"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":F30F
            Key             =   "HTML"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":F735
            Key             =   "JS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":F7DC
            Key             =   "MYCOMPUTER"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":FBF7
            Key             =   "DISK"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":FE44
            Key             =   "SHAREDDISK"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":10249
            Key             =   "CDROM"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":10674
            Key             =   "SHAREDCDROM"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":10A9A
            Key             =   "FLOPPY"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":10EC9
            Key             =   "SHAREDFLOPPY"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":112E5
            Key             =   "FOLDERCLOSE"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":113AE
            Key             =   "FOLDEROPEN"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":114B6
            Key             =   "MYD"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":1165F
            Key             =   "MYN"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":11AA9
            Key             =   "DES"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":11D40
            Key             =   "HISTORY"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":12184
            Key             =   "HTODAY"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":12414
            Key             =   "HPAST"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":126A6
            Key             =   "EMPTY"
         EndProperty
      EndProperty
   End
   Begin VB.Shape shResizeBorder 
      BorderColor     =   &H00C0C0C0&
      Height          =   360
      Left            =   3495
      Top             =   840
      Width           =   195
   End
   Begin VB.Image imgSplit 
      Height          =   945
      Left            =   3390
      MouseIcon       =   "frmEditor.frx":1291E
      MousePointer    =   99  'Custom
      Picture         =   "frmEditor.frx":131E8
      Top             =   4995
      Width           =   105
   End
   Begin VB.Image imgTest 
      Height          =   315
      Left            =   7350
      Top             =   5190
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgResize 
      Height          =   6705
      Left            =   3240
      MousePointer    =   9  'Size W E
      Stretch         =   -1  'True
      Top             =   795
      Width           =   30
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuAddFolder 
         Caption         =   "Add Folder"
      End
      Begin VB.Menu mnuAddFile 
         Caption         =   "AddFile"
      End
      Begin VB.Menu mnuhy3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu myhy2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewsite 
         Caption         =   "New Site"
      End
      Begin VB.Menu mnuRenameSite 
         Caption         =   "Edit Site"
      End
      Begin VB.Menu mnuRemoveSite 
         Caption         =   "Delete Site"
      End
   End
   Begin VB.Menu mnuApplication 
      Caption         =   "Application"
      Visible         =   0   'False
      Begin VB.Menu mnuNewConnection 
         Caption         =   "New Connection..."
      End
      Begin VB.Menu mnuhy1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditConnection 
         Caption         =   "Edit Connection..."
      End
      Begin VB.Menu mnuDeleteConnection 
         Caption         =   "Delete Connection"
      End
      Begin VB.Menu mnuTestConnection 
         Caption         =   "Test Connection"
      End
      Begin VB.Menu mnuhy2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertCode 
         Caption         =   "Insert Code"
      End
      Begin VB.Menu mnuViewRecords 
         Caption         =   "View Records"
      End
      Begin VB.Menu mnuHy4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenerateCode 
         Caption         =   "Generate Script to Code Window"
         Begin VB.Menu mnuSelectGC 
            Caption         =   "Select"
         End
         Begin VB.Menu mnuInsertGC 
            Caption         =   "Insert"
         End
         Begin VB.Menu mnuUpdateGC 
            Caption         =   "Update"
         End
         Begin VB.Menu mnuDeleteGC 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu mnuGenerateFile 
         Caption         =   "Generate Script to File"
         Begin VB.Menu mnuSelectGF 
            Caption         =   "Select"
         End
         Begin VB.Menu mnuInsertGF 
            Caption         =   "Insert"
         End
         Begin VB.Menu mnuUpdateGF 
            Caption         =   "Update"
         End
         Begin VB.Menu mnuDeleteGF 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu mnuGenerateClipboard 
         Caption         =   "Generate Script to Clipboard"
         Begin VB.Menu mnuSelectGP 
            Caption         =   "Select"
         End
         Begin VB.Menu mnuInsertGP 
            Caption         =   "Insert"
         End
         Begin VB.Menu mnuUpdateGP 
            Caption         =   "Update"
         End
         Begin VB.Menu mnuDeleteGP 
            Caption         =   "Delete"
         End
      End
   End
   Begin VB.Menu mnuHistory 
      Caption         =   "History"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveHistory 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuHy5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearhistory 
         Caption         =   "Clear History"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'format:
Rem --------------------------------
Rem   Property bar colors(toolbox)
Rem --------------------------------
Rem    Border - rgb(220,217,200)
Rem    Mouseover - rgb(172,168,153)
Rem --------------------------------

Rem --------------------------------
Rem     Format of site details
Rem --------------------------------
Rem   Sitename^path^url^connectionname~connectionstring1^connectionname~connectionstring2^...
Rem --------------------------------

Rem --------------------------------
Rem       Public declarations
Rem --------------------------------
Public Mdocumentype As Integer '1-Asp,0-HTML,(-1)-Cancel
Public Mrecentfile As String 'Recent filename which is opened from the templates
Public Mopendialog As Boolean 'Flag for dont show the template at startup

Rem --------------------------------
Rem       Private declarations
Rem --------------------------------
Private WithEvents mFTP As clsFTP
Attribute mFTP.VB_VarHelpID = -1
Private WithEvents mStandardMenu As cPopupMenu 'Menu
Attribute mStandardMenu.VB_VarHelpID = -1
Private WithEvents mSiteFrm As frmSiteDetails 'Site detail form for multiuse
Attribute mSiteFrm.VB_VarHelpID = -1
Private mBlankPage As Integer 'Count for new documents
Private mEditorIndex As Integer 'Index of selected rtb editor
Private mIndex As Integer 'Index of selected panel in toolbox
Private mEnable As Boolean 'To avoid the menu enable frequently
Private mSitename As String 'Sitename for set to combo after edit/remove
Private mDrag As Boolean 'Drag Flag
Private mSource As String 'Source file/folder to copy when drag drop
Private mSourceType As String 'Source type(file/folder) to copy when drag drop
Private mCopy As Boolean 'Source is to copy/move
Private mServer As String 'ftp server
Private mPassword As String 'ftp password
Private mUsername As String 'ftp username
Private mPort As Integer 'ftp port
Private mLoading As Boolean 'for loading completed; to avoid the raising statechanged event frequently

Rem --------------------------------
Rem           Properties
Rem --------------------------------
Private mWordWrap As Boolean 'Wordwrap setting
Private mLineNo As Boolean 'Show lineno setting
Private mToolbox As Boolean 'Show toolbox setting
Private mToolboxSize As Single 'Show toolbox size
Private mAutoCompletion As Boolean 'Autocompletion setting
Private mSyntaxHighlighting As Boolean 'Syntaxhighlighting setting
Private mIntelisense As Boolean 'Show intelisense setting
Private mFullmodePreview As Boolean 'View preview in fullmode
Private mSite As Integer 'Last visited site

Rem --------------------------------
Rem          Action Events
Rem --------------------------------

Private Sub cboFind_KeyPress(KeyAscii As Integer)
Dim lResult As Boolean
Dim li As Integer
  If KeyAscii = vbKeyReturn Then
    For li = 0 To cboFind.ListCount - 1
      If LCase(cboFind.List(li)) = LCase(cboFind.Text) Then
        lResult = True
      End If
    Next
    If lResult = False Then cboFind.AddItem cboFind.Text, 0
    If cboFind.ListCount > 5 Then cboFind.RemoveItem 10 'storage range is 5
    frmFind.txtFindwhat.Text = cboFind.Text
    frmFind.FindNextButton_Click
    KeyAscii = 0
  End If
End Sub

Private Sub cboSites_Click()
Dim lSite As clsSite
Dim lConstr As String
Dim li As Long
Dim LConArr
Dim Lconname As String
Dim Lconstring As String

  On Error Resume Next
  If cboSites.Text = "Define sites..." Then
    mnuNewsite_Click
  ElseIf cboSites.Text = "---------------" Then
   If cboSites.ListIndex <> 0 Then
    cboSites.ListIndex = 0
   Else
    mSitename = ""
    tvSiteFiles.Nodes.Clear
    tvApplication.Nodes.Clear
   End If
  Else
    cboSites.Tag = cboSites.ListIndex
    If cboSites.Text <> tvSiteFiles.Tag Then
      Set lSite = Msitedetails.Item(cboSites.Text)
      mServer = lSite.Server
      mUsername = lSite.Username
      mPassword = lSite.Password
      mPort = lSite.Port
      If lSite.UseFTP Then 'For remote folders
        Screen.MousePointer = vbHourglass
        tvSiteFiles.Nodes.Clear
        tvSiteFiles.Nodes.Add(, , "F0", lSite.Server, "FOLDEROPENONLINE").Tag = "R" 'for remote
        LoadRemoteFiles "//", "F0"
        tvSiteFiles.Nodes("F0").Expanded = True
        Screen.MousePointer = vbDefault
      Else 'For local folders
        LoadSiteFiles
        If tvSiteFiles.Nodes.Count = 0 Then tvSiteFiles.Nodes.Add(, , "F0", "  (No files)", "FOLDERCLOSE").ForeColor = vbGrayText
      End If
      tvApplication.Nodes.Clear
      If Not lSite.ConnectionString Is Nothing Then
        For li = 1 To lSite.ConnectionString.Count
          lConstr = lSite.ConnectionString(li)
          Lconname = Split(lConstr, "~")(0)
          Lconstring = Split(lConstr, "~")(1)
          If Lconname <> "" And Lconstring <> "" Then
              tvApplication.Nodes.Add(, , Lconname, Lconname, "CONNECTION").Tag = Lconstring
              LoadConnection Lconname, Lconstring, False
          End If
        Next
      End If
      tvSiteFiles.Tag = cboSites.Text
    End If
  End If
  
  Err.Clear
End Sub

Private Sub ctbHeader_ButtonClick(ByVal lButton As Long)
Dim lKey As String
Dim lSite As clsSite
  On Error Resume Next
  lKey = ctbHeader.ButtonKey(lButton)
  Select Case UCase(lKey)
  Case "NEW"
    Set lSite = Msitedetails.Item(cboSites.Text)
    If lSite Is Nothing Then
      LoadDocument
    Else
      LoadDocument , , lSite.LocalPath, , lSite.URL
    End If
  Case "SAVE"
    SaveDocument
  Case "OPEN"
    OpenDocument
  Case "CUT"
    RTB(mEditorIndex).Cut
    ctbHeader.ButtonEnabled("Paste") = True
    mStandardMenu.Enabled(mStandardMenu.IndexForKey("mnuPaste")) = True
  Case "COPY"
    RTB(mEditorIndex).Copy
    ctbHeader.ButtonEnabled("Paste") = True
    mStandardMenu.Enabled(mStandardMenu.IndexForKey("mnuPaste")) = True
  Case "PASTE"
    RTB(mEditorIndex).Paste
  Case "REDO"
    RTB(mEditorIndex).Redo
  Case "UNDO"
    RTB(mEditorIndex).Undo
  Case "FIND"
    SetParent frmFind.hwnd, Me.hwnd
    If RTB(mEditorIndex).SelText <> "" Then
      frmFind.txtFindwhat.Text = RTB(mEditorIndex).SelText
    End If
    frmFind.Show
  Case "BOLD"
    RTB(mEditorIndex).Paste "<B>" & RTB(mEditorIndex).SelText & "</B>"
  Case "ITALICS"
    RTB(mEditorIndex).Paste "<I>" & RTB(mEditorIndex).SelText & "</I>"
  Case "UNDERLINE"
    RTB(mEditorIndex).Paste "<U>" & RTB(mEditorIndex).SelText & "</U>"
  Case "LEFT"
    RTB(mEditorIndex).Paste "<p align=""left"">" & RTB(mEditorIndex).SelText & "</p>"
  Case "CENTER"
    RTB(mEditorIndex).Paste "<p align=""center"">" & RTB(mEditorIndex).SelText & "</p>"
  Case "RIGHT"
    RTB(mEditorIndex).Paste "<p align=""right"">" & RTB(mEditorIndex).SelText & "</p>"
  Case "INDENT"
    RTB(mEditorIndex).Indent
  Case "UNINDENT"
    RTB(mEditorIndex).Outdent
  Case "TABLE"
    frmTable.Show vbModal
  Case "IMAGE"
    frmImage.Show vbModal
  Case "LINK"
    frmLink.Show vbModal
  End Select
  
  Err.Clear
End Sub

Private Sub Form_Load()
Dim lCommand As String
Dim lWidth As Single
Dim lCode As Integer

  S110_WriteLog "Main form starting...", True
  
  'Splash screen and verify registration
  Load frmSplash
  If frmSplash.CheckRegistration = False Then
    Unload frmSplash
    End
  End If
  
  'Load settings
  LoadSettings
  
  Me.Caption = Mtitle
  
  'Loading forms
  Load frmTemplates
  Load frmFonts
  
  'Set ftp
  Set mFTP = New clsFTP
  
  'Initialise the image list
  tvFiles.ImageList = imlFiles
  tvHistory.ImageList = imlFiles
  tvSiteFiles.ImageList = imlFiles
  tvApplication.ImageList = imlApplication
  
  'Load files/history/sites/connection
  LoadDrives
  LoadHistory
  LoadSites
  mIndex = IIf(cboSites.ListCount > 1, 3, 1)
  If mToolbox Then picTools.Visible = True
  
  'Load menus/toolbar
  BuildToolBar
  EnableMenuMM False
  
  'Form actions
  Form_Resize
  Show
  
  Unload frmSplash
  
  'Command line arguments
  lCommand = Command
  If lCommand <> "" Then
    lCommand = Replace(lCommand, """", "")
    LoadDocument lCommand
  Else
    ctbHeader_ButtonClick ctbHeader.ButtonIndex("New")
  End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  crbHeader.RebarSize
  
  picTools.Left = 0
  picTools.Top = tbrMenu.Height + ctbHeader.Height ' + Screen.TwipsPerPixelY * 5 'Screen.TwipsPerPixelY * 40
  picTools.Height = Me.ScaleHeight - picTools.Top - stBar.Height + Screen.TwipsPerPixelY * 2
  
  'picDrag.Width = picTools.Width
  'picDrag.Height = picTools.Height - picTools.Height / 3
  'picDrag.Top = picTools.Top
  
  picToolbox.Left = Screen.TwipsPerPixelX
  picToolbox.Width = picTools.Width - picToolbox.Left
  picToolbox.Top = Screen.TwipsPerPixelY
  
  CollapseTool mIndex
  
  imgSplit.Left = picTools.Left + picTools.Width
  imgSplit.Top = picTools.Top + ((picTools.Height / 2) - (imgSplit.Height / 2))
  
  imgResize.Left = picTools.Left + picTools.Width
  imgResize.Top = picTools.Top
  imgResize.Width = imgSplit.Width
  imgResize.Height = picTools.Height
  
  picResize.Left = imgResize.Left
  picResize.Height = imgResize.Height
  picResize.Top = imgResize.Top
  picResize.Width = imgSplit.Width
  
  shResizeBorder.Top = imgResize.Top
  shResizeBorder.Height = imgResize.Height - Screen.TwipsPerPixelY * 2
  shResizeBorder.Left = imgResize.Left
  shResizeBorder.Width = imgResize.Width
  
  picEditor.Left = imgResize.Left + imgResize.Width  'Screen.TwipsPerPixelX * 2
  picEditor.Top = picTools.Top
  picEditor.Height = picTools.Height
  picEditor.Width = Me.ScaleWidth - picEditor.Left
  
  tabMain.Left = 0
  tabMain.Top = 0
  tabMain.Width = picEditor.Width - Screen.TwipsPerPixelX * 5
  tabMain.Height = picEditor.Height - Screen.TwipsPerPixelY * 5
  
  tmrResize.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim li As Integer
Dim lResult As VbMsgBoxResult
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  For li = 1 To RTB.Count - 1
    If RTB(li).Changed Then
      lResult = MsgBox(IIf(RTB(li).FileName = "", RTB(li).Key, RTB(li).FileName) & vbCrLf & vbCrLf & "The above document has been changed. Do you want to save changes?", vbYesNoCancel + vbQuestion, Mtitle)
      If lResult = vbYes Then
        If RTB(li).FileName = "" Then
          RTB(li).SaveAsFile
        Else
          RTB(li).SaveFile RTB(li).FileName
        End If
      ElseIf lResult = vbCancel Then
        Cancel = 1
        Screen.MousePointer = vbDefault
        Exit Sub
      End If
      RTB(li).Changed = False
    End If
  Next
  Set mSiteFrm = Nothing
  Set Msitedetails = Nothing
  mFTP.CloseConnection
  Set mFTP = Nothing
  crbHeader.RemoveAllRebarBands
  Unload frmFind
  Unload frmFonts
  Unload frmTemplates
  S105_Delete App.Path & "\Temp", True
  SaveSettings
  Screen.MousePointer = vbDefault
  'TerminateProcess GetCurrentProcess(), 0& 'Raga
  'End
End Sub


Private Sub imgApplication_Click()
  picAProperty_Click
End Sub

Private Sub imgApplication_DblClick()
  picAProperty_DblClick
End Sub

Private Sub imgHistory_Click()
  picHProperty_Click
End Sub

Private Sub imgHistory_DblClick()
  picHProperty_DblClick
End Sub

Private Sub imgSites_Click()
  picSProperty_Click
End Sub

Private Sub imgSites_DblClick()
  picSProperty_DblClick
End Sub

Private Sub imgSplit_Click()
  mStandardMenu_Click mStandardMenu.IndexForKey("mnuToolbox")
End Sub

Private Sub imgWorkspace_Click()
  picWSProperty_Click
End Sub

Private Sub imgWorkspace_DblClick()
  picWSProperty_DblClick
End Sub

Private Sub lblApplication_Click()
  picAProperty_Click
End Sub

Private Sub lblApplication_DblClick()
  picAProperty_DblClick
End Sub

Private Sub lblHistory_Click()
  picHProperty_Click
End Sub

Private Sub lblHistory_DblClick()
  picHProperty_DblClick
End Sub

Private Sub lblSite_Click()
  picSProperty_Click
End Sub

Private Sub lblSite_DblClick()
  picSProperty_DblClick
End Sub

Private Sub lblWorkspace_Click()
  picWSProperty_Click
End Sub

Private Sub imgResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If picTools.Visible Then
    picResize.Left = imgResize.Left
    picResize.Visible = True
  End If
End Sub

Private Sub imgResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lPos As Integer
Dim Lrange As Integer
  If picResize.Visible = True Then
    lPos = picTools.Width + x
    Lrange = (picEditor.Left + picEditor.Width) - val(picTools.Tag)
    If lPos > val(picTools.Tag) And lPos < Lrange Then
      picResize.Left = (picTools.Width + x) - (picResize.Width / 2)
    End If
  End If
End Sub

Private Sub imgResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If picResize.Visible = True Then
    picResize.Visible = False
    picTools.Width = picResize.Left
    mToolboxSize = picTools.Width
    Form_Resize
  End If
End Sub

Private Sub lblWorkspace_DblClick()
  picWSProperty_DblClick
End Sub

Private Sub mFTP_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long, Percentage As Integer, TransferRate As Double, EstTime As Double)
  RaiseProgress 100, Percentage, ""
End Sub

Private Sub mnuAddFile_Click()
Dim Lcnt As Integer
Dim LFileName As String
Dim lFso As New Scripting.FileSystemObject
Dim lNode As Node
Dim Lselnode As Node
Dim lFile As TextStream
Dim lPath As String
Dim lContent As String
  
  lContent = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" & vbCrLf & _
             "<html>" & vbCrLf & _
             "<head>" & vbCrLf & _
             "<title>Untitled Document</title>" & vbCrLf & vbTab & _
             "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbCrLf & vbTab & _
             "</head>" & vbCrLf & _
             "<body>" & vbCrLf & vbCrLf & _
             "</body>" & vbCrLf & _
             "</html>"
  If mIndex = 1 Then 'Workspace
    If Not tvFiles.SelectedItem Is Nothing Then
      Set Lselnode = tvFiles.SelectedItem
      If Lselnode.Image = "FOLDEROPEN" Or Lselnode.Image = "FOLDERCLOSE" Or Lselnode.Image = "DISK" Or Lselnode.Image = "FLOPPY" Then
        'Get filename
        lPath = IIf(Lselnode.Key <> "F0", Lselnode.Key, Lselnode.Text)
        Do Until Not lFso.FileExists(lPath & "\untitled" & IIf(Lcnt <> 0, Lcnt, "") & ".asp")
        Lcnt = Lcnt + 1
        Loop
        'Write on temp file
        LFileName = "untitled" & IIf(Lcnt <> 0, Lcnt, "") & ".asp"
        Set lFile = lFso.OpenTextFile(lPath & "\" & LFileName, ForWriting, True)
        lFile.Write lContent
        lFile.Close
        'add to treeview
        Set lNode = tvFiles.Nodes.Add(Lselnode.Key, tvwChild, lPath & "\" & LFileName, LFileName, "ASP")
        lNode.Selected = True
        lNode.Tag = LFileName
        lNode.EnsureVisible
        tvFiles.StartLabelEdit
      End If
    End If
  ElseIf mIndex = 3 Then 'Site files
    If Not tvSiteFiles.SelectedItem Is Nothing Then
      Set Lselnode = tvSiteFiles.SelectedItem
      If Lselnode.Image = "FOLDEROPEN" Or Lselnode.Image = "FOLDERCLOSE" Then
        'Get filename
        lPath = IIf(Lselnode.Key <> "F0", Lselnode.Key, Lselnode.Text)
        Do Until Not lFso.FileExists(lPath & "\untitled" & IIf(Lcnt <> 0, Lcnt, "") & ".asp")
        Lcnt = Lcnt + 1
        Loop
        'Write on temp file
        LFileName = "untitled" & IIf(Lcnt <> 0, Lcnt, "") & ".asp"
        Set lFile = lFso.OpenTextFile(lPath & "\" & LFileName, ForWriting, True)
        lFile.Write lContent
        lFile.Close
        'add to treeview
        Set lNode = tvSiteFiles.Nodes.Add(Lselnode.Key, tvwChild, lPath & "\" & LFileName, LFileName, "ASP")
        lNode.Selected = True
        lNode.Tag = LFileName
        lNode.EnsureVisible
        tvSiteFiles.StartLabelEdit
      End If
    End If
  End If
  Set lNode = Nothing
  Set lFso = Nothing
End Sub

Private Sub mnuAddFolder_Click()
Dim Lcnt As Integer
Dim LFoldername As String
Dim lFso As New Scripting.FileSystemObject
Dim lNode As Node
Dim Lselnode As Node
Dim lPath As String
  If mIndex = 3 Then 'Site files
    If Not tvSiteFiles.SelectedItem Is Nothing Then
      Set Lselnode = tvSiteFiles.SelectedItem
      If Lselnode.Image = "FOLDEROPEN" Or Lselnode.Image = "FOLDERCLOSE" Then
        lPath = IIf(Lselnode.Key <> "F0", Lselnode.Key, Lselnode.Text)
        Do Until Not lFso.FolderExists(lPath & "\untitled" & IIf(Lcnt <> 0, Lcnt, ""))
          Lcnt = Lcnt + 1
        Loop
        LFoldername = "untitled" & IIf(Lcnt <> 0, Lcnt, "")
        S101_Make_Dir lPath & "\" & LFoldername
        Set lNode = tvSiteFiles.Nodes.Add(Lselnode.Key, tvwChild, lPath & "\" & LFoldername, LFoldername, "FOLDERCLOSE", "FOLDEROPEN")
        lNode.Selected = True
        lNode.EnsureVisible
        lNode.Tag = LFoldername
        tvSiteFiles.StartLabelEdit
      End If
    End If
  ElseIf mIndex = 1 Then 'Workspace
    If Not tvFiles.SelectedItem Is Nothing Then
      Set Lselnode = tvFiles.SelectedItem
      If Lselnode.Image = "FOLDEROPEN" Or Lselnode.Image = "FOLDERCLOSE" Or Lselnode.Image = "DISK" Then
        lPath = IIf(Lselnode.Key <> "F0", Lselnode.Key, Lselnode.Text)
        If InStr(lPath, "[") > 0 Then lPath = Trim(Split(lPath, "[")(0))
        Do Until Not lFso.FolderExists(lPath & "\untitled" & IIf(Lcnt <> 0, Lcnt, ""))
          Lcnt = Lcnt + 1
        Loop
        LFoldername = "untitled" & IIf(Lcnt <> 0, Lcnt, "")
        S101_Make_Dir lPath & "\" & LFoldername
        Set lNode = tvFiles.Nodes.Add(Lselnode.Key, tvwChild, lPath & "\" & LFoldername, LFoldername, "FOLDERCLOSE", "FOLDEROPEN")
        lNode.Selected = True
        lNode.EnsureVisible
        lNode.Tag = LFoldername
        tvFiles.StartLabelEdit
      End If
    End If
  End If
  Set lNode = Nothing
  Set lFso = Nothing
End Sub

Private Sub mnuClearhistory_Click()
  If MsgBox("Are you sure to clear history?", vbQuestion + vbYesNo, Mtitle) = vbYes Then
    Screen.MousePointer = vbHourglass
    Mhistories.RemoveAll
    Mhistories.Save
    tvHistory.Nodes.Clear
    tvHistory.Nodes.Add(, , "H0", "    (No files)", "EMPTY").ForeColor = RGB(123, 123, 123)
    Screen.MousePointer = vbDefault
  End If
End Sub

Private Sub mnuDelete_Click()
  If mIndex = 1 Then 'WD
    Call tvFiles_KeyDown(vbKeyDelete, 0)
  ElseIf mIndex = 3 Then 'Site
    Call tvSiteFiles_KeyDown(vbKeyDelete, 0)
  End If
End Sub

Private Sub mnuDeleteConnection_Click()
 If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "CONNECTION" Then
      If cboSites.Text = "Define sites..." Then
        Call tvSiteFiles_GotFocus
      End If
      If cboSites.ListIndex >= 0 Then
        If MsgBox("Are you sure to delete this Connection?", vbYesNo + vbQuestion, Mtitle) Then
          DeleteConnectionString
        End If
      End If
    End If
  End If
End Sub

Private Sub mnuDeleteGC_Click()
Dim lCode As String
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateDelete(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      RTB(mEditorIndex).Paste lCode
    End If
  End If
End Sub

Private Sub mnuDeleteGF_Click()
Dim lCode As String
Dim lFileNum As Integer
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateDelete(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      With cdMain
        .FileName = ""
        .CancelError = False
        .Filter = "SQL Files (*.sql)|*.sql|All Files(*.*)|*.*"
        .ShowSave
        If .FileName <> "" Then
          lFileNum = FreeFile
          Open .FileName For Output As #lFileNum
          Print #lFileNum, lCode
          Close #lFileNum
        End If
      End With
    End If
  End If
End Sub

Private Sub mnuEditConnection_Click()
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "CONNECTION" Then
      If cboSites.Text = "Define sites..." Then
        Call tvSiteFiles_GotFocus
      End If
      If cboSites.ListIndex >= 0 Then
        frmConnection.mSitename = cboSites.Text
        frmConnection.txtConnectionString.Text = tvApplication.SelectedItem.Tag
        frmConnection.txtConnectionString.Tag = tvApplication.SelectedItem.Tag
        frmConnection.txtName.Text = tvApplication.SelectedItem.Text
        frmConnection.txtName.Tag = tvApplication.SelectedItem.Text
        frmConnection.mEdit = True
        frmConnection.Show vbModal
      End If
    End If
  End If
End Sub

Private Sub mnuInsertCode_Click()
Dim lCode As String
  If Not tvApplication Is Nothing Then
    Select Case tvApplication.SelectedItem.Image
    Case "CONNECTION"
      lCode = vbTab & "Dim " & tvApplication.SelectedItem.Text & vbCrLf & _
              vbTab & "Set " & tvApplication.SelectedItem.Text & " = Server.CreateObject(" & quote & "ADODB.Connection" & quote & ")" & vbCrLf & _
              vbTab & tvApplication.SelectedItem.Text & ".CursorLocation = 3 'UseClient" & vbCrLf & _
              vbTab & tvApplication.SelectedItem.Text & ".Open " & quote & tvApplication.SelectedItem.Tag & quote
    Case "FIELD"
      If InStr(tvApplication.SelectedItem.Text, "(") > 0 Then
        lCode = Split(tvApplication.SelectedItem.Text, "(")(0)
      Else
        lCode = tvApplication.SelectedItem.Text
      End If
    Case "TABLE"
      lCode = tvApplication.SelectedItem.Text
    End Select
    RTB(mEditorIndex).Paste lCode
  End If
End Sub

Private Sub mnuInsertGC_Click()
Dim lCode As String
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateInsert(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      RTB(mEditorIndex).Paste lCode
    End If
  End If
End Sub

Private Sub mnuInsertGF_Click()
Dim lCode As String
Dim lFileNum As Integer
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateInsert(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      With cdMain
        .FileName = ""
        .CancelError = False
        .Filter = "SQL Files (*.sql)|*.sql|All Files(*.*)|*.*"
        .ShowSave
        If .FileName <> "" Then
          lFileNum = FreeFile
          Open .FileName For Output As #lFileNum
          Print #lFileNum, lCode
          Close #lFileNum
        End If
      End With
    End If
  End If
End Sub

Private Sub mnuInsertGP_Click()
Dim lCode As String
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateInsert(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      Clipboard.SetText lCode
    End If
  End If
End Sub

Private Sub mnuNewConnection_Click()
  If cboSites.Text = "Define sites..." Then
    Call tvSiteFiles_GotFocus
  End If
  If cboSites.ListIndex >= 0 Then
    frmConnection.mSitename = cboSites.Text
    frmConnection.txtConnectionString.Text = ""
    frmConnection.txtName.Text = ""
    frmConnection.Show vbModal
  End If
End Sub

Private Sub mnuNewsite_Click()
  Set mSiteFrm = New frmSiteDetails
  mSiteFrm.LoadDetails
  mSiteFrm.Show vbModal
End Sub

Private Sub mnuOpen_Click()
  If mIndex = 1 Then 'Workspace
    If Not tvFiles.SelectedItem Is Nothing Then
      If tvFiles.SelectedItem.Image = "FOLDEROPEN" Or tvFiles.SelectedItem.Image = "FOLDERCLOSE" Then
        If Not tvFiles.SelectedItem.Expanded Then tvFiles.SelectedItem.Expanded = True 'Not tvFiles.SelectedItem.Expanded
      Else
        tvFiles_DblClick
      End If
    End If
  ElseIf mIndex = 3 Then 'Sitefiles
    If Not tvSiteFiles.SelectedItem Is Nothing Then
      If tvSiteFiles.SelectedItem.Image = "FOLDEROPEN" Or tvSiteFiles.SelectedItem.Image = "FOLDERCLOSE" Then
        If Not tvSiteFiles.SelectedItem.Expanded Then tvSiteFiles.SelectedItem.Expanded = True 'Not tvSiteFiles.SelectedItem.Expanded
      Else
        tvSiteFiles_DblClick
      End If
    End If
  End If
End Sub

Private Sub mnuRemoveHistory_Click()
Dim lParentKey As String
  If Not tvHistory.SelectedItem Is Nothing Then
    If tvHistory.SelectedItem.Key <> "H0" Then
      If tvHistory.SelectedItem.Tag = "1" Then
        If MsgBox("Are you sure to remove the file '" & tvHistory.SelectedItem.Text & "' from history?", vbQuestion + vbYesNo, Mtitle) = vbYes Then
          Screen.MousePointer = vbHourglass
          'Remove and save
          Mhistories.Remove tvHistory.SelectedItem.Key
          Mhistories.Save
          'Udpate the treeview
          lParentKey = tvHistory.SelectedItem.Parent.Key
          tvHistory.Nodes.Remove tvHistory.SelectedItem.Key
          If tvHistory.Nodes(lParentKey).Children = 0 Then
            tvHistory.Nodes.Remove lParentKey
            If tvHistory.Nodes.Count > 0 Then
              tvHistory.Nodes.Clear
              tvHistory.Nodes.Add(, , "H0", "    (No files)", "EMPTY").ForeColor = RGB(123, 123, 123)
            End If
          End If
          Screen.MousePointer = vbDefault
        End If
      Else
        If MsgBox("Are you sure to remove the folder '" & tvHistory.SelectedItem.Text & "' from the history?", vbQuestion + vbYesNo, Mtitle) = vbYes Then
          Screen.MousePointer = vbHourglass
          'Remove and save
          Mhistories.RemoveFor tvHistory.SelectedItem.Key
          Mhistories.Save
          'Update the treeview
          tvHistory.Nodes.Remove tvHistory.SelectedItem.Key
          If tvHistory.Nodes.Count > 0 Then
            tvHistory.Nodes.Clear
            tvHistory.Nodes.Add(, , "H0", "    (No files)", "EMPTY").ForeColor = RGB(123, 123, 123)
          End If
          Screen.MousePointer = vbDefault
        End If
      End If
    End If
  End If
End Sub

Private Sub mnuRemoveSite_Click()
  On Error Resume Next
  If cboSites.Text <> "" And cboSites.Text <> "Define sites..." Then
    If MsgBox("Do you want to delete the site '" & cboSites.Text & "'?", vbQuestion + vbYesNo, Mtitle) = vbYes Then
      Msitedetails.Remove cboSites.Text
      Msitedetails.Save
      cboSites.RemoveItem cboSites.ListIndex
      cboSites.ListIndex = cboSites.ListIndex + 1
    End If
  End If
  
  Err.Clear
End Sub

Private Sub mnuRename_Click()
  If mIndex = 1 Then 'WD
    Call tvFiles_KeyDown(vbKeyF2, 0)
  ElseIf mIndex = 3 Then 'Site
    Call tvSiteFiles_KeyDown(vbKeyF2, 0)
  End If
End Sub

Private Sub mnuRenameSite_Click()
  On Error Resume Next
  If cboSites.Text <> "" And cboSites.Text <> "Define sites..." Then
    Set mSiteFrm = New frmSiteDetails
    mSiteFrm.LoadDetails cboSites.Text
    mSiteFrm.Show vbModal
  End If
  
  Err.Clear
End Sub

Private Sub mnuSelectGC_Click()
Dim lCode As String
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateSelect(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      RTB(mEditorIndex).Paste lCode
    End If
  End If
End Sub

Private Sub mnuSelectGF_Click()
Dim lCode As String
Dim lFileNum As Integer
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateSelect(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      With cdMain
        .FileName = ""
        .CancelError = False
        .Filter = "SQL Files (*.sql)|*.sql|All Files(*.*)|*.*"
        .ShowSave
        If .FileName <> "" Then
          lFileNum = FreeFile
          Open .FileName For Output As #lFileNum
          Print #lFileNum, lCode
          Close #lFileNum
        End If
      End With
    End If
  End If
End Sub

Private Sub mnuSelectGP_Click()
Dim lCode As String
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateSelect(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      Clipboard.SetText lCode
    End If
  End If
End Sub

Private Sub mnuTestConnection_Click()
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "CONNECTION" Then
      TestConnection tvApplication.SelectedItem.Tag
    End If
  End If
End Sub

Private Sub mnuUpdateGC_Click()
Dim lCode As String
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateUpdate(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      RTB(mEditorIndex).Paste lCode
    End If
  End If
End Sub

Private Sub mnuUpdateGF_Click()
Dim lCode As String
Dim lFileNum As Integer
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateUpdate(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      With cdMain
        .FileName = ""
        .CancelError = False
        .Filter = "SQL Files (*.sql)|*.sql|All Files(*.*)|*.*"
        .ShowSave
        If .FileName <> "" Then
          lFileNum = FreeFile
          Open .FileName For Output As #lFileNum
          Print #lFileNum, lCode
          Close #lFileNum
        End If
      End With
    End If
  End If
End Sub

Private Sub mnuUpdateGP_Click()
Dim lCode As String
  If Not tvApplication.SelectedItem Is Nothing Then
    If tvApplication.SelectedItem.Image = "TABLE" Then
      lCode = GenerateUpdate(tvApplication.SelectedItem.Parent.Parent.Text, tvApplication.SelectedItem.Parent.Parent.Tag, tvApplication.SelectedItem.Text)
      Clipboard.SetText lCode
    End If
  End If
End Sub

Private Sub mnuViewRecords_Click()
  tmrOpen.Enabled = True
End Sub

Private Sub mSiteFrm_SiteSaved(ByVal pInfo As String, ByVal pNew As Boolean)
  If pNew Then
    cboSites.Tag = ""
  End If
  LoadSites pInfo
End Sub

Private Sub mStandardMenu_Click(ItemNumber As Long)
Dim lKey As String
Dim lCursor As Boolean
Dim lSite As clsSite
  On Error Resume Next
  lCursor = IIf(RTB(mEditorIndex).SelText = "", True, False)
  lKey = mStandardMenu.ItemKey(ItemNumber)
  Select Case UCase(lKey)
  'FILE MENU
    Case "MNUNEW"
      Set lSite = Msitedetails.Item(cboSites.Text)
      If lSite Is Nothing Then
        LoadDocument
      Else
        LoadDocument , , lSite.LocalPath, , lSite.URL
      End If
    Case "MNUOPEN"
      OpenDocument
    Case "MNUSAVE"
      SaveDocument
    Case "MNUSAVEAS"
      SaveDocument True
    Case "MNUCLOSE"
      CloseDocument
    Case "MNUCLOSEALL"
      CloseAllDocuments
    Case "MNUPRINT"
      RTB(mEditorIndex).PrintText
    Case "MNUEXIT"
      Unload Me
      Exit Sub
  'EDIT MENU
  Case "MNUREDO"
    RTB(mEditorIndex).Redo
  Case "MNUUNDO"
    RTB(mEditorIndex).Undo
  Case "MNUCUT"
    RTB(mEditorIndex).Cut
    ctbHeader.ButtonEnabled("Paste") = True
    mStandardMenu.Enabled(mStandardMenu.IndexForKey("mnuPaste")) = True
  Case "MNUCOPY"
    RTB(mEditorIndex).Copy
    ctbHeader.ButtonEnabled("Paste") = True
    mStandardMenu.Enabled(mStandardMenu.IndexForKey("mnuPaste")) = True
  Case "MNUPASTE"
    RTB(mEditorIndex).Paste
  Case "MNUDELETE"
    RTB(mEditorIndex).Delete
  Case "MNUSELECTALL"
    RTB(mEditorIndex).SelectAll
  Case "MNUFIND"
    SetParent frmFind.hwnd, Me.hwnd
    If RTB(mEditorIndex).SelText <> "" Then
      frmFind.txtFindwhat.Text = RTB(mEditorIndex).SelText
    End If
    frmFind.Show
  Case "MNUGOTOLINE"
    If mWordWrap = False Then frmGotoline.Show vbModal
  'VIEW
  Case "MNUFULLMODEPREVIEW"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    mFullmodePreview = mStandardMenu.Checked(ItemNumber)
    'SetFullmode mFullmodePreview
  Case "MNUWEBPREVIEW"
    PreviewInBrowser
  Case "MNUREFRESHTOOLBOX"
    If mIndex = 1 Then
      lvPaths.ListItems.Clear
      LoadDrives
    ElseIf mIndex = 2 Then
      LoadHistory
    ElseIf mIndex = 3 Then
      lsvSitePath.ListItems.Clear
      LoadSites
    End If
  Case "MNUCODE"
    RTB(mEditorIndex).View 1
  Case "MNUVIEW"
    RTB(mEditorIndex).View 2
  Case "MNUCODEVIEW"
    RTB(mEditorIndex).View 3
  'OPTIONS
  Case "MNUWORDWRAP"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    SetWordwrap Not mStandardMenu.Checked(ItemNumber)
    mWordWrap = mStandardMenu.Checked(ItemNumber)
    stBar.Panels("PRow").Visible = Not mWordWrap
    stBar.Panels("PCol").Visible = Not mWordWrap
  Case "MNUTOOLBOX"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    picTools.Width = IIf(mStandardMenu.Checked(ItemNumber), mToolboxSize, 0)
    picTools.Visible = mStandardMenu.Checked(ItemNumber)
    mToolbox = mStandardMenu.Checked(ItemNumber)
    Form_Resize
  Case "MNULINENUMBER"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    mLineNo = mStandardMenu.Checked(ItemNumber)
    SetLinenumber mLineNo
  Case "MNUOPENDIALOG"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    Mopendialog = mStandardMenu.Checked(ItemNumber)
    frmTemplates.chkOpenDialog.Value = IIf(Mopendialog, 1, 0)
  Case "MNUSYNTAXHIGHLIGHTING"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    mSyntaxHighlighting = mStandardMenu.Checked(ItemNumber)
    SetSyntaxHighlighting mSyntaxHighlighting
  Case "MNUAUTOCOMPLETION"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    mAutoCompletion = mStandardMenu.Checked(ItemNumber)
    SetAutocompletion mAutoCompletion
  Case "MNUINTELISENSE"
    mStandardMenu.Checked(ItemNumber) = Not mStandardMenu.Checked(ItemNumber)
    mIntelisense = mStandardMenu.Checked(ItemNumber)
    SetIntellisense mIntelisense
  'HTML
  Case "MNULINK"
    frmLink.Show vbModal
  Case "MNUEMAILLINK"
    frmEmailLink.Show vbModal
  Case "MNUIMAGE"
    frmImage.Show vbModal
  Case "MNUROLLOVERIMAGE"
    frmRolloverImage.Show vbModal
  Case "MNUBOOKMARK"
    frmBmark.Show vbModal
  Case "MNUTABLES"
    frmTable.Show vbModal
  Case "MNUSTYLESHEET"
    LoadStyleClasses
  Case "MNUCSS"
    frmCSSEditor.Show vbModal
  Case "MNULIST"
    frmList.Show vbModal
  Case "MNUMARQUEE"
    RTB(mEditorIndex).Paste "<marquee>" & RTB(mEditorIndex).SelText & "</marquee>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 10
  Case "MNUSPAN"
    RTB(mEditorIndex).Paste "<span>" & RTB(mEditorIndex).SelText & "</span>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNUDIV"
    RTB(mEditorIndex).Paste "<div>" & RTB(mEditorIndex).SelText & "</div>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 6
  Case "MNUHORIZONTALRULE"
    RTB(mEditorIndex).Paste "<hr>"
  Case "MNUDATE"
    frmDate.Show vbModal
  Case "MNUCLIENT"
    RTB(mEditorIndex).Paste "<script language=javascript>" & vbCrLf & "<!--" & vbCrLf & vbCrLf & "-->" & vbCrLf & "</script>"
    RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 14
  Case "MNUSERVER"
    RTB(mEditorIndex).Paste "<script language=vbscript runat=server>" & vbCrLf & vbCrLf & "</script>"
    RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 10
  Case "MNUFORMVALIDATION"
    frmFormValidation.Show vbModal
  Case "MNUDEFAULTVALUE"
    frmDefaultValue.Show vbModal
  Case "MNUFORM"
    RTB(mEditorIndex).Paste "<form id=form1 name=form1 action="""" method=post></form>"
    RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNUTEXTBOX"
    RTB(mEditorIndex).Paste "<input type=text name=text1 value=""text1"">"
  Case "MNUTEXTAREA"
    RTB(mEditorIndex).Paste "<textarea name=textarea1></textarea>"
  Case "MNUSUBMITBUTTON"
    RTB(mEditorIndex).Paste "<input type=submit name=submitbutton1 value=""submitbutton1"">"
  Case "MNURESETBUTTON"
    RTB(mEditorIndex).Paste "<input type=reset name=resetbutton1 value=""resetbutton1"">"
  Case "MNUHIDDENBOX"
    RTB(mEditorIndex).Paste "<input type=hidden name=hiddenbox1 value=""hiddenbox1"">"
  Case "MNUOPTIONBUTTON"
    RTB(mEditorIndex).Paste "<input type=radio name=radio1 value=""radio1"">"
  Case "MNUPUSHBUTTON"
    RTB(mEditorIndex).Paste "<input type=button name=button1 value=""button1"">"
  Case "MNUCHECKBOX"
    RTB(mEditorIndex).Paste "<input type=checkbox name=checkbox1 value=""checkbox1"" checked>"
  Case "MNULABEL"
    RTB(mEditorIndex).Paste "<label>" & RTB(mEditorIndex).SelText & "</label>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 8
  'WIZARDS
  Case "MNUDSNCONNECTION"
    frmDSN.Show vbModal
  Case "MNUDBCONNECTION"
    frmDB.Show vbModal
  Case "MNUCOOKIE"
    frmCookie.Show vbModal
  'FORMAT
  Case "MNUBOLD"
    RTB(mEditorIndex).Paste "<b>" & RTB(mEditorIndex).SelText & "</b>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 4
  Case "MNUITALIC"
    RTB(mEditorIndex).Paste "<i>" & RTB(mEditorIndex).SelText & "</i>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 4
  Case "MNUUNDERLINE"
    RTB(mEditorIndex).Paste "<u>" & RTB(mEditorIndex).SelText & "</u>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 4
  Case "MNUSUPERSCRIPT"
    RTB(mEditorIndex).Paste "<sup>" & RTB(mEditorIndex).SelText & "</sup>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 6
  Case "MNUSUBSCRIPT"
    RTB(mEditorIndex).Paste "<sub>" & RTB(mEditorIndex).SelText & "</sub>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 6
  Case "MNULEFT"
    RTB(mEditorIndex).Paste "<p align=""left"">" & RTB(mEditorIndex).SelText & "</p>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 4
  Case "MNUCENTER"
    RTB(mEditorIndex).Paste "<p align=""center"">" & RTB(mEditorIndex).SelText & "</p>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 4
  Case "MNURIGHT"
    RTB(mEditorIndex).Paste "<p align=""right"">" & RTB(mEditorIndex).SelText & "</p>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 4
  Case "MNUSTRIKETHROUGH"
    RTB(mEditorIndex).Paste "<s>" & RTB(mEditorIndex).SelText & "</s>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 4
  Case "MNUTELETYPE"
    RTB(mEditorIndex).Paste "<tt>" & RTB(mEditorIndex).SelText & "</tt>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUEMPHASIS"
    RTB(mEditorIndex).Paste "<em>" & RTB(mEditorIndex).SelText & "</em>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUSTROMG"
    RTB(mEditorIndex).Paste "<strong>" & RTB(mEditorIndex).SelText & "</strong>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 9
  Case "MNUCITETION"
    RTB(mEditorIndex).Paste "<cite>" & RTB(mEditorIndex).SelText & "</cite>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNUSTRONG"
    RTB(mEditorIndex).Paste "<strong>" & RTB(mEditorIndex).SelText & "</strong>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNUDEFINITION"
    RTB(mEditorIndex).Paste "<dfn>" & RTB(mEditorIndex).SelText & "</dfn>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 6
  Case "MNUDELETED"
    RTB(mEditorIndex).Paste "<del>" & RTB(mEditorIndex).SelText & "</del>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 6
  Case "MNUINSERTED"
    RTB(mEditorIndex).Paste "<ins>" & RTB(mEditorIndex).SelText & "</ins>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 6
  Case "MNUCODE"
    RTB(mEditorIndex).Paste "<code>" & RTB(mEditorIndex).SelText & "</code>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNUVARIABLE"
    RTB(mEditorIndex).Paste "<var>" & RTB(mEditorIndex).SelText & "</var>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 6
  Case "MNUSAMPLE"
    RTB(mEditorIndex).Paste "<samp>" & RTB(mEditorIndex).SelText & "</samp>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNUKEYBOARD"
    RTB(mEditorIndex).Paste "<kbd>" & RTB(mEditorIndex).SelText & "</kbd>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 6
  Case "MNUBACKGROUNDCOLOR"
    'mnuBackcolor_CLICK
  Case "MNU1"
    RTB(mEditorIndex).Paste "<font size=""1"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU2"
    RTB(mEditorIndex).Paste "<font size=""2"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU3"
    RTB(mEditorIndex).Paste "<font size=""3"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU4"
    RTB(mEditorIndex).Paste "<font size=""4"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU5"
    RTB(mEditorIndex).Paste "<font size=""5"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU6"
    RTB(mEditorIndex).Paste "<font size=""6"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU7"
    RTB(mEditorIndex).Paste "<font size=""7"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU+1"
    RTB(mEditorIndex).Paste "<font size=""+1"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU+2"
    RTB(mEditorIndex).Paste "<font size=""+2"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU+3"
    RTB(mEditorIndex).Paste "<font size=""+3"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU+4"
    RTB(mEditorIndex).Paste "<font size=""+4"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU-1"
    RTB(mEditorIndex).Paste "<font size=""-1"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU-2"
    RTB(mEditorIndex).Paste "<font size=""-2"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNU-3"
    RTB(mEditorIndex).Paste "<font size=""-3"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNUDEFAULTFONT"
    RTB(mEditorIndex).Paste "<font face=""0Font"">" & RTB(mEditorIndex).SelText & "</font>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
  Case "MNUOTHERFONTS"
    frmFonts.Show vbModal
  Case "MNUUNORDEREDLIST"
    RTB(mEditorIndex).Paste "<ul>" & RTB(mEditorIndex).SelText & "</ul>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUORDEREDLIST"
    RTB(mEditorIndex).Paste "<ol>" & RTB(mEditorIndex).SelText & "</ol>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUDEFINITIONLIST"
    RTB(mEditorIndex).Paste "<dl>" & RTB(mEditorIndex).SelText & "</dl>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUPARAGRAPH"
    RTB(mEditorIndex).Paste "<p>" & RTB(mEditorIndex).SelText & "</p>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 4
  Case "MNUHEADING1"
    RTB(mEditorIndex).Paste "<h1>" & RTB(mEditorIndex).SelText & "</h1>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUHEADING2"
    RTB(mEditorIndex).Paste "<h2>" & RTB(mEditorIndex).SelText & "</h2>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUHEADING3"
    RTB(mEditorIndex).Paste "<h3>" & RTB(mEditorIndex).SelText & "</h3>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUHEADING4"
    RTB(mEditorIndex).Paste "<h4>" & RTB(mEditorIndex).SelText & "</h4>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUHEADING5"
    RTB(mEditorIndex).Paste "<h5>" & RTB(mEditorIndex).SelText & "</h5>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUHEADING6"
    RTB(mEditorIndex).Paste "<h6>" & RTB(mEditorIndex).SelText & "</h6>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 5
  Case "MNUPRETEXT"
    RTB(mEditorIndex).Paste "<pre>" & RTB(mEditorIndex).SelText & "</pre>"
    If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 6
  'Special characters
  Case "MNULINEBREAK"
    RTB(mEditorIndex).Paste "<br>"
  Case "MNUNBSP"
    RTB(mEditorIndex).Paste "&nbsp;"
  Case "MNUCOPYRIGHT"
    RTB(mEditorIndex).Paste "&copy;"
  Case "MNUREGISTERED"
    RTB(mEditorIndex).Paste "&reg;"
  Case "MNUTRADEMARK"
    RTB(mEditorIndex).Paste "&#8482;"
  Case "MNUPOUND"
    RTB(mEditorIndex).Paste "&pound;"
  Case "MNUYEN"
    RTB(mEditorIndex).Paste "&yen;"
  Case "MNUEURO"
    RTB(mEditorIndex).Paste "&#8364;"
  Case "MNULEFTQUOTE"
    RTB(mEditorIndex).Paste "&#8220;"
  Case "MNURIGHTQUOTE"
    RTB(mEditorIndex).Paste "&#8221;"
  Case "MNUEMDASH"
    RTB(mEditorIndex).Paste "&#8212;"
  Case "MNUOTHERSSC"
    frmSpecialCharacters.Show vbModal
  Case "MNUMETA"
    frmMeta.Show vbModal
  Case "MNUKEYWORDS"
    frmKeyword.mType = 1
    frmKeyword.Show vbModal
  Case "MNUDESCRIPTION"
    frmKeyword.mType = 2
    frmKeyword.Show vbModal
  Case "MNUREFRESH"
    frmRefresh.Show vbModal
  Case "MNUBASE"
    frmBase.Show vbModal
  Case "MNUMETALINK"
    frmMetaLink.Show vbModal
  Case "MNUINDENT"
    RTB(mEditorIndex).Indent
  Case "MNUOUTDENT"
    RTB(mEditorIndex).Outdent
  'Tools
  Case "MNUSITESHOW"
    picSProperty_Click
  Case "MNUNEWSITE"
    picSProperty_Click
    frmSites.Show vbModal
  Case "MNUEDITSITE"
    picSProperty_Click
    frmSites.Show vbModal
  Case "MNUHISTORY"
    picHProperty_Click
  Case "MNUCLEARHISTORY"
    picHProperty_Click
    ClearHistory
  Case "MNUDATAEXPLORER"
    picAProperty_Click
  Case "MNUNEWCONNECTION"
    picAProperty_Click
    mnuNewConnection_Click
  'Help
  Case "MNUREGISTRATION"
    frmRegistration.Show vbModal
  Case "MNUABOUTUS"
    frmSplash.Show vbModal
  'Files list
  Case Else
    If LCase(Left(lKey, 4)) = "file" Then
      UncheckAllMenu
      mStandardMenu.Checked(mStandardMenu.IndexForKey(lKey)) = True
      tabMain.Tabs.Item(Mid(lKey, 5)).Selected = True
    ElseIf LCase(Left(lKey, 7)) = "classes" Then
      If Len(RTB(mEditorIndex).SelText) > 0 Then
        RTB(mEditorIndex).Paste """" & mStandardMenu.Caption(ItemNumber) & """"
      Else
        RTB(mEditorIndex).Paste "class=""" & mStandardMenu.Caption(ItemNumber) & """"
      End If
    ElseIf LCase(Left(lKey, 4)) = "font" Then
      RTB(mEditorIndex).Paste "<font face=""" & mStandardMenu.Caption(ItemNumber) & """>" & RTB(mEditorIndex).SelText & "</font>"
      If lCursor = True Then RTB(mEditorIndex).SelStart = RTB(mEditorIndex).SelStart - 7
    End If
  End Select
  If RTB(mEditorIndex).SelLength = 0 Then RTB(mEditorIndex).SelColor = vbBlack
  
  Err.Clear
End Sub

Private Sub picAProperty_Click()
  If mIndex <> 4 Then
    mIndex = 4
    CollapseTool mIndex
  End If
End Sub

Private Sub picAProperty_DblClick()
  If mIndex = 4 Then
    mIndex = 3
    CollapseTool mIndex
  Else
    picAProperty_Click
  End If
End Sub

Private Sub picHProperty_Click()
  If mIndex <> 2 Then
    mIndex = 2
    CollapseTool mIndex
  End If
End Sub

Private Sub picHProperty_DblClick()
  If mIndex = 2 Then
    mIndex = 3
    CollapseTool mIndex
  Else
    picHProperty_Click
  End If
End Sub

Private Sub picHProperty_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If spBorderH.Visible = False Then
    picTools_MouseMove Button, Shift, x, y
    spBorderH.Visible = True
  End If
End Sub

Private Sub picSProperty_Click()
  If mIndex <> 3 Then
    mIndex = 3
    CollapseTool mIndex
  End If
End Sub

Private Sub picSProperty_DblClick()
  If mIndex = 3 Then
    mIndex = 4
    CollapseTool mIndex
  Else
    picSProperty_Click
  End If
End Sub

Private Sub picSProperty_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If spBorderS.Visible = False Then
    picTools_MouseMove Button, Shift, x, y
    spBorderS.Visible = True
  End If
End Sub

Private Sub picToolbox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'  picDrag.Visible = True
'  FormDrag picDrag
'  picDrag.Visible = False
End Sub

Private Sub picTools_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If spBorderWS.Visible Then spBorderWS.Visible = False
  If spBorderH.Visible Then spBorderH.Visible = False
  If spBorderS.Visible Then spBorderS.Visible = False
End Sub

Private Sub picWSProperty_Click()
  If mIndex <> 1 Then
    mIndex = 1
    CollapseTool mIndex
  End If
End Sub

Private Sub picWSProperty_DblClick()
  If mIndex = 1 Then
    mIndex = 2
    CollapseTool mIndex
  Else
    picWSProperty_Click
  End If
End Sub

Private Sub picWSProperty_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If spBorderWS.Visible = False Then
    picTools_MouseMove Button, Shift, x, y
    spBorderWS.Visible = True
  End If
End Sub

Private Sub RTB_BrowserStatusChanged(Index As Integer, ByVal pText As String)
  stBar.Panels("P1").Text = pText
End Sub

Private Sub RTB_Changed(Index As Integer, ByVal pChange As Boolean)
  If pChange Then
    If Right(tabMain.SelectedTab.Caption, 1) <> "*" Then tabMain.SelectedTab.Caption = tabMain.SelectedTab.Caption & "*"
  Else
    If Right(tabMain.SelectedTab.Caption, 1) = "*" Then tabMain.SelectedTab.Caption = Left(tabMain.SelectedTab.Caption, Len(tabMain.SelectedTab.Caption) - 1)
  End If
End Sub

Private Sub RTB_DocumentOpened(Index As Integer, ByVal pNew As Boolean)
  If mEnable = False Then
    EnableMenuMM True
  End If
End Sub

Private Sub RTB_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
  If mDrag Then
    If mSource <> "" Then
      If mIndex = 1 Then
        tvFiles.Nodes(mSource).Selected = True
        tvFiles_DblClick
      ElseIf mIndex = 2 Then
        tvHistory.Nodes(mSource).Selected = True
        tvHistory_DblClick
      ElseIf mIndex = 3 Then
        tvSiteFiles.Nodes(mSource).Selected = True
        tvSiteFiles_DblClick
      End If
    End If
  End If
End Sub

Private Sub RTB_FileSaved(Index As Integer, ByVal pFilename As String, ByVal pNew As Boolean)
  If RTB(Index).IsRemote Then 'if remote path, before preview upload the file
    If pNew = False Then
      UploadFile RTB(Index).FileName, RTB(Index).VirtualPath
    End If
  Else
    If pNew Then
      lsvSitePath.ListItems.Clear
      LoadSites
    End If
  End If
End Sub

Private Sub RTB_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
  If ctbHeader.ButtonEnabled("Undo") = False Then
    mStandardMenu.Enabled(mStandardMenu.IndexForKey("mnuUndo")) = True
    ctbHeader.ButtonEnabled("Undo") = True
  End If
End Sub

Private Sub RTB_LockControl(Index As Integer)
  LockWindowUpdate stBar.hwnd
End Sub

Private Sub RTB_ModeChanged(Index As Integer, ByVal Mode As Integer)
  If Mode = 2 Then 'View to fullmode
    If mFullmodePreview Then
      picTools.Width = 0
      Form_Resize
      RTB(Index).Fullmode = True
    End If
  Else
    If mToolbox = True Then
      picTools.Width = mToolboxSize
      Form_Resize
    End If
    RTB(Index).Fullmode = False
  End If
End Sub

Private Sub RTB_Position(Index As Integer, ByVal x As Long, ByVal y As Long)
  stBar.Panels("PRow").Text = "Ln " & y
  stBar.Panels("PCol").Text = "Col " & x
End Sub

Private Sub RTB_ProgressStatus(Index As Integer, ByVal pValue As Integer, ByVal pMax As Integer, ByVal pStatus As String)
  RaiseProgress pMax, pValue, pStatus, True
End Sub

Private Sub RTB_ReleaseControl(Index As Integer)
  LockWindowUpdate 0&
End Sub

Private Sub RTB_StateChanged(Index As Integer)
  UpdateEditMenu
End Sub

Private Sub tabMain_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
  On Error Resume Next
  mEditorIndex = val(theTab.Tag)
  UpdateEditMenu
  RTB(mEditorIndex).Visible = True
  RTB_ModeChanged mEditorIndex, RTB(mEditorIndex).Mode
  RTB(mEditorIndex).SetFocus
  
  Err.Clear
End Sub

Private Sub tabMain_TabClose(theTab As vbalDTab6.cTab, bCancel As Boolean)
Dim lResult As VbMsgBoxResult
  On Error Resume Next
  If RTB(val(theTab.Tag)).Changed Then
    lResult = MsgBox(IIf(RTB(val(theTab.Tag)).FileName = "", RTB(val(theTab.Tag)).Key, RTB(val(theTab.Tag)).FileName) & vbCrLf & vbCrLf & "The above document has been changed. Do you want to save changes?", vbQuestion + vbYesNoCancel, Mtitle)
    If lResult = vbYes Then
      If RTB(val(theTab.Tag)).FileName = "" Then
        RTB(val(theTab.Tag)).SaveAsFile
      Else
        RTB(val(theTab.Tag)).SaveFile RTB(val(theTab.Tag)).FileName
      End If
    ElseIf lResult = vbCancel Then
      bCancel = True
      Exit Sub
    End If
  End If
  RTB(val(theTab.Tag)).Changed = False
  If RTB(val(theTab.Tag)).IsRemote Then 'If remote then delete the temp page
    S105_Delete RTB(val(theTab.Tag)).FileName
  Else 'Else save to history
    If RTB(val(theTab.Tag)).IsHistory = False Then SaveHistory RTB(val(theTab.Tag)).FileName
  End If
  Unload RTB(val(theTab.Tag))
  tmr.Enabled = True
  If tabMain.Tabs.Count = 1 Then
    tabMain.Visible = False
    mEditorIndex = 0
  End If
  mStandardMenu.RemoveItem "file" & theTab.Key
  If tabMain.Visible = False Then
    mStandardMenu.AddItem "  No Documents", , , mStandardMenu.IndexForKey("mnuWindows"), , , False, "mnuNo"
    EnableMenuMM False
  End If
  
  Err.Clear
End Sub

Private Sub tmr_Timer()
  If Not tabMain.SelectedTab Is Nothing Then
    tabMain_TabClick tabMain.SelectedTab, vbLeftButton, vbAltMask, 0, 0
  End If
  tmr.Enabled = False
End Sub

Private Sub tmrFocus_Timer()
  On Error Resume Next
  tmrFocus.Enabled = False
  RTB(mEditorIndex).SetFocus
End Sub

Private Sub tmrOpen_Timer()
Dim lNode As Object
  If Not tvApplication.SelectedItem Is Nothing Then
    Set lNode = tvApplication.SelectedItem
    frmRecordset.MConnectionName = lNode.Parent.Parent.Text
    frmRecordset.mConnectionString = lNode.Parent.Parent.Tag
    frmRecordset.MTable = lNode.Text
    frmRecordset.Show vbModal
  End If
  tmrOpen.Enabled = False
End Sub

Private Sub tmrResize_Timer()
  tmrResize.Enabled = False
  prgbarMain.Left = stBar.Panels("P1").Width + Screen.TwipsPerPixelX * 5
  prgbarMain.Top = stBar.Top + Screen.TwipsPerPixelY * 4 'picEditor.Top + picEditor.Height + Screen.TwipsPerPixelY * 10
  prgbarMain.Width = stBar.Panels("P2").Width - Screen.TwipsPerPixelX * 4
End Sub

Private Sub tvApplication_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lNode As Object
  If Button = vbRightButton Then
    EnableMenuDE
    PopupMenu mnuApplication, , picTools.Left + picApplication.Left + tvApplication.Left + x, picTools.Top + picApplication.Top + tvApplication.Top + y
  End If
End Sub

Private Sub tvFiles_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim li As Long
Dim lj As Long
Dim lStart As Long
  Screen.MousePointer = vbHourglass
  If NewString <> tvFiles.Tag Then
    If S104_Rename(tvFiles.SelectedItem.Key, NewString, IIf(InStr(tvFiles.SelectedItem.Image, "FOLDER") > 0, True, False)) = True Then
      For li = 1 To tvFiles.SelectedItem.Children
        lStart = tvFiles.SelectedItem.Index + 1
        For lj = lStart To tvFiles.Nodes.Count
          If tvFiles.Nodes(lj).Parent.Key = tvFiles.SelectedItem.Key Then
            tvFiles.Nodes.Remove lj
            Exit For
          End If
        Next
      Next
      tvFiles.SelectedItem.Key = Mid(tvFiles.SelectedItem.Key, 1, InStrRev(tvFiles.SelectedItem.Key, "\")) & NewString 'Ucase Changed
      If PathExists(tvFiles.SelectedItem.Key) Then lvPaths.ListItems.Remove tvFiles.SelectedItem.Key
      LoadFiles tvFiles, tvFiles.SelectedItem.Key, tvFiles.SelectedItem.Key
    Else
      MsgBox "Unable to rename the item!", vbInformation + vbOKOnly, Mtitle
      Cancel = 1
    End If
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub tvFiles_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Image = "FOLDEROPEN" Then Node.Image = "FOLDERCLOSE"
End Sub

Private Sub tvFiles_DragDrop(Source As Control, x As Single, y As Single)
Dim lDest As String
Dim lNode As Object
Dim lResult As Boolean
  On Error Resume Next
  If mDrag Then
    mDrag = False
    Set tvFiles.DropHighlight = tvFiles.HitTest(x, y)
    mSource = Replace(mSource, "\\", "\")
    If Left(mSource, 1) = "\" Then mSource = "\" & mSource 'Networkpath
    If mSource <> "" Then
      If Not tvFiles.DropHighlight Is Nothing Then
        lDest = tvFiles.DropHighlight.Key
        If lDest = "D1" Then lDest = GetFolder(ftMYDOCUMENTS, Me.hwnd) 'For my documents
        If mSource = lDest Then Exit Sub
        If (tvFiles.DropHighlight.Image = "DISK" Or tvFiles.DropHighlight.Image = "MYD" Or tvFiles.DropHighlight.Image = "FLOPPY" Or tvFiles.DropHighlight.Image = "FOLDEROPEN" Or tvFiles.DropHighlight.Image = "FOLDERCLOSE" Or mSourceType = "FOLDEROPEN") And (mSourceType = "FOLDERCLOSE" Or mSourceType = "FOLDEROPEN") Then
          'For folders
          lsvSitePath.ListItems.Remove mSource 'Ucase Changed
          If mCopy = False Then
            lResult = S109_Copy_Folder(mSource, lDest & Mid(mSource, InStrRev(mSource, "\")), True)
            If lResult Then tvFiles.Nodes.Remove mSource 'Ucase Changed
          Else
            lResult = S109_Copy_Folder(mSource, lDest & Mid(mSource, InStrRev(mSource, "\")))
          End If
          If lResult Then
            Set lNode = tvFiles.Nodes.Add(IIf(lDest = GetFolder(ftMYDOCUMENTS, Me.hwnd), "D1", lDest), tvwChild, lDest & Mid(mSource, InStrRev(mSource, "\")), Mid(mSource, InStrRev(mSource, "\") + 1), "FOLDERCLOSE") 'Ucase Changed
            lNode.Tag = lDest & Mid(mSource, InStrRev(mSource, "\"))
            lNode.Selected = True
            lNode.Expanded = True
            tvFiles_Expand lNode
          End If
        ElseIf (tvFiles.DropHighlight.Image = "DISK" Or tvFiles.DropHighlight.Image = "MYD" Or tvFiles.DropHighlight.Image = "FLOPPY" Or tvFiles.DropHighlight.Image = "FOLDEROPEN" Or tvFiles.DropHighlight.Image = "FOLDERCLOSE" Or mSourceType = "FOLDEROPEN") And (mSourceType <> "FOLDERCLOSE" Or mSourceType <> "FOLDEROPEN") Then
          'For Files
          If mCopy = False Then
            lResult = S108_Copy_File(mSource, lDest & Mid(mSource, InStrRev(mSource, "\")), True)
            If lResult Then tvFiles.Nodes.Remove mSource 'Ucase Changed
          Else
            lResult = S108_Copy_File(mSource, lDest & Mid(mSource, InStrRev(mSource, "\")))
          End If
          If lResult Then
            tvFiles.Nodes.Add(IIf(lDest = GetFolder(ftMYDOCUMENTS, Me.hwnd), "D1", lDest), tvwChild, lDest & Mid(mSource, InStrRev(mSource, "\")), Mid(mSource, InStrRev(mSource, "\") + 1), GetFileImg(mSource)).Tag = lDest & Mid(mSource, InStrRev(mSource, "\")) 'Ucase Changed
          End If
        End If
      End If
    End If
    tvFiles.Drag vbEndDrag
  End If
  
  Err.Clear
End Sub

Private Sub tvFiles_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  Set tvFiles.DropHighlight = tvFiles.HitTest(x, y)
End Sub

Private Sub tvFiles_Expand(ByVal Node As MSComctlLib.Node)
  On Error Resume Next
  If Node.Parent.Key = "D3" Then
    LoadNetworkfiles Node
    Node.EnsureVisible
    Exit Sub
  End If
  If Node.Image <> "ASP" And Node.Image <> "HTML" And Node.Image <> "HTM" Or Node.Image <> "JS" Or Node.Image <> "DEFAULT" Then
    If Node.Image = "FOLDERCLOSE" And Node.Children > 0 Then Node.Image = "FOLDEROPEN"
    LoadFiles tvFiles, Node.Key, Node.Key
    'Node.Selected = True
    If Node.Image = "DISK" Or Node.Image = "CDROM" Or Node.Image = "FLOPPY" Then tvFiles.Nodes.Remove Node.Key & " "
  End If
  
  Err.Clear
End Sub

Private Sub tvFiles_DblClick()
Dim lNode As Node
Dim lExt As String
  If Not tvFiles.SelectedItem Is Nothing Then
    Set lNode = tvFiles.SelectedItem
    If S102_File_Exists(lNode.Key) = True Then
      lExt = UCase(Mid(lNode.Key, InStrRev(lNode.Key, ".") + 1))
      If lExt = "ASP" Or lExt = "HTM" Or lExt = "HTML" Or lExt = "JS" Or lExt = "TXT" Or lExt = "XML" Or lExt = "INI" Or lExt = "CSS" Then
        LoadDocument lNode.Key, lNode.Text
      Else
        Call ShellExecute(Me.hwnd, "open", lNode.Text, vbNullString, Mid(lNode.Key, 1, InStrRev(lNode.Key, "\") - 1), 1&)
      End If
    End If
  End If
End Sub

Private Sub tvFiles_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lNode As Node
  If KeyCode = vbKeyF2 Then
    Set lNode = tvFiles.SelectedItem
    If Not lNode Is Nothing Then
      If InStr("DES MYD MYCOMPUTER MYN FLOPPY DISK CDROM", lNode.Image) = 0 Then
        tvFiles.Tag = lNode.Text
        tvFiles.StartLabelEdit
      End If
    End If
    KeyCode = 0
  ElseIf KeyCode = vbKeyDelete Then
    Set lNode = tvFiles.SelectedItem
    If Not lNode Is Nothing Then
      If InStr("DES MYD MYCOMPUTER MYN FLOPPY DISK CDROM", lNode.Image) = 0 Then
        If MsgBox("Are you sure to delete '" & lNode.Text & "'?", vbQuestion + vbYesNo + vbDefaultButton2, Mtitle) = vbYes Then
          If S105_Delete(lNode.Key, IIf(lNode.Image = "FOLDERCLOSE", True, False)) Then
            tvFiles.Nodes.Remove lNode.Key
          Else
            MsgBox "File cannot be deteted. Access Denied.", vbInformation, Mtitle
          End If
        End If
      End If
    End If
    KeyCode = 0
  End If
End Sub

Private Sub tvFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not tvFiles.HitTest(x, y) Is Nothing Then
    tvFiles.HitTest(x, y).Selected = True
    If tvFiles.HitTest(x, y).Key = "F1" Then
      mSource = ""
      mSourceType = ""
    Else
      mSource = tvFiles.HitTest(x, y).Tag
      mSourceType = tvFiles.HitTest(x, y).Image
    End If
    If Shift = 2 Then
      mCopy = True
    Else
      mCopy = False
    End If
  End If
End Sub

Private Sub tvFiles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    mDrag = True
    tvFiles.Drag vbBeginDrag
    If mCopy Then
      Set tvFiles.DragIcon = imlFiles.ListImages("DRAGCOPY").Picture
    Else
      Set tvFiles.DragIcon = imlFiles.ListImages("DRAG").Picture
    End If
  End If
End Sub

Private Sub tvFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  mDrag = False
  mSource = ""
  mSourceType = ""
  mCopy = False
  If Button = vbRightButton Then
    If Not tvFiles.SelectedItem Is Nothing Then
      EnableMenuWS tvFiles.SelectedItem.Image
    Else
      EnableMenuWS ""
    End If
    PopupMenu mnuFile, , picTools.Left + picWorkSpace.Left + tvFiles.Left + x, picTools.Top + picWorkSpace.Top + tvFiles.Top + y
  End If
End Sub

Private Sub tvHistory_DblClick()
Dim lNode As Node
Dim lFile As String
Dim lExt As String
  If Not tvHistory.SelectedItem Is Nothing Then
    Set lNode = tvHistory.SelectedItem
    If lNode.Image <> "EMPTY" And lNode.Image <> "HISTORY" And lNode.Image <> "HTODAY" And lNode.Image <> "HPAST" Then
      lFile = Split(lNode.Key, "^")(1)
      If S102_File_Exists(lFile) = True Then
        lExt = UCase(Mid(lFile, InStrRev(lFile, ".") + 1))
        If lExt = "ASP" Or lExt = "HTM" Or lExt = "HTML" Or lExt = "JS" Or lExt = "TXT" Or lExt = "XML" Or lExt = "INI" Or lExt = "CSS" Then
          LoadDocument lFile, lNode.Text, , True
        End If
      Else
        lNode.ForeColor = vbRed
      End If
    End If
  End If
End Sub

Private Sub tvHistory_DragDrop(Source As Control, x As Single, y As Single)
  If mDrag Then
    mDrag = False
    tvHistory.Drag vbEndDrag
  End If
End Sub

Private Sub tvHistory_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    If Shift = 1 Then 'Clear the history
      mnuClearhistory_Click
    Else 'Remove the file/folder from history
      mnuRemoveHistory_Click
    End If
  End If
End Sub

Private Sub tvHistory_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not tvHistory.HitTest(x, y) Is Nothing Then
    tvHistory.HitTest(x, y).Selected = True
  End If
End Sub

Private Sub tvHistory_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    mDrag = True
    Set tvHistory.DragIcon = imlFiles.ListImages("DRAG").Picture
    tvHistory.Drag vbBeginDrag
  End If
End Sub

Private Sub tvHistory_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  mDrag = False
  If Button = vbRightButton Then
    mnuRemoveHistory.Enabled = False
    If Not tvHistory.SelectedItem Is Nothing Then
      If tvHistory.SelectedItem.Key <> "H0" Then mnuRemoveHistory.Enabled = True
    End If
    PopupMenu mnuHistory, , picTools.Left + picHistory.Left + tvHistory.Left + x, picTools.Top + picHistory.Top + tvHistory.Top + y
  End If
End Sub

Private Sub tvSiteFiles_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim li As Long
Dim lj As Long
Dim lStart As Long
Dim lExt As String
Dim lFile As String
Dim lNode As Object
  Screen.MousePointer = vbHourglass
  If NewString <> tvSiteFiles.Tag Then
    If S104_Rename(tvSiteFiles.SelectedItem.Key, NewString, IIf(InStr(tvSiteFiles.SelectedItem.Image, "FOLDER") > 0, True, False)) = True Then
      'Remove the children
      For li = 1 To tvSiteFiles.SelectedItem.Children
        lStart = tvSiteFiles.SelectedItem.Index + 1
        For lj = lStart To tvSiteFiles.Nodes.Count
          If tvSiteFiles.Nodes(lj).Parent.Key = tvSiteFiles.SelectedItem.Key Then
            tvSiteFiles.Nodes.Remove lj
            Exit For
          End If
        Next
      Next
      lFile = Mid(tvSiteFiles.SelectedItem.Key, 1, InStrRev(tvSiteFiles.SelectedItem.Key, "\")) & NewString 'Ucase Changed
      'Change the key
      tvSiteFiles.SelectedItem.Key = lFile
      If PathExists(tvSiteFiles.SelectedItem.Key, True) Then lsvSitePath.ListItems.Remove tvSiteFiles.SelectedItem.Key
      LoadFiles tvSiteFiles, tvSiteFiles.SelectedItem.Key, tvSiteFiles.SelectedItem.Key, True
    Else
      MsgBox "Unable to rename the item!", vbInformation + vbOKOnly, Mtitle
      Cancel = 1
    End If
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub tvSiteFiles_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Image = "FOLDEROPEN" Then
    Node.Image = "FOLDERCLOSE"
  ElseIf Node.Image = "FOLDEROPENONLINE" Then
    Node.Image = "FOLDERCLOSEONLINE"
  End If
End Sub

Private Sub tvSiteFiles_DblClick()
Dim lNode As Node
Dim lExt As String
Dim lFile As String
Dim lSite As clsSite
  If Not tvSiteFiles.SelectedItem Is Nothing Then
    Set lNode = tvSiteFiles.SelectedItem
    Set lSite = Msitedetails.Item(cboSites.Text)
    If Not lSite Is Nothing Then
      If S102_File_Exists(lNode.Key) = True Then
        'Open local site files
        lExt = UCase(Mid(lNode.Key, InStrRev(lNode.Key, ".") + 1))
        If lExt = "ASP" Or lExt = "HTM" Or lExt = "HTML" Or lExt = "JS" Or lExt = "TXT" Or lExt = "XML" Or lExt = "INI" Or lExt = "CSS" Then
          LoadDocument lNode.Key, lNode.Text, lSite.LocalPath, , lSite.URL
        Else
          Call ShellExecute(Me.hwnd, "open", lNode.Text, vbNullString, Mid(lNode.Key, 1, InStrRev(lNode.Key, "\") - 1), 1&)
        End If
      ElseIf tvSiteFiles.Nodes("F0").Tag = "R" Then
        'Open remote site files
        lFile = DownloadFile(lNode.Key, App.Path & "\Temp")
        If lFile <> "" Then
          lExt = UCase(Mid(lNode.Key, InStrRev(lNode.Key, ".") + 1))
          If lExt = "ASP" Or lExt = "HTM" Or lExt = "HTML" Or lExt = "JS" Or lExt = "TXT" Or lExt = "XML" Or lExt = "INI" Or lExt = "CSS" Then
            LoadDocument lFile, lNode.Text, App.Path & "\Temp", , lSite.URL, True
          Else
            Call ShellExecute(Me.hwnd, "open", lNode.Text, vbNullString, Mid(lFile, 1, InStrRev(lFile, "\")), 1&)
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub tvSiteFiles_DragDrop(Source As Control, x As Single, y As Single)
Dim lDest As String
Dim lNode As Object
Dim lResult As Boolean
  On Error Resume Next
  If mDrag And tvSiteFiles.Nodes("F0").Tag <> "R" Then 'not for remote site
    mDrag = False
    Set tvSiteFiles.DropHighlight = tvSiteFiles.HitTest(x, y)
    If mSource <> "" Then
      If Not tvSiteFiles.DropHighlight Is Nothing Then
        lDest = tvSiteFiles.DropHighlight.Key
        If lDest = "F0" Then lDest = tvSiteFiles.DropHighlight.Text 'For root folder
        If mSource = lDest Then Exit Sub
        If (tvSiteFiles.DropHighlight.Image = "FOLDEROPEN" Or tvSiteFiles.DropHighlight.Image = "FOLDERCLOSE" Or mSourceType = "FOLDEROPEN") And (mSourceType = "FOLDERCLOSE" Or mSourceType = "FOLDEROPEN") Then
          'For folders
          lsvSitePath.ListItems.Remove mSource 'Ucase Changed
          If mCopy = False Then
            lResult = S109_Copy_Folder(mSource, lDest & Mid(mSource, InStrRev(mSource, "\")), True)
            If lResult Then tvSiteFiles.Nodes.Remove mSource 'Ucase Changed
          Else
            lResult = S109_Copy_Folder(mSource, lDest & Mid(mSource, InStrRev(mSource, "\")))
          End If
          If lResult Then
            Set lNode = tvSiteFiles.Nodes.Add(IIf(lDest = tvSiteFiles.DropHighlight.Text, "F0", lDest), tvwChild, lDest & Mid(mSource, InStrRev(mSource, "\")), Mid(mSource, InStrRev(mSource, "\") + 1), "FOLDERCLOSE") 'Ucase Changed
            lNode.Tag = lDest & Mid(mSource, InStrRev(mSource, "\"))
            lNode.Selected = True
            lNode.Expanded = True
            tvSiteFiles_Expand lNode
          End If
        ElseIf (tvSiteFiles.DropHighlight.Image = "FOLDEROPEN" Or tvSiteFiles.DropHighlight.Image = "FOLDERCLOSE" Or mSourceType = "FOLDEROPEN") And (mSourceType <> "FOLDERCLOSE" Or mSourceType <> "FOLDEROPEN") Then
          'For Files
          If mCopy = False Then
            lResult = S108_Copy_File(mSource, lDest & Mid(mSource, InStrRev(mSource, "\")), True)
            If lResult Then tvSiteFiles.Nodes.Remove mSource 'Ucase Changed
          Else
            lResult = S108_Copy_File(mSource, lDest & Mid(mSource, InStrRev(mSource, "\")))
          End If
          If lResult Then
            tvSiteFiles.Nodes.Add(IIf(lDest = tvSiteFiles.DropHighlight.Text, "F0", lDest), tvwChild, lDest & Mid(mSource, InStrRev(mSource, "\")), Mid(mSource, InStrRev(mSource, "\") + 1), GetFileImg(mSource)).Tag = lDest & Mid(mSource, InStrRev(mSource, "\")) 'Ucase Changed
          End If
        End If
      End If
    End If
    tvSiteFiles.Drag vbEndDrag
  End If
  
  Err.Clear
End Sub

Private Sub tvSiteFiles_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  If mDrag Then
    Set tvSiteFiles.DropHighlight = tvSiteFiles.HitTest(x, y)
  End If
End Sub

Private Sub tvSiteFiles_Expand(ByVal Node As MSComctlLib.Node)
Dim lSite As clsSite
  On Error Resume Next
  If Node.Image <> "ASP" And Node.Image <> "HTML" And Node.Image <> "HTM" Or Node.Image <> "JS" Or Node.Image <> "DEFAULT" Then
    If Node.Image = "FOLDERCLOSE" Then
      Node.Image = "FOLDEROPEN"
    ElseIf Node.Image = "FOLDERCLOSEONLINE" Then
      Node.Image = "FOLDEROPENONLINE"
    End If
    If tvSiteFiles.Nodes("F0").Tag = "R" Then
      tvSiteFiles.Nodes.Remove Node.Key & " "
      LoadRemoteFiles Node.Key, Node.Key
    Else
      LoadFiles tvSiteFiles, Node.Key, Node.Key, True
    End If
  End If
  
  Err.Clear
End Sub

Private Sub tvSiteFiles_GotFocus()
  On Error Resume Next
  If tvSiteFiles.Nodes.Count > 0 Then
    If cboSites.Text <> tvSiteFiles.Tag Then
      cboSites.Text = mSitename
    End If
  End If
End Sub

Private Sub tvSiteFiles_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lNode As Node
  If KeyCode = vbKeyReturn Then 'Open file
    tvSiteFiles_DblClick
  ElseIf KeyCode = vbKeyF2 Then 'Rename file
    Set lNode = tvSiteFiles.SelectedItem
    If Not lNode Is Nothing Then
      If lNode.Key <> "F0" Then
        tvSiteFiles.Tag = lNode.Text
        tvSiteFiles.StartLabelEdit
      End If
    End If
    KeyCode = 0
  ElseIf KeyCode = vbKeyDelete Then 'Delete file
    Set lNode = tvSiteFiles.SelectedItem
    If Not lNode Is Nothing Then
      If InStr("DES MYD MYCOMPUTER MYN FLOPPY DISK CDROM", lNode.Image) = 0 Then
        If MsgBox("Are you sure to delete '" & lNode.Text & "'?", vbQuestion + vbYesNo + vbDefaultButton2, Mtitle) = vbYes Then
          If S105_Delete(lNode.Key, IIf(lNode.Image = "FOLDERCLOSE" Or lNode.Image = "FOLDEROPEN", True, False)) Then
            tvSiteFiles.Nodes.Remove lNode.Key
          Else
            MsgBox "File/Folder cannot be deteted. Access Denied.", vbInformation, Mtitle
          End If
        End If
      End If
    End If
    KeyCode = 0
  ElseIf KeyCode = 93 Then 'Property key
    'show the right click menu
  End If
  If mDrag Then
    If Shift = 2 Then
      mCopy = True
      Set tvSiteFiles.DragIcon = imlFiles.ListImages("DRAGCOPY").Picture
    Else
      mCopy = False
      Set tvSiteFiles.DragIcon = imlFiles.ListImages("DRAG").Picture
    End If
  End If
End Sub

Private Sub tvSiteFiles_KeyUp(KeyCode As Integer, Shift As Integer)
  If mDrag Then
    'If Shift = 2 Then
      'mCopy = True
      'Set tvSiteFiles.DragIcon = imlFiles.ListImages("DRAGCOPY").Picture
    'Else
      mCopy = False
      Set tvSiteFiles.DragIcon = imlFiles.ListImages("DRAG").Picture
    'End If
  End If
End Sub

Private Sub tvSiteFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    If Not tvSiteFiles.HitTest(x, y) Is Nothing Then
      tvSiteFiles.HitTest(x, y).Selected = True
      If tvSiteFiles.HitTest(x, y).Key = "F1" Then
        mSource = ""
        mSourceType = ""
      Else
        mSource = tvSiteFiles.HitTest(x, y).Tag
        mSourceType = tvSiteFiles.HitTest(x, y).Image
      End If
      If Shift = 2 Then
        mCopy = True
      Else
        mCopy = False
      End If
    End If
  End If
End Sub

Private Sub tvSiteFiles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    If tvSiteFiles.Nodes.Count > 0 Then
      If tvSiteFiles.Nodes("F0").Tag = "R" Then 'For remote site
        mDrag = False
        Exit Sub
      End If
      mDrag = True
      tvSiteFiles.Drag vbBeginDrag
      If mCopy Then
        Set tvSiteFiles.DragIcon = imlFiles.ListImages("DRAGCOPY").Picture
      Else
        Set tvSiteFiles.DragIcon = imlFiles.ListImages("DRAG").Picture
      End If
    End If
  Else
    mDrag = False
    tvSiteFiles.Drag vbEndDrag
  End If
End Sub

Private Sub tvSiteFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  mDrag = False
  mSource = ""
  mSourceType = ""
  mCopy = False
  If Button = vbRightButton Then
    If tvSiteFiles.Nodes.Count > 0 Then
      If Not tvSiteFiles.SelectedItem Is Nothing Then
        EnableMenuWS tvSiteFiles.SelectedItem.Image, True, cboSites.Text, Not tvSiteFiles.Nodes("F0").Tag <> "R"
      Else
        EnableMenuWS "", True, cboSites.Text, Not tvSiteFiles.Nodes("F0").Tag <> "R"
      End If
    End If
    PopupMenu mnuFile, , picTools.Left + picSite.Left + tvSiteFiles.Left + x, picTools.Top + picSite.Top + tvSiteFiles.Top + y
  End If
End Sub

Rem --------------------------------
Rem         User Functions
Rem --------------------------------

Private Function AddFileToMenu(ByVal pFilename As String)
'
'Add the open file to the recent files of window menu
'
  On Error Resume Next
  If tabMain.Tabs.Count = 1 Then
    mStandardMenu.RemoveItem "mnuNo"
  End If
  mStandardMenu.AddItem Mid(pFilename, InStrRev(pFilename, "\") + 1), , , mStandardMenu.IndexForKey("mnuWindows"), , , , "file" & pFilename
End Function

Private Sub BuildMenu()
'
'Build the menus
'
Dim iP As Long
Dim iP2 As Long
  Set mStandardMenu = New cPopupMenu
  With mStandardMenu
    .ImageList = imlMenu
    .hWndOwner = Me.hwnd
    .OfficeXpStyle = True
    'Files
    iP = .AddItem("&File", , , , , , , "mnuFile")
      .AddItem "&New" & vbTab & "Ctrl+N", , , iP, GetIndex("NEW"), , , "mnuNew"
      .AddItem "&Open" & vbTab & "Ctrl+O", , , iP, GetIndex("OPEN"), , , "mnuOpen"
      .AddItem "-", , , iP
      .AddItem "&Save" & vbTab & "Ctrl+S", , , iP, GetIndex("SAVE"), , , "mnuSave"
      .AddItem "Save As...", , , iP, , , , "mnuSaveAs"
      .AddItem "-", , , iP
      .AddItem "&Close", , , iP, , , , "mnuClose"
      .AddItem "Close All", , , iP, , , , "mnuCloseAll"
      .AddItem "-", , , iP
      .AddItem "&Print...", , , iP, GetIndex("PRINT"), , , "mnuPrint"
      .AddItem "-", , , iP
      .AddItem "E&xit" & vbTab & "Ctrl+Q", , , iP, , , , "mnuExit"
    'Edit
    iP = .AddItem("&Edit", , , , , , , "mnuEdit")
      .AddItem "&Redo" & vbTab & "Ctrl+Y", , , iP, GetIndex("REDO"), , , "mnuRedo"
      .AddItem "&Undo" & vbTab & "Ctrl+Z", , , iP, GetIndex("UNDO"), , , "mnuUndo"
      .AddItem "-", , , iP
      .AddItem "Cu&t", , , iP, GetIndex("CUT"), , , "mnuCut"
      .AddItem "&Copy", , , iP, GetIndex("COPY"), , , "mnuCopy"
      .AddItem "&Paste", , , iP, GetIndex("PASTE"), , , "mnuPaste"
      .AddItem "&Delete", , , iP, , , , "mnuDelete"
      .AddItem "-", , , iP
      .AddItem "Select All" & vbTab & "Ctrl+A", , , iP, , , , "mnuSelectAll"
      .AddItem "-", , , iP
      .AddItem "&Find and Replace..." & vbTab & "Ctrl+F", , , iP, GetIndex("SEARCH"), , , "mnuFind"
      .AddItem "&Goto Line..." & vbTab & "Ctrl+G", , , iP, , , , "mnuGotoline"
    'Views
    iP = .AddItem("&View", , , , , , , "mnuViewM")
      .AddItem "Fullmode Preview", , , iP, GetIndex("FULLMODE"), mFullmodePreview, , "mnuFullmodePreview"
      .AddItem "Preview in Browser" & vbTab & "F12", GetIndex("BROWSERPREVIEW"), , iP, , , , "mnuWebPreview"
      .AddItem "-", , , iP
      .AddItem "Refresh Toolbox" & vbTab & "F8", , , iP, GetIndex("REFRESHTOOLBOX"), , , "mnuRefreshToolbox"
      .AddItem "-", , , iP
      .AddItem "Code" & vbTab & "F4", , , iP, GetIndex("CODE"), , , "mnuCode"
      .AddItem "View" & vbTab & "F5", , , iP, GetIndex("VIEW"), , , "mnuView"
      .AddItem "Code/View" & vbTab & "F6", , , iP, GetIndex("CODEVIEW"), , , "mnuCodeView"
    'Inserts
    iP = .AddItem("&Insert", , , , , , , "mnuInsert")
      .AddItem "Hyperlink", , , iP, GetIndex("LINK"), , , "mnuLink"
      .AddItem "Named Anchor", , , iP, GetIndex("BOOKMARK"), , , "mnuBookMark"
      .AddItem "Email Link", , , iP, , , , "mnuEmailLink"
      .AddItem "-", , , iP
      .AddItem "&Image...", , , iP, GetIndex("IMAGE"), , , "mnuImage"
      .AddItem "&Rollover Image...", , , iP, , , , "mnuRolloverImage"
      .AddItem "-", , , iP
      .AddItem "&Form", , , iP, , , , "mnuForm"
      iP2 = .AddItem("&Form Objects", , , iP, , , , "mnuInput")
        .AddItem "Text Box", , , iP2, GetIndex("TEXTBOX"), , , "mnuTextbox"
        .AddItem "Text Area", , , iP2, GetIndex("TEXTAREA"), , , "mnuTextarea"
        .AddItem "Submit Button", , , iP2, GetIndex("SUBMIT"), , , "mnuSubmitButton"
        .AddItem "Reset Button", , , iP2, GetIndex("RESET"), , , "mnuResetButton"
        .AddItem "Hidden Box", , , iP2, GetIndex("HIDDEN"), , , "mnuHiddenBox"
        .AddItem "L&ist/Menu", , , iP2, GetIndex("COMBO"), , , "mnuList"
        .AddItem "&Option Button", , , iP2, GetIndex("OPTION"), , , "mnuOptionbutton"
        .AddItem "&Push Button", , , iP2, GetIndex("BUTTON"), , , "mnuPushbutton"
        .AddItem "&Check Box", , , iP2, GetIndex("CHECK"), , , "mnuCheckbox"
        .AddItem "&Label", , , iP2, GetIndex("LABEL"), , , "mnuLabel"
        .AddItem "-", , , iP
        .AddItem "&Table", , , iP, GetIndex("TABLE"), , , "mnuTables"
        .AddItem "&Marquee", , , iP, , , , "mnuMarquee"
        .AddItem "&Div", , , iP, , , , "mnuDiv"
        .AddItem "&Span", , , iP, , , , "mnuSpan"
        .AddItem "-", , , iP
        .AddItem "Date", , , iP, , , , "mnuDate"
        .AddItem "Horizontal Rule", , , iP, , , , "mnuHorizontalRule"
        .AddItem "-", , , iP
      iP2 = .AddItem("Head Tags", , , iP, , , , "mnuHeadTags")
        .AddItem "Meta", , , iP2, , , , "mnuMeta"
        .AddItem "Keywords", , , iP2, , , , "mnuKeywords"
        .AddItem "Description", , , iP2, , , , "mnuDescription"
        .AddItem "Refresh", , , iP2, , , , "mnuRefresh"
        .AddItem "Base", , , iP2, , , , "mnuBase"
        .AddItem "Link", , , iP2, , , , "mnuMetaLink"
      iP2 = .AddItem("Script Block", , , iP, , , , "mnuScripBlock")
        .AddItem "Client", , , iP2, , , , "mnuClient"
        .AddItem "Server", , , iP2, , , , "mnuServer"
        .AddItem "-", , , iP2
        .AddItem "Validation..." & vbTab & "F11", , , iP2, , , , "mnuFormvalidation"
        .AddItem "Set Default Value..." & vbTab & "F10", , , iP2, , , , "mnuDefaultValue"
      iP2 = .AddItem("Special Characters", , , iP, , , , "mnuSpecialCharacters")
        .AddItem "Line Break", , , iP2, , , , "mnuLineBreak"
        .AddItem "Non-Breaking Space", , , iP2, , , , "mnuNBSP"
        .AddItem "-", , , iP2
        .AddItem "Copyright", , , iP2, , , , "mnuCopyright"
        .AddItem "Registered", , , iP2, , , , "mnuRegistered"
        .AddItem "Trade Mark", , , iP2, , , , "mnuTrademark"
        .AddItem "Pound", , , iP2, , , , "mnuPound"
        .AddItem "Yen", , , iP2, , , , "mnuYen"
        .AddItem "Euro", , , iP2, , , , "mnuEuro"
        .AddItem "Left Quote", , , iP2, , , , "mnuLeftQuote"
        .AddItem "Right Quote", , , iP2, , , , "mnuRightQuote"
        .AddItem "Em-Dash", , , iP2, , , , "mnuEmDash"
        .AddItem "Others...", , , iP2, , , , "mnuOthersSC"
    'Texts
    iP = .AddItem("&Text", , , , , , , "mnuTextMain")
      .AddItem "Indent" & vbTab & "Ctrl+Shift+I", , , iP, GetIndex("INDENT"), , , "mnuIndent"
      .AddItem "Outdent" & vbTab & "Ctrl+Shift+O", , , iP, GetIndex("OUTDENT"), , , "mnuOutdent"
      .AddItem "-", , , iP
      iP2 = .AddItem("Paragraph Format", , , iP, , , , "mnuParagraphFormat")
        .AddItem "Paragraph", , , iP2, , , , "mnuParagraph"
        .AddItem "Heading 1" & vbTab & "Ctrl+1", , , iP2, , , , "mnuHeading1"
        .AddItem "Heading 2" & vbTab & "Ctrl+2", , , iP2, , , , "mnuHeading2"
        .AddItem "Heading 3" & vbTab & "Ctrl+3", , , iP2, , , , "mnuHeading3"
        .AddItem "Heading 4" & vbTab & "Ctrl+4", , , iP2, , , , "mnuHeading4"
        .AddItem "Heading 5" & vbTab & "Ctrl+5", , , iP2, , , , "mnuHeading5"
        .AddItem "Heading 6" & vbTab & "Ctrl+6", , , iP2, , , , "mnuHeading6"
        .AddItem "Preformatted Text", , , iP2, , , , "mnuPretext"
      iP2 = .AddItem("List", , , iP, , , , "mnuListText")
        .AddItem "Unordered List", , , iP2, , , , "mnuUnorderedList"
        .AddItem "Ordered List", , , iP2, , , , "mnuOrderedList"
        .AddItem "Definition List", , , iP2, , , , "mnuDefinitionList"
      .AddItem "-", , , iP
      iP2 = .AddItem("Font", , , iP, , , , "mnuFont")
        LoadFontsMenu
      iP2 = .AddItem("Styles", , , iP, , , , "mnuStyles")
        .AddItem "Bold" & vbTab & "Ctrl+B", , , iP2, GetIndex("BOLD"), , , "mnuBold"
        .AddItem "Italic" & vbTab & "Ctrl+I", , , iP2, GetIndex("ITALIC"), , , "mnuItalic"
        .AddItem "Underline" & vbTab & "Ctrl+U", , , iP2, GetIndex("UNDERLINE"), , , "mnuUnderline"
        .AddItem "-", , , iP2
        .AddItem "Superscript", , , iP2, , , , "mnuSuperscript"
        .AddItem "Subscript", , , iP2, , , , "mnuSubscript"
        .AddItem "-", , , iP2
        .AddItem "Left", , , iP2, GetIndex("LEFT"), , , "mnuLeft"
        .AddItem "Center", , , iP2, GetIndex("CENTEER"), , , "mnuCenter"
        .AddItem "Right", , , iP2, GetIndex("RIGHT"), , , "mnuRight"
        .AddItem "-", , , iP2
        .AddItem "Strikethrough", , , iP2, , , , "mnuStrikethrough"
        .AddItem "Teletype", , , iP2, , , , "mnuTeletype"
        .AddItem "Emphasis", , , iP2, , , , "mnuEmphasis"
        .AddItem "Strong", , , iP2, , , , "mnuStrong"
        .AddItem "-", , , iP2
        .AddItem "Code", , , iP2, , , , "mnuCode"
        .AddItem "Variable", , , iP2, , , , "mnuVariable"
        .AddItem "Sample", , , iP2, , , , "mnuSample"
        .AddItem "Keyboard", , , iP2, , , , "mnuKeyboard"
        .AddItem "-", , , iP2
        .AddItem "Citation", , , iP2, , , , "mnuCitetion"
        .AddItem "Definition", , , iP2, , , , "mnuDefinition"
        .AddItem "Deleted", , , iP2, , , , "mnuDeleted"
        .AddItem "Inserted", , , iP2, , , , "mnuInserted"
      iP2 = .AddItem("CSS Styles", , , iP, , , , "mnuCSSStyles")
        .AddItem "Attach style sheet...", , , iP2, , , , "mnuStylesheet"
        .AddItem "Style sheet editor..." & vbTab & "Ctrl+Shift+E", , , iP2, , , , "mnuCSS"
        .AddItem "-", , , iP2
      iP2 = .AddItem("Size", , , iP, , , , "mnuSize")
        .AddItem "1", , , iP2, , , , "mnu1"
        .AddItem "2", , , iP2, , , , "mnu2"
        .AddItem "3", , , iP2, , , , "mnu3"
        .AddItem "4", , , iP2, , , , "mnu4"
        .AddItem "5", , , iP2, , , , "mnu5"
        .AddItem "6", , , iP2, , , , "mnu6"
        .AddItem "7", , , iP2, , , , "mnu7"
      iP2 = .AddItem("Size Change", , , iP, , , , "mnuSizeChange")
        .AddItem "+1", , , iP2, , , , "mnu+1"
        .AddItem "+2", , , iP2, , , , "mnu+2"
        .AddItem "+3", , , iP2, , , , "mnu+3"
        .AddItem "+4", , , iP2, , , , "mnu+4"
        .AddItem "-", , , iP2
        .AddItem Space(1) & Chr("0173") & "1", , , iP2, , , , "mnu-1" 'Char('0173') equal to - to avoid separator
        .AddItem Space(1) & Chr("0173") & "2", , , iP2, , , , "mnu-2"
        .AddItem Space(1) & Chr("0173") & "3", , , iP2, , , , "mnu-3"
    'Wizards
    iP = .AddItem("&Wizards", , , , , , , "mnuWizards")
      .AddItem "&DSN Connection", , , iP, GetIndex("DSNCON"), , , "mnuDSNConnection"
      .AddItem "&Database Connection", , , iP, GetIndex("DBCON"), , , "mnuDBConnection"
      .AddItem "&Cookie", , , iP, GetIndex("COOKIES"), , , "mnuCookie"
    'Tools
    iP = .AddItem("Tools", , , , , , , "mnuTools")
      .AddItem "Site", , , iP, GetIndex("SITE"), , , "mnuSiteShow"
      .AddItem "New site", , , iP, , , , "mnuNewSite"
      .AddItem "Edit site", , , iP, , , , "mnuEditSite"
      .AddItem "-", , , iP
      .AddItem "History", , , iP, GetIndex("HISTORY"), , , "mnuHistory"
      .AddItem "Clear History", , , iP, GetIndex("CLEARHISTORY"), , , "mnuClearHistory"
      .AddItem "-", , , iP
      .AddItem "Data Explorer", , , iP, GetIndex("DATAEXPLORER"), , , "mnuDataExplorer"
      .AddItem "New connection", , , iP, , , , "mnuNewConnection"
    'Options
    iP = .AddItem("Options", , , , , , , "mnuOptions")
      .AddItem "Auto Completion", , , iP, , mAutoCompletion, , "mnuAutoCompletion"
      .AddItem "Syntax Highlighting", , , iP, , mSyntaxHighlighting, , "mnuSyntaxHighlighting"
      .AddItem "Intelisense", , , iP, , mIntelisense, , "mnuIntelisense"
      .AddItem "Line Number" & vbTab & "Ctrl+L", , , iP, , mLineNo, , "mnuLinenumber"
      .AddItem "-", , , iP
      .AddItem "Word Wrap", , , iP, , mWordWrap, , "mnuWordwrap"
      .AddItem "-", , , iP
      .AddItem "Toolbox" & vbTab & "Ctrl+T", , , iP, GetIndex("TOOLBOX"), mToolbox, , "mnuToolbox"
      .AddItem "Open Default Document", , , iP, , Mopendialog, , "mnuOpenDialog"
    'Windows
    iP = .AddItem("Wi&ndows", , , , , , , "mnuWindows")
      .AddItem "  No Documents", , , iP, , , False, "mnuNo"
    iP = .AddItem("Help", , , , , , , "mnuHelp")
      .AddItem "About...", , , iP, , , , "mnuAboutUs"
      If GetSetting(App.Title, "regkey", "regkey") = "" Then
        .AddItem "-", , , iP
        .AddItem "Registration...", , , iP, , , , "mnuRegistration"
      End If
  End With
End Sub

Private Sub BuildToolBar()
'
'Build the menubar and toolbar
'
  With ctbHeader
    .ImageSource = CTBExternalImageList
    .SetImageList imlMenu, CTBImageListNormal
    .ImageStandardBitmapType = CTBHistorySmallColor
    .DrawStyle = CTBDrawOfficeXPStyle
    .CreateToolbar 16, True, True, True
    .AddButton "New", GetIndex("NEW"), , , "", CTBAutoSize, "New"
    .AddButton "Open", GetIndex("OPEN"), , , "", CTBAutoSize, "Open"
    .AddButton "Save", GetIndex("SAVE"), , , "", CTBAutoSize, "Save"
    .AddButton , , , , , CTBSeparator
    .AddButton "Cut", GetIndex("CUT"), , , "", CTBAutoSize, "Cut"
    .AddButton "Copy", GetIndex("COPY"), , , "", CTBAutoSize, "Copy"
    .AddButton "Paste", GetIndex("PASTE"), , , "", CTBAutoSize, "Paste"
    .AddButton , , , , , CTBSeparator
    .AddButton "Undo", GetIndex("UNDO"), , , "", CTBAutoSize, "Undo"
    .AddButton "Redo", GetIndex("REDO"), , , "", CTBAutoSize, "Redo"
    .AddButton , , , , , CTBSeparator
    .AddButton "Find", GetIndex("SEARCH"), , , "", CTBAutoSize, "Find"
    .AddControl cboFind.hwnd
    .AddButton , , , , , CTBSeparator
    .AddButton "Bold", GetIndex("BOLD"), , , "", CTBAutoSize, "Bold"
    .AddButton "Italics", GetIndex("ITALIC"), , , "", CTBAutoSize, "Italics"
    .AddButton "Underline", GetIndex("UNDERLINE"), , , "", CTBAutoSize, "Underline"
    .AddButton , , , , , CTBSeparator
    .AddButton "Left", GetIndex("LEFT"), , , "", CTBAutoSize, "Left"
    .AddButton "Center", GetIndex("CENTER"), , , "", CTBAutoSize, "Center"
    .AddButton "Right", GetIndex("RIGHT"), , , "", CTBAutoSize, "Right"
    .AddButton , , , , , CTBSeparator
    .AddButton "Indent", GetIndex("INDENT"), , , "", CTBAutoSize, "Indent"
    .AddButton "Unindent", GetIndex("OUTDENT"), , , "", CTBAutoSize, "Unindent"
    .AddButton , , , , , CTBSeparator
    .AddButton "Table", GetIndex("TABLE"), , , "", CTBAutoSize, "Table"
    .AddButton "Link", GetIndex("LINK"), , , "", CTBAutoSize, "Link"
    .AddButton "Image", GetIndex("IMAGE"), , , "", CTBAutoSize, "Image"
    .Visible = True
    .Wrappable = True
  End With
  
  Call BuildMenu
  
  With tbrMenu
    .CreateFromMenu mStandardMenu
    .Wrappable = True
    .DrawStyle = CTBDrawOfficeXPStyle
    On Error Resume Next
    SetWindowTheme tbrMenu.hwnd, StrPtr(" "), StrPtr(" ")
    On Error GoTo 0
  End With
  
  With tbhMenu
    .MDIToolbar = True
    .Capture tbrMenu
    .Width = tbhMenu.MDIToolbarMinWidth * Screen.TwipsPerPixelX
  End With
  
  With crbHeader
    .CreateRebar Me.hwnd
    'Top Menu
    .AddBandByHwnd tbhMenu.hwnd, , , , "MenuBar"
    .BandChildMinWidth(0) = 64
    'Toolbar
    .AddBandByHwnd ctbHeader.hwnd, , , , "Toolbar"
    .BandChildMinWidth(crbHeader.BandCount - 1) = 24
    .Visible = True
  End With
End Sub

Private Function ClearHistory()
'
'Clear the history
'
  If MsgBox("Are you sure to clear the history?", vbQuestion + vbYesNo, Mtitle) = vbYes Then
    If S102_File_Exists(App.Path & "\history.dat") Then
      Call S105_Delete(App.Path & "\history.dat")
      LoadHistory
      MsgBox "Successfully cleared.", vbInformation, Mtitle
    End If
  End If
End Function

Private Function CloseAllDocuments()
'
'Close all opened documents
'
Dim li As Integer
  On Error Resume Next
  For li = tabMain.Tabs.Count To 1 Step -1
    If RTB(tabMain.Tabs.Item(li).Tag).IsHistory = False Then SaveHistory RTB(tabMain.Tabs.Item(li).Tag).FileName
    Unload RTB(tabMain.Tabs.Item(li).Tag)
    tabMain.Tabs.Remove li
  Next
  tabMain.Visible = False
  
  Err.Clear
End Function

Private Function CloseDocument()
'
'Close the active opened document
'
  On Error Resume Next
  If RTB(tabMain.SelectedTab.Tag).IsHistory = False Then SaveHistory RTB(tabMain.SelectedTab.Tag).FileName
  Unload RTB(tabMain.SelectedTab.Tag)
  tabMain.Tabs.Remove tabMain.SelectedTab.Key
  
  Err.Clear
End Function

Private Sub CollapseTool(ByVal pIndex As Integer)
'
'Expand the clicked property bar in the toolbox
'
Dim lHeight As Single
Dim lTools As Integer
  If pIndex = 0 Then pIndex = 1
  lTools = 4
  'Workspace
  picWorkSpace.Height = picWSProperty.Height
  picWSProperty.BackColor = RGB(220, 217, 200)
  lblWorkspace.ForeColor = vbBlack
  'History
  picHistory.Height = picHProperty.Height
  picHProperty.BackColor = RGB(220, 217, 200)
  lblHistory.ForeColor = vbBlack
  'Site
  picSite.Height = picSProperty.Height
  picSProperty.BackColor = RGB(220, 217, 200)
  lblSite.ForeColor = vbBlack
  'Application
  picApplication.Height = picAProperty.Height
  picAProperty.BackColor = RGB(220, 217, 200)
  lblApplication.ForeColor = vbBlack
  'Selected band height
  lHeight = picTools.Height - (picWorkSpace.Height * lTools) - Screen.TwipsPerPixelY * (lTools + 3)
  Select Case pIndex
  Case 1 'Workspace
    picWSProperty.BackColor = RGB(172, 168, 153)
    lblWorkspace.ForeColor = vbWhite
    picWorkSpace.Height = lHeight
  Case 2 'History
    picHProperty.BackColor = RGB(172, 168, 153)
    lblHistory.ForeColor = vbWhite
    picHistory.Height = lHeight
  Case 3 'Site
    picSProperty.BackColor = RGB(172, 168, 153)
    lblSite.ForeColor = vbWhite
    picSite.Height = lHeight
  Case 4 'Application
    picAProperty.BackColor = RGB(172, 168, 153)
    lblApplication.ForeColor = vbWhite
    picApplication.Height = lHeight
  End Select
  ResizeTool
End Sub

Private Function DeleteConnectionString()
'
'Delete the connection details
'
Dim lConn As String
Dim li As Integer
Dim lSite As clsSite
Dim lResult As Boolean
  Set lSite = Msitedetails.Item(cboSites.Text)
  If Not lSite Is Nothing Then
    If lSite.ConnectionString.Count > 0 Then
      lConn = tvApplication.SelectedItem.Text & "~" & tvApplication.SelectedItem.Tag
      For li = 1 To lSite.ConnectionString.Count
        If lConn = lSite.ConnectionString(li) Then
          lSite.ConnectionString.Remove li
          lResult = True
        End If
      Next
      If lResult Then Msitedetails.Save
    End If
  End If
  tvApplication.Nodes.Remove tvApplication.SelectedItem.Key
End Function

Private Function DownloadFile(ByVal pRemoteFile As String, ByVal pLocalPath As String) As String
'
'Download the file from remote for editing/view
'Remotefile(path+file) , pLocalpath (Path only)
'if successfully downloaded, it returns the local filename
'else nullstring
'
Dim lConnected As Boolean
Dim lFile As String
Dim lPath As String
  RaiseProgress 0, 0, "Downloading... " & Mid(pRemoteFile, InStrRev(pRemoteFile, "/") + 1)
  'Test for connection
  lConnected = mFTP.IsConnected
  If lConnected = False Then
    lConnected = mFTP.OpenConnection(mServer, mUsername, mPassword)
  End If
  If lConnected Then
    'Get remote path/file
    lPath = Mid(pRemoteFile, 1, InStrRev(pRemoteFile, "/") - 1)
    lFile = Mid(pRemoteFile, InStrRev(pRemoteFile, "/") + 1)
    If Left(lPath, 1) = "/" Then lPath = Mid(lPath, 2)
    'Set local file
    If Right(pLocalPath, 1) <> "\" Then pLocalPath = pLocalPath & "\"
    pLocalPath = pLocalPath & lPath & "\"
    pLocalPath = Replace(pLocalPath, "/", "\")
    S101_Make_Dir pLocalPath
    pLocalPath = pLocalPath & lFile
    'Download the file
    If mFTP.SetFTPDirectory("/\/" & lPath) Then
      If mFTP.FTPDownloadFile(pLocalPath, lFile) Then
        DownloadFile = pLocalPath
        RaiseProgress 0, 0, "Done"
      Else
        DownloadFile = ""
        RaiseProgress 0, 0, "Error on download... " & lFile
      End If
    End If
  End If
  
End Function

Private Sub EnableMenuDE(Optional ByVal pEnable As Boolean = True)
'
'Enable the dataexpoler menus
'
Dim lNode As Object
  mnuEditConnection.Enabled = False
  mnuDeleteConnection.Enabled = False
  mnuViewRecords.Enabled = False
  mnuInsertCode.Enabled = False
  mnuTestConnection.Enabled = False
  mnuInsertGC.Enabled = False
  mnuSelectGC.Enabled = False
  mnuUpdateGC.Enabled = False
  mnuDeleteGC.Enabled = False
  mnuInsertGP.Enabled = False
  mnuSelectGP.Enabled = False
  mnuUpdateGP.Enabled = False
  mnuDeleteGP.Enabled = False
  mnuInsertGF.Enabled = False
  mnuSelectGF.Enabled = False
  mnuUpdateGF.Enabled = False
  mnuDeleteGF.Enabled = False
  Set lNode = tvApplication.SelectedItem
  If Not lNode Is Nothing Then
    If mEditorIndex > 0 Then mnuInsertCode.Enabled = True
    Select Case lNode.Image
    Case "FOL"
      mnuInsertCode.Enabled = False
    Case "TABLE"
      mnuViewRecords.Enabled = True
      mnuInsertGC.Enabled = True
      mnuSelectGC.Enabled = True
      mnuUpdateGC.Enabled = True
      mnuDeleteGC.Enabled = True
      mnuInsertGP.Enabled = True
      mnuSelectGP.Enabled = True
      mnuUpdateGP.Enabled = True
      mnuDeleteGP.Enabled = True
      mnuInsertGF.Enabled = True
      mnuSelectGF.Enabled = True
      mnuUpdateGF.Enabled = True
      mnuDeleteGF.Enabled = True
    Case "CONNECTION"
      mnuEditConnection.Enabled = True
      mnuTestConnection.Enabled = True
      mnuDeleteConnection.Enabled = True
    End Select
  End If
End Sub

Private Function EnableMenuMM(ByVal pEnable As Boolean)
'
'Enable the main menu
'
Dim li As Integer
  On Error Resume Next
  'Menus
  With mStandardMenu
    'File menu
      .Enabled(.IndexForKey("mnuSave")) = pEnable
      .Enabled(.IndexForKey("mnuSaveAs")) = pEnable
      .Enabled(.IndexForKey("mnuClose")) = pEnable
      .Enabled(.IndexForKey("mnuCloseAll")) = pEnable
      .Enabled(.IndexForKey("mnuPrint")) = pEnable
    'Edit menu
      If pEnable = False Then .Enabled(.IndexForKey("mnuRedo")) = pEnable
      If pEnable = False Then .Enabled(.IndexForKey("mnuUndo")) = pEnable
      .Enabled(.IndexForKey("mnuCopy")) = pEnable
      .Enabled(.IndexForKey("mnuCut")) = pEnable
      If pEnable = False Then .Enabled(.IndexForKey("mnuPaste")) = pEnable
      .Enabled(.IndexForKey("mnuDelete")) = pEnable
      .Enabled(.IndexForKey("mnuSelectAll")) = pEnable
      .Enabled(.IndexForKey("mnuFind")) = pEnable
      .Enabled(.IndexForKey("mnuGotoline")) = pEnable
    'View menu
      .Enabled(.IndexForKey("mnuWebPreview")) = pEnable
      .Enabled(.IndexForKey("mnuCode")) = pEnable
      .Enabled(.IndexForKey("mnuView")) = pEnable
      .Enabled(.IndexForKey("mnuCode/View")) = pEnable
    'Insert
      .Enabled(.IndexForKey("mnuImage")) = pEnable
      .Enabled(.IndexForKey("mnuRolloverImage")) = pEnable
      .Enabled(.IndexForKey("mnuLink")) = pEnable
      .Enabled(.IndexForKey("mnuEmailLink")) = pEnable
      .Enabled(.IndexForKey("mnuHorizontalRule")) = pEnable
      .Enabled(.IndexForKey("mnuDate")) = pEnable
      .Enabled(.IndexForKey("mnuTables")) = pEnable
      .Enabled(.IndexForKey("mnuCSS")) = pEnable
      .Enabled(.IndexForKey("mnuIndent")) = pEnable
      .Enabled(.IndexForKey("mnuOutdent")) = pEnable
      .Enabled(.IndexForKey("mnuParagraph")) = pEnable
      .Enabled(.IndexForKey("mnuHeading1")) = pEnable
      .Enabled(.IndexForKey("mnuHeading2")) = pEnable
      .Enabled(.IndexForKey("mnuHeading3")) = pEnable
      .Enabled(.IndexForKey("mnuHeading4")) = pEnable
      .Enabled(.IndexForKey("mnuHeading5")) = pEnable
      .Enabled(.IndexForKey("mnuHeading6")) = pEnable
      .Enabled(.IndexForKey("mnuPretext")) = pEnable
      .Enabled(.IndexForKey("mnuUnorderedList")) = pEnable
      .Enabled(.IndexForKey("mnuOrderedList")) = pEnable
      .Enabled(.IndexForKey("mnuDefinitionList")) = pEnable
      .Enabled(.IndexForKey("mnuDefaultFont")) = pEnable
      .Enabled(.IndexForKey("mnuOtherFonts")) = pEnable
      .Enabled(.IndexForKey("mnuStrikethrough")) = pEnable
      .Enabled(.IndexForKey("mnuTeletype")) = pEnable
      .Enabled(.IndexForKey("mnuEmphasis")) = pEnable
      .Enabled(.IndexForKey("mnuStrong")) = pEnable
      .Enabled(.IndexForKey("mnuCode")) = pEnable
      .Enabled(.IndexForKey("mnuVariable")) = pEnable
      .Enabled(.IndexForKey("mnuSample")) = pEnable
      .Enabled(.IndexForKey("mnuKeyboard")) = pEnable
      .Enabled(.IndexForKey("mnuCitetion")) = pEnable
      .Enabled(.IndexForKey("mnuDefinition")) = pEnable
      .Enabled(.IndexForKey("mnuDeleted")) = pEnable
      .Enabled(.IndexForKey("mnuInserted")) = pEnable
      .Enabled(.IndexForKey("mnu1")) = pEnable
      .Enabled(.IndexForKey("mnu2")) = pEnable
      .Enabled(.IndexForKey("mnu3")) = pEnable
      .Enabled(.IndexForKey("mnu4")) = pEnable
      .Enabled(.IndexForKey("mnu5")) = pEnable
      .Enabled(.IndexForKey("mnu6")) = pEnable
      .Enabled(.IndexForKey("mnu7")) = pEnable
      .Enabled(.IndexForKey("mnu+1")) = pEnable
      .Enabled(.IndexForKey("mnu+2")) = pEnable
      .Enabled(.IndexForKey("mnu+3")) = pEnable
      .Enabled(.IndexForKey("mnu+4")) = pEnable
      .Enabled(.IndexForKey("mnu-1")) = pEnable
      .Enabled(.IndexForKey("mnu-2")) = pEnable
      .Enabled(.IndexForKey("mnu-3")) = pEnable
      .Enabled(.IndexForKey("mnuStylesheet")) = pEnable
      .Enabled(.IndexForKey("mnuMarquee")) = pEnable
      .Enabled(.IndexForKey("mnuSpan")) = pEnable
      .Enabled(.IndexForKey("mnuDiv")) = pEnable
      .Enabled(.IndexForKey("mnuBookMark")) = pEnable
      .Enabled(.IndexForKey("mnuForm")) = pEnable
      .Enabled(.IndexForKey("mnuTextbox")) = pEnable
      .Enabled(.IndexForKey("mnuTextarea")) = pEnable
      .Enabled(.IndexForKey("mnuSubmitButton")) = pEnable
      .Enabled(.IndexForKey("mnuResetButton")) = pEnable
      .Enabled(.IndexForKey("mnuHiddenBox")) = pEnable
      .Enabled(.IndexForKey("mnuOptionbutton")) = pEnable
      .Enabled(.IndexForKey("mnuPushbutton")) = pEnable
      .Enabled(.IndexForKey("mnuList")) = pEnable
      .Enabled(.IndexForKey("mnuCheckbox")) = pEnable
      .Enabled(.IndexForKey("mnuLabel")) = pEnable
      .Enabled(.IndexForKey("mnuClient")) = pEnable
      .Enabled(.IndexForKey("mnuServer")) = pEnable
      .Enabled(.IndexForKey("mnuLineBreak")) = pEnable
      .Enabled(.IndexForKey("mnuNBSP")) = pEnable
      .Enabled(.IndexForKey("mnuCopyright")) = pEnable
      .Enabled(.IndexForKey("mnuRegistered")) = pEnable
      .Enabled(.IndexForKey("mnuTrademark")) = pEnable
      .Enabled(.IndexForKey("mnuYen")) = pEnable
      .Enabled(.IndexForKey("mnuEuro")) = pEnable
      .Enabled(.IndexForKey("mnuRightQuote")) = pEnable
      .Enabled(.IndexForKey("mnuLeftQuote")) = pEnable
      .Enabled(.IndexForKey("mnuOthersSC")) = pEnable
      .Enabled(.IndexForKey("mnuPound")) = pEnable
      .Enabled(.IndexForKey("mnuEmDash")) = pEnable
      .Enabled(.IndexForKey("mnuBold")) = pEnable
      .Enabled(.IndexForKey("mnuItalic")) = pEnable
      .Enabled(.IndexForKey("mnuUnderline")) = pEnable
      .Enabled(.IndexForKey("mnuLeft")) = pEnable
      .Enabled(.IndexForKey("mnuCenter")) = pEnable
      .Enabled(.IndexForKey("mnuRight")) = pEnable
      .Enabled(.IndexForKey("mnuSuperscript")) = pEnable
      .Enabled(.IndexForKey("mnuSubscript")) = pEnable
      .Enabled(.IndexForKey("mnuFormvalidation")) = pEnable
      .Enabled(.IndexForKey("mnuDefaultValue")) = pEnable
      .Enabled(.IndexForKey("mnuMeta")) = pEnable
      .Enabled(.IndexForKey("mnuMetaLink")) = pEnable
      .Enabled(.IndexForKey("mnuKeywords")) = pEnable
      .Enabled(.IndexForKey("mnuBase")) = pEnable
      .Enabled(.IndexForKey("mnuRefresh")) = pEnable
      .Enabled(.IndexForKey("mnuDescription")) = pEnable
      'fonts
      For li = 0 To val(frmFonts.lsFontslist_b.Tag) - 1
        .Enabled(.IndexForKey("font" & li)) = pEnable
      Next
    'Wizards
      .Enabled(.IndexForKey("mnuDSNConnection")) = pEnable
      .Enabled(.IndexForKey("mnuDBConnection")) = pEnable
      .Enabled(.IndexForKey("mnuCookie")) = pEnable
    'Options
      '.Enabled(.IndexForKey("mnuLinenumber")) = pEnable
      '.Enabled(.IndexForKey("mnuIntelisense")) = pEnable
      '.Enabled(.IndexForKey("mnuAutoCompletion")) = pEnable
      '.Enabled(.IndexForKey("mnuSyntaxHighlighting")) = pEnable
      .Enabled(.IndexForKey("mnuWordwrap")) = pEnable
  End With
  'Toolbar
  With ctbHeader
      .ButtonEnabled("Cut") = pEnable
      .ButtonEnabled("Copy") = pEnable
      If pEnable = False Then .ButtonEnabled("Paste") = pEnable
      If pEnable = False Then .ButtonEnabled("Redo") = pEnable
      If pEnable = False Then .ButtonEnabled("Undo") = pEnable
      .ButtonEnabled("Find") = pEnable
      .ButtonEnabled("Bold") = pEnable
      .ButtonEnabled("Italics") = pEnable
      .ButtonEnabled("Underline") = pEnable
      .ButtonEnabled("Left") = pEnable
      .ButtonEnabled("Center") = pEnable
      .ButtonEnabled("Right") = pEnable
      .ButtonEnabled("Indent") = pEnable
      .ButtonEnabled("Unindent") = pEnable
      .ButtonEnabled("Image") = pEnable
      .ButtonEnabled("Link") = pEnable
      .ButtonEnabled("Table") = pEnable
      .ButtonEnabled("Save") = pEnable
      cboFind.Enabled = pEnable
  End With
  mEnable = pEnable
End Function

Private Sub EnableMenuWS(ByVal pImage As String, Optional ByVal pSite As Boolean = False, Optional ByVal pSitename As String, Optional pRemote As Boolean)
'
'Enable the site/workspace menus
'
  mnuAddFile.Enabled = False
  mnuAddFolder.Enabled = False
  mnuOpen.Enabled = False
  mnuRename.Enabled = False
  mnuDelete.Enabled = False
  mnuRenameSite.Enabled = False
  mnuRemoveSite.Enabled = False
  If pRemote = False Then
    Select Case pImage
    Case "DES", "MYCOMPUTER", "MYN", "CDROM"
    Case "DISK", "MYD", "FLOPPY"
      mnuAddFile.Enabled = True
      mnuAddFolder.Enabled = True
    Case "FOLDERCLOSE", "FOLDEROPEN"
      mnuAddFile.Enabled = True
      mnuAddFolder.Enabled = True
      mnuRename.Enabled = True
      mnuDelete.Enabled = True
    Case Else
      mnuOpen.Enabled = True
      mnuRename.Enabled = True
      mnuDelete.Enabled = True
    End Select
  Else
    If InStr("DES,MYCOMPUTER,MYN,CDROM,DISK,MYD,FLOPPY,FOLDERCLOSE,FOLDEROPEN", pImage) = 0 Then
      mnuOpen.Enabled = True
    End If
  End If
  If pSite Then
    myhy2.Visible = True
    mnuNewsite.Visible = True
    mnuRenameSite.Visible = True
    mnuRemoveSite.Visible = True
    If pSitename <> "" And pSitename <> "Define sites..." Then
      mnuRenameSite.Enabled = True
      mnuRemoveSite.Enabled = True
    End If
  Else
    myhy2.Visible = False
    mnuNewsite.Visible = False
    mnuRenameSite.Visible = False
    mnuRemoveSite.Visible = False
  End If
End Sub

Private Function GenerateDelete(ByVal pConnection As String, ByVal pConnectionStr As String, ByVal pTable As String) As String
'
'Generate the insert query
'
Dim li As Long
Dim lCon As Object
Dim lRecordset As Object
Dim Lfields As String
Dim Lsql As String
Dim lTmp As String
  Lfields = ""
  Screen.MousePointer = vbHourglass
  Set lCon = CreateObject("ADODB.Connection")
  lCon.CursorLocation = 3 'Useclient
  Call lCon.Open(pConnectionStr)
  If lCon.State = 1 Then ' adStateOpen
    Set lRecordset = CreateObject("ADODB.Recordset") ' CreateObject("ADODB.Recordset") ' New ADODB.Recordset
    lRecordset.CursorType = 2 ' adOpenDynamic
    Set lRecordset = lCon.Execute("Select * from " & pTable)
    If lRecordset.Fields.Count > 0 Then
      For li = 0 To lRecordset.Fields.Count - 1
        Lfields = Lfields & "[" & lRecordset.Fields(li).Name & "]='" & quote & " &  & " & quote & "',"
      Next
      Lfields = Trim(Lfields)
      Lfields = UCase(Lfields)
      If Len(Lfields) > 0 Then Lfields = Left(Lfields, Len(Lfields) - 1)
      'make query
      If Lfields <> "" Then
        Lsql = pConnection & ".Execute  " & quote & "delete from " & pTable & " where " & Lfields & quote
      End If
    End If
  End If
  GenerateDelete = Lsql
  Screen.MousePointer = vbDefault
End Function

Private Function GenerateInsert(ByVal pConnection As String, ByVal pConnectionStr As String, ByVal pTable As String) As String
'
'Generate the insert query
'
Dim li As Long
Dim lCon As Object
Dim lRecordset As Object
Dim Lfields As String
Dim lValues As String
Dim Lsql As String
Dim lTmp As String
  Lfields = ""
  lValues = ""
  Screen.MousePointer = vbHourglass
  Set lCon = CreateObject("ADODB.Connection")
  lCon.CursorLocation = 3 'Useclient
  Call lCon.Open(pConnectionStr)
  If lCon.State = 1 Then ' adStateOpen
    Set lRecordset = CreateObject("ADODB.Recordset") ' CreateObject("ADODB.Recordset") ' New ADODB.Recordset
    lRecordset.CursorType = 2 ' adOpenDynamic
    Set lRecordset = lCon.Execute("Select * from " & pTable)
    If lRecordset.Fields.Count > 0 Then
      For li = 0 To lRecordset.Fields.Count - 1
        If lRecordset.Fields(li).Attributes <> 16 Then 'leave Auto increment fields
          Lfields = Lfields & "[" & lRecordset.Fields(li).Name & "],"
          lValues = lValues & "'" & quote & " &  & " & quote & "',"
        End If
      Next
      Lfields = Trim(Lfields)
      Lfields = UCase(Lfields)
      If Len(Lfields) > 0 Then Lfields = Left(Lfields, Len(Lfields) - 1)
      'values
      lValues = Trim(lValues)
      lValues = (lValues)
      If Len(lValues) > 0 Then lValues = Left(lValues, Len(lValues) - 1)
      'make query
      If Lfields <> "" And lValues <> "" Then
        Lsql = pConnection & ".Execute  " & quote & "insert into " & pTable & " (" & Lfields & ") values (" & lValues & ")" & quote
      End If
    End If
  End If
  GenerateInsert = Lsql
  Screen.MousePointer = vbDefault
End Function

Private Function GenerateSelect(ByVal pConnection As String, ByVal pConnectionStr As String, ByVal pTable As String) As String
'
'Generate the insert query
'
Dim li As Long
Dim lCon As Object
Dim lRecordset As Object
Dim Lfields As String
Dim lValues As String
Dim Lsql As String
Dim lTmp As String
  Lfields = ""
  lValues = ""
  Screen.MousePointer = vbHourglass
  Set lCon = CreateObject("ADODB.Connection")
  lCon.CursorLocation = 3 'Useclient
  Call lCon.Open(pConnectionStr)
  If lCon.State = 1 Then ' adStateOpen
    Set lRecordset = CreateObject("ADODB.Recordset")
    lRecordset.CursorType = 2 ' adOpenDynamic
    Set lRecordset = lCon.Execute("Select * from " & pTable)
    If lRecordset.Fields.Count > 0 Then
      For li = 0 To lRecordset.Fields.Count - 1
        Lfields = Lfields & "[" & lRecordset.Fields(li).Name & "],"
      Next
      Lfields = Trim(Lfields)
      Lfields = UCase(Lfields)
      If Len(Lfields) > 0 Then Lfields = Left(Lfields, Len(Lfields) - 1)
      'make query
      If Lfields <> "" Then
        Lsql = pConnection & ".Execute  " & quote & "select " & Lfields & " from " & pTable & " where <condition>" & quote
      End If
    End If
  End If
  GenerateSelect = Lsql
  Screen.MousePointer = vbDefault
End Function

Private Function GenerateUpdate(ByVal pConnection As String, ByVal pConnectionStr As String, ByVal pTable As String) As String
'
'Generate the insert query
'
Dim li As Long
Dim lCon As Object
Dim lRecordset As Object
Dim Lfields As String
Dim lValues As String
Dim Lsql As String
Dim lTmp As String
  Lfields = ""
  lValues = ""
  Screen.MousePointer = vbHourglass
  Set lCon = CreateObject("ADODB.Connection")
  lCon.CursorLocation = 3 'Useclient
  Call lCon.Open(pConnectionStr)
  If lCon.State = 1 Then ' adStateOpen
    Set lRecordset = CreateObject("ADODB.Recordset")
    lRecordset.CursorType = 2 ' adOpenDynamic
    Set lRecordset = lCon.Execute("Select * from " & pTable)
    If lRecordset.Fields.Count > 0 Then
      For li = 0 To lRecordset.Fields.Count - 1
        If lRecordset.Fields(li).Attributes <> 16 Then 'leave Auto Increment
          Lfields = Lfields & "[" & lRecordset.Fields(li).Name & "],"
          lValues = lValues & "'" & quote & " &  & " & quote & "',"
        End If
      Next
      Lfields = Trim(Lfields)
      Lfields = UCase(Lfields)
      If Len(Lfields) > 0 Then Lfields = Left(Lfields, Len(Lfields) - 1)
      'make query
      If Lfields <> "" Then
        Lsql = pConnection & ".Execute  " & quote & "update " & pTable & " set " & Lfields & " where <condition>" & quote
      End If
    End If
  End If
  GenerateUpdate = Lsql
  Screen.MousePointer = vbDefault
End Function

Public Function GetActiveRTB() As Integer
'
'Get the active rtb editor index
'
  GetActiveRTB = mEditorIndex
End Function

Private Function GetDriveImg(ByVal pDrive As String) As String
'
'Get the drives img
'
Dim lTmp As String
Dim lFso As New FileSystemObject
Dim lDrive As Drive
  On Error Resume Next
  If InStr(pDrive, ":") > 0 And InStr(pDrive, "[") > 0 Then
    pDrive = Trim(Mid(pDrive, 1, InStr(pDrive, "[") - 1)) & "\"
  Else
    pDrive = pDrive & "\"
  End If
  Set lDrive = lFso.GetDrive(pDrive)
  If Not lDrive Is Nothing Then
    If lDrive.DriveType = CDRom Then
      lTmp = "CDROM"
    ElseIf lDrive.DriveType = Removable Then
      lTmp = "FLOPPY"
    Else
      lTmp = "DISK"
    End If
  Else
    lTmp = "DISK"
  End If
  GetDriveImg = lTmp
  
  Err.Clear
End Function

Private Function GetFieldType(ByVal pFType As Integer, ByVal pSize As Long, ByVal pPrimary As Boolean, ByVal pAttribute As Integer) As String
'
'Get the table field types
'
    Select Case pFType
        Case 21 'adUnsignedBigInt
          GetFieldType = "long " & pSize
        Case 17 'adUnsignedTinyInt
            GetFieldType = "byte " & pSize
        Case 11 'adBoolean
            GetFieldType = "bit " & pSize
        Case 135 'adDBTimeStamp
            GetFieldType = "datetime " & pSize
        Case 5 'adDouble
            GetFieldType = "float " & pSize
        Case 3 'adInteger
            GetFieldType = "int " & pSize
        Case 2 'adSmallInt
            GetFieldType = "smallint " & pSize
        Case 129 'adChar
            GetFieldType = "char " & pSize
        Case 200 'adVarChar
            GetFieldType = "varchar " & pSize
        Case 130 'adWChar
            GetFieldType = "wchar " & pSize
        Case 201 'adLongVarChar
            GetFieldType = "varchar " & pSize
        Case 202 'adVarWChar
            GetFieldType = "varwchar " & pSize
        Case 203 'adLongVarWChar
            GetFieldType = "memo " & pSize
    End Select
    GetFieldType = GetFieldType & IIf(pPrimary, " primary ", "") & IIf(CBool(pAttribute And 32), " null", "") '32-for adFldNullable
End Function

Private Function GetFileImg(ByVal pFilename As String) As String
'
'Get the file types image
'
Dim lTmp As String
  On Error GoTo Cerr
  lTmp = Mid(pFilename, InStrRev(pFilename, ".") + 1)
  imgTest.Picture = imlFiles.ListImages(UCase(lTmp)).Picture
  GetFileImg = UCase(lTmp)
  Exit Function
Cerr:
  GetFileImg = "DEFAULT"
End Function

Private Function GetIndex(ByVal pImage As String) As Integer
'
'Get the image index in the imglist for menus
'
  On Error GoTo Cerr
  GetIndex = imlMenu.ListImages(UCase(pImage)).Index - 1
  Exit Function
Cerr:
  GetIndex = 0
End Function

Private Function GetOpenedDocument(ByVal pFilename As String) As String
'
'Get the document key if it is opened
'
Dim li As Integer
  On Error Resume Next
  If pFilename = "" Then Exit Function
  For li = 1 To RTB.Count - 1
    If RTB(li).Key <> "" Then
      If RTB(li).FileName = pFilename Then 'Ucase Changed
        GetOpenedDocument = RTB(li).Key
        Exit Function
      End If
    End If
  Next
  GetOpenedDocument = ""
  
  If Err.nnumber > 0 Then S110_WriteLog "frmEditor:GetOpenedDocument... " & Err.Description
  Err.Clear
End Function

Public Function LoadConnection(ByVal pConnectionName As String, ByVal pConnectionString As String, Optional ByVal pNew As Boolean)
'
'Execute the connection as load the results
'
Dim lCon As Object
Dim lTables As Object
Dim lRecordset As Object
Dim lTablename As String
Dim Lsql As String
Dim lType As String
Dim li As Integer
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  If pNew Then
    tvApplication.Nodes.Clear
    tvApplication.Nodes.Add(, , pConnectionName, pConnectionName, "CONNECTION").Tag = pConnectionString 'Ucase Changed
  End If
  tvApplication.Nodes.Add pConnectionName, tvwChild, "T" & pConnectionName, "Tables", "TABLES" 'Ucase Changed
  tvApplication.Nodes.Add pConnectionName, tvwChild, "V" & pConnectionName, "Views", "VIEW" 'Ucase Changed
  Set lCon = CreateObject("ADODB.Connection")
  lCon.CursorLocation = 3 'Useclient
  Call lCon.Open(pConnectionString)
  If lCon.State = 1 Then ' adStateOpen
    Set lTables = CreateObject("ADODB.Recordset") ' CreateObject("ADODB.Recordset") ' New ADODB.Recordset
    lTables.CursorType = 2 ' adOpenDynamic
    Set lTables = lCon.OpenSchema(20, Array(Empty, Empty, Empty, "TABLE,VIEW"))
    Do While lTables.EOF = False
      lTablename = lTables.Fields("TABLE_NAME").Value
      lTablename = Trim(lTablename)
      Lsql = "SELECT TOP 1 * FROM [" & lTablename & "]"
      Set lRecordset = CreateObject("ADODB.Recordset") ' CreateObject("ADODB.Recordset") ' New ADODB.Recordset
      Call lRecordset.Open(Lsql, lCon)
      If lRecordset.State = 1 Then 'adStateOpen
        If lTables.Fields("TABLE_TYPE") = "TABLE" Then
          tvApplication.Nodes.Add("T" & pConnectionName, tvwChild, pConnectionName & "_" & lTablename, lTablename, "TABLE").Tag = lTablename 'Ucase Changed
        Else
          tvApplication.Nodes.Add("V" & pConnectionName, tvwChild, pConnectionName & "_" & lTablename, lTablename, "TABLEV").Tag = lTablename 'Ucase Changed
        End If
        For li = 0 To lRecordset.Fields.Count - 1
          lType = GetFieldType(lRecordset.Fields(li).Type, lRecordset.Fields(li).DefinedSize, lRecordset.Fields(li).Properties("KEYCOLUMN").Value, lRecordset.Fields(li).Attributes)
          If lType <> "" Then lType = " (" & lType & ")"
          tvApplication.Nodes.Add pConnectionName & "_" & lTablename, tvwChild, , lRecordset.Fields(li).Name & lType, "FIELD" 'Ucase Changed
        Next
      End If
      lTables.MoveNext
    Loop
  End If
  Screen.MousePointer = vbDefault
  
  If Err.nnumber > 0 Then S110_WriteLog "frmEditor:LoadConnection... " & Err.Description
  Err.Clear
End Function

Public Function LoadDocument(Optional ByVal pFilename As String, Optional ByVal pSortpath As String, Optional ByVal pVirtualFolder As String, Optional ByVal pHistory As Boolean, Optional ByVal pLocalHost As String, Optional ByVal pRemote As Boolean)
'
'Load the file; if file is blank then open blank document
'
Dim lKey As String
Dim lTab As vbalDTab6.cTab

  On Error Resume Next
  mLoading = True
  lKey = GetOpenedDocument(pFilename)
  If lKey = "" Then 'new
    If pFilename <> "" Then
      Screen.MousePointer = vbHourglass
      If S102_File_Exists(pFilename) Then
        If LCase(Right(pFilename, 3)) = "css" Then
          frmCSSEditor.Mfilename = pFilename
          frmCSSEditor.Show vbModal
        Else
          stBar.Panels("P1").Text = "Loading " & Mid(pFilename, InStrRev(pFilename, "\") + 1) & "..."
          mEditorIndex = RTB.UBound + 1
          Load RTB(mEditorIndex)
          RTB(mEditorIndex).Title = Mtitle
          RTB(mEditorIndex).IsRemote = pRemote
          RTB(mEditorIndex).VirtualPath = pVirtualFolder
          RTB(mEditorIndex).Localhost = pLocalHost
          RTB(mEditorIndex).Path = Mid(pFilename, 1, InStrRev(pFilename, "\") - 1)
          RTB(mEditorIndex).FileName = pFilename
          RTB(mEditorIndex).AppPath = App.Path
          RTB(mEditorIndex).Key = pFilename
          RTB(mEditorIndex).AutoComplete = mAutoCompletion
          RTB(mEditorIndex).Intelisense = mIntelisense
          RTB(mEditorIndex).SyntaxHighlighting = mSyntaxHighlighting
          RTB(mEditorIndex).IsHistory = pHistory
          RTB(mEditorIndex).OpenFile pFilename
          RTB(mEditorIndex).WordWrap = Not mStandardMenu.Checked(mStandardMenu.IndexForKey("mnuWordwrap"))
          RTB(mEditorIndex).Lineno = mStandardMenu.Checked(mStandardMenu.IndexForKey("mnuLinenumber"))
          Set lTab = tabMain.Tabs.Add(pFilename, , IIf(pSortpath = "", Mid(pFilename, InStrRev(pFilename, "\") + 1), pSortpath))
          lTab.Panel = RTB(mEditorIndex)
          RTB(mEditorIndex).SetFocus
          lTab.Selected = True
          lTab.Tag = mEditorIndex
          If tabMain.Visible = False Then tabMain.Visible = True
          AddFileToMenu lTab.Key
          stBar.Panels("P1").Text = "Done"
        End If
      Else
        MsgBox "Specified file not found!", vbCritical, Mtitle
        Screen.MousePointer = vbDefault
        Exit Function
      End If
    Else 'Blank
      Mrecentfile = ""
      If Mopendialog = False Then
        frmTemplates.LoadRecent
        frmTemplates.Show vbModal
        mStandardMenu.Checked(mStandardMenu.IndexForKey("mnuOpenDialog")) = Mopendialog
      Else
        frmTemplates.cmdOk_Click
      End If
      If Mrecentfile <> "" Then 'if recent file is opened
        LoadDocument Mrecentfile
        Exit Function
      End If
      If Mdocumentype > -1 Then 'if new document is seleted
        Screen.MousePointer = vbHourglass
        mBlankPage = mBlankPage + 1
        mEditorIndex = RTB.UBound + 1
        Load RTB(mEditorIndex)
        RTB(mEditorIndex).Title = Mtitle
        RTB(mEditorIndex).Path = App.Path
        RTB(mEditorIndex).AppPath = App.Path
        RTB(mEditorIndex).VirtualPath = pVirtualFolder
        RTB(mEditorIndex).Localhost = pLocalHost
        RTB(mEditorIndex).Key = "New Document " & mBlankPage
        RTB(mEditorIndex).Blankpage = mBlankPage
        RTB(mEditorIndex).AutoComplete = mAutoCompletion
        RTB(mEditorIndex).Intelisense = mIntelisense
        RTB(mEditorIndex).SyntaxHighlighting = mSyntaxHighlighting
        RTB(mEditorIndex).IsHistory = False
        RTB(mEditorIndex).WordWrap = Not mStandardMenu.Checked(mStandardMenu.IndexForKey("mnuWordwrap"))
        RTB(mEditorIndex).Lineno = mStandardMenu.Checked(mStandardMenu.IndexForKey("mnuLinenumber"))
        RTB(mEditorIndex).LoadNew Mdocumentype
        Set lTab = tabMain.Tabs.Add("New Document " & mBlankPage, , "New Document " & mBlankPage)
        lTab.Panel = RTB(mEditorIndex)
        RTB(mEditorIndex).Visible = True
        RTB(mEditorIndex).SetFocus
        lTab.Selected = True
        lTab.Tag = mEditorIndex
        If tabMain.Visible = False Then tabMain.Visible = True
        AddFileToMenu lTab.Key
        tmrFocus.Enabled = True
      End If
    End If
  Else 'set focus(select the already opened)
    tabMain.Tabs.Item(lKey).Selected = True
    mEditorIndex = val(tabMain.Tabs.Item(lKey).Tag)
    RTB(mEditorIndex).VirtualPath = pVirtualFolder
    RTB(mEditorIndex).SetFocus
  End If
  mLoading = False
  Screen.MousePointer = vbDefault
  
  If Err.nnumber > 0 Then S110_WriteLog "frmEditor:LoadDocument... " & Err.Description
  Err.Clear
End Function

Private Sub LoadDrives()
'
'Load the desktop/mycomputer/mydocuments/mynetworkplaces
'
Dim li As Integer
Dim lDrive As String
  drvBox.Refresh
  LockWindowUpdate Me.hwnd
  Screen.MousePointer = vbHourglass
  
  tvFiles.Nodes.Clear
  tvFiles.Nodes.Add , , "D0", "Desktop", "DES"
  tvFiles.Nodes.Add "D0", tvwChild, "D1", "My Documents", "MYD"
  LoadFiles tvFiles, GetFolder(ftMYDOCUMENTS, Me.hwnd), "D1"
  tvFiles.Nodes.Add "D0", tvwChild, "D2", "My Computer", "MYCOMPUTER"
  tvFiles.Nodes.Add "D0", tvwChild, "D3", "My Network Places", "MYN"
  LoadFiles tvFiles, GetFolder(ftNETHOOD, Me.hwnd), "D3"
  For li = 0 To drvBox.ListCount - 1
    lDrive = Trim(Split(drvBox.List(li), "[")(0)) 'Ucase Changed
    tvFiles.Nodes.Add "D2", tvwChild, lDrive, drvBox.List(li), GetDriveImg(drvBox.List(li))
    tvFiles.Nodes.Add lDrive, tvwChild, lDrive & " ", " ", "FOLDERCLOSE"
  Next
  LoadFiles tvFiles, GetFolder(ftDESKTOP, Me.hwnd), "D0"
  tvFiles.Nodes("D2").Expanded = True
  tvFiles.Nodes("D2").Selected = True
  Screen.MousePointer = vbDefault
  LockWindowUpdate 0&
End Sub

Private Sub LoadFiles(ByVal pTv As TreeView, ByVal pPath As String, ByVal pParent As String, Optional ByVal pSite As Boolean)
'
'Load the files/subfolders of given folder
'
Dim li As Integer
Dim lj As Integer
Dim lParent As String
Dim lPath As String
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  If PathExists(pPath, pSite) = False Then
    If InStr(pPath, ":") > 0 And InStr(pPath, "[") > 0 Then
      drBox.Path = Trim(Mid(pPath, 1, InStr(pPath, "[") - 1)) & "\"
    Else
      drBox.Path = pPath & "\"
    End If
    If Err.Number = 0 Then
      For li = 0 To drBox.ListCount - 1
        pTv.Nodes.Add pParent, tvwChild, drBox.List(li), Mid(drBox.List(li), InStrRev(drBox.List(li), "\") + 1), "FOLDERCLOSE" 'Ucase Changed
        Err.Clear
        'Get sub folders
        subDr.Path = drBox.List(li)
        If Err.Number = 0 Then
          For lj = 0 To subDr.ListCount - 1
            pTv.Nodes.Add(drBox.List(li), tvwChild, subDr.List(lj), Mid(subDr.List(lj), InStrRev(subDr.List(lj), "\") + 1), "FOLDERCLOSE").Tag = subDr.List(lj) 'Ucase Changed
          Next
        End If
        Err.Clear
        'Get files of subfolders
        flBox.Path = drBox.List(li)
        If Err.Number = 0 Then
          For lj = 0 To flBox.ListCount - 1
            lPath = drBox.List(li) & "\" & flBox.List(lj) 'Ucase Changed
            lPath = Replace(lPath, "\\", "\")
            If Left(lPath, 1) = "\" Then lPath = "\" & lPath 'For network path
            pTv.Nodes.Add(drBox.List(li), tvwChild, lPath, flBox.List(lj), GetFileImg(flBox.List(lj))).Tag = lPath 'Ucase Changed
          Next
        End If
        Err.Clear
      Next
      'get files of folder
      flBox.Path = drBox.Path
      If Err.Number = 0 Then
        For li = 0 To flBox.ListCount - 1
          lPath = drBox.Path & "\" & flBox.List(li) 'Ucase Changed
          lPath = Replace(lPath, "\\", "\")
          If Left(lPath, 1) = "\" Then lPath = "\" & lPath 'For network path
          pTv.Nodes.Add(pParent, tvwChild, lPath, flBox.List(li), GetFileImg(flBox.List(li))).Tag = lPath
        Next
      End If
    End If
    pTv.Nodes(pParent).EnsureVisible
    If pSite Then
      lsvSitePath.ListItems.Add , pPath, pPath
    Else
      lvPaths.ListItems.Add , pPath, pPath
    End If
    Err.Clear
  End If
  Screen.MousePointer = vbDefault
  
  If Err.nnumber > 0 Then S110_WriteLog "frmEditor:Loadfiles... " & Err.Description
  Err.Clear
End Sub

Public Sub LoadFontsMenu()
'
'Load the fonts family from the family list form
'
Dim iP2 As Integer
Dim li As Integer
  On Error Resume Next
  With mStandardMenu
    .RemoveItem "mnuDefaultFont"
    .RemoveItem "sepFont"
    .RemoveItem "mnuOtherFonts"
    For li = 0 To val(frmFonts.lsFontslist_b.Tag) - 1
      .RemoveItem "font" & li
    Next
    iP2 = .IndexForKey("mnuFont")
    .AddItem "Default", , , iP2, , , , "mnuDefaultFont"
    For li = 0 To frmFonts.lsFontsList.ListCount - 1
      If LCase(frmFonts.lsFontsList.List(li)) <> "(new fonts list)" Then
        .AddItem frmFonts.lsFontsList.List(li), , , iP2, , , , "font" & li
      End If
    Next
    .AddItem "-", , , iP2, , , , "sepFont"
    .AddItem "Others...", , , iP2, , , , "mnuOtherFonts"
  End With
End Sub

Private Function LoadHistory()
'
'Load the history of recent files/past opened documents
'
Dim lHistory As clsHistory
Dim LIDS As Variant
Dim li As Long
  LockWindowUpdate Me.hwnd
  Screen.MousePointer = vbHourglass
  'Get the history
  If Mhistories Is Nothing Then
    Set Mhistories = New clsHistories
    Mhistories.Load
  End If
  On Error Resume Next
  tvHistory.Nodes.Clear
  LIDS = Mhistories.IDs
  'Set the root
  If UBound(LIDS) < 0 Then
    tvHistory.HideSelection = True
    tvHistory.Nodes.Add(, , "H0", "    (No files)", "EMPTY").ForeColor = RGB(123, 123, 123)
  Else
    tvHistory.HideSelection = False
    tvHistory.Nodes.Add , , "H0", "Recent files", "HISTORY"
  End If
  'Load the history to treeview
  For li = UBound(LIDS) To LBound(LIDS) Step -1
    If LIDS(li) <> "" Then
      Set lHistory = Mhistories.Item(LIDS(li))
      If Not lHistory Is Nothing Then
        With lHistory
          tvHistory.Nodes.Add "H0", tvwChild, .HDate, IIf(DateDiff("d", .HDate, Date) = 0, "Today", .HDate), IIf(DateDiff("d", .HDate, Date) = 0, "HTODAY", "HPAST")
          tvHistory.Nodes.Add(.HDate, tvwChild, LIDS(li), Mid(.LocalPath, InStrRev(.LocalPath, "\") + 1), GetFileImg(.LocalPath)).Tag = 1     'Ucase Changed
        End With
      End If
    End If
  Next
  tvHistory.Nodes("H0").Expanded = True
  tvHistory.Nodes("H0").EnsureVisible
  Screen.MousePointer = vbDefault
  LockWindowUpdate 0&
  
  If Err.nnumber > 0 Then S110_WriteLog "frmEditor:Loadhistory... " & Err.Description
  Err.Clear
End Function

Private Function LoadNetworkfiles(ByVal pNode As Node)
'
'Load the network files of mapped drives/folders
'
Dim fn As Integer
Dim lPath As String
Dim lFso As New FileSystemObject
Dim li As Integer
  If Not pNode.Parent Is Nothing Then
    If pNode.Parent.Key = "D3" Then
      lPath = lFso.OpenTextFile(pNode.Key & "\target.lnk").ReadAll
      lPath = Mid(lPath, InStr(lPath, "\\"), InStr(InStr(lPath, "\\"), lPath, CStr(Chr(0))) - InStr(lPath, "\\"))
      LoadFiles tvFiles, lPath, pNode.Key
      tvFiles.Nodes.Remove pNode.Key & "\target.lnk" 'Ucase Changed
    End If
  End If
End Function

Private Function LoadRemoteFiles(ByVal pDir As String, ByVal pParent As String)
'
'Load the remote files for online
'
Dim lFiles As New Collection
Dim lConnected As Boolean
Dim hFile As Long
Dim lData As WIN32_FIND_DATA
Dim lFile As String
Dim lTemp As String
Dim lDetails As Variant
Dim li As Integer
  If PathExists(pDir, True) = False Then
    lConnected = mFTP.IsConnected
    If lConnected = False Then
      lConnected = mFTP.OpenConnection(mServer, mUsername, mPassword)
    End If
    If lConnected Then
      pDir = Replace(pDir, "//", "/")
      If Right(pDir, 1) = "/" Then pDir = Left(pDir, Len(pDir) - 1)
      If Left(pDir, 1) = "/" Then pDir = Mid(pDir, 2)
      If mFTP.SetFTPDirectory("/\/" & pDir) Then
        hFile = FtpFindFirstFile(mFTP.GetConnection, "*.*", lData, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)
        If hFile Then
          Do
            lFile = Left(lData.cFileName, InStr(1, lData.cFileName, Chr(0)) - 1)
            If lFile <> "" Then
              If lData.dwFileAttributes And vbDirectory Then
                'Add the folders
                tvSiteFiles.Nodes.Add(pParent, tvwChild, pDir & "/" & lFile, lFile, "FOLDERCLOSE").Tag = pDir & "/" & lFile
                Call tvSiteFiles.Nodes.Add(pDir & "/" & lFile, tvwChild, pDir & "/" & lFile & " ") 'For expand symbol
              Else
                'Collect the files for uniform
                lTemp = pParent & "^" & pDir & "/" & lFile & "^" & lFile & "^" & GetFileImg(lFile)
                lFiles.Add lTemp
              End If
            End If
          Loop While InternetFindNextFile(hFile, lData)
          'Now add the files
          For li = 1 To lFiles.Count
            If lFiles(li) <> "" Then
              lDetails = Split(lFiles(li), "^")
              tvSiteFiles.Nodes.Add(lDetails(0), tvwChild, lDetails(1), lDetails(2), lDetails(3)).Tag = lDetails(1)
            End If
          Next
        End If
      End If
      'Add the path to collection to avoid reloading
      lsvSitePath.ListItems.Add , pDir, pDir
    End If
    'Close the handle
    InternetCloseHandle hFile
  End If
End Function

Private Function LoadSettings()
'
'Load the codepiler settings when start up
'
  mSite = val(GetSetting(App.Title, "LastVisitedSite", "LastVisitedSite"))
  mIntelisense = CBool(val(GetSetting(App.Title, "Intelisense", "Intelisense", "1")))
  mWordWrap = CBool(val(GetSetting(App.Title, "Wordwrap", "Wordwrap")))
  mLineNo = CBool(val(GetSetting(App.Title, "Lineno", "Lineno", "1")))
  mFullmodePreview = CBool(val(GetSetting(App.Title, "Fullmode", "Fullmode", "1")))
  mToolbox = CBool(val(GetSetting(App.Title, "Toolbox", "Toolbox", "1")))
  mToolboxSize = val(GetSetting(App.Title, "ToolboxSize", "ToolboxSize", picTools.Width))
  Mopendialog = CBool(val(GetSetting(App.Title, "OpenDialog", "OpenDialog", "0")))
  mSyntaxHighlighting = CBool(val(GetSetting(App.Title, "SyntaxHighlighting", "SyntaxHighlighting", "1")))
  mAutoCompletion = CBool(val(GetSetting(App.Title, "AutoCompletion", "AutoCompletion", "1")))
  'apply to statusbar setting
  stBar.Panels("PRow").Visible = Not mWordWrap
  stBar.Panels("PCol").Visible = Not mWordWrap
  'apply to site
  cboSites.Tag = mSite
  'apply to toolbox
  picTools.Visible = False
  picTools.Tag = picTools.Width
  If Not mToolbox Then
    picTools.Width = 0
  Else
    picTools.Width = mToolboxSize
  End If
  Exit Function
End Function

Private Function LoadSiteFiles()
'
'Load the selected site
'
Dim lPath As String
Dim lSite As clsSite
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  If cboSites.ListCount > 0 Then
    Set lSite = Msitedetails.Item(cboSites.Text)
    If Not lSite Is Nothing Then
      tvSiteFiles.Nodes.Clear
      lPath = lSite.LocalPath
      If lPath <> "" Then
        tvSiteFiles.Nodes.Add , , "F0", lPath, "FOLDEROPEN"
        lsvSitePath.ListItems.Clear
        LoadFiles tvSiteFiles, lPath, "F0", True
        tvSiteFiles.Nodes("F0").Expanded = True
        mSitename = cboSites.List(cboSites.ListIndex)
      End If
    End If
  End If
  Screen.MousePointer = vbDefault
  
  If Err.nnumber > 0 Then S110_WriteLog "frmEditor:Loadsitefiles... " & Err.Description
  Err.Clear
End Function

Public Function LoadSites(Optional ByVal pSite As String)
'
'Load the sites of user created to the combo box
'
Dim fn As Integer
Dim lContent As String
Dim lFiles As Variant
Dim lIndex As Integer
Dim lSite As clsSite
Dim lFile As String
Dim lKey As Variant
  Screen.MousePointer = vbHourglass
  cboSites.Clear
  lIndex = -1
  If Msitedetails Is Nothing Then
    Set Msitedetails = New clsSites
    Msitedetails.Load
  End If
  If Msitedetails.Count > 0 Then
    For Each lKey In Msitedetails.IDs
      Set lSite = Msitedetails.Item(lKey)
      If Not lSite Is Nothing Then
        cboSites.AddItem lSite.Name
        If LCase(lSite.Name) = LCase(pSite) Then
          lIndex = cboSites.NewIndex
        End If
      End If
    Next
    cboSites.AddItem "---------------"
    cboSites.AddItem "Define sites..."
    tvSiteFiles.Tag = ""
    tvSiteFiles.Nodes.Clear
    cboSites.ListIndex = IIf(lIndex > -1, lIndex, IIf(cboSites.Tag <> "", val(cboSites.Tag), 0))
  Else
    cboSites.AddItem "Define sites..."
  End If
  Screen.MousePointer = vbDefault
End Function

Private Sub LoadStyleClasses()
'
'Load the stylesheet classes form the stylesheet linking
'
Dim li As Long
Dim lID As Integer
Dim lCount As Long
Dim lClasses As Variant
  On Error Resume Next
  frmStylesheet.Show vbModal
  RTB(mEditorIndex).SplitClasses
  For li = 0 To RTB(mEditorIndex).Styles
    mStandardMenu.RemoveItem "classes" & li
  Next
  If InStr(RTB(mEditorIndex).StylesList, vbCrLf) > 0 Then
    lClasses = Split(RTB(mEditorIndex).StylesList, vbCrLf)
    lCount = UBound(lClasses) + 1
    lID = mStandardMenu.IndexForKey("mnuCSSStyles")
    For li = 0 To lCount - 2
      mStandardMenu.AddItem Split(lClasses(li), ",")(2), , , lID, , , , "classes" & li
    Next
  End If
  
  If Err.nnumber > 0 Then S110_WriteLog "frmEditor:Loadstyleclasses... " & Err.Description
  Err.Clear
End Sub

Private Sub OpenDocument()
'
'Open the document
'
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

Private Function PathExists(ByVal pPath As String, Optional ByVal pSite As Boolean) As Boolean
'
'Check for path exist of ws/site tree to avoid reopen the folders/drive (if already opened)
'
Dim Litem As ListItem
  On Error Resume Next
  If pSite Then
    Set Litem = lsvSitePath.ListItems(pPath)
  Else
    Set Litem = lvPaths.ListItems(pPath)
  End If
  PathExists = Not Litem Is Nothing
  Err.Clear
End Function

Private Sub PreviewInBrowser()
'
'Preview the page in browser
'
Dim Ret As Long
Dim FileNum As Integer
Dim lPath As String
Dim lFile As String
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  If RTB(mEditorIndex).VirtualPath <> "" Then
    lFile = tabMain.SelectedTab.Key
    If RTB(mEditorIndex).FileName <> "" Then
      RTB(mEditorIndex).SaveFile lFile
    Else
      S101_Make_Dir RTB(mEditorIndex).VirtualPath & "\tmp\"
      lFile = RTB(mEditorIndex).VirtualPath & "\tmp\" & "rnd" & Rnd(1000) & ".asp"
      RTB(mEditorIndex).SaveFile lFile, True
    End If
    lPath = GetVirtualPath(RTB(mEditorIndex).VirtualPath, lFile)
    lPath = RTB(mEditorIndex).Localhost & lPath
    lPath = Replace(lPath, "\", "/")
    lPath = Replace(lPath, "//", "/")
    lPath = Replace(lPath, "http:", "http:/")
    Ret = ShellExecute(Me.hwnd, "Open", lPath, vbNullString, vbNullString, SW_SHOWMAXIMIZED)
    S105_Delete lFile
  Else
    lFile = "rnd" & Rnd(1000) & ".asp"
    RTB(mEditorIndex).SaveFile App.Path & "\" & lFile, True
    Ret = ShellExecute(Me.hwnd, "Open", lFile, vbNullString, App.Path, SW_SHOWMAXIMIZED)
    S105_Delete App.Path & "\" & lFile
  End If
  Screen.MousePointer = vbDefault
  
  If Err.nnumber > 0 Then S110_WriteLog "frmEditor:PreviewinBrowser... " & Err.Description
  Err.Clear
End Sub

Private Function RaiseProgress(ByVal pMax As Long, ByVal pValue As Long, ByVal pStatus As String, Optional ByVal pShowPercent As Boolean = True)
'
'Show the status text and progress if any
'
  On Error Resume Next
  If pStatus <> "" Then stBar.Panels("P1").Text = pStatus
  If pMax > 0 Then prgbarMain.Max = pMax
  If pValue >= pMax Then
    prgbarMain.Max = 1
    prgbarMain.Value = 0
    prgbarMain.Visible = False
    stBar.Panels("P3").Text = ""
    Screen.MousePointer = vbDefault
    Exit Function
  Else
    Screen.MousePointer = vbHourglass
    If prgbarMain.Visible = False Then prgbarMain.Visible = True
    prgbarMain.Value = pValue
    prgbarMain.Refresh
    If pShowPercent Then stBar.Panels("P3").Text = CInt((pValue / pMax) * 100) & "%"
  End If
  DoEvents
End Function

Private Function ResizeTool()
'
'Resize the toolbox when click the property bar
'
  On Error Resume Next
  'Workspace
  picWorkSpace.Left = picToolbox.Left
  picWorkSpace.Width = picToolbox.Width
  picWorkSpace.Top = picToolbox.Top + picToolbox.Height + Screen.TwipsPerPixelX * 2
  
  picWSProperty.Left = 0
  picWSProperty.Width = picWorkSpace.Width
  picWSProperty.Top = 0
  
  spBorderWS.Left = 0
  spBorderWS.Top = 0
  spBorderWS.Width = picWSProperty.Width
  spBorderWS.Height = picWSProperty.Height
  
  spBorder.Left = picWSProperty.Left
  spBorder.Top = picWSProperty.Height
  spBorder.Width = picWorkSpace.Width
  spBorder.Height = picWorkSpace.Height - spBorder.Top
  
  tvFiles.Left = Screen.TwipsPerPixelX * spBorder.BorderWidth
  tvFiles.Width = spBorder.Width - (Screen.TwipsPerPixelX * spBorder.BorderWidth * 2)
  tvFiles.Height = spBorder.Height - (Screen.TwipsPerPixelY * spBorder.BorderWidth * 2)
  tvFiles.Top = spBorder.Top + Screen.TwipsPerPixelY * spBorder.BorderWidth
  
  'History
  picHistory.Left = picToolbox.Left
  picHistory.Width = picToolbox.Width
  picHistory.Top = picWorkSpace.Top + picWorkSpace.Height + Screen.TwipsPerPixelX * 2
  
  picHProperty.Left = 0
  picHProperty.Width = picHistory.Width
  picHProperty.Top = 0
  
  spBorderH.Left = 0
  spBorderH.Top = 0
  spBorderH.Width = picHProperty.Width
  spBorderH.Height = picHProperty.Height
  
  spBorderHT.Left = picHProperty.Left
  spBorderHT.Top = picHProperty.Height
  spBorderHT.Width = picHistory.Width
  spBorderHT.Height = picHistory.Height - spBorderHT.Top
  
  tvHistory.Left = Screen.TwipsPerPixelX * spBorderHT.BorderWidth
  tvHistory.Width = spBorderHT.Width - (Screen.TwipsPerPixelX * spBorderHT.BorderWidth * 2)
  tvHistory.Height = spBorderHT.Height - (Screen.TwipsPerPixelY * spBorderHT.BorderWidth * 2)
  tvHistory.Top = spBorderHT.Top + Screen.TwipsPerPixelY * spBorderHT.BorderWidth
  
  'Site files
  picSite.Left = picToolbox.Left
  picSite.Width = picToolbox.Width
  picSite.Top = picHistory.Top + picHistory.Height + Screen.TwipsPerPixelX * 2
  
  picSProperty.Left = 0
  picSProperty.Width = picSite.Width
  picSProperty.Top = 0
  
  spBorderS.Left = 0
  spBorderS.Top = 0
  spBorderS.Width = picSProperty.Width
  spBorderS.Height = picSProperty.Height
  
  cboSites.Left = 0
  cboSites.Top = picSProperty.Height
  cboSites.Width = picSProperty.Width + Screen.TwipsPerPixelX * 2
  
  spBorderST.Left = picSProperty.Left
  spBorderST.Top = cboSites.Top + cboSites.Height
  spBorderST.Width = picSite.Width
  spBorderST.Height = picSite.Height - spBorderST.Top '- Screen.TwipsPerPixelY * 5
  
  tvSiteFiles.Left = Screen.TwipsPerPixelX * spBorderST.BorderWidth
  tvSiteFiles.Width = spBorderST.Width - (Screen.TwipsPerPixelX * spBorderST.BorderWidth * 2)
  tvSiteFiles.Height = spBorderST.Height - (Screen.TwipsPerPixelY * spBorderST.BorderWidth * 2)
  tvSiteFiles.Top = spBorderST.Top + Screen.TwipsPerPixelY * spBorderST.BorderWidth
  
  'Application
  picApplication.Left = picToolbox.Left
  picApplication.Width = picToolbox.Width
  picApplication.Top = picSite.Top + picSite.Height + Screen.TwipsPerPixelX * 2
  
  picAProperty.Left = 0
  picAProperty.Width = picApplication.Width
  picAProperty.Top = 0
  
  spBorderA.Left = 0
  spBorderA.Top = 0
  spBorderA.Width = picAProperty.Width
  spBorderA.Height = picAProperty.Height
  
  spBorderAT.Left = picAProperty.Left
  spBorderAT.Top = picAProperty.Height
  spBorderAT.Width = picApplication.Width
  spBorderAT.Height = picApplication.Height - spBorderAT.Top
  
  tvApplication.Left = Screen.TwipsPerPixelX * spBorderAT.BorderWidth
  tvApplication.Width = spBorderAT.Width - (Screen.TwipsPerPixelX * spBorderAT.BorderWidth * 2)
  tvApplication.Height = spBorderAT.Height - (Screen.TwipsPerPixelY * spBorderAT.BorderWidth * 2)
  tvApplication.Top = spBorderAT.Top + Screen.TwipsPerPixelY * spBorderAT.BorderWidth
End Function

Public Sub SaveDocument(Optional ByVal pSaveAs As Boolean)
'
'Save/SaveAs the document
'
Dim LFileName As String
  Screen.MousePointer = vbHourglass
  If RTB(mEditorIndex).FileName = "" Or pSaveAs Then
    LFileName = RTB(mEditorIndex).SaveAsFile
    If S102_File_Exists(LFileName) Then
      RTB(mEditorIndex).FileName = LFileName
      tabMain.Tabs.Item(RTB(mEditorIndex).Key).Caption = Mid(LFileName, InStrRev(LFileName, "\") + 1)
    End If
  Else
    RTB(mEditorIndex).SaveFile RTB(mEditorIndex).FileName
    RTB(mEditorIndex).Changed = False
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Function SaveHistory(ByVal pFilename As String)
'
'Save the file to history
'
Dim lNode As Object
Dim lHistory As clsHistory
  Err.Clear
  On Error Resume Next
  'Save to history
  Set lHistory = New clsHistory
  lHistory.HDate = Date
  lHistory.LocalPath = pFilename
  Mhistories.Add lHistory
  Mhistories.Save
  'Add to history treeview
  Set lNode = tvHistory.Nodes(CStr(Date))
  If Err.Number > 0 Then
    Err.Clear
    If tvHistory.Nodes.Count = 1 Then
      tvHistory.Nodes.Clear
      tvHistory.Nodes.Add , , "H0", "Recent files", "HISTORY"
    End If
    tvHistory.Nodes.Add "H0", tvwChild, CStr(Date), "Today", "HTODAY"
  End If
  tvHistory.Nodes.Add(CStr(Date), tvwChild, Date & pFilename, Mid(pFilename, InStrRev(pFilename, "\") + 1), GetFileImg(pFilename)).Tag = 1 'Ucase Changed
  
  If Err.nnumber > 0 Then S110_WriteLog "frmEditor:Savehistory... " & Err.Description
  Err.Clear
End Function

Private Function SaveSettings()
'
'Save the codepiler settings when exit of codepiler
'
  SaveSetting App.Title, "Intelisense", "Intelisense", CInt(mIntelisense)
  SaveSetting App.Title, "AutoCompletion", "AutoCompletion", CInt(mAutoCompletion)
  SaveSetting App.Title, "SyntaxHighlighting", "SyntaxHighlighting", CInt(mSyntaxHighlighting)
  SaveSetting App.Title, "OpenDialog", "OpenDialog", CInt(Mopendialog)
  SaveSetting App.Title, "Wordwrap", "Wordwrap", CInt(mWordWrap)
  SaveSetting App.Title, "Lineno", "Lineno", CInt(mLineNo)
  SaveSetting App.Title, "Fullmode", "Fullmode", CInt(mFullmodePreview)
  SaveSetting App.Title, "Toolbox", "Toolbox", CInt(mToolbox)
  SaveSetting App.Title, "ToolboxSize", "ToolboxSize", mToolboxSize
  SaveSetting App.Title, "LastVisitedSite", "LastVisitedSite", cboSites.ListIndex
End Function

Private Sub SetAutocompletion(ByVal pVal As Boolean)
'
'Set all opened documents autocompletion setting
'
Dim li As Integer
  Screen.MousePointer = vbHourglass
  If RTB.Count > 0 Then
    For li = 1 To RTB.Count - 1
      If Not RTB(li) Is Nothing Then
        RTB(li).AutoComplete = pVal
      End If
    Next
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub SetFullmode(ByVal pVal As Boolean)
'
'Set all opened documents fullmode preview setting
'
Dim li As Integer
  Screen.MousePointer = vbHourglass
  If RTB.Count > 0 Then
    For li = 1 To RTB.Count - 1
      If Not RTB(li) Is Nothing Then
        RTB(li).Fullmode = pVal
      End If
    Next
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub SetIntellisense(ByVal pVal As Boolean)
'
'Set all opened documents intellisense setting
'
Dim li As Integer
  Screen.MousePointer = vbHourglass
  If RTB.Count > 0 Then
    For li = 1 To RTB.Count - 1
      If Not RTB(li) Is Nothing Then
        RTB(li).Intelisense = pVal
      End If
    Next
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub SetLinenumber(ByVal pVal As Boolean)
'
'Set all opened documents linenumber setting
'
Dim li As Integer
  Screen.MousePointer = vbHourglass
  If RTB.Count > 0 Then
    For li = 1 To RTB.Count - 1
      If Not RTB(li) Is Nothing Then
        RTB(li).Lineno = pVal
      End If
    Next
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub SetSyntaxHighlighting(ByVal pVal As Boolean)
'
'Set all opened documents syntax highlighting setting
'
Dim li As Integer
  Screen.MousePointer = vbHourglass
  If RTB.Count > 0 Then
    For li = 1 To RTB.Count - 1
      If Not RTB(li) Is Nothing Then
        RTB(li).SyntaxHighlighting = pVal
      End If
    Next
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub SetWordwrap(ByVal pVal As Boolean)
'
'Set all opened documents wordwrap setting
'
Dim li As Integer
  Screen.MousePointer = vbHourglass
  If RTB.Count > 0 Then
    For li = 1 To RTB.Count - 1
      If Not RTB(li) Is Nothing Then
        RTB(li).WordWrap = pVal
      End If
    Next
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Function TestConnection(ByVal pConnectionString As String) As Boolean
'
'Test the connection for valid
'
Dim lCon As Object
  Err.Clear
  On Error GoTo Cerr
  Set lCon = CreateObject("ADODB.Connection")
  TestConnection = True
  Call lCon.Open(pConnectionString)
  If lCon.State = 1 Then 'Open state
    TestConnection = True
  Else
    TestConnection = False
  End If
  MsgBox "Connection succeeded.", vbInformation, Mtitle
  Exit Function
Cerr:
  MsgBox Err.Description, vbCritical, Mtitle
  TestConnection = False
End Function

Private Function UncheckAllMenu()
'
'Uncheck all recent file menus in the window menu
'
Dim li As Integer
  On Error Resume Next
  For li = 1 To tabMain.Tabs.Count
    mStandardMenu.Checked(mStandardMenu.IndexForKey("file" & tabMain.Tabs.Item(li).Key)) = False
  Next
End Function

Private Sub UpdateEditMenu()
'
'Update the edit menus when changes in the editor
'
Dim lEnable As Boolean
  If mLoading Then Exit Sub
  With mStandardMenu
    If RTB.Count = 1 Then
      lEnable = False
      If .Enabled(.IndexForKey("mnuRedo")) <> lEnable Then .Enabled(.IndexForKey("mnuRedo")) = lEnable
      If .Enabled(.IndexForKey("mnuUndo")) <> lEnable Then .Enabled(.IndexForKey("mnuUndo")) = lEnable
      If .Enabled(.IndexForKey("mnuCopy")) <> lEnable Then .Enabled(.IndexForKey("mnuCopy")) = lEnable
      If .Enabled(.IndexForKey("mnuCut")) <> lEnable Then .Enabled(.IndexForKey("mnuCut")) = lEnable
      If .Enabled(.IndexForKey("mnuPaste")) <> lEnable Then .Enabled(.IndexForKey("mnuPaste")) = lEnable
    Else
      If .Enabled(.IndexForKey("mnuRedo")) <> RTB(mEditorIndex).CanRedo Then .Enabled(.IndexForKey("mnuRedo")) = RTB(mEditorIndex).CanRedo
      If .Enabled(.IndexForKey("mnuUndo")) <> RTB(mEditorIndex).CanUndo Then .Enabled(.IndexForKey("mnuUndo")) = RTB(mEditorIndex).CanUndo
      If .Enabled(.IndexForKey("mnuCopy")) <> RTB(mEditorIndex).CanCopy Then .Enabled(.IndexForKey("mnuCopy")) = RTB(mEditorIndex).CanCopy
      If .Enabled(.IndexForKey("mnuCut")) <> RTB(mEditorIndex).CanCut Then .Enabled(.IndexForKey("mnuCut")) = RTB(mEditorIndex).CanCut
      If .Enabled(.IndexForKey("mnuPaste")) <> RTB(mEditorIndex).CanPaste Then .Enabled(.IndexForKey("mnuPaste")) = RTB(mEditorIndex).CanPaste
    End If
    If ctbHeader.ButtonEnabled("Undo") <> .Enabled(.IndexForKey("mnuUndo")) Then ctbHeader.ButtonEnabled("Undo") = .Enabled(.IndexForKey("mnuUndo"))
    If ctbHeader.ButtonEnabled("Redo") <> .Enabled(.IndexForKey("mnuRedo")) Then ctbHeader.ButtonEnabled("Redo") = .Enabled(.IndexForKey("mnuRedo"))
    If ctbHeader.ButtonEnabled("Copy") <> .Enabled(.IndexForKey("mnuCopy")) Then ctbHeader.ButtonEnabled("Copy") = .Enabled(.IndexForKey("mnuCopy"))
    If ctbHeader.ButtonEnabled("Cut") <> .Enabled(.IndexForKey("mnuCut")) Then ctbHeader.ButtonEnabled("Cut") = .Enabled(.IndexForKey("mnuCut"))
    If ctbHeader.ButtonEnabled("Paste") <> .Enabled(.IndexForKey("mnuPaste")) Then ctbHeader.ButtonEnabled("Paste") = .Enabled(.IndexForKey("mnuPaste"))
  End With
End Sub

Public Function UpdateSiteLocalHost(oldname As String, newname As String)
'
'Set the rtb editor localhost when changes is made in the site details
'
Dim Lcnt As Integer
 For Lcnt = RTB.LBound To RTB.UBound
   If RTB(Lcnt).Localhost = oldname Then RTB(Lcnt).Localhost = newname
 Next
End Function

Private Function UploadFile(ByVal pLocalPath As String, ByVal pVirtualPath As String) As Boolean
'
'Uplaod the file to remote for view
'pLocalpath (Path+file), pVirtualPath (Base Path of file which is downloaded)
'
Dim lConnected As Boolean
Dim lFile As String
Dim lPath As String
  RaiseProgress 0, 0, "Uploading... " & Mid(pLocalPath, InStrRev(pLocalPath, "\") + 1)
  'Test for connection
  lConnected = mFTP.IsConnected
  If lConnected = False Then
    lConnected = mFTP.OpenConnection(mServer, mUsername, mPassword)
  End If
  If lConnected Then
    'Get remote path/file
    lPath = Mid(pLocalPath, 1, InStrRev(pLocalPath, "\") - 1)
    lPath = Replace(lPath, pVirtualPath, "")
    If Left(lPath, 1) = "\" Then lPath = Mid(lPath, 2)
    lPath = Replace(lPath, "\", "/")
    lFile = Mid(pLocalPath, InStrRev(pLocalPath, "\") + 1)
    'Upload the file
    If mFTP.SetFTPDirectory("/\/" & lPath) Then
      If mFTP.FTPUploadFile(pLocalPath, lFile) Then
        UploadFile = True
        RaiseProgress 0, 0, "Done"
      Else
        UploadFile = False
        RaiseProgress 0, 0, "Error on upload... " & lFile
      End If
    End If
  End If
End Function



VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "CCRPFTV6.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Casper Advanced HTML Editor"
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10455
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9630
      Top             =   2475
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PTAB 
      Align           =   3  'Align Left
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   7380
      Left            =   0
      MouseIcon       =   "frmMain.frx":0E42
      ScaleHeight     =   7380
      ScaleWidth      =   2340
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   2340
      Begin VB.PictureBox PicCloseTab 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2055
         Picture         =   "frmMain.frx":1284
         ScaleHeight     =   135
         ScaleWidth      =   180
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Hide File Manager"
         Top             =   15
         Width           =   210
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   1
         Left            =   0
         Picture         =   "frmMain.frx":140A
         ScaleHeight     =   150
         ScaleWidth      =   6000
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   30
         Width           =   6000
      End
      Begin TabDlg.SSTab XTAB 
         Height          =   6960
         Left            =   45
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   12277
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         MouseIcon       =   "frmMain.frx":432C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   " "
         TabPicture(0)   =   "frmMain.frx":4348
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "PnewMenu"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Pmenu1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "PMenuExit"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "PPrint"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "DRV"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "DIR1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "FILE"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Sizer"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "DIR"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "ListFiles"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   " "
         TabPicture(1)   =   "frmMain.frx":46FA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "JavaList"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " "
         TabPicture(2)   =   "frmMain.frx":4A8C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SnippList"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "<>"
         TabPicture(3)   =   "frmMain.frx":4C5E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "TRV"
         Tab(3).Control(1)=   "TreeIMG"
         Tab(3).ControlCount=   2
         Begin ComctlLib.ListView ListFiles 
            Height          =   3045
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   3270
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   5371
            View            =   2
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ImageList TreeIMG 
            Left            =   -74280
            Top             =   3240
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   14
            ImageHeight     =   14
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":4C7A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":4F3A
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView TRV 
            Height          =   6375
            Left            =   -74880
            TabIndex        =   15
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   11245
            _Version        =   393217
            Indentation     =   0
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "TreeIMG"
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CCRPFolderTV6.FolderTreeview DIR 
            Height          =   2700
            Left            =   120
            TabIndex        =   14
            Top             =   510
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   4763
         End
         Begin VB.PictureBox Sizer 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   585
            Left            =   720
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   42
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   5385
            Visible         =   0   'False
            Width           =   630
         End
         Begin MSComctlLib.ListView JavaList 
            Height          =   6285
            Left            =   -74865
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   105
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   11086
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imlToolbarIcons"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Java Snippets"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.FileListBox FILE 
            Appearance      =   0  'Flat
            Height          =   2760
            Left            =   120
            ReadOnly        =   0   'False
            System          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Available Files"
            Top             =   3300
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.DirListBox DIR1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Available Directories"
            Top             =   2790
            Width           =   135
         End
         Begin VB.DriveListBox DRV 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Available Drives"
            Top             =   60
            Width           =   1935
         End
         Begin MSComctlLib.ListView SnippList 
            Height          =   6285
            Left            =   -74865
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   105
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   11086
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imlToolbarIcons"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "HTML Snippets"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.PictureBox PPrint 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   840
            Picture         =   "frmMain.frx":510E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   12
            Top             =   5910
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox PMenuExit 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   840
            Picture         =   "frmMain.frx":5210
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   13
            Top             =   4470
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Pmenu1 
            Height          =   195
            Left            =   840
            Picture         =   "frmMain.frx":5312
            Top             =   2670
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image PnewMenu 
            Height          =   195
            Left            =   840
            Picture         =   "frmMain.frx":53FC
            Top             =   4350
            Visible         =   0   'False
            Width           =   195
         End
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7800
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   476
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8361
            Text            =   "Status: Ready to Edit"
            TextSave        =   "Status: Ready to Edit"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4410
            Text            =   "Line:1 Character:1"
            TextSave        =   "Line:1 Character:1"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:41 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   9645
      Top             =   465
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   9600
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54E6
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55F8
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":570A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":581C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":592E
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A40
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B52
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C64
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D76
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E88
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F9A
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60AC
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61BE
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":62D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6724
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":760C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7720
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":80DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8530
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8984
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":916C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9490
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMainToolBarImageList 
      Left            =   9600
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9858
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9BEC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar Cool 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   741
      VariantHeight   =   0   'False
      EmbossShadow    =   12632256
      _CBWidth        =   10455
      _CBHeight       =   420
      _Version        =   "6.0.8169"
      Child1          =   "tbToolBar"
      MinHeight1      =   330
      Width1          =   8955
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "CoFonts"
      MinHeight2      =   315
      Width2          =   1515
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      MinHeight3      =   360
      NewRow3         =   0   'False
      BandStyle3      =   1
      Begin VB.ComboBox CoFonts 
         Height          =   315
         Left            =   9150
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   45
         Width           =   1185
      End
      Begin MSComctlLib.Toolbar tbToolBar 
         Height          =   330
         Left            =   165
         TabIndex        =   17
         Top             =   45
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   30
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "New"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "OpenWWW"
               ImageIndex      =   27
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Bold"
               ImageKey        =   "Bold"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Italic"
               ImageKey        =   "Italic"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Left"
               Object.ToolTipText     =   "Align Left"
               ImageKey        =   "Align Left"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "Center"
               ImageKey        =   "Center"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Right"
               Object.ToolTipText     =   "Align Right"
               ImageKey        =   "Align Right"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Image Map"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Spell Check"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Char"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Tags"
               ImageIndex      =   21
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TagsIns"
               ImageIndex      =   23
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Frames"
               ImageIndex      =   25
            EndProperty
            BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Tables"
               ImageIndex      =   26
            EndProperty
            BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Colors"
               ImageIndex      =   24
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   6
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "html"
                     Text            =   "HTML Coloring"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "vb"
                     Text            =   "VB Coloring"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "java"
                     Text            =   "Java Coloring"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "mnuseprt"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ColorType"
                     Text            =   "Color as Type"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "NewLine"
                     Text            =   "New Line"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9585
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":9F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":A22A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFILE 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Template 
         Caption         =   "&Template"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu OpenWeb 
         Caption         =   "&Open from The Web ..."
      End
      Begin VB.Menu ConvertTextFile 
         Caption         =   "&Convert text file"
      End
      Begin VB.Menu mnusepp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Begin VB.Menu PrintCode 
            Caption         =   "&Print Code"
            Shortcut        =   ^P
         End
         Begin VB.Menu PrintWeb 
            Caption         =   "&Print As WebPage"
         End
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu SelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu MNUQ 
         Caption         =   "-"
      End
      Begin VB.Menu SelectWrod 
         Caption         =   "&Select Word"
      End
      Begin VB.Menu SentenceNow 
         Caption         =   "&Select Sentence"
      End
      Begin VB.Menu mnuerer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWordWrap 
         Caption         =   "Word Wrap"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu FileManager 
         Caption         =   "&File Manager"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Open Doc"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Doc"
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
      Begin VB.Menu MetaTags 
         Caption         =   "&Meta tags Wizard"
      End
      Begin VB.Menu TableWiz 
         Caption         =   "&Table Wizard"
      End
      Begin VB.Menu FramesWiz 
         Caption         =   "&Frames Wizard"
      End
      Begin VB.Menu ColorSel 
         Caption         =   "&Color Selector"
      End
      Begin VB.Menu TagIns 
         Caption         =   "&Tag Chooser"
      End
      Begin VB.Menu ImageMap 
         Caption         =   "&Image Map"
      End
      Begin VB.Menu CharSet 
         Caption         =   "&Special Character Set"
      End
      Begin VB.Menu Mmnus 
         Caption         =   "-"
      End
      Begin VB.Menu DocWheight 
         Caption         =   "&Document Weight"
      End
   End
   Begin VB.Menu Tags 
      Caption         =   "&Fast Tags"
      Begin VB.Menu StartTag 
         Caption         =   "&Start Tag <>"
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu EndTag 
         Caption         =   "&End Tag </>"
      End
      Begin VB.Menu mnusepppp 
         Caption         =   "-"
      End
      Begin VB.Menu Anchor 
         Caption         =   "&Anchor"
      End
      Begin VB.Menu Image 
         Caption         =   "&Image"
      End
      Begin VB.Menu sepppppp 
         Caption         =   "-"
      End
      Begin VB.Menu Bold 
         Caption         =   "&Bold"
      End
      Begin VB.Menu Italic 
         Caption         =   "&Italic"
      End
      Begin VB.Menu FoSizeoneplus 
         Caption         =   "Font Size +1"
      End
      Begin VB.Menu FoSize1Minus 
         Caption         =   "&Font Size -1"
      End
      Begin VB.Menu seppppp 
         Caption         =   "-"
      End
      Begin VB.Menu Block 
         Caption         =   "&Blockquote"
      End
      Begin VB.Menu Break 
         Caption         =   "&Break"
      End
      Begin VB.Menu Center 
         Caption         =   "&Center"
      End
      Begin VB.Menu Comment 
         Caption         =   "&Comment"
      End
      Begin VB.Menu HorRule 
         Caption         =   "&Horizontal Rule"
      End
      Begin VB.Menu nobreak 
         Caption         =   "&Non-breaking Break"
      End
      Begin VB.Menu Paragraph 
         Caption         =   "&Paragraph"
      End
      Begin VB.Menu sepuuuuuuu 
         Caption         =   "-"
      End
      Begin VB.Menu Head1 
         Caption         =   "&Heading 1"
      End
      Begin VB.Menu head2 
         Caption         =   "&Heading 2"
      End
      Begin VB.Menu head3 
         Caption         =   "&Heading 3"
      End
      Begin VB.Menu sepullll 
         Caption         =   "-"
      End
      Begin VB.Menu EditCurrtag 
         Caption         =   "&Edit Current Tag"
      End
   End
   Begin VB.Menu Syn 
      Caption         =   "&Syntaxing"
      Visible         =   0   'False
      Begin VB.Menu mnuOptionsLower 
         Caption         =   "&Lower Case"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsUpper 
         Caption         =   "&Upper Case"
      End
      Begin VB.Menu mnuOptionsComplete 
         Caption         =   "&Syntaxing ON/OFF"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu Files 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu Filter 
         Caption         =   "&Filter"
         Begin VB.Menu HTM 
            Caption         =   "&HTML Files"
         End
         Begin VB.Menu Pictures 
            Caption         =   "&Picture Files"
         End
         Begin VB.Menu TXT 
            Caption         =   "&Cgi &Pl &Txt Files"
         End
         Begin VB.Menu m 
            Caption         =   "-"
         End
         Begin VB.Menu Custom 
            Caption         =   "&Custom ..."
         End
      End
      Begin VB.Menu Sort 
         Caption         =   "&Sort Files"
         Begin VB.Menu ABC 
            Caption         =   "&ABCDE"
            Enabled         =   0   'False
         End
         Begin VB.Menu EDC 
            Caption         =   "&EDCBA"
         End
      End
      Begin VB.Menu Style 
         Caption         =   "&Style"
         Begin VB.Menu Longa 
            Caption         =   "&Long Style"
         End
         Begin VB.Menu Shorta 
            Caption         =   "&Short Style"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu Java 
      Caption         =   "Java"
      Visible         =   0   'False
      Begin VB.Menu Add 
         Caption         =   "&Add New Java Script"
      End
      Begin VB.Menu Refresh 
         Caption         =   "&Refresh List"
      End
      Begin VB.Menu mnus 
         Caption         =   "-"
      End
      Begin VB.Menu sABC 
         Caption         =   "&Sorting ABC"
      End
      Begin VB.Menu sCBA 
         Caption         =   "&Sorting CBA"
      End
   End
   Begin VB.Menu Snippets 
      Caption         =   "Snippets"
      Visible         =   0   'False
      Begin VB.Menu Addsnipp 
         Caption         =   "&Add Snippet"
      End
      Begin VB.Menu RefreshSnipp 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuss 
         Caption         =   "-"
      End
      Begin VB.Menu ABCSnip 
         Caption         =   "&Sorting ABC"
      End
      Begin VB.Menu CBAsnip 
         Caption         =   "&Sorting CBA"
      End
   End
   Begin VB.Menu RTFMenu 
      Caption         =   "RTFMenu"
      Visible         =   0   'False
      Begin VB.Menu Edittag 
         Caption         =   "&Edit Current Tag"
         Shortcut        =   ^E
      End
      Begin VB.Menu InsertTag 
         Caption         =   "&Insert Tag"
      End
      Begin VB.Menu MNUUUUU 
         Caption         =   "-"
      End
      Begin VB.Menu Fileopt 
         Caption         =   "&File"
         Begin VB.Menu NewDoc1 
            Caption         =   "&New"
         End
         Begin VB.Menu OpenDoc 
            Caption         =   "&Open"
         End
         Begin VB.Menu SaveDoc 
            Caption         =   "&Save"
         End
         Begin VB.Menu SaveAsDoc 
            Caption         =   "&Save As"
         End
      End
      Begin VB.Menu mnunn2 
         Caption         =   "-"
      End
      Begin VB.Menu Copy1 
         Caption         =   "&Copy"
      End
      Begin VB.Menu Paste1 
         Caption         =   "&Paste"
      End
      Begin VB.Menu Cut1 
         Caption         =   "&Cut"
      End
      Begin VB.Menu DateTime 
         Caption         =   "&Time/Date"
         Begin VB.Menu Date 
            Caption         =   "&Date"
         End
         Begin VB.Menu Time 
            Caption         =   "&Time"
         End
         Begin VB.Menu DateAndTime 
            Caption         =   "&Date and Time"
         End
      End
      Begin VB.Menu mnuuu1 
         Caption         =   "-"
      End
      Begin VB.Menu NewDoc 
         Caption         =   "&New Document"
      End
      Begin VB.Menu CloseDoc 
         Caption         =   "&Close Current Document"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
' Project:  Casper HTML   v.2.0                              *
' Filename: n/a                                              *
' Author:   Vladimir S. Pekulas Jr.                          *
' Date:     8/16/2000                                        *
' Copyright © 2000 Vladimir S. Pekulas Jr.                   *
'                                                            *
' Use this program as you wish, but please let me know       *
' if you like it. Anyway, you can do whatever you want       *
' with it. I'm not responsible for any demage tough :)       *
'*************************************************************
      '**  SEE frmAbout FOR FULL CREDITS ! **
      
Public NumberI As Integer
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private BMove As Boolean
Private SplitCoord As Single
Private OldX As Single
Private OldY As Single
'
Public CustomExtention As String
Public FileType As String
Private Type ViewSnipps
    intID As Integer
    strTitle As String * 99
    strArtist As String * 100
End Type
Public lDocumentCount As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7

Private Sub ABC_Click()
 ListFiles.SortOrder = lvwAscending
 ABC.Enabled = False
 EDC.Enabled = True
End Sub

Private Sub ABCSnip_Click()
 SnippList.SortOrder = lvwAscending
End Sub

Private Sub Add_Click()
 frmJava.Show 1, fMainForm
End Sub

Private Sub Addsnipp_Click()
 frmSnipp.Show 1, fMainForm
End Sub

Private Sub Anchor_Click()
 TabNumber = 0
 frmTagEdit.Show 1, fMainForm
End Sub

Private Sub Block_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<blockquote></blockquote>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 13
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub Bold_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<b></b>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 4
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub Break_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<br>"
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub CBAsnip_Click()
 SnippList.SortOrder = lvwDescending
End Sub


Private Sub Center_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<center></center>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 9
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub CharSet_Click()
On Error Resume Next
 Unload frmChar
 frmChar.Show 1, fMainForm
End Sub

Private Sub CloseDoc_Click()
On Error Resume Next
  If MsgBox("Are you sure to close current document ?", vbQuestion + vbYesNo, "Close document ?") = vbYes Then
        lDocumentCount = lDocumentCount - 1
        Unload fMainForm.ActiveForm
   Else
        Exit Sub
   End If
End Sub

Private Sub CoFonts_Click()
 On Error Resume Next
 ActiveForm.rtfText.SelText = "<font face=" & Chr(34) & CoFonts.Text & Chr(34) & ">"
End Sub

Private Sub ColorSel_Click()
 frmColor.Show 1, fMainForm
End Sub

Private Sub Comment_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<!---  --->"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 5
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub ConvertTextFile_Click()
On Error Resume Next
Dim intFileNum As Integer
Dim strFilename As String
 
 With dlgCommonDialog
  .ShowOpen
  .CancelError = False
  sFile = .FileName
 End With
  
If sFile = "" Then Exit Sub
Me.MousePointer = 11
  LoadNewDoc
  'ActiveForm.rtfText3.Language = hlNOHighLight
strTextLineNew = vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>Untitled</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body bgcolor=" & Chr(34) & "white" & Chr(34) & ">" & vbCrLf & vbCrLf
intFileNum = FreeFile
Open sFile For Input As #intFileNum
Do While Not EOF(intFileNum)
 Line Input #intFileNum, strTextLine
 strTextLineNew = strTextLineNew & strTextLine & "<br>" & vbCrLf
Loop
'ActiveForm.rtfText3.Text = ActiveForm.rtfText3.Text & strTextLine & "<br>" & vbCrLf
'ActiveForm.rtfText3.Text = ActiveForm.rtfText3.Text & vbCrLf & "<br>"
Close #intFileNum
 ActiveForm.rtfText.Text = ""
 ActiveForm.rtfText.SelText = strTextLineNew & vbCrLf & "</body>" & vbCrLf & "</html>"
Me.MousePointer = 0
End Sub

Private Sub Cool_HeightChanged(ByVal NewHeight As Single)
On Error Resume Next
 XTAB.Height = Me.Height - Cool.Height - 900 - sbStatusBar.Height
 FILE.Height = XTAB.Height - 3700
 ListFiles.left = FILE.left
 ListFiles.top = FILE.top
 ListFiles.Height = FILE.Height
 JavaList.Height = XTAB.Height - 600
 SnippList.Height = XTAB.Height - 600
End Sub

Private Sub Copy1_Click()
 On Error Resume Next
 Clipboard.SetText ActiveForm.rtfText.SelText
End Sub

Private Sub Cut1_Click()
 On Error Resume Next
 Clipboard.SetText ActiveForm.rtfText.SelText
 ActiveForm.rtfText.SelText = vbNullString
End Sub

Private Sub Date_Click()
 ActiveForm.rtfText.SelText = Date
End Sub

Private Sub DIR_Change()
On Error Resume Next
FILE.Path = DIR.SelectedFolder
ListFilesWicons
End Sub

Private Sub DateAndTime_Click()
ActiveForm.rtfText.SelText = Now
End Sub

Private Sub DIR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0
End Sub

Private Sub DIR_SelectionChange(Folder As CCRPFolderTV6.Folder, PreChange As Boolean, Cancel As Boolean)
On Error Resume Next
PreChange = False
  If PreChange Then
    
    ' If by chance the Folder about to be set is the desktop folder
    ' (set to a different drive than Drive1 is set to), then the drive's
    ' disk may have been removed causing the FTV's AutoUpdate
    ' property to kick in and change it's root to the desktop folder.
    ' This is default behavior (if the FTV's root is no longer valid.).
    With DIR
      If Folder = .GetSpecialFolderName(ftvDesktop) Then
        ' We'll set Drive1 (and the FTV) to the C drive, invoking a
        ' Drive1_Change. We could (and should) set Cancel to True
        ' here, but since we're changing the FTV's root folder, the FTV
        ' will figure it out and not process the desktop folder's selection.
        DRV = "c:\"
      End If
    End With
    
  Else   ' PreChange = False
  
    ' Set File1's contents
    
'    FILE.Path = "c:\" 'Folder.FullPath
    If DIR.SelectedFolder = "Desktop" Then
     FILE.Path = "c:\windows\desktop\"
     GoTo Continue:
    End If
     
    If DIR.SelectedFolder = "My Computer" Then
     FILE.Path = "c:\windows\desktop\"
     GoTo Continue:
    End If
     
    If DIR.SelectedFolder = "My Documents" Then
     FILE.Path = "c:\My Documents\"
     GoTo Continue:
    End If
     FILE.Path = DIR.SelectedFolder
    
Continue:

    ListFiles.ListItems.Clear
    For i = 0 To FILE.ListCount - 1
     ListFiles.ListItems.Add , , FILE.List(i), , 2
    Next i
    
    If FILE.ListCount Then
    ' Select the first file in File1, invokes a File1_Click
      FILE.ListIndex = 0
    Else
      ' No files in this folder, use the folder's path
      'Label1 = File1.Path
    End If
  
  End If   ' PreChange = False


End Sub

Private Sub DocWheight_Click()
Dim intFile As Integer
 intFile = FreeFile
 Open App.Path & "\Casper~temp.html" For Output As #intFile
 Print #intFile, fMainForm.ActiveForm.rtfText.Text
 Close #intFile

 frmDocSize.Show 1, fMainForm
End Sub

Private Sub DRV_Change()
  Dim sDrive As String
  
  
  sDrive = DRV.Drive
  m_sGoodDrive = sDrive
  
  DIR.RootFolder.Selected = True
  DIR.RootFolder = sDrive & "\"
mDRV = Mid(DRV.Drive, 1, 2)
If mDRV = "c:" Then
DIR.SelectedFolder = "C:\"
End If

End Sub

Private Function NormalizePath(ByVal sPath As String) As String
'  If Right$(sPath, 1) = "\" Then
'    NormalizePath = sPath
'  Else
'    NormalizePath = sPath & "\"
'  End If
End Function
  
' Returns True if sPath is a valid file system or UNC folder path.
' Used to validate Drive1's Drive property.

Private Function IsExplicitFolderPath(ByVal sPath As String) As Boolean
  On Error GoTo Out
  
  ' Prevents relative finds...
  'If Left$(sPath, 2) = "\\" Or Mid$(sPath, 2, 2) = ":\" Then
    
    ' Will err if sPath points to a file
   ' IsExplicitFolderPath = Len(DRV(NormalizePath(sPath) & "*.*", _
                                                    vbDirectory Or vbHidden Or _
                                                    vbReadOnly Or vbSystem))   ' 23
 ' End If
' sExplicitFolderPath = True
Out:
End Function


Private Sub EDC_Click()
 ListFiles.SortOrder = lvwDescending
 ABC.Enabled = True
 EDC.Enabled = False
End Sub

Private Sub EditCurrtag_Click()
 frmTagEdit.Show 1, fMainForm
End Sub

Private Sub Edittag_Click()
On Error Resume Next
Dim TagRec As String
Dim FullTag As String
  Me.ActiveForm.rtfText.Span "<", False, True        ' Select Full Tag
  Me.ActiveForm.rtfText.Span ">", True, True         ' Select Full Tag
  TagRec = "<" & Me.ActiveForm.rtfText.SelText & ">" ' Add <> to the tag

  TagRec = Mid(Me.ActiveForm.rtfText.SelText, 1, 2)
    If Mid(TagRec, 1, 1) = "/" Then
      MsgBox "Can't edit the end of TAG", vbInformation
      Exit Sub
    End If
'body
    If LCase(TagRec) = "bo" Then
     TabNumber = 1
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'div
    If LCase(TagRec) = "di" Then
     TabNumber = 3
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'anchor
    If LCase(TagRec) = "a " Then
     TabNumber = 0
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'font
    If LCase(TagRec) = "fo" Then
     TabNumber = 4
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'img
    If LCase(TagRec) = "im" Then
     TabNumber = 6
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'select
    If LCase(TagRec) = "se" Then
     TabNumber = 8
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'hr
    If LCase(TagRec) = "hr" Then
     TabNumber = 5
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'textarea
    If LCase(TagRec) = "te" Then
     TabNumber = 10
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'table
    If LCase(TagRec) = "ta" Then
     TabNumber = 13
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'Td
    If LCase(TagRec) = "td" Then
     TabNumber = 14
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'tr
    If LCase(TagRec) = "tr" Then
     TabNumber = 15
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
  
  If LCase(TagRec) = "in" Then
    TagRecNew = LCase(Mid(Me.ActiveForm.rtfText.SelText, 13, 3))

'Submit
     If TagRecNew = Chr(34) & "sub" Then
      TabNumber = 9
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "sub" Then
      TabNumber = 9
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
'Radio
     If TagRecNew = Chr(34) & "rad" Then
      TabNumber = 7
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "rad" Then
      TabNumber = 7
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
'Checkbox
     If TagRecNew = Chr(34) & "che" Then
      TabNumber = 2
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "che" Then
      TabNumber = 2
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
'text
     If TagRecNew = Chr(34) & "tex" Then
      TabNumber = 12
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "tex" Then
      TabNumber = 12
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
'hidden
     If TagRecNew = Chr(34) & "hid" Then
      TabNumber = 11
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "hid" Then
      TabNumber = 11
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
  End If
If TabNumber = 0 Then MsgBox "Tag not supported by TagEditor", vbInformation, "Unsupported Tag"
' End If
End Sub

Private Sub EndTag_Click()
 fMainForm.ActiveForm.rtfText.SelText = "</>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 1
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub FILE_DblClick()
 OpenAFile
End Sub

Private Sub FILE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
   Me.PopupMenu Files
  End If
End Sub

Private Sub FILE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0
End Sub

Private Sub FileManager_Click()
FileManager.Checked = Not FileManager.Checked
PTAB.Visible = FileManager.Checked
End Sub

Private Sub FoSize1Minus_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<font size=" & Chr(34) & "-1" & Chr(34) & "></font>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 7
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub FoSizeoneplus_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<font size=" & Chr(34) & "+1" & Chr(34) & "></font>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 7
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub FramesWiz_Click()
 frmFramesWiz.Show 1, fMainForm
End Sub

Private Sub Head1_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<h1></h1>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 5
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub head2_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<h2></h2>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 5
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub head3_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<h3></h3>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 5
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub HorRule_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<hr></hr>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 5
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub HTM_Click()
 Dim All As Integer
 All = FILE.ListCount - 1
  ListFiles.ListItems.Clear
 For i = 0 To All
   
   SplitName = FILE.List(i)
   Extension = vbNullString
   intPos = Len(SplitName)
  Do While intPos > 0
   Select Case Mid$(SplitName, intPos, 1)
   Case "."
   Extension = Mid$(SplitName, intPos + 1)
   Exit Do
   Case Else
   End Select
   intPos = intPos - 1
  Loop
           If Extension = "html" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "htm" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "HTML" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "Html" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "HTM" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "Htm" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
 Next i
End Sub

Private Sub Image_Click()
 TabNumber = 6
 frmTagEdit.Show 1, fMainForm
End Sub

Private Sub ImageMap_Click()
frmImageMap.Show 1, fMainForm
End Sub

Private Sub InsertTag_Click()
 frmTags.Show 1, fMainForm
End Sub

Private Sub Italic_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<i></i>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 4
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub JavaList_DblClick()
 Dim OurPath As String
If JavaList.SelectedItem.Text = "News Ticker Script" Then
    OurPath = App.Path & "\Java\news.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If


If JavaList.SelectedItem.Text = "Email Form Script" Then
    OurPath = App.Path & "\Java\email.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Status Bar Script" Then
    OurPath = App.Path & "\Java\statusbar.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Counter Script" Then
    OurPath = App.Path & "\Java\counter.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Gallery Script" Then
    OurPath = App.Path & "\Java\gallery.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "IP Address Script" Then
    OurPath = App.Path & "\Java\ip.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Text Effect Script" Then
    OurPath = App.Path & "\Java\textfx.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Redirection Script" Then
    OurPath = App.Path & "\Java\redirect.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Resolution Script" Then
    OurPath = App.Path & "\Java\resolution.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Scroller Script" Then
    OurPath = App.Path & "\Java\scroller.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

' Now we are going to extract the path of snippet file
' that's going to be open. That's for Users added snipps.
 Dim udtCDToView As ViewSnipps
 Dim intCDFile As Integer, lngRecLength As Long
 Dim lngTotalRecords As Long, lngCDID As Long
 Dim intFileNum As Integer
 Dim NumRecords As Long
 '
 Dim Text As String
 intFileNum = FreeFile
 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtCDToView)
 Open App.Path & "\JavaIndex.dat" For Random As #intCDFile Len = lngRecLength

 '# of Rec.
 If LOF(intFileNum) Mod lngRecLength = 0 Then
 NumRecords = (LOF(intCDFile) \ lngRecLength)
 Else
 NumRecords = (LOF(intCDFile) \ lngRecLength) + 1
 End If
 lngTotalRecords = NumRecords

 'View Rec if Valid
     If lngTotalRecords = 0 Then
 MsgBox "Error 001 - Can not read the record"
 End If
 lngCDID = 0
 Do
 lngCDID = lngCDID + 1
 Get #intCDFile, lngCDID, udtCDToView
 If udtCDToView.strTitle = JavaList.SelectedItem.Text Then
 Trim udtCDToView.strArtist
 LoadNewDoc
 fMainForm.ActiveForm.rtfText.LoadFile udtCDToView.strArtist
 Exit Sub
 End If
 ' udtCDToView.strTitle  ' That's what we search for
 ' udtCDToView.strArtist ' and this is the path to the file
 Loop
 Close #intCDFile
End Sub

Private Sub JavaList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
 Me.PopupMenu Java
End If
End Sub

Function OpenAFile()
Dim intPos As Integer
Dim intFileNum As Integer
Dim strTextLine As String, strFilename As String
 
 'Get Path First
 'If DIR.SelectedFolder = "C:\" Then
 ' MyPath = Mid(DIR.SelectedFolder, 1, 2)
 'Else
 ' MyPath = DIR.SelectedFolder
 'End If
 MyPath = DIR.SelectedFolder
  If Len(MyPath) = 3 Then
   MyPath = Mid(MyPath, 1, 2)
  End If
 
'// I'm still having some troubles with virtual folders so for now
'// let's do a temporary solution.     <:)
 
    If MyPath = "Desktop" Then
     MyPath = "c:\windows\desktop"
    End If
     
    If MyPath = "My Computer" Then
     MyPath = "c:\windows\desktop"

    End If
     
    If MyPath = "My Documents" Then
     MyPath = "c:\My Documents"

    End If
 
 '# Is it a Picture ?
 sbStatusBar.Panels(1).Text = "Status: Opening File ..."
 

 
 If FileType = "List" Then SplitName = ListFiles.SelectedItem.Text
 If FileType = "FILE" Then SplitName = FILE.FileName
 
 'Get The Extention
 Extension = vbNullString
 intPos = Len(SplitName)
 Do While intPos > 0
      Select Case Mid$(SplitName, intPos, 1)
          Case "."
            Extension = Mid$(SplitName, intPos + 1)
            Exit Do
          Case Else
      End Select
  intPos = intPos - 1
 Loop
 
If LCase(Extension) = "gif" Then
 Sizer.Picture = LoadPicture(MyPath & "\" & SplitName)
 PicRatioW = Int(Sizer.Width * (0.064367816091954) - 1.33)
 PicRatioH = Int(Sizer.Height * (0.064367816091954) - 1.33)
 ActiveForm.rtfText.SelText = "<img src=" & Chr(34) & MyPath & "\" & SplitName & Chr(34) & " width=" & Chr(34) & PicRatioW & Chr(34) & " height=" & Chr(34) & PicRatioH & Chr(34) & " alt=" & Chr(34) & Chr(34) & " Border=" & Chr(34) & "0" & Chr(34) & ">"
 Exit Function
End If

If LCase(Extension) = "jpg" Then
 Sizer.Picture = LoadPicture(MyPath & "\" & SplitName)
 PicRatioW = Int(Sizer.Width * (0.064367816091954) - 1.33)
 PicRatioH = Int(Sizer.Height * (0.064367816091954) - 1.33)
 ActiveForm.rtfText.SelText = "<img src=" & Chr(34) & MyPath & "\" & SplitName & Chr(34) & " width=" & Chr(34) & PicRatioW & Chr(34) & " height=" & Chr(34) & PicRatioH & Chr(34) & " alt=" & Chr(34) & Chr(34) & " Border=" & Chr(34) & "0" & Chr(34) & ">"
 Exit Function
End If
 frmDocument.rtfText.LoadFile MyPath & "\" & SplitName
 lDocumentCount = lDocumentCount + 1
 fMainForm.ActiveForm.rtfText.ToolTipText = MyPath & "\" & SplitName
 sbStatusBar.Panels(1).Text = "Status: Ready to Edit"
 fMainForm.ActiveForm.rtfText.ToolTipText = MyPath & "\" & SplitName
 
 sFile = MyPath & "\" & SplitName
 '# Recent.dat
 intFileNum = FreeFile
 Open App.Path & "\recent.dat" For Append As #intFileNum
 Print #intFileNum, sFile
 Close #intFileNum
 
End Function

Private Sub JavaList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.MousePointer = 0
End Sub

Private Sub ListFiles_DblClick()
 OpenAFile
End Sub

Private Sub ListFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then Me.PopupMenu Files
End Sub

Private Sub ListFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0
End Sub

Private Sub Longa_Click()
 FileType = "FILE"
 FILE.Visible = True
 ListFiles.Visible = False
 Filter.Enabled = False
 Sort.Enabled = False
 Shorta.Enabled = True
 Longa.Enabled = False
End Sub



Private Sub MDIForm_Load()
On Error GoTo ErrorHand:
'################################################################################
'# The Following statemnet will basicly run only he first time you run this app #
'# it will create a file containing the names of all the font on your PC        #
'# and load it to the invisible combo box from which the font names will be     #
'# transfred wherever necessary.                                                #
'#                                                                              #
'# It is way (I mean WAY!) faster this way then loading the fonts each time     #
'# you need them by screen.font ...                                             #
'################################################################################
        '# Check if we have created the file
        Dim intFileNum As Integer, strFilename As String
        strFilename = App.Path & "\fontsQ.txt"
        intFileNum = FreeFile
        Open strFilename For Input As #intFileNum
         Do While Not EOF(intFileNum)
              Line Input #intFileNum, FontQ
              If FontQ = "" Then GoTo Continue:
         Loop
Continue:
        Close #intFileNum
        '##
             If Trim(FontQ) = "" Then
        '# Create a file with all the fonts
        strFilename = App.Path & "\fonts.txt"
        intFileNum = FreeFile
        Open strFilename For Output As #intFileNum
        For i = 1 To Screen.FontCount
        Print #intFileNum, Screen.Fonts(i)
         Next i
        Close #intFileNum
        'Close it
        Open App.Path & "\fontsQ.txt" For Output As #intFileNum
            Print #intFileNum, "1"
        GoTo LoadFonts:
        '##
            Else
        '# Addd fonts to the combo box
LoadFonts:
        Close #intFileNum
        strFilename = App.Path & "\fonts.txt"
        intFileNum = FreeFile
        Open strFilename For Input As #intFileNum
        Do While Not EOF(intFileNum)
            Line Input #intFileNum, FontNameA
            If Trim(FontNameA) = "" Then GoTo Contin:
            CoFonts.AddItem FontNameA
Contin:
        Loop
        Close #intFileNum
         CoFonts.ListIndex = 0
                End If
'###########################################################################
    TabNumber = 0
    FileType = "List"
    LoadNewDoc
    Refresh_Click
    RefreshSnipp_Click
    ListFilesWicons
    Call DRV_Change
    PTAB.Width = 3000
'##
'    With Cool
'        Set .Bands(2).Child = CoFonts
'        Set .Bands(1).Child = tbToolBar
'    End With
'fMainForm.Cool.Bands(1).Width = 6000
'##
Exit Sub
ErrorHand:
    MsgBox "An Error has occured: " & Err.Description & " = " & Err.Number
End Sub

Function ListFilesWicons()
On Error GoTo ErrorHand:
Dim All As Integer
 All = FILE.ListCount - 1
 ListFiles.ListItems.Clear
 For i = 0 To All
  ListFiles.ListItems.Add , , FILE.List(i), , 2
 Next i
Exit Function
ErrorHand:
    MsgBox "An Error has occured: " & Err.Description & " = " & Err.Number
End Function

Public Sub LoadNewDoc()
On Error GoTo ErrorHand:
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Show
    ActiveForm.rtfText.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//CASPER HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    ActiveForm.rtfText.Text = ActiveForm.rtfText.Text & vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "         <title>Untitled</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
Exit Sub
ErrorHand:
    MsgBox "An Error has occured: " & Err.Description & " = " & Err.Number
End Sub

Function LoadNewDocFunction()
On Error Resume Next
    Dim frmD As frmDocument
    fMainForm.lDocumentCount = fMainForm.lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Show
    ActiveForm.rtfText.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//CASPER HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    ActiveForm.rtfText.Text = ActiveForm.rtfText.Text & "<html>" & vbCrLf & "<head>" & vbCrLf & "         <title>Untitled</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
End Function

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 If Button = 2 Then
     CloseDoc.Enabled = False
     Paste1.Enabled = False
     Cut1.Enabled = False
     Copy1.Enabled = False
     InsertTag.Enabled = False
     Edittag.Enabled = False
     Fileopt.Enabled = False
     Date.Enabled = False
     Me.PopupMenu RTFMenu
 End If
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0
End Sub

Private Sub MDIForm_Resize()
On Error GoTo Error:
 If Me.WindowState = 1 Then Exit Sub '// On minimize go lost
 XTAB.Height = Me.Height - Cool.Height - 900 - sbStatusBar.Height
 FILE.Height = XTAB.Height - 3700
 ListFiles.left = FILE.left
 ListFiles.top = FILE.top
 ListFiles.Height = FILE.Height
 JavaList.Height = XTAB.Height - 600
 SnippList.Height = XTAB.Height - 600
 Exit Sub
Error:
 MsgBox "An error occured, restoring aplication's environment.", vbCritical, "Error Resizing"
 Me.Width = 10575
 Me.Height = 8760
 Resume Next
End Sub

Private Sub MDIForm_Terminate()
End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo MyErr:
 'Unload fMainForm.ActiveForm
 End
Exit Sub
MyErr:
 MsgBox Err.Description & " #:" & Err.Number
End Sub


Private Sub MetaTags_Click()
 frmMetaTags.Show 1, fMainForm
End Sub

Private Sub mnuEditWordWrap_Click()
 If mnuEditWordWrap.Checked = True Then
  ActiveForm.rtfText.RightMargin = ActiveForm.rtfText.Width * 10
  mnuEditWordWrap.Checked = False
  Exit Sub
 Else
  ActiveForm.rtfText.RightMargin = 0
  mnuEditWordWrap.Checked = True
 End If
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show 1, fMainForm
End Sub

Private Sub mnuOptionsComplete_Click()
 mnuOptionsComplete.Checked = Not mnuOptionsComplete.Checked
End Sub

Private Sub mnuOptionsLower_Click()
    mnuOptionsLower.Checked = Not mnuOptionsLower.Checked
    mnuOptionsUpper.Checked = Not mnuOptionsUpper.Checked
End Sub

Private Sub mnuOptionsUpper_Click()
 mnuOptionsLower_Click
End Sub

Private Sub NewDoc_Click()
    LoadNewDoc
End Sub

Private Sub NewDoc1_Click()
 LoadNewDoc
End Sub

Private Sub nobreak_Click()
 fMainForm.ActiveForm.rtfText.SelText = "&nbsp;"
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub OpenDoc_Click()
 mnuFileOpen_Click
End Sub

Private Sub OpenWeb_Click()
 frmOpenWWW.Show 1, fMainForm
End Sub

Private Sub Paragraph_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<p></p>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 4
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub Paste1_Click()
 On Error Resume Next
 ActiveForm.rtfText.SelText = Clipboard.GetText
End Sub

Private Sub PicCloseTab_Click()
 PTAB.Visible = False
 FileManager.Checked = False
End Sub

Private Sub PicCloseTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fMainForm.MousePointer = 0
End Sub

Private Sub Picture1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
fMainForm.MousePointer = 0
End Sub

Private Sub Pictures_Click()
On Error Resume Next
 Dim All As Integer
 All = FILE.ListCount - 1
  ListFiles.ListItems.Clear
 For i = 0 To All
   
   SplitName = FILE.List(i)
   Extension = vbNullString
   intPos = Len(SplitName)
  Do While intPos > 0
   Select Case Mid$(SplitName, intPos, 1)
   Case "."
   Extension = Mid$(SplitName, intPos + 1)
   Exit Do
   Case Else
   End Select
   intPos = intPos - 1
  Loop
           If Extension = "GIF" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "Gif" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "gif" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "jpg" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "JPG" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "Jpg" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
 Next i
End Sub

Private Sub PrintCode_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .flags = .flags + cdlPDAllPages
        Else
            .flags = .flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hdc
        End If
    End With
End Sub

Private Sub PrintWeb_Click()
 'frmDocument.Web.ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER)
End Sub

Private Sub PTAB_Resize()
If PTAB.Width < 2530 Then PTAB.Width = 2530
If PTAB.Width > 6000 Then PTAB.Width = 6000
 XTAB.Width = PTAB.Width - 150
 ListFiles.Width = XTAB.Width - 250
 DIR.Width = XTAB.Width - 250
 DRV.Width = XTAB.Width - 250
 FILE.Width = ListFiles.Width
 'Java
 JavaList.Width = FILE.Width
 'HTML
 SnippList.Width = FILE.Width
 TRV.Width = SnippList.Width
 TRV.Height = SnippList.Height
 TRV.top = SnippList.top
 'TRV.Left = SnippList.Left
 'CloseButt
 PicCloseTab.left = PTAB.Width - 250
 Picture1(1).Width = PTAB.Width - 250
 JavaList.ColumnHeaders(1).Width = JavaList.Width - 65    'Estetics thingi
 SnippList.ColumnHeaders(1).Width = SnippList.Width - 65  'Estetics thingi
 Cool.Refresh
 tbToolBar.Visible = False
 tbToolBar.Visible = True
End Sub

Private Sub Refresh_Click()
On Error Resume Next
 Dim IRef As Integer
 'Delete it first !
 For IRef = 1 To JavaList.ListItems.Count
  JavaList.ListItems.Remove (1)
 Next IRef
 ' Load it again ! (Our Own)
 JavaList.ListItems.Add , , "Counter Script", , 14
 JavaList.ListItems.Add , , "Gallery Script", , 14
 JavaList.ListItems.Add , , "IP Address Script", , 14
 JavaList.ListItems.Add , , "Text Effect Script", , 14
 JavaList.ListItems.Add , , "Redirection Script", , 14
 JavaList.ListItems.Add , , "Resolution Script", , 14
 JavaList.ListItems.Add , , "Scroller Script", , 14
 JavaList.ListItems.Add , , "Status Bar Script", , 14
 JavaList.ListItems.Add , , "Email Form Script", , 14
 JavaList.ListItems.Add , , "News Ticker Script", , 14
 ' Load it again ! (Users)
 Dim udtJavaToView As ViewSnipps
 Dim intJavaFile As Integer, lngRecLengthJava As Long
 Dim lngTotalRecordsJava As Long, lngJavaID As Long
 Dim intFileNumJava As Integer
 Dim NumRecordsJava As Long
 intFileNumJava = FreeFile
 'Open File
 intJavaFile = FreeFile
 lngRecLengthJava = LenB(udtJavaToView)
 Open App.Path & "\JavaIndex.dat" For Random As #intJavaFile Len = lngRecLengthJava

 
 If LOF(intFileNumJava) Mod lngRecLengthJava = 0 Then
  NumRecordsJava = (LOF(intJavaFile) \ lngRecLengthJava)
 Else
  NumRecordsJava = (LOF(intJavaFile) \ lngRecLengthJava) + 1
 End If
 lngTotalRecordsJava = NumRecordsJava

 'View Rec if Valid
 If lngTotalRecordsJava = 0 Then
 Exit Sub
 End If
 lngJavaID = 0
 Do
     If lngJavaID = lngTotalRecordsJava Then
     Close #intJavaFile
     Exit Sub
     Else
 lngJavaID = lngJavaID + 1
  If lngJavaID > 0 And lngJavaID <= lngTotalRecordsJava Then
 Get #intJavaFile, lngJavaID, udtJavaToView
 JavaList.ListItems.Add , , udtJavaToView.strTitle, , 14
  End If
     End If
 Loop
 Close #intJavaFile
End Sub

Private Sub RefreshSnipp_Click()
On Error Resume Next
 Dim IRef As Integer
 'Delete it first !
 For IRef = 1 To SnippList.ListItems.Count
  SnippList.ListItems.Remove (1)
 Next IRef
 ' Load it again ! (Our Own)
 SnippList.ListItems.Add , , "Bohemia Gift Finder", , 14
 SnippList.ListItems.Add , , "GoTo.com Search Engine", , 14
 SnippList.ListItems.Add , , "InfoSeek.com Search Engine", , 14
 ' Load it again ! (Users)
 Dim udtCDToView As ViewSnipps
 Dim intCDFile As Integer, lngRecLength As Long
 Dim lngTotalRecords As Long, lngCDID As Long
 Dim intFileNum As Integer
 Dim NumRecords As Long
 intFileNum = FreeFile
 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtCDToView)
 Open App.Path & "SnippetIndex.dat" For Random As #intCDFile Len = lngRecLength
 '# of Rec.
 If LOF(intFileNum) Mod lngRecLength = 0 Then
  NumRecords = (LOF(intCDFile) \ lngRecLength)
 Else
  NumRecords = (LOF(intCDFile) \ lngRecLength) + 1
 End If
 lngTotalRecords = NumRecords

 'View Rec if Valid
     If lngTotalRecords = 0 Then
     Exit Sub
     Close #intCDFile
     End If
 lngCDID = 0
 Do
 lngCDID = lngCDID + 1
     If lngCDID > lngTotalRecords Then
 Exit Sub
     Close #intCDFile
     Else
  If lngCDID > 0 And lngCDID <= lngTotalRecords Then
 Get #intCDFile, lngCDID, udtCDToView
 SnippList.ListItems.Add , , udtCDToView.strTitle, , 14
  End If
     End If
 Loop
 Close #intCDFile
End Sub

Private Sub sABC_Click()
 JavaList.SortOrder = lvwAscending
End Sub

Private Sub SaveAsDoc_Click()
 mnuFileSaveAs_Click
End Sub

Private Sub SaveDoc_Click()
 mnuFileSave_Click
End Sub

Private Sub sCBA_Click()
 JavaList.SortOrder = lvwDescending
End Sub

Private Sub SelectAll_Click()
   With fMainForm.ActiveForm.rtfText
        .SetFocus
        .SelStart = 0
        .SelLength = Len(fMainForm.ActiveForm.rtfText.Text)
    End With
End Sub

Private Sub SelectWrod_Click()
 ActiveForm.rtfText.Span " ", False, True
  ActiveForm.rtfText.Span " ", True, True
End Sub

Private Sub SentenceNow_Click()
    With ActiveForm.rtfText
        .Span ".?!:", True, True
        .SelLength = .SelLength + 1
    End With
End Sub

Private Sub Shorta_Click()
On Error Resume Next
 FILE.Visible = False
 ListFiles.Visible = True
 Filter.Enabled = True
 Sort.Enabled = True
 Shorta.Enabled = False
 Longa.Enabled = True
End Sub

Private Sub SnippList_DblClick()
On Error Resume Next
' First load up snippets that comes with Casper HTML
If SnippList.SelectedItem.Text = "InfoSeek.com Search Engine" Then
    OurPath = App.Path & "\Snipps\infoseek.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath

    Exit Sub
End If

If SnippList.SelectedItem.Text = "GoTo.com Search Engine" Then
    OurPath = App.Path & "\Snipps\goto.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If SnippList.SelectedItem.Text = "Bohemia Gift Finder" Then
    OurPath = App.Path & "\Snipps\bohemia.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

' Now we are going to extract the path of snippet file
' that's going to be open.
 Dim udtCDToView As ViewSnipps
 Dim intCDFile As Integer, lngRecLength As Long
 Dim lngTotalRecords As Long, lngCDID As Long
 Dim intFileNum As Integer
 Dim NumRecords As Long
 Dim Text As String
 intFileNum = FreeFile
 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtCDToView)
 Open App.Path & "\SnippetIndex.dat" For Random As #intCDFile Len = lngRecLength

'# of Rec.
 If LOF(intFileNum) Mod lngRecLength = 0 Then
  NumRecords = (LOF(intCDFile) \ lngRecLength)
 Else
  NumRecords = (LOF(intCDFile) \ lngRecLength) + 1
 End If
 lngTotalRecords = NumRecords

 'View Rec if Valid
     If lngTotalRecords = 0 Then
 MsgBox "Error 001 - Can not read the record"
 End If
 lngCDID = 0
 Do
 lngCDID = lngCDID + 1
 Get #intCDFile, lngCDID, udtCDToView
 If udtCDToView.strTitle = SnippList.SelectedItem.Text Then
 Trim udtCDToView.strArtist
 LoadNewDoc
 fMainForm.ActiveForm.rtfText.LoadFile udtCDToView.strArtist
  Exit Sub
 End If
 Loop
 Close #intCDFile
End Sub

Private Sub SnippList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Me.PopupMenu Snippets
End Sub

Private Sub SnippList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.MousePointer = 0
End Sub

Private Sub StartTag_Click()
 fMainForm.ActiveForm.rtfText.SelText = "<>"
 fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 1
 fMainForm.ActiveForm.rtfText.SetFocus
End Sub

Private Sub TableWiz_Click()
 frmTables.Show 1, fMainForm
End Sub

Private Sub TagIns_Click()
 frmTags.Show 1, fMainForm
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
            ActiveForm.rtfText.SelText = "<B> </B>"
        Case "Italic"
            ActiveForm.rtfText.SelText = "<I> </I>"
        Case "Align Left"
            ActiveForm.rtfText.SelText = "<Div align=" & Chr(34) & "Left" & Chr(34) & "> </Div>"
        Case "Center"
            ActiveForm.rtfText.SelText = "<Center> </Center>"
        Case "Align Right"
            ActiveForm.rtfText.SelText = "<Div align=" & Chr(34) & "Right" & Chr(34) & "> </Div>"
        Case "Image Map"
            frmImageMap.Show 1, fMainForm
        Case "Spell Check"
        
        Case "Char"
            frmChar.Show 1, fMainForm
        Case "Tags"
            frmTagEdit.Show 1, fMainForm
        Case "TagsIns"
            frmTags.Show 1, fMainForm
        Case "OpenWWW"
            frmOpenWWW.Show 1, fMainForm
        Case "Frames"
            frmFramesWiz.Show 1, fMainForm
        Case "Tables"
            frmTables.Show 1, fMainForm
    End Select
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuViewWebBrowser_Click()
    'ToDo: Add 'mnuViewWebBrowser_Click' code.
    MsgBox "Add 'mnuViewWebBrowser_Click' code."
End Sub

Private Sub mnuViewOptions_Click()
 frmSettings.Show
End Sub

Private Sub mnuViewRefresh_Click()
 MsgBox "See readme file"
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    Cool.Visible = mnuViewToolbar.Checked
    If Cool.Visible = False Then
      XTAB.Height = XTAB.Height + Cool.Height
    End If
    If Cool.Visible = True Then
      XTAB.Height = XTAB.Height - Cool.Height
    End If
    
    PTAB_Resize
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelText = Clipboard.GetText
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelText
End Sub

Private Sub mnuEditCut_Click()
 On Error Resume Next
 Clipboard.SetText ActiveForm.rtfText.SelText
 ActiveForm.rtfText.SelText = vbNullString
End Sub

Private Sub mnuEditUndo_Click()
Call frmDocument.Undo(True)
End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    End
End Sub

Private Sub mnuFileSend_Click()
    MsgBox "Huh ? Oh, sorry I felt asleep ...."
End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub


Private Sub mnuFileSaveAs_Click()
 SaveDocument
End Sub

Private Sub mnuFileSave_Click()
 SaveDocument
End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String
    If ActiveForm Is Nothing Then LoadNewDoc
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hdc
        End If
        sFile = .FileName
    
'# Recent.dat
 Dim intFileNum As Integer
 Dim strTextLine As String, strFilename As String
 intFileNum = FreeFile
 Open App.Path & "\recent.dat" For Append As #intFileNum
 Print #intFileNum, sFile
 Close #intFileNum
    
    End With
    ActiveForm.rtfText.LoadFile sFile
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub

Private Sub tbToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 Select Case ButtonMenu.index
  Case 1
    fMainForm.ActiveForm.rtfText.Language = hlhtml
  Case 2
    fMainForm.ActiveForm.rtfText.Language = hlVisualBasic
  Case 3
    fMainForm.ActiveForm.rtfText.Language = hlJava
 Case 5
    fMainForm.ActiveForm.rtfText.HighlightCode = hlAsType
 Case 6
    fMainForm.ActiveForm.rtfText.HighlightCode = hlOnNewLine
 End Select
End Sub

Private Sub tbToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0
End Sub

Private Sub Template_Click()
On Error Resume Next
 Dim intFileNum As Integer
 Dim strTextLine As String, strFilename As String
 intFileNum = FreeFile
 sDat = "0"
 Open App.Path & "\exitfront.txt" For Output As #intFileNum
 Print #intFileNum, sDat
 Close #intFileNum
 frmFront.Show 1, fMainForm
End Sub

Private Sub Time_Click()
ActiveForm.rtfText.SelText = Time
End Sub

Private Sub TRV_DblClick()
On Error Resume Next
 fMainForm.ActiveForm.rtfText.SelText = TRV.SelectedItem.Text
End Sub

Private Sub TRV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 Me.MousePointer = 0
End Sub

Private Sub TXT_Click()
On Error Resume Next
 Dim All As Integer
 All = FILE.ListCount - 1
  ListFiles.ListItems.Clear
 For i = 0 To All
   
   SplitName = FILE.List(i)
   Extension = vbNullString
   intPos = Len(SplitName)
  Do While intPos > 0
   Select Case Mid$(SplitName, intPos, 1)
   Case "."
   Extension = Mid$(SplitName, intPos + 1)
   Exit Do
   Case Else
   End Select
   intPos = intPos - 1
  Loop
           If Extension = "txt" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "TXT" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "pl" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "PL" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "Pl" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "Cgi" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "CGI" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
           If Extension = "cgi" Then ListFiles.ListItems.Add , , FILE.List(i), , 2
 Next i
End Sub



Private Sub Custom_Click()
On Error Resume Next
Dim All As Integer
'Ask first? Duh ...
ExtensionOwn = InputBox("Enter your own extension such as 'txt'." & vbCrLf & "Case sensitive!", "Custom Extension")
If ExtensionOwn = "" Then Exit Sub
  
  All = FILE.ListCount - 1
  ListFiles.ListItems.Clear
 For i = 0 To All
   SplitName = FILE.List(i)
   Extension = vbNullString
   intPos = Len(SplitName)
  Do While intPos > 0
   Select Case Mid$(SplitName, intPos, 1)
   Case "."
   Extension = Mid$(SplitName, intPos + 1)
   Exit Do
   Case Else
   End Select
   intPos = intPos - 1
  Loop
          If Extension = ExtensionOwn Then ListFiles.ListItems.Add , , FILE.List(i), , 2
 Next i
End Sub







'Resize on Fly
Private Sub PTAB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
Me.MousePointer = 9
    BMove = True
    RetVal = SetCapture(PTAB.hWnd)
    OldX = -32
End Sub

Private Sub PTAB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 9
    If BMove Then
        PTAB.DrawMode = 6
        If OldX <> -32 Then
            PTAB.Line (OldX - 15!, 0!)-(OldX + 15!, Height), , BF
        End If
        PTAB.Line (X - 15!, 0!)-(X + 15!, Height), , BF
        OldX = X
    End If
End Sub

Private Sub PTAB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
Me.MousePointer = 0
    If BMove Then
        RetVal = ReleaseCapture()
        BMove = False
        PTAB.Cls
        SplitCoord = X
        PTAB.Width = SplitCoord
    End If
End Sub

Private Sub XTAB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0
End Sub


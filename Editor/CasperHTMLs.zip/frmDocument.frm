VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmDocument 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   5700
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab EV 
      Height          =   5100
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   8996
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit"
      TabPicture(0)   =   "frmDocument.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtfText2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture9(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "AutoSyntaxPic"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "BrowserTest"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Picture3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdTable"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture9(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExitDoc"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdFind"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdFullView"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdReplace"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdOpenDoc"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdColorEdit"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdSepartate"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "rtfText"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "rtfText3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "PicLines"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "frmDocument.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "P2"
      Tab(1).Control(1)=   "pResizeWeb"
      Tab(1).Control(2)=   "cmdFor"
      Tab(1).Control(3)=   "cmdBack"
      Tab(1).Control(4)=   "cmdRef"
      Tab(1).ControlCount=   5
      Begin VB.PictureBox PicLines 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7695
         Left            =   480
         ScaleHeight     =   7665
         ScaleWidth      =   375
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   400
      End
      Begin RichTextLib.RichTextBox rtfText3 
         Height          =   630
         Left            =   4680
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1111
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         TextRTF         =   $"frmDocument.frx":0038
      End
      Begin Project1.CodeHighlight rtfText 
         CausesValidation=   0   'False
         Height          =   2415
         Left            =   960
         TabIndex        =   23
         Top             =   360
         Width           =   4455
         _extentx        =   8281
         _extenty        =   4260
         language        =   3
         keywordcolor    =   0
         operatorcolor   =   0
         delimitercolor  =   16711680
         forecolor       =   0
         functioncolor   =   0
         highlightcode   =   0
         font            =   "frmDocument.frx":00F9
      End
      Begin VB.PictureBox cmdRef 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   -74903
         Picture         =   "frmDocument.frx":011D
         ScaleHeight     =   195
         ScaleWidth      =   150
         TabIndex        =   20
         ToolTipText     =   "Refresh"
         Top             =   1455
         Width           =   150
      End
      Begin VB.PictureBox cmdBack 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   -74910
         Picture         =   "frmDocument.frx":0179
         ScaleHeight     =   165
         ScaleWidth      =   135
         TabIndex        =   19
         ToolTipText     =   "Back"
         Top             =   960
         Width           =   135
      End
      Begin VB.PictureBox cmdFor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   -74910
         Picture         =   "frmDocument.frx":01CC
         ScaleHeight     =   165
         ScaleWidth      =   135
         TabIndex        =   18
         ToolTipText     =   "Forward"
         Top             =   1185
         Width           =   135
      End
      Begin VB.PictureBox cmdSepartate 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   120
         Picture         =   "frmDocument.frx":0221
         ScaleHeight     =   165
         ScaleWidth      =   135
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   945
         Width           =   135
      End
      Begin VB.PictureBox pResizeWeb 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   -74925
         Picture         =   "frmDocument.frx":0278
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   15
         ToolTipText     =   "Size test"
         Top             =   645
         Width           =   195
      End
      Begin VB.PictureBox P2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   3180
         Left            =   -74640
         ScaleHeight     =   3180
         ScaleWidth      =   6015
         TabIndex        =   13
         Top             =   555
         Width           =   6015
         Begin VB.PictureBox PRuler2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   11670
            Left            =   -20
            Picture         =   "frmDocument.frx":02F1
            ScaleHeight     =   11670
            ScaleWidth      =   360
            TabIndex        =   24
            Top             =   285
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.PictureBox Pruler 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   0
            Picture         =   "frmDocument.frx":DE03
            ScaleHeight     =   330
            ScaleWidth      =   18555
            TabIndex        =   22
            ToolTipText     =   "Different Screen Width"
            Top             =   0
            Visible         =   0   'False
            Width           =   18585
         End
         Begin SHDocVwCtl.WebBrowser Web 
            Height          =   1935
            Left            =   0
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
            ExtentX         =   5318
            ExtentY         =   3413
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   0
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin VB.Shape Shape 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   5  'Downward Diagonal
            Height          =   2985
            Left            =   0
            Top             =   240
            Width           =   5940
         End
      End
      Begin VB.PictureBox cmdColorEdit 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   60
         Picture         =   "frmDocument.frx":21D45
         ScaleHeight     =   240
         ScaleWidth      =   270
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Color Editor"
         Top             =   2610
         Width           =   270
      End
      Begin VB.PictureBox cmdOpenDoc 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   60
         Picture         =   "frmDocument.frx":22107
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "List of Open Documents"
         Top             =   2070
         Width           =   270
      End
      Begin VB.PictureBox cmdReplace 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   60
         Picture         =   "frmDocument.frx":22421
         ScaleHeight     =   240
         ScaleWidth      =   270
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Replace in Text"
         Top             =   1530
         Width           =   270
      End
      Begin VB.PictureBox cmdFullView 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   75
         Picture         =   "frmDocument.frx":227E3
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Full Screen View"
         Top             =   2355
         Width           =   255
      End
      Begin VB.PictureBox cmdFind 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   60
         Picture         =   "frmDocument.frx":22B31
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Find in Text"
         Top             =   1245
         Width           =   270
      End
      Begin VB.PictureBox cmdExitDoc 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   105
         Picture         =   "frmDocument.frx":22E4B
         ScaleHeight     =   135
         ScaleWidth      =   180
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Close Current Document"
         Top             =   600
         Width           =   180
      End
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   0
         Left            =   60
         Picture         =   "frmDocument.frx":22FD1
         ScaleHeight     =   60
         ScaleWidth      =   270
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1890
         Width           =   270
      End
      Begin VB.PictureBox cmdTable 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   90
         Picture         =   "frmDocument.frx":230F3
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Table Wizard"
         Top             =   2970
         Width           =   225
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Left            =   60
         Picture         =   "frmDocument.frx":23405
         ScaleHeight     =   60
         ScaleWidth      =   270
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   270
      End
      Begin VB.PictureBox BrowserTest 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   75
         Picture         =   "frmDocument.frx":23527
         ScaleHeight     =   255
         ScaleWidth      =   240
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Test With Favorite Browser"
         Top             =   3315
         Width           =   240
      End
      Begin VB.PictureBox AutoSyntaxPic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   75
         Picture         =   "frmDocument.frx":23899
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Gutter"
         Top             =   3930
         Width           =   240
      End
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   1
         Left            =   60
         Picture         =   "frmDocument.frx":23B7B
         ScaleHeight     =   60
         ScaleWidth      =   270
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3750
         Width           =   270
      End
      Begin RichTextLib.RichTextBox rtfText2 
         Height          =   1755
         Left            =   930
         TabIndex        =   17
         Top             =   2925
         Visible         =   0   'False
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   3096
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         RightMargin     =   3
         TextRTF         =   $"frmDocument.frx":23C9D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   9360
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":23D5E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":23E70
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":23F82
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24094
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":241A6
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":242B8
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":243CA
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":244DC
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":245EE
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24700
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24812
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24924
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24A36
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":24F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":250AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   9480
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":251C0
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":252D2
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":253E4
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":254F6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":25608
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":2571A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":2582C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":2593E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":25A50
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":25B62
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":25C74
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":25D86
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":25E98
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":25FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":266DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":26B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":26F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":273DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":2782E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":27C82
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":280D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":2852A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":28982
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":28DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":2922E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":29682
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":29C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":29F02
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
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

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const vbKeyLessThan = 60

Dim pt          As POINTAPI
Dim lngStart    As Long

Dim TextHeigth As Long, fTop As Integer  '// Text height - important
Dim LineCountChange As Integer           '// This is used to determin if we need _
                                             to redraw the numbers
Dim FirstLine As Long                    '// Dim the First visible line
Dim FirstLineNow As Long

Public WebSize As Integer '# What Size is the Web Window in ?
Public Separate As Integer '# Do we have two desktops ?

Private Sub eXIT_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub AutoSyntaxPic_Click()
On Error Resume Next
' fMainForm.mnuOptionsComplete.Checked = Not fMainForm.mnuOptionsComplete.Checked
' rtfText.SetFocus
' If fMainForm.mnuOptionsComplete.Checked = True Then AutoSyntaxPic.ToolTipText = "Syntaxing On"
' If fMainForm.mnuOptionsComplete.Checked = False Then AutoSyntaxPic.ToolTipText = "Syntaxing Off"


If PicLines.Visible = True Then
   PicLines.Visible = False
   Form_Resize
Else
   PicLines.Visible = True
   Form_Resize
   DrawNumbers
End If
End Sub

Private Sub BrowserTest_Click() '# test edited file in default Browser
On Error Resume Next
 Dim strView As String
 Dim intFile As Integer
 intFile = FreeFile
 Open "c:\Casper~temp.html" For Output As #intFile
 Print #intFile, fMainForm.ActiveForm.rtfText.Text
 Close #intFile
 Shell ("start c:\Casper~temp.html")
End Sub

Private Sub cmdBack_Click()
On Error Resume Next
 Web.GoBack
End Sub

Private Sub cmdColorEdit_Click()
On Error Resume Next
 frmColor.Show 1, fMainForm
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
  Dim s As String
  If rtfText.SelLength > 0 Then s = rtfText.SelText Else s = "Find Text"
  ShowFind Me, rtfText, FR_SHOWHELP, s
End Sub

Private Sub cmdFor_Click()
On Error Resume Next
 Web.GoForward
End Sub

Private Sub cmdFullView_Click()
On Error Resume Next
 fMainForm.FileManager.Checked = Not fMainForm.FileManager.Checked
 fMainForm.PTAB.Visible = fMainForm.FileManager.Checked

 fMainForm.mnuViewStatusBar.Checked = Not fMainForm.mnuViewStatusBar.Checked
 fMainForm.sbStatusBar.Visible = fMainForm.mnuViewStatusBar.Checked

 fMainForm.mnuViewToolbar.Checked = Not fMainForm.mnuViewToolbar.Checked
 fMainForm.tbToolBar.Visible = fMainForm.mnuViewToolbar.Checked
End Sub

Private Sub cmdOpenDoc_Click()
On Error Resume Next
MsgBox "Function not yet available." & vbCrLf & "Currently " & fMainForm.lDocumentCount & " document(s) opened.", vbInformation, "Unavailable"
End Sub

Private Sub cmdRef_Click()
On Error Resume Next
 Web.Refresh
End Sub

Private Sub cmdReplace_Click()
On Error Resume Next
  Dim s As String
  If rtfText.SelLength > 0 Then s = rtfText.SelText Else s = "Find Text"
  ShowFind Me, rtfText, FR_SHOWHELP, s, True, "Replace Text"
End Sub

Private Sub cmdSepartate_Click()
On Error Resume Next
 If Separate = 1 Then
  rtfText2.Text = rtfText.Text
  rtfText2.Visible = True
  rtfText.Height = (rtfText.Height / 2) - 50
  rtfText2.left = rtfText.left
  rtfText2.Width = rtfText.Width
  rtfText2.top = (Me.ScaleHeight - 330) - rtfText.Height + 200
  rtfText2.Height = rtfText.Height
  Separate = 0
 Else
  rtfText2.Visible = False
  rtfText.Move 370, 300, Me.ScaleWidth - 470, Me.ScaleHeight - 330
  Separate = 1
 End If
End Sub

Private Sub cmdExitDoc_Click()
On Error Resume Next
  If MsgBox("Are you sure to close current document ?", vbQuestion + vbYesNo, "Close document ?") = vbYes Then
        fMainForm.lDocumentCount = fMainForm.lDocumentCount - 1
fMainForm.sbStatusBar.Panels(2).Text = "Line: 0  Character: 0"
        Unload Me
   Else
        Exit Sub
   End If
End Sub

Private Sub cmdTable_Click()
On Error Resume Next
frmTables.Show 1, fMainForm
End Sub



Private Sub EV_Click(PreviousTab As Integer)
On Error Resume Next
'// Save for View
Dim strView As String
Dim intFile As Integer
intFile = FreeFile
Open App.Path & "\Casper~temp.html" For Output As #intFile
Print #intFile, rtfText.Text
Close #intFile

If EV.Caption = "Edit" Then
 rtfText.Visible = True
 P2.Visible = False
 Web.Visible = False
 fMainForm.sbStatusBar.Panels(1).Text = "Status: Ready to Edit"
 DrawNumbers
 rtfText.SetFocus
Else
 rtfText.Visible = False
 P2.Visible = True
 Web.Visible = True
 Web.Navigate (App.Path & "\Casper~temp.html")
 fMainForm.sbStatusBar.Panels(1).Text = "Status: Visual Viewing"
End If
End Sub



Private Sub EV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fMainForm.MousePointer = 0
End Sub

Private Sub Form_Activate()
 rtfText.SelStart = 128
 DrawNumbers
End Sub

Private Sub Form_Load()
On Error Resume Next
 Separate = 1
 WebSize = 1 '# Set WebWindow Resizer to it's full size
 rtfText.Visible = True
 P2.Visible = False
 Web.Visible = False
 'fMainForm.CoDocs.AddItem "Untitled"
fMainForm.mnuOptionsUpper.Checked = False
rtfText.RightMargin = rtfText.Width * 10
End Sub

Private Sub Form_Paint()
 fMainForm.tbToolBar.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Error:
Unload Me
Exit Sub
Error:
MsgBox Err.Description & " Number:" & Err.Number, vbCritical, "Error"
End Sub

Private Sub Form_Resize()
On Error GoTo Error:
    EV.top = 0
    EV.left = 0
    EV.Height = Me.Height
    EV.Width = Me.Width
    rtfText.Move 370, 300, Me.ScaleWidth - 470, Me.ScaleHeight - 330
    Web.Move 370, 300, Me.ScaleWidth - 470, Me.ScaleHeight - 330
    P2.Move 370, 300, Me.ScaleWidth - 470, Me.ScaleHeight - 330
    Web.top = 0
    Web.left = 0
    Shape.Height = P2.Height
    Shape.Width = P2.Width
    rtfText2.Width = rtfText.Width

If PicLines.Visible = True Then
    PicLines.left = 370
    PicLines.top = 300
    rtfText.Move 370 + PicLines.Width, 300, Me.ScaleWidth - 470 - PicLines.Width, Me.ScaleHeight - 330
    PicLines.Height = rtfText.Height
    DrawNumbers
End If


Exit Sub
Error:
MsgBox Err.Description & " Number:" & Err.Number, vbCritical, "Error"
'Stop
End Sub



Private Sub pResizeWeb_Click() '# Determine and/or change size of WebWindow
On Error GoTo Error:
 If WebSize = 0 Then '# If WebSize=0 then resize to original shape
   'Call Form_Resize
   Web.top = 0
   Web.left = 0
   Web.Width = Me.ScaleWidth - 470
   Web.Height = Me.ScaleHeight - 330
   Pruler.Visible = False
   PRuler2.Visible = False
   WebSize = 1
 Else  '# If WebSize = 1 then resize to smaller WebWindow
   Pruler.Visible = True
   PRuler2.Visible = True
   Pruler.top = 25
   Pruler.left = 0
   PRuler2.left = -20
   PRuler2.top = 285
   'Web.Height = Web.Height - 360
   Web.top = 360
   Web.left = 350
   Web.Width = ((Me.Width - Web.left) / 100) * 80
   Web.Height = ((Me.Height - Web.top) / 100) * 80
   WebSize = 0
 End If
Exit Sub
Error:
MsgBox Err.Description & " Number:" & Err.Number, vbCritical, "Error"
'Stop
End Sub


Private Sub rtfText_Change()
On Error Resume Next
 rtfText2.Text = rtfText.Text
 
  '// Get number of lines in Rtftext
 LineCount = rtfText.GetLineCount
 LineCount = LineCount - 1  '// Change start from 0 to 1

    If LineCount = LineCountChange Then
      Exit Sub    '// Line count is still the same
    Else
      DrawNumbers '// new Line count is required
    End If
 
End Sub

Private Sub rtfText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 If Button = 2 Then
  fMainForm.Fileopt.Enabled = True
  fMainForm.CloseDoc.Enabled = False 'Temporarly: For some reason my VB6 is crashing while closing one of more then 3 documents from pop-up menu ?@#$$!!$
  fMainForm.Paste1.Enabled = True

  If rtfText.SelLength > 0 Then
   fMainForm.Cut1.Enabled = True
  Else
   fMainForm.Cut1.Enabled = False
  End If
   
  If rtfText.SelLength > 0 Then
   fMainForm.Copy1.Enabled = True
  Else
   fMainForm.Copy1.Enabled = False
  End If

  fMainForm.InsertTag.Enabled = True
  fMainForm.Edittag.Enabled = True
  fMainForm.Date.Enabled = True
  
  fMainForm.PopupMenu fMainForm.RTFMenu
 End If
End Sub

Private Sub rtfText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fMainForm.MousePointer = 0
End Sub

Private Sub rtfText2_KeyPress(KeyAscii As Integer)
On Error Resume Next
 rtfText.Text = rtfText2.Text
End Sub

Private Sub lsMain_Done(ByVal Text As String)
On Error Resume Next
    ' Hide the popup window and add the text
    If fMainForm.mnuOptionsComplete.Checked = True Then
        ' Add the tag and close it
        rtfText.SelText = Text & "></" & Text & ">"
        ' Move the caret in between the two tags
        rtfText.SelStart = rtfText.SelStart - Len("</" & Text & ">")
    Else
        ' Add the tag without closing it
        rtfText.SelText = Text & ">"
    End If
    lsMain.Visible = False
    rtfText.SetFocus
End Sub

Private Sub lsMain_Escape()
On Error Resume Next
    ' Hide the popup window and dont add the text
    lsMain.Visible = False
    rtfText.SetFocus
End Sub

Private Sub rtfText_KeyPress(KeyAscii As Integer)
On Error Resume Next
 If fMainForm.mnuOptionsComplete.Checked = True Then
    If KeyAscii = vbKeyLessThan Then
        ' Get the position of the caret
        GetCaretPos pt
        ' Get the selstart
        lngStart = rtfText.SelStart
        ' Move the popup window to the caret
        'lsMain.Move pt.x + rtfText.Font.Size, pt.y + (2 * rtfText.Font.Size)
        ' Check if the popup window is within the form
        'If lsMain.Left + lsMain.Width > ScaleWidth Then lsMain.Move pt.x - lsMain.Width
        'If lsMain.Top + lsMain.Height > ScaleHeight Then lsMain.Move lsMain.Left, pt.y - lsMain.Height
        ' Fill the popup window with tags (only if there are no errors!)
        If lsMain.FillWithTags(App.Path & "\tags.lst", fMainForm.mnuOptionsUpper.Checked) = 0 Then Exit Sub
        ' Fill the popup window with fonts 'lsMain.FillWithFonts ' Fill the popup window with available drives
        'lsMain.FillWithDrives mnuOptionsUpper.Checked ' Show the popup window
        lsMain.Visible = True
        ' Give the window focus
        lsMain.SetFocus
    End If
   Else
 End If
End Sub


Public Sub Undo(ByVal bUndo As Boolean)
On Error Resume Next
    On Error Resume Next
    Dim OK As Long
    
    OK = SendMessageLong(Screen.ActiveForm.ActiveControl.hWnd, EM_UNDO, 0&, 0&)
    If (bUndo) Then
        'mnuRightUndo.Enabled = False
        'mnuUndo.Enabled = False
        'mnuRightRedo.Enabled = True
        'mnuRedo.Enabled = True
        'Toolbar1.Buttons(16).Enabled = False
        'Toolbar1.Buttons(17).Enabled = True
    Else
        'mnuRightUndo.Enabled = True
        'mnuUndo.Enabled = True
        'mnuRightRedo.Enabled = False
        'mnuRedo.Enabled = False
        'Toolbar1.Buttons(16).Enabled = True
        'Toolbar1.Buttons(17).Enabled = False
    End If
    
    Exit Sub
End Sub

Private Sub rtfText2_SelChange()
Call CarrotStatus2
End Sub

Private Sub Web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
 If EV.Tab = 1 Then fMainForm.sbStatusBar.Panels(1).Text = Web.LocationURL
End Sub









Sub DrawNumbers()
On Error GoTo Error:
Dim LineCount As Long '// How many lines in total
Dim i As Integer      '// Just an integer

'// Get number of lines in Rtftext
LineCount = rtfText.GetLineCount
LineCount = LineCount - 1  '// Change start from 0 to 1


'// Same lines ?
LineCountChange = LineCount


'// Get first visible line in rtfText
FirstLine = rtfText.GetFirstLine
FirstLine = FirstLine   '// Change start from 0 to 1 if necessary

PicLines.Cls '// Clear the PicLines
PicLines.CurrentY = 40  '// Move the .top text by 40 twips

'// Print the number of each line on a picture
For i = 0 To LineCount - FirstLine
    PicLines.CurrentY = PicLines.CurrentY + 7.49 '// Where on Y
    PicLines.CurrentX = 20 '-2                   '// Where on X
    PicLines.Print i + FirstLine + 1             '// print the number
Next
 'LineCountChange = LineCount '// Remember the last line count
 FirstLineNow = FirstLine     '// Is the first visible line still the same ?

'// Error Handle
   Exit Sub
Error:
 MsgBox "An Error Occured.", vbCritical
 PicLines.Visible = False
 Form_Resize
End Sub


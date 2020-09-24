VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSnipp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HTML Snippets Library ..."
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Locate Snippet File"
      Height          =   2655
      Left            =   165
      TabIndex        =   9
      Top             =   120
      Width           =   5415
      Begin VB.DriveListBox drvDrive 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.DirListBox DirDirectory 
         Height          =   1890
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.FileListBox filFileName 
         Height          =   1845
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin MSComDlg.CommonDialog Cmd 
         Left            =   1560
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   360
         Left            =   2280
         TabIndex        =   13
         Top             =   2160
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   635
         ButtonWidth     =   5239
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Or Click Here ...                                "
               Key             =   "Click"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   840
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSnipp.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Snippet Properties"
      Height          =   1335
      Left            =   150
      TabIndex        =   0
      Top             =   2940
      Width           =   5415
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Snippet"
         Height          =   340
         Left            =   3840
         TabIndex        =   5
         Top             =   332
         Width           =   1455
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtCDID 
         Height          =   285
         Left            =   5280
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   340
         Left            =   3840
         TabIndex        =   1
         Top             =   812
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Snipp ID:"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "File:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Title:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmSnipp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
' Project:  Casper HTML   v.2.0                              *
' Filename: n/a                                              *
' Author:   Vladimir S. Pekulas Jr.                          *
' Date:     7/22/2000                                        *
' Copyright © 2000 Vladimir S. Pekulas Jr.                   *
'                                                            *
' Use this program as you wish, but please let me know       *
' if you like it. Anyway, you can do whatever you want       *
' with it. I'm not responsible for any demage tough :)       *
'*************************************************************

Option Explicit
Dim NumRecords As Integer
Dim intFileNum As Integer
Dim lngRecLength As Long
Private Type ViewSnipps
    intID As Integer
    strTitle As String * 99
    strArtist As String * 100
End Type
Private Type AddCD
    intID As Integer
    strTitle As String * 99
    strArtist As String * 100
End Type

Private Sub cmdDone_Click()
 Unload Me
End Sub

Private Sub Command1_Click()
 'Refresh List
 Dim IRef As Integer
 'Delete it first !
 For IRef = 1 To fMainForm.SnippList.ListItems.Count
  fMainForm.SnippList.ListItems.Remove (1)
 Next IRef
 ' Load it again ! (Our Own)
 fMainForm.SnippList.ListItems.Add , , "Bohemia Gift Finder", , 14
 fMainForm.SnippList.ListItems.Add , , "GoTo.com Search Engine", , 14
 fMainForm.SnippList.ListItems.Add , , "InfoSeek.com Search Engine", , 14
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
 Unload Me
     Exit Sub
     Close #intCDFile
     End If
 lngCDID = 0
 Do
 lngCDID = lngCDID + 1
     If lngCDID > lngTotalRecords Then
 Unload Me
 Exit Sub
     Close #intCDFile
     Else
  If lngCDID > 0 And lngCDID <= lngTotalRecords Then
 Get #intCDFile, lngCDID, udtCDToView
 fMainForm.SnippList.ListItems.Add , , udtCDToView.strTitle, , 14
  End If
     End If
 Loop
 Close #intCDFile
 Unload Me
End Sub

Private Sub DirDirectory_Change()
 filFileName.Path = DirDirectory.Path
End Sub

Private Sub drvDrive_Change()
 DirDirectory.Path = drvDrive.Drive
End Sub

Private Sub filFileName_Click()
 Dim intFileNum As Integer
 Dim strTextLine As String, strFilename As String
 If Right(DirDirectory.Path, 1) = "\" Then
  strFilename = filFileName.Path & filFileName.Filename
 Else
  strFilename = filFileName.Path & "\" & filFileName.Filename
 End If
 txtArtist.Text = strFilename
 txtArtist.SelStart = Len(txtArtist.Text)
 txtTitle.SetFocus
End Sub

Private Sub Form_Load()
 Dim udtCD As AddCD
 Dim intCDFile As Integer, lngRecLength As Long, lngNextCDID As Long
 Dim NumRecords As Integer
 Dim intFileNum As Integer
 intFileNum = FreeFile
 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtCD)
 Open App.Path & "\SnippetIndex.dat" For Random As #intCDFile Len = lngRecLength
 'Next rec (NUMRECORDS)
 If LOF(intFileNum) Mod lngRecLength = 0 Then
  NumRecords = (LOF(intFileNum) \ lngRecLength)
 Else
  NumRecords = (LOF(intFileNum) \ lngRecLength) + 1
 End If
 lngNextCDID = NumRecords + 1
 txtCDID.Text = lngNextCDID
 txtCDID.Enabled = False
 Close #intCDFile
End Sub


Private Sub cmdAdd_Click()
 Dim udtNewCD As AddCD
 Dim intCDFile As Integer, lngRecLength As Long, lngCDID As Long
 'Check if not Title ""
 If txtTitle.Text = "" Then
 MsgBox ("Please Name the Snippet")
 Exit Sub
 End If
 ' check if not path to file ""
 If txtArtist.Text = "" Then
 MsgBox ("Please Select File to Use as a Snippet.")
 Exit Sub
 End If

 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtNewCD)
 Open App.Path & "\SnippetIndex.dat" For Random As #intCDFile Len = lngRecLength

 'Adds New CD
 lngCDID = txtCDID.Text
 udtNewCD.strArtist = txtArtist.Text
 udtNewCD.strTitle = txtTitle.Text
 
 Put #intCDFile, lngCDID, udtNewCD
 'Make txt ""
 txtCDID.Text = lngCDID + 1
 txtTitle.Text = ""
 txtArtist.Text = ""
 Close #intCDFile
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Click"
Cmd.ShowOpen
txtArtist.Text = Cmd.Filename
End Select
End Sub


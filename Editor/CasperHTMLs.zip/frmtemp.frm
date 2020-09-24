VERSION 5.00
Begin VB.Form frmChar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Special Characters"
   ClientHeight    =   3390
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.ListBox LV2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.ListBox LV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Available Characters:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Currently Selected:"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Menu menu 
      Caption         =   "&Format"
      Begin VB.Menu Character 
         Caption         =   "&Show As Character"
      End
      Begin VB.Menu HTMLCode 
         Caption         =   "&Show as HTML Code"
      End
   End
End
Attribute VB_Name = "frmChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Character_Click()
On Error Resume Next
Dim intFileNum As Integer
intFileNum = FreeFile
 LV.Clear
Open App.Path & "\temp\specialcharvis.txt" For Input As #intFileNum
 Do While Not EOF(intFileNum)
  Line Input #intFileNum, Value
  LV.AddItem Value
 Loop
End Sub

Private Sub Command1_Click()
On Error Resume Next
All = LV2.ListCount - 1
For i = 1 To All
 fMainForm.ActiveForm.rtfText.SelText = LV2.List(i) & " "
Next i
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo Error:
LV2.RemoveItem LV2.ListIndex

Exit Sub
Error:
If Err.Number = 5 Then
MsgBox "No item to remove.", vbInformation
End If
Resume Next
End Sub

Private Sub Command3_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim intFileNum As Integer
intFileNum = FreeFile

Open App.Path & "\temp\specialcharvis.txt" For Input As #intFileNum
 Do While Not EOF(intFileNum)
  Line Input #intFileNum, Value
  LV.AddItem Value
 Loop
End Sub

Private Sub HTMLCode_Click()
On Error Resume Next
Dim intFileNum As Integer
intFileNum = FreeFile
 LV.Clear
Open App.Path & "\temp\specialchar.txt" For Input As #intFileNum
 Do While Not EOF(intFileNum)
  Line Input #intFileNum, Value
  LV.AddItem Value
 Loop
End Sub

Private Sub LV_Click()
On Error Resume Next
 LV2.AddItem LV.Text
End Sub



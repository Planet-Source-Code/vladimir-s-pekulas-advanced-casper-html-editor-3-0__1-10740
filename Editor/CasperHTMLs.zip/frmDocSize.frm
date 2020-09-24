VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDocSize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Document Weight"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView lstLinks 
      Height          =   1455
      Left            =   240
      TabIndex        =   18
      Top             =   4560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "IMG"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Other Dependecies"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   6120
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   3
      Left            =   240
      Picture         =   "frmDocSize.frx":0000
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   9
      Top             =   2520
      Width           =   360
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   2
      Left            =   240
      Picture         =   "frmDocSize.frx":018A
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   360
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   240
      Picture         =   "frmDocSize.frx":0314
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   360
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   240
      Picture         =   "frmDocSize.frx":049E
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   360
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   240
      Picture         =   "frmDocSize.frx":0628
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   360
      Width           =   240
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   3120
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDocSize.frx":072A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Other content then code such as  images is  not included in document weight,"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Download time is calculated using favorable conditions.  This number can change depending on the Web  traffic."
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label Modem28 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 11.34 KB"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1605
      Width           =   1215
   End
   Begin VB.Label Modem56 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 11.34 KB"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   2085
      Width           =   1215
   End
   Begin VB.Label ModemLAN 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 11.34 KB"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   2565
      Width           =   1215
   End
   Begin VB.Label Modem14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 11.34 KB"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1125
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "with a cable modem:"
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   10
      Top             =   2565
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "with a 28.8 modem:"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   8
      Top             =   1605
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "with a 56.6 modem:"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   2085
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "with a 14.4 modem:"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1125
      Width           =   1455
   End
   Begin VB.Label DocSize 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 11.34 KB"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   353
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Document weight:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   353
      Width           =   2175
   End
End
Attribute VB_Name = "frmDocSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROBLEMS:
'
' When calculating download time from a file size bellow 5 kb (+/-)
' the division gives you something like 9.90021324563210321E-02
' which I need to convert to 0.090021324563210321.
' I have written a small if then statement, but it doesn't work as
' I'd like and since I'm pretty sure VB has some in-build conversion for
' it, I won't bother with it. Anyway if you know how to do that send me
' an email at vpekulas@home.com

Dim DocumentSize As Double

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Error:
'Get other content
Dim arrLinks() As String, lngFile As Long, strHTML As String
Dim lngLink As Long
lngFile& = FreeFile
  
 strHTML$ = fMainForm.ActiveForm.rtfText.Text
 'Open txtFile.Text For Input As lngFile&
  'strHTML$ = Input(LOF(lngFile&), lngFile&)
 'Close lngFile&
  Call LinksFromHTML(arrLinks(), strHTML$)
 For lngLink& = 1 To UBound(arrLinks)
  lstLinks.ListItems.Add , , (arrLinks(lngLink&)), , 1
 Next lngLink&
 
'Calculate Download Time
 DocumentSize = FileLen(App.Path & "\Casper~temp.html") / 1024 'Windows use 1000 bytes for a each KB not 1024, but let's stick with 1024.
 Modem14Func
 Modem28Func
 Modem56Func
 ModemLANFunc
  If DocumentSize < 10 Then DocSize.Caption = Mid(DocumentSize, 1, 4) & " KB"
  If DocumentSize > 10 Then DocSize.Caption = Mid(DocumentSize, 1, 5) & " KB"
Exit Sub
Error:
 MsgBox Err.Description, vbCritical, "Error"
End Sub

Function Modem14Func()
On Error GoTo Error:
Dim TMPTime As String
 TMPTime = DocumentSize / 1.612
 
If InStr(1, TMPTime, "E") > 0 Then
 Pos = Mid(TMPTime, Len(TMPTime) - 1, 3)
 PosPos = Mid(TMPTime, Len(TMPTime) - 1, 3)
If Pos = "01" Then Pos = "0."
If Pos = "02" Then Pos = "0.0"
If Pos = "03" Then Pos = "0.00"
If Pos = "04" Then Pos = "0.000"
If Pos = "05" Then Pos = "0.0000"
 TMPTime = Pos & Mid(TMPTime, 1, PosPos - 1) & Mid(TMPTime, PosPos + 1, Len(TMPTime) - PosPos)
End If
 
  If TMPTime < 10 Then TMPTime = Mid(TMPTime, 1, 4)
  If TMPTime > 10 Then TMPTime = Mid(TMPTime, 1, 5)
 Modem14.Caption = TMPTime & " Sec."
 Exit Function
Error:
 Modem14.Caption = "0.00 Sec."
End Function

Function Modem28Func()
On Error GoTo Error:
Dim TMPTime As String
 TMPTime = DocumentSize / 3.225
If InStr(1, TMPTime, "E") > 0 Then
 Pos = Mid(TMPTime, Len(TMPTime) - 1, 3)
 PosPos = Mid(TMPTime, Len(TMPTime) - 1, 3)
If Pos = "01" Then Pos = "0."
If Pos = "02" Then Pos = "0.0"
If Pos = "03" Then Pos = "0.00"
If Pos = "04" Then Pos = "0.000"
If Pos = "05" Then Pos = "0.0000"
 TMPTime = Pos & Mid(TMPTime, 1, PosPos - 1) & Mid(TMPTime, PosPos + 1, Len(TMPTime) - PosPos)
End If

  If TMPTime < 10 Then TMPTime = Mid(TMPTime, 1, 4)
  If TMPTime > 10 Then TMPTime = Mid(TMPTime, 1, 5)
 Modem28.Caption = TMPTime & " Sec."
 Exit Function
Error:
 Modem28.Caption = "0.00 Sec."
End Function

Function Modem56Func()
On Error GoTo Error:
Dim TMPTime As String
 TMPTime = DocumentSize / 6.25
If InStr(1, TMPTime, "E") > 0 Then
 Pos = Mid(TMPTime, Len(TMPTime) - 1, 3)
 PosPos = Mid(TMPTime, Len(TMPTime) - 1, 3)
If Pos = "01" Then Pos = "0."
If Pos = "02" Then Pos = "0.0"
If Pos = "03" Then Pos = "0.00"
If Pos = "04" Then Pos = "0.000"
If Pos = "05" Then Pos = "0.0000"
 TMPTime = Pos & Mid(TMPTime, 1, PosPos - 1) & Mid(TMPTime, PosPos + 1, Len(TMPTime) - PosPos)
End If
  If TMPTime < 10 Then TMPTime = Mid(TMPTime, 1, 4)
  If TMPTime > 10 Then TMPTime = Mid(TMPTime, 1, 5)
 Modem56.Caption = TMPTime & " Sec."
 Exit Function
Error:
 Modem56.Caption = "0.00 Sec."
End Function

Function ModemLANFunc()
On Error GoTo Error:
Dim TMPTime As String
 TMPTime = DocumentSize / 24.68
If InStr(1, TMPTime, "E") > 0 Then
 Pos = Mid(TMPTime, Len(TMPTime) - 1, 3)
 PosPos = Mid(TMPTime, Len(TMPTime) - 1, 3)
If Pos = "01" Then Pos = "0."
If Pos = "02" Then Pos = "0.0"
If Pos = "03" Then Pos = "0.00"
If Pos = "04" Then Pos = "0.000"
If Pos = "05" Then Pos = "0.0000"
 
 
 TMPTime = Pos & Mid(TMPTime, 1, PosPos - 1) & Mid(TMPTime, PosPos + 1, Len(TMPTime) - PosPos)
End If
  If TMPTime < 10 Then TMPTime = Mid(TMPTime, 1, 4)
  If TMPTime > 10 Then TMPTime = Mid(TMPTime, 1, 5)
 ModemLAN.Caption = TMPTime & " Sec."
 Exit Function
Error:
 ModemLAN.Caption = "0.00 Sec."
End Function




'Other Content that might affect download time

Public Function URLFromTag(strTag As String) As String
On Error GoTo Error:
Dim lngChar As Long, strChar As String, blnFlag As Boolean, strBuffer As String

      For lngChar& = 1 To Len(strTag$)
         strChar$ = Mid$(strTag$, lngChar&, 1)
      
         If blnFlag = False And LCase$(Mid$(strTag$, lngChar&, 3)) = "src" Then
            blnFlag = True
            lngChar& = lngChar& + 3
            If Mid$(strTag$, lngChar& + 1, 1) = """" Then lngChar& = lngChar& + 1
         
         ElseIf blnFlag = True Then
            If strChar$ = """" Or strChar$ = ">" Or strChar$ = " " Then Exit For
            strBuffer$ = strBuffer$ & strChar$
         End If
         
      Next lngChar&
   URLFromTag = strBuffer$
Exit Function
Error:
 MsgBox Err.Description, vbCritical, "Error"
End Function

Public Function LinksFromHTML(arrLinks() As String, strHTML As String) As Long
On Error GoTo Error:
Dim lngSpot As Long, strCurTag As String, strURL As String
ReDim arrLinks(0) As String
   
      Do
         lngSpot& = InStr(lngSpot& + 1, LCase$(strHTML$), "<img")
         If lngSpot& = 0 Then Exit Do
          strCurTag$ = Mid$(strHTML$, lngSpot&, InStr(lngSpot&, strHTML$, ">") - lngSpot&)
          strURL$ = URLFromTag(strCurTag$)
             If Len(strURL$) > 0 Then
              ReDim Preserve arrLinks(UBound(arrLinks) + 1) As String
              arrLinks(UBound(arrLinks)) = strURL$
            End If
      Loop
  LinksFromHTML = UBound(arrLinks) + 1
Exit Function
Error:
 MsgBox Err.Description, vbCritical, "Error"
End Function


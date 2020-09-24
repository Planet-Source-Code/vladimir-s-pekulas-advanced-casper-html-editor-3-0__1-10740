VERSION 5.00
Begin VB.Form frmChar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Special Characters"
   ClientHeight    =   3810
   ClientLeft      =   150
   ClientTop       =   105
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Txtcopy 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      HideSelection   =   0   'False
      Left            =   240
      TabIndex        =   7
      Top             =   3255
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   960
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   240
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   6
      Top             =   480
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Special Character  Set:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
        ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As _
        Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hdc As Long, ByVal hObject As _
        Long)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function MoveToEx& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
        ByVal Y As Long, lpPoint As POINTAPI)
Private Declare Function CreateRectRgnIndirect& Lib "gdi32" (lprect As RECT)
Private Declare Function CreateRectRgn& Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As _
        Long, ByVal X2 As Long, ByVal Y2 As Long)
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Private Declare Function LineTo& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal _
        Y As Long)
Private Declare Function Rectangle& Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Private Const HORZRES = 8
Private Const VERTRES = 10
Const SRCCOPY = &HCC0020

Dim asciiList()
Dim sizeX, sizeY, previousX, previousY
Dim mouseDown As Boolean, mouseVisible As Boolean

Private Sub CboFonts_Click()

    drawSquare "Arial"
    Picture2.Font = "Arial"
    Picture2.FontSize = 18
    drawfocusColour previousX, previousY
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Txtcopy.Text, vbCFText
    Picture1.SetFocus
End Sub

Private Sub cmdSelect_Click()
    inserttext
    Picture1.SetFocus
End Sub
Sub inserttext()
    Dim X1, Y1, char$, lprect As RECT, offsetx, offsety, s
    s = selectedsquare
    Y1 = s \ 32
    X1 = s Mod 32
    char$ = Chr$((Y1 * 32) + (X1 + 1) + 30) '1)
    Txtcopy.SelText = char$
End Sub


Private Sub Command1_Click()
   fMainForm.ActiveForm.rtfText.SelText = Txtcopy.Text
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc("A") And (Shift And vbAltMask) Then
        MsgBox "alt+ A=frm key"
        Txtcopy.SelStart = 0
        Txtcopy.SelLength = Len(Txtcopy.Text)
        Txtcopy.SetFocus
    End If
    If KeyCode = Asc("F") And (Shift And vbAltMask) Then

    End If
    If KeyCode = Asc("S") And (Shift And vbAltMask) Then
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc("A") And (Shift And vbAltMask) Then
        Txtcopy.SelStart = 0
        Txtcopy.SelLength = Len(Txtcopy.Text)
        Txtcopy.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim X, Y
    Dim index
    sizeX = (Picture1.ScaleWidth \ 32)
    sizeY = (Picture1.ScaleHeight \ 7)
    index = 32
    drawSquare "Arial"
    Picture2.Visible = False
    Picture3.Visible = False
    mouseDown = False
    Picture1_MouseDown 0&, 0&, 0, 0
    Picture1_MouseUp 0&, 0&, 0, 0
    frmChar.Show
    Picture1.SetFocus
    cmdCopy.Enabled = False
    selectedsquare = 1
End Sub


Private Sub Picture1_DblClick()
    inserttext
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown
        If selectedsquare + 32 < 225 Then
            selectedsquare = selectedsquare + 32
        End If
    Case vbKeyUp
        If selectedsquare - 32 > 0 Then
            selectedsquare = selectedsquare - 32
        End If
    Case vbKeyRight
        If selectedsquare + 1 < 225 Then
            selectedsquare = selectedsquare + 1
        End If
    Case vbKeyLeft
        If selectedsquare - 1 > 0 Then
            selectedsquare = selectedsquare - 1
        End If
    Case Else
        Exit Sub
    End Select
    drawselected (selectedsquare - 1)
    updateLabel (selectedsquare - 1) Mod 32, (selectedsquare - 1) \ 32

End Sub

Sub drawselected(s As Long)
    Dim X1, Y1, char$, lprect As RECT, offsetx, offsety
    Y1 = s \ 32
    X1 = s Mod 32
    Picture1.Line (previousX * sizeX + 1, previousY * sizeY + 1)-(previousX * sizeX + (sizeX - 1), previousY * sizeY + (sizeY - 1)), vbWhite, BF
    Picture1.CurrentX = (previousX * sizeX) + 3
    Picture1.CurrentY = (previousY * sizeY)
    Picture1.Print Chr$((previousY * 32) + (previousX + 1) + 31);
    previousX = X1
    previousY = Y1
    char$ = Chr$((Y1 * 32) + (X1 + 1) + 31)
    Picture2.Visible = False: Picture3.Visible = False
    offsetx = (Picture2.ScaleWidth - Picture2.TextWidth(char$)) \ 2
    offsety = (Picture2.ScaleHeight - Picture2.TextHeight(char$)) \ 2
    Picture2.left = (X1 * sizeX - 5) + 10
    Picture2.top = (Y1 * sizeY - 5) + 35
    Picture3.left = Picture2.left + 5
    Picture3.top = Picture2.top + 5
    Picture2.CurrentX = offsetx
    Picture2.CurrentY = offsety '    Chr$((y1 * 32) + (x1 + 1) + 31)
    Picture2.Picture = LoadPicture()
    Picture2.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
    Picture2.Visible = True: Picture3.Visible = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim X1, Y1, ret, lprect As RECT, offsetx, offsety, char$
    X1 = X \ sizeX
    Y1 = Y \ sizeY
    If Button = vbRightButton Then
        Exit Sub
    End If
    If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then

        If Not (IsEmpty(previousX) And IsEmpty(previousY)) Then
            lprect.left = X1 * sizeX + 1
            lprect.top = Y1 * sizeY + 1
            lprect.right = X1 * sizeX + (sizeX - 1) + 1 '- 1
            lprect.bottom = Y1 * sizeY + (sizeY - 1) + 1
            Picture1.Line (previousX * sizeX, previousY * sizeY)-(previousX * sizeX + (sizeX), previousY * sizeY + (sizeY)), vbBlack, BF
            Picture1.Line (previousX * sizeX + 1, previousY * sizeY + 1)-(previousX * sizeX + (sizeX - 1), previousY * sizeY + (sizeY - 1)), vbWhite, BF
            char$ = Chr$((previousY * 32) + (previousX + 1) + 31)
            offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
            offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
            Picture1.CurrentX = (previousX * sizeX) + offsetx
            Picture1.CurrentY = (previousY * sizeY) + offsety
            Picture1.Print char$;
        End If
        Picture2.Visible = False
        Picture3.Visible = False
        Picture2.left = (X1 * sizeX - 5) + 10
        Picture2.top = (Y1 * sizeY - 5) + 35
        Picture3.left = Picture2.left + 5
        Picture3.top = Picture2.top + 5
        Picture2.Visible = True
        Picture3.Visible = True
        selectedsquare = (Y1 * 32) + (X1 + 1)
        previousX = X1
        previousY = Y1
    End If
  
    If mouseDown = False Then
        ret = ShowCursor(False)
        While ret >= 0
            ret = ShowCursor(False)
        Wend

        mouseVisible = False
    End If
    mouseVisible = False
    Picture2.Visible = True
    Picture3.Visible = True
    mouseDown = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim X1, Y1, ret, char$, Key
    Dim offsetx, offsety
    Static lastx
    Static lasty
    If mouseDown = True Then
        X1 = X \ sizeX
        Y1 = Y \ sizeY
        If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then
            If mouseVisible = True Then
                makeCursorInvisible
            End If
            If lastx = X1 And lasty = Y1 Then Exit Sub
            lastx = X1: lasty = Y1
            Key = (Y1 * 32) + (X1 + 1)
            Picture2.Visible = False
            Picture3.Visible = False
            Picture2.left = (X1 * sizeX - 5) + 10
            Picture2.top = (Y1 * sizeY - 5) + 35
            Picture3.left = Picture2.left + 5
            Picture3.top = Picture2.top + 5
            char$ = Chr$((Y1 * 32) + (X1 + 1) + 31)
            If Picture2.Tag = char$ Then
            Else
                previousX = X1
                previousY = Y1
                Picture2.Tag = char$
                offsetx = (Picture2.ScaleWidth - Picture2.TextWidth(char$)) \ 2
                offsety = (Picture2.ScaleHeight - Picture2.TextHeight(char$)) \ 2
                Picture2.CurrentX = offsetx
                Picture2.CurrentY = offsety
                Picture2.Picture = LoadPicture()
                Picture2.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
                Picture2.Visible = True
                Picture3.Visible = True
            End If
            previousX = X1
            previousY = Y1
        Else
            makeCursorVisible
            Exit Sub
        End If
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ret, X1, Y1, lprect As RECT
    X1 = X \ sizeX
    Y1 = Y \ sizeY
    If mouseVisible = False Then
        ret = ShowCursor(True)
        While ret < 0
            ret = ShowCursor(True)
        Wend
        mouseVisible = True
    End If
    drawfocusColour previousX, previousY
    If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then

    Else
        If mouseDown = True Then
            Picture2.Visible = False
            Picture3.Visible = False
            drawfocusColour previousX, previousY
        End If

    End If
    Picture2.Visible = False
    Picture3.Visible = False
    mouseDown = False
End Sub

Sub makeCursorInvisible()
    Dim ret
    ret = ShowCursor(False)
    While ret >= 0
        ret = ShowCursor(False)
    Wend
    mouseVisible = False
End Sub

Sub makeCursorVisible()
    Dim ret
    ret = ShowCursor(True)
    While ret < 0
        ret = ShowCursor(True)
    Wend
    mouseVisible = True
End Sub

Sub drawSquare(f As String)
    Dim X As Long, Y As Long, char$, lpPT As POINTAPI
    Dim offsetx, offsety
    Picture1.Visible = False
    Picture1.FontName = f
    Picture1.FontSize = 8
    Picture1.Picture = LoadPicture()
    For X = 0 To 31
        For Y = 0 To 6
            char$ = Chr$((Y * 32) + (X + 1) + 31)
            offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
            offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
            Picture1.CurrentX = (X * sizeX) + offsetx
            Picture1.CurrentY = (Y * sizeY) + offsety
            Picture1.Print char$;

        Next Y
    Next X
    For X = 0 To 7
        MoveToEx Picture1.hdc, 0, X * sizeY, lpPT
        LineTo Picture1.hdc, sizeX * 32, X * sizeY
    Next X
    For X = 0 To 32
        MoveToEx Picture1.hdc, X * sizeX, 0, lpPT
        LineTo Picture1.hdc, X * sizeX, sizeY * 7 + 1
    Next X
    Picture1.Visible = True
End Sub

Private Sub Txtcopy_Change()
    If Txtcopy.Text = "" Then
        cmdCopy.Enabled = False
    Else

        cmdCopy.Enabled = True
    End If
End Sub

Sub drawfocusColour(X, Y)
    Dim lprect As RECT, offsetx, offsety, char$
    Picture1.Line (X * sizeX + 1, Y * sizeY + 1)-(X * sizeX + (sizeX - 1), _
            Y * sizeY + (sizeY - 1)), vbHighlight, BF
    char$ = Chr$((Y * 32) + (X + 1) + 31)
    offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
    offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
    Picture1.CurrentX = (X * sizeX) + offsetx
    Picture1.CurrentY = (Y * sizeY) + offsety
    Picture1.ForeColor = vbWhite
    Picture1.Print char$;
    Picture1.ForeColor = vbBlack
    lprect.left = X * sizeX + 1
    lprect.top = Y * sizeY + 1
    lprect.right = X * sizeX + (sizeX - 1) + 1 '- 1
    lprect.bottom = Y * sizeY + (sizeY - 1) + 1  '- 1
    DrawFocusRect Picture1.hdc, lprect
End Sub


Private Sub Txtcopy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc("A") And (Shift And vbAltMask) Then
        MsgBox "tkdwn"
        Txtcopy.SelStart = 0
        Txtcopy.SelLength = Len(Txtcopy.Text)
    End If
End Sub




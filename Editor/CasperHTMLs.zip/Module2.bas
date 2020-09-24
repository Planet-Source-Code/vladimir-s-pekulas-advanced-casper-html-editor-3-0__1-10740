Attribute VB_Name = "Module2"
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Public Const MIIM_ID = &H2
    Public Const MIIM_TYPE = &H10
    Public Const MFT_STRING = &H0&


'@@@@

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'Public Const HTBORDER = 18 'Not used in this example
'Public Const HTBOTTOM = 15 'Not used in this example
'Public Const HTBOTTOMLEFT = 16 'Not used in this example
'Public Const HTBOTTOMRIGHT = 17 'Not used in this example

'// Constants
Public Const HTTOP = 12
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CYSMCAPTION = 51
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const EM_GETLINECOUNT = &HBA        '// Total Line Count
Public Const EM_GETFIRSTVISIBLELINE = &HCE '// First Visible Line





Public Sub CarrotStatus()
   Dim lLine As Long, lCol As Long
   Dim cCol As Long, lChar As Long, I As Long

   lChar = fMainForm.ActiveForm.rtfText.SelStart + 1

   ' Get the line number
   lLine = 1 + SendMessageLong(fMainForm.ActiveForm.rtfText.hWnd, EM_LINEFROMCHAR, _
           fMainForm.ActiveForm.rtfText.SelStart, 0&)

   ' Get the Character Position
   cCol = SendMessageLong(fMainForm.ActiveForm.rtfText.hWnd, EM_LINELENGTH, lChar - 1, 0&)

   I = SendMessageLong(fMainForm.ActiveForm.rtfText.hWnd, EM_LINEINDEX, lLine - 1, 0&)
   lCol = lChar - I

   ' Caption of Label1 is set to Cursor Position.
   ' This could also be a panel in a StatusBar.
   fMainForm.sbStatusBar.Panels(2).Text = "Line: " & lLine & ", Character: " & lCol

End Sub



Public Sub CarrotStatus2()
   Dim lLine As Long, lCol As Long
   Dim cCol As Long, lChar As Long, I As Long

   lChar = fMainForm.ActiveForm.rtfText2.SelStart + 1

   ' Get the line number
   lLine = 1 + SendMessageLong(fMainForm.ActiveForm.rtfText2.hWnd, EM_LINEFROMCHAR, _
           fMainForm.ActiveForm.rtfText2.SelStart, 0&)

   ' Get the Character Position
   cCol = SendMessageLong(fMainForm.ActiveForm.rtfText2.hWnd, EM_LINELENGTH, lChar - 1, 0&)

   I = SendMessageLong(fMainForm.ActiveForm.rtfText2.hWnd, EM_LINEINDEX, lLine - 1, 0&)
   lCol = lChar - I

   ' Caption of Label1 is set to Cursor Position.
   ' This could also be a panel in a StatusBar.
   fMainForm.sbStatusBar.Panels(2).Text = "Line: " & lLine & ", Character: " & lCol

End Sub


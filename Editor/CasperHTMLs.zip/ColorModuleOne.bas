Attribute VB_Name = "ColorModuleOne"
Option Explicit
' Public Constants

'Public Const EM_GETLINECOUNT = &HBA        '// Total Line Count
'Public Const EM_GETFIRSTVISIBLELINE = &HCE '// First Visible Line
'Public Const WM_VSCROLL = &H115            '// Vertical Scrolling



Public Const WM_USER = &H400
Public Const MAX_PATH = 260
Public Const EM_GETLINECOUNT = &HBA
'Public Const EM_LINELENGTH = &HC1
'Public Const EM_LINEINDEX = &HBB

'Public Const EM_HIDESELECTION = WM_USER + 63

Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_COPYDATA = &H4A
Public Const WM_SYSCOMMAND = &H112
Public Const WM_SETREDRAW = &HB

Public Const GWL_STYLE = (-16)

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const SW_SHOW = 5
Public Const SW_HIDE = 0


Public Enum ColourStatus
    InTag = 0
    OutTag = 1
    inComment = 2
    OutComment = 3
    inscript = 4
End Enum


Public Enum ModeConstants
    vbwHTML = 1
End Enum






Public Type COPYDATASTRUCT
   dwData As Long
   cbData As Long
   lpData As Long
End Type

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Public Type CharRange
  cpMin As Long     '// First character of range (0 for start of doc)
  cpMax As Long     '// Last character of range (-1 for end of doc)
End Type



Public Type CMDLG_VALUES
    FileName As String
    FileTitle As String
    FilterIndex As Long
End Type


Public blnStartUp       As Boolean
Public blnOpening       As Boolean
Public blnBusy          As Boolean
Public blnClosing       As Boolean
Public DocsChanged      As Boolean


Public CmDlg            As CMDLG_VALUES
Public sVBPath          As String
Private m_bInDevelopment As Boolean
Public blnDoingMultiple As Boolean


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const EM_CANUNDO = &HC6
Public Const EM_CANPASTE = (WM_USER + 50)
Public Const WM_VSCROLL = &H115

'// formatting constants
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_GETCHARFORMAT = (WM_USER + 58)
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const SCF_SELECTION = &H1&
Public Const LF_FACESIZE = 32
Public Const CFM_BACKCOLOR = &H4000000

'// get pos constants
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_POSFROMCHAR = &HD6&
Public Const EM_CHARFROMPOS = &HD7&


'// Text Modes
Public Enum TextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2 ' /* default behavior */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8 ' /* default behavior */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32 ' /* default behavior */
End Enum

Sub ErrHandler(Optional lngErrNum As Long = 0, Optional strErrorText As String = "", Optional strSource As String = "<Unknown>", Optional blnMustExit As Boolean = False, Optional strExtra As String = Empty, Optional bNoError As Boolean = False)
 MsgBox "Err " & lngErrNum & ": " & strErrorText & vbCrLf
End Sub



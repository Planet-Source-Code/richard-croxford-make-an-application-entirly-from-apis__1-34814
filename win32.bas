Attribute VB_Name = "win32"
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function ValidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long




Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const CS_CLASSDC = &H40
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_SHOW = 5
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Public Const CS_OWNDC = &H20
Public Const CS_HREDRAW = &H2
Public Const CS_VREDRAW = &H1
Public Const COLOR_APPWORKSPACE = 12

Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_SIZEBOX = WS_THICKFRAME

Public Const WS_TABSTOP = &H10000

Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000


' Window Messages
Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Const WM_MOVE = &H3
Const WM_SIZE = &H5

Const WM_ACTIVATE = &H6
'
'  WM_ACTIVATE state values

Const WA_INACTIVE = 0
Const WA_ACTIVE = 1
Const WA_CLICKACTIVE = 2

Const WM_SETFOCUS = &H7
Const WM_KILLFOCUS = &H8
Const WM_ENABLE = &HA
Const WM_SETREDRAW = &HB
Const WM_SETTEXT = &HC
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Const WM_CLOSE = &H10
Const WM_QUERYENDSESSION = &H11
Const WM_QUIT = &H12
Const WM_QUERYOPEN = &H13
Const WM_ERASEBKGND = &H14
Const WM_SYSCOLORCHANGE = &H15
Const WM_ENDSESSION = &H16
Const WM_SHOWWINDOW = &H18
Const WM_WININICHANGE = &H1A
Const WM_DEVMODECHANGE = &H1B
Const WM_ACTIVATEAPP = &H1C
Const WM_FONTCHANGE = &H1D
Const WM_TIMECHANGE = &H1E
Const WM_CANCELMODE = &H1F
Const WM_SETCURSOR = &H20
Const WM_MOUSEACTIVATE = &H21
Const WM_CHILDACTIVATE = &H22
Const WM_QUEUESYNC = &H23

Const WM_GETMINMAXINFO = &H24

Public Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type


Public Type WNDCLASS
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
'    pt As POINTAPI
End Type


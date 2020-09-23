Attribute VB_Name = "mmain"
Dim hwnd As Long



Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case uMsg
    Case WM_DESTROY
        PostQuitMessage (0)
    Case WM_PAINT
        Dim hdc As Long
        Dim pt As PAINTSTRUCT
        Dim rc As RECT
        
        hdc = BeginPaint(hwnd, pt)
        
        GetClientRect hwnd, rc
        rc.Top = rc.Bottom / 2
        DrawText hdc, "This is a window, without VB forms!!!", 37, rc, 1
        
        EndPaint hwnd, pt
    Case Else
        WindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
  End Select
End Function


Sub Main()
    
    Dim classinfo As WNDCLASSEX
    
    classinfo.style = 0
    classinfo.lpfnWndProc = FarProc(AddressOf WindowProc)
    classinfo.cbClsExtra = 0
    classinfo.cbWndExtra = 0
    classinfo.hInstance = App.hInstance
    classinfo.hbrBackground = COLOR_APPWORKSPACE
    classinfo.lpszMenuName = ""
    classinfo.lpszClassName = "NoForms"
    classinfo.cbSize = Len(classinfo)

      
    
     RegisterClassEx classinfo
    
    
    
    hwnd = CreateWindowEx(WS_EX_DLGMODALFRAME, "NoForms", "Look No Forms!", WS_OVERLAPPEDWINDOW, 300, 300, 300, 300, GetDesktopWindow(), 0, App.hInstance, 0)
        
    
    ShowWindow hwnd, SW_SHOW
    UpdateWindow hwnd
    
    Dim msg As msg
    Do While GetMessage(msg, 0, 0, 0)
        TranslateMessage msg
        DispatchMessage msg
    Loop

    UnregisterClass "NoForms", App.hInstance
    
End Sub


Public Function FarProc(lpProcName As Long) As Long
    FarProc = lpProcName
End Function


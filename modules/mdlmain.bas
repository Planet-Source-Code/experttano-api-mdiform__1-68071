Attribute VB_Name = "mdlmain"
Option Explicit

Private Const MF_POPUP = &H10&
Private Const MF_STRING = &H0&
Private Const GWL_WNDPROC = (-4)

Private Const WM_COMMAND = &H111
Private Const WM_DESTROY As Long = &H2
Private Const WM_PAINT As Long = &HF&
Private Const WM_CREATE As Long = &H1
Private Const WM_NULL As Long = &H0
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCPAINT As Long = &H85

Private Const CW_USEDEFAULT As Long = &H80000000

Private Const WS_EX_TOPMOST = &H8
Private Const WS_EX_CLIENTEDGE = &H200&

Private Const WS_HSCROLL = &H100000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_OVERLAPPED As Long = &H0&
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_OVERLAPPEDWINDOW As Long = _
  (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or _
    WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const SW_NORMAL As Long = 1
Private Const SW_SHOW As Long = 5

Private Const CS_HREDRAW As Long = &H2
Private Const CS_VREDRAW As Long = &H1
Private Const COLOR_WINDOW As Long = 5

Private Const IDC_ARROW As Long = 32512&

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type PAINTSTRUCT
  hdc As Long
  fErase As Long
  rcPaint As RECT
  fRestore As Long
  fIncUpdate As Long
  rgbReserved As Byte
End Type

Private Type WNDCLASSEX
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

Private Type Msg
  hWnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type

Private Type CLIENTCREATESTRUCT
  hWindowMenu As Long
  idFirstChild As Long
End Type

Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function UnregisterClass Lib "user32.dll" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long

Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, _
    ByVal lpWindowName As String, ByVal dwStyle As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hWndParent As Long, _
    ByVal hMenu As Long, ByVal hInstance As Long, _
    ByRef lpParam As Any) As Long
    
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function BeginPaint Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CreateMDIWindow Lib "user32.dll" Alias "CreateMDIWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String, _
    ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, _
    ByVal hInstance As Long, ByVal lParam As Long) As Long

Private Declare Function DefFrameProc Lib "user32" Alias "DefFrameProcA" _
  (ByVal hWnd As Long, ByVal hWndMDIClient As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" _
  (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long


Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal handle_of_window As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal handle_of_window As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal handle_of_window As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal handle_of_window As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function GetClientRect Lib "user32" (ByVal handle_of_window As Long, lpRect As RECT) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
                                                    
'global variables
Public hWnd As Long, MDIhwnd As Long
Public hMenu1 As Long, hMenu2 As Long

Const ID_EXIT = 100
Const ID_NEW_WND = 101
Const APP_NAME = "MDIMain"
Const CHILD_NAME = "MDI Child"

Public Sub Main()
    WinMDIMain
End Sub

'subclassing
Public Function MDIChildHandler(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim ps As PAINTSTRUCT
Dim hdc As Long

  Select Case Msg
    Case WM_PAINT
      hdc = BeginPaint(hWnd, ps)
      
      EndPaint hWnd, ps
    Case Else
      
      MDIChildHandler = DefMDIChildProc(hWnd, Msg, wParam, lParam)
   End Select
End Function

'subclassing Child
Public Function MDIClientHandler(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim ps As PAINTSTRUCT
  Dim chHwnd As Long
  Dim cs As CLIENTCREATESTRUCT

  Select Case Msg
    Case WM_PAINT
      BeginPaint hWnd, ps
      EndPaint hWnd, ps
      
    Case WM_DESTROY
      DestroyWindow hWnd
      UnregisterClass CHILD_NAME, App.hInstance
      UnregisterClass APP_NAME, App.hInstance
      
    Case WM_CREATE
      'Client Frame MDI form
      cs.idFirstChild = 0
      
      MDIhwnd = CreateWindowEx(WS_EX_TOPMOST And WS_EX_CLIENTEDGE, "MDIClient", _
                       "Main", WS_CHILD Or WS_CLIPCHILDREN Or WS_VISIBLE, _
                       0, 0, 0, 0, hWnd, 0&, _
                       App.hInstance, ByVal (VarPtr(cs)))
      
    Case WM_COMMAND
      Select Case wParam
        Case ID_EXIT
          DestroyWindow hWnd
          UnregisterClass CHILD_NAME, App.hInstance
          UnregisterClass APP_NAME, App.hInstance
          
        Case ID_NEW_WND
          'Create Child Window
          chHwnd = CreateMDIWindow(CHILD_NAME, _
                       "MDI Client", WS_CHILD Or WS_CLIPCHILDREN Or WS_VISIBLE, _
                       CW_USEDEFAULT, CW_USEDEFAULT, _
                       CW_USEDEFAULT, CW_USEDEFAULT, MDIhwnd, _
                       App.hInstance, 0&)

          If chHwnd Then
            ShowWindow chHwnd, SW_SHOW And SW_NORMAL
            UpdateWindow chHwnd
            SetFocus hWnd
          Else
            MsgBox "Child window not crerated!"
            Exit Function
          End If
      End Select
    Case Else
      MDIClientHandler = DefFrameProc(hWnd, MDIhwnd, Msg, wParam, lParam)
   End Select
End Function

Public Function WinMDIMain() As Long
    Dim wndcls As WNDCLASSEX
    Dim message As Msg
    
    
    wndcls.cbSize = Len(wndcls)
    wndcls.style = CS_HREDRAW + CS_VREDRAW

    wndcls.lpfnWndProc = GetFuncPtr(AddressOf MDIClientHandler)
    wndcls.cbClsExtra = 0
    wndcls.cbWndExtra = 0
    wndcls.hInstance = App.hInstance
    wndcls.hIcon = 0
    wndcls.hCursor = LoadCursor(0, IDC_ARROW)
    wndcls.hbrBackground = COLOR_WINDOW
    wndcls.lpszMenuName = 0
    wndcls.lpszClassName = APP_NAME
    
    If RegisterClassEx(wndcls) = 0 Then
      MsgBox "Can't Register Class"
      Exit Function
    End If

    wndcls.style = 0
    wndcls.lpfnWndProc = GetFuncPtr(AddressOf MDIChildHandler)
    wndcls.cbClsExtra = 0
    wndcls.cbWndExtra = 0
    wndcls.hInstance = App.hInstance
    wndcls.hIcon = 0
    wndcls.hCursor = LoadCursor(0, IDC_ARROW)
    wndcls.hbrBackground = COLOR_WINDOW
    wndcls.lpszMenuName = 0
    wndcls.lpszClassName = CHILD_NAME
    
    If RegisterClassEx(wndcls) = 0 Then
      MsgBox "Can't Register Class"
      Exit Function
    End If

    'Main form
    hWnd = CreateWindowEx(0&, APP_NAME, _
                       App.ProductName, WS_OVERLAPPEDWINDOW, _
                       CW_USEDEFAULT, 0, _
                       CW_USEDEFAULT, 0, 0&, 0&, _
                       App.hInstance, 0&)
    'Menu
    hMenu2 = CreateMenu()
    AppendMenu hMenu2, MF_STRING, ID_NEW_WND, "&New windows"
    AppendMenu hMenu2, MF_STRING, ID_EXIT, "&Exit"

    hMenu1 = CreateMenu()
    AppendMenu hMenu1, MF_POPUP, hMenu2, "&File"

    SetMenu hWnd, hMenu1
    
    SetWindowLong hWnd, GWL_WNDPROC, AddressOf MDIClientHandler

    ShowWindow hWnd, SW_NORMAL
    UpdateWindow hWnd
    
    SetFocus hWnd
    
    While (GetMessage(message, hWnd, 0, 0)) And message.message <> WM_NULL
        TranslateMessage message
        DispatchMessage message
        DoEvents
    Wend
        
    WinMDIMain = message.wParam
End Function

Private Function GetFuncPtr(ByVal lngFnPtr As Long) As Long
    GetFuncPtr = lngFnPtr
End Function

Private Function ProcAddress(lpfn As Long) As Long
    ProcAddress = lpfn
End Function

Attribute VB_Name = "Module1"
Option Explicit
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Const WM_SYSCOMMAND = &H112
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const GWL_WNDPROC = (-4)
Public Const IDM_ABOUT As Long = 1010
Dim lProcOld As Long

Private Type WNDCLASSEX     ' Same as WNDCLASS but has a few advanced values
    cbSize As Long
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long               ' Handle to large icon (Alt-Tab icon)
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long             ' Handle to Small icon (Top Left Icon/Taskbar Icon)
End Type

Private Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long               ' Handle to icon (only 1 size)
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwflags As Long
    szexeFile As String * 260
End Type
Private Declare Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASSEX) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlgas As Long, ByVal lProcessID As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const GA_ROOT As Long = 2
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const WM_GETICON As Long = &H7F
Private Const GWL_HINSTANCE As Long = -6
Private Const GCL_HICON As Long = -14

Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long





Public Function MenuHandler(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If iMsg = WM_SYSCOMMAND Then
If wParam = IDM_ABOUT Then
    MsgBox "Programmed By:  Billy Conner" & Chr(13) & Chr(10) & "Jackyl_xyu@yahoo.com", vbInformation, "About"
Exit Function
End If
End If
MenuHandler = CallWindowProc(lProcOld, hwnd, iMsg, wParam, lParam)
End Function
Public Function SubClass(hwnd As Long)
Dim lhSysMenu As Long, lRet As Long
lhSysMenu = GetSystemMenu(hwnd, 0&)
lRet = AppendMenu(lhSysMenu, MF_SEPARATOR, 0&, vbNullString)
lRet = AppendMenu(lhSysMenu, MF_STRING, IDM_ABOUT, "About...")
lProcOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf MenuHandler)
End Function

'find a handle of a icon used by a form
Private Function GetIconHandle(hwnd As Long) As Long
  
  Dim ClassName As String
  Dim WCX As WNDCLASSEX
  Dim hInstance As Long
  Dim hIcon As Long
  Dim X As Long   ' temp variable
  Dim WC As WNDCLASS

    'Method: SendMessage (Small Icon)
    hIcon = SendMessage(hwnd, WM_GETICON, CLng(0), CLng(0))
  
    If hIcon > 0 Then ' found it
        GetIconHandle = hIcon
        Exit Function
    End If
    'Method: SendMessage (Large Icon)
    hIcon = SendMessage(hwnd, WM_GETICON, CLng(1), CLng(0))
    If hIcon > 0 Then ' found it
        GetIconHandle = hIcon
        Exit Function
    End If
    'Method: GetClassInfoEx (Small or Large with Small Pref.)
    hInstance = GetWindowLong(hwnd, GWL_HINSTANCE)
    WCX.cbSize = Len(WCX)
    ClassName = Space$(255)
    X = GetClassName(hwnd, ClassName, 255)
    X = GetClassInfoEx(hInstance, ClassName, WCX)
    If X > 0 Then
        If WCX.hIconSm = 0 Then 'No small icon
            hIcon = WCX.hIcon ' No small icon.. Windows should have given default.. weird
          Else
            hIcon = WCX.hIconSm ' Small Icon is better
        End If
        GetIconHandle = hIcon   ' found it =]
        Exit Function
    End If
 
    '*************************************
    'Method: GetClassInfo (Large Icon)
    '*************************************
    X = GetClassInfo(hInstance, ClassName, WC)
    If X > 0 Then
        hIcon = WC.hIcon
        GetIconHandle = hIcon
        Exit Function    ' Found it
    End If
        
    '*************************************
    'Method: GetClassLong (Large Icon)
    '*************************************
    X = GetClassLong(hwnd, GCL_HICON)
    If X > 0 Then
        hIcon = X
      Else
        hIcon = 0
    End If

    If hIcon < 0 Then
        hIcon = 0
    End If
    GetIconHandle = hIcon

End Function

Private Function GrabIcon(Optional ay = "") As Long

  Dim cc As Long
  Dim iconmod As String, numicons As Long
  Dim hModule As Long, iconh As Long
  Dim mainhwnd As Long

    If ay = "" Then
        cc = mainhwnd
      Else
        cc = CLng(Mid$(ay, 2, Len(ay) - 1))
    End If
    hModule = GetModuleHandle(0)
    iconmod$ = GetExeFromHandle(cc) + Chr$(0)  'prepares filename
    iconh = ExtractIcon(hModule, iconmod, -1) 'gets number of icons
    numicons = iconh - 1 'puts it into a variable
    If numicons > 0 Then
        iconh = ExtractIcon(hModule, iconmod, 0)     'Extracts the first icon
    End If
    GrabIcon = iconh

End Function

'extracts an icon from a file
Private Function GrabIconFromFile(File_name As String, IconNumber As Long) As Long

    GrabIconFromFile = ExtractIcon(GetModuleHandle(0), File_name, IconNumber)

End Function
Private Function GetExeFromHandle(wnd As Long) As String

  Dim ThreadId As Long, ProcessId As Long, hSnapshot As Long
  Dim uProcess As PROCESSENTRY32, rProcessFound As Long
  Dim i As Integer, szExename As String

    ' Get ID for window thread
    ThreadId = GetWindowThreadProcessId(wnd, ProcessId)
    ' Check if valid
    If ThreadId = 0 Or ProcessId = 0 Then
        Exit Function
    End If
    ' Create snapshot of current processes
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    ' Check if snapshot is valid
    If hSnapshot = -1 Then
        Exit Function
    End If
    'Initialize uProcess with correct size
    uProcess.dwSize = Len(uProcess)
    'Start looping through processes
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    Do While rProcessFound
        If uProcess.th32ProcessID = ProcessId Then
            'Found it, now get name of exefile
            i = InStr(1, uProcess.szexeFile, Chr$(0))
            If i Then
                szExename = Left$(uProcess.szexeFile, i - 1)
            End If
            Exit Do
          Else
            'Wrong ID, so continue looping
            rProcessFound = ProcessNext(hSnapshot, uProcess)
        End If
    Loop
    CloseHandle hSnapshot
    GetExeFromHandle = szExename

End Function
'my method of finding the best icon to use for my treeview
Public Function DetermineBestIcon(hwnd) As Long

  Dim iconh As Long
  Dim RetLen As Integer
  Dim sysdirbuff As String

    iconh = GetIconHandle(GetAncestor(hwnd, GA_ROOT))
    If iconh = 0 Then
        iconh = GrabIcon("t" & hwnd)
    End If
    If iconh = 0 Then
        sysdirbuff = String$(255, 0)
        RetLen = GetSystemDirectory(sysdirbuff, 255)
        sysdirbuff = Left$(sysdirbuff, RetLen)
        iconh = GrabIconFromFile(sysdirbuff & "\shell32.dll", 2)
    End If
    DetermineBestIcon = iconh

End Function
'returns the state of the window :normal,maximized,minimized
Public Function GetWindowState(hwnd As Long) As Long

  'Finds The Windowstate
  
  Dim h As Long

    h = 2 'Normal
    If IsIconic(hwnd) Then 'Minimized
        h = 0
    End If
    If IsZoomed(hwnd) Then 'Maximized
        h = 1
    End If
    GetWindowState = h

End Function
Public Function SetWindowState(hwnd As Long, NewState As Long) As Long   'Returns NewState

    ShowWindow hwnd, NewState
    SetWindowState = GetWindowState(hwnd)

End Function


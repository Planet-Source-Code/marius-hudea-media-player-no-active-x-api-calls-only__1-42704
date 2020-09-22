Attribute VB_Name = "ModAPI"
Option Explicit
'
'   These are all the API functions,types & constants this
' aplication uses. perhaps it would have been better to
' declare each one as private in their module bu there are
' 2 or more modules that uses the same functions.
'
'
'
' Used by modAVI
Public Declare Function AVIFileOpen Lib "avifil32" Alias "AVIFileOpenA" (ppfile As Long, ByVal szFile As String, ByVal mode As Long, pclsidHandler As Any) As Long
Public Declare Function AVIFileRelease Lib "avifil32" (ByVal pfile As Long) As Long
Public Declare Function AVIFileInfo Lib "avifil32" Alias "AVIFileInfoA" (ByVal pfile As Long, pfi As AVIFileInfo, ByVal lSize As Long) As Long
Public Declare Sub AVIFileInit Lib "avifil32" ()
Public Declare Sub AVIFileExit Lib "avifil32" ()
' Used by modMCI
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
' Used by modMCI, modTop & the main form
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' Used by modOpenFile
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
' Used by modProgressBar
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' Used by modStatus
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function InitCommonControlsEx Lib "comctl32" (init As InitCommonControlsExType) As Boolean
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'Used by modSubtitles
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WindowPlacement) As Long
'Used by frmToolBar
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointApi) As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'Used by modPlayList
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'Used by modPlayList - Add Folder Sub
Public Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As Win32_Find_Data)
Public Declare Function FindNextFile& Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As Win32_Find_Data)
Public Declare Function FindClose& Lib "kernel32" (ByVal hFindFile As Long)
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Public Type RECT 'Mantains the position of a window
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Public Type PointApi 'Point on the screen
        X As Long
        Y As Long
End Type

Public Type WindowPlacement
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As PointApi
        ptMaxPosition As PointApi
        rcNormalPosition As RECT
End Type

Public Type InitCommonControlsExType 'Used by modStatus
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type

Public Type AVIFileInfo          'Used by modAVI
    dwMaxBytesPerSec As Long
    dwFlags As Long
    dwCaps As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwLength As Long
    dwEditCount As Long
    szFileType As String * 64
End Type

Public Type OpenFileName 'Used by modOpenFile
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type BrowseInfo 'Select Folder Information
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type Win32_Find_Data
        dwFileAttributes As Long
        ftCreationTime As FileTime
        ftLastAccessTime As FileTime
        ftLastWriteTime As FileTime
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 255
        cAlternate As String * 14
End Type


Public Enum OnOrOff
AlwaysOn = 1
AlwaysOff = 0
End Enum


'
'   Folder/File attributes constants
'
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
'
'   Folder/File constants
'
Public Const INVALID_HANDLE_VALUE = -1
Public Const ERROR_NO_MORE_FILES = 18&
'
'   Open File Constants
'
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const SW_MINIMIZE = 6
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_HIDEREADONLY = &H4
Public Const OF_SHARE_DENY_WRITE = &H20

'
'   Always On Top ( SetWindowPos ) Constants
'
Public Const HWND_TOPMOST = -1          'Always On Top
Public Const HWND_NOTOPMOST = -2        'Always Off Top
Public Const SWP_NOSIZE = &H1           '
Public Const SWP_NOMOVE = &H2           '  Play with these to understand what
Public Const SWP_NOACTIVATE = &H10      ' they are for.
Public Const SWP_SHOWWINDOW = &H40      '
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOOWNERZORDER = &H200  '  Don't do owner Z ordering
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
'
'   Browse For Folder Constants
'
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
'
'   TaskBar Name Constant
'
Public Const TaskBar_Name As String = "Shell_traywnd"


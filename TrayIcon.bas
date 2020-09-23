Attribute VB_Name = "TrayIcon"
Public Const MAX_TOOLTIP As Integer = 64
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
End Type
Public nfIconData As NOTIFYICONDATA
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public TrayIconOn As Boolean 'This keeps the information are we currently using tray icon
Sub AddToTray()

With nfIconData
    .hwnd = Form_main.hwnd
    .uID = Form_main.Icon
    .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = Form_main.Icon.Handle
    .szTip = "VisualReminder" & vbNullChar
    .cbSize = Len(nfIconData)
End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)
TrayIconOn = True
End Sub


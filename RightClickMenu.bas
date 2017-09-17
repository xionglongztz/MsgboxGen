Attribute VB_Name = "RightClickMenu"
Option Explicit
'Download by http://www.codefans.net
' API函数声明
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

' 常数声明
Public Const WM_SYSCOMMAND = &H112
' 单击控制框产生此消息
Public Const MF_SEPARATOR = &H800&
' 为菜单加一条分隔线
Public Const MF_STRING = &H0&
' 在菜单中加一个字符串
Public Const GWL_WNDPROC = (-4)
' 全局变量
Public OldWindowProc As Long
' 保存默认的窗口函数地址
Public SysMenuHwnd As Long
' 保存系统菜单句柄
Public Function SubClass1_WndMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    If Msg <> WM_SYSCOMMAND Then
        SubClass1_WndMessage = CallWindowProc(OldWindowProc, hwnd, Msg, wp, lp)
        ' 如果消息不是WM_SYSCOMMAND，就调用默认的窗口函数处理
        Exit Function
    End If
    Select Case wp
        Case 2001
            Form1.Command1.Value = True
        Case 2002
            Form1.Command2.Value = True
        Case 2003
            Form1.Command3.Value = True
        Case 2004
            Form1.Command4.Value = True
        Case 2006
            MsgBox "开发者：雄龙ztz"
        'Case 2010
            'Call GetSystemMenu(Form1.hwnd, True)
            'Call SetWindowLong(Form1.hwnd, GWL_WNDPROC, OldWindowProc)
            'Call MsgBox("已经恢复了默认的系统菜单", vbOKOnly + vbInformation, "Hydrogen Browser")
        'Case 2011
        
        Case Else
            SubClass1_WndMessage = CallWindowProc(OldWindowProc, hwnd, Msg, wp, lp)
            Exit Function
    End Select
    SubClass1_WndMessage = True
End Function






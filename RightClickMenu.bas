Attribute VB_Name = "RightClickMenu"
Option Explicit
'Download by http://www.codefans.net
' API��������
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

' ��������
Public Const WM_SYSCOMMAND = &H112
' �������ƿ��������Ϣ
Public Const MF_SEPARATOR = &H800&
' Ϊ�˵���һ���ָ���
Public Const MF_STRING = &H0&
' �ڲ˵��м�һ���ַ���
Public Const GWL_WNDPROC = (-4)
' ȫ�ֱ���
Public OldWindowProc As Long
' ����Ĭ�ϵĴ��ں�����ַ
Public SysMenuHwnd As Long
' ����ϵͳ�˵����
Public Function SubClass1_WndMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    If Msg <> WM_SYSCOMMAND Then
        SubClass1_WndMessage = CallWindowProc(OldWindowProc, hwnd, Msg, wp, lp)
        ' �����Ϣ����WM_SYSCOMMAND���͵���Ĭ�ϵĴ��ں�������
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
            MsgBox "�����ߣ�����ztz"
        'Case 2010
            'Call GetSystemMenu(Form1.hwnd, True)
            'Call SetWindowLong(Form1.hwnd, GWL_WNDPROC, OldWindowProc)
            'Call MsgBox("�Ѿ��ָ���Ĭ�ϵ�ϵͳ�˵�", vbOKOnly + vbInformation, "Hydrogen Browser")
        'Case 2011
        
        Case Else
            SubClass1_WndMessage = CallWindowProc(OldWindowProc, hwnd, Msg, wp, lp)
            Exit Function
    End Select
    SubClass1_WndMessage = True
End Function






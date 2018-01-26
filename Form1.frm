VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MsgBox对话框代码生成器"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11475
   ForeColor       =   &H8000000C&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   11475
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "基础"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11295
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label Label1 
         Caption         =   "标题："
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "内容："
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "生成及调试"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   11295
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "生成并执行(&R)"
         Default         =   -1  'True
         Height          =   1215
         Left            =   7320
         TabIndex        =   31
         Top             =   720
         Width           =   3855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "重置所有选项(&Z)"
         Height          =   855
         Left            =   2880
         TabIndex        =   36
         Top             =   1080
         Width           =   4455
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "复制到剪贴板(&C)"
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "生成代码(&G)"
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   11055
      End
      Begin VB.Label Label4 
         Caption         =   "返回值："
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "高级"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11295
      Begin VB.Frame Frame7 
         Caption         =   "其他"
         Height          =   1695
         Left            =   7560
         TabIndex        =   16
         Top             =   240
         Width           =   3615
         Begin VB.OptionButton Option17 
            Caption         =   "运行时系统被挂起"
            Height          =   225
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   3375
         End
         Begin VB.OptionButton Option16 
            Caption         =   "运行时应用程序挂起"
            Height          =   180
            Left            =   120
            TabIndex        =   34
            Top             =   960
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "从右往左显示"
            Height          =   180
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   3135
         End
         Begin VB.CheckBox Check3 
            Caption         =   "文本右对齐"
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   3135
         End
         Begin VB.CheckBox Check2 
            Caption         =   "指定消息框为前景窗口"
            Height          =   180
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "图标样式"
         Height          =   1695
         Left            =   5160
         TabIndex        =   15
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton Option15 
            Caption         =   "通知信息"
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton Option14 
            Caption         =   "警告信息"
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton Option13 
            Caption         =   "询问信息"
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton Option12 
            Caption         =   "错误信息"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton Option11 
            Caption         =   "无"
            Height          =   180
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "默认按钮"
         Height          =   1695
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton Option10 
            Caption         =   "第四个"
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton Option9 
            Caption         =   "第三个"
            Height          =   180
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton Option8 
            Caption         =   "第二个"
            Height          =   180
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton Option7 
            Caption         =   "第一个"
            Height          =   180
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "按钮形式"
         Height          =   1695
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton Option6 
            Caption         =   "重试和取消"
            Height          =   180
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton Option5 
            Caption         =   "是和否"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   2175
         End
         Begin VB.OptionButton Option4 
            Caption         =   "是,否和取消"
            Height          =   180
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton Option3 
            Caption         =   "中止,重试和忽略"
            Height          =   180
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton Option2 
            Caption         =   "确定和取消"
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "仅确定"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim anniu, moren, tubiao
Dim zonghe
Dim fanhuizhi
Private Sub Command1_Click()
zonghe = anniu + tubiao + moren
If Check1.Value = 1 Then zonghe = zonghe + 1048576
If Check2.Value = 1 Then zonghe = zonghe + 65536
If Check3.Value = 1 Then zonghe = zonghe + 524288
If Option17.Value = True Then zonghe = zonghe + 4096
Text3.Text = "MsgBox " & """" & Text2.Text & """" & "," & zonghe & "," & """" & Text1.Text & """"
End Sub

Private Sub Command2_Click()
Command1.Value = True
fanhuizhi = MsgBox(Text2.Text, zonghe, Text1.Text)
Label4.Caption = "返回值：" & fanhuizhi
End Sub

Private Sub Command3_Click()
Clipboard.Clear
Clipboard.SetText Text3.Text
MsgBox "已复制到剪贴板"
End Sub

Private Sub Command4_Click()
Dim chongzhi
chongzhi = MsgBox("你确定要重置所有选项？", 4)
If chongzhi = 6 Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label4.Caption = "返回值："
Option1.Value = True
Option7.Value = True
Option11.Value = True
Option16.Value = True
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
End If
End Sub

Private Sub Form_Load()
'命令行
Dim cmdstring As String
cmdstring = Command()
cmdstring = Trim(cmdstring) '替换变量

If cmdstring = "/?" Then
MsgBox "用法：" & Chr(10) & App.EXEName & " [/?]"
End
End If

'检测是否两个同时运行
If App.PrevInstance = True Then
MsgBox "程序已经运行！", 48
End
End If
'窗口菜单
OldWindowProc = GetWindowLong(Me.hwnd, GWL_WNDPROC) ' 取得窗口函数的地址
Call SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf SubClass1_WndMessage) ' 用SubClass1_WndMessage代替窗口函数处理消息
SysMenuHwnd = GetSystemMenu(Me.hwnd, False)
SysMenuHwnd = GetSystemMenu(Me.hwnd, False)
Call AppendMenu(SysMenuHwnd, MF_SEPARATOR, 2000, vbNullString)
Call AppendMenu(SysMenuHwnd, MF_STRING, 2001, "生成代码(&G)")
Call AppendMenu(SysMenuHwnd, MF_STRING, 2002, "生成并执行(&R)")
Call AppendMenu(SysMenuHwnd, MF_STRING, 2003, "复制到剪贴板(&C)")
Call AppendMenu(SysMenuHwnd, MF_STRING, 2004, "重置所有选项(&Z)")
Call AppendMenu(SysMenuHwnd, MF_SEPARATOR, 2005, vbNullString)
Call AppendMenu(SysMenuHwnd, MF_STRING, 2006, "开发者信息(&D)")

'Call AppendMenu(SysMenuHwnd, MF_STRING, 2010, "恢复菜单(&R)")
End Sub

Private Sub Option1_Click()
anniu = 0
End Sub

Private Sub Option10_Click()
moren = 768
End Sub

Private Sub Option11_Click()
tubiao = 0
End Sub

Private Sub Option12_Click()
tubiao = 16
End Sub

Private Sub Option13_Click()
tubiao = 32
End Sub

Private Sub Option14_Click()
tubiao = 48
End Sub

Private Sub Option15_Click()
tubiao = 64
End Sub

Private Sub Option2_Click()
anniu = 1
End Sub

Private Sub Option3_Click()
anniu = 2
End Sub

Private Sub Option4_Click()
anniu = 3
End Sub

Private Sub Option5_Click()
anniu = 4
End Sub

Private Sub Option6_Click()
anniu = 5
End Sub

Private Sub Option7_Click()
moren = 0
End Sub

Private Sub Option8_Click()
moren = 256
End Sub

Private Sub Option9_Click()
moren = 512
End Sub

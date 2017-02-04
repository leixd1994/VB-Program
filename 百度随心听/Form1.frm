VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "时钟 BY thunder"
   ClientHeight    =   8190
   ClientLeft      =   5730
   ClientTop       =   615
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   14430
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      ExtentX         =   25426
      ExtentY         =   14420
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5520
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindows Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10
'new


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, _
 ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, _
 ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
 
Private Sub Form_Load()


 'myval = SetWindowPos(Form1.hwnd, -1, 0, 0, 0, 0, 3)
 Me.BackColor = vbBlue

Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbBlue, 190, LWA_COLORKEY

Label1.BackColor = vbBlue
WebBrowser1.Navigate ("http://fm.baidu.com")

'new
End Sub
Private Sub Webbrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

Shell "c:\auto.exe"
Sleep 1500
'winhwnd = FindWindows("Windows 任务管理器", vbNullString) '知道窗口类名，关闭
winhwnd = FindWindows(vbNullString, "auto") '知道窗口标题，关闭
Call PostMessage(winhwnd, WM_CLOSE, 0&, 0&)




End Sub




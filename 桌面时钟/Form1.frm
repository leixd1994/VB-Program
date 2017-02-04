VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   " ±÷” BY thunder"
   ClientHeight    =   13665
   ClientLeft      =   -1905
   ClientTop       =   -1065
   ClientWidth     =   25005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   13665
   ScaleWidth      =   25005
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Ω‚À¯"
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3015
      Left            =   3960
      TabIndex        =   0
      Top             =   3360
      Width           =   8775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, _
 ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
 ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
 
Private Sub Command1_Click()
If Text1.Text = "abc1234" Then
Unload Form1
Else
Text1.Text = "√‹¬Î¥ÌŒÛ"
End If
End Sub

Private Sub Form_Load()


 myval = SetWindowPos(Form1.hwnd, -1, 0, 0, 0, 0, 3)
 Me.BackColor = vbBlue
Label1.BackColor = vbBlue
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbBlue, 190, LWA_COLORKEY
Timer1.Interval = 300
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Now()
 myval = SetWindowPos(Form1.hwnd, -1, 0, 0, 0, 0, 3)
End Sub

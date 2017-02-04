VERSION 5.00
Begin VB.Form IAudioEndpointVolume 
   BorderStyle     =   0  'None
   Caption         =   "win7Ö÷ÒôÁ¿¿ØÖÆ"
   ClientHeight    =   525
   ClientLeft      =   17205
   ClientTop       =   210
   ClientWidth     =   1740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleWidth      =   1740
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtAudioLevel 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton VolumeStepDown 
      Caption         =   "VolumeStepDown "
      Height          =   360
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton VolumeStepUp 
      Caption         =   "VolumeStepUp"
      Height          =   360
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton GetMasterVolume 
      Caption         =   "GetMasterVolume"
      Height          =   360
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton SetMasterVolume 
      Caption         =   "SetMasterVolume"
      Height          =   360
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton SetMute 
      Caption         =   "SetMute"
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   990
   End
   Begin VB.CommandButton GetMute 
      Caption         =   "GetMute"
      Height          =   360
      Left            =   2760
      TabIndex        =   0
      Top             =   2400
      Width           =   990
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   0
   End
   Begin VB.Label lblVolumeLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0-100)"
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   1680
      Width           =   540
   End
   Begin VB.Label lblGetMasterVolume 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GetMasterVolume"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   555
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2220
   End
   Begin VB.Label lblGetMute 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GetMute"
      Height          =   195
      Left            =   3960
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
End
Attribute VB_Name = "IAudioEndpointVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, _
 ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
 ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Sub Form_Load()
myval = SetWindowPos(IAudioEndpointVolume.hwnd, -1, 0, 0, 0, 0, 3)
' Me.BackColor = vbBlue
lblGetMasterVolume.BackColor = vbBlue
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbBlue, 190, LWA_COLORKEY
End Sub

Private Sub GetMasterVolume_Click()
    Dim Audio As CAudioEndpointVolume
    Set Audio = New CAudioEndpointVolume
    lblGetMasterVolume = Audio.GetMasterVolumeLevelScalar    '0.5 = %50
    
    
End Sub

Private Sub GetMute_Click()
    Dim Audio As CAudioEndpointVolume
    Set Audio = New CAudioEndpointVolume
    lblGetMute = Audio.GetMute & "   (0:un-Mute / 1: Mute)"
End Sub

Private Function CheckExeIsRun(exeName As String) As Boolean
    On Error GoTo Err
    Dim WMI
    Dim Obj
    Dim Objs
    CheckExeIsRun = False
    Set WMI = GetObject("WinMgmts:")
    Set Objs = WMI.InstancesOf("Win32_Process")
    For Each Obj In Objs
      If (InStr(UCase(exeName), UCase(Obj.Description)) <> 0) Then
            CheckExeIsRun = True
            If Not Objs Is Nothing Then Set Objs = Nothing
            If Not WMI Is Nothing Then Set WMI = Nothing
            Exit Function
      End If
    Next
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
    Exit Function
Err:
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
End Function


Private Sub SetMasterVolume_Click()
    Dim Audio As CAudioEndpointVolume
    Set Audio = New CAudioEndpointVolume
    Dim AudioLevel As Double
    If TxtAudioLevel = "" Then Exit Sub
    AudioLevel = CLng(TxtAudioLevel.Text)
    AudioLevel = AudioLevel / 100
    'Audio.SetMasterVolumeLevelScalar 0.25 ' 25%
    Audio.SetMasterVolumeLevelScalar AudioLevel
    
End Sub

Private Sub SetMute_Click()
    
    Dim Audio As CAudioEndpointVolume
    Set Audio = New CAudioEndpointVolume
    If Audio.GetMute = 0 Then
        Audio.SetMute 1 ' mute
    Else
        Audio.SetMute 0 ' un-mute
    End If
    
End Sub


Private Sub Timer1_Timer()
Dim Audio As CAudioEndpointVolume
    Set Audio = New CAudioEndpointVolume
    lblGetMasterVolume = Int(Audio.GetMasterVolumeLevelScalar * 100)  '0.5 = %50
    If lblGetMasterVolume > 20 Then
lblGetMasterVolume.ForeColor = &HFF&

myval = SetWindowPos(IAudioEndpointVolume.hwnd, -1, 0, 0, 0, 0, 3)
Else
lblGetMasterVolume.ForeColor = &HC000&
End If
If CheckExeIsRun("QQPlayer.exe") Then
Else
If lblGetMasterVolume > 26 Then
TxtAudioLevel.Text = 7
SetMasterVolume_Click
  End If
End If
End Sub

Private Sub VolumeStepDown_Click()
    Dim Audio As CAudioEndpointVolume
    Set Audio = New CAudioEndpointVolume
    Audio.VolumeStepDown
End Sub

Private Sub VolumeStepUp_Click()
    Dim Audio As CAudioEndpointVolume
    Set Audio = New CAudioEndpointVolume
    Audio.VolumeStepUp
End Sub

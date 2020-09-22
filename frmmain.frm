VERSION 5.00
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4980
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   3660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmmain.frx":030A
   ScaleHeight     =   4980
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Minimise"
      Height          =   495
      Left            =   1815
      TabIndex        =   12
      Top             =   4140
      Width           =   990
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1995
      Top             =   465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Timer"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   4140
      Width           =   1005
   End
   Begin VB.OptionButton optlo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log Off"
      Height          =   330
      Left            =   2490
      TabIndex        =   7
      Top             =   3480
      Width           =   900
   End
   Begin VB.OptionButton optr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Restart"
      Height          =   330
      Left            =   1515
      TabIndex        =   6
      Top             =   3480
      Width           =   900
   End
   Begin VB.OptionButton optsd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shut Down"
      Height          =   330
      Left            =   225
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   27.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   195
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2310
      Width           =   3210
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1575
      Top             =   465
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1650
      TabIndex        =   11
      Top             =   30
      Width           =   315
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   360
      TabIndex        =   10
      Top             =   300
      Width           =   330
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2940
      TabIndex        =   9
      Top             =   300
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose Your Option"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   2925
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Desired Time"
      Height          =   210
      Left            =   1230
      TabIndex        =   2
      Top             =   2115
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Time"
      Height          =   225
      Left            =   1305
      TabIndex        =   1
      Top             =   1185
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   180
      TabIndex        =   0
      Top             =   1335
      Width           =   3240
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const EWX_POWEROFF = 1
Private Const EWX_REBOOT = 2
Private Const EWX_LOGOFF = 0
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Sub Command2_Click()
frmmain.WindowState = 1
End Sub

Private Sub Form_Load()
'the form must be fully visible before calling Shell_NotifyIcon
Me.Show
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = "Timer" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim msg As Long
'the value of X will vary depending upon the scalemode setting
If Me.ScaleMode = vbPixels Then
msg = X
Else
msg = X / Screen.TwipsPerPixelX
End If
Select Case msg
Case WM_LBUTTONUP '514 restore form window
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
Me.Show
Case WM_LBUTTONDBLCLK '515 restore form window
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
Me.Show
Case WM_RBUTTONUP '517 display popup menu
Result = SetForegroundWindow(Me.hwnd)
Me.PopupMenu Me.mPopupSys
End Select
End Sub

Private Sub Form_Resize()
'this is necessary to assure that the minimized window is hidden
If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
'this removes the icon from the system tray
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mPopExit_Click()
'called when user clicks the popup menu Exit command
Unload Me
End Sub

Private Sub mPopRestore_Click()
'called when the user clicks the popup menu Restore command
Me.WindowState = vbNormal
'Result = SetForegroundWindow(Me.hwnd)
Me.Show
End Sub
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox ("Please Enter a Time")
Else
If optsd.Value = False And optlo.Value = False And optr.Value = False Then
MsgBox ("Please select an option")
Else
MsgBox ("Timer set at : ") + Text1.Text
Timer2.Enabled = True
End If
End If
End Sub
Private Sub Label5_Click()
Unload Me
End Sub

Private Sub Label6_Click()
SendKeys "{F1}", True
'MsgBox ("This is where what you have to click on if you want help <may make this press the F1 button>")
End Sub

Private Sub Label7_Click()
frmabout.Visible = True
End Sub

Private Sub mnuAbout_Click()
MsgBox ("This is where the about box will appear from")
End Sub
Private Sub Timer1_Timer()
Label1.Caption = Time
End Sub
Private Sub Timer2_Timer()
If Text1.Text = Label1.Caption Then
    If optsd.Value = True Then
    'Shut the computer off
    Unload Me
    ExitWindowsEx EWX_POWEROFF, 1
     
    ElseIf optr.Value = True Then
    'Reboot the computer
    Unload Me
    ExitWindowsEx EWX_REBOOT, 2
    'End If
     
    ElseIf optlo.Value = True Then
    'Logoff
    Unload Me
    ExitWindowsEx EWX_LOGOFF, 0
    End If

End If
End Sub

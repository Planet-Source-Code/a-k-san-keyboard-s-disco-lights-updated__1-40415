VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard's Disco Lights"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Lights Indicator"
      Height          =   975
      Left            =   3960
      TabIndex        =   13
      Top             =   2160
      Width           =   2655
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Form1.frx":0442
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1080
         Picture         =   "Form1.frx":074C
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1920
         Picture         =   "Form1.frx":0A56
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   240
         Picture         =   "Form1.frx":0D60
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   1080
         Picture         =   "Form1.frx":106A
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   1920
         Picture         =   "Form1.frx":1374
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   360
      Top             =   5160
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   360
      Top             =   5160
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   360
      Top             =   5160
   End
   Begin VB.Frame Frame2 
      Caption         =   "2 Lights"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2040
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
      Begin VB.OptionButton random2 
         Caption         =   "Random"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton ltr2 
         Caption         =   "Left to Right"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton rtl2 
         Caption         =   "Right to Left"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "1 Light"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
      Begin VB.OptionButton rtl1 
         Caption         =   "Right to Left"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton ltr1 
         Caption         =   "Left to Right"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton random1 
         Caption         =   "Random"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   360
      Top             =   5160
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   360
      Top             =   5160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Minimize to System Tray"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Turn off monitor"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   360
      Top             =   5160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sansoft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   4
      Top             =   3360
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Including the monitor."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3060
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "If possible, please turn off the lights in the room."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************
'* Author: A.K.San     *
'* Date  : 05 NOV 2002 *
'***********************
Dim nid As NOTIFYICONDATA 'to hold data for the notifycondata type
Dim X As Integer, Y As Integer, z As Integer, l As Byte 'three variables to hold the default LED state and one for the counter

Private Sub Command1_Click()
'not really turning off the monitor
'just put on a black colored and maximized form only :)
Load Form2
Form2.Show
Label2.Caption = "Please rate it if you like it."
Label3.Caption = "Thanks for trying it."
End Sub

Private Sub Command2_Click()
'getting all the data needed by the function
nid.cbSize = Len(nid)
nid.hWnd = Form1.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Form1.Icon
nid.szTip = Form1.Caption & vbNullChar
'minimize and put on the system tray's icon
Me.WindowState = 1
Me.Visible = False
Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_Activate()
Call linitial 'reset the image lights
End Sub

Private Sub Form_Click()
Me.WindowState = 1
End Sub

Private Sub Form_Load()
l = 1
'get the default settings of the lights
X = GetKeyState(num)
Y = GetKeyState(cap)
z = GetKeyState(scr)
'reset all the lights first
Call keyreset
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim but As Long
but = X / Screen.TwipsPerPixelX
Select Case but
Case WM_LBUTTONDOWN
Shell_NotifyIcon NIM_DELETE, nid
Me.WindowState = 0
Me.Visible = True
Me.SetFocus
End Select
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid 'clear off the system icon made by the program
End Sub

Private Sub Form_Unload(Cancel As Integer)
'reset the lights again before set it back to default
Call keyreset
'set the lights back to default
If X = 1 Then Call keynum
If Y = 1 Then Call keycap
If z = 1 Then Call keyscr
End Sub

Private Sub ltr1_Click() 'on the required timer and set off all the option values
Call keyreset
l = 1
random2.Value = False
rtl2.Value = False
ltr2.Value = False
Call falsealltimer
Timer2.Enabled = True
End Sub

Private Sub ltr2_Click() 'on the required timer and set off all the option values
Call keyreset
l = 1
random1.Value = False
rtl1.Value = False
ltr1.Value = False
Call falsealltimer
Timer4.Enabled = True
End Sub

Private Sub random1_Click() 'on the required timer and set off all the option values
Call keyreset
random2.Value = False
rtl2.Value = False
ltr2.Value = False
Call falsealltimer
Timer1.Enabled = True
End Sub

Private Sub random2_Click() 'on the required timer and set off all the option values
Call keyreset
random1.Value = False
rtl1.Value = False
ltr1.Value = False
Call falsealltimer
Timer3.Enabled = True
End Sub

Private Sub rtl1_Click() 'on the required timer and set off all the option values
Call keyreset
l = 1
random2.Value = False
rtl2.Value = False
ltr2.Value = False
Call falsealltimer
Timer5.Enabled = True
End Sub

Private Sub rtl2_Click() 'on the required timer and set off all the option values
Call keyreset
l = 1
random1.Value = False
rtl1.Value = False
ltr1.Value = False
Call falsealltimer
Timer6.Enabled = True
End Sub

Private Sub Timer1_Timer() 'timer for 1 light random
Dim light As Byte 'a variable to hold randomly generated integer
Randomize
light = Int(Rnd * 3) + 1 'generate the integer
If light = 1 Then 'when = 1 on\off the numlock light
If GetKeyState(num) <> 1 Then Call keynum
If GetKeyState(cap) = 1 Then Call keycap
If GetKeyState(scr) = 1 Then Call keyscr
ElseIf light = 2 Then 'when = 2 on\off the capslock light
If GetKeyState(cap) <> 1 Then Call keycap
If GetKeyState(num) = 1 Then Call keynum
If GetKeyState(scr) = 1 Then Call keyscr
ElseIf light = 3 Then 'when = 3 on\off the scroll lock light
If GetKeyState(scr) <> 1 Then Call keyscr
If GetKeyState(cap) = 1 Then Call keycap
If GetKeyState(num) = 1 Then Call keynum
End If
End Sub

Private Sub Timer2_Timer() 'timer for 1 light move from left to right
'these codes are rather hard to explain so DIY
If l = 1 Then
Call keynum
l = 2
ElseIf l = 2 Then
Call keynum
Call keycap
l = 3
ElseIf l = 3 Then
Call keycap
Call keyscr
l = 4
ElseIf l = 4 Then
Call keyscr
l = 1
End If
End Sub

Private Sub Timer3_Timer() 'timer for 2 lights random
Dim light As Byte 'a variable to hold randomly generated integer
Randomize
light = Int(Rnd * 3) + 1 'generate the integer
If light = 1 Then 'when = 1 on\off the numlock light
Call keynum
Call keycap
ElseIf light = 2 Then 'when = 2 on\off the capslock light
Call keynum
Call keyscr
ElseIf light = 3 Then 'when = 3 on\off the scroll lock light
Call keyscr
Call keycap
End If
End Sub

Private Sub Timer4_Timer() 'timer for 2 lights move from left to right
'these codes are rather hard to explain so DIY
If l = 1 Then
Call keynum
l = 2
ElseIf l = 2 Then
Call keycap
l = 3
ElseIf l = 3 Then
Call keynum
Call keyscr
l = 4
ElseIf l = 4 Then
Call keycap
l = 5
ElseIf l = 5 Then
Call keyscr
l = 1
End If
End Sub

Private Sub Timer5_Timer() 'timer for 1 light move from right to left
'these codes are rather hard to explain so DIY
If l = 1 Then
Call keyscr
l = 2
ElseIf l = 2 Then
Call keyscr
Call keycap
l = 3
ElseIf l = 3 Then
Call keycap
Call keynum
l = 4
ElseIf l = 4 Then
Call keynum
l = 1
End If
End Sub

Private Sub Timer6_Timer() 'timer for 2 lights move from right to left
'these codes are rather hard to explain so DIY
If l = 1 Then
Call keyscr
l = 2
ElseIf l = 2 Then
Call keycap
l = 3
ElseIf l = 3 Then
Call keyscr
Call keynum
l = 4
ElseIf l = 4 Then
Call keycap
l = 5
ElseIf l = 5 Then
Call keynum
l = 1
End If
End Sub

Private Sub falsealltimer() 'a sub to set all the timer.enabled to false
Timer6.Enabled = False
Timer5.Enabled = False
Timer4.Enabled = False
Timer3.Enabled = False
Timer2.Enabled = False
Timer1.Enabled = False
End Sub

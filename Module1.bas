Attribute VB_Name = "Module1"
'**************************
'* Module made by A.K.San *
'* Date: 05 NOV 2002      *
'**************************

Public Type NOTIFYICONDATA 'the data type needed by the function
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'some constants for shell_notifyicon's function
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDOWN = &H201

'all the API declarations needed for this program
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim a As Integer, b As Integer, c As Integer

'some constants for later use
Public Const KEYEVENTF_KEYUP = &H2
Public Const num = &H90
Public Const scr = &H91
Public Const cap = &H14

Public Sub keycap()
Dim ret As Long
ret = MapVirtualKey(cap, 0)
keybd_event cap, ret, 0, 0
keybd_event cap, ret, KEYEVENTF_KEYUP, 0
Call l2
End Sub

Public Sub keynum()
Dim ret As Long
ret = MapVirtualKey(num, 0)
keybd_event num, ret, 0, 0
keybd_event num, ret, KEYEVENTF_KEYUP, 0
Call l1
End Sub

Public Sub keyscr()
Dim ret As Long
ret = MapVirtualKey(scr, 0)
keybd_event scr, ret, 0, 0
keybd_event scr, ret, KEYEVENTF_KEYUP, 0
Call l3
End Sub

Public Sub keyreset()
a = GetKeyState(num)
b = GetKeyState(cap)
c = GetKeyState(scr)
If a = 1 Then Call keynum
If b = 1 Then Call keycap
If c = 1 Then Call keyscr
End Sub

Public Sub l1()
If Form1.Image1.Visible = True Then
Form1.Image1.Visible = False
Form1.Image4.Visible = True
Else
Form1.Image4.Visible = False
Form1.Image1.Visible = True
End If
End Sub

Public Sub l2()
If Form1.Image2.Visible = True Then
Form1.Image2.Visible = False
Form1.Image5.Visible = True
Else
Form1.Image5.Visible = False
Form1.Image2.Visible = True
End If
End Sub

Public Sub l3()
If Form1.Image3.Visible = True Then
Form1.Image3.Visible = False
Form1.Image6.Visible = True
Else
Form1.Image6.Visible = False
Form1.Image3.Visible = True
End If
End Sub

Public Sub linitial()
If a = 1 Then Call l1
If b = 1 Then Call l2
If c = 1 Then Call l3
End Sub


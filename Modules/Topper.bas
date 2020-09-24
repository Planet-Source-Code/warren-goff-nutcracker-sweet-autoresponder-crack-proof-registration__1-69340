Attribute VB_Name = "Topper"
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Option Explicit
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


'declare for moving the form
Public Declare Function ReleaseCapture Lib "user32" () As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Global c As String 'to store current form's caption




Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
 If Topmost = True Then 'Make the window topmost
  SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
 Else
  SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
  SetTopMostWindow = False
 End If
End Function



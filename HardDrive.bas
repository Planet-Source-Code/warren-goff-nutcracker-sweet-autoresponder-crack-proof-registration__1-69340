Attribute VB_Name = "HardDrive"
Option Explicit
Global HardDriveS As String
Global ChrHardDriveS As String
Global ReceiveFlag As Boolean
Global MacAddress As String
Private Declare Function GetVolumeInformation Lib _
"kernel32.dll" Alias "GetVolumeInformationA" (ByVal _
lpRootPathName As String, ByVal lpVolumeNameBuffer As _
String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength _
As Long, lpFileSystemFlags As Long, ByVal _
lpFileSystemNameBuffer As String, ByVal _
nFileSystemNameSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long


Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Sub Delay(HowLong As Date)
Dim TempTime As Variant
TempTime = DateAdd("s", HowLong, Now)
While TempTime > Now
DoEvents 'Allows windows to handle other stuff
Wend
End Sub
Sub Main()
    'You will need to rename this variable to Schnibble in order to embed a different serial number
    HardDriveS = "Schnibble"
    Load ServerInfo
    ServerInfo.Show
End Sub


Public Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    Res = GetVolumeInformation(strDrive, Temp1, _
    Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
End Function



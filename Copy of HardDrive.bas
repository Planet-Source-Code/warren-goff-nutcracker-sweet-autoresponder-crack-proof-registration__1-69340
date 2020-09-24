Attribute VB_Name = "HardDrive"
Option Explicit
Global HardDriveS As String
Global ChrHardDriveS As String

Private Declare Function GetVolumeInformation Lib _
"kernel32.dll" Alias "GetVolumeInformationA" (ByVal _
lpRootPathName As String, ByVal lpVolumeNameBuffer As _
String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength _
As Long, lpFileSystemFlags As Long, ByVal _
lpFileSystemNameBuffer As String, ByVal _
nFileSystemNameSize As Long) As Long
Sub Main()
    HardDriveS = "1150874195"
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



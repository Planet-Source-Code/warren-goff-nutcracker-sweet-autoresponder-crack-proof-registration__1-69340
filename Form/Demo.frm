VERSION 5.00
Begin VB.Form Demo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12495
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H80000009&
      Caption         =   "Login"
      Height          =   270
      Left            =   1005
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5775
      Width           =   945
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H80000009&
      Caption         =   "Logout"
      Height          =   270
      Left            =   2025
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5775
      Width           =   945
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6165
      Top             =   4755
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   7245
      TabIndex        =   21
      Top             =   15
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   7605
      TabIndex        =   20
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtMail 
      Enabled         =   0   'False
      Height          =   2220
      Left            =   -15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   3510
      Width           =   4005
   End
   Begin VB.ListBox lstMail 
      Enabled         =   0   'False
      Height          =   1620
      Left            =   0
      TabIndex        =   10
      Top             =   1875
      Width           =   3960
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H80000014&
      Height          =   1650
      Left            =   -30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Demo.frx":08CA
      Top             =   240
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "POP3 mail server"
      Height          =   1815
      Left            =   4110
      TabIndex        =   1
      Top             =   1410
      Width           =   7695
      Begin VB.CommandButton Command4 
         Caption         =   "r"
         Height          =   255
         Left            =   7395
         TabIndex        =   19
         Top             =   645
         Width           =   210
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Compile"
         Height          =   270
         Left            =   6630
         TabIndex        =   18
         Top             =   615
         Width           =   705
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Serial"
         Height          =   270
         Left            =   6000
         TabIndex        =   17
         Top             =   645
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send Email"
         Height          =   315
         Left            =   4335
         TabIndex        =   16
         Top             =   660
         Width           =   1410
      End
      Begin VB.TextBox tbxCount 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox tbxSize 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6105
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1380
         Width           =   1455
      End
      Begin VB.TextBox tbxPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   8
         Text            =   "YourPassword"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox tbxUser 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "Your Name"
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox tbxHost 
         Height          =   285
         Left            =   1185
         TabIndex        =   3
         Text            =   "Your POP3 server"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox tbxPort 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "110"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Number of messages :"
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Total size of messages :"
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Password :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "User name :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Server name :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Server port :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Receive Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1005
      TabIndex        =   22
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents POP3 As EVICT_POP3.POP3
Attribute POP3.VB_VarHelpID = -1
Private Declare Function GetVolumeInformation Lib _
"kernel32.dll" Alias "GetVolumeInformationA" (ByVal _
lpRootPathName As String, ByVal lpVolumeNameBuffer As _
String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength _
As Long, lpFileSystemFlags As Long, ByVal _
lpFileSystemNameBuffer As String, ByVal _
nFileSystemNameSize As Long) As Long



Public Sub cmdLogin_Click()
Dim intloop As Integer
Dim id As Integer, X As Integer, y As Integer, i As Long, Subj As String, Ret As String
    Set POP3 = Nothing
    Set POP3 = New EVICT_POP3.POP3
    POP3.POPHost = Me.tbxHost
    POP3.POPPort = Me.tbxPort
    POP3.UserName = Me.tbxUser
    POP3.Password = Me.tbxPassword
    POP3.Connect
    
    If POP3.MessageCount = 0 Then
       'MsgBox "Sorry there are no messages in your inbox."
       ServerInfo.Caption = "Nutcracker Sweet: no messages"
    Else
        Me.tbxCount = POP3.MessageCount
        Me.tbxSize = POP3.TotalMessageSize
        ServerInfo.Caption = "Nutcracker Sweet: " & POP3.MessageCount & " messages"
        lstMail.Clear
        For intloop = 1 To POP3.MessageCount
        DoEvents
            lstMail.AddItem intloop & ", " & POP3.MessageSize(intloop) & ", " & POP3.UniqueMessageID(intloop) & ", " & POP3.Subject(intloop)
            Subj = POP3.Subject(intloop)
            Ret = POP3.From(intloop)
            If InStr(Subj, "Schnibble") <> 0 Then
                'id = Left(lstMail.List(intloop - 1), InStr(1, lstMail.List(intloop - 1), ",") - 1)
                'MsgBox id
                'frmMain.txtTo = ret
                X = InStr(Trim(Subj), ">,") + 2
                y = Len(Trim(Subj)) - X + 1
                HardDriveS = Mid(Trim(Subj), X, y)
                'For i = 0 To List2.ListCount - 1
                    'If InStr(List2.List(i), HardDriveS) <> 0 Then Exit For
                'Next
                List2.AddItem HardDriveS
                Command2_Click
                Exit For
            End If
        Next
    End If
End Sub


Public Sub Command2_Click()
Dim i As Long, X As String, y As String, MacAd As String
HardDriveS = GetSerialNumber("C:\")
Open App.Path & "\MacAddress" For Input As #1
    Do While Not EOF(1)
        Line Input #1, MacAd
        X = "        Physical Address. . . . . . . . . : "
        If InStr(MacAd, X) <> 0 Then
            MacAddress = Replace(MacAd, X, "")
            'MsgBox MacAddress
        End If
    Loop
Close #1
y = Str(HardDriveS)
ChrHardDriveS = ""
X = ""
'converts the serial number into chr(asc) format
For i = 1 To Len(y)
    X = "chr(" & Trim(Asc(Mid(y, i, 1))) & ")"
    If i <> Len(y) Then
        ChrHardDriveS = ChrHardDriveS & X & " & "
    Else
        ChrHardDriveS = ChrHardDriveS & X
    End If
Next
ChrHardDriveS = ChrHardDriveS & " & " & """=++=""" & " & "
y = MacAddress
X = ""
'converts the serial number into chr(asc) format
For i = 1 To Len(y)
    X = "chr(" & Trim(Asc(Mid(y, i, 1))) & ")"
    If i <> Len(y) Then
        ChrHardDriveS = ChrHardDriveS & X & " & "
    Else
        ChrHardDriveS = ChrHardDriveS & X
    End If
Next

' Schnibble is the keyword in the subject for the autoresponder and for the
' replacement text in HardDrive.bas

frmMain.txtSubject.Text = "Schnibble:" & ChrHardDriveS
If ReceiveFlag = True Then
    ServerInfo.Caption = "Nutcracker Sweet: Auto-Responding!"
    Command4_Click
End If
End Sub

Private Sub Command3_Click()
    Dim retval As Long
    'Timer1.Enabled = False
    'ServerInfo.Label1.Caption = "Recv"
    'ServerInfo.Caption = "Nutcracker Sweet:"
' If the project being compiled is open in the IDE, this will generate an error.
' If the exe file is running this will generate an error as it cannot write the compiled project exe
    If ServerInfo.Check1.Value = 1 Then
        Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /make " & _
            ServerInfo.Text1.Text, vbNormalFocus
    End If
    frmMain.txtAttach = ServerInfo.Text3.Text
    
    Do While Dir(ServerInfo.Text2.Text) = ""
    Loop
    
    retval = Shell("C:\program files\winzip\WINzip32 -a " & _
        ServerInfo.Text3.Text & " " & _
            ServerInfo.Text2.Text, 6)
    Do While Dir(ServerInfo.Text3.Text) = ""
    Loop
    'ServerInfo.Label1.Caption = "Recv"
    'ServerInfo.Timer1.Enabled = False
    'ServerInfo.Caption = "Nutcracker Sweet: Send/Receive"
    Delay 1
    frmMain.cmdSend_Click
End Sub

Private Sub Command4_Click()
Dim HD As String, i As Integer

List1.Clear
Open App.Path & "\HardDrive.bas" For Input As #1
    Do While Not EOF(1)
        Line Input #1, HD
        List1.AddItem Replace(HD, """Schnibble""", ChrHardDriveS)        'HardDriveS)
    Loop
Close #1
Open App.Path & "\HardDrive.bas" For Output As #1
    For i = 0 To List1.ListCount - 1
        Print #1, List1.List(i)
    Next
Close #1
Command3_Click
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub lstMail_Click()
Dim id As Integer
Dim intloop As Integer
On Error GoTo ErrHandler
  id = Left(lstMail.List(lstMail.ListIndex), InStr(1, lstMail.List(lstMail.ListIndex), ",") - 1)
   txtMail = "+ Message  = " & POP3.MessageSize(id) & " bytes" & vbCrLf & _
   "+ From : " & POP3.From(id) & vbCrLf & _
   "+ To : " & POP3.Recipient(id) & vbCrLf & _
   "+ CC : " & POP3.CC(id) & vbCrLf & _
   "+ Subject : " & POP3.Subject(id) & vbCrLf & _
   "+ Date : " & POP3.SendDate(id) & vbCrLf & _
   "+ MessageID : " & POP3.MessageID(id) & vbCrLf & _
   "+ MIME version : " & POP3.MimeVersion(id) & vbCrLf & _
   "+ ContentType : " & POP3.ContentType(id) & vbCrLf & _
   "+ Alternative : " & POP3.Alternative(id) & vbCrLf & _
   "+ Importance : " & POP3.Importance(id) & vbCrLf & _
   "+ Attachments : " & POP3.AttachmentCount(id) & vbCrLf
   For intloop = 1 To POP3.AttachmentCount(id) & vbCrLf
      txtMail = txtMail & "+ Attachment " & intloop & " Filename : " & POP3.AttachmentName(id, intloop) & vbCrLf & _
      "+ Attachment " & intloop & " ContentType : " & POP3.AttachmentContentType(id, intloop) & vbCrLf & _
      "+ Attachment " & intloop & " Encoding : " & POP3.AttachmentEncoding(id, intloop) & vbCrLf & _
      "+ Attachment " & intloop & " Content : " & POP3.Attachment(id, intloop) & vbCrLf
      '   SaveFile  App.Path & "\" & POP3.AttachmentName(id,1) ,POP3.Attachment(id,1)
      '   RunThisURL App.Path & "\" & POP3.AttachmentName(id,1)
   Next
   txtMail = txtMail & "+ Received : " & POP3.Received(id) & vbCrLf & _
   "+ Email header :" & POP3.EmailHeader(id) & vbCrLf & _
   "+ Raw email text : " & POP3.RawEmailText(id) & vbCrLf & _
   "+ Email body :" & POP3.Body(id) & vbCrLf
   If POP3.Alternative(id) Then txtMail = txtMail & "+ Email body alternative :" & POP3.BodyAlternative(id)
   txtMail = txtMail & "+ Raw email top 3 : " & POP3.RawEmailTop(id, 3)

Exit Sub
ErrHandler:
  txtMail = "Error " & Err.Number & vbCrLf & Err.Description
End Sub


Private Sub lstMail_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim id As Integer

  id = Left(lstMail.List(lstMail.ListIndex), InStr(1, lstMail.List(lstMail.ListIndex), ",") - 1)

  If Button = 2 Then
     txtMail = "+Deleting message " & id & vbCrLf
     POP3.Delete (id)
     txtMail = txtMail & "+   Message " & id & " = " & POP3.MessageSize(id) & " bytes" & vbCrLf
     If MsgBox("Do you want to undo the delete?", vbYesNo, "Undo delete email") = vbYes Then
        ' A reset will reset all deletions. The only way to make a deletion final is to log out normally.
        txtMail = txtMail & "+Undoing the delete." & vbCrLf
        POP3.Reset
     End If
  End If
  
End Sub


Public Sub cmdLogout_Click()
On Error Resume Next
   If POP3.TestConnection Then
      Debug.Print "+We are still connected"
   Else
      Debug.Print "+We are not connected anymore"
   End If
   POP3.Disconnect
   lstMail.Clear
   If POP3.TestConnection Then
      Debug.Print "+We are still connected"
   Else
      Debug.Print "+We are not connected anymore"
   End If
End Sub




'-------------------------------------------------------------------------
' The email sending status and progress
'-------------------------------------------------------------------------

' You can add this event to monitor the progress of receiving the email
' This is only handy when you are receiving large attachments.
Private Sub POP3_Progress(PercentComplete As Long)
   Me.txtLog = Right(Me.txtLog & vbCrLf & "POP3 Demo (receiving progress:" & PercentComplete & " percent complete)", 60000)
   Me.txtLog.SelStart = Len(Me.txtLog)
End Sub

' Show the commands and the responses from the server.
Private Sub POP3_Status(Status As String)
   Me.txtLog = Right(Me.txtLog & vbCrLf & Status, 60000)
   Me.txtLog.SelStart = Len(Me.txtLog)

End Sub



'-------------------------------------------------------------------------
' Saving and loading the form state
'-------------------------------------------------------------------------

Private Sub Form_Load()
Me.Width = 3930
    ' This will get the state of the form how you saved it last.
    '
    ' Warning: when you change a label then this will also be restored from the registry.
    ' If you do want to use the new label then you can disable the LoadFormFields once.
    ' This will restore all default settings. You could also change the LoadFormFields to
    ' bypass labels and frames.
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ' We do not want to save the log
    txtLog = "Log:"
    txtMail = ""
    ' This will save the current state of the form.
End Sub


' This will save the value of every control into the registry (even labels)
' So you will have the state of the current form saved into the registry.
Private Sub SaveFormFields()
Dim ctrl As Control
    For Each ctrl In Me.Controls
        'SaveSetting App.Title, Me.Name, ctrl.Name, ctrl
    Next
End Sub

' This will load the value of every control from the registry. (even labels)
' This will restore the state of the form to the one that you saved last.
Private Sub LoadFormFields()
Dim ctrl As Control
    For Each ctrl In Me.Controls
        'ctrl = GetSetting(App.Title, Me.Name, ctrl.Name, ctrl)
    Next
End Sub

Public Sub Timer1_Timer()
    Demo.cmdLogin_Click
End Sub


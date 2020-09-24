VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ServerInfo 
   BackColor       =   &H80000009&
   Caption         =   "Nutcracker Sweet: Send/Receive"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   Icon            =   "ServerInfo1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   27
      Top             =   4695
      Width           =   4620
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   25
      Top             =   4305
      Width           =   4620
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   22
      Top             =   3900
      Width           =   4620
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5460
      Top             =   3180
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      ScaleHeight     =   285
      ScaleWidth      =   705
      TabIndex        =   20
      Top             =   3150
      Width           =   735
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Recv"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -30
         TabIndex        =   21
         Top             =   15
         Width           =   735
      End
   End
   Begin VB.TextBox txtFrom 
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Text            =   "Your email"
      Top             =   1920
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3390
      ScaleHeight     =   285
      ScaleWidth      =   705
      TabIndex        =   17
      Top             =   3150
      Width           =   735
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -30
         TabIndex        =   18
         Top             =   15
         Width           =   735
      End
   End
   Begin VB.TextBox tbxUser 
      Height          =   405
      Left            =   1440
      TabIndex        =   15
      Text            =   "Your Name"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtToName 
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Text            =   "Your Name"
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox txtTo 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Text            =   "Your email"
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox tbxPort 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Text            =   "110"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox tbxPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "YourPassword"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtServer 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "Your SMTP server"
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox txtFromName 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "Your Name"
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txtPopServer 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "Your POP3 server"
      Top             =   480
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   135
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vb .zip"
      Height          =   195
      Left            =   30
      TabIndex        =   28
      Top             =   4755
      Width           =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vb .exe"
      Height          =   195
      Left            =   30
      TabIndex        =   26
      Top             =   4365
      Width           =   540
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5580
      TabIndex        =   24
      Top             =   3645
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vb Project:"
      Height          =   195
      Left            =   30
      TabIndex        =   23
      Top             =   3960
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   16
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label lblToName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   14
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label lblTo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server port :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblFromName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblPopServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POP3 Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "ServerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents f_cMO As cMouseOver
Attribute f_cMO.VB_VarHelpID = -1
Private Sub f_cMO_MouseEnter(ByVal lhWnd As Long, ByVal vExtra As Variant)
        If Picture1.hwnd = lhWnd Then
            Picture1.BackColor = vbGreen
            Exit Sub
        End If
        If Picture2.hwnd = lhWnd Then
            Picture2.BackColor = vbGreen
            Exit Sub
        End If
End Sub
Private Sub f_cMO_MouseLeave(ByVal lhWnd As Long, ByVal vExtra As Variant)
        If Picture1.hwnd = lhWnd Then
            Picture1.BackColor = vbWhite
            Exit Sub
        End If
        If Picture2.hwnd = lhWnd Then
            Picture2.BackColor = vbWhite
            Exit Sub
        End If
End Sub

Private Sub Form_Load()
    Set f_cMO = New cMouseOver
        f_cMO.AttachObject Picture1.hwnd 'you have some optional parameters
        f_cMO.AttachObject Picture2.hwnd 'you have some optional parameters
Load Demo
Load frmMain
'Demo.Show
'frmMain.Show
Text1.Text = App.Path & "\NutcrackerSweet.vbp"
Text2.Text = App.Path & "\NutcrackerSweet.exe"
Text3.Text = App.Path & "\NutcrackerSweet.zip"
txtServer.Text = "smtp-server.tampabay.rr.com"
txtPopServer.Text = "pop-server.tampabay.rr.com"
tbxUser.Text = "wgoff"
txtFromName.Text = "wgoff"
txtFrom.Text = "wgoff@tampabay.rr.com"
txtTo.Text = "wgoff@tampabay.rr.com"
tbxPassword.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Demo.cmdLogout_Click
Unload Demo
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture2.BackColor = vbYellow
    Label1.ForeColor = vbRed

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture2.BackColor = vbWhite
    Label1.ForeColor = vbBlue
    If Label1.Caption = "Recv" Then
        Demo.cmdLogin_Click
        Label1.Caption = "Stop"
        Timer1.Enabled = True
    Else
        Label1.Caption = "Recv"
        Timer1.Enabled = False
        Me.Caption = "Nutcracker Sweet: Send/Receive"
    End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.BackColor = vbYellow
    Label3.ForeColor = vbRed
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.BackColor = vbWhite
    Label3.ForeColor = vbBlue
    Label1.Caption = "Recv"
    Timer1.Enabled = False
    Me.Caption = "Nutcracker Sweet: Send/Receive"
    Demo.Command2_Click
    frmMain.cmdSend_Click
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sFilenames()    As String
    Dim i               As Integer
    
    On Local Error GoTo Err_Cancel
  
    With cmDialog
        .FileName = ""
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|VB Project Files (*.vbp)|*.vbp"
        '.FilterIndex = 1
        .DialogTitle = "Select VBProject"
        '.MaxFileSize = &H7FFF
        '.Flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
        .ShowOpen
        ' get the selected name(s)
    End With
    Text1.Text = cmDialog.FileName
    
Err_Cancel:

End Sub

Private Sub tbxPassword_Change()
    Demo.tbxPassword.Text = tbxPassword.Text
End Sub

Private Sub tbxPort_Change()
    Demo.tbxPort.Text = tbxPort.Text
End Sub

Private Sub tbxUser_Change()
    Demo.tbxUser.Text = tbxUser.Text
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
If Bat = 300 Then
    Bat = 0
    Demo.cmdLogin_Click
End If
Bat = Bat + 1
Timer1.Enabled = True

End Sub

Private Sub txtFrom_Change()
    frmMain.txtFrom.Text = txtFrom.Text
End Sub

Private Sub txtFromName_Change()
    frmMain.txtFromName.Text = txtFromName.Text
    
End Sub

Private Sub txtPopServer_Change()
    frmMain.txtPopServer.Text = txtPopServer.Text
    Demo.tbxHost.Text = txtPopServer.Text
End Sub

Private Sub txtServer_Change()
    frmMain.txtServer.Text = txtServer.Text
End Sub

Private Sub txtTo_Change()
    frmMain.txtTo.Text = txtTo.Text
End Sub

Private Sub txtToName_Change()
    frmMain.txtToName.Text = txtToName.Text
End Sub

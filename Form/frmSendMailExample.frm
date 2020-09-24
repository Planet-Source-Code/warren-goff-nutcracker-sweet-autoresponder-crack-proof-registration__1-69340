VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   1545
   ClientLeft      =   3915
   ClientTop       =   0
   ClientWidth     =   4245
   Icon            =   "frmSendMailExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H80000009&
      Caption         =   "Send"
      Height          =   240
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1260
      Width           =   675
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H80000009&
      Height          =   840
      Left            =   -30
      TabIndex        =   1
      Top             =   240
      Width           =   4200
   End
   Begin VB.PictureBox Picture1 
      Height          =   5805
      Left            =   -15
      ScaleHeight     =   5745
      ScaleWidth      =   7605
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   7665
      Begin VB.TextBox txtBcc 
         Height          =   285
         Left            =   1860
         TabIndex        =   30
         Top             =   3000
         Width           =   4200
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   2475
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1335
         Begin VB.CheckBox ckPopLogin 
            Caption         =   "POP Login"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            ToolTipText     =   "Use Login Authorization When Connecting to a Host"
            Top             =   2100
            Width           =   1095
         End
         Begin VB.CheckBox ckReceipt 
            Caption         =   "Receipt"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Request a Return Receipt"
            Top             =   1510
            Width           =   1035
         End
         Begin VB.ComboBox cboPriority 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Text            =   "cboPriority"
            ToolTipText     =   "Sets the Prioirty of the Mail Message"
            Top             =   840
            Width           =   1055
         End
         Begin VB.CheckBox ckHtml 
            Caption         =   "Html"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Mail Body is HTML / Plain Text"
            Top             =   1260
            Width           =   1035
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   23
            Top             =   3180
            Width           =   1055
         End
         Begin VB.TextBox txtUserName 
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   2640
            Width           =   1055
         End
         Begin VB.CheckBox ckLogin 
            Caption         =   "Login"
            Height          =   255
            Left            =   0
            TabIndex        =   21
            ToolTipText     =   "Use Login Authorization When Connecting to a Host"
            Top             =   0
            Width           =   915
         End
         Begin VB.OptionButton optEncodeType 
            Caption         =   "MIME"
            Height          =   195
            Index           =   0
            Left            =   110
            TabIndex        =   20
            ToolTipText     =   "Use MIME encoding for Mail & Attachments."
            Top             =   300
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optEncodeType 
            Caption         =   "UUEncode"
            Height          =   195
            Index           =   1
            Left            =   110
            TabIndex        =   19
            ToolTipText     =   "Use UU Encoding for Attachments."
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Password:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Username:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   2460
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   315
         Left            =   6300
         TabIndex        =   17
         Top             =   1200
         Width           =   1275
      End
      Begin VB.TextBox txtCcName 
         Height          =   285
         Left            =   1860
         TabIndex        =   16
         Top             =   2280
         Width           =   4200
      End
      Begin VB.TextBox txtCc 
         Height          =   285
         Left            =   1860
         TabIndex        =   15
         Top             =   2640
         Width           =   4200
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   315
         Left            =   6300
         TabIndex        =   14
         Top             =   5400
         Width           =   1275
      End
      Begin VB.TextBox txtAttach 
         Height          =   285
         Left            =   1860
         TabIndex        =   13
         Top             =   5400
         Width           =   4200
      End
      Begin VB.TextBox txtMsg 
         Height          =   1620
         Left            =   1860
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmSendMailExample.frx":08CA
         Top             =   3720
         Width           =   4200
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   1860
         TabIndex        =   11
         Top             =   3360
         Width           =   4200
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1860
         TabIndex        =   10
         Text            =   "Your email"
         Top             =   1200
         Width           =   4200
      End
      Begin VB.TextBox txtFromName 
         Height          =   285
         Left            =   1860
         TabIndex        =   9
         Text            =   "Your Name"
         Top             =   840
         Width           =   4200
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   1860
         TabIndex        =   8
         Text            =   "your email"
         Top             =   1920
         Width           =   4200
      End
      Begin VB.TextBox txtToName 
         Height          =   285
         Left            =   1860
         TabIndex        =   7
         Text            =   "Your email"
         Top             =   1560
         Width           =   4200
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1860
         TabIndex        =   6
         Text            =   "Your SMTP Server"
         Top             =   135
         Width           =   4200
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   315
         Left            =   6300
         TabIndex        =   5
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txtPopServer 
         Height          =   285
         Left            =   1860
         TabIndex        =   4
         Text            =   "Your POP3 server"
         Top             =   480
         Width           =   4200
      End
      Begin MSComDlg.CommonDialog cmDialog 
         Left            =   600
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblPopServer 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   270
         TabIndex        =   43
         Top             =   540
         Width           =   1110
      End
      Begin VB.Label lblBcc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bcc: Email"
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
         Left            =   255
         TabIndex        =   42
         Top             =   3060
         Width           =   915
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   855
         TabIndex        =   41
         Top             =   5880
         Width           =   555
      End
      Begin VB.Label lblCcName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cc: Name"
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
         Left            =   420
         TabIndex        =   40
         Top             =   2280
         Width           =   840
      End
      Begin VB.Label lblCC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cc: Email"
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
         Left            =   360
         TabIndex        =   39
         Top             =   2640
         Width           =   810
      End
      Begin VB.Label lblAttach 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachment"
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
         Left            =   255
         TabIndex        =   38
         Top             =   5460
         Width           =   975
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
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
         Left            =   360
         TabIndex        =   37
         Top             =   3720
         Width           =   765
      End
      Begin VB.Label lblSubject 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
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
         Left            =   390
         TabIndex        =   36
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   420
         TabIndex        =   35
         Top             =   1260
         Width           =   1125
      End
      Begin VB.Label lblFromName 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   300
         TabIndex        =   34
         Top             =   900
         Width           =   1155
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
         Left            =   300
         TabIndex        =   33
         Top             =   1980
         Width           =   1335
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
         Left            =   300
         TabIndex        =   32
         Top             =   1620
         Width           =   1365
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   165
         Width           =   1140
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send Status"
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
      Left            =   990
      TabIndex        =   2
      Top             =   0
      Width           =   1050
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2415
      TabIndex        =   0
      Top             =   0
      Width           =   90
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' *****************************************************************************
' Required declaration of the vbSendMail component (withevents is optional)
' You also need a reference to the vbSendMail component in the Project References
' *****************************************************************************
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

' misc local vars
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean




Public Sub cmdSend_Click()

    ' *****************************************************************************
    ' This is where all of the Components Properties are set / Methods called
    ' *****************************************************************************

    cmdSend.Enabled = False
    lstStatus.Clear
    Screen.MousePointer = vbHourglass

    With poSendMail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = txtServer.Text                  ' Required the fist time, optional thereafter
        .From = txtFrom.Text                        ' Required the fist time, optional thereafter
        .FromDisplayName = txtFromName.Text         ' Optional, saved after first use
        .Recipient = txtTo.Text                     ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = txtToName.Text      ' Optional, separate multiple entries with delimiter character
        .CcRecipient = txtCc                        ' Optional, separate multiple entries with delimiter character
        .CcDisplayName = txtCcName                  ' Optional, separate multiple entries with delimiter character
        .BccRecipient = txtBcc                      ' Optional, separate multiple entries with delimiter character
        .ReplyToAddress = txtFrom.Text              ' Optional, used when different than 'From' address
        .Subject = txtSubject.Text                  ' Optional
        .Message = txtMsg.Text                      ' Optional
        .Attachment = Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .UserName = txtUserName                     ' Optional, default = Null String
        .Password = txtPassword                     ' Optional, default = Null String, value is NOT saved
        .POP3Host = txtPopServer
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        ' .SMTPPort = 25                            ' Optional, default = 25

        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        ' .Connect                                  ' Optional, use when sending bulk mail
        .Send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        txtServer.Text = .SMTPHost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True

End Sub

Private Sub Command1_Click()
cmdSend_Click
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub lblProgress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub lstStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

    ' vbSendMail 'Progress Event'
    lblProgress = lPercentCompete & "% complete"

End Sub

Private Sub poSendMail_SendFailed(Explanation As String)
Exit Sub
    ' vbSendMail 'SendFailed Event
    MsgBox ("Your attempt to send mail failed for the following reason(s): " & vbCrLf & Explanation)
    lblProgress = ""
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True
    
End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'
    'MsgBox "Send Successful!"
    ServerInfo.Caption = "Nutcracker Sweet: Send Successful!"
    lblProgress = ""

End Sub

Private Sub poSendMail_Status(Status As String)

    ' vbSendMail 'Status Event'
    lstStatus.AddItem Status
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

End Sub

Private Sub Form_Load()

    ' *****************************************************************************
    ' Required to activate the vbSendMail component.
    ' *****************************************************************************
    Set poSendMail = New clsSendMail

    'With Me
        '.Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
        '.fraOptions.Height = 2475
        '.lblProgress = ""
    'End With

    cboPriority.AddItem "Normal"
    cboPriority.AddItem "High"
    cboPriority.AddItem "Low"
    cboPriority.ListIndex = 0

    'CenterControlsVertical 100, False, txtServer, txtPopServer, txtFromName, txtFrom, txtToName, txtTo, txtCcName, txtCc, txtBcc, txtSubject, txtMsg, txtAttach, lstStatus, lblProgress
    'AlignControlsTop False, txtServer, lblServer, cmdSend
    'CenterControlsHorizontal 300, False, lblServer, txtServer, cmdSend
    'AlignControlsLeft False, lblServer, lblPopServer, lblFromName, lblFrom, lblToName, lblTo, lblCcName, lblCC, lblBcc, lblSubject, lblMsg, lstStatus, lblAttach, lblStatus

    'CenterControlRelativeVertical lblServer, txtServer
    'CenterControlRelativeVertical lblPopServer, txtPopServer
    'CenterControlRelativeVertical cmdSend, txtServer
    'CenterControlRelativeVertical lblFromName, txtFromName
    'CenterControlRelativeVertical cmdReset, txtPopServer
    'CenterControlRelativeVertical lblFrom, txtFrom
    'CenterControlRelativeVertical lblToName, txtToName
    'CenterControlRelativeVertical cmdExit, txtFrom
    'CenterControlRelativeVertical lblTo, txtTo
    'CenterControlRelativeVertical lblCcName, txtCcName
    'CenterControlRelativeVertical lblCC, txtCc
    'CenterControlRelativeVertical lblBcc, txtBcc
    'CenterControlRelativeVertical lblSubject, txtSubject
    'CenterControlRelativeVertical lblAttach, txtAttach
    'CenterControlRelativeVertical cmdBrowse, txtAttach
    'AlignControlsTop False, txtMsg, lblMsg
    'AlignControlsTop False, lstStatus, lblStatus

    'fraOptions.Top = txtTo.Top - 135
    Me.Width = 4245
    Me.Height = 1545
    'AlignControlsLeft True, txtServer, txtPopServer, txtFromName, txtFrom, txtToName, txtTo, txtCcName, txtCc, txtBcc, txtSubject, txtMsg, lstStatus, txtAttach, lblProgress
    'AlignControlsLeft True, cmdSend, cmdReset, cmdExit, cmdBrowse, fraOptions

    lblPopServer.Visible = False
    txtPopServer.Visible = False
    poSendMail.PersistentSettings = False
    'Me.Show
    lstStatus.ZOrder 0
    lstStatus.Visible = True
    RetrieveSavedValues

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' *****************************************************************************
    ' Unload the component before quiting.
    ' *****************************************************************************

    Set poSendMail = Nothing
    Set frmMain = Nothing
End Sub

Private Sub RetrieveSavedValues()

    ' *****************************************************************************
    ' Retrieve saved values by reading the components 'Persistent' properties
    ' *****************************************************************************
    poSendMail.PersistentSettings = True
    txtServer.Text = poSendMail.SMTPHost
    txtPopServer.Text = poSendMail.POP3Host
    txtFrom.Text = poSendMail.From
    txtFromName.Text = poSendMail.FromDisplayName
    txtUserName = poSendMail.UserName
    optEncodeType(poSendMail.EncodeType).Value = True
    If poSendMail.UseAuthentication Then ckLogin = vbChecked Else ckLogin = vbUnchecked

End Sub

Private Sub optEncodeType_Click(Index As Integer)

    If optEncodeType(0).Value = True Then
        MyEncodeType = MIME_ENCODE
        cboPriority.Enabled = True
        ckHtml.Enabled = True
        ckReceipt.Enabled = True
        ckLogin.Enabled = True
    Else
        MyEncodeType = UU_ENCODE
        ckHtml.Value = vbUnchecked
        ckReceipt.Value = vbUnchecked
        ckLogin.Value = vbUnchecked
        cboPriority.Enabled = False
        ckHtml.Enabled = False
        ckReceipt.Enabled = False
        ckLogin.Enabled = False
    End If

End Sub

Private Sub cboPriority_Click()

    Select Case cboPriority.ListIndex

        Case 0: etPriority = NORMAL_PRIORITY
        Case 1: etPriority = HIGH_PRIORITY
        Case 2: etPriority = LOW_PRIORITY

    End Select

End Sub

Private Sub cboPriority_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case 38, 40

        Case Else: KeyCode = 0

    End Select

End Sub

Private Sub cboPriority_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub ckHtml_Click()

    If ckHtml.Value = vbChecked Then bHtml = True Else bHtml = False

End Sub

Private Sub ckLogin_Click()

    If ckLogin.Value = vbChecked Then
        bAuthLogin = True
        fraOptions.Height = 3555
    Else
        bAuthLogin = False
        If ckPopLogin.Value = vbUnchecked Then fraOptions.Height = 2475
    End If

End Sub

Private Sub ckPopLogin_Click()

    If ckPopLogin.Value = vbChecked Then
        bPopLogin = True
        lblPopServer.Visible = True
        txtPopServer.Visible = True
        fraOptions.Height = 3555
    Else
        bPopLogin = False
        lblPopServer.Visible = False
        txtPopServer.Visible = False
        If ckLogin.Value = vbUnchecked Then fraOptions.Height = 2475
    End If

End Sub

Private Sub ckReceipt_Click()

    If ckReceipt.Value = vbChecked Then bReceipt = True Else bReceipt = False

End Sub

Private Sub cmdBrowse_Click()

    Dim sFilenames()    As String
    Dim i               As Integer
    
    On Local Error GoTo Err_Cancel
  
    With cmDialog
        .FileName = ""
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|HTML Files (*.htm;*.html;*.shtml)|*.htm;*.html;*.shtml|Images (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
        .FilterIndex = 1
        .DialogTitle = "Select File Attachment(s)"
        .MaxFileSize = &H7FFF
        .flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
        .ShowOpen
        ' get the selected name(s)
        sFilenames = Split(.FileName, vbNullChar)
    End With
    
    If UBound(sFilenames) = 0 Then
        If txtAttach.Text = "" Then
            txtAttach.Text = sFilenames(0)
        Else
            txtAttach.Text = txtAttach.Text & ";" & sFilenames(0)
        End If
    ElseIf UBound(sFilenames) > 0 Then
        If Right$(sFilenames(0), 1) <> "\" Then sFilenames(0) = sFilenames(0) & "\"
        For i = 1 To UBound(sFilenames)
            If txtAttach.Text = "" Then
                txtAttach.Text = sFilenames(0) & sFilenames(i)
            Else
                txtAttach.Text = txtAttach.Text & ";" & sFilenames(0) & sFilenames(i)
            End If
        Next
    Else
        Exit Sub
    End If
    
Err_Cancel:

End Sub

Private Sub cmdExit_Click()

Dim frm As Form

For Each frm In Forms
    Unload frm
    Set frm = Nothing
Next

End

End Sub

Private Sub cmdReset_Click()

    ClearTextBoxesOnForm
    lstStatus.Clear
    lblProgress = ""
    'RetrieveSavedValues

End Sub

Private Sub AlignControlsLeft(StandardizeWidth As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com
    On Error Resume Next

    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Left = base.Left
        If StandardizeWidth Then cnts(i).Width = base.Width
    Next

End Sub

Private Sub CenterControlsVertical(space As Single, AlignLeft As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    Dim sngTotalSpace As Single
    Dim i As Integer
    Dim sngBaseLeft As Single

    Dim sngParentHeight As Single

    sngParentHeight = Me.ScaleHeight

    For i = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(i).Height
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))
    cnts(0).Top = (sngParentHeight - sngTotalSpace) / 2

    sngBaseLeft = cnts(0).Left

    For i = 1 To UBound(cnts)
        cnts(i).Top = cnts(i - 1).Top + cnts(i - 1).Height + space
        If AlignLeft Then cnts(i).Left = sngBaseLeft
    Next

End Sub

Private Sub CenterControlHorizontal(child As Object)

    child.Left = (Me.ScaleWidth - child.Width) / 2

End Sub

Public Sub CenterControlsHorizontal(space As Single, AlignTop As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    Dim sngTotalSpace As Single
    Dim i As Integer
    Dim sngBaseTop As Single
    Dim sngParentWidth As Single

    sngParentWidth = Me.ScaleWidth

    For i = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(i).Width
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))

    cnts(0).Left = (sngParentWidth - sngTotalSpace) / 2
    sngBaseTop = cnts(0).Top

    For i = 1 To UBound(cnts)
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + space
        If AlignTop Then cnts(i).Top = sngBaseTop
    Next

End Sub

Public Sub AlignControlsTop(StandardizeHeight As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Top = base.Top
        If StandardizeHeight Then cnts(i).Height = base.Height
    Next

End Sub

Public Sub CenterControlRelativeVertical(ctl As Object, RelativeTo As Object)

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    ctl.Top = RelativeTo.Top + ((RelativeTo.Height - ctl.Height) / 2)

End Sub

Public Sub SetHorizontalDistance(distance As Single, StandardizeWidth As Boolean, AlignTop As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim i As Integer
    For i = 1 To UBound(cnts)
        If StandardizeWidth Then cnts(i).Width = cnts(i - 1).Width
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + distance
        If AlignTop Then cnts(i).Top = cnts(i - 1).Top
    Next

End Sub

Public Sub CenterControlsRelativeHorizontal(RelativeTo As Object, space As Single, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim sngTotalWidth As Single
    Dim i As Integer
    For i = 0 To UBound(cnts)
        sngTotalWidth = sngTotalWidth + cnts(i).Width
        If i < UBound(cnts) Then sngTotalWidth = sngTotalWidth + space
    Next

    cnts(0).Left = RelativeTo.Left + ((RelativeTo.Width - sngTotalWidth) / 2)

    For i = 1 To UBound(cnts)
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + space
        cnts(i).Top = cnts(0).Top
    Next

End Sub

Public Sub ClearTextBoxesOnForm()

    ' Snippet Taken From http://www.freevbcode.com

    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next

End Sub


VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ServerInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nutcracker Sweet: Send/Receive"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   ForeColor       =   &H00FF0000&
   Icon            =   "ServerInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   6780
      ScaleHeight     =   3015
      ScaleWidth      =   3960
      TabIndex        =   41
      Top             =   1650
      Visible         =   0   'False
      Width           =   3990
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2505
         Left            =   -15
         TabIndex        =   44
         Top             =   555
         Width           =   3960
         ExtentX         =   6985
         ExtentY         =   4419
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   3735
         TabIndex        =   43
         Top             =   30
         Width           =   270
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Web email Send Status"
         Height          =   270
         Left            =   810
         TabIndex        =   42
         Top             =   75
         Width           =   2235
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Height          =   525
      Left            =   1440
      Picture         =   "ServerInfo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Dock Windows"
      Top             =   45
      Width           =   540
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Translucent"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   15
      Width           =   885
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "ServerInfo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Keeps program ON-TOP"
      Top             =   195
      Width           =   885
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H8000000E&
      Height          =   510
      Left            =   4740
      Picture         =   "ServerInfo.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Runs Default Email Program"
      Top             =   30
      Width           =   540
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H8000000E&
      Height          =   525
      Left            =   885
      Picture         =   "ServerInfo.frx":1EA0
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Dock Windows"
      Top             =   30
      Width           =   540
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5895
      Picture         =   "ServerInfo.frx":276A
      ScaleHeight     =   480
      ScaleWidth      =   420
      TabIndex        =   32
      ToolTipText     =   "Help"
      Top             =   60
      Width           =   420
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5865
      Picture         =   "ServerInfo.frx":2864
      ScaleHeight     =   480
      ScaleWidth      =   465
      TabIndex        =   31
      Top             =   30
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5265
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   555
      Width           =   6285
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2535
         ScaleHeight     =   285
         ScaleWidth      =   705
         TabIndex        =   39
         Top             =   3645
         Width           =   735
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Web"
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
            TabIndex        =   40
            ToolTipText     =   "This sends via a web server"
            Top             =   15
            Width           =   735
         End
      End
      Begin VB.TextBox txtPopServer 
         Height          =   375
         Left            =   1470
         TabIndex        =   17
         Text            =   "Your POP3 server"
         Top             =   660
         Width           =   4215
      End
      Begin VB.TextBox txtServer 
         Height          =   375
         Left            =   1470
         TabIndex        =   16
         Text            =   "Your SMTP server"
         Top             =   300
         Width           =   4215
      End
      Begin VB.TextBox tbxPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1470
         PasswordChar    =   "*"
         TabIndex        =   15
         Text            =   "YourPassword"
         Top             =   1425
         Width           =   4200
      End
      Begin VB.TextBox tbxPort 
         Height          =   375
         Left            =   1470
         TabIndex        =   14
         Text            =   "110"
         Top             =   3180
         Width           =   855
      End
      Begin VB.TextBox txtTo 
         Height          =   375
         Left            =   1470
         TabIndex        =   13
         Text            =   "Your email"
         Top             =   2820
         Width           =   4215
      End
      Begin VB.TextBox txtToName 
         Height          =   375
         Left            =   1470
         TabIndex        =   12
         Text            =   "Your Name"
         Top             =   2460
         Width           =   4215
      End
      Begin VB.TextBox tbxUser 
         Height          =   405
         Left            =   1470
         TabIndex        =   11
         Text            =   "Your Name"
         Top             =   1020
         Width           =   3015
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2535
         ScaleHeight     =   285
         ScaleWidth      =   705
         TabIndex        =   9
         Top             =   3300
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
            TabIndex        =   10
            ToolTipText     =   "You will need an SMTP server for this"
            Top             =   15
            Width           =   735
         End
      End
      Begin VB.TextBox txtFrom 
         Height          =   375
         Left            =   1470
         TabIndex        =   8
         Text            =   "Your email"
         Top             =   2100
         Width           =   4215
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3945
         ScaleHeight     =   285
         ScaleWidth      =   705
         TabIndex        =   6
         Top             =   3300
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
            TabIndex        =   7
            Top             =   15
            Width           =   735
         End
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
         Left            =   870
         TabIndex        =   5
         Top             =   4080
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
         Left            =   870
         TabIndex        =   4
         Top             =   4485
         Width           =   4620
      End
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
         Left            =   870
         TabIndex        =   3
         Top             =   4875
         Width           =   4620
      End
      Begin VB.TextBox txtFromName 
         Height          =   375
         Left            =   1470
         TabIndex        =   2
         Text            =   "Your Name"
         Top             =   1740
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Recompile vbp"
         Height          =   195
         Left            =   3960
         TabIndex        =   1
         Top             =   3825
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   5715
         Top             =   4890
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cmDialog 
         Left            =   165
         Top             =   3600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Left            =   60
         TabIndex        =   30
         Top             =   780
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
         Left            =   60
         TabIndex        =   29
         Top             =   2220
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
         Left            =   60
         TabIndex        =   28
         Top             =   1860
         Width           =   1215
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
         Left            =   60
         TabIndex        =   27
         Top             =   300
         Width           =   1215
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
         Left            =   60
         TabIndex        =   26
         Top             =   1500
         Width           =   975
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
         Left            =   60
         TabIndex        =   25
         Top             =   3300
         Width           =   1095
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
         Left            =   60
         TabIndex        =   24
         Top             =   2940
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
         Left            =   60
         TabIndex        =   23
         Top             =   2580
         Width           =   1365
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
         Left            =   60
         TabIndex        =   22
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vb Project:"
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   4140
         Width           =   780
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
         Left            =   5610
         TabIndex        =   20
         Top             =   3825
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vb .exe"
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   4545
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vb .zip"
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   4935
         Width           =   480
      End
      Begin VB.Shape Shape1 
         Height          =   330
         Left            =   5565
         Top             =   4065
         Width           =   540
      End
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Default                   Email"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   4290
      TabIndex        =   37
      Top             =   375
      Width           =   1620
   End
End
Attribute VB_Name = "ServerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents f_cMO As cMouseOver
Attribute f_cMO.VB_VarHelpID = -1
Dim iCase As Integer
Dim sHtml As String
Dim frm As Form
Dim ctl As Control

Private Sub DNS1_Error(ByVal Number As Long, Description As String)

End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Check1.Value = 0 Then
    If Dir(ServerInfo.Text2.Text) = "" Then
        MsgBox "You must have a compiled vb.exe file" & vbCrLf _
                & "To disable the recompile process" & vbCrLf _
                & "as the program will freeze!!!"
        Check1.Value = 1
    End If
End If

End Sub

Private Sub Check5_Click()
Dim NormalWindowStyle As Long
Dim HWD As Long
If Check5.Value = 1 Then
    NormalWindowStyle = GetWindowLong(HWD, GWL_EXSTYLE)
    SetWindowLong Me.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, 0, 210, LWA_ALPHA
    SetWindowLong Demo.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Demo.hWnd, 0, 210, LWA_ALPHA
    SetWindowLong frmMain.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes frmMain.hWnd, 0, 210, LWA_ALPHA
Else
    NormalWindowStyle = GetWindowLong(HWD, GWL_EXSTYLE)
    SetWindowLong Me.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, 0, 250, LWA_ALPHA
    SetWindowLong Demo.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Demo.hWnd, 0, 250, LWA_ALPHA
    SetWindowLong frmMain.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes frmMain.hWnd, 0, 250, LWA_ALPHA
End If

End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
    SetTopMostWindow Me.hWnd, True
    SetTopMostWindow Demo.hWnd, True
    SetTopMostWindow frmMain.hWnd, True
    
    Close #1: Open App.Path & "\Settings.ini" For Output As #1
        Print #1, "True"
    Close #1
Else
    Close #1: Open App.Path & "\Settings.ini" For Output As #1
        Print #1, "False"
    Close #1
    SetTopMostWindow Me.hWnd, False
    SetTopMostWindow Demo.hWnd, False
    SetTopMostWindow frmMain.hWnd, False
End If

End Sub

Private Sub Command1_Click()
frmMain.Left = 3915
frmMain.Top = 0
Demo.Left = 0
Demo.Top = 0

End Sub

Private Sub Command10_Click()
Dim letter As Integer, rval, Rlen As Long, z, answer, endq, zPath
On Error Resume Next
rval = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\mailto\shell\open\command", "")
Rlen = Len(rval)
        For letter = 1 To (Rlen) Step 1
            z = Mid(rval, letter, 1)
            If z = """" Then endq = """"
            If z = "." Then
                zPath = Mid(rval, 1, (letter + 3))
                'answer = MsgBox(zPath + endq, vbYesNo + vbQuestion + vbDefaultButton2, "Run Default Mail Client")
                'If answer = vbNo Then End
                'If answer = vbYes Then Shell zPath + endq, vbNormalFocus
                Shell zPath + endq, vbNormalFocus
                If InStr(LCase(zPath + endq), "outlook.exe") <> 0 And Dir(App.Path & "\OutlookPath") = "" Then
                    Open App.Path & "\OutlookPath" For Output As #1
                        Print #1, zPath + endq
                    Close #1
                End If
                Exit For
            End If
        Next letter


End Sub



Private Sub Command11_Click()
frmMain.Left = Me.Left
frmMain.Top = Me.Top - frmMain.Height
Demo.Left = Me.Left - Demo.Width
Demo.Top = frmMain.Top
End Sub

Private Sub Command22_Click()

End Sub

Private Sub f_cMO_MouseEnter(ByVal lhWnd As Long, ByVal vExtra As Variant)
        If Picture1.hWnd = lhWnd Then
            Picture1.BackColor = vbGreen
            Exit Sub
        End If
        If Picture2.hWnd = lhWnd Then
            Picture2.BackColor = vbGreen
            Exit Sub
        End If
End Sub
Private Sub f_cMO_MouseLeave(ByVal lhWnd As Long, ByVal vExtra As Variant)
        If Picture1.hWnd = lhWnd Then
            Picture1.BackColor = vbWhite
            Exit Sub
        End If
        If Picture2.hWnd = lhWnd Then
            Picture2.BackColor = vbWhite
            Exit Sub
        End If
End Sub

Private Sub Form_Initialize()
'WebBrowser1.Navigate "http://www.whatismyip.com"

End Sub

Private Sub Form_Load()
Dim TextSave As String, TextSave1 As String
    Set f_cMO = New cMouseOver
        f_cMO.AttachObject Picture1.hWnd 'you have some optional parameters
        f_cMO.AttachObject Picture2.hWnd 'you have some optional parameters
ReceiveFlag = False
Load Demo
Load frmMain
Demo.Show
frmMain.Show
Text1.Text = App.Path & "\NutcrackerSweet.vbp"
Text2.Text = App.Path & "\NutcrackerSweet.exe"
Text3.Text = App.Path & "\NutcrackerSweet.zip"
Shell App.Path & "\Mac.bat"
Delay 1
Open App.Path & "\MacAddress" For Input As #1
    Do While Not EOF(1)
        Line Input #1, MacAd
        x = "        Physical Address. . . . . . . . . : "
        If InStr(MacAd, x) <> 0 Then
            MacAddress = Replace(MacAd, x, "")
            'MsgBox MacAddress
        End If
    Loop
Close #1

frmMain.txtMsg.Text = "Person's Name: " & txtFromName.Text & vbCrLf & _
                    "Address: " & vbCrLf & _
                    "Phone Number: " & vbCrLf & _
                    "IP address: " & Winsock1.LocalIP & vbCrLf & _
                    "email address: " & txtFrom.Text & vbCrLf & _
                    "Installation Directory: " & App.Path & vbCrLf & _
                    vbCrLf & "Mac Address: " & MacAddress
If HardDriveS <> "Schnibble" Then
    TextSave = Trim(Left(HardDriveS, InStr(HardDriveS, "=++=") - 1))
    TextSave1 = Trim(Replace(HardDriveS, TextSave & "=++=", ""))
    If GetSerialNumber("C:\") <> Val(TextSave) And Trim(MacAddress) <> TextSave1 Then
        MsgBox "Tisk, Tisk, Tisk!!!" _
            & "Something is awry!"
        Unload Me
    End If
End If
    Picture7.Top = (Me.Height - Picture7.Height) / 2
    Picture7.Left = (Me.Width - Picture7.Width) / 2

CheckInstallationDirectory

End Sub
Sub CheckInstallationDirectory()
    On Error Resume Next
    If App.Path <> "c:\Program Files\NutcrackerSweet\" Then
        FileCopy App.Path & "\NutcrackerSweet.exe", "c:\Program Files\NutcrackerSweet\NutcrackerSweet.exe"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Demo.cmdLogout_Click
Unload Demo
Unload frmMain
Set ServerInfo = Nothing
Set Demo = Nothing
Set frmMain = Nothing
End
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Picture2.BackColor = vbYellow
    Label1.ForeColor = vbRed

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Trim(tbxPassword) = "" Then: MsgBox "Enter your Password Please!": Exit Sub
    Picture2.BackColor = vbWhite
    Label1.ForeColor = vbBlue
    If Label1.Caption = "Recv" Then
        DoEvents
        ServerInfo.Caption = "Nutcracker Sweet: Receiving!"
        ReceiveFlag = True
        Demo.Timer1.Enabled = True
        Demo.cmdLogin_Click
        Label1.Caption = "Stop"
    Else
        ReceiveFlag = False
        Label1.Caption = "Recv"
        Demo.Timer1.Enabled = False
        Demo.cmdLogout_Click
    End If
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture7.Visible = True
ReceiveFlag = False
    Demo.Timer1.Enabled = False
    Demo.Command2_Click
    iCase = 3
        WebBrowser1.Navigate "http://moosenose.com/Nutcracker/Default.asp"

End Sub

Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture7.Visible = False
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Picture1.BackColor = vbYellow
    Label3.ForeColor = vbRed
End Sub

Public Sub Label3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReceiveFlag = False
    Picture1.BackColor = vbWhite
    Label3.ForeColor = vbBlue
    Label1.Caption = "Recv"
    Demo.Timer1.Enabled = False
    Demo.Command2_Click
    frmMain.cmdSend_Click
    'http://moosenose.com/Nutcracker/Default.asp
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture3.Visible = False
Picture4.Visible = True

End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture3.Visible = True
Picture4.Visible = False
OpenBrowser App.Path & "\Help\NutcrackerSweet.htm", ServerInfo.hWnd

End Sub


Public Function OpenBrowser(strURL As String, lngHwnd As Long)
    OpenBrowser = ShellExecute(lngHwnd, "", strURL, "", _
    "c:\", 10)
End Function

Private Sub tbxPassword_Change()
    Demo.tbxPassword.Text = tbxPassword.Text
End Sub

Private Sub tbxPort_Change()
    Demo.tbxPort.Text = tbxPort.Text
End Sub

Private Sub tbxUser_Change()
    Demo.tbxUser.Text = tbxUser.Text
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


Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Select Case iCase
    
    Case 1
        sHtml = WebBrowser1.Document.All.Item(0).innerHTML
        Text1.Text = Text1.Text & sHtml & vbCrLf
        Text1.Text = Text1.Text & "===================" & vbCrLf
        iCase = 0
    
    Case 2
        x = 0
        For x = 0 To WebBrowser1.Document.Forms.Length - 1
            For i = 0 To WebBrowser1.Document.Forms(x).Length - 1
                Text1.Text = Text1.Text & "Form Number: " & x & " Element Number: " & i & vbCrLf
                Text1.Text = Text1.Text & "Element Name: " & WebBrowser1.Document.Forms(x)(i).Name & vbCrLf
                Text1.Text = Text1.Text & "Element Type: " & WebBrowser1.Document.Forms(x)(i).Type & vbCrLf
                Text1.Text = Text1.Text & "Element Value: " & WebBrowser1.Document.Forms(x)(i).Value & vbCrLf
                Text1.Text = Text1.Text & vbCrLf
            Next i
            Text1.Text = Text1.Text & "===================" & vbCrLf
    
           DoEvents
        Next x
        iCase = 0
        
    Case 3
       
        With WebBrowser1.Document
            .Forms(0)("txtSubject").Value = frmMain.txtSubject.Text
            .Forms(0)("txtEmail").Value = txtFrom.Text
            .Forms(0)("txtTo").Value = "wgoff@tampabay.rr.com"
            .Forms(0)("txtMessage").Value = frmMain.txtMsg.Text
            .Forms(0).submit.Click
            '.All("btnG").Click
        End With
        
        iCase = 0
        
    Case 4
        WebBrowser1.Document.Forms(0)(0).Checked = True
        iCase = 0
        
    
    Case 5
        
        WebBrowser1.Document.Forms(0)(1).Checked = True
        iCase = 0

End Select

End Sub


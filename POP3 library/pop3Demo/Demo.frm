VERSION 5.00
Begin VB.Form Demo 
   Caption         =   "EVICT_POP3 demo"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMail 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   3720
      Width           =   7695
   End
   Begin VB.ListBox lstMail 
      Height          =   1620
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   7695
   End
   Begin VB.Frame Frame1 
      Caption         =   "POP3 mail server"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   210
         Left            =   210
         TabIndex        =   18
         Top             =   1545
         Width           =   540
      End
      Begin VB.CommandButton cmdLogout 
         Caption         =   "Logout"
         Height          =   375
         Left            =   6000
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         Top             =   240
         Width           =   1575
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
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox tbxPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   8
         Text            =   "goffpauq"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox tbxUser 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "wgoff"
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox tbxHost 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "pop-server.tampabay.rr.com"
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
   Begin VB.TextBox txtLog 
      BackColor       =   &H80000018&
      Height          =   6735
      Left            =   7920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Demo.frx":0000
      Top             =   240
      Width           =   3975
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



Private Sub cmdLogin_Click()
Dim intloop As Integer
Dim id As Integer

    Set POP3 = Nothing
    Set POP3 = New EVICT_POP3.POP3
    POP3.POPHost = Me.tbxHost
    POP3.POPPort = Me.tbxPort
    POP3.UserName = Me.tbxUser
    POP3.Password = Me.tbxPassword
    POP3.Connect
    If POP3.MessageCount = 0 Then
       MsgBox "Sorry there are no messages in your inbox."
    Else
        Me.tbxCount = POP3.MessageCount
        Me.tbxSize = POP3.TotalMessageSize
    
        lstMail.Clear
        For intloop = 1 To POP3.MessageCount
            lstMail.AddItem intloop & ", " & POP3.MessageSize(intloop) & ", " & POP3.UniqueMessageID(intloop) & ", " & POP3.Subject(intloop)
        Next
    End If
   
End Sub


Private Sub Command1_Click()
    'MsgBox POP3.MessageID(1)
    Dim xxx As Long
    On Error Resume Next
    xxx = 1     'Val(Text1.Text)
    

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


Private Sub lstMail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim id As Integer

  id = Left(lstMail.List(lstMail.ListIndex), InStr(1, lstMail.List(lstMail.ListIndex), ",") - 1)

  If Button = 2 Then
     txtMail = "+Deleting message " & id & vbCrLf
     POP3.Delete (id)
     txtMail = txtMail & "+   Message " & id & " = " & POP3.MessageSize(id) & " bytes" & vbCrLf
     If MsgBox("Do you want to udo the delete?", vbYesNo, "Undo delete email") = vbYes Then
        ' A reset will reset all deletions. The only way to make a deletion final is to log out normally.
        txtMail = txtMail & "+Undoing the delete." & vbCrLf
        POP3.Reset
     End If
  End If
  
End Sub


Private Sub cmdLogout_Click()
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
    ' This will get the state of the form how you saved it last.
    '
    ' Warning: when you change a label then this will also be restored from the registry.
    ' If you do want to use the new label then you can disable the LoadFormFields once.
    ' This will restore all default settings. You could also change the LoadFormFields to
    ' bypass labels and frames.
    LoadFormFields
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set POP3 = Nothing
    Unload Me
    ' We do not want to save the log
    txtLog = "Log:"
    txtMail = ""
    ' This will save the current state of the form.
    SaveFormFields
End Sub


' This will save the value of every control into the registry (even labels)
' So you will have the state of the current form saved into the registry.
Private Sub SaveFormFields()
Dim ctrl As Control
    For Each ctrl In Me.Controls
        SaveSetting App.Title, Me.Name, ctrl.Name, ctrl
    Next
End Sub

' This will load the value of every control from the registry. (even labels)
' This will restore the state of the form to the one that you saved last.
Private Sub LoadFormFields()
Dim ctrl As Control
    For Each ctrl In Me.Controls
        ctrl = GetSetting(App.Title, Me.Name, ctrl.Name, ctrl)
    Next
End Sub



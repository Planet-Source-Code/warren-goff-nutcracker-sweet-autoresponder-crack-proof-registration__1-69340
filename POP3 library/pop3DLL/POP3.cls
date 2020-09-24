VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose:
' With this class object you can receive Email using the POP protoclol.

Option Explicit
Option Compare Text

' misc defaults
Private Const CONNECT_TIMEOUT = 10          ' seconds to wait before giving up
Private Const CONNECT_RETRY = 5             ' number of times to try before giving up
Private Const MSG_TIMEOUT = 60              ' seconds before timing out on message transmission
Private Const MSG_ATTACH_TIMEOUT = 600      ' seconds before timing out on attachment transmission
Private Const DEFAULT_PORT As Long = 110    ' default socket port for POP

Private plConnectTimeout        As Long
Private plCommandTimeout        As Long
Private plMessageTimeOut        As Long
Private plConnectRetry          As Long

Private plMessageCount As Long
Private plTotalMessageSize As Long
Private plMessageSize() As Long
Private psMessageID() As String
Private psMessageHeader() As String
Private strAttatchments() As String
Private lngID As Long

Private PlPopPort As Long
Private psPopHost As String
Private gsUserName                       As String   ' Auth username - optional, not supported by all servers
Private gsPassword                       As String   ' Auth password - optional, not supported by all servers

Private blnConnected As Boolean
Private strLastResponse As String

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

' constant & arrays used in Base64 & UUEncode encode/decode functions
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63) As String
Private pbBase64Byt(0 To 63) As Byte
Private psUUEncodeChr(0 To 63) As String

' Create an instance of the client
Private WithEvents oClient As GenericClient
Attribute oClient.VB_VarHelpID = -1

'Purpose:
' During the sending of the mail this event will show the current status of this opperations
Public Event Status(Status As String)
'Purpose:
' This event will return a number between 0 and 100 indicating the estimated percentage of the send in progress.
Public Event Progress(PercentComplete As Long)

Private Sub Class_Initialize()

' setup default values...

1     POPPort = DEFAULT_PORT                     ' Standard POP port (110)
2     ConnectTimeout = CONNECT_TIMEOUT
3     MessageTimeout = MSG_TIMEOUT
4     ConnectRetry = CONNECT_RETRY

5     ReDim plMessageSize(0) As Long
6     ReDim psMessageID(0) As String
7     ReDim psMessageHeader(0) As String

    ' initialize arrays for base64 & uuencode encoders
8 Dim iPTR As Integer
9     For iPTR = 0 To 63
10         pbBase64Byt(iPTR) = Asc(Mid$(BASE64CHR, iPTR + 1, 1))   ' byte array value
11         psBase64Chr(iPTR) = Mid$(BASE64CHR, iPTR + 1, 1)        ' string array value
12         psUUEncodeChr(iPTR) = Chr$(iPTR + &H20)
13     Next iPTR
14     psUUEncodeChr(0) = Chr$(&H60)

End Sub

Private Sub Class_Terminate()

15     On Error Resume Next ' We want to clean up whatever happens

        ' just in case things were accidently left open...
16         oClient.Connection.CloseSocket
17         Set oClient = Nothing

End Sub

'Purpose:
' You can give the command to connect and disconnect instead of letting the Send method doing that. This way it is possible to send multiple emails without setting up the connection for each email.

Public Sub Connect()

18     On Error GoTo ErrorHandler
19 Dim datnow As Date
20 Dim shello As String
21 Dim strTemp() As String

22     plMessageCount = 0
23     plTotalMessageSize = 0
24     blnConnected = False
25     Set oClient = Nothing
26     Set oClient = New GenericClient
27     RaiseEvent Status("Connecting to Server...")
28     oClient.Connect psPopHost, PlPopPort

    ' Wait for a connection
29     datnow = Now
30     While (blnConnected = False Or Left$(strLastResponse, 1) <> "+") And DateDiff("s", datnow, Now) < ConnectTimeout
31         DoEvents
32     Wend
    ' In case we could not connect (timeout)
33     If blnConnected = False Or Left$(strLastResponse, 1) <> "+" Then
34         RaiseEvent Status("Connect failed!")
35         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Connect : 6  - Connect failed!"
36         SendFail
37         Exit Sub
38     End If

    ' once a connection is established, login with username
39     RaiseEvent Status("Login in...")
40     If Not POPCommand("USER " & gsUserName & vbCrLf) Then
41         RaiseEvent Status("User " & gsUserName & " failed!")
42         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Connect : 6  - USER " & gsUserName & " failed!"
43         SendFail
44         Exit Sub
45     End If

    ' once a connection is established, login with password
46     If Not POPCommand("PASS " & gsPassword & vbCrLf) Then
47         RaiseEvent Status("PASS " & gsPassword & " failed!")
48         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Connect : 6  - PASS " & gsPassword & " failed!"
49         SendFail
50         Exit Sub
51     Else
52         RaiseEvent Status("Successfully loged in to pop server ")
53     End If

54     plTotalMessageSize = 0
55     If blnConnected Then
56         RaiseEvent Status("Getting the message count...")
57         If Not POPCommand("STAT" & vbCrLf) Then
58             RaiseEvent Status("TotalMessageSize : STAT failed!")
59             Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method MessageCount : 6  - STAT failed!"
60             SendFail
61             Exit Sub
62         End If
63     Else
64         RaiseEvent Status("TotalMessageSize : STAT failed")
65         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method MessageCount : 6  - STAT failed, Not connected to a POP server."
66         SendFail
67         Exit Sub
68     End If
69     strTemp = Split(strLastResponse, " ", , vbTextCompare)
70     If UBound(strTemp) > 1 Then
71         plMessageCount = CLng(strTemp(1))
72         plTotalMessageSize = CLng(strTemp(2))
73         ReDim psMessageHeader(plMessageCount) As String
74     End If

75 Exit Sub

76 ErrorHandler:
77     Resume Next
78     RaiseEvent Status("Error in method Connect on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description)
79     Err.Raise (vbObjectError Or Err), "POPmail", "POPmail :: Error in method Connect on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Purpose:
' You can give the command to connect and disconnect instead of letting the Send method doing that. This way it is possible to send multiple emails without setting up the connection for each email.

Public Sub Disconnect()

80     On Error Resume Next

81         If blnConnected Then
82             RaiseEvent Status("Closing Connection...")
83             POPCommand "QUIT" & vbCrLf
84             oClient.Connection.CloseSocket
85             lngID = 0
86             plMessageCount = 0
87         End If

End Sub

'Purpose:
' Just send a command and return the response

Private Function POPCommand(strCommand As String) As Boolean

88 Dim datnow As Date

89     DoEvents
90     datnow = Now
91     strLastResponse = ""
92     oClient.Connection.Send strCommand
93     If Left$(strCommand, 4) = "LIST" Or Left$(strCommand, 3) = "TOP" Or Left$(strCommand, 4) = "RETR" Or Left$(strCommand, 4) = "UIDL" Then
94         While blnConnected = True And InStr(1, strLastResponse, vbLf & "." & vbCrLf) = 0 And DateDiff("s", datnow, Now) < plMessageTimeOut
95             DoEvents
96         Wend
97     Else
98         While blnConnected = True And LenB(strLastResponse) = 0 And DateDiff("s", datnow, Now) < plMessageTimeOut
99             DoEvents
100         Wend
101     End If
102     POPCommand = (Left$(strLastResponse, 1) = "+")
103     RaiseEvent Status(strLastResponse)

End Function

Private Sub oClient_OnConnect()

104     blnConnected = True

End Sub

Private Sub oClient_OnClose()

105     blnConnected = False

End Sub

Private Sub oClient_OnDataArrive()

106 Dim strTemp As String

107     oClient.Connection.Recv strTemp
    '   If left$(strTemp, 1) = "+" Then
    '      strLastResponse = strTemp
    '   Else
108     strLastResponse = strLastResponse + strTemp
    '   End If

End Sub

Private Sub SendFail()

109     On Local Error Resume Next

        ' close the connection with the remote host
110         If blnConnected Then
111             POPCommand "QUIT" & vbCrLf
112             oClient.Connection.CloseSocket
113             blnConnected = False
114         End If
115         Set oClient = Nothing

End Sub

'----------------- Properties -------------------

'Purpose:
' To get/set the host name of the POP enabled email server that will be used.

Public Property Get POPHost() As String

116     POPHost = psPopHost

End Property

Public Property Let POPHost(NewValue As String)

117     psPopHost = NewValue

End Property

'Purpose:
' To get/set the IP port of the POP enabled email server that will be used. The default POP port is 25.

Public Property Get POPPort() As Long

118     POPPort = PlPopPort

End Property

Public Property Let POPPort(ByVal NewValue As Long)

119     If NewValue < 1 Or NewValue > 65535 Then
120         Err.Raise (vbObjectError Or 1), "POPmail", "POPmail :: Error in property POPPort : 1 - Invalid value"
121     Else
122         PlPopPort = NewValue
123     End If

End Property

'Purpose:
' To get/set the username that will log in to the POP server if it is required to log in. (UseAuthentication = true)

Public Property Get UserName() As String

124     UserName = gsUserName

End Property

Public Property Let UserName(ByVal NewValue As String)

125     gsUserName = NewValue

End Property

'Purpose:
' To get/set the password that will log in to the POP server if it is required to log in. (UseAuthentication = true)

Public Property Get Password() As String

126     Password = gsPassword

End Property

Public Property Let Password(ByVal NewValue As String)

127     gsPassword = NewValue

End Property

'Purpose:
' To get/set the number of seconds for setting up a connection with the host. The default is 10 seconds.

Public Property Get ConnectTimeout() As Long

128     ConnectTimeout = plConnectTimeout

End Property

Public Property Let ConnectTimeout(ByVal NewValue As Long)

129     If NewValue > 0 And NewValue <= 120 Then
130         plConnectTimeout = NewValue
131     End If

End Property

'Purpose:
' To get/set the number of seconds for sending a message. The default = 60

Public Property Get MessageTimeout() As Long

132     MessageTimeout = plMessageTimeOut

End Property

Public Property Let MessageTimeout(ByVal NewValue As Long)

133     If NewValue > 0 And NewValue <= 3600 Then
134         plMessageTimeOut = NewValue
135     End If

End Property

'Purpose:
' To get/set the number of tries if a send failed. The default is 5.

Public Property Get ConnectRetry() As Long

136     ConnectRetry = plConnectRetry

End Property

Public Property Let ConnectRetry(ByVal NewValue As Long)

137     If NewValue > 0 And NewValue <= 20 Then
138         plConnectRetry = NewValue
139     End If

End Property

'Purpose:
' To get the number of messages that is waiting on the mail server

Public Property Get MessageCount() As Long

140     MessageCount = plMessageCount

End Property

'Purpose:
' To get the size of all the messages that is waiting on the mail server

Public Property Get TotalMessageSize() As Long

141     TotalMessageSize = plTotalMessageSize

End Property

'Purpose:
' To get the size of all the messages that is waiting on the mail server

Public Property Get MessageSize(lngMessageID As Variant) As Long

142 Dim strTemp() As String
143 Dim intLoop As Integer

144     If UBound(plMessageSize) <> plMessageCount Then
145         ReDim plMessageSize(plMessageCount) As Long

146         If blnConnected Then
147             RaiseEvent Status("Getting the message size of each message ...")
148             If Not POPCommand("LIST" & vbCrLf) Then
149                 RaiseEvent Status("MessageSize : LIST failed!")
150                 Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method MessageSize : 6  - LIST failed!"
151                 SendFail
152                 Exit Property
153             End If
154             strTemp = Split(strLastResponse, vbCrLf, , vbTextCompare)
155             For intLoop = 1 To UBound(strTemp) - 2
156                 plMessageSize(intLoop) = CLng(Mid$(strTemp(intLoop), InStr(1, strTemp(intLoop), " ") + 1))
157             Next intLoop
158         Else
159             RaiseEvent Status("MessageSize : LIST failed")
160             Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method MessageSize : 6  - LIST failed, Not connected to a POP server."
161             SendFail
162             Exit Property
163         End If
164     End If
165     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
166         RaiseEvent Status("MessageSize(" & lngMessageID & ") failed! The message ID does not exist.")
167         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method MessageSize(" & lngMessageID & ") failed! The message ID does not exist."
168     Else
169         MessageSize = plMessageSize(CInt(lngMessageID))
170     End If

End Property

'Purpose:
' Getting the raw body of an email

Public Property Get RawEmailText(lngMessageID As Variant) As String

171 Static strRaw As String

' In case we already got the message

172     If lngID = lngMessageID Then
173         RawEmailText = strRaw
174         Exit Property
175     End If
176     strRaw = ""
177     lngID = 0

178     If blnConnected Then
        ' Of course we must be connected
179         If CInt(lngMessageID) < 1 Or CInt(lngMessageID) > UBound(plMessageSize) Then
            ' The message ID does not exist.
180             RaiseEvent Status("RawEmailText : RETR " & lngMessageID & " failed! The message ID does not exist.")
181             Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method RawEmailText : 6  - RETR " & lngMessageID & " failed! The message ID does not exist."
182             Exit Property
183         Else
184             If plMessageSize(CInt(lngMessageID)) = 0 Then
                ' Didn't we already delete it
185                 RaiseEvent Status("RawEmailText : RETR " & lngMessageID & " failed! The message was deleted.")
186                 Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method RawEmailTop : 6  - TOP " & lngMessageID & " failed! The message was deleted."
187                 Exit Property
188             Else
                ' Send the retr command and wait for the response
189                 RaiseEvent Status("Getting the message.")
190                 If POPCommand("RETR " & lngMessageID & vbCrLf) = False Then
191                     RaiseEvent Status("RawEmailText : RETR failed!")
192                     Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method RawEmailText : 6  - RETR failed!"
193                     SendFail
194                     Exit Property
195                 End If
196             End If
197         End If
198     Else
        ' We are not connected to a server
199         RaiseEvent Status("RawEmailText : RETR failed")
200         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method RawEmailText : 6  - RETR failed, Not connected to a POP server."
201         SendFail
202         Exit Property
203     End If
204     strRaw = Mid$(strLastResponse, 6, Len(strLastResponse) - 10)
205     RawEmailText = strRaw
206     lngID = lngMessageID
207     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
208         RaiseEvent Status("psMessageHeader(" & lngMessageID & ") failed! The message ID does not exist.")
209         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method psMessageHeader(" & lngMessageID & ") failed! The message ID does not exist."
210     Else
211         psMessageHeader(CLng(lngMessageID)) = Left$(RawEmailText, InStr(1, RawEmailText, vbCrLf & vbCrLf))
212     End If

End Property

'Purpose:
' Replace special HTML characters by HTML tags

Public Function HTMLencode(strText As String) As String

213     HTMLencode = Replace(HTMLencode, "&", "&amp;")
214     HTMLencode = Replace(strText, ">", "&gt;")
215     HTMLencode = Replace(HTMLencode, "<", "&lt;")
216     HTMLencode = Replace(HTMLencode, vbCrLf, "<BR>")
217     HTMLencode = Replace(HTMLencode, vbCr, "<BR>")
218     HTMLencode = Replace(HTMLencode, vbLf, "<BR>")
219     HTMLencode = Replace(HTMLencode, vbTab, "&#09;")
220     HTMLencode = Replace(HTMLencode, vbBack, "&#09;")
221     HTMLencode = Replace(HTMLencode, Chr(34), "&quot;")

End Function

'Purpose:
' The only thing needed for a field is the replace of a double quote by &quot;

Public Function HTMLfieldEncode(strText As String) As String

222     HTMLfieldEncode = Replace(strText, Chr(34), "&quot;")

End Function

'Purpose:
' Replace special HTML tags by the corresponding characters.

Public Function HTMLdecode(strText As String) As String

223     HTMLdecode = Replace(strText, "&gt;", ">")
224     HTMLdecode = Replace(HTMLdecode, "&lt;", "<")
225     HTMLdecode = Replace(HTMLdecode, "<BR>", vbCrLf)
226     HTMLdecode = Replace(HTMLdecode, "&#09;", vbTab)
227     HTMLdecode = Replace(HTMLdecode, "&#09;", vbBack)
228     HTMLdecode = Replace(HTMLdecode, "&quot;", Chr(34))
229     HTMLdecode = Replace(HTMLdecode, "&amp;", "&")

End Function

'Purpose:
' The only thing needed for a field is the replace of &quot; by a  double quotes

Public Function HTMLfieldDecode(strText As String) As String

230     HTMLfieldDecode = Replace(strText, "&quot;", Chr(34))

End Function

'Purpose:
' Getting the first number of lines of the body of an email including the header.

Public Property Get RawEmailTop(lngMessageID As Variant, lngNumberOfLines As Variant) As String

231     If blnConnected Then
        ' Of course we must be connected
232         If CInt(lngMessageID) < 1 Or CInt(lngMessageID) > UBound(plMessageSize) Then
            ' The message ID does not exist.
233             RaiseEvent Status("RawEmailTop : TOP " & lngMessageID & " " & lngNumberOfLines & " failed! The message ID does not exist.")
234             Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method RawEmailTop : 6  - TOP " & lngMessageID & " " & lngNumberOfLines & " failed! The message ID does not exist."
235             Exit Property
236         Else
237             If plMessageSize(CInt(lngMessageID)) = 0 Then
                ' Didn't we already delete it
238                 RaiseEvent Status("RawEmailText : RETR " & lngMessageID & " failed! The message was deleted.")
239                 Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method RawEmailTop : 6  - TOP " & lngMessageID & " " & lngNumberOfLines & " failed! The message was deleted."
240                 Exit Property
241             Else
                ' Send the delete command and wait for the response
242                 RaiseEvent Status("Getting the top of message.")
243                 If Not POPCommand("TOP " & lngMessageID & " " & lngNumberOfLines & vbCrLf) Then
244                     RaiseEvent Status("RawEmailTop : TOP " & lngMessageID & " " & lngNumberOfLines & " failed!")
245                     Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method RawEmailTop : 6  - TOP " & lngMessageID & " " & lngNumberOfLines & " failed!"
246                     SendFail
247                     Exit Property
248                 End If
249             End If
250         End If
251     Else
        ' We are not connected to a server
252         RaiseEvent Status("RawEmailTop : TOP " & lngMessageID & " " & lngNumberOfLines & " failed")
253         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method RawEmailTop : 6  - TOP " & lngMessageID & " " & lngNumberOfLines & " failed, Not connected to a POP server."
254         SendFail
255         Exit Property
256     End If
257     RawEmailTop = Mid$(strLastResponse, 6, Len(strLastResponse) - 10)

258     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
259         RaiseEvent Status("psMessageHeader(" & lngMessageID & ") failed! The message ID does not exist.")
260         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method psMessageHeader(" & lngMessageID & ") failed! The message ID does not exist."
261     Else
262         psMessageHeader(CLng(lngMessageID)) = Left$(RawEmailTop, InStr(1, RawEmailTop, vbCrLf & vbCrLf))
263     End If

End Property

'Purpose:
' Delete the specified message.

Public Sub Delete(lngMessageID As Variant)

264     If blnConnected Then
        ' Of course we must be connected
265         If CInt(lngMessageID) < 1 Or CInt(lngMessageID) > UBound(plMessageSize) Then
            ' The message ID does not exist.
266             RaiseEvent Status("Delete : DELE " & lngMessageID & " failed! The message ID does not exist.")
267             Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Delete : 6  - DELE " & lngMessageID & " failed! The message ID does not exist."
268             Exit Sub
269         Else
270             If plMessageSize(CInt(lngMessageID)) = 0 Then
                ' Didn't we already delete it
271                 RaiseEvent Status("RawEmailText : RETR " & lngMessageID & " failed! The message was already deleted.")
272             End If
            ' Send the delete command and wait for the response
273             RaiseEvent Status("Deleting the message.")
274             If Not POPCommand("DELE " & lngMessageID & vbCrLf) Then
275                 RaiseEvent Status("Delete : DELE " & lngMessageID & " failed!")
276                 Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Delete : 6  - DELE " & lngMessageID & " failed!"
277                 SendFail
278                 Exit Sub
279             Else
280                 RaiseEvent Status("Delete : DELE " & lngMessageID & " was successfull.")
281                 plMessageSize(CInt(lngMessageID)) = 0
282             End If
283         End If
284     Else
        ' We are not connected to a server
285         RaiseEvent Status("Delete : DELE " & lngMessageID & " failed")
286         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Delete : 6  - DELE " & lngMessageID & " failed, Not connected to a POP server."
287         SendFail
288         Exit Sub
289     End If

End Sub

'Purpose:
' Cansel all deletions.

Public Sub Reset()

290     ReDim plMessageSize(0) As Long
291     If blnConnected Then
292         RaiseEvent Status("Cancel all email deletions.")
293         If Not POPCommand("RSET" & vbCrLf) Then
294             RaiseEvent Status("Reset : RSET failed!")
295             Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Reset : RSET failed!"
296             SendFail
297             Exit Sub
298         Else
299             RaiseEvent Status("Reset : RSET was successfull.")
300         End If
301     Else
        ' We are not connected to a server
302         RaiseEvent Status("Reset : RSET  failed")
303         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Reset : RSET failed, Not connected to a POP server."
304         SendFail
305     End If

End Sub

'Purpose:
' This will send a 'NOOP' to the server in order to see if we are still connected. This will also prevent losing the connection because of a time out.

Public Function TestConnection() As Boolean

306     TestConnection = False
307     ReDim plMessageSize(0) As Long
308     If blnConnected Then
309         RaiseEvent Status("Test if we are still connected (this will also keep the connection up).")
310         If Not POPCommand("NOOP" & vbCrLf) Then
311             RaiseEvent Status("TestConnection : NOOP failed!")
312         Else
313             RaiseEvent Status("TestConnection : NOOP was successfull.")
314             TestConnection = True
315         End If
316     Else
        ' We are not connected to a server
317         RaiseEvent Status("TestConnection : NOOP  failed")
318     End If

End Function

'Purpose:
' To get the size of all the messages that is waiting on the mail server

Public Property Get UniqueMessageID(lngMessageID As Variant) As String

319 Dim strTemp() As String
320 Dim intLoop As Integer

321     If UBound(psMessageID) <> plMessageCount Then
322         ReDim psMessageID(plMessageCount) As String

323         If blnConnected Then
324             RaiseEvent Status("Getting the message size of each message ...")
325             If Not POPCommand("UIDL" & vbCrLf) Then
326                 RaiseEvent Status("UniqueMessageID : UIDL failed!")
327                 Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method UniqueMessageID : 6  - UIDL failed!"
328                 SendFail
329                 Exit Property
330             End If
331             strTemp = Split(strLastResponse, vbCrLf, , vbTextCompare)
332             For intLoop = 1 To UBound(strTemp) - 2
333                 psMessageID(intLoop) = Mid$(strTemp(intLoop), InStr(1, strTemp(intLoop), " ") + 1)
334             Next intLoop
335         Else
336             RaiseEvent Status("UniqueMessageID : UIDL failed")
337             Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method UniqueMessageID : 6  - UIDL failed, Not connected to a POP server."
338             SendFail
339             Exit Property
340         End If
341     End If
342     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
343         RaiseEvent Status("UniqueMessageID(" & lngMessageID & ") failed! The message ID does not exist.")
344         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method UniqueMessageID(" & lngMessageID & ") failed! The message ID does not exist."
345     Else
346         UniqueMessageID = psMessageID(CInt(lngMessageID))
347     End If

End Property

'Purpose:
' Getting the first number of lines of the body of an email including the header.

Public Property Get EmailHeader(lngMessageID As Variant) As String

348     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
349         RaiseEvent Status("EmailHeader(" & lngMessageID & ") failed! The message ID does not exist.")
350         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method EmailHeader(" & lngMessageID & ") failed! The message ID does not exist."
351     Else
352         If LenB(Trim$(psMessageHeader(CLng(lngMessageID)) & "")) = 0 Then
353             EmailHeader = RawEmailTop(lngMessageID, 1)
354         End If
355         EmailHeader = Trim$(psMessageHeader(CLng(lngMessageID)) & "")
356         EmailHeader = Replace(EmailHeader, vbCrLf & vbTab, " ")
357     End If

End Property

'Purpose:
' Getting the first number of lines of the body of an email including the header.

Public Property Get Body(lngMessageID As Variant) As String

358 Dim strSep As String
359 Dim lngLoop As Long

360     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
361         RaiseEvent Status("Body(" & lngMessageID & ") failed! The message ID does not exist.")
362         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Body(" & lngMessageID & ") failed! The message ID does not exist."
363     Else
364         Body = RawEmailText(lngMessageID)
365         If Left$(ContentType(lngMessageID), 9) = "multipart" Then
366             Body = Mid$(Body, InStr(1, Body, vbCrLf & vbCrLf) + 4)      ' First start at the warning
367             Body = Mid$(Body, InStr(1, Body, Seperator(lngMessageID)) + Len(Seperator(lngMessageID)) + 2) ' Start after the first seperator
368             Body = Mid$(Body, InStr(1, Body, vbCrLf & vbCrLf) + 4)      ' Skip the info header
369             Body = Left$(Body, InStr(1, Body, Seperator(lngMessageID)) - 3)
370         Else
371             Body = Mid$(Body, InStr(1, Body, vbCrLf & vbCrLf) + 4)
372         End If
373     End If

End Property

Public Property Get BodyAlternative(lngMessageID As Variant) As String

374     If Alternative(lngMessageID) Then
375         BodyAlternative = Attachment(lngMessageID, 0)
376     Else
377         BodyAlternative = Empty
378     End If

End Property

'Purpose:
' Get info from the header

Private Function GetFromHeader(lngMessageID As Variant, strWhat As String) As String

379 Dim strTemp() As String
380 Dim intLoop As Integer

381     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
382         RaiseEvent Status("GetFromHeader(" & strWhat & ", " & lngMessageID & ") failed! The message ID does not exist.")
383         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method GetFromHeader(" & strWhat & ", " & lngMessageID & ") failed! The message ID does not exist."
384     Else
385         GetFromHeader = ""
386         strTemp = Split(EmailHeader(lngMessageID), vbCrLf, -1, vbTextCompare)
387         For intLoop = 0 To UBound(strTemp)
388             If Left$(strTemp(intLoop), Len(strWhat)) = strWhat Then
389                 If LenB(GetFromHeader) = 0 Then
390                     GetFromHeader = Mid$(strTemp(intLoop), Len(strWhat) + 2)
391                 Else
392                     GetFromHeader = GetFromHeader & vbCrLf & Mid$(strTemp(intLoop), Len(strWhat) + 2)
393                 End If
394             End If
395         Next intLoop
396     End If

End Function

'Purpose:
' Getting the from out of the header

Public Property Get From(lngMessageID As Variant) As String

397     From = GetFromHeader(lngMessageID, "From:")

End Property

'Purpose:
' Getting the To out of the header

Public Property Get Recipient(lngMessageID As Variant) As String

398     Recipient = GetFromHeader(lngMessageID, "To:")

End Property

'Purpose:
' Getting the CC out of the header

Public Property Get CC(lngMessageID As Variant) As String

399     CC = GetFromHeader(lngMessageID, "CC:")

End Property

'Purpose:
' Getting the Subject out of the header

Public Property Get Subject(lngMessageID As Variant) As String

400     Subject = GetFromHeader(lngMessageID, "Subject:")

End Property

'Purpose:
' Getting the send date out of the header

Public Property Get SendDate(lngMessageID As Variant) As String

401     SendDate = GetFromHeader(lngMessageID, "Date:")

End Property

'Purpose:
' Getting the Message ID out of the header

Public Property Get MessageID(lngMessageID As Variant) As String

402     MessageID = GetFromHeader(lngMessageID, "Message-ID:")

End Property

'Purpose:
' Getting the MIME-Version out of the header

Public Property Get MimeVersion(lngMessageID As Variant) As String

403     MimeVersion = GetFromHeader(lngMessageID, "MIME-Version:")

End Property

'Purpose:
' Getting the Content type out of the header

Public Property Get ContentType(lngMessageID As Variant) As String

404     ContentType = GetFromHeader(lngMessageID, "Content-Type:")

End Property

Public Property Get Alternative(lngMessageID As Variant) As Boolean

405 Dim ContentType As String

406     ContentType = GetFromHeader(lngMessageID, "Content-Type:")
407     If InStr(1, ContentType, "alternative") > 0 Then
408         Alternative = True
409     Else
410         Alternative = False
411     End If

End Property

'Purpose:
' Getting the Content type out of the header

Public Property Get Importance(lngMessageID As Variant) As String

412     Importance = GetFromHeader(lngMessageID, "Importance:")

End Property

'Purpose:
' Getting the Message Flag (out of the header if supplied)

Public Property Get MessageFlag(lngMessageID As Variant) As String

413     MessageFlag = GetFromHeader(lngMessageID, "X-Message-Flag:")

End Property

'Purpose:
' Getting the Expire date of the message (out of the header if supplied)

Public Property Get ExpireDate(lngMessageID As Variant) As String

414     ExpireDate = GetFromHeader(lngMessageID, "Expiry-Date:")

End Property

'Purpose:
' Getting the Expire date of the message (out of the header if supplied)

Public Property Get Sensitivity(lngMessageID As Variant) As String

415     Sensitivity = GetFromHeader(lngMessageID, "Sensitivity:")

End Property

'Purpose:
' Return true if there are attachments

Public Property Get HasAttachments(lngMessageID As Variant) As Boolean

416     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
417         RaiseEvent Status("HasAttachments(" & lngMessageID & ") failed! The message ID does not exist.")
418         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method HasAttachments(" & lngMessageID & ") failed! The message ID does not exist."
419     Else
420         If Left$(ContentType(lngMessageID), 9) = "multipart" Then
421             HasAttachments = True
422         Else
423             HasAttachments = False
424         End If
425     End If

End Property

'Purpose:
' Getting the Content type out of the header

Public Property Get Received(lngMessageID As Variant) As String

426     Received = GetFromHeader(lngMessageID, "Received:")

End Property

'Purpose:
' Getting the number of attachments

Public Property Get AttachmentCount(lngMessageID As Variant) As String

427 Dim strTemp As String

428     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
429         RaiseEvent Status("AttachmentCount(" & lngMessageID & ") failed! The message ID does not exist.")
430         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method AttachmentCount(" & lngMessageID & ") failed! The message ID does not exist."
431     Else
432         strTemp = RawEmailText(lngMessageID)
433         If Left$(ContentType(lngMessageID), 9) = "multipart" Then
434             strAttatchments = Split(strTemp, Seperator(lngMessageID), -1, vbTextCompare)
435             AttachmentCount = UBound(strAttatchments) - 2
436             If Alternative(lngMessageID) Then AttachmentCount = AttachmentCount - 1
437         Else
438             AttachmentCount = 0
439         End If
440     End If

End Property

'Purpose:
' Getting the attachments

Public Property Get Attachment(lngMessageID As Variant, intAttachment As Variant) As String

441 Dim strAttatchments() As String
442 Dim strTemp As String
443 Dim strEnc As String
444 Dim intAlt As Integer

445     If Alternative(lngMessageID) Then intAlt = 1 Else intAlt = 0

446     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
447         RaiseEvent Status("Attachment(" & lngMessageID & ") failed! The message ID does not exist.")
448         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Attachment(" & lngMessageID & ") failed! The message ID does not exist."
449     Else
450         strTemp = RawEmailText(lngMessageID)
451         If Left$(ContentType(lngMessageID), 9) = "multipart" Then
452             strAttatchments = Split(strTemp, Seperator(lngMessageID), -1, vbTextCompare)
453             Attachment = strAttatchments(intAttachment + 1 + intAlt)
454             Attachment = Mid$(Attachment, InStr(1, Attachment, vbCrLf & vbCrLf) + 4)
455             Attachment = Left$(Attachment, Len(Attachment) - 2)
456             strEnc = AttachmentEncoding(lngMessageID, intAttachment + intAlt)
457             If InStr(1, strEnc, "base64") > 0 Then
458                 Attachment = DecodeBase64String(Attachment)
459             End If
460         Else
461             Attachment = ""
462         End If
463     End If

End Property

'Purpose:
' Saving the attachments

Public Sub SaveAttachment(lngMessageID As Variant, intAttachment As Variant, strFileName As String)

464     SaveFile strFileName, Attachment(lngMessageID, intAttachment)

End Sub

'Purpose:
' Getting the attachment name

Public Property Get AttachmentName(lngMessageID As Variant, intAttachment As Variant) As String

465 Dim strTemp As String
466 Dim strAttatchments() As String
467 Dim n As Long
468 Dim nl As Long

469     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
470         RaiseEvent Status("AttachmentName(" & lngMessageID & ") failed! The message ID does not exist.")
471         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method AttachmentName(" & lngMessageID & ") failed! The message ID does not exist."
472     End If

473     strTemp = RawEmailText(lngMessageID)
474     If Left$(ContentType(lngMessageID), 9) = "multipart" Then
475         strAttatchments = Split(strTemp, Seperator(lngMessageID), -1, vbTextCompare)
476         n = InStr(1, strAttatchments(intAttachment + 1), "name=") + 5
477         nl = InStr(n, strAttatchments(intAttachment + 1), vbCrLf) - n
478         AttachmentName = Trim$(Mid$(strAttatchments(intAttachment + 1), n, nl))
479         If Left$(AttachmentName, 1) = Chr(34) Then AttachmentName = Mid$(AttachmentName, 2)
480         If Right$(AttachmentName, 1) = Chr(34) Then AttachmentName = Left$(AttachmentName, Len(AttachmentName) - 1)
481     Else
482         AttachmentName = ""
483     End If

End Property

'Purpose:
' Getting the attachment Encoding

Public Property Get AttachmentEncoding(lngMessageID As Variant, intAttachment As Variant) As String

484 Dim strTemp As String
485 Dim strAttatchments() As String
486 Dim n As Long
487 Dim nl As Long

488     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
489         RaiseEvent Status("AttachmentEncoding(" & lngMessageID & ") failed! The message ID does not exist.")
490         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method AttachmentEncoding(" & lngMessageID & ") failed! The message ID does not exist."
491     End If

492     strTemp = RawEmailText(lngMessageID)
493     If Left$(ContentType(lngMessageID), 9) = "multipart" Then
494         strAttatchments = Split(strTemp, Seperator(lngMessageID), -1, vbTextCompare)
495         n = InStr(1, strAttatchments(intAttachment + 1), "Encoding:") + 9
496         nl = InStr(n, strAttatchments(intAttachment + 1), vbCrLf) - n
497         AttachmentEncoding = Trim$(Mid$(strAttatchments(intAttachment + 1), n, nl))
498     Else
499         AttachmentEncoding = ""
500     End If

End Property

'Purpose:
' Getting the attachment content type

Public Property Get AttachmentContentType(lngMessageID As Variant, intAttachment As Variant) As String

501 Dim strTemp As String
502 Dim strAttatchments() As String
503 Dim n As Long
504 Dim nl As Long

505     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
506         RaiseEvent Status("AttachmentContentType(" & lngMessageID & ") failed! The message ID does not exist.")
507         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method AttachmentContentType(" & lngMessageID & ") failed! The message ID does not exist."
508     End If

509     strTemp = RawEmailText(lngMessageID)
510     If Left$(ContentType(lngMessageID), 9) = "multipart" Then
511         strAttatchments = Split(strTemp, Seperator(lngMessageID), -1, vbTextCompare)
512         n = InStr(1, strAttatchments(intAttachment + 1), "Content-Type:") + 13
513         nl = InStr(n, strAttatchments(intAttachment + 1), vbCrLf) - n
514         AttachmentContentType = Trim$(Mid$(strAttatchments(intAttachment + 1), n, nl))
515     Else
516         AttachmentContentType = ""
517     End If

End Property

'Purpose:
' To get the multipart seperator.

Private Function Seperator(lngMessageID As Variant)

518 Dim strSep As String

519     If lngMessageID < 1 Or lngMessageID > UBound(plMessageSize) Then
        ' The message ID does not exist.
520         RaiseEvent Status("Seperator(" & lngMessageID & ") failed! The message ID does not exist.")
521         Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in method Seperator(" & lngMessageID & ") failed! The message ID does not exist."
522     End If

523     If Left$(ContentType(lngMessageID), 9) = "multipart" Then
524         strSep = Trim$(Mid$(ContentType(lngMessageID), InStr(1, ContentType(lngMessageID), "boundary") + 8))
525         If Left$(strSep, 1) = "=" Then strSep = Mid$(strSep, 2)
526         If Left$(strSep, 1) = Chr(34) Then strSep = Mid$(strSep, 2)
527         If Right$(strSep, 1) = vbCr Then strSep = Left$(strSep, Len(strSep) - 1)
528         If Right$(strSep, 1) = Chr(34) Then strSep = Left$(strSep, Len(strSep) - 1)
529         Seperator = "--" & strSep
530     Else
531         Seperator = ""
532     End If

End Function

' Description:
'   Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
'   values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
'   ascii character equivalent. Stop converting at the end of the input string
'   or when the first '=' (equal sign) is encountered.

Private Function DecodeBase64String(str2Decode As String) As String

533     On Error GoTo ErrorHandler
534 Dim lPtr            As Long
535 Dim iValue          As Integer
536 Dim iLen            As Integer
537 Dim iCtr            As Integer
538 Dim Bits(1 To 4)    As Byte
539 Dim strDecode       As String
540 Dim lPtr2           As Long

    ' Clean it up and create the destination buffer
541     str2Decode = Replace(str2Decode, vbCrLf, "")
542     lPtr2 = 0
543     strDecode = String$(Len(str2Decode) * 3 / 4 + 3, " ")
    ' for each 4 character group....
544     For lPtr = 1 To Len(str2Decode) Step 4
545         iLen = 4
546         For iCtr = 0 To 3
            ' retrive the base 64 value, 4 at a time
547             iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
548             Select Case iValue
                ' A~Za~z0~9+/
            Case 1 To 64: Bits(iCtr + 1) = iValue - 1
                ' =
549             Case 65
550                 iLen = iCtr
551                 Exit For
                ' not found
552             Case 0
553                 Exit Function
554             End Select
555         Next iCtr

        ' convert the 4, 6 bit values into 3, 8 bit values
556         Bits(1) = Bits(1) * &H4 + (Bits(2) And &H30) \ &H10
557         Bits(2) = (Bits(2) And &HF) * &H10 + (Bits(3) And &H3C) \ &H4
558         Bits(3) = (Bits(3) And &H3) * &H40 + Bits(4)

        ' add the three new characters to the output string
559         For iCtr = 1 To iLen - 1
560             lPtr2 = lPtr2 + 1
561             Mid$(strDecode, lPtr2, 1) = Chr$(Bits(iCtr))
562         Next iCtr

563     Next lPtr

564     DecodeBase64String = Left$(strDecode, lPtr2)

565 Exit Function

566 ErrorHandler:
567     Err.Raise (vbObjectError Or 6), "POPmail", "POPmail :: Error in function DecodeBase64String(" & Left$(str2Decode, 16) & "....) failed! The message ID does not exist."
568     Resume Next

End Function

'Purpose:
' This function will save a string to the specified file.

Private Sub SaveFile(FileName As String, FileContent As String)

569     On Error GoTo ErrorHandler
570 Dim intnextfreefile As Integer

571     intnextfreefile = FreeFile
572     Open FileName For Binary As #intnextfreefile
573     Put #intnextfreefile, , FileContent  ' might need something like arrByteArray()
574     Close #intnextfreefile

575 Exit Sub

576 ErrorHandler:
577     Err.Raise (vbObjectError Or Err), "Functions", "Error in method SaveFile on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

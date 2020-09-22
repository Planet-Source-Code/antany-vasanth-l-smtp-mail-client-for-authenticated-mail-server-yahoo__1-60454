VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form2"
   ScaleHeight     =   4995
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStat 
      Height          =   3735
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Main 
      Caption         =   "Main"
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtBody 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtSub 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblBody 
         Alignment       =   2  'Center
         Caption         =   "Body:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label lblSubject 
         Caption         =   "Subject:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblTo 
         Caption         =   "To:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   240
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IP As String
Private Connected As Boolean 'This defines "connected" it will be either True or False, see WS Connect.
Dim Tfrom As String, Tto As String, Tsub As String, Tbody As String 'these are all variables...
Dim ver As String

Private Sub cmdSend_Click()
WS.Close
WS.Connect Form1.Text1.Text, Val(Form1.Text2.Text) 'connected to the server(cmbServer's text) on the defined port, in this case 25
Pause 2 'Waits 2 seconds so winsock can connect before the next event
If Connected = False Then
MsgBox "Connection problem try again.", vbCritical, "Anonymous Mailer"
Me.Hide
Form1.Show
Else
  Form2.Caption = "Connect to Server " & Form1.Text1.Text
End If


Tfrom = txtFrom.Text
Tto = txtTo.Text
Tsub = txtSub.Text
Tbody = txtBody.Text
IP = WS.LocalIP
If txtFrom.Text = "" Then 'if there is no text it txtFrom the following Message Box will be displayed
MsgBox "Enter the senders email address", vbCritical, "Anonymous Mailer"
Exit Sub
End If
If txtTo.Text = "" Then ''if there is no text it txtTo the following Message Box will be displayed
MsgBox "Enter Receivers Email Address", vbCritical, "Anonymous Mailer"
Exit Sub
End If
If txtBody.Text = "" Then
MsgBox "Enter Some Content to Body"
Exit Sub
End If
WS.SendData "HELO " & IP & Chr(13) & Chr(10) 'Send the command "HELO " and whatever IP is set to, to the server
LogText "HELO " & IP & Chr(13) & Chr(10) 'Send the command "HELO " and whatever IP is set to, to the server
Pause 2
WS.SendData "AUTH LOGIN" & Chr(13) & Chr(10)
LogText "AUTH LOGIN" & Chr(13) & Chr(10)
Pause 2
WS.SendData EncodeStr64(Form1.Text3.Text) & Chr(13) & Chr(10)
LogText EncodeStr64(Form1.Text3.Text) & Chr(13) & Chr(10)
Pause 2
WS.SendData EncodeStr64(Form1.Text4.Text) & Chr(13) & Chr(10)
LogText EncodeStr64(Form1.Text4.Text) & Chr(13) & Chr(10)
Pause 4
'WS.GetData ver
MsgBox ver
If Val(Mid$(ver, 1, 3)) = 535 Then
  MsgBox "Invalid UserName and Password"
  Form1.Show
  Unload Me
End If


WS.SendData "MAIL FROM: " & txtFrom & Chr(13) & Chr(10) 'sends the e-mail address
LogText "MAIL FROM: " & txtFrom & Chr(13) & Chr(10) 'sends the e-mail address
Pause 1
WS.SendData "RCPT TO: " & "<" & Tto & ">" & Chr(13) & Chr(10) 'tells the server who will be receiving the e-mail
LogText "RCPT TO: " & "<" & Tto & ">" & Chr(13) & Chr(10)
Pause 1
WS.SendData "DATA" & Chr(13) & Chr(10) 'starts sending Data
LogText "DATA" & Chr(13) & Chr(10) 'starts sending Data
WS.SendData "From:" & Tfrom & Chr(13) & Chr(10)
LogText "From:" & Tfrom & Chr(13) & Chr(10)
WS.SendData "To:" & Tto & Chr(13) & Chr(10)
LogText "To:" & Tto & Chr(13) & Chr(10)
WS.SendData "SUBJECT: " & Tsub & " from: " & WS.LocalHostName & Chr(13) & Chr(13) 'specifies the subject of the e-mail
LogText "SUBJECT: " & Tsub & " from: " & WS.LocalHostName & Chr(13) & Chr(13) 'specifies the subject of the e-mail
WS.SendData Tbody & Chr(13) & Chr(10) 'send the body of the e-mail
LogText Tbody & Chr(13) & Chr(10) 'send the body of the e-mail
WS.SendData "Date at Send " & Date & Chr(13)
WS.SendData "Sender Ip :" & WS.LocalIP & Chr(13)
WS.SendData "Sender Ip :" & WS.LocalHostName & Chr(13) & Chr(10)
WS.SendData "." & Chr(13) & Chr(10) 'ends the sending of Data
Pause 2

Pause 1
MsgBox "Messages Sent"
LogText "quit"
WS.SendData "QUIT" & vbCrLf 'Exits server
End Sub

Private Sub Form_Load()
WS.Close
WS.Connect Form1.Text1.Text, Val(Form1.Text2.Text) 'connected to the server(cmbServer's text) on the defined port, in this case 25
Pause 2 'Waits 2 seconds so winsock can connect before the next event
If Connected = False Then
MsgBox "Connection problem try again.", vbCritical, "Anonymous Mailer"
Me.Hide
Form1.Show
Else
  Form2.Caption = "Connect to Server " & Form1.Text1.Text
End If

End Sub
Public Sub Pause(Duration As Double) 'I needed this Sub to make intervals between the sending of commands
Dim Current As Long 'Duration ican be change (i mean the amout of time)
Current = Timer
Do Until Timer - Current >= Duration 'Loops event until the current time matches the Duration defined
DoEvents
DoEvents

Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub

Private Sub WS_Connect() 'when Winsock connects
Form2.Caption = "Conneted To: " & WS.RemoteHost & vbCrLf 'changes the text of txtstat to "Connected To: " and displays the IP of the server.
Connected = True 'It is connected to somthing therefor Connected will be true.
End Sub 'ends sub


Private Sub WS_DataArrival(ByVal bytesTotal As Long) 'bytes recieved counter fixed thnx to nspyrou@yahoo.com
'If WS.BytesReceived <> 0 Then ...
    If WS.BytesReceived <> 0 Then
'add the total bytes received to the caption of lblBR
       
    End If

Dim Data As String 'defines data
WS.GetData Data 'gets data
ver = Data
LogText Data 'logs data on txtStat
End Sub 'ends sub
Sub LogText(Text As String) 'Now when "LogText" is typed it will display the text in txtStat w/o using "txtStat.text"
txtStat.Text = txtStat.Text & Text & Chr(13) & Chr(10) 'the text of txtStat will be the same when new Text is added, "Text" is the new text and it will be added to txtStat
End Sub 'Ends sub


VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat - Server"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   300
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   240
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   300
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtServerPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "< PORT >"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtServerIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "< IP ADDRESS>"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   300
      Left            =   3840
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Hello!!!"
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Local Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   530
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   170
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Caption         =   "Server is Closed..."
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   540
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' * Basic Chat Program *

' description: this is a basic chat program consist of Client and Server.

' · Server - can start and stop listening for incoming connection
'            and can send data (message) to the client...
' · Client - connects to the listening Server with the specified Port
'            and ip, and can send data (message) to the server...

' author: Vanni Alora
' email:  vanjo08@msn.com
' url:    vanjo08@cjb.net

' this is my third submission to PSC, feel great coz i have done my
' first chat program. its a basic but for beginners like me... its
' really COOL. im a newbie in winsock and hope you may find it usefull
' and help you specially for beginners like me. i need comments,
' suggestions, or criticism about my little app...
' my credits goes to my friends Cris "Coding Genius" Waddell and...
' Yariv Sarafraz for helping me learn about Winsock, thanx guys.
' Good Luck and God Bless... happy coding...

' thanx to PSC.Com and thanx to you for downloading this program...

'==============================================================================

Option Explicit

' define a global variable...
Dim svrPort As String
Dim msg As String

Private Sub cmdListen_Click()
    ' set the value for global variable...
    svrPort = txtServerPort.Text
    ' close any current open connection...
    wskServer.Close
    ' set the Port to listen...
    wskServer.LocalPort = svrPort
    ' start listening...
    wskServer.Listen
    ' set the Server status...
    lblStatus.Caption = "Waiting for connection..."
    ' disable the Port Textbox for any input...
    txtServerPort.Enabled = False
    ' disable the Listen Button and enable the Stop Button...
    cmdListen.Enabled = False
    cmdStop.Enabled = True
End Sub

Private Sub cmdSend_Click()
    ' set the value of "msg" variable...
    msg = txtMsg.Text
    ' make a condition...
    If wskServer.State = sckConnected Then
        ' add the message to the ChatBox...
        txtChat.Text = txtChat & vbNewLine & "<Server:> " & msg
        ' send the msg to client...
        wskServer.SendData msg
        DoEvents
        ' clear the txtMsg.text...
        txtMsg.Text = ""
    Else
        txtChat.Text = "* No connection found..."
    End If
End Sub

Private Sub cmdStop_Click()
    ' close the server and stop listening...
    wskServer.Close
    ' update the server status...
    lblStatus.Caption = "Server is closed..."
    ' enable the Port Textbox...
    txtServerPort.Enabled = True
    ' disable the Stop Button and enable the Listen Button...
    cmdStop.Enabled = False
    cmdListen.Enabled = True
    ' display a "Closed message" in the ChatBox...
    txtChat.Text = "* Session Closed..."
End Sub

Private Sub Form_Load()
    ' display your local IP Address...
    txtServerIP.Text = wskServer.LocalIP
    ' disable IP Address Textbox for any input...
    txtServerIP.Enabled = False
    ' display the default listening Port...
    txtServerPort.Text = "12345"
    ' disable the Stop Button by default...
    cmdStop.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' confirm the user to exit the program...
    If MsgBox("Are you sure to Exit the program?", vbQuestion + vbYesNo + vbDefaultButton1, "Exit") = vbNo Then
        Cancel = 1
    Else
        Unload Me
        End
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ' display a About Info...
    MsgBox "Simple Chat Program" & vbNewLine & "Vanni Alora" & vbNewLine & _
            "vanjo08@msn.com" & vbNewLine & vbNewLine & _
            "http://www.planet-source-code.com", vbInformation, "About"
End Sub

Private Sub txtChat_Change()
    ' this will always focus on the new line in ChatBox...
    txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub wskServer_Close()
    ' if the client is disconnected and try to connect again, then do this...
    If wskServer.State <> sckClosed Then wskServer.Close
    ' call the Listen command...
    cmdListen_Click
    ' display a message that the Client is Disconnected...
    txtChat.Text = "* The Client have Quit the Session..."
End Sub

Private Sub wskServer_ConnectionRequest(ByVal requestID As Long)
    ' close any current connection...
    wskServer.Close
    ' if the client try to connect, then do this...
    wskServer.Accept requestID
    ' display the Successfull connection status...
    lblStatus.Caption = "Connection Success..."
    ' display something in the ChatBox...
    txtChat.Text = "* The Client is Successfully Connected..."
End Sub

Private Sub wskServer_DataArrival(ByVal bytesTotal As Long)
    ' if the client send a data, then do this...
    wskServer.GetData msg
    DoEvents
    ' display it on the ChatBox...
    txtChat.Text = txtChat.Text & vbNewLine & "<Client:> " & msg
End Sub

Private Sub wskServer_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' if any error occured, then display a message...
    MsgBox "Unexpected error occured...", vbCritical + vbOKOnly, "Error"
End Sub

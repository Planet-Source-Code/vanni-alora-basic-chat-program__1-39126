VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat - Client"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   735
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
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   240
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   300
      Left            =   3840
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
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
   Begin VB.TextBox txtServerIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "< SERVER IP >"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtServerPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "< SERVER PORT >"
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   300
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   300
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Caption         =   "Ready to connect..."
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   540
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Server IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   165
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Server Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   525
      Width           =   975
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
Dim svrIP As String
Dim msg As String

Private Sub cmdConnect_Click()
    ' set the value for global variables...
    svrPort = txtServerPort.Text
    svrIP = txtServerIP.Text
    ' close any current open connection...
    wskClient.Close
    ' set the server IP and Port to connect...
    wskClient.RemoteHost = svrIP
    wskClient.RemotePort = svrPort
    ' connect to server...
    wskClient.Connect
    ' disable the textboxes and Connect button...
    txtServerPort.Enabled = False
    txtServerIP.Enabled = False
    cmdConnect.Enabled = False
    ' enable the Disconnect button...
    cmdDisconnect.Enabled = True
    ' display the client status...
    lblStatus.Caption = "Connecting to server..."
End Sub

Private Sub cmdDisconnect_Click()
    ' close the current connection...
    wskClient.Close
    ' update the client status...
    lblStatus.Caption = "Disconnected to server..."
    ' enable the Textboxes and Connect button...
    txtServerPort.Enabled = True
    txtServerIP.Enabled = True
    cmdConnect.Enabled = True
    ' disable the Disconnect button...
    cmdDisconnect.Enabled = False
    ' display a "Disconnected State" in the ChatBox...
    txtChat.Text = "* Disconnected..."
End Sub

Private Sub cmdSend_Click()
    ' set the value of "msg" variable...
    msg = txtMsg.Text
    ' make a condition...
    If wskClient.State = sckConnected Then
        ' add the message to the ChatBox...
        txtChat.Text = txtChat & vbNewLine & "<Client:> " & msg
        ' send the msg to client...
        wskClient.SendData msg
        DoEvents
        ' clear the txtMsg.text...
        txtMsg.Text = ""
    Else
        txtChat.Text = "* Not connected to server..."
    End If
End Sub

Private Sub Form_Load()
    ' set and display the default port and ip...
    txtServerPort.Text = "12345"
    txtServerIP.Text = wskClient.LocalIP
    ' disable the Disconnect button by default...
    cmdDisconnect.Enabled = False
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

Private Sub wskClient_Close()
    ' if the client is disconnecting to server or server is disconnected...
    If wskClient.State <> sckClosed Then wskClient.Close
    ' display a message...
    MsgBox "Connection to server lost...", vbInformation + vbOKOnly, "Error 2"
    ' call the Disconnect Click button event...
    cmdDisconnect_Click
End Sub

Private Sub wskClient_Connect()
    ' if the client is connected to the server successfully...
    If wskClient.State <> sckClosed Then
        ' update the status...
        lblStatus.Caption = "Connected to server..."
    End If
    ' display a "Welcome Message from Server"...
    txtChat.Text = "* Welcome To Chat, Feel Free To Talk Anything..."
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
    ' if the server send a data, then do this...
    wskClient.GetData msg
    ' display it on the ChatBox...
    txtChat.Text = txtChat.Text & vbNewLine & "<Server:> " & msg
    DoEvents
End Sub

Private Sub wskClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' if unexpected error occured, then do this...
    MsgBox "Could not connect to server...", vbInformation + vbOKOnly, "Error 1"
    ' call the Disconnect Click button event...
    cmdDisconnect_Click
End Sub

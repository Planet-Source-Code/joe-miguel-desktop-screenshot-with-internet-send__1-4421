VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   Caption         =   "NetHPLC"
   ClientHeight    =   4575
   ClientLeft      =   6465
   ClientTop       =   6105
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   Begin VB.TextBox txtRemotePort 
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Text            =   "1002"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtLocalPort 
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Text            =   "1001"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   840
   End
   Begin MSWinsockLib.Winsock sckTCP 
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton go 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox IP 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3855
   End
   Begin VB.PictureBox picsource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   2400
      Picture         =   "screenshot v3.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   3
      Top             =   1320
      Width           =   4830
   End
   Begin VB.CommandButton Process 
      Caption         =   "Process"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address for client or blank for server"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' By Joe Miguel!  joe_miguel@hotmail.com

' Winsock state commands
Enum sckState
    sckClosed = 0
    sckOpen = 1
    sckListening = 2
    sckConnectionPending = 3
    sckResolvingHost = 4
    sckHostResolved = 5
    sckConnecting = 6
    sckConnected = 7
    sckClosing = 8
    sckError = 9
End Enum

Dim Mode As String
Dim Busy As Boolean 'true when busy sending data

Private Sub Form_Load()
'setup the form
    MainForm.go.BackColor = RGB(0, 255, 0)
    MainForm.go.Caption = "Go!"
    MainForm.Caption = "NetHPLC"
    sckTCP.RemoteHost = sckTCP.LocalIP
    MainForm.IP.Text = sckTCP.RemoteHost
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub go_Click()
    sckTCP.Close

If MainForm.go.Caption = "Go!" Then
'turn on machine, change button to stop
    MainForm.go.BackColor = RGB(255, 0, 0)
    MainForm.go.Caption = "Stop"
    MainForm.SetFocus
        
    MainForm.IP.Enabled = False
    
    If IP <> "" Then
        MainForm.Timer.Enabled = True
        MainForm.Timer.Interval = 100   '10000ms = 10s
        
        Mode = "Server"
        MainForm.Caption = "NetHPLC - Server Mode"
        sckTCP.LocalPort = txtLocalPort.Text
        sckTCP.RemotePort = txtRemotePort.Text
        sckTCP.RemoteHost = IP.Text

    Else
        MainForm.Timer.Enabled = True
        MainForm.Timer.Interval = 100   '10000ms = 10s

        Mode = "Client"
        MainForm.Caption = "NetHPLC - Client Mode"
        sckTCP.LocalPort = txtRemotePort.Text
        sckTCP.RemotePort = txtLocalPort.Text
        sckTCP.Listen
    End If
    

Else
' change to go!
    MainForm.go.BackColor = RGB(0, 255, 0)
    MainForm.go.Caption = "Go!"
    MainForm.SetFocus
    
    MainForm.Timer.Enabled = False
    
    MainForm.IP.Enabled = True
    
    Mode = ""
    
    MainForm.Caption = "NetHPLC  by  joe_miguel@hotmail.com"
    
    sckTCP.Close
    
    Exit Sub
End If

End Sub

Private Sub Process_Click()
'Add this code to the form_load event
'or whatever you want to make it occur
'Get the hWnd of the desktop
DeskhWnd& = GetDesktopWindow()

'BitBlt needs the DC to copy the image. So, we
'need the GetDC API.
deskDC& = GetDC(DeskhWnd&)

'Copy whole screen into picture box!
StretchBlt picsource.hdc, 0, 0, _
                picsource.Width, _
                picsource.Height, _
                deskDC&, 0, 0, _
                Screen.Width \ Screen.TwipsPerPixelX, _
                Screen.Height \ Screen.TwipsPerPixelY, _
                SRCCOPY
picsource.Refresh


End Sub

Private Sub sckTCP_ConnectionRequest(ByVal RequestID As Long)
    sckTCP.Close
    sckTCP.Accept RequestID
    MainForm.Caption = "NetHPLC - Client Mode : Accepting Connection"
End Sub

Private Sub sckTCP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckTCP.Close
End Sub


Private Sub sckTCP_SendComplete()
    Busy = False
End Sub

Private Sub Timer_Timer()
    'tens of seconds
    
    Const PacketSize = 1024
    
    Static Seconds As Long
    
    Dim Buffer As Variant
    Dim Buffer2() As Byte
    Dim Buffer3() As Byte
    
    Dim Temp As Long
    
    Dim X As Long
    Dim ArrayCounter As Long
    Dim Y As Byte
    Dim a As Long
    Dim Test As Variant
    
If Mode = "Server" Then
    
    'Allow only one instance
    MainForm.Timer.Enabled = False
    
    Seconds = Seconds + 1
    
'    If Seconds * 10 >= 20 Then
    If Seconds * 10 >= 10 Then
    '60 seconds have pased
        Seconds = 0 ' reset seconds counter
    'set to server mode
    
    
    
    'Get the hWnd of the desktop
    DeskhWnd& = GetDesktopWindow()
'BitBlt needs the DC to copy the image. So, we
'need the GetDC API.
    deskDC& = GetDC(DeskhWnd&)
    StretchBlt picsource.hdc, 0, 0, _
        picsource.Width, _
        picsource.Height, _
        deskDC&, 0, 0, _
        Screen.Width \ Screen.TwipsPerPixelX, _
        Screen.Height \ Screen.TwipsPerPixelY, _
        SRCCOPY
        
        Buffer = GetDIB(picsource)
    
    Open App.Path & "\temp.txt" For Binary As #1
        Put #1, , Buffer
    Close 1
    Open App.Path & "\temp.txt" For Binary As #1
        ReDim Buffer2(1 To LOF(1))
        For X = 1 To LOF(1)
            Get #1, , Y
            Buffer2(X) = Y
'            DoEvents
        Next X
    Close 1
           
        DoEvents                    ' required to prevent some errors
        sckTCP.Close
        sckTCP.Connect              ' request connection

'put back in
        Do Until sckTCP.State = sckState.sckConnected Or sckTCP.State = sckState.sckClosed
            MainForm.Caption = "NetHPLC - Server Mode : " & sckTCP.State
            DoEvents
        Loop
        If sckTCP.State = sckState.sckClosed Then Exit Sub

        Temp = UBound(Buffer2)
        
        Busy = True
        sckTCP.SendData Temp
        Do Until Busy = False: DoEvents: Loop
        sckTCP.SendData Buffer2
    End If
Exit Sub
Else ' client
    
    Dim Upper As Long
    Dim Counter As Long
    
    MainForm.Timer.Enabled = False
   
    Do Until sckTCP.State = sckState.sckConnected: DoEvents: Loop
    Do Until sckTCP.BytesReceived >= 3: DoEvents: Loop
           
    sckTCP.GetData X, vbLong
        
    Do Until sckTCP.BytesReceived = X
        MainForm.Caption = "NetHPLC - Client Mode : Bytes Received " & sckTCP.BytesReceived & " of " & X
        DoEvents
    Loop
    
    sckTCP.GetData Buffer2, vbByte + vbArray, X
    
    Open App.Path & "\temp.txt" For Binary As #1
        Put #1, , Buffer2
    Close 1
    Open App.Path & "\temp.txt" For Binary As #1
        Get #1, , Buffer
    Close 1
    Buffer3 = Buffer
   
    SetDIB Buffer3, picsource
    
End If
End Sub




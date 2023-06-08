VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPv4StreamSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const COMMAND_ERROR = -1
Private Const RECV_ERROR = -1
Private Const NO_ERROR = 0
Private m_socketId As Long

''' <summary> Initializes an internet protocol Stream Socket with a socket id. </summary>
Private Sub Class_Initialize()

    ' see that the singleton windock class is instantiated
    On Error Resume Next
    
    ' this is required to initialize Winsock.  It will only ran once.
    Winsock.Initialize
    
    On Error GoTo 0

    ' Create a new socket
    
    m_socketId = wsock32.CreateSocket(wsock32.AF_INET, wsock32.SOCK_STREAM, 0)
    If m_socketId = wsock32.INVALID_SOCKET Then
        MsgBox ("ERROR: wsock32.CreateSocket returned " + Str$(m_socketId))
    Else
        wnsock.Register Me
    End If
End Sub

''' <summary> terminates the socket and releases its Winsock registration. </summary>
Private Sub Class_Terminate()

    ' see that the singleton windock class is instantiated
    On Error Resume Next
    
    ' unregister this socket.
    Winsock.UnRegister Me
    
    On Error GoTo 0

End Sub

''' <summary> Returns the socket id. </summary>
''' <returns> The socket identity. </returns>
Public Property Get SocketId() As Long
    SocketId = m_socketId
End Property

''' <summary> Opens a socket connection and sets and returns the socket id. </summary>
''' <param name=""> </param>
''' <param name=""> </param>
''' <returns> The socket identity. </returns>
Function OpenConnection(ByVal host As String, ByVal port As Integer) As Integer
   
    If host = "localhost" Then
        host = "127.0.0.1"
    End If
    
    ' Open a connection to a server
    ' https://www.allscoop.com/tcp-listen.php
    
    Dim address As wsock32.sockaddr_in
    address.sin_addr.s_addr = wsock32.inet_addr(host)
    address.sin_family = wsock32.AF_INET
    address.sin_port = wsock32.htons(port)
    
    Dim connectResult As Long
    connectResult = wsock32.connect(m_socketId, address, Len(address))
    If connectResult < 0 Then
        MsgBox ("ERROR: connection failed = " + Str$(connectResult))
        OpenConnection = COMMAND_ERROR
        Exit Function
    End If
    
    OpenConnection = m_socketId

End Function

''' <summary> Closes the socket connection and zeros the socket id. </summary>
Function CloseConnection() As Long

    Dim result As Long
    result = closesocket(m_socketId)
    m_socketId = 0
    If result < 0 Then
        MsgBox ("ERROR: closing connection = " + Str$(result))
    End If
    CloseConnection = result

End Function


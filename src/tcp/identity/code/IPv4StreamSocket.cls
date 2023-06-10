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

''' <summary> Creates a socket that is bound to a specific transport service provider. </summary>
''' <remarks> </remarks>
''' <param name="addressFamily"> [in] The address family specification.
''' The values currently supported are <see cref="AF_INET"/> or <see cref="AF_INET6"/>, which are the Internet
''' address family formats for IPv4 and IPv6. Other options for address family (AF_NETBIOS for use with NetBIOS,
''' for example) are supported if a Windows Sockets service provider for the address family is installed.
''' Note that the values for the AF_ address family and PF_ protocol family constants are identical
''' (for example, AF_INET and PF_INET), so either constant can be used.
''' </param>
''' <param name="socketType">    [in] The type specification for the new socket.
''' In Windows Sockets 1.1, the only possible socket types are SOCK_DGRAM and SOCK_STREAM. </param>
''' <param name="protocol"> The protocol to be used. The possible options for the protocol parameter are specific
''' to the address family and socket type specified. </param>
''' <returns> If no error occurs, socket returns a descriptor referencing the new socket.
''' Otherwise, a value of INVALID_SOCKET is returned, and a specific error code can be retrieved by
''' calling WSAGetLastError. <returns>
Private Declare PtrSafe Function CreateSocket Lib "wsock32.dll" Alias "socket" (ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As Long

Private m_socketId As Long
Private m_connected As Boolean
Private m_timeout As Integer

''' <summary> Initializes an internet protocol Stream Socket with a socket id. </summary>
Private Sub Class_Initialize()

    ' see that the singleton windock class is instantiated
    On Error Resume Next
    
    ' this is required to initialize Winsock.  It will only ran once.
    Winsock.Initialize
    
    On Error GoTo 0

    ' Create a new socket
    
    m_socketId = WinsockApi.CreateIPv4StreamSocket()
    If m_socketId = wsock32.INVALID_SOCKET Then
        Err.Raise ErrorCodes.SocketCreationError.Code, "IdentityReader.IPv4StreamSocket.Class_Initialize", _
                  ErrorCodes.SocketCreationError.Description & " Last error: " & CStr(Winsock.lastError)
    Else
        Winsock.Register Me
        
        ' set the default timeout
        SetTimeout 500
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
Public Property Get socketId() As Long
    socketId = m_socketId
End Property

''' <summary> Returns the connection state of the socket. </summary>
''' <returns> True if the socket is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    Connected = m_connected
End Property

''' <summary> Opens a socket connection and sets and returns the socket id. </summary>
''' <param name="host"> An IPv4 dotted-decimal host address. </param>
''' <param name="port"> The port that the server is listening on. </param>
''' <returns> True if the socket is connected. </returns>
Public Function OpenConnection(ByVal host As String, ByVal port As Integer) As Boolean
   
    m_connected = False
    If host = "localhost" Then
        host = "127.0.0.1"
    End If
    
    ' Open a connection to a server
    ' can use this: https://www.allscoop.com/tcp-listen.php to listen.
    
    Dim address As wsock32.sockaddr_in
    
    ' set the binary address
    address.sin_addr.s_addr = wsock32.inet_addr(host)
    
    address.sin_family = wsock32.AF_INET
    
    ' set the port in network byte order
    address.sin_port = wsock32.htons(port)
    
    Dim connectResult As Long
    connectResult = wsock32.connect(m_socketId, address, Len(address))
    If connectResult = wsock32.SOCKET_ERROR Then
        Err.Raise ErrorCodes.SocketConnectionError.Code, "IdentityReader.IPv4StreamSocket.OpenConnection", _
                  ErrorCodes.SocketConnectionError.Description & " Last error: " & CStr(Winsock.lastError)
        m_connected = False
    Else
        m_connected = True
    End If
    
    OpenConnection = m_connected

End Function

''' <summary> Closes the socket connection and zeros the socket id. </summary>
''' <returns> True if socket disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    Dim result As Long
    result = wsock32.closesocket(m_socketId)
    m_socketId = 0
    If result = wsock32.SOCKET_ERROR Then
        Err.Raise ErrorCodes.SocketDisconnectionError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                  ErrorCodes.SocketDisconnectionError.Description & " Last error: " & CStr(Winsock.lastError)
        CloseConnection = False
    Else
        m_connected = False
        CloseConnection = True
    End If

End Function

''' <summary> Sends a terminated message to the server. </summary>
''' <param name="command"> The message. </param>
''' <returns> The number of bytes that were sent to the server. </returns>
Public Function SendCommand(ByVal command As String) As Integer

    If StringExtensions.IsNullOrEmpty(command) Then
        SendCommand = 0
    Else
        count = wsock32.send(Me.socketId, ByVal command, Len(command), 0)
        
        If count < 0 Then
            Err.Raise ErrorCodes.SocketSendError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                      ErrorCodes.SocketSendError.Description & " Last error: " & CStr(Winsock.lastError)
            SendCommand = COMMAND_ERROR
        Else
            SendCommand = count
        End If
    End If

End Function

''' <summary> Receives a message from the server. </summary>
''' <param name="buffer"> Where the message will go. </param>
''' <param name="maxLength"> The maximum number of bytes to read. </param>
''' <param name="termination"> [Optional] The character the signifies the end of the stream. </param>
''' <returns> If no error, the number of byte that were sent to the server.
''' Otherwise, RECV_ERROR.
''' </returns>
Public Function Receive(buffer As String, ByVal maxLength As Integer, _
                        Optional ByVal termination As String = vbLf) As Integer

    Dim c As String * 1
    Dim length As Integer
    Dim count As Integer
    
    buffer = ""
    While length < maxLength
    
        DoEvents
        
        c = ""
        Dim l As Long
        l = Len(c)
        count = wsock32.recv(socketId, c, l, 0)
        
        If count < 1 Then
            Receive = RECV_ERROR
            buffer = Chr$(0)
            Exit Function
        End If
        
        If c = termination Then
           buffer = buffer + Chr$(0)
           Receive = length
           Exit Function
        End If
        
        length = length + count
        buffer = buffer + c
        
    Wend
    
    Receive = RECV_ERROR
    
End Function

''' <summary> Sets the socket timeout. </summary>
''' <param name="timeout"> The timeout interval in milliseconds. </param>
Public Sub SetTimeout(ByVal interval As Long)

    Dim result As Long
    result = wsock32.setsockopt(m_socketId, wsock32.SOL_SOCKET, wsock32.SO_RCVTIMEO, interval, 4)

    If result = wsock32.SOCKET_ERROR Then
        Err.Raise ErrorCodes.SocketOptionSettingError.Code, "IdentityReader.IPV4StreamSocket.SetTimeout", _
                  ErrorCodes.SocketOptionSettingError.Description & " Last error: " & CStr(Winsock.lastError)
    Else
        m_timeout = interval
    End If
    
End Sub

''' <summary> Gets the socket timeout in milliseconds. </summary>
Public Property Get Timeout() As Long
    Timeout = m_timeout
End Property

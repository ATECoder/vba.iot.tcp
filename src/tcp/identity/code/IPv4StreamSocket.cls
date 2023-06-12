VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPv4StreamSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const COMMAND_ERROR = -1
Private Const RECV_ERROR = -1
Private Const NO_ERROR = 0

''' <summary> Permits an incoming connection attempt on a socket. </summary>
''' <remarks> </remarks>
''' <param name="s">                   [in] A descriptor that identifies a socket that has been placed in a listening state with the listen function.
'''                                    The connection is actually made with the socket that is returned by accept.</param>
''' <param name="clientAddress">       [out] An optional pointer to a buffer that receives the address of the connecting entity,
'''                                    as known to the communications layer. The exact format of the addr parameter is determined by the address family
'''                                    that was established when the socket from the ws32_SocketAddress structure was created.</param>
''' <param name="clientAddressLength"> [in, out] An optional pointer to an integer that contains the length of structure pointed to by
'''                                    the addr parameter. </param>
''' <returns> If no error occurs, accept returns a value of type SOCKET that is a descriptor for the new socket.
''' This returned value is a handle for the socket on which the actual connection is made. Otherwise, a value of
''' ws32_INVALID_SOCKET is returned, and a specific error code can be retrieved by calling WSAGetLastError.
''' The integer referred to by clientAddressLength initially contains the amount of space pointed to by clientAddress. On return it
''' will contain the actual length in bytes of the address returned. <returns>
Private Declare PtrSafe Function accept_ Lib "wsock32.dll" Alias "accept" ( _
        ByVal s As Long, clientAddress As wsock32.ws32_SocketAddress, clientAddressLength As Integer) As Long

''' <summary> Associates a local address with a socket. </summary>
''' <remarks> </remarks>
''' <param name="s">             [in] A descriptor identifying an unbound socket. </param>
''' <param name="address">       [in] A pointer to a ws32_SocketAddress_in structure of the local address
'''                              to assign to the bound socket . </param>
''' <param name="addressLength"> [in] The length, in bytes, of the value pointed to by address. </param>
''' <returns> Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function bind_ Lib "wsock32.dll" Alias "bind" ( _
        ByVal s As Long, address As wsock32.ws32_SocketAddress_in, ByVal addressLength As Integer) As Long

''' <summary> Closes an existing socket. </summary>
''' <remarks> </remarks>
''' <param name="s"> [in] A descriptor identifying the socket to close. </param>
''' <returns> If no error occurs, closesocket returns zero. Otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function closesocket_ Lib "wsock32.dll" Alias "closesocket" ( _
        ByVal s As Long) As Long

''' <summary> Establishes a connection to a specified socket. </summary>
''' <remarks> </remarks>
''' <param name="s">          [in] A descriptor identifying an unconnected socket. </param>
''' <param name="address">    [in] A pointer to the <see cref="ws32_SocketAddress_in"/> structure to which the
'''                           connection should be established. </param>
''' <param name="addressLen"> [in] The length, in bytes, of the ws32_SocketAddress structure pointed to by the
'''                           <paramref name="address"/> parameter. </param>
''' <returns> Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function connect_ Lib "wsock32.dll" Alias "connect" ( _
        ByVal s As Long, ByRef address As wsock32.ws32_SocketAddress_in, ByVal addressLen As Long) As Long

''' <summary> Places a socket in a state in which it is listening for an incoming connection. </summary>
''' <remarks> winsock2 only? </remarks>
''' <param name="s">       [in] A descriptor identifying a bound, unconnected socket. </param>
''' <param name="backlog"> [in] The maximum length of the queue of pending connections. If set to SOMAXCONN,
'''                         the underlying service provider responsible for socket s will set the backlog to a
'''                         maximum reasonable value. If set to SOMAXCONN_HINT(N) (where N is a number), the
'''                         backlog value will be N, adjusted to be within the range (200, 65535). Note that
'''                         SOMAXCONN_HINT can be used to set the backlog to a larger value than possible with SOMAXCONN.
'''                         SOMAXCONN_HINT is only supported by the Microsoft TCP/IP service provider. There is no
'''                         standard provision to obtain the actual backlog value.
''' </param>
''' <returns> Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function listen_ Lib "wsock32.dll" Alias "listen" ( _
        ByVal s As Long, ByVal backlog As Integer) As Long

''' <summary> Sends data on a connected socket. </summary>
''' <remarks> </remarks>
''' <param name="s">            [in] A descriptor identifying a connected socket. </param>
''' <param name="buffer">       [in] A pointer to a buffer containing the data to be transmitted. </param>
''' <param name="bufferLength"> [in] The length, in bytes, of the data in buffer pointed to by the buffer parameter. </param>
''' <param name="flags">        [in] A set of flags that specify the way in which the call is made. This parameter is
'''                             constructed by using the bitwise OR operator with any of the following values:
'''                             ws32_MSG_DONTROUTE: Specifies that the data should not be subject to routing. A Windows Sockets
'''                             service provider can choose to ignore this flag.
'''                             ws32_MSG_OOB: Sends OOB data (stream-style socket such as ws32_SOCK_STREAM only).
'''                             </param>
''' <returns> Returns the total number of bytes sent, which can be less than the number requested to be sent
''' in the len parameter. If an error occured, a value of ws32_SOCKET_ERROR is returned, and a specific error code
''' can be retrieved by calling WSAGetLastError.
''' <returns>
Private Declare PtrSafe Function send_ Lib "wsock32.dll" Alias "send" ( _
        ByVal s As Long, buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long

''' <summary> Sets a socket option. </summary>
''' <remarks> </remarks>
''' <param name="s">       [in] A descriptor that identifies a socket. </param>
''' <param name="level">   [in] The level at which the option is defined (for example, ws32_SOL_SOCKET). </param>
''' <param name="optname"> [in] The socket option for which the value is to be set (for example, ws32_SO_BROADCAST).
'''                        The optname parameter must be a socket option defined within the specified level,
'''                        or behavior is undefined. </param>
''' <param name="optval">  [in] A pointer to the buffer in which the value for the requested option is specified. </param>
''' <param name="optlen">  [in] The size, in bytes, of the buffer pointed to by the optval parameter. </param>
''' <returns> If no error occurs, setsockopt returns zero. Otherwise, a value of ws32_SOCKET_ERROR is returned,
''' and a specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function setsockopt_ Lib "wsock32.dll" Alias "setsockopt" ( _
        ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Integer) As Long

''' <summary> Receives data from a connected socket or a bound connectionless socket. </summary>
''' <remarks>
''' The flags parameter can be used to influence the behavior of the function invocation beyond the options specified
''' for the associated socket. The semantics of this function are determined by the socket options and the flags parameter.
''' The possible value of flags parameter is constructed by using the bitwise OR operator with any of the following values:
''' ws32_MSG_PEEK    Peeks at the incoming data. The data is copied into the buffer, but is not removed from the input queue.
''' ws32_MSG_OOB     Processes Out Of Band (OOB) data.
''' ws32_MSG_WAITALL The receive request will complete only when one of the following events occurs:
'''                  The buffer supplied by the caller is completely full.
'''                  The connection has been closed.
'''                  The request has been canceled or an error occurred.
''' Note that if the underlying transport does not support ws32_MSG_WAITALL, or if the socket is in a non-blocking mode,
''' then this call will fail with WSAEOPNOTSUPP. Also, if ws32_MSG_WAITALL is specified along with ws32_MSG_OOB, ws32_MSG_PEEK, or
''' ws32_MSG_PARTIAL, then this call will fail with WSAEOPNOTSUPP. This flag is not supported on datagram sockets
''' or message-oriented sockets.
''' </remarks>
''' <param name="s">            [in] A descriptor identifying a connected socket. </param>
''' <param name="buffer">       [out] A pointer to the buffer to receive the incomming data. </param>
''' <param name="bufferLength"> [in] The length, in bytes, of the data in buffer pointed to by the buffer parameter. </param>
''' <param name="flags">        [in] A set of flags that influences the behavior of this function. </param>
''' <returns>
''' If no error occurs, recv returns the number of bytes received and the buffer pointed to by the buffre parameter will
''' contain this data received. If the connection has been gracefully closed, the return value is zero.
''' Otherwise, a value of ws32_SOCKET_ERROR is returned, and a specific error code can be retrieved by calling WSAGetLastError.
''' <returns>
Private Declare PtrSafe Function recv_ Lib "wsock32.dll" Alias "recv" ( _
        ByVal s As Long, ByVal buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long


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
    If m_socketId = wsock32.ws32_INVALID_SOCKET Then
        Err.Raise UserDefinedErrors.SocketCreationError.Code, "IdentityReader.IPv4StreamSocket.Initialize", _
                  UserDefinedErrors.SocketCreationError.Description & " Last error: " & CStr(Winsock.lastError)
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
Public Property Get SocketId() As Long
    SocketId = m_socketId
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
    
    Dim address As wsock32.ws32_SocketAddress_in
    
    ' set the binary address
    address.sin_addr.s_addr = WinsockApi.ToInetAddress(host)
    
    address.sin_family = wsock32.ws32_AF_INET
    
    ' set the port in network byte order
    address.sin_port = WinsockApi.ToInetByteOrder(port)
    
    Dim connectResult As Long
    connectResult = connect_(m_socketId, address, Len(address))
    If connectResult = wsock32.ws32_SOCKET_ERROR Then
        Err.Raise UserDefinedErrors.SocketConnectionError.Code, "IdentityReader.IPv4StreamSocket.OpenConnection", _
                  UserDefinedErrors.SocketConnectionError.Description & " Last error: " & CStr(Winsock.lastError)
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
    result = closesocket_(m_socketId)
    m_socketId = 0
    If result = wsock32.ws32_SOCKET_ERROR Then
        Err.Raise UserDefinedErrors.SocketDisconnectionError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                  UserDefinedErrors.SocketDisconnectionError.Description & " Last error: " & CStr(Winsock.lastError)
        CloseConnection = False
    Else
        m_connected = False
        CloseConnection = True
    End If

End Function

''' <summary> Sends a terminated message to the server. </summary>
''' <param name="command"> The message. </param>
''' <returns> The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal command As String) As Integer

    If StringExtensions.IsNullOrEmpty(command) Then
        SendMessage = 0
    Else
        Dim commandLength As Integer
        commandLength = Len(command)
        Dim sentLength As Integer
        sentLength = send_(Me.SocketId, ByVal command, Len(command), 0)
        
        If sentLength < 0 Then
            ' if the sent length is negative, we have an API error.
            Err.Raise UserDefinedErrors.SocketSendError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                      UserDefinedErrors.SocketSendError.Description & " Last error: " & CStr(Winsock.lastError)
            SendMessage = COMMAND_ERROR
        ElseIf sentLength <> commandLength Then
        Else
            SendMessage = sentLength
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
        count = recv_(Me.SocketId, c, l, 0)
        
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
    result = setsockopt_(m_socketId, wsock32.ws32_SOL_SOCKET, wsock32.ws32_SO_RCVTIMEO, interval, 4)

    If result = wsock32.ws32_SOCKET_ERROR Then
        Err.Raise UserDefinedErrors.SocketOptionSettingError.Code, "IdentityReader.IPV4StreamSocket.SetTimeout", _
                  UserDefinedErrors.SocketOptionSettingError.Description & " Last error: " & CStr(Winsock.lastError)
    Else
        m_timeout = interval
    End If
    
End Sub

''' <summary> Gets the socket timeout in milliseconds. </summary>
Public Property Get Timeout() As Long
    Timeout = m_timeout
End Property

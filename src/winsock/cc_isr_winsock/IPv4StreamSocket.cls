VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPv4StreamSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SEND_MESSAGE_ERROR = -1
Private Const SEND_PARTIAL_MESSAGE_ERROR = -2
Private Const RECEIVE_ERROR = -1
Private Const DEFAULT_TIMEOUT_MS = 500

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByRef eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   Permits an incoming connection attempt on a socket. </summary>
''' <param name="s">                     [in] A descriptor that identifies a socket that has been placed in a listening state with the listen function.
'''                                      The connection is actually made with the socket that is returned by accept. </param>
''' <param name="clientAddress">         [out] An optional pointer to a buffer that receives the address of the connecting entity,
'''                                      as known to the communications layer. The exact format of the address parameter is determined by the address family
'''                                      that was established when the socket from the ws32_SocketAddress structure was created. </param>
''' <param name="clientAddressLength">   [in, out] An optional pointer to an integer that contains the length of structure pointed to by
'''                                      the address parameter. </param>
''' <returns>   If no error occurs, accept returns a value of type SOCKET that is a descriptor for the new socket.
''' This returned value is a handle for the socket on which the actual connection is made. Otherwise, a value of
''' ws32_INVALID_SOCKET is returned, and a specific error code can be retrieved by calling WSAGetLastError.
''' The integer referred to by clientAddressLength initially contains the amount of space pointed to by clientAddress. On return it
''' will contain the actual length in bytes of the address returned. <returns>
Private Declare PtrSafe Function accept_ Lib "wsock32.dll" Alias "accept" ( _
        ByVal s As Long, ByRef clientAddress As wsock32.ws32_SocketAddress, ByVal clientAddressLength As Integer) As Long

''' <summary>   Associates a local address with a socket. </summary>
''' <param name="s">               [in] A descriptor identifying an unbound socket. </param>
''' <param name="address">         [in] A pointer to a ws32_SocketAddress_in structure of the local address
'''                                to assign to the bound socket . </param>
''' <param name="addressLength">   [in] The length, in bytes, of the value pointed to by address. </param>
''' <returns>   Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function bind_ Lib "wsock32.dll" Alias "bind" ( _
        ByVal s As Long, ByRef address As wsock32.ws32_SocketAddress_in, ByVal addressLength As Integer) As Long

''' <summary>   Closes an existing socket. </summary>
''' <param name="s">   [in] A descriptor identifying the socket to close. </param>
''' <returns>   If no error occurs, closesocket returns zero. Otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function closesocket_ Lib "wsock32.dll" Alias "closesocket" ( _
        ByVal s As Long) As Long

''' <summary>   Establishes a connection to a specified socket. </summary>
''' <param name="s">            [in] A descriptor identifying an unconnected socket. </param>
''' <param name="address">      [in] A pointer to the <see cref="ws32_SocketAddress_in"/>   structure to which the
'''                             connection should be established. </param>
''' <param name="addressLen">   [in] The length, in bytes, of the ws32_SocketAddress structure pointed to by the
'''                             <paramref name="address"/>   parameter. </param>
''' <returns>   Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function connect_ Lib "wsock32.dll" Alias "connect" ( _
        ByVal s As Long, ByRef address As wsock32.ws32_SocketAddress_in, ByVal addressLen As Long) As Long

''' <summary>   Places a socket in a state in which it is listening for an incoming connection. </summary>
''' <param name="s">         [in] A descriptor identifying a bound, unconnected socket. </param>
''' <param name="backlog">   [in] The maximum length of the queue of pending connections. If set to SOMAXCONN,
'''                          the underlying service provider responsible for socket s will set the backlog to a
'''                          maximum reasonable value. If set to SOMAXCONN_HINT(N) (where N is a number), the
'''                          backlog value will be N, adjusted to be within the range (200, 65535). Note that
'''                          SOMAXCONN_HINT can be used to set the backlog to a larger value than possible with SOMAXCONN.
'''                          SOMAXCONN_HINT is only supported by the Microsoft TCP/IP service provider. There is no
'''                          standard provision to obtain the actual backlog value.
''' </param>
''' <returns>   Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function listen_ Lib "wsock32.dll" Alias "listen" ( _
        ByVal s As Long, ByVal backlog As Integer) As Long

''' <summary>   Sends data on a connected socket. </summary>
''' <param name="s">              [in] A descriptor identifying a connected socket. </param>
''' <param name="buffer">         [in] A pointer to a buffer containing the data to be transmitted. </param>
''' <param name="bufferLength">   [in] The length, in bytes, of the data in buffer pointed to by the buffer parameter. </param>
''' <param name="flags">          [in] A set of flags that specify the way in which the call is made. This parameter is
'''                               constructed by using the bitwise OR operator with any of the following values:
'''                               ws32_MSG_DONTROUTE: Specifies that the data should not be subject to routing. A Windows Sockets
'''                               service provider can choose to ignore this flag.
'''                               ws32_MSG_OOB: Sends OOB data (stream-style socket such as ws32_SOCK_STREAM only).
'''                               </param>
''' <returns>   Returns the total number of bytes sent, which can be less than the number requested to be sent
''' in the length parameter. If an error occurred, a value of ws32_SOCKET_ERROR is returned, and a specific error code
''' can be retrieved by calling WSAGetLastError.
''' <returns>
Private Declare PtrSafe Function send_ Lib "wsock32.dll" Alias "send" ( _
        ByVal s As Long, ByRef buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long

''' <summary>   Sets a socket option. </summary>
''' <param name="s">         [in] A descriptor that identifies a socket. </param>
''' <param name="level">     [in] The level at which the option is defined (for example, ws32_SOL_SOCKET). </param>
''' <param name="optname">   [in] The socket option for which the value is to be set (for example, ws32_SO_BROADCAST).
'''                          The optname parameter must be a socket option defined within the specified level,
'''                          or behavior is undefined. </param>
''' <param name="optval">    [in] A pointer to the buffer in which the value for the requested option is specified. </param>
''' <param name="optlen">    [in] The size, in bytes, of the buffer pointed to by the optval parameter. </param>
''' <returns>   If no error occurs, setsockopt returns zero. Otherwise, a value of ws32_SOCKET_ERROR is returned,
''' and a specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function setsockopt_ Lib "wsock32.dll" Alias "setsockopt" ( _
        ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Integer) As Long

''' <summary>   Receives data from a connected socket or a bound connectionless socket. </summary>
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
''' <param name="s">              [in] A descriptor identifying a connected socket. </param>
''' <param name="buffer">         [out] A pointer to the buffer to receive the incoming data. </param>
''' <param name="bufferLength">   [in] The length, in bytes, of the data in buffer pointed to by the buffer parameter. </param>
''' <param name="flags">          [in] A set of flags that influences the behavior of this function. </param>
''' <returns>
''' If no error occurs, receive returns the number of bytes received and the buffer pointed to by the buffer parameter will
''' contain this data received. If the connection has been gracefully closed, the return value is zero.
''' Otherwise, a value of ws32_SOCKET_ERROR is returned, and a specific error code can be retrieved by calling WSAGetLastError.
''' <returns>
Private Declare PtrSafe Function recv_ Lib "wsock32.dll" Alias "recv" ( _
        ByVal s As Long, ByVal buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long

Private m_socketPort As Long
Private m_socketHost As String
Private m_socketAddress As String
Private m_socketId As Long
Private m_connected As Boolean
Private m_listening As Boolean
Private m_receiveTimeout As Long
Private m_fdSet As wsock32.ws32_fd_set

''' <summary>   Initializes an internet protocol Stream Socket with a socket id. </summary>
Private Sub Class_Initialize()

    ' see that the singleton Winsock class is instantiated
    On Error Resume Next
    
    ' this is required to initialize Winsock.  It will only ran once.
    Winsock.Initialize
    
    m_socketHost = ""
    m_socketAddress = ""
    
    On Error GoTo 0

    ' Create a new socket
    
    m_socketId = WinsockApi.CreateIPv4StreamSocket()
    If m_socketId = wsock32.ws32_INVALID_SOCKET Then
        Err.Raise UserDefinedErrors.SocketCreationError.Code, "IdentityReader.IPv4StreamSocket.Initialize", _
                  UserDefinedErrors.SocketCreationError.Description & " Last error: " & CStr(Winsock.lastError)
    Else
        Winsock.Register Me
        
        ' set the receive timeout to its default value
        
        Me.SetReceiveTimeout DEFAULT_TIMEOUT_MS
    End If
    
End Sub

''' <summary>   terminates the socket and releases its Winsock registration. </summary>
Private Sub Class_Terminate()

    ' see that the singleton Winsock class is instantiated
    On Error Resume Next
    
    ' unregister this socket.
    Winsock.UnRegister Me
    
    On Error GoTo 0

End Sub

''' <summary>   Returns the port the server is listening to. </summary>
''' <value>   [Long] The port the server listens to. </value>
Public Property Get SocketPort() As Long
    SocketPort = m_socketPort
End Property

''' <summary>   Gets the IPv4 dotted-decimal host address. </summary>
''' <value>   [String] The IPv4 dotted-decimal host address. </value>
Public Property Get SocketHost() As String
    SocketHost = m_socketHost
End Property

''' <summary>   Gets the socket address consisting of the IPv4 dotted-decimal host
''' and the port the server is listening too. </summary>
''' <value>   [String] The IPv4 dotted-decimal host address and port delimited with a colon. </value>
Public Property Get SocketAddress() As String
    SocketAddress = m_socketAddress
End Property

''' <summary>   Returns a descriptor identifying the socket. </summary>
''' <value>   The socket identifying descriptor. </value>
Public Property Get SocketId() As Long
Attribute SocketId.VB_Description = "A descriptor identifying the socket"
Attribute SocketId.VB_UserMemId = 0
    SocketId = m_socketId
End Property

''' <summary>   Returns the connection state of the socket. </summary>
''' <value>   True if the socket is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
    Connected = m_connected
End Property

''' <summary>   Initializes an Inet stream socket using the provided socket id. </summary>
''' <param name="connectedSocketId">   The ID of a connected socket. </param>
''' <param name="timeoutMs">           [Optional, 500] The receive timeout interval in milliseconds. </param>
Public Sub Initialize(ByVal connectedSocketId As Long, Optional ByVal timeoutMs As Long = 500)
    m_socketId = connectedSocketId
    m_connected = True
    Me.SetReceiveTimeout timeoutMs
End Sub

''' <summary>   Opens a socket connection and sets and returns the socket id. </summary>
''' <param name="host">   An IPv4 dotted-decimal host address. </param>
''' <param name="port">   The port that the server is listening on. </param>
''' <returns>   True if the socket is connected. </returns>
Public Function OpenConnection(ByVal host As String, ByVal port As Long) As Boolean
   
    m_connected = False
    If StringExtensions.IsNullOrEmpty(host) Then
        Err.Raise UserDefinedErrors.InvalidArgumentError.Code, "IdentityReader.IPv4StreamSocket.OpenConnection", _
                  UserDefinedErrors.InvalidArgumentError.Description & " Argument: host"
    End If
    
    If host = "localhost" Then
        host = "127.0.0.1"
    End If
    
    ' Open a connection to a server
    ' can use this: https://www.allscoop.com/Tcp-listen.php to listen.
    
    Dim address As wsock32.ws32_SocketAddress_in
    
    ' set the binary address
    address.sin_addr.s_addr = WinsockApi.ToInetAddress(host)
    
    address.sin_family = wsock32.ws32_AF_INET
    
    ' set the port in network byte order
    address.sin_port = WinsockApi.ToInetByteOrder(port)
    
    Dim connectResult As Long
    connectResult = connect_(m_socketId, address, Len(address))
    If connectResult = wsock32.ws32_SOCKET_ERROR Then
        m_connected = False
        Err.Raise UserDefinedErrors.SocketConnectionError.Code, "IdentityReader.IPv4StreamSocket.OpenConnection", _
                  UserDefinedErrors.SocketConnectionError.Description & " Last error: " & CStr(Winsock.lastError)
    Else
        m_connected = True
        
        OnConnectionChanged Me.Connected

    End If
    
    m_socketPort = port
    m_socketHost = host
    m_socketAddress = host & ":" & CStr(port)
    
    OpenConnection = m_connected

End Function

''' <summary>   Closes the socket connection and zeros the socket id. </summary>
''' <returns>   True if socket disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    Dim result As Long
    result = closesocket_(m_socketId)
    m_socketId = 0
    If result = wsock32.ws32_SOCKET_ERROR Then
        CloseConnection = False
        Err.Raise UserDefinedErrors.SocketDisconnectionError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                  UserDefinedErrors.SocketDisconnectionError.Description & " Last error: " & CStr(Winsock.lastError)
    Else
        m_connected = False
        
        CloseConnection = True
        
        OnConnectionChanged Me.Connected
        
    End If

End Function

''' <summary>   Associates a local address with a socket. </summary>
''' <param name="address">   The INet address to bind to. </param>
''' <param name="port">   The port that the server is listening on. </param>
''' <returns>   True if the socket is bound. </returns>
Public Function BindTo(ByVal address As Long, ByVal port As Long) As Boolean

    Dim endpoint As wsock32.ws32_SocketAddress_in
    endpoint.sin_family = wsock32.ws32_AF_INET
    endpoint.sin_addr.s_addr = address
    endpoint.sin_port = WinsockApi.ToInetByteOrder(port)
    
    Dim bindResult As Long
    bindResult = bind_(Me.SocketId, endpoint, 16)
    
    If bindResult = wsock32.ws32_SOCKET_ERROR Then
        Err.Raise UserDefinedErrors.SocketBindError.Code, "IdentityReader.IPv4StreamSocket.BindTo", _
                  UserDefinedErrors.SocketBindError.Description & " Last error: " & CStr(Winsock.lastError)
        m_connected = False
    Else
        m_connected = True
    End If
    BindTo = Me.Connected
   
End Function

''' <summary>   Returns the listening state of the socket. </summary>
''' <value>   True if the socket is Listening; otherwise, False. </value>
Public Property Get Listening() As Boolean
    Listening = m_listening
End Property


''' <summary>   Places a socket in a state in which it is listening for an incoming connection. </summary>
''' <param name="backlog">   [Optional, 10] [in] The maximum length of the queue of pending connections. </para>
''' <returns>   True if the socket is Listening; otherwise, False. </returns>
Public Function StartListening(Optional ByVal backlog As Integer = 10) As Boolean

    Dim listenResult As Long
    listenResult = listen_(Me.SocketId, backlog)
    If listenResult = wsock32.ws32_SOCKET_ERROR Then
        Err.Raise UserDefinedErrors.SocketListenError.Code, "IdentityReader.IPv4StreamSocket.Listen", _
                  UserDefinedErrors.SocketListenError.Description & " Last error: " & CStr(Winsock.lastError)
        m_listening = False
    Else
        m_listening = True
    End If
    StartListening = Me.Listening

End Function

''' <summary>   Permits an incoming connection attempt on a socket. </summary>
''' <param name="serverSocket">   [in] Reference to the server socket. </param>
''' <param name="timeoutMs">      [Optional, 500] The maximum milliseconds time for select to wait. </param>
''' <returns>   the connected socket. </returns>
Public Function AcceptTcpClient(ByVal serverSocket As IPv4StreamSocket, _
        Optional ByVal timeoutMs As Long = 500) As IPv4StreamSocket

    ' add this socket to a socket set.
    WinsockApi.FD_SET_INIT m_fdSet
    WinsockApi.FD_SET_ADD m_socketId, m_fdSet
    
    ' check if we have available socket for readability
    
    Dim readableSocketCount As Integer
    readableSocketCount = WinsockApi.DetermineReadability(m_fdSet, timeoutMs)
    
    If readableSocketCount = 0 Then
        Set AcceptTcpClient = Nothing
        Exit Function
    End If
    
    Dim socket_id As Long
    Dim SocketAddress As wsock32.ws32_SocketAddress
    
    ' get a connected socket id
    
    socket_id = accept_(serverSocket.SocketId, SocketAddress, wsock32.ws32_SocketAddressLen)
    
    If socket_id = wsock32.ws32_INVALID_SOCKET Then
        Err.Raise UserDefinedErrors.SocketAcceptError.Code, "IdentityReader.IPv4StreamSocket.AcceptTcpClient", _
                  UserDefinedErrors.SocketAcceptError.Description & " Last error: " & CStr(Winsock.lastError)
        Exit Function
    End If
    
    Dim clientSocket As IPv4StreamSocket
    Set clientSocket = New IPv4StreamSocket
    clientSocket.Initialize socket_id, timeoutMs
    
    Set AcceptTcpClient = clientSocket
End Function

''' <summary>   Sends a terminated message to the server. </summary>
''' <param name="message">   The terminated message. </param>
''' <returns>   The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal message As String) As Integer

    If StringExtensions.IsNullOrEmpty(message) Then
        SendMessage = 0
    Else
        Dim messageLength As Integer
        messageLength = Len(message)
        Dim sentLength As Integer
        sentLength = send_(Me.SocketId, ByVal message, Len(message), 0)
        
        If sentLength < 0 Then
            ' if the sent length is negative, we have an API error.
            Err.Raise UserDefinedErrors.SocketSendError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                      UserDefinedErrors.SocketSendError.Description & " Last error: " & CStr(Winsock.lastError)
            SendMessage = SEND_MESSAGE_ERROR
        ElseIf sentLength <> messageLength Then
            Err.Raise UserDefinedErrors.SocketSendError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                      UserDefinedErrors.SocketSendError.Description & _
                      " A partial message was sent: " & CStr(sentLength) & " out of " & CStr(messageLength)
            SendMessage = SEND_PARTIAL_MESSAGE_ERROR
        Else
            SendMessage = sentLength
        End If
    End If

End Function

''' <summary>   Receives a message from the server until reaching
''' the specified termination. </summary>
''' <param name="buffer">        Where the message will go. </param>
''' <param name="maxLength">     The maximum number of bytes to read. </param>
''' <param name="termination">   [Optional, vbLf] The character that signifies the end of the stream. </param>
''' <returns>   If no error, the number of byte that were sent to the server.
''' Otherwise, RECEIVE_ERROR.
''' </returns>
Public Function ReceiveTerminatedMessage(ByRef buffer As String, ByVal maxLength As Integer, _
                                         Optional ByVal termination As String = vbLf) As Integer

    Dim c As String * 1
    Dim Length As Integer
    Dim count As Integer
    
    buffer = ""
    While Length < maxLength
    
        DoEvents
        
        c = ""
        Dim l As Long
        l = Len(c)
        count = recv_(Me.SocketId, c, l, 0)
        
        If count < 1 Then
            ReceiveTerminatedMessage = RECEIVE_ERROR
            buffer = Chr$(0)
            Exit Function
        End If
        
        If c = termination Then
           buffer = buffer + Chr$(0)
           ReceiveTerminatedMessage = Length
           Exit Function
        End If
        
        Length = Length + count
        buffer = buffer + c
        
    Wend
    
    ReceiveTerminatedMessage = RECEIVE_ERROR
    
End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="maxLength">   [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">   [Optional, True] True to return the string without the termination. </param>
''' <param name="termination">   [Optional, 10 ASCII] The character that signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function ReceiveRaw(Optional ByVal maxLength As Long = &H7FFFF, _
                           Optional ByVal trimEnd As Boolean = True, _
                           Optional ByVal termination As String = vbLf) As String

    Dim c As String * 1
    Dim count As Integer
    
    Dim builder As StringBuilder: Set builder = cc_isr_Core.Constructor.CreateStringBuilder()
    Dim currentLength As Long
    currentLength = 0
    Do While currentLength < maxLength
    
        DoEvents
        
        c = ""
        count = recv_(Me.SocketId, c, 1, 0)
        
        If count < 1 Then
            Err.Raise UserDefinedErrors.SocketReceiveError.Code, "IdentityReader.IPV4StreamSocket.receiveRaw", _
                      UserDefinedErrors.SocketReceiveError.Description & " Last error: " & CStr(Winsock.lastError)
            Exit Do
        ElseIf c = termination Then
            If Not trimEnd Then
                builder.Append c
            End If
            Exit Do
        Else
            currentLength = currentLength + count
            builder.Append c
        End If
    Loop
    ReceiveRaw = builder.ToString
    
End Function

''' <summary>   Receives a message from the server. </summary>
''' <param name="maxLength">   The maximum number of maxLength to read. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveBytes(ByVal maxLength As Long) As String
    
    Dim buffer As String
    buffer = StringExtensions.Repeat(Chr(0), maxLength)
    
    Dim actualLength As Long
    actualLength = recv_(Me.SocketId, buffer, maxLength, 0)
    
    If actualLength > 0 Then
        ReceiveBytes = StringExtensions.Substring(buffer, 0, actualLength)
    Else
        ReceiveBytes = ""
    End If
    
End Function

''' <summary>   Receives a message from the server. </summary>
''' <remarks>   Data is read until exhausted. </remarks>
''' <param name="bufferSize">   [Optional, 1024] The number of character to read on each trial. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveString(Optional ByVal bufferSize As Integer = 1024) As String
    
    Dim buffer As String
    Dim message As String
    message = ""
    
    Do
        buffer = Trim(Me.ReceiveBytes(bufferSize))

        If Len(buffer) > 0 Then
            message = message & buffer
        End If
    Loop While Len(buffer) > 0
    
    ReceiveString = Trim(message)
End Function

''' <summary>   Sets the socket receive timeout. </summary>
''' <param name="timeoutMs">   The timeout interval in milliseconds. </param>
Public Sub SetReceiveTimeout(ByVal timeoutMs As Long)

    Dim result As Long
    
    ' check if the requested receive timeout is different from the
    ' timeout that was already set.
    If m_receiveTimeout <> timeoutMs Then
    
        result = setsockopt_(m_socketId, wsock32.ws32_SOL_SOCKET, wsock32.ws32_SO_RCVTIMEO, timeoutMs, 4)
    
        If result = wsock32.ws32_SOCKET_ERROR Then
            Err.Raise UserDefinedErrors.SocketOptionSettingError.Code, "IdentityReader.IPV4StreamSocket.SetReceiveTimeout", _
                      UserDefinedErrors.SocketOptionSettingError.Description & " Last error: " & CStr(Winsock.lastError)
        Else
            m_receiveTimeout = timeoutMs
        End If
    End If
    
End Sub

''' <summary>   Gets the socket receive timeout in milliseconds. </summary>
Public Property Get receiveTimeout() As Long
    receiveTimeout = m_receiveTimeout
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Connection Changed

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal isConnected As Boolean)
    RaiseEvent ConnectionChanged( _
            cc_isr_Winsock.Constructor.CreateConnectionChangedEventArgs(isConnected))
End Sub




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

''' <summary> Establishes a connection to a specified socket. </summary>
''' <remarks> </remarks>
''' <param name="s">          [in] A descriptor identifying an unconnected socket. </param>
''' <param name="address">    [in] A pointer to the <see cref="wsock32.sockaddr_in"/> structure to which the
'''                           connection should be established. </param>
''' <param name="addressLen"> [in] The length, in bytes, of the sockaddr structure pointed to by the
'''                           <paramref name="address"/> parameter. </param>
''' <returns> If no error occurs, connect returns zero. Otherwise, it returns SOCKET_ERROR.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function connect_ Lib "wsock32.dll" Alias "connect" ( _
        ByVal s As Long, ByRef address As sockaddr_in, ByVal addressLen As Long) As Long

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
''' <returns> If no error occurs, listen returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function listen_ Lib "wsock32.dll" Alias "listen" ( _
        ByVal s As Long, ByVal backlog As Integer) As Long

''' <summary> Sends data on a connected socket. </summary>
''' <remarks> </remarks>
''' <param name="s">            [in] A descriptor identifying a connected socket. </param>
''' <param name="buffer">       [in] A pointer to a buffer containing the data to be transmitted. </param>
''' <param name="bufferLength"> [in] The length, in bytes, of the data in buffer pointed to by the buffer parameter. </param>
''' <param name="flags">        [in] A set of flags that specify the way in which the call is made. This parameter is
'''                             constructed by using the bitwise OR operator with any of the following values:
'''                             MSG_DONTROUTE: Specifies that the data should not be subject to routing. A Windows Sockets
'''                                            service provider can choose to ignore this flag.
'''                             MSG_OOB: Sends OOB data (stream-style socket such as SOCK_STREAM only).
'''                             </param>
''' <returns>
''' If no error occurs, send returns the total number of bytes sent, which can be less than the number requested to be sent
''' in the len parameter. Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be retrieved by calling
''' WSAGetLastError.
''' <returns>
Private Declare PtrSafe Function send_ Lib "wsock32.dll" Alias "send" ( _
        ByVal s As Long, buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long

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
    address.sin_addr.s_addr = WinsockApi.ToInetAddress(host)
    
    address.sin_family = wsock32.AF_INET
    
    ' set the port in network byte order
    address.sin_port = WinsockApi.ToInetByteOrder(port)
    
    Dim connectResult As Long
    connectResult = connect_(m_socketId, address, Len(address))
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
Public Function SendMessage(ByVal command As String) As Integer

    If StringExtensions.IsNullOrEmpty(command) Then
        SendMessage = 0
    Else
        Dim commandLength As Integer
        commandLength = Len(command)
        Dim sentLength As Integer
        sentLength = send_(Me.socketId, ByVal command, Len(command), 0)
        
        If sentLength < 0 Then
            ' if the sent length is negative, we have an API error.
            Err.Raise ErrorCodes.SocketSendError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                      ErrorCodes.SocketSendError.Description & " Last error: " & CStr(Winsock.lastError)
            SendMessage = COMMAND_ERROR
        ElseIf sentLength <> commandLength Then
        Else
            SendMessage = count
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

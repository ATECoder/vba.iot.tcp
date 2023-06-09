Attribute VB_Name = "wsock32"
''' <summary> The winsock implementation version </summary.
''' <remarks>
''' Version 1.1 (1*256 + 1) = 257
''' version 2.0 (2*256 + 0) = 512
''' </remarks>
Public Const WINSOCK_VERSION = 257

Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYS_STATUS_LEN = 128

Public Const WSADESCRIPTION_LEN_ARRAY = WSADESCRIPTION_LEN + 1
Public Const WSASYS_STATUS_LEN_ARRAY = WSASYS_STATUS_LEN + 1

''' <summary> A data structure that receives the information returned from
''' the WSAStartup() function. </summary>
Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN_ARRAY
    szSystemStatus As String * WSASYS_STATUS_LEN_ARRAY
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type

' Define address families
Public Const AF_UNSPEC = 0             ' unspecified
Public Const AF_UNIX = 1               ' local to host (pipes, portals)
Public Const AF_INET = 2               ' The Internet Protocol version 4 (IPv4) address family.

' Define socket types

''' <summar> A socket type that provides sequenced, reliable, two-way, connection-based byte streams with an
''' OOB data transmission mechanism. This socket type uses the Transmission Control Protocol (TCP) for the
''' Internet address family (AF_INET or AF_INET6). </summary>
Public Const SOCK_STREAM = 1

''' <summary>
''' A socket type that supports datagrams, which are connectionless, unreliable buffers of a fixed (typically
''' small) maximum length. This socket type uses the User Datagram Protocol (UDP) for the Internet address family
''' (AF_INET or AF_INET6).
''' </summary>
Public Const SOCK_DGRAM = 2

Public Const SOCK_RAW = 3              ' Raw data socket
Public Const SOCK_RDM = 4              ' Reliable Delivery socket
Public Const SOCK_SEQPACKET = 5        ' Sequenced Packet socket

Public Const INADDR_ANY = 0
Public Const INADDR_NONE = &HFFFFFFFF

''' <summary> Sets the Internet address type as a long integer (32-bit) </summary>
Public Type IN_ADDR
    s_addr As Long
End Type

''' <summary> Sets the socket IPv4 address expressed in network byte order. </summary>
Public Type sockaddr
    sa_family As Integer
    sa_data As String * 14
End Type

''' <summary> Sets the socket IPv4 address expressed in network byte order. </summary>
Public Type sockaddr_in
    sin_family As Integer  ' Address family of the socket, such as AF_INET.
    sin_port As Integer    ' sock address port number, e.g.,  htons(5150);
    sin_addr As IN_ADDR    ' the internet address as a long integer type.
    sin_zero As String * 8 '
End Type

Public Const FD_SETSIZE = 64

Public Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Long
End Type

Public Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

' Define socket return codes
Public Const INVALID_SOCKET = &HFFFF
Public Const SOCKET_ERROR = -1

Public Const SOL_SOCKET = 65535   ' socket options
Public Const SO_RCVTIMEO = &H1006 ' receive timeout option

Public Const MSG_OOB = &H1       ' Process out-of-band data.
Public Const MSG_PEEK = &H2      ' Peek at incoming messages.
Public Const MSG_DONTROUTE = &H4 ' Don't use local routing
Public Const MSG_WAITALL = &H8   ' do not complete until packet is completely filled

''' <summary> Creates a socket that is bound to a specific transport service provider. </summary>
''' <remarks> </remarks>
''' <param name="addressFamily"> [in] The address family specification. The values currently supported are
'''                              <see cref="AF_INET"/> or <see cref="AF_INET6"/>, which are the Internet
'''                              address family formats for IPv4 and IPv6. Other options for address family
'''                              (AF_NETBIOS for use with NetBIOS, for example) are supported if a Windows
'''                              Sockets service provider for the address family is installed. Note that the values
'''                              for the AF_ address family and PF_ protocol family constants are identical
'''                              (for example, AF_INET and PF_INET), so either constant can be used. </param>
''' <param name="socketType">    [in] The type specification for the new socket.
'''                              In Windows Sockets 1.1, the only possible socket types are SOCK_DGRAM and SOCK_STREAM. </param>
''' <param name="protocol">      [in] The protocol to be used. The possible options for the protocol parameter
'''                              are specific''' to the address family and socket type specified. </param>
''' <returns> If no error occurs, socket returns a descriptor referencing the new socket.
''' Otherwise, a value of INVALID_SOCKET is returned, and a specific error code can be retrieved by
''' calling WSAGetLastError. <returns>
Public Declare PtrSafe Function CreateSocket Lib "wsock32.dll" Alias "socket" (ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As Long

''' <summary> Establishes a connection to a specified socket. </summary>
''' <remarks> </remarks>
''' <param name="s">          [in] A descriptor identifying an unconnected socket. </param>
''' <param name="address">    [in] A pointer to the <see cref="wnsoc32.sockaddr_in"/> structure to which the
'''                           connection should be established. </param>
''' <param name="addressLen"> [in] The length, in bytes, of the sockaddr structure pointed to by the
'''                           <paramref name="address"/> parameter. </param>
''' <returns> If no error occurs, connect returns zero. Otherwise, it returns SOCKET_ERROR.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Public Declare PtrSafe Function connect Lib "wsock32.dll" (ByVal s As Long, ByRef address As sockaddr_in, ByVal addressLen As Long) As Long

''' <summary> Converts a u_short from host to TCP/IP network byte order (which is big-endian). </summary>
''' <remarks>
''' The htons function takes a 16-bit number in host byte order and returns a 16-bit number in network byte order
''' used in TCP/IP networks (the AF_INET or AF_INET6 address family).
'''
''' The htons function can be used to convert an IP port number in host byte order to the IP port number
''' in network byte order.
'''
''' The htons function does not require that the Winsock DLL has previously been loaded with a successful call
''' to the WSAStartup function.
''' </remarks>
''' <param name="hostshort"> [in] A 16-bit number in host byte order. </param>
''' <returns>  the value in TCP/IP network byte order. <returns>
Public Declare PtrSafe Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer

''' <summary> associates a local address with a socket. </summary>
''' <remarks> </remarks>
''' <param name="s">             [in] A descriptor identifying an unbound socket. </param>
''' <param name="address">       [in] A pointer to a sockaddr_in structure of the local address
'''                              to assign to the bound socket . </param>
''' <param name="addressLength"> [in] The length, in bytes, of the value pointed to by address. </param>
''' <returns> <returns>
Public Declare PtrSafe Function bind Lib "wsock32.dll" (ByVal s As Long, address As sockaddr_in, ByVal addressLength As Integer) As Long

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
Public Declare PtrSafe Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Integer) As Long

''' <summary> determines the status of one or more sockets, waiting if necessary, to perform synchronous I/O. </summary>
''' <remarks> winsock2  only? </remarks>
''' <param name="nfds">      [in] Ignored. The nfds parameter is included only for compatibility with Berkeley sockets.</param>
''' <param name="readfds">   [in, out] An optional pointer to a set of sockets to be checked for readability. </param>
''' <param name="writefds">  [in, out] An optional pointer to a set of sockets to be checked for writability. </param>
''' <param name="exceptfds"> [in, out] An optional pointer to a set of sockets to be checked for errors. </param>
''' <param name="timeout">   [in] const The maximum time for select to wait, provided in the form of a TIMEVAL structure.
'''                          Set the timeout parameter to null for blocking operations. </param>
''' <returns>
''' The total number of socket handles that are ready and contained in the fd_set structures, zero if the time limit expired,
''' or SOCKET_ERROR if an error occurred. If the return value is SOCKET_ERROR, WSAGetLastError can be used to retrieve
''' a specific error code.
''' <returns>
Public Declare PtrSafe Function SelectSockets Lib "wsock32.dll" Alias "select" (ByVal nfds As Integer, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Integer

''' <summary> Permits an incoming connection attempt on a socket. </summary>
''' <remarks> </remarks>
''' <param name="s">                   [in] A descriptor that identifies a socket that has been placed in a listening state with the listen function.
'''                                    The connection is actually made with the socket that is returned by accept.</param>
''' <param name="clientAddress">       [out] An optional pointer to a buffer that receives the address of the connecting entity,
'''                                    as known to the communications layer. The exact format of the addr parameter is determined by the address family that was established when the socket from the sockaddr structure was created.</param>
''' <param name="clientAddressLength"> [in, out] An optional pointer to an integer that contains the length of structure pointed to by
'''                                    the addr parameter. </param>
''' <returns> If no error occurs, accept returns a value of type SOCKET that is a descriptor for the new socket.
''' This returned value is a handle for the socket on which the actual connection is made. Otherwise, a value of
''' INVALID_SOCKET is returned, and a specific error code can be retrieved by calling WSAGetLastError.
''' The integer referred to by clientAddressLength initially contains the amount of space pointed to by clientAddress. On return it
''' will contain the actual length in bytes of the address returned. <returns>
Public Declare PtrSafe Function accept Lib "wsock32.dll" (ByVal s As Long, clientAddress As sockaddr, clientAddressLength As Integer) As Long

''' <summary> sets a socket option. </summary>
''' <remarks> </remarks>
''' <param name="s">       [in] A descriptor that identifies a socket. </param>
''' <param name="level">   [in] The level at which the option is defined (for example, SOL_SOCKET). </param>
''' <param name="optname"> [in] The socket option for which the value is to be set (for example, SO_BROADCAST).
'''                        The optname parameter must be a socket option defined within the specified level,
'''                        or behavior is undefined. </param>
''' <param name="optval">  [in] A pointer to the buffer in which the value for the requested option is specified. </param>
''' <param name="optlen">  [in] The size, in bytes, of the buffer pointed to by the optval parameter. </param>
''' <returns> If no error occurs, setsockopt returns zero. Otherwise, a value of SOCKET_ERROR is returned,
''' and a specific error code can be retrieved by calling WSAGetLastError. <returns>
Public Declare PtrSafe Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Integer) As Long

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
Public Declare PtrSafe Function send Lib "wsock32.dll" (ByVal s As Long, buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long

''' <summary> Receives data from a connected socket or a bound connectionless socket. </summary>
''' <remarks>
''' The flags parameter can be used to influence the behavior of the function invocation beyond the options specified
''' for the associated socket. The semantics of this function are determined by the socket options and the flags parameter.
''' The possible value of flags parameter is constructed by using the bitwise OR operator with any of the following values:
''' MSG_PEEK    Peeks at the incoming data. The data is copied into the buffer, but is not removed from the input queue.
''' MSG_OOB     Processes Out Of Band (OOB) data.
''' MSG_WAITALL The receive request will complete only when one of the following events occurs:
'''             The buffer supplied by the caller is completely full.
'''             The connection has been closed.
'''             The request has been canceled or an error occurred.
''' Note that if the underlying transport does not support MSG_WAITALL, or if the socket is in a non-blocking mode, then this call will fail with WSAEOPNOTSUPP. Also, if MSG_WAITALL is specified along with MSG_OOB, MSG_PEEK, or MSG_PARTIAL, then this call will fail with WSAEOPNOTSUPP. This flag is not supported on datagram sockets or message-oriented sockets.
''' </remarks>
''' <param name="s">            [in] A descriptor identifying a connected socket. </param>
''' <param name="buffer">       [out] A pointer to the buffer to receive the incomming data. </param>
''' <param name="bufferLength"> [in] The length, in bytes, of the data in buffer pointed to by the buffer parameter. </param>
''' <param name="flags">        [in] A set of flags that influences the behavior of this function. </param>
''' <returns>
''' If no error occurs, recv returns the number of bytes received and the buffer pointed to by the buffre parameter will
''' contain this data received. If the connection has been gracefully closed, the return value is zero.
''' Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be retrieved by calling WSAGetLastError.
''' <returns>
Public Declare PtrSafe Function recv Lib "wsock32.dll" (ByVal s As Long, ByVal buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long

''' <summary> The inet_addr function converts a string containing an IPv4 dotted-decimal address into a
''' proper address for the IN_ADDR structure. </summary>
''' <remarks> </remarks>
''' <param name="hostname"> [in] An IPv4 dotted-decimal address. </param>
''' <returns> If no error occurs, the inet_addr function returns an unsigned long value containing a suitable binary
''' representation of the Internet address given. If the string in the hostname parameter does not contain a legitimate
''' Internet address, for example if a portion of an "a.b.c.d" address exceeds 255, then inet_addr returns the value
''' INADDR_NONE. <returns>
Public Declare PtrSafe Function inet_addr Lib "wsock32.dll" (ByVal hostname As String) As Long

''' <summary> Closes an existing socket. </summary>
''' <remarks> </remarks>
''' <param name="s"> [in] A descriptor identifying the socket to close. </param>
''' <returns> If no error occurs, closesocket returns zero. Otherwise, a value of SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Public Declare PtrSafe Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long


Public Sub FD_ZERO_MACRO(ByRef s As fd_set)
    s.fd_count = 0
End Sub


Public Sub FD_SET_MACRO(ByVal fd As Long, ByRef s As fd_set)
    Dim i As Integer
    i = 0
    
    Do While i < s.fd_count
        If s.fd_array(i) = fd Then
            Exit Do
        End If
        
        i = i + 1
    Loop
    
    If i = s.fd_count Then
        If s.fd_count < FD_SETSIZE Then
            s.fd_array(i) = fd
            s.fd_count = s.fd_count + 1
        End If
    End If
End Sub



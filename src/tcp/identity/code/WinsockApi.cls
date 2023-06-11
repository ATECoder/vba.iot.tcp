VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinsockApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' <summary> Creates a socket that is bound to a specific transport service provider. </summary>
''' <remarks> </remarks>
''' <param name="addressFamily"> [in] The address family specification.
''' The values currently supported are <see cref="ws32_AF_INET"/> or <see cref="ws32_AF_INET6"/>, which are the Internet
''' address family formats for IPv4 and IPv6. Other options for address family (ws32_AF_NETBIOS for use with NetBIOS,
''' for example) are supported if a Windows Sockets service provider for the address family is installed.
''' Note that the values for the ws32_AF_ address family and PF_ protocol family constants are identical
''' (for example, ws32_AF_INET and PF_INET), so either constant can be used.
''' </param>
''' <param name="socketType">    [in] The type specification for the new socket.
''' In Windows Sockets 1.1, the only possible socket types are ws32_SOCK_DGRAM and ws32_SOCK_STREAM. </param>
''' <param name="protocol"> The protocol to be used. The possible options for the protocol parameter are specific
''' to the address family and socket type specified. </param>
''' <returns> If no error occurs, socket returns a descriptor referencing the new socket.
''' Otherwise, a value of ws32_INVALID_SOCKET is returned, and a specific error code can be retrieved by
''' calling WSAGetLastError. <returns>
Private Declare PtrSafe Function createSocket_ Lib "wsock32.dll" Alias "socket" ( _
        ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As Long

''' <summary> Converts a u_short from host to TCP/IP network byte order (which is big-endian). </summary>
''' <remarks>
''' The htons function takes a 16-bit number in host byte order and returns a 16-bit number in network byte order
''' used in TCP/IP networks (the ws32_AF_INET or ws32_AF_INET6 address family).
'''
''' The htons function can be used to convert an IP port number in host byte order to the IP port number
''' in network byte order.
'''
''' The htons function does not require that the Winsock DLL has previously been loaded with a successful call
''' to the WSAStartup function.
''' </remarks>
''' <param name="hostshort"> [in] A 16-bit number in host byte order. </param>
''' <returns>  the value in TCP/IP network byte order. <returns>
Private Declare PtrSafe Function toInetByteOrder_ Lib "wsock32.dll" Alias "htons" ( _
        ByVal hostshort As Long) As Integer
        
''' <summary> Converts an IPv4 dotted-decimal address into a binary address value for the ws32_IN_ADDR structure. </summary>
''' <remarks> </remarks>
''' <param name="hostname"> [in] An IPv4 dotted-decimal address. </param>
''' <returns> If no error occurs, returns an unsigned long value containing a suitable binary
''' representation of the Internet address given. If the string in the hostname parameter does not contain a legitimate
''' Internet address, for example if a portion of an "a.b.c.d" address exceeds 255, then returns the value
''' ws32_INADDR_NONE. <returns>
Private Declare PtrSafe Function toInetAddress_ Lib "wsock32.dll" Alias "inet_addr" ( _
        ByVal hostname As String) As Long
        
''' <summary> Determines the status of one or more sockets, waiting if necessary, to perform synchronous I/O. </summary>
''' <remarks> winsock2  only? </remarks>
''' <param name="nfds">      [in] Ignored. The nfds parameter is included only for compatibility with Berkeley sockets.</param>
''' <param name="readfds">   [in, out] An optional pointer to a set of sockets to be checked for readability. </param>
''' <param name="writefds">  [in, out] An optional pointer to a set of sockets to be checked for writability. </param>
''' <param name="exceptfds"> [in, out] An optional pointer to a set of sockets to be checked for errors. </param>
''' <param name="timeout">   [in] const The maximum time for select to wait, provided in the form of a TIMEVAL structure.
'''                          Set the timeout parameter to null for blocking operations. </param>
''' <returns>
''' The total number of socket handles that are ready and contained in the ws32_fd_set structures, zero if the time limit expired,
''' or ws32_SOCKET_ERROR if an error occurred. If the return value is ws32_SOCKET_ERROR, WSAGetLastError can be used to retrieve
''' a specific error code.
''' <returns>
Private Declare PtrSafe Function select_ Lib "wsock32.dll" Alias "select" ( _
        ByVal nfds As Integer, readfds As wsock32.ws32_fd_set, writefds As wsock32.ws32_fd_set, _
        exceptfds As wsock32.ws32_fd_set, Timeout As wsock32.ws32_Time_Value) As Integer

''' <summary> Creates a socket that is bound to a IPv4 transport service. </summary>
''' <returns> If no error occurs, socket returns a descriptor referencing the new socket.
''' Otherwise, a value of ws32_INVALID_SOCKET is returned, and a specific error code can be retrieved by
''' calling WSAGetLastError. <returns>
Public Function CreateIPv4StreamSocket() As Long
    CreateIPv4StreamSocket = createSocket_(wsock32.ws32_AF_INET, wsock32.ws32_SOCK_STREAM, 0)
End Function


''' <summary> Converts an unsigned short value from host to TCP/IP network byte order (which is big-endian). </summary>
''' <param name="value"> [in] A 16-bit number in host byte order. </param>
''' <returns>  The value in TCP/IP network byte order. <returns>
Public Function ToInetByteOrder(ByVal value As Long) As Integer
    ToInetByteOrder = toInetByteOrder_(value)
End Function

''' <summary> Converts an IPv4 dotted-decimal address into a binary address value for the ws32_IN_ADDR structure. </summary>
''' <returns> If no error occurs, returns an unsigned long value containing a suitable binary
''' representation of the Internet address given. If the string in the hostname parameter does not contain a legitimate
''' Internet address, for example if a portion of an "a.b.c.d" address exceeds 255, then returns the value
''' ws32_INADDR_NONE. <returns>
Public Function ToInetAddress(ByVal host As String) As Long
    ToInetAddress = toInetAddress_(host)
End Function

''' <summary> Initializes the socket set to the empty set. A set should always be cleared before using./ </summary>
''' <param name="socketSet">   [in, out] A pointer to the set of sockets. </param>
Public Sub FD_SET_INIT(ByRef socketSet As wsock32.ws32_fd_set)
    socketSet.fd_count = 0
End Sub

''' <summary> Removes socket s from set. </summary>
''' <param name="socket_Id">   [in] A descriptor that identifies a socket. </param>
''' <param name="socketSet">   [in, out] A pointer to the set of sockets. </param>
Public Sub FD_SET_REMOVE(ByVal s As Long, ByRef socketSet As wsock32.ws32_fd_set)
    Dim i As Integer
    i = 0
    Do While i < socketSet.fd_count
        If socketSet.fd_array(i) = socket_Id Then
            socketSet.fd_array(i) = 0
            socketSet.fd_count = socketSet.fd_count - 1
            Exit Do
        End If
        i = i + 1
    Loop
End Sub

''' <summary> Checks to see if s is a member of set and returns TRUE if so.. </summary>
''' <param name="socket_Id">   [in] A descriptor that identifies a socket. </param>
''' <param name="socketSet">   [in, out] A pointer to the set of sockets. </param>
Public Function FD_SET_CONTAINS(ByVal socket_Id As Long, ByRef socketSet As wsock32.ws32_fd_set) As Boolean
    Dim isSet As Boolean
    isSet = False
    Dim i As Integer
    i = 0
    
    Do While i < socketSet.fd_count
        If socketSet.fd_array(i) = socket_Id Then
            isSet = True
            Exit Do
        End If
        i = i + 1
    Loop
    FD_ISSET = isSet
   
End Function

''' <summary> Adds a socket to the set. </summary>
''' <param name="socket_Id">   [in] A descriptor that identifies a socket. </param>
''' <param name="socketSet">   [in, out] A pointer to the set of sockets. </param>
Public Sub FD_SET_ADD(ByVal socket_Id As Long, ByRef socketSet As wsock32.ws32_fd_set)
    Dim i As Integer
    i = 0
    
    Do While i < socketSet.fd_count
        If socketSet.fd_array(i) = socket_Id Then
            Exit Do
        End If
        i = i + 1
    Loop
    
    If i = socketSet.fd_count Then
        If socketSet.fd_count < wsock32.ws32_FD_SETSIZE Then
            socketSet.fd_array(i) = socket_Id
            socketSet.fd_count = socketSet.fd_count + 1
        End If
    End If
End Sub




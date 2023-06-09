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

''' <summary> Sets the Internet address type as a long integer (32-bit) </summary>
Public Type IN_ADDR
    s_addr As Long
End Type

''' <summary> Sets the socket address  </summary>
Public Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As IN_ADDR
    sin_zero As String * 8
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

''' <summary> A data type to store Internet addresses. </summary>
Public Type sockaddr
    sa_family As Integer
    sa_data As String * 14
End Type

' Define socket return codes
Public Const INVALID_SOCKET = &HFFFF
Public Const SOCKET_ERROR = -1

Public Const SOL_SOCKET = 65535
Public Const SO_RCVTIMEO = &H1006

''' <summary> Creates a socket that is bound to a specific transport service provider. </summary>
''' <remarks> </remarks>
''' <param name="addressFamily"> [in] The address family specification.
''' The values currently supported are <see cref="AF_INET"/> or <see cref="AF_INET6"/>, which are the Internet
''' address family formats for IPv4 and IPv6. Other options for address family (AF_NETBIOS for use with NetBIOS,
''' for example) are supported if a Windows Sockets service provider for the address family is installed.
''' Note that the values for the AF_ address family and PF_ protocol family constants are identical
''' (for example, AF_INET and PF_INET), so either constant can be used.
''' </para>
''' <param name="socketType">    [in] The type specification for the new socket.
''' In Windows Sockets 1.1, the only possible socket types are SOCK_DGRAM and SOCK_STREAM. </para>
''' <param name="protocol"> The protocol to be used. The possible options for the protocol parameter are specific
''' to the address family and socket type specified. </para>
''' <returns> If no error occurs, socket returns a descriptor referencing the new socket.
''' Otherwise, a value of INVALID_SOCKET is returned, and a specific error code can be retrieved by
''' calling WSAGetLastError. <returns>
Public Declare PtrSafe Function CreateSocket Lib "wsock32.dll" Alias "socket" (ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As Long

''' <summary> </summary>
''' <remarks> </remarks>
''' <param name=""> </para>
''' <returns> <returns>
Public Declare PtrSafe Function connect Lib "wsock32.dll" (ByVal s As Long, ByRef address As sockaddr_in, ByVal namelen As Long) As Long

Public Declare PtrSafe Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer

Public Declare PtrSafe Function bind Lib "wsock32.dll" (ByVal socket As Long, name As sockaddr_in, ByVal nameLength As Integer) As Long

Public Declare PtrSafe Function listen Lib "wsock32.dll" (ByVal socket As Long, ByVal backlog As Integer) As Long

Public Declare PtrSafe Function select_ Lib "wsock32.dll" Alias "select" (ByVal nfds As Integer, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Integer

Public Declare PtrSafe Function accept Lib "wsock32.dll" (ByVal socket As Long, clientAddress As sockaddr, clientAddressLength As Integer) As Long

Public Declare PtrSafe Function setsockopt Lib "wsock32.dll" (ByVal socket As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Integer) As Long

Public Declare PtrSafe Function send Lib "wsock32.dll" (ByVal socket As Long, buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long

Public Declare PtrSafe Function recv Lib "wsock32.dll" (ByVal socket As Long, ByVal buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long

Public Declare PtrSafe Function inet_addr Lib "wsock32.dll" (ByVal hostname As String) As Long

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



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
Public Const SOCK_STREAM = 1           ' Stream socket
Public Const SOCK_DGRAM = 2            ' Datagram socket
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

''' <summary> Initiates use of the Winsock DLL by a process. </summary>
''' <param name="was"> A pointer to the WSADATA data structure that is to receive
''' details of the Windows Sockets implementation. </para>
''' <returns> If successful, the WSAStartup function returns zero. Otherwise, it returns one of
''' the error codes listed below. The WSAStartup function directly returns the extended error code
''' in the return value for this function. A call to the WSAGetLastError function is not needed and should not be used.
''' </returns>
Public Declare PtrSafe Function WSAStartup Lib "wsock32.dll" (ByVal versionRequired As Long, wsa As WSADATA) As Long

Public Declare PtrSafe Function WSAGetLastError Lib "wsock32.dll" () As Long

''' <summary> terminates use of the Winsock dll. </summary>
''' <remarks> In a multithreaded environment, WSACleanup terminates Windows Sockets operations
'''   for all threads. </remarks>
''' <param name=""> </para>
''' <returns>
'''   The return value is zero if the operation was successful. Otherwise, the value
'''   SOCKET_ERROR is returned, and a specific error number can be retrieved by calling WSAGetLastError.
''' <returns>
Public Declare PtrSafe Function WSACleanup Lib "wsock32.dll" () As Long

''' <summary> </summary>
''' <remarks> </remarks>
''' <param name=""> </para>
''' <returns> <returns>
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



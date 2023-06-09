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
        MsgBox "ERROR: wsock32.CreateSocket returned " + Str$(m_socketId)
    Else
        Winsock.Register Me
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

Public Sub SetTimeout(ByVal timeout As Long)

    Dim result As Long
    result = wsock32.setsockopt(m_socketId, wsock32.SOL_SOCKET, wsock32.SO_RCVTIMEO, timeout, 4)

    If result = wnsock32.SOCKET_ERROR Then
        err.Raise 513, "IPv4StreamSocket", "ERROR: failed setting timeout. Last error = " + Str$(Winsock.LastError())
    End If
    
End Sub



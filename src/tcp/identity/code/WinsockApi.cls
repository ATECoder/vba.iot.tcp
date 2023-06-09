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
Private Declare PtrSafe Function CreateSocket Lib "wsock32.dll" Alias "socket" (ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As Long


''' <summary> Creates a socket that is bound to a IPv4 transport service. </summary>
Public Function CreateIPv4StreamSocket()
    CreateIPv4StreamSocket = CreateSocket(wsock32.AF_INET, wsock32.SOCK_STREAM, 0)
End Function

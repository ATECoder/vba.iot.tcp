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
''' </param>
''' <param name="socketType">    [in] The type specification for the new socket.
''' In Windows Sockets 1.1, the only possible socket types are SOCK_DGRAM and SOCK_STREAM. </param>
''' <param name="protocol"> The protocol to be used. The possible options for the protocol parameter are specific
''' to the address family and socket type specified. </param>
''' <returns> If no error occurs, socket returns a descriptor referencing the new socket.
''' Otherwise, a value of INVALID_SOCKET is returned, and a specific error code can be retrieved by
''' calling WSAGetLastError. <returns>
Private Declare PtrSafe Function createSocket_ Lib "wsock32.dll" Alias "socket" ( _
        ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As Long

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
Private Declare PtrSafe Function toInetByteOrder_ Lib "wsock32.dll" Alias "htons" ( _
        ByVal hostshort As Long) As Integer
        
''' <summary> Converts an IPv4 dotted-decimal address into a binary address value for the IN_ADDR structure. </summary>
''' <remarks> </remarks>
''' <param name="hostname"> [in] An IPv4 dotted-decimal address. </param>
''' <returns> If no error occurs, returns an unsigned long value containing a suitable binary
''' representation of the Internet address given. If the string in the hostname parameter does not contain a legitimate
''' Internet address, for example if a portion of an "a.b.c.d" address exceeds 255, then returns the value
''' INADDR_NONE. <returns>
Private Declare PtrSafe Function toInetAddress_ Lib "wsock32.dll" Alias "inet_addr" ( _
        ByVal hostname As String) As Long

''' <summary> Creates a socket that is bound to a IPv4 transport service. </summary>
''' <returns> If no error occurs, socket returns a descriptor referencing the new socket.
''' Otherwise, a value of INVALID_SOCKET is returned, and a specific error code can be retrieved by
''' calling WSAGetLastError. <returns>
Public Function CreateIPv4StreamSocket() As Long
    CreateIPv4StreamSocket = createSocket_(wsock32.AF_INET, wsock32.SOCK_STREAM, 0)
End Function

''' <summary> Converts an unsigned short value from host to TCP/IP network byte order (which is big-endian). </summary>
''' <param name="value"> [in] A 16-bit number in host byte order. </param>
''' <returns>  The value in TCP/IP network byte order. <returns>
Public Function ToInetByteOrder(ByVal value As Long) As Integer
    ToInetByteOrder = toInetByteOrder_(value)
End Function

''' <summary> Converts an IPv4 dotted-decimal address into a binary address value for the IN_ADDR structure. </summary>
''' <returns> If no error occurs, returns an unsigned long value containing a suitable binary
''' representation of the Internet address given. If the string in the hostname parameter does not contain a legitimate
''' Internet address, for example if a portion of an "a.b.c.d" address exceeds 255, then returns the value
''' INADDR_NONE. <returns>
Public Function ToInetAddress(ByVal host As String) As Long
    ToInetAddress = toInetAddress_(host)
End Function


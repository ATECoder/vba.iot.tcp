VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_tcpServerSocket As IPv4StreamSocket

''' <summary> Gets the Tcp Server socket. </summary>
''' <value> A IPv4StreamSocket. </value>
Public Property Get TcpServerSocket() As IPv4StreamSocket
    Set TcpServerSocket = m_tcpServerSocket
End Property

''' <summary> Associates any local address with a socket and start listening. </summary>
''' <param name="port">    The port that the server is listening on. </param>
''' <param name="backlog"> [Optional, 10] The maximum length of the queue of pending connections. </para>
''' <returns> True if the socket is bound. </returns>
Public Function BindTo(ByVal port As Long, Optional ByVal backlog As Integer = 10) As Boolean

    On Error GoTo Finally
    
    Set m_tcpServerSocket = New IPv4StreamSocket
    
    ' bind the socket to the specified port for listening on any address
    
    m_tcpServerSocket.BindTo wsock32.ws32_INADDR_ANY, port
    
    ' get the socket to start listening
    
    BindTo = m_tcpServerSocket.StartListening(backlog)
    
Finally:

    If Me.Connected Then
        Me.TcpServerSocket.CloseConnection
    End If
    m_tcpServerSocket = Nothing
    
    Err.Raise Err

End Function

''' <summary> Returns the connection state of the client. </summary>
''' <returns> True if the socket is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    If m_tcpServerSocket Is Nothing Then
        Connected = False
    Else
        Connected = m_tcpServerSocket.Connected
    End If
End Property

''' <summary> Permits an incoming connection attempt on a socket. </summary>
''' <param name="timeout"> [Optional, 500] The maximum milliseconds time for select to wait. </param>
''' <returns> the connected tcp client. </returns>
Public Function AcceptTcpClient(Optional ByVal timeoutMs As Long = 500) As TcpClient

    Dim connectedTcpClientSocket As IPv4StreamSocket
    connectedTcpClientSocket = Me.TcpServerSocket.AcceptTcpClient(Me, timeoutMs)
    
    Dim client As TcpClient
    Set client = New TcpClient
    client.Initialize connectedTcpClientSocket
    
    Set AcceptTcpClient = client
End Function

''' <summary> Closes and releases the socket. </summary>
Public Sub Dispose()
    If Me.Connected Then
        m_tcpServerSocket.CloseConnection
    End If
    m_tcpServerSocket = Nothing
End Sub

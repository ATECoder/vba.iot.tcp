VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_tcpClientSocket As IPv4StreamSocket

Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Gets the Tcp Client socket. </summary>
''' <value>   A IPv4StreamSocket. </value>
Public Property Get TcpClientSocket() As IPv4StreamSocket
    Set TcpClientSocket = m_tcpClientSocket
End Property

''' <summary>   Opens a socket connection and sets and returns the socket id. </summary>
''' <param name="host">   An IPv4 dotted-decimal host address. </param>
''' <param name="port">   The port that the server is listening on. </param>
''' <returns>   True if the socket is connected. </returns>
Public Function ConnectTo(ByVal host As String, ByVal port As Long) As Boolean
    
    If host = "localhost" Then
        host = "127.0.0.1"
    End If

    Set m_tcpClientSocket = New IPv4StreamSocket

    m_tcpClientSocket.OpenConnection host, port
    
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the socket is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    If m_tcpClientSocket Is Nothing Then
        Connected = False
    Else
        Connected = Me.TcpClientSocket.Connected
    End If
End Property

''' <summary>   Closes the connection. </summary>
''' <returns>   True if disconnection succeeded; otherwise, false. </returns>
Public Function Disconnect() As Boolean

    Me.TcpClientSocket.CloseConnection
    
    Set m_tcpClientSocket = Nothing
        
    Disconnect = True

End Function

''' <summary>   Initializes this TCP Client using the provided connected socket. </summary>
''' <param name="connectedSocket ">   A connected INET stream socket. </param>
Public Sub Initialize(ByVal connectedSocket As IPv4StreamSocket)
    Set m_tcpClientSocket = connectedSocket
End Sub

''' <summary>   Sends a terminated message to the server. </summary>
''' <param name="command">   The message. </param>
''' <returns>   The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal message As String) As Integer
    SendMessage = Me.TcpClientSocket.SendMessage(message)
End Function

''' <summary>   Receives a message from the server. </summary>
''' <param name="maxLength">   The maximum number of characters to read. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveBytes(ByVal maxLength As Long) As String
    
    ReceiveBytes = Me.TcpClientSocket.ReceiveBytes(maxLength)

End Function

''' <summary>   Receives a message from the server. </summary>
''' <remarks>   Data is read until exhausted. </remarks>
''' <param name="bufferSize">   [Optional, 1024] The number of character to read on each trial. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveString(Optional ByVal bufferSize As Integer = 1024) As String
    
    ReceiveString = Me.TcpClientSocket.ReceiveString(bufferSize)

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <param name="termination">   [Optional, 10 ASCII] The character the signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function ReceiveRaw(Optional ByVal maxLength As Long = &H7FFF, _
                           Optional ByVal trimEnd As Boolean = True, _
                           Optional ByVal termination As String = vbLf) As String
    
    ReceiveRaw = Me.TcpClientSocket.ReceiveRaw(maxLength, trimEnd, termination)
  
End Function

''' <summary>   Closes and releases the socket. </summary>
Public Sub Dispose()
    If Me.Connected Then
        m_tcpClientSocket.CloseConnection
    End If
    Set m_tcpClientSocket = Nothing
End Sub


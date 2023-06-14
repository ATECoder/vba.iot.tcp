VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_readTermination As String
Private m_writeTermination As String
Private m_timeoutMs As Long
Private m_tcpClient As TcpClient

''' <summary>   Initializes an IEEE488 Session. </summary>
Private Sub Class_Initialize()
    m_readTermination = vbLf
    m_writeTermination = vbLf
    m_timeoutMs = 500
End Sub

''' <summary>   Termiantes an IEEE488 Session. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Closes and releases the TCP Client. </summary>
Public Sub Dispose()
    If Me.Connected Then
        Me.CloseConnection
    End If
    Set m_tcpClient = Nothing
End Sub

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the TCP Client socket is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    If m_tcpClient Is Nothing Then
        Connected = False
    Else
        Connected = m_tcpClient.Connected
    End If
End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <returns>   True if a connection can be made. </returns>
Public Property Get CanConnect() As Boolean
    CanConnect = Not Me.Connected
End Property

''' <summary>   Returns a descriptor identifying the socket. </summary>
''' <returns>   The socket identifying descriptor. </returns>
Public Property Get SocketId() As Long
    If Me.Connected Then
        SocketId = m_tcpClient.TcpClientSocket.SocketId
    Else
        SocketId = -1
    End If
End Property

''' <summary>   Gets the read termination. </summary>
''' <value>   A string. </value>
Public Property Get ReadTermination() As String
    ReadTermination = m_readTermination
End Property

''' <summary>   Sets the read termination. </summary>
''' <param name="value">   The read termiantion. </param>
Public Property Let ReadTermination(value As String)
    m_readTermination = value
End Property

''' <summary>   Gets the Write termination. </summary>
''' <value>   A string. </value>
Public Property Get WriteTermination() As String
    WriteTermination = m_writeTermination
End Property

''' <summary>   Sets the Write termination. </summary>
''' <param name="value">   The Write termiantion. </param>
Public Property Let WriteTermination(value As String)
    m_writeTermination = value
End Property

''' <summary>   Gets the Timeout in Milliseconds. </summary>
''' <value>   A string. </value>
Public Property Get timeoutMs() As String
    timeoutMs = m_timeoutMs
End Property

''' <summary>   Sets the socket timeout. </summary>
''' <param name="timeout">   The timeout interval in milliseconds. </param>
Public Sub SetTimeout(ByVal interval As Long)
    If Me.Connected Then
        If interval <> m_tcpClient.TcpClientSocket.Timeout Then
            m_tcpClient.TcpClientSocket.SetTimeout interval
        End If
    End If
    m_timeoutMs = interval
    
End Sub

''' <summary>   Opens a TCP Client connection. </summary>
''' <param name="host">   An IPv4 dotted-decimal host address. </param>
''' <param name="port">   The port that the server is listening on. </param>
''' <returns>   True if the TCP Client is connected. </returns>
Public Function OpenConnection(ByVal host As String, ByVal port As Long, Optional timeoutMs As Long) As Boolean
   
    Set m_tcpClient = New TcpClient
   
    m_tcpClient.ConnectTo host, port
    
    Me.SetTimeout timeoutMs
    
    OpenConnection = Me.Connected
   
End Function

''' <summary>   Closes the TCP Client connection. </summary>
''' <returns>   True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    CloseConnection = m_tcpClient.Disconnect
    
    Set m_tcpClient = Nothing
    
End Function

''' <summary>   Sends a message. </summary>
''' <param name="message">   The message. </param>
''' <returns>   The number of sent characters. </returns>
Public Function WriteLine(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    If appendTermination Then
        message = message & Me.WriteTermination
    End If
    WriteLine = m_tcpClient.SendMessage(message)
End Function

''' <summary>   Receives a message from the server untill reaching the specified termiantion
''' or reading the specified number of characters. </summary>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <returns>   The received string. </returns>
Public Function Read(Optional ByVal maxLength As Long = &H7FFF, _
                     Optional ByVal trimEnd As Boolean = True) As String
   
    Read = m_tcpClient.ReceiveRaw(maxLength, trimEnd, Me.ReadTermination)
  
End Function
    
''' <summary>   Sends a message and receives a reply. </summary>
''' <param name="message">   The message. </param>
''' <returns>   The received string. </returns>
Public Function QueryLine(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    If Me.WriteLine(message, appendTermination) > 0 Then
        QueryLine = Me.Read()
    Else
        QueryLine = ""
    End If
End Function


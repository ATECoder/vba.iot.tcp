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
Private Const PROLOGIX_PORT_NO As Integer = 1234
Private m_readTermination As String
Private m_writeTermination As String
Private m_receiveTimeout As Long
Private m_tcpClient As TcpClient
Private m_usingPrologix As Boolean
Private m_readAfterWriteEnabled As Boolean

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByRef eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   Handles the Class Initialize event; constructs the <see cref="TcpClient"/>. </summary>
Private Sub Class_Initialize()
    Set m_tcpClient = cc_isr_Winsock.Constructor.CreateTcpClient
    m_readTermination = vbLf
    m_writeTermination = vbLf
    m_receiveTimeout = 500
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="TcpClient"/>. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Closes and releases the TCP Client. </summary>
Public Sub Dispose()
    If Me.Connected Then
        Me.CloseConnection
    End If
    If Not m_tcpClient Is Nothing Then
        m_tcpClient.Dispose
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
    CanConnect = Not Me.Connected And _
                 Not m_tcpClient.TcpClientSocket Is Nothing
End Property

''' <summary>   Returns a descriptor identifying the socket. </summary>
''' <value>   The socket identifying descriptor if connected; otherwise, -1. </value>
Public Property Get SocketId() As Long
    If m_tcpClient Is Nothing Then
        SocketId = cc_isr_Winsock.wsock32.ws32_INVALID_SOCKET
    ElseIf m_tcpClient.TcpClientSocket Is Nothing Then
        SocketId = cc_isr_Winsock.wsock32.ws32_INVALID_SOCKET
    Else
        SocketId = m_tcpClient.TcpClientSocket.SocketId
    End If
End Property

''' <summary>   Returns the socket address as IP plus port. </summary>
''' <value>   The socket address. </value>
Public Property Get SocketAddress() As String
    If Me.Connected Then
        SocketAddress = m_tcpClient.TcpClientSocket.SocketAddress
    Else
        SocketAddress = ""
    End If
End Property

''' <summary>   Gets the read termination. </summary>
''' <value>   A string. </value>
Public Property Get ReadTermination() As String
    ReadTermination = m_readTermination
End Property

''' <summary>   Sets the read termination. </summary>
''' <param name="value">   The read termination. </param>
Public Property Let ReadTermination(value As String)
    m_readTermination = value
End Property

''' <summary>   Gets the Write termination. </summary>
''' <value>   A string. </value>
Public Property Get WriteTermination() As String
    WriteTermination = m_writeTermination
End Property

''' <summary>   Sets the Write termination. </summary>
''' <param name="value">   The Write termination. </param>
Public Property Let WriteTermination(value As String)
    m_writeTermination = value
End Property

''' <summary>   Gets the receive timeout in Milliseconds. </summary>
''' <value>   A Long. </value>
Public Property Get ReceiveTimeout() As Long
    ReceiveTimeout = m_receiveTimeout
End Property

''' <summary>   Sets the socket receive timeout. </summary>
''' <param name="timeoutMs">   The timeout interval in milliseconds. </param>
Public Sub SetReceiveTimeout(ByVal timeoutMs As Long)
    If Me.Connected Then
        If timeoutMs <> m_tcpClient.TcpClientSocket.ReceiveTimeout Then
            m_tcpClient.TcpClientSocket.SetReceiveTimeout timeoutMs
        End If
    End If
    m_receiveTimeout = timeoutMs
    
End Sub

''' <summary>   Opens a TCP Client connection. </summary>
''' <param name="host">        An IPv4 dotted-decimal host address. </param>
''' <param name="port">        The port that the server is listening on. </param>
''' <param name="timeoutMs">   [Optional, 500] The receive timeout in milliseconds. </param>
''' <returns>   True if the TCP Client is connected. </returns>
Public Function OpenConnection(ByVal host As String, ByVal port As Long, Optional ByVal timeoutMs As Long = 500) As Boolean
   
    Me.UsingPrologix = port = PROLOGIX_PORT_NO
    
    m_tcpClient.ConnectTo host, port
    
    Me.SetReceiveTimeout timeoutMs
    
    If Me.UsingPrologix Then
        ' set prologix to auto read after write.
        ' this forces the opposite value
        m_readAfterWriteEnabled = False
        Me.SetReadAfterWrite Not m_readAfterWriteEnabled
    End If
    
    OnConnectionChanged Me.Connected

    OpenConnection = Me.Connected
   
End Function

''' <summary>   Closes the TCP Client connection. </summary>
''' <returns>   True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    CloseConnection = m_tcpClient.Disconnect
    
    OnConnectionChanged Me.Connected
    
End Function

''' <summary>   Returns true if the TCP Connection uses the Prologix
''' GPIB-Lan interface device. </summary>
''' <value>   True if using the Prologix GPIB-Lan interface device. </value>
Public Property Get UsingPrologix() As Boolean
    UsingPrologix = m_usingPrologix
End Property

''' <summary>   Set true if the TCP Connection uses the Prologix
''' GPIB-Lan interface device. </summary>
Public Property Let UsingPrologix(ByVal value As Boolean)
    m_usingPrologix = value
End Property

''' <summary>   Sends a message. </summary>
''' <remarks>   If using the Prologix device at port 1234, this
''' method first sets the Prologix to auto off (++auto 0) to prevent it from
''' setting the device to talk prematurely which might cause the device
''' (e.g., the Keithley 2700 scanning multimeter) to issue error -420 Query Unterminated. </remarks>
''' <param name="message">   The message. </param>
''' <returns>   The number of sent characters. </returns>
Public Function WriteLine(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    
    If appendTermination Then
        message = message & Me.WriteTermination
    End If
    WriteLine = m_tcpClient.SendMessage(message)

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <remarks>   If using the Prologix device at port 1234, this
''' method first sets the Prologix to auto on (++auto 1) assuming auto was turned off on
''' the previous read on <see cref="WriteLine"/>. </remarks>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <returns>   The received string. </returns>
Public Function Read(Optional ByVal maxLength As Long = &H7FFF, _
                     Optional ByVal trimEnd As Boolean = True) As String
    
    Read = m_tcpClient.ReceiveRaw(maxLength, trimEnd, Me.ReadTermination)

End Function
    
Public Function PrologixRead(Optional ByVal maxLength As Long = &H7FFF, _
                     Optional ByVal trimEnd As Boolean = True) As String
    m_tcpClient.SendMessage "++read eoi" & Me.WriteTermination
    PrologixRead = m_tcpClient.ReceiveRaw(32767)
End Function


''' <summary>   Sends a message and receives a reply. </summary>
''' <param name="message">   The message. </param>
''' <returns>   The received string. </returns>
Public Function QueryLine(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    If Me.WriteLine(message, appendTermination) > 0 Then
        DoEvents
        QueryLine = Me.Read()
    Else
        QueryLine = ""
    End If
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' PROLOGIX

''' <summary>   Gets the read-after-write enabled status. </summary>
''' <value>   True if wrte-after-write is enabled.  </value>
Public Property Get ReadAfterWriteEnabled() As Boolean
    ReadAfterWriteEnabled = m_readAfterWriteEnabled
End Property

''' <summary>   Configure the Read-After-Write mode of the Prologix controller. </summary>
''' <remarks>
''' Prologix GPIB-ETHERNET controller can be configured to automatically address
''' instruments to talk after sending them a command in order to read their response. The
''' feature called, Read-After-Write, saves the user from having to issue read commands
''' repeatedly. This command enabled or disabled the Read-After-Write feature.
'''
''' In addition, auto command also addresses the instrument at the currently specified
''' address to TALK or LISTEN. ++auto 0 addresses the instrument to LISTEN and
''' ++auto 1 addresses the instrument to TALK.
''' If the command is issued without any arguments it returns the current state of the
''' readafter-write feature.
'''
''' NOTE:
''' Some instruments generate “Query Unterminated” or “-420” error if they are addressed
''' to talk after sending a command that does not generate a response (often called non
''' query commands). In effect the instrument is saying, I have been asked to talk but I have
''' nothing to say. The error is often benign and may be ignored. Otherwise, use the
''' ++read command to read the instrument response. For example:
''' <code>
''' ++auto 0 — Turn off read-after-write and address instrument to listen
''' SET VOLT 1.0 — Non-query command
''' *idn? — Query command
''' ++read eoi — Read until EOI asserted by instrument
''' "HP54201A" — Response from instrument
''' </code>
''' <param name="enable">   True to enable read-after-write. </param>
''' </remarks>
Public Sub SetReadAfterWrite(ByVal enable As Boolean)
    
    If Me.UsingPrologix And m_readAfterWriteOn <> Me.ReadAfterWriteEnabled Then
        Dim output As String: output = IIf(readAfterWriteOn, "1", "0")
        m_tcpClient.SendMessage "++auto " & output & Me.WriteTermination
        m_readAfterWriteEnabled = readAfterWriteOn
    End If

End Sub

''' <summary>   Query the Read-After-Write mode of the Prologix controller. </summary>
Public Function GetReadAfterWrite() As Boolean
    m_readAfterWriteEnabled = "1" = Me.QueryLine("++auto")
End Function


' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Connection Changed

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal isConnected As Boolean)
    RaiseEvent ConnectionChanged( _
            cc_isr_Winsock.Constructor.CreateConnectionChangedEventArgs(isConnected))
End Sub






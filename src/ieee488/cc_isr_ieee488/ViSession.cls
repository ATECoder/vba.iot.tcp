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

Implements cc_isr_Winsock.IConnectable

Private Const GPIB_LAN_PORT_NO As Integer = 1234

Private Const READ_AFTER_WRITE_DELAY_DEFAULT As Integer = 5

Private Const SESSION_READ_TIMEOUT_DEFAULT As Integer = 3000

Private Const SOCKET_RECEIVE_TIMEOUT_DEFAULT As Integer = 500

Private m_gpibLan As GpibLanController

Private m_readAfterWriteDelay As Integer

Private m_readTermination As String

Private m_socketReceiveTimeout As Long

Private m_sessionReadTimeout As Long

Private m_readAfterWriteEnabled As Boolean

Private m_stopper As StopWatch

Private WithEvents m_tcpClient As TcpClient

Private m_usingGpibLan As Boolean

Private m_writeTermination As String

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event args. </param>
Public Event ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Handles the Class Initialize event; constructs the <see cref="TcpClient"/>. </summary>
Private Sub Class_Initialize()
    Set m_stopper = cc_isr_Core.Constructor.CreateStopWatch
    Set m_tcpClient = cc_isr_Winsock.Constructor.CreateTcpClient
    m_readTermination = vbLf
    m_writeTermination = vbLf
    m_socketReceiveTimeout = SOCKET_RECEIVE_TIMEOUT_DEFAULT
    m_sessionReadTimeout = SESSION_READ_TIMEOUT_DEFAULT
    Me.ReadAfterWriteDelay = READ_AFTER_WRITE_DELAY_DEFAULT
    
    Set m_gpibLan = New GpibLanController
    m_gpibLan.Initialize m_tcpClient, New ErrTracer, m_readTermination, m_writeTermination, READ_AFTER_WRITE_DELAY_DEFAULT
    
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="TcpClient"/>. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Initializes the <see cref="IErrTracer"/> for the <see cref="ViSession"/>. </summary>
''' <param name="a_tracer">   <see cref="IErrTracer"/> implementation. </param>
Public Sub Initialize(ByVal a_tracer As IErrTracer)
    
    m_gpibLan.Initialize m_tcpClient, a_tracer, m_readTermination, m_writeTermination, READ_AFTER_WRITE_DELAY_DEFAULT
   
End Sub

''' <summary>   Gets a reference to the GPIB Lan. </summary>
Public Property Get GpibLan() As GpibLanController
    Set GpibLan = m_gpibLan
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
        SocketAddress = vbNullString
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

''' <summary>   Gets the Read After Write Delay. </summary>
''' <value>   An Integer. </value>
Public Property Get ReadAfterWriteDelay() As Integer
    ReadAfterWriteDelay = m_readAfterWriteDelay
End Property

''' <summary>   Sets the Read After Write Delay. </summary>
Public Property Let ReadAfterWriteDelay(ByVal a_value As Integer)
    m_readAfterWriteDelay = a_value
End Property

''' <summary>   Gets the Session Read Timeout. </summary>
''' <value>   An Integer. </value>
Public Property Get SessionReadTimeout() As Integer
    SessionReadTimeout = m_sessionReadTimeout
End Property

''' <summary>   Sets the Session Read Timeout. </summary>
Public Property Let SessionReadTimeout(ByVal a_value As Integer)
    m_sessionReadTimeout = a_value
End Property

''' <summary>   Gets the socket receive timeout in Milliseconds. </summary>
''' <value>   A Long. </value>
Public Property Get SocketReceiveTimeout() As Long
    SocketReceiveTimeout = m_socketReceiveTimeout
End Property

''' <summary>   Sets the socket receive timeout. </summary>
''' <param name="a_timeoutMs">   The timeout interval in milliseconds. </param>
Public Sub SetSocketReceiveTimeout(ByVal a_timeoutMs As Long)
    If Me.Connected Then
        If a_timeoutMs <> m_tcpClient.TcpClientSocket.ReceiveTimeout Then
            m_tcpClient.TcpClientSocket.SetReceiveTimeout a_timeoutMs
        End If
    End If
    m_socketReceiveTimeout = a_timeoutMs
    
End Sub

''' <summary>   Returns true if the TCP Connection uses the Prologix
''' GPIB-Lan interface device. </summary>
''' <value>   [Boolean] True if using the Prologix GPIB-Lan interface device. </value>
Public Property Get UsingGpibLan() As Boolean
    UsingGpibLan = m_usingGpibLan
End Property

''' <summary>   Sends a message. </summary>
''' <remarks>   If using the Prologix device at port 1234, this
''' method first sets the Prologix to auto off (++auto 0) to prevent it from
''' setting the device to talk prematurely which might cause the device
''' (e.g., the Keithley 2700 scanning multimeter) to issue error -420 Query Unterminated. </remarks>
''' <param name="a_message">             [String] The message to send to the instrument. </param>
''' <param name="a_appendTermination">   [Optional, Boolean, True] True to append termination to
'''                                      the message. </param>
''' <returns>   [Long] The number of sent characters. </returns>
Public Function WriteLine(ByVal a_message As String, _
        Optional ByVal a_appendTermination As Boolean = True) As Long
    
    If Me.UsingGpibLan Then
        WriteLine = Me.GpibLan.SendToDevice(a_message, a_appendTermination)
    Else
        If a_appendTermination Then a_message = a_message & Me.WriteTermination
        WriteLine = m_tcpClient.SendMessage(a_message)
        m_stopper.Wait Me.ReadAfterWriteDelay
    End If

End Function

''' <summary>   Receives a message from the server until reaching the specified termination,
''' reading the specified number of characters, or <see cref="SeassionReadTimeout">timeout</see>. </summary>
''' <param name="a_maxLength">  [Optional, Long, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">    [Optional, Boolean, True] True to return the string without
'''                             the termination. </param>
''' <returns>   [String] The received message. </returns>
Public Function Read(Optional ByVal a_maxLength As Long = &H7FFF, _
        Optional ByVal a_trimEnd As Boolean = True) As String
    
    Dim p_reading As String
    If Me.SessionReadTimeout > 0 Then
        p_reading = AwaitReading(Me.SessionReadTimeout, a_maxLength, a_trimEnd)
    Else
        On Error Resume Next
        If Me.UsingGpibLan Then
            p_reading = Me.GpibLan.ReceiveFromDevice(a_maxLength, a_trimEnd)
        Else
            p_reading = m_tcpClient.ReceiveRaw(a_maxLength, a_trimEnd, Me.ReadTermination)
        End If
        On Error GoTo 0
    End If
        
    ' report an error on failure to read.
    If 0 = Len(p_reading) Then
    
        ' raise the read error adding some information
        Err.Raise UserDefinedErrors.TimeoutError.Code, "cc_isr_Ieee488.ViSession.Read", _
                  UserDefinedErrors.TimeoutError.Description & _
                  ". Data not received reading the instrument at " & Me.SocketAddress & _
                  " with a timeout of " & CStr(Me.SessionReadTimeout) & "ms."
    
    End If
    
    Read = p_reading

End Function

''' <summary>   Receives a message from the server until reaching the specified termination, reading the
''' specified number of characters, or timeout. </summary>
''' <remarks>   If using a GPIB-Lan device, such as the Prologix GPIB-Lan interface, this
''' method first uses the device <c>++read</c> to read the instrument. </remarks>
''' <param name="a_timeout">     [Integer] time to wait for reply in milliseconds. </param>
''' <param name="a_maxLength">   [Optional, Long, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">     [Optional, Boolean, True] True to return the string without the termination. </param>
''' <returns>   [String] The reading. </returns>
Public Function AwaitReading(ByVal a_timeout As Integer, Optional ByVal a_maxLength As Long = &H7FFF, _
        Optional ByVal a_trimEnd As Boolean = True) As String

    On Error Resume Next
    
    Dim p_reading As String
    Dim p_stopper As cc_isr_Core.StopWatch
    Set p_stopper = cc_isr_Core.Constructor.CreateStopWatch()
    p_stopper.Restart
    
    ' wait for data or timeout
    Do Until (Len(p_reading) > 0) Or _
             (p_stopper.ElapsedMilliseconds > a_timeout)
    
        DoEvents
        
        ' take a reading
        
        If Me.UsingGpibLan Then
            p_reading = Me.GpibLan.ReceiveFromDevice(a_maxLength, a_trimEnd)
        Else
            p_reading = m_tcpClient.ReceiveRaw(a_maxLength, a_trimEnd, Me.ReadTermination)
        End If
        
    Loop
        
    AwaitReading = p_reading

    On Error GoTo 0
    
End Function
   
''' <summary>   Sends a message and receives a reply. </summary>
''' <param name="a_message">             [String] The message to send to the instrument. </param>
''' <param name="a_appendTermination">   [Optional, Boolean, True] True to append termination to
'''                                      the message. </param>
''' <param name="a_maxLength">           [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">             [Optional, Boolean, True] True to return the string without the termination. </param>
''' <returns>   [String] The received string. </returns>
Public Function QueryLine(ByVal a_message As String, _
        Optional ByVal a_appendTermination As Boolean = True, _
        Optional ByVal a_maxLength As Long = &H7FFF, _
        Optional ByVal a_trimEnd As Boolean = True) As String
    If Me.WriteLine(a_message, a_appendTermination) > 0 Then
        DoEvents
        QueryLine = Me.Read(a_maxLength, a_trimEnd)
    Else
        QueryLine = vbNullString
    End If
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Connectable implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets a reference to the connectable <see cref="cc_isr_Winsoc.TcpClient"/> object . </summary>
''' <value>   [<see cref="cc_isr_Winsock.IConnectable"/>]. </value>
Public Property Get Connectable() As cc_isr_Winsock.IConnectable
    Set Connectable = m_tcpClient
End Property

''' <summary>   Returns the connection state of the client. </summary>
''' <value>   [Boolean] True if the TCP Client socket is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
    
    Connected = IConnectable_Connected

End Property

''' <summary>   Returns the connection state of the client. </summary>
''' <value>   [Boolean] True if the TCP Client socket is connected; otherwise, False. </value>
Private Property Get IConnectable_Connected() As Boolean
    
    If m_tcpClient Is Nothing Then
        IConnectable_Connected = False
    Else
        IConnectable_Connected = Me.Connectable.Connected
    End If

End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean
    
    CanConnect = IConnectable_CanConnect

End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Private Property Get IConnectable_CanConnect() As Boolean

    If Me.Connectable Is Nothing Then
        IConnectable_CanConnect = False
    Else
        IConnectable_CanConnect = Me.Connectable.CanConnect
    End If

End Property


''' <summary>   Closes and releases the TCP Client. </summary>
Public Sub Dispose()
    
    IConnectable_Dispose

End Sub

''' <summary>   Closes and releases the TCP Client. </summary>
Private Sub IConnectable_Dispose()

    If Not Me.Connectable Is Nothing Then
        If Me.Connectable.Connected Then
            Me.Connectable.CloseConnection
        End If
    End If
    
    ' dispose the GPIB-Lan controller only after closing the connection.
    
    If Not m_gpibLan Is Nothing Then
        m_gpibLan.Dispose
    End If
    Set m_gpibLan = Nothing
    
    If Not Me.Connectable Is Nothing Then
        Me.Connectable.Dispose
    End If
    
    Set m_tcpClient = Nothing
    
    Set m_stopper = Nothing
    
End Sub

''' <summary>   Opens a TCP Client connection to the specified host at the specified port
''' as sets the connection timeout. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_timeoutMs">   [Optional, Long, 500] The receive timeout in milliseconds. </param>
''' <returns>   True if the TCP Client is connected. </returns>
Public Function OpenConnection(ByVal a_host As String, ByVal a_port As Long, Optional ByVal a_timeoutMs As Long = 500) As Boolean
    
    IConnectable_OpenConnection a_host, a_port, a_timeoutMs

End Function

''' <summary>   Opens a TCP Client connection to the specified host at the specified port
''' as sets the connection timeout. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_timeoutMs">   [Optional, Long, 500] The receive timeout in milliseconds. </param>
''' <returns>   True if the TCP Client is connected. </returns>
Private Function IConnectable_OpenConnection(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_timeoutMs As Long = 500&) As Boolean

    m_usingGpibLan = a_port = GPIB_LAN_PORT_NO
    
    Dim p_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs
    Set p_eventArgs = cc_isr_Winsock.Constructor.CreateConnectionChangingEventArgs(Me.Connected)
    IConnectable_OnConnectionChanging p_eventArgs
    
    If Not p_eventArgs.cancel Then
    
        ' enable/disable the GPIB-Lan controller; this ensure that the controller correctly handles the
        ' Tcp Client connection events.
        
        m_gpibLan.Enabled = Me.UsingGpibLan
        
        Connectable.OpenConnection a_host, a_port, a_timeoutMs
        
        ' update the socket receive timeout.
        
        m_socketReceiveTimeout = a_timeoutMs
        
        OnConnectionChanged Connectable.Connected
        
    End If

    IConnectable_OpenConnection = Connectable.Connected
    
End Function

''' <summary>   Closes the TCP Client connection. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean
    
    CloseConnection = IConnectable_CloseConnection
    
End Function

''' <summary>   Closes the TCP Client connection. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Private Function IConnectable_CloseConnection() As Boolean

    Dim p_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs
    Set p_eventArgs = cc_isr_Winsock.Constructor.CreateConnectionChangingEventArgs(Me.Connected)
    IConnectable_OnConnectionChanging p_eventArgs
    
    If Not p_eventArgs.cancel Then
    
        IConnectable_CloseConnection = Connectable.CloseConnection
    
        OnConnectionChanged Connectable.Connected
    
    End If
    
    IConnectable_CloseConnection = Not Connectable.Connected

End Function

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal a_isConnected As Boolean)
    
    IConnectable_OnConnectionChanged cc_isr_Winsock.Constructor.CreateConnectionChangedEventArgs(a_isConnected)

End Sub

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_eventArgs">   [<<see cref="cc_isr_Winsock.ConnectionchangedEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
    
    RaiseEvent ConnectionChanged(a_eventArgs)

End Sub

''' <summary>   Raises the <see cref="ConnectionChanging"/> event. </summary>
''' <para name="a_eventArgs">   [<<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
    
    RaiseEvent ConnectionChanging(a_eventArgs)

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  TCP Client Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Winsock.TcpClient.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_tcpClient_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
End Sub

''' <summary>   Handles the <see cref="cc_isr_Winsock.TcpClient.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_tcpClient_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub


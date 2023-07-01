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

Private Const READ_AFTER_WRITE_DELAY_DEFAULT = 5

Private m_gpibLan As GpibLanController

Private m_readAfterWriteDelay As Integer

Private m_readTermination As String

Private m_receiveTimeout As Long

Private m_readAfterWriteEnabled As Boolean

Private m_stopper As StopWatch

Private WithEvents m_tcpClient As TcpClient

Private m_usingGpibLan As Boolean

Private m_writeTermination As String

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before change of connection. </summary>
''' <param name="eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event args. </param>
Public Event ConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Handles the Class Initialize event; constructs the <see cref="TcpClient"/>. </summary>
Private Sub Class_Initialize()
    Set m_stopper = cc_isr_Core.Constructor.CreateStopWatch
    Set m_tcpClient = cc_isr_Winsock.Constructor.CreateTcpClient
    m_readTermination = vbLf
    m_writeTermination = vbLf
    m_receiveTimeout = 500
    Me.ReadAfterWriteDelay = READ_AFTER_WRITE_DELAY_DEFAULT
    
    Set m_gpibLan = New GpibLanController
    m_gpibLan.Initialize m_tcpClient, New ErrTracer, m_readTermination, m_writeTermination, READ_AFTER_WRITE_DELAY_DEFAULT
    
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="TcpClient"/>. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Initializes the <see cref="IErrTracer"/> for the <see cref="ViSession"/>. </summary>
''' <param name="tracer">   <see cref="IErrTracer"/> implementation. </param>
Public Sub Initialize(ByVal tracer As IErrTracer)
    
    m_gpibLan.Initialize m_tcpClient, tracer, m_readTermination, m_writeTermination, READ_AFTER_WRITE_DELAY_DEFAULT
   
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
Friend Property Get ReadAfterWriteDelay() As Integer
    ReadAfterWriteDelay = m_readAfterWriteDelay
End Property

''' <summary>   Sets the Read After Write Delay. </summary>
Friend Property Let ReadAfterWriteDelay(ByVal value As Integer)
    m_readAfterWriteDelay = value
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
''' <param name="message">             [String] The message to send to the instrument. </param>
''' <param name="appendTermination">   [Option, Boolean, True] True to append termination to
'''                                    the message. </param>
''' <returns>   [Long] The number of sent characters. </returns>
Public Function WriteLine(ByVal message As String, _
        Optional ByVal appendTermination As Boolean = True) As Long
    
    If Me.UsingGpibLan Then
        WriteLine = Me.GpibLan.SendToDevice(message, appendTermination)
    Else
        If appendTermination Then message = message & Me.WriteTermination
        WriteLine = m_tcpClient.SendMessage(message)
        m_stopper.Wait Me.ReadAfterWriteDelay
    End If

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <remarks>   If using the Prologix device at port 1234, this
''' method first sets the Prologix to auto on (++auto 1) assuming auto was turned off on
''' the previous read on <see cref="WriteLine"/>. </remarks>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <returns>   [String] The received string. </returns>
Public Function Read(Optional ByVal maxLength As Long = &H7FFF, _
                     Optional ByVal trimEnd As Boolean = True) As String
    If Me.UsingGpibLan Then
        Read = Me.GpibLan.ReceiveFromDevice(maxLength, trimEnd)
    Else
        Read = m_tcpClient.ReceiveRaw(maxLength, trimEnd, Me.ReadTermination)
    End If

End Function
    
''' <summary>   Sends a message and receives a reply. </summary>
''' <param name="message">             [String] The message to send to the instrument. </param>
''' <param name="appendTermination">   [Option, Boolean, True] True to append termination to
'''                                    the message. </param>
''' <param name="maxLength">           [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">             [Optional, True] True to return the string without the termination. </param>
''' <returns>   [String] The received string. </returns>
Public Function QueryLine(ByVal message As String, _
        Optional ByVal appendTermination As Boolean = True, _
        Optional ByVal maxLength As Long = &H7FFF, _
        Optional ByVal trimEnd As Boolean = True) As String
    If Me.WriteLine(message, appendTermination) > 0 Then
        DoEvents
        QueryLine = Me.Read(maxLength, trimEnd)
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

    If Me.Connectable.Connected Then
        Me.Connectable.CloseConnection
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
''' <param name="host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="port">        [Long] The port that the server is listening on. </param>
''' <param name="timeoutMs">   [Optional, Long, 500] The receive timeout in milliseconds. </param>
''' <returns>   True if the TCP Client is connected. </returns>
Public Function OpenConnection(ByVal host As String, ByVal port As Long, Optional ByVal timeoutMs As Long = 500) As Boolean
    
    IConnectable_OpenConnection host, port, timeoutMs

End Function

''' <summary>   Opens a TCP Client connection to the specified host at the specified port
''' as sets the connection timeout. </summary>
''' <param name="host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="port">        [Long] The port that the server is listening on. </param>
''' <param name="timeoutMs">   [Optional, Long, 500] The receive timeout in milliseconds. </param>
''' <returns>   True if the TCP Client is connected. </returns>
Private Function IConnectable_OpenConnection(ByVal host As String, ByVal port As Long, _
        Optional ByVal timeoutMs As Long = 500&) As Boolean

    m_usingGpibLan = port = GPIB_LAN_PORT_NO
    
    Dim p_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs
    Set p_eventArgs = cc_isr_Winsock.Constructor.CreateConnectionChangingEventArgs(Me.Connected)
    IConnectable_OnConnectionChanging p_eventArgs
    
    If Not p_eventArgs.cancel Then
    
        ' enable/disable the GPIB-Lan controller; this ensure that the controller correctly handles the
        ' Tcp Client connection events.
        
        m_gpibLan.Enabled = Me.UsingGpibLan
        
        Connectable.OpenConnection host, port, timeoutMs
        
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
''' <para name="isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal isConnected As Boolean)
    IConnectable_OnConnectionChanged cc_isr_Winsock.Constructor.CreateConnectionChangedEventArgs(isConnected)
End Sub

''' <summary>   Raises the <see cref="Connectionchanged"/> event. </summary>
''' <para name="eventArgs">   [<<see cref="cc_isr_Winsock.ConnectionchangedEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanged(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
    RaiseEvent ConnectionChanged(eventArgs)
End Sub

''' <summary>   Raises the <see cref="ConnectionChanging"/> event. </summary>
''' <para name="eventArgs">   [<<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
    RaiseEvent ConnectionChanging(eventArgs)
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  TCP Client Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Winsock.TcpClient.ConnectionChanged"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_tcpClient_ConnectionChanged(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
End Sub

''' <summary>   Handles the <see cref="cc_isr_Winsock.TcpClient.ConnectionChanging"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_tcpClient_ConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub


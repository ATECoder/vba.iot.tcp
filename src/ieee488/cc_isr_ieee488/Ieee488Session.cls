VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ieee488Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cc_isr_Winsock.IConnectable

Private WithEvents m_viSession As ViSession

Private m_identity As String

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event args. </param>
Public Event ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Handles the Class Initialize event; constructs the <see cref="ViSession"/>. </summary>
Private Sub Class_Initialize()
    
    Set m_viSession = New ViSession
    m_viSession.Initialize New ErrTracer
    
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="ViSession"/>. </summary>
Private Sub Class_Terminate()
    
    Me.Dispose

End Sub

''' <summary>   Initializes the <see cref="IErrTracer"/> for the <see cref="Ieee488Session"/>. </summary>
''' <param name="a_tracer">   <see cref="IErrTracer"/> implementation. </param>
Public Sub Initialize(ByVal a_tracer As IErrTracer)
   
    m_viSession.Initialize a_tracer
   
End Sub

''' <summary>   Returns a descriptor identifying the socket. </summary>
''' <value>   The socket identifying descriptor if connected; otherwise, -1. </value>
Public Property Get SocketId() As Long
    If Me.ViSession Is Nothing Then
        SocketId = cc_isr_Winsock.wsock32.ws32_INVALID_SOCKET
    Else
        SocketId = Me.ViSession.SocketId
    End If
End Property

''' <summary>   Sends a message. </summary>
''' <param name="a_message">             [String] The message to send to the instrument. </param>
''' <param name="a_queryEAV">            [Optional, Boolean, True] True to check the Error Available status bit
'''                                      after sending the message if <see cref="UsingGpibLan"/> or VXI. </param>
''' <param name="a_appendTermination">   [Optional, Boolean, True] True to append termination to
'''                                      the message. </param>
''' <returns>   [Long] The number of sent characters. </returns>
Public Function WriteLine(ByVal a_message As String, _
        Optional ByVal a_queryEAV As Boolean = True, _
        Optional ByVal a_appendTermination As Boolean = True) As Long
    
    WriteLine = Me.ViSession.WriteLine(a_message, a_appendTermination)
    
    If Me.ViSession.UsingGpibLan And a_queryEAV Then
    
        ' read the status byte and check for errors.
        Dim p_status As Integer: p_status = Me.ViSession.GpibLan.SerialPoll()
        
    
        ' check if we have an error.
        If ServiceRequests.ErrorAvailable = (p_status And ServiceRequests.ErrorAvailable) Then
        
            ' raise the write error adding some information
            Err.Raise UserDefinedErrors.SocketSendError.Code, "cc_isr_Ieee488.Ieee488Session.WriteLine", _
                      UserDefinedErrors.SocketSendError.Description & _
                      " SRQ=0x" & Hex(p_status) & _
                      ". Error Available (0x" & Hex(ServiceRequests.ErrorAvailable) & _
                      ") after sending " & a_message & " to the instrument at " & _
                      Me.ViSession.SocketAddress & "."
        End If
    
    End If

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="a_awaitMAV">    [Optional, Boolean, True] True to wait for the Message Available
'''                              status bit before reading if <see cref="UsingGpibLan"/> or VXI. </param>
''' <param name="a_maxLength">   [Optional, Long, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">     [Optional, Boolean, True] True to return the string without
'''                              the termination. </param>
''' <returns>   [String] The received message. </returns>
Public Function Read(Optional ByVal a_awaitMAV As Boolean = True, _
        Optional ByVal a_maxLength As Long = &H7FFF, _
        Optional ByVal a_trimEnd As Boolean = True) As String
    
    Dim p_statusByte As Integer
    
    If Me.ViSession.UsingGpibLan And a_awaitMAV Then
        
        ' wait for the message available bits.
        
        p_statusByte = Me.ViSession.GpibLan.AwaitStatus(Me.ViSession.SessionReadTimeout, _
                                                        ServiceRequests.MessageAvailable)
        
    End If
    
    ' either way, try reading the instrument
    
    On Error Resume Next
    If Me.ViSession.SessionReadTimeout > 0 Then
        Read = Me.ViSession.AwaitReading(Me.ViSession.SessionReadTimeout, a_maxLength, a_trimEnd)
    Else
        Read = Me.ViSession.Read(a_maxLength, a_trimEnd)
    End If
    On Error GoTo 0
        
    ' report an error on failure to read.
    If 0 = Len(Read) Then
    
        ' check if message available.
        If 0 = (p_statusByte And ServiceRequests.MessageAvailable) Then
        
            ' raise the read error adding some information
            Err.Raise UserDefinedErrors.TimeoutError.Code, "cc_isr_Ieee488.Ieee488Session.Read", _
                      UserDefinedErrors.TimeoutError.Description & _
                      " SRQ=0x" & Hex(p_statusByte) & _
                      ". No Message Available (0x" & Hex(ServiceRequests.MessageAvailable) & _
                      ") reading the instrument at " & Me.ViSession.SocketAddress & _
                      " with a timeout of " & CStr(Me.ViSession.SessionReadTimeout) & "ms."
                      
        Else
        
            ' raise the read error adding some information
            Err.Raise UserDefinedErrors.TimeoutError.Code, "cc_isr_Ieee488.Ieee488Session.Read", _
                      UserDefinedErrors.TimeoutError.Description & _
                      " SRQ=0x" & Hex(p_statusByte) & _
                      ". Data not received after message available (0x" & Hex(ServiceRequests.MessageAvailable) & _
                      ") reading the instrument at " & Me.ViSession.SocketAddress & _
                      " with a timeout of " & CStr(Me.ViSession.SessionReadTimeout) & "ms."
        
        End If
   
    End If

End Function

''' <summary>   Sends a message and receives a reply. </summary>
''' <param name="a_message">             [String] The message to send to the instrument. </param>
''' <param name="a_queryEAV">            [Optional, Boolean, True] True to check the Error Available status bit
'''                                      after sending the message if <see cref="UsingGpibLan"/> or VXI. </param>
''' <param name="a_awaitMAV">            [Optional, Boolean, True] True to wait for the Message Available
'''                                      status bit before reading if <see cref="UsingGpibLan"/> or VXI. </param>
''' <param name="a_appendTermination">   [Optional, Boolean, True] True to append termination to
'''                                      the message. </param>
''' <param name="a_maxLength">           [Optional, Long, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">             [Optional, Boolean True] True to return the string without
'''                                      the termination. </param>
''' <param name="a_timeout">             [Optional, Integer, 0] positive value indicates
'''                                      the timeout to wait for a reading and message available status. </param>
''' <returns>   [String] The received message. </returns>
Public Function QueryLine(ByVal a_message As String, _
        Optional ByVal a_queryEAV As Boolean = True, _
        Optional ByVal a_awaitMAV As Boolean = True, _
        Optional ByVal a_appendTermination As Boolean = True, _
        Optional ByVal a_maxLength As Long = &H7FFF, _
        Optional ByVal a_trimEnd As Boolean = True) As String
    
    If Me.WriteLine(a_message, a_queryEAV, a_appendTermination) > 0 Then
        DoEvents
        QueryLine = Me.Read(a_awaitMAV, a_maxLength, a_trimEnd)
    Else
        QueryLine = vbNullString
    End If

End Function

''' <summary>   Returns a reference to the VI Session. </summary>
''' <value>   [<see cref="ViSession"/>. </returns>
Public Property Get ViSession() As ViSession
    Set ViSession = m_viSession
End Property

''' <summary>   Clears Status (CLS) command. </summary>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*CLS</c> command by querying <c>*CLS; *OPC?</c></param>
Public Sub ClearExecutionState(Optional ByVal a_awaitOpc As Boolean = True)
    
    Dim p_command As String: p_command = Ieee488Syntax.ClearExecutionStateCommand
    If a_awaitOpc Then
        p_command = p_command & ";" & Ieee488Syntax.OperationCompletedQueryCommand
        Me.QueryLine p_command, False, True
    Else
        Me.WriteLine p_command, False
    End If

End Sub

''' <summary>   Returns the identity. </summary>
''' <value>   [String]. </value>
Public Property Get Identity() As String
    If StringExtensions.IsNullOrEmpty(m_identity) And Me.ViSession.Connected Then
        m_identity = Me.QueryIdentity
    End If
    Identity = m_identity
End Property

''' <summary>   Returns the instrument identity using the *IDN? query command. </summary>
''' <returns>   [String]. </returns>
Public Function QueryIdentity() As String
    m_identity = Me.QueryLine(Ieee488Syntax.IdentityQueryCommand)
    QueryIdentity = m_identity
End Function

''' <summary>   Operation Complete (*OPC) command. </summary>
''' <returns>   [Long] The number of characters that were sent. </returns>
Public Function OperationComplete() As Long
    OperationComplete = Me.WriteLine(Ieee488Syntax.OperationCompleteCommand)
End Function

''' <summary>   Returns 1 if operation was completed; otherwise 0. </summary>
''' <returns>   A string. 1 if completed; otherwise 0. </returns>
Public Function QueryOperationCompleted() As String
    QueryOperationCompleted = Me.QueryLine(Ieee488Syntax.OperationCompletedQueryCommand)
End Function

''' <summary>   Returns option (instrument specific). </summary>
''' <returns>   [String]. </returns>
Public Function QueryOptions() As String
    
    QueryOptions = Me.QueryLine(Ieee488Syntax.OptionsQueryCommand)
    
End Function

''' <summary>   Reads the status byte. </summary>
''' <param name="a_canQuery">   [Optional, Boolean, False] True to send <c>*STB?</c>
'''                             if not <see cref="UsingGpibLan"/>. </param>
''' <returns>   The status byte. </returns>
Public Function ReadStatusByte(Optional ByVal a_canQuery As Boolean = False) As Integer
    
    If Me.ViSession.UsingGpibLan Then
        ReadStatusByte = Me.ViSession.GpibLan.SerialPoll()
    ElseIf a_canQuery Then
        ReadStatusByte = Me.QueryServiceRequestStatus()
    Else
        ' querying the service request status could cause
        ' a query unterminated error.
        ReadStatusByte = 0
    End If
    
End Function

''' <summary>   Checks if a service is requested. </summary>
''' <param name="canQuery">   [Optional, Boolean, False] True to send <c>*STB?</c>
''' if not <see cref="UsingGpibLan"/>. </param>
''' <returns>   The status byte. </returns>
Public Function ServiceRequested(Optional ByVal a_canQuery As Boolean = False) As Boolean
    
    If Me.ViSession.UsingGpibLan Then
        ServiceRequested = Me.ViSession.GpibLan.ServiceRequested()
    ElseIf a_canQuery Then
        ServiceRequested = ServiceRequests.RequestingService = _
            (ServiceRequests.RequestingService And Me.QueryServiceRequestStatus())
    Else
        ServiceRequested = False
    End If
    
End Function

''' <summary>   Issues a wait (*WAI) command. </summary>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                            after issuing the <c>*WAI</c> command by querying <c>*WAI; *OPC?</c></param>
Public Sub Wait(Optional ByVal a_awaitOpc As Boolean = True)
    
    Dim p_command As String: p_command = Ieee488Syntax.WaitCommand
    If a_awaitOpc Then
        p_command = p_command & ";" & Ieee488Syntax.OperationCompletedQueryCommand
        Me.QueryLine p_command, False, True
    Else
        Me.WriteLine p_command
    End If

End Sub

''' <summary>   Enables standard events using the (*ESE {0}) command. </summary>
''' <param name="a_bitMask">    [Integer] Defines the bits corresponding to the standard events
'''                             to enable. <param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*ESE #</c> command by querying <c>*ESE #; *OPC?</c></param>
Public Sub EnableStandardEvents(ByVal a_bitMask As Integer, _
        Optional ByVal a_awaitOpc As Boolean = True)
    
    Dim p_command As String
    p_command = StringExtensions.StringFormat(Ieee488Syntax.StandardEventEnableCommand, a_bitMask)
    If a_awaitOpc Then
        p_command = p_command & ";" & Ieee488Syntax.OperationCompletedQueryCommand
        Me.QueryLine p_command, False, True
    Else
        Me.WriteLine p_command
    End If
    
End Sub

''' <summary>   Returns the standard events enable byte using the *ESE? query command. </summary>
''' <returns>   [Integer] The standard event register enable status value. </returns>
Public Function QueryStandardEventsEnable() As Integer
    
    Dim p_enabledBits As String
    p_enabledBits = Me.QueryLine(Ieee488Syntax.StandardEventEnableQueryCommand)
    QueryStandardEventsEnable = CInt(p_enabledBits)

End Function

''' <summary>   Returns the standard events status byte using the *ESR? query command. </summary>
''' <returns>   The standard event register status value. </returns>
Public Function QueryStandardEventsStatus() As Integer
    
    Dim p_StatusBits As String
    p_StatusBits = Me.QueryLine(Ieee488Syntax.StandardEventStatusQueryCommand)
    QueryStandardEventsStatus = CInt(p_StatusBits)
    
End Function

''' <summary>   Enables Service Request using the (*SRE {0}) command. </summary>
''' <param name="a_bitMask">    [Integer] Defines the bits corresponding to the service request
'''                             events to enable. <param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*SRE #</c> command by querying
'''                             <c>*SRE #; *OPC?</c></param>
Public Sub EnableServiceRequest(ByVal a_bitMask As Integer, _
        Optional ByVal a_awaitOpc As Boolean = True)
    
    Dim p_command As String
    p_command = StringExtensions.StringFormat(Ieee488Syntax.ServiceRequestEnableCommand, a_bitMask)
    If a_awaitOpc Then
        p_command = p_command & ";" & Ieee488Syntax.OperationCompletedQueryCommand
        Me.QueryLine p_command, False, True
    Else
        Me.WriteLine p_command
    End If
    
End Sub

''' <summary>   Returns the Service Request enable byte using the *SRE? query command. </summary>
''' <returns>   The service request event register enable status value. </returns>
Public Function QueryServiceRequestEnable() As Integer
    
    Dim p_enabledBits As String
    p_enabledBits = Me.QueryLine(Ieee488Syntax.ServiceRequestEnableQueryCommand)
    QueryServiceRequestEnable = CInt(p_enabledBits)
    
End Function

''' <summary>   Returns the Service Request status byte using the *STB? query command. </summary>
''' <returns>   The service request event register status value. </returns>
Public Function QueryServiceRequestStatus() As Integer
    
    Dim p_eventStatus As String
    p_eventStatus = Me.QueryLine(Ieee488Syntax.ServiceRequestQueryCommand)
    QueryServiceRequestStatus = CInt(p_eventStatus)
    
End Function

''' <summary>   Reset to known state (*RST) command. </summary>
''' <param name="awaitOpc">    [Optional, Boolean, True] True to wait for operation completion
'''                            after issuing the <c>*SRE</c> command by querying <c>*SRE; *OPC?</c></param>
Public Sub ResetKnownState(Optional ByVal a_awaitOpc As Boolean = True)
    
    Dim p_command As String: p_command = Ieee488Syntax.ResetKnownStateCommand
    If a_awaitOpc Then
        p_command = p_command & ";" & Ieee488Syntax.OperationCompletedQueryCommand
        Me.QueryLine p_command, False, True
    Else
        Me.WriteLine p_command, False
    End If
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Connectable implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets a reference to the <see cref="cc_isr_Winsock.IConnectable"/>
''' <see cref="ViSession">VI Session</see>. </summary>
''' <value>   [<see cref="cc_isr_Winsock.IConnectable"/>]. </value>
Public Property Get Connectable() As cc_isr_Winsock.IConnectable
    Set Connectable = m_viSession
End Property

''' <summary>   Returns the connection state of the <see cref="ViSession">VI Session</see>. </summary>
''' <value>   [Boolean] True if the <see cref="ViSession">VI Session</see> is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
    
    Connected = IConnectable_Connected

End Property

''' <summary>   Returns the connection state of the <see cref="cc_isr_Winsock.IConnectable"/>. </summary>
''' <value>   [Boolean] True if the <see cref="cc_isr_Winsock.IConnectable"/> is connected; otherwise, False. </value>
Private Property Get IConnectable_Connected() As Boolean
    
    If Me.Connectable Is Nothing Then
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

''' <summary>   Closes the connection and releases the <see cref="ViSession">VI Session</see>. </summary>
Public Sub Dispose()
    
    IConnectable_Dispose

End Sub

''' <summary>   Closes the connection and releases the <see cref="cc_isr_Winsock.IConnectable"/> </summary>
Private Sub IConnectable_Dispose()

    If Not Me.Connectable Is Nothing Then
        If Me.Connectable.Connected Then
            Me.Connectable.CloseConnection
        End If
    End If

    If Not Me.Connectable Is Nothing Then
        Me.Connectable.Dispose
    End If
    
    Set m_viSession = Nothing
    
End Sub

''' <summary>   Opens a <see cref="ViSession">VI Session</see> connection to the specified host
''' at the specified port as sets the socket receive timeout. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name=a_"timeoutMs">   [Optional, Long, 500] The socket receive timeout in milliseconds. </param>
''' <returns>   [Boolean] True if the VI Session is connected. </returns>
Public Function OpenConnection(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_timeoutMs As Long = 500) As Boolean
    
    IConnectable_OpenConnection a_host, a_port, a_timeoutMs

End Function

''' <summary>   Opens a <see cref="ViSession">VI Session</see> connection to the specified host
''' at the specified port as sets the socket receive timeout. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_timeoutMs">   [Optional, Long, 500] The socket receive timeout in milliseconds. </param>
''' <returns>   [Boolean] True if the VI Session is connected. </returns>
Private Function IConnectable_OpenConnection(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_timeoutMs As Long = 500&) As Boolean

    Dim p_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs
    Set p_eventArgs = cc_isr_Winsock.Constructor.CreateConnectionChangingEventArgs(Me.Connected)
    IConnectable_OnConnectionChanging p_eventArgs
    
    If Not p_eventArgs.cancel Then
    
        Connectable.OpenConnection a_host, a_port, a_timeoutMs
        
        OnConnectionChanged Connectable.Connected
        
    End If

    IConnectable_OpenConnection = Connectable.Connected
    
End Function

''' <summary>   Closes the <see cref="cc_isr_Winsock.IConnectable"/> connection. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean
    
    CloseConnection = IConnectable_CloseConnection
    
End Function

''' <summary>   Closes the <see cref="cc_isr_Winsock.IConnectable"/> connection. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Private Function IConnectable_CloseConnection() As Boolean

    Dim p_args As cc_isr_Winsock.ConnectionChangingEventArgs
    Set p_args = cc_isr_Winsock.Constructor.CreateConnectionChangingEventArgs(Me.Connected)
    IConnectable_OnConnectionChanging p_args
   
    If Not p_args.cancel Then
    
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
''' <para name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionchangedEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
    
    RaiseEvent ConnectionChanged(a_eventArgs)

End Sub

''' <summary>   Raises the <see cref="ConnectionChanging"/> event. </summary>
''' <para name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
    
    RaiseEvent ConnectionChanging(a_eventArgs)

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  VI Session Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Ieee488.ViSession.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_viSession_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
End Sub

''' <summary>   Handles the <see cref="cc_isr_Ieee488.ViSession.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_viSession_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub


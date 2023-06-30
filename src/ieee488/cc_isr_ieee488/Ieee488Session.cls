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

Private WithEvents m_viSession As ViSession
Private m_identity As String

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before change of connection. </summary>
''' <param name="eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event args. </param>
Public Event ConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Handles the Class Initialize event; constructs the <see cref="ViSession"/>. </summary>
Private Sub Class_Initialize()
    
    Set m_viSession = New ViSession
    
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="ViSession"/>. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Closes the connection and releases the <see cref="ViSession"/>. </summary>
Public Sub Dispose()
    
    If Me.Connected Then
        Me.CloseConnection
    End If
    
    If Not m_viSession Is Nothing Then
        m_viSession.Dispose
    End If
    Set m_viSession = Nothing

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


''' <summary>   Returns the connection state of the <see cref="ViSession"/>. </summary>
''' <value>   [Boolean] True if the <see cref="ViSession"/> is connected via the TCP Client socket;
''' otherwise, False. </value>
Public Property Get Connected() As Boolean
    
    If m_viSession Is Nothing Then
        Connected = False
    Else
        Connected = Me.ViSession.Connected
    End If

End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean
    
    If Me.ViSession Is Nothing Then
        CanConnect = False
    Else
        CanConnect = Me.ViSession.CanConnect
    End If

End Property

''' <summary>   Opens a TCP Client connection to the specified host at the specified port
''' as sets the connection timeout. </summary>
''' <param name="host">        An IPv4 dotted-decimal host address. </param>
''' <param name="port">        The port that the server is listening on. </param>
''' <param name="timeoutMs">   [Optional, 500] The receive timeout in milliseconds. </param>
''' <returns>   True if the TCP Client is connected. </returns>
Public Function OpenConnection(ByVal host As String, ByVal port As Long, Optional ByVal timeoutMs As Long) As Boolean
   
    Dim eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs
    Set eventArgs = cc_isr_Winsock.Constructor.CreateConnectionChangingEventArgs(Me.Connected)
    OnConnectionChanging eventArgs
    
    If Not eventArgs.cancel Then
    
        OpenConnection = Me.ViSession.OpenConnection(host, port, timeoutMs)
        
        OnConnectionChanged Me.Connected
     
    End If
   
    OpenConnection = Me.Connected
   
End Function

''' <summary>   Closes the TCP Client connection. </summary>
''' <returns>   True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    Dim p_args As cc_isr_Winsock.ConnectionChangingEventArgs
    Set p_args = cc_isr_Winsock.Constructor.CreateConnectionChangingEventArgs(Me.Connected)
    OnConnectionChanging p_args
    
    If Not p_args.cancel Then
    
        CloseConnection = Me.ViSession.CloseConnection
        
        OnConnectionChanged Me.Connected
     
    End If

    CloseConnection = Not Me.Connected

End Function

''' <summary>   Sends a message. </summary>
''' <param name="message">                 [String] The message to send to the instrument. </param>
''' <param name="queryStatusAfterWrite">   [Option, Boolean, True] True to read the status byte and
'''                                        check for errors after sending the message, which
'''                                        requires <see cref="UsingGpibLan"/> or VXI support. </param>
''' <param name="appendTermination">       [Option, Boolean, True] True to append termination to
'''                                        the message. </param>
''' <returns>   [Long] The number of sent characters. </returns>
Public Function WriteLine(ByVal message As String, _
        Optional ByVal queryStatusAfterWrite As Boolean = True, _
        Optional ByVal appendTermination As Boolean = True) As Long
    
    WriteLine = Me.ViSession.WriteLine(message, appendTermination)
    
    If Me.ViSession.UsingGpibLan And queryStatusAfterWrite Then
    
        ' read the status byte and check for errors.
        Dim p_status As Integer: p_status = Me.ViSession.GpibLan.SerialPoll()
        
    
        ' check if we have an error.
        If ServiceRequests.ErrorAvailable = (p_status And ServiceRequests.ErrorAvailable) Then
        
            ' raise the write error adding some information
            Err.Raise UserDefinedErrors.SocketSendError.Code, "cc_isr_Ieee488.Ieee488Session.WriteLine", _
                      UserDefinedErrors.SocketSendError.Description & _
                      " SRQ=0x" & Hex(p_status) & _
                      ". Error Available (0x" & Hex(ServiceRequests.ErrorAvailable) & _
                      ") after sending " & message & " to the instrument at " & _
                      Me.ViSession.SocketAddress & "."
        End If
    
    End If

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="queryStatusBeforeRead">   [Option, Boolean, True] True to read the status byte and
'''                                        check for errors before reading, which
'''                                        requires <see cref="UsingGpibLan"/> or VXI support. </param>
''' <param name="maxLength">               [Optional, Long, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">                 [Optional, Boolean True] True to return the string without
'''                                         the termination. </param>
''' <returns>   [String] The received message. </returns>
Public Function Read(Optional ByVal queryStatusBeforeRead As Boolean = True, _
        Optional ByVal maxLength As Long = &H7FFF, _
        Optional ByVal trimEnd As Boolean = True) As String
    
    If Me.ViSession.UsingGpibLan And queryStatusBeforeRead Then
        
        ' check if we have data to read.
    
        ' read the status byte and check for errors.
        Dim StatusByte As Integer: StatusByte = Me.ViSession.GpibLan.SerialPoll()
    
        ' check if message available.
        If 0 = (StatusByte And ServiceRequests.MessageAvailable) Then
        
            ' raise the write error adding some information
            Err.Raise UserDefinedErrors.InvalidOperationError.Code, "cc_isr_Ieee488.Ieee488Session.Read", _
                      UserDefinedErrors.InvalidOperationError.Description & _
                      " SRQ=0x" & Hex(StatusByte) & _
                      ". No Message Available (0x" & Hex(ServiceRequests.MessageAvailable) & _
                      ") trying reading the instrument at " & Me.ViSession.SocketAddress & "."
        
        End If
    
    End If
    
    Read = Me.ViSession.Read(maxLength, trimEnd)

End Function
    
''' <summary>   Sends a message and receives a reply. </summary>
''' <param name="message">                 [String] The message to send to the instrument. </param>
''' <param name="queryStatusAfterWrite">   [Option, Boolean, True] True to read the status byte and
'''                                        check for errors after sending the message, which
'''                                        requires <see cref="UsingGpibLan"/> or VXI support. </param>
''' <param name="queryStatusBeforeRead">   [Option, Boolean, True] True to read the status byte and
'''                                        check for errors before reading, which
'''                                        requires <see cref="UsingGpibLan"/> or VXI support. </param>
''' <param name="appendTermination">       [Option, Boolean, True] True to append termination to
'''                                        the message. </param>
''' <param name="maxLength">               [Optional, Long, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">                 [Optional, Boolean True] True to return the string without
'''                                         the termination. </param>
''' <returns>   [String] The received message. </returns>
Public Function QueryLine(ByVal message As String, _
        Optional ByVal queryStatusAfterWrite As Boolean = True, _
        Optional ByVal queryStatusBeforeRead As Boolean = True, _
        Optional ByVal appendTermination As Boolean = True, _
        Optional ByVal maxLength As Long = &H7FFF, _
        Optional ByVal trimEnd As Boolean = True) As String
    
    If Me.WriteLine(message, queryStatusAfterWrite, appendTermination) > 0 Then
        DoEvents
        QueryLine = Me.Read(queryStatusBeforeRead, maxLength, trimEnd)
    Else
        QueryLine = ""
    End If

End Function

''' <summary>   Returns a reference to the VI Session. </summary>
''' <value>   [<see cref="ViSession"/>. </returns>
Public Property Get ViSession() As ViSession
    Set ViSession = m_viSession
End Property

''' <summary>   Clears Status (CLS) command. </summary>
''' <param name="awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                           after issuing the <c>*CLS</c> command by querying <c>*CLS; *OPC?</c></param>
Public Sub ClearExecutionState(Optional ByVal awaitOpc As Boolean = True)
    If awaitOpc Then
        Me.QueryLine Ieee488Syntax.ClearExecutionStateOpcQueryCommand, False, False
    Else
        Me.WriteLine Ieee488Syntax.ClearExecutionStateCommand, False
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
''' <param name="canQuery">   [Optional, False] True to send <c>*STB?</c>
''' if not <see cref="UsingGpibLan"/>. </param>
''' <returns>   The status byte. </returns>
Public Function ReadStatusByte(Optional ByVal canQuery As Boolean = False) As Integer
    
    If Me.ViSession.UsingGpibLan Then
        ReadStatusByte = Me.ViSession.GpibLan.SerialPoll()
    ElseIf canQuery Then
        ReadStatusByte = Me.QueryServiceRequestStatus()
    Else
        ' querying the service request status could cause
        ' a query unterminated error.
        ReadStatusByte = 0
    End If
    
End Function

''' <summary>   Checks if a service is requested. </summary>
''' <param name="canQuery">   [Optional, False] True to send <c>*STB?</c>
''' if not <see cref="UsingGpibLan"/>. </param>
''' <returns>   The status byte. </returns>
Public Function ServiceRequested(Optional ByVal canQuery As Boolean = False) As Boolean
    
    If Me.ViSession.UsingGpibLan Then
        ServiceRequested = Me.ViSession.GpibLan.ServiceRequested()
    ElseIf canQuery Then
        ServiceRequested = ServiceRequests.RequestingService = (ServiceRequests.RequestingService And Me.QueryServiceRequestStatus())
    Else
        ServiceRequested = False
    End If
    
End Function

''' <summary>   Issues a wait (*WAI) command. </summary>
''' <param name="awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                           after issuing the <c>*WAI</c> command by querying <c>*WAI; *OPC?</c></param>
Public Sub Wait(Optional ByVal awaitOpc As Boolean = True)
    If awaitOpc Then
        Me.QueryLine Ieee488Syntax.WaitOpcQueryCommand, False, False
    Else
        Me.WriteLine Ieee488Syntax.WaitCommand
    End If
End Sub

''' <summary>   Enables standard events using the (*ESE {0}) command. </summary>
''' <param name="standardEvents">   The standard event bits to enable. <param>
''' <param name="awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                           after issuing the <c>*ESE #</c> command by querying <c>*ESE #; *OPC?</c></param>
Public Sub EnableStandardEvents(ByVal StandardEvents As Integer, _
        Optional ByVal awaitOpc As Boolean = True)
    
    Dim command As String
    If awaitOpc Then
        command = StringExtensions.StringFormat(StandardEventEnableOpcQueryCommand, StandardEvents)
        Me.QueryLine command, False, False
    Else
        command = StringExtensions.StringFormat(StandardEventEnableCommand, StandardEvents)
        Me.WriteLine Ieee488Syntax.WaitCommand
    End If
    
End Sub

''' <summary>   Returns the standard events enable byte using the *ESE? query command. </summary>
''' <returns>   The standard event register enable status value. </returns>
Public Function QueryStandardEventsEnable() As Integer
    
    Dim StandardEvents As String
    StandardEvents = Me.QueryLine(Ieee488Syntax.StandardEventEnableQueryCommand)
    QueryStandardEventsEnable = CInt(StandardEvents)

End Function

''' <summary>   Returns the standard events status byte using the *ESR? query command. </summary>
''' <returns>   The standard event register status value. </returns>
Public Function QueryStandardEventsStatus() As Integer
    
    Dim StandardEvents As String
    StandardEvents = Me.QueryLine(Ieee488Syntax.StandardEventStatusQueryCommand)
    QueryStandardEventsStatus = CInt(StandardEvents)
    
End Function

''' <summary>   Enables Service Request using the (*SRE {0}) command. </summary>
''' <param name="ServiceRequest">   [Integer] The standard event bits to enable. <param>
''' <param name="awaitOpc">         [Optional, Boolean, True] True to wait for operation completion
'''                                 after issuing the <c>*SRE #</c> command by querying
'''                                 <c>*SRE #; *OPC?</c></param>
Public Sub EnableServiceRequest(ByVal serviceRequest As Integer, _
        Optional ByVal awaitOpc As Boolean = True)
    
    Dim command As String
    If awaitOpc Then
        command = StringExtensions.StringFormat(ServiceRequestEnableOpcQueryCommand, serviceRequest)
        Me.QueryLine command, False, False
    Else
        command = StringExtensions.StringFormat(ServiceRequestEnableCommand, serviceRequest)
        Me.WriteLine command
    End If
    
End Sub

''' <summary>   Returns the Service Request enable byte using the *SRE? query command. </summary>
''' <returns>   The service request event register enable status value. </returns>
Public Function QueryServiceRequestEnable() As Integer
    
    Dim serviceRequest As String
    serviceRequest = Me.QueryLine(Ieee488Syntax.ServiceRequestEnableQueryCommand)
    QueryServiceRequestEnable = CInt(serviceRequest)
    
End Function

''' <summary>   Returns the Service Request status byte using the *STB? query command. </summary>
''' <returns>   The service request event register status value. </returns>
Public Function QueryServiceRequestStatus() As Integer
    
    Dim serviceRequest As String
    serviceRequest = Me.QueryLine(Ieee488Syntax.ServiceRequestQueryCommand)
    QueryServiceRequestStatus = CInt(serviceRequest)
    
End Function

''' <summary>   Reset to known state (*RST) command. </summary>
''' <param name="awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                           after issuing the <c>*SRE</c> command by querying <c>*SRE; *OPC?</c></param>
Public Sub ResetKnownState(Optional ByVal awaitOpc As Boolean = True)
    
    If awaitOpc Then
        Me.QueryLine Ieee488Syntax.ResetKnownStateOpcQueryCommand, False, False
    Else
        Me.WriteLine Ieee488Syntax.ResetKnownStateCommand, False
    End If
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Connection Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal isConnected As Boolean)
    RaiseEvent ConnectionChanged( _
            cc_isr_Winsock.Constructor.CreateConnectionChangedEventArgs(isConnected))
End Sub

''' <summary>   Raises the <see cref="ConnectionChanging"/> event. </summary>
''' <para name="eventArgs">   [<<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>]. </param>
Private Sub OnConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
    RaiseEvent ConnectionChanging(eventArgs)
End Sub


' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  VI Session Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Ieee488.m_viSession.ConnectionChanged"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_viSession_ConnectionChanged(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

End Sub

''' <summary>   Handles the <see cref="cc_isr_Ieee488.ViSession.ConnectionChanging"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_viSession_ConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

End Sub


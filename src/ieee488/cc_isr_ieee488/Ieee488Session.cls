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

Private m_viSession As ViSession
Private m_identity As String
Private m_readTimeout As Integer

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByRef eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

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
''' <value>   True if the <see cref="ViSession"/> is connected via the TCP Client socket;
''' otherwise, False. </value>
Public Property Get Connected() As Boolean
    
    If m_viSession Is Nothing Then
        Connected = False
    Else
        Connected = Me.ViSession.Connected
    End If

End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   True if a connection can be made. </value>
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
   
    OpenConnection = Me.ViSession.OpenConnection(host, port, timeoutMs)
   
    OnConnectionChanged Me.Connected
   
End Function

''' <summary>   Closes the TCP Client connection. </summary>
''' <returns>   True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    CloseConnection = Me.ViSession.CloseConnection
    
    OnConnectionChanged Me.Connected

End Function

''' <summary>   Sends a message. </summary>
''' <param name="message">   The message. </param>
''' <returns>   The number of sent characters. </returns>
Public Function WriteLine(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    
    WriteLine = Me.ViSession.WriteLine(message, appendTermination)
    
    If Me.ViSession.UsingPrologix Then
    
        ' read the status byte and check for errors.
        Dim statusByte As Integer: statusByte = Me.SeriaPoll()
    
        ' check if we have an error.
        If ServiceRequests.ErrorAvailable = (statusByte And ServiceRequests.ErrorAvailable) Then
        
            ' raise the write error adding some information
            Err.Raise UserDefinedErrors.SocketSendError.Code, "cc_isr_Ieee488.Ieee488Session.WriteLine", _
                      UserDefinedErrors.SocketSendError.Description & _
                      " SRQ=0x" & Hex(statusByte) & _
                      ". Error Available (0x" & Hex(ServiceRequests.ErrorAvailable) & _
                      ") after sending " & message & " to the instrument at " & _
                      Me.ViSession.SocketAddress & "."
        End If
    
    End If

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <returns>   The received string. </returns>
Public Function Read(Optional ByVal maxLength As Long = &H7FFF, _
                     Optional ByVal trimEnd As Boolean = True) As String
    
    If Me.ViSession.UsingPrologix Then
        
        ' check if we have data to read.
    
        ' read the status byte and check for errors.
        Dim statusByte As Integer: statusByte = Me.ReadStatusByte()
    
        ' check if message available.
        If 0 = (statusByte And ServiceRequests.MessageAvailable) Then
        
            ' raise the write error adding some information
            Err.Raise UserDefinedErrors.InvalidOperationError.Code, "cc_isr_Ieee488.Ieee488Session.Read", _
                      UserDefinedErrors.InvalidOperationError.Description & _
                      " SRQ=0x" & Hex(statusByte) & _
                      ". No Message Available (0x" & Hex(ServiceRequests.MessageAvailable) & _
                      ") trying reading the instrument at " & Me.ViSession.SocketAddress & "."
        
        End If
    
    End If
    
    Read = Me.ViSession.Read(maxLength, trimEnd)

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

''' <summary>   Returns a reference to the VI Session. </summary>
''' <value>   A ViSession. </returns>
Public Property Get ViSession() As ViSession
    Set ViSession = m_viSession
End Property

''' <summary> Clears Status (CLS) command. </summary>
''' <returns> the number of characters that were sent. </returns>
Public Function ClearExecutionState() As Integer
    ClearExecutionState = Me.WriteLine(Ieee488Syntax.ClearExecutionStateCommand)
End Function

''' <summary>   Returns the identity. </summary>
''' <value>   A string. </returns>
Public Property Get identity() As String
    If StringExtensions.IsNullOrEmpty(m_identity) And Me.ViSession.Connected Then
        m_identity = Me.QueryIdentity
    End If
    identity = m_identity
End Property

''' <summary>   Returns the instrument identity using the *IDN? query command. </summary>
''' <returns>   A string. </returns>
Public Function QueryIdentity() As String
    m_identity = Me.QueryLine(Ieee488Syntax.IdentityQueryCommand)
    QueryIdentity = m_identity
End Function

''' <summary> Operation Complete (*OPC) command. </summary>
''' <returns> The number of characters that were sent. </returns>
Public Function OperationComplete() As Integer
    OperationComplete = Me.WriteLine(Ieee488Syntax.OperationCompleteCommand)
End Function

''' <summary>   Returns 1 if operation was completed; otherwise 0. </summary>
''' <returns>   A string. 1 if completed; otherwise 0. </returns>
Public Function QueryOperationCompleted() As String
    QueryOperationCompleted = Me.QueryLine(Ieee488Syntax.OperationCompletedQueryCommand)
End Function

''' <summary>   Returns option (instrument specific). </summary>
''' <returns>   A string. </returns>
Public Function QueryOptions() As String
    
    QueryOptions = Me.QueryLine(Ieee488Syntax.OptionsQueryCommand)
    
End Function

''' <summary> Reads the status byte. </summary>
''' <param name="canQuery">   [Optional, False] True to send <c>*STB?</c>
''' if not <see cref="UsingPrologix"/>. </param>
''' <returns>   The status byte. </returns>
Public Function ReadStatusByte(Optional ByVal canQuery As Boolean = False) As Integer
    
    If Me.ViSession.UsingPrologix Then
        ReadStatusByte = Me.SerialPoll()
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
''' if not <see cref="UsingPrologix"/>. </param>
''' <returns>   The status byte. </returns>
Public Function SRQ(Optional ByVal canQuery As Boolean = False) As Boolean
    
    If Me.ViSession.UsingPrologix Then
        SRQ = Me.IsSRQ()
    ElseIf canQuery Then
        SRQ = ServiceRequests.RequestingService = (ServiceRequests.RequestingService And Me.QueryServiceRequestStatus())
    Else
        SRQ = 0
    End If
    
End Function

''' <summary> Issues a wait (*WAI) command. </summary>
''' <returns>   The number of characters that were sent. </returns>
Public Function Wait() As Integer
    
    Wait = Me.WriteLine(Ieee488Syntax.WaitCommand)
    
End Function

''' <summary> Enables standard events using the (*ESE {0}) command. </summary>
''' <param name="standardEvents">   The standard event bits to enable. <param>
''' <returns>   The number of characters that were sent. </returns>
Public Function EnableStandardEvents(ByVal StandardEvents As Integer) As Integer
    
    Dim command As String
    command = StringExtensions.StringFormat(StandardEventEnableCommandFormat, StandardEvents)
    EnableStandardEvents = Me.WriteLine(command)
    
End Function

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
''' <param name="ServiceRequest">   The standard event bits to enable. <param>
''' <returns>   The number of characters that were sent. </returns>
Public Function EnableServiceRequest(ByVal serviceRequest As Integer) As Integer
    
    Dim command As String
    command = StringExtensions.StringFormat(ServiceRequestEnableCommandFormat, serviceRequest)
    EnableServiceRequest = Me.WriteLine(command)
    
End Function

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

''' <summary> Reset to known state (*RST) command. </summary>
''' <returns>   The number of characters that were sent. </returns>
Public Function ResetKnownState() As Integer
    
    ResetKnownState = Me.WriteLine(Ieee488Syntax.ResetKnownStateCommand)
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' PROLOGIX GPIB METHODS

''' <summary>   Configure the GPIB address. The meaning of the GPIB address depends on the
''' operating mode of the controller.
'''
''' In CONTROLLER mode, it refers to the GPIB address of the instrument being controlled.
''' In DEVICE mode, it is the address of the GPIB peripheral that Prologix GPIB-ETHERNET
''' controller is emulating. </summary>
''' <remarks>
''' An optional secondary address may also be specified.
'''
''' Internally, the secondary address, which is offset by 96, must be separated from the primary
''' address by a space character. Specifying secondary address has no effect in DEVICE mode.
''' </remarks>
''' <para name="primaryAddress">     [Integer] Specifies the primary GPIB address
'''                                   between 0 and 30. </param>
''' <para name="secondaryAddress)>   [Optional, Integer] Specifies the secondard GPIB Address
'''                                  between 0 and 30. </param>
''' <returns>   The status byte. </returns>
Public Sub SetGpibAddress(ByVal primaryAddress As Integer, _
                          Optional ByVal secondaryAddress As Integer = -1)
    
    If Me.ViSession.UsingPrologix Then
        If primaryAddress >= 0 And secondaryAddress < 0 Then
            Me.ViSession.WriteLine ("++addr " & CStr(primaryAddress))
        ElseIf primaryAddress >= 0 And secondaryAddress >= 0 Then
            Me.ViSession.QueryLine ("++spoll " & CStr(primaryAddress) & _
                " " & CStr(secondaryAddress + 96))
        End If
    End If
    
End Sub

''' <summary>   Query the GPIB address. </summary>
''' <value>    An string including the primary and secondary address separated by a space. </value>
Public Function GetGpibAddress() As String
    
    GetGpibAddress = "-1"
    If Me.ViSession.UsingPrologix Then
        GetGpibAddress = CInt(Me.ViSession.QueryLine("++addr"))
        Dim addresses() As String: addresses = Split("", " ")
        If Len(addresses) = 1 Then
            GetGpibAddress = addresses(1)
            If Len(addresses) > 1 Then
                GetGpibAddress = GetGpibAddress & " " & addresses(2)
            End If
        End If
        
    End If
    
End Function


''' <summary> Issues a local lockout (++llo). </summary>
''' <remarks>   Valid if <see cref="UsingPrologix"/>. </remarks>
Public Sub LocalLockout()
    
    If Me.ViSession.UsingPrologix Then
        Me.ViSession.WriteLine "++llo"
    End If
    
End Sub

''' <summary> Issues a Go To Local (GTL) (++loc). </summary>
''' <remarks>   Valid if <see cref="UsingPrologix"/>. </remarks>
Public Sub GoToLocal()
    
    If Me.ViSession.UsingPrologix Then
        Me.ViSession.WriteLine "++loc"
    End If
    
End Sub

''' <summary>   Gets the device read timeout in Milliseconds. </summary>
''' <value>   An Integer. </value>
Public Property Get ReadTimeout() As Integer
    ReadTimeout = m_readTimeout
End Property

''' <summary>   Sets the device read timeout, in milliseconds, to be used in the
'''  Prologix read and spoll commands. Timeout may be set to any value between 1 and 3000
''' milliseconds.
''' </summary>
''' <param name="timeoutMs">   The timeout interval in milliseconds. </param>
Public Sub SetReadTimeout(ByVal timeoutMs As Integer)
    
    If Me.ViSession.UsingPrologix Then
        If Me.Connected Then
            
            If timeoutMs < 1 Then timeoutMs = 1
            If timeoutMs > 3000 Then timeoutMs = 3000
            
            If timeoutMs <> m_readTimeout Then
                Me.ViSession.WriteLine "++read_tmo_ms " & CStr(timeoutMs)
            End If
            m_readTimeout = timeoutMs
        End If
    End If
    
End Sub

''' <summary>   Sets the device read timeout, in milliseconds, that is used in the
'''  Prologix read and spoll commands.
''' </summary>
Public Function GetReadTimeout() As Integer
    
    If Me.ViSession.UsingPrologix Then
        If Me.Connected Then
            
            m_readTimeout = CInt(Me.ViSession.QueryLine("++read_tmo_ms"))
            
            If timeoutMs <> m_readTimeout Then
            End If
            m_readTimeout = timeoutMs
        End If
    Else
        m_readTimeout = -1
    End If
    
End Function

''' <summary> Issues an SDC. </summary>
''' <remarks>   Valid if <see cref="UsingPrologix"/>. </remarks>
Public Sub SelectiveDeviceClear()
    
    If Me.ViSession.UsingPrologix Then
        Me.ViSession.WriteLine "++clr"
    End If
    
End Sub

''' <summary> Performs a serial poll of the instrument at the current or the specified address.
''' If no address is specified then this command serial polls the currently addressed instrument (as
''' set by a previous ++addr command). This command uses the time-out value specified
''' by the read_tmo_ms command.
''' </summary>
''' <para name="primaryAddress">     [Optional, Integer] Specifies the primary GPIB address
'''                                   between 0 and 30. </param>
''' <para name="secondaryAddress)>   [Optional, Integer] Specifies the secondard GPIB Address
'''                                  between 0 and 30. </param>
''' <returns>   The status byte. </returns>
Public Function SerialPoll(Optional ByVal primaryAddress As Integer = -1, _
                           Optional ByVal secondaryAddress As Integer = -1) As Integer
    
    If Me.ViSession.UsingPrologix Then
        If primaryAddress >= 0 And secondaryAddress < 0 Then
            SeriaPoll = CInt(Me.ViSession.QueryLine("++spoll " & CStr(primaryAddress)))
        ElseIf primaryAddress >= 0 And secondaryAddress >= 0 Then
            SeriaPoll = CInt(Me.ViSession.QueryLine("++spoll " & CStr(primaryAddress) & _
                        " " & CStr(secondaryAddress + 96)))
        Else
            SeriaPoll = CInt(Me.ViSession.QueryLine("++spoll"))
        End If
    Else
        SeriaPoll = 0
    End If
    
End Function

''' <summary>   Query Prologix service request status. </summary>
''' <returns>   True if Server requested. </returns>
Friend Function IsSRQ() As Boolean
    
    If Me.ViSession.UsingPrologix Then
        IsSRQ = 1 = CInt(Me.ViSession.QueryLine("++srq"))
    End If
    
End Function


' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Connection Changed

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal isConnected As Boolean)
    RaiseEvent ConnectionChanged( _
            cc_isr_Winsock.Constructor.CreateConnectionChangedEventArgs(isConnected))
End Sub


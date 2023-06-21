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

''' <summary> Initializes an IEEE488 Session. </summary>
Private Sub Class_Initialize()
    
    Set m_viSession = New ViSession
    
End Sub

''' <summary> Termiantes an IEEE488 Session. </summary>
Private Sub Class_Terminate()
    
    Me.Dispose
    
End Sub

''' <summary>   Closes and releases the VI Session. </summary>
Public Sub Dispose()
    
    If Not m_viSession Is Nothing Then
        m_viSession.Dispose
    End If
    Set m_viSession = Nothing

End Sub

''' <summary>   Returns the connection state of the client. </summary>
''' <vaule>   True if the TCP Client socket is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
    
    If m_viSession Is Nothing Then
        Connected = False
    Else
        Connected = m_viSession.Connected
    End If

End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean
    
    CanConnect = Not Me.Connected

End Property

''' <summary>   Opens a TCP Client connection to the specified host at the specified port
''' as sets the connection timeout. </summary>
''' <param name="host">        An IPv4 dotted-decimal host address. </param>
''' <param name="port">        The port that the server is listening on. </param>
''' <param name="timeoutMs">   [Optional, 500] The receive timeout in milliseconds. </param>
''' <returns>   True if the TCP Client is connected. </returns>
Public Function OpenConnection(ByVal host As String, ByVal port As Long, Optional ByVal timeoutMs As Long) As Boolean
   
    OpenConnection = m_viSession.OpenConnection(host, port, timeoutMs)
   
End Function

''' <summary>   Closes the TCP Client connection. </summary>
''' <returns>   True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    CloseConnection = m_viSession.CloseConnection
    
End Function

''' <summary>   Returns a descriptor identifying the socket. </summary>
''' <value>   The socket identifying descriptor. </value>
Public Property Get SocketId() As Long
    If Me.Connected Then
        SocketId = m_viSession.SocketId
    Else
        SocketId = -1
    End If
End Property

''' <summary>   Sends a message. </summary>
''' <param name="message">   The message. </param>
''' <returns>   The number of sent characters. </returns>
Public Function WriteLine(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    
    WriteLine = m_viSession.WriteLine(message, appendTermination)

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <returns>   The received string. </returns>
Public Function Read(Optional ByVal maxLength As Long = &H7FFF, _
                     Optional ByVal trimEnd As Boolean = True) As String
    
    Read = m_viSession.Read(maxLength, trimEnd)

End Function
    
''' <summary>   Sends a message and receives a reply. </summary>
''' <param name="message">   The message. </param>
''' <returns>   The received string. </returns>
Public Function QueryLine(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    
    QueryLine = m_viSession.QueryLine(message, appendTermination)

End Function

''' <summary>   Returns a reference to the VI Session. </summary>
''' <value>   A ViSession. </returns>
Public Property Get ViSession() As ViSession
    Set ViSession = m_viSession
End Property

''' <summary> Clears Status (CLS) command. </summary>
''' <returns> the number of characters that were sent. </returns>
Public Function ClearExecutionState() As Integer
    ClearExecutionState = Me.ViSession.WriteLine(Ieee488Syntax.ClearExecutionStateCommand)
End Function

''' <summary>   Returns the identity. </summary>
''' <value>   A string. </returns>
Public Property Get Identity() As String
    If StringExtensions.IsNullOrEmpty(m_identity) And Me.ViSession.Connected Then
        m_identity = Me.QueryIdentity
    End If
    Identity = m_identity
End Property

''' <summary>   Returns the instrument identity using hte *IDN? query command. </summary>
''' <returns>   A string. </returns>
Public Function QueryIdentity() As String
    m_identity = Me.ViSession.QueryLine(Ieee488Syntax.IdentityQueryCommand)
    QueryIdentity = m_identity
End Function

''' <summary> Operation Complete (*OPC) command. </summary>
''' <returns> The number of characters that were sent. </returns>
Public Function OperationComplete() As Integer
    OperationComplete = Me.ViSession.WriteLine(Ieee488Syntax.OperationCompleteCommand)
End Function

''' <summary>   Returns 1 if operation was coompleted; otherwize 0. </summary>
''' <returns>   A string. 1 if complete; otehrwise 0. </returns>
Public Function QueryOperationCompleted() As String
    QueryOperationCompleted = Me.ViSession.QueryLine(Ieee488Syntax.OperationCompletedQueryCommand)
End Function

''' <summary>   Returns option (instrument specific). </summary>
''' <returns>   A string. </returns>
Public Function QueryOptions() As String
    
    QueryOptions = Me.ViSession.QueryLine(Ieee488Syntax.OptionsQueryCommand)
    
End Function

''' <summary> Issues a local lockout (++llo). </summary>
''' <remarks>   Valid if <see cref="UsingPrologix"/>. </remarks>
''' <returns>   The number of characters that were sent. </returns>
Public Function LocalLockout() As Integer
    
    If Me.ViSession.UsingPrologix Then
        LocalLockout = Me.WriteLine("++llo")
    Else
        LocalLockout = 0
    End If
    
End Function

''' <summary> Issues a Go To Local (GTL) (++loc). </summary>
''' <remarks>   Valid if <see cref="UsingPrologix"/>. </remarks>
''' <returns>   The number of characters that were sent. </returns>
Public Function GoToLocal() As Integer
    
    If Me.ViSession.UsingPrologix Then
        GoToLocal = Me.WriteLine("++loc")
    Else
        GoToLocal = 0
    End If
    
End Function

''' <summary> Issues an SDC. </summary>
''' <remarks>   Valid if <see cref="UsingPrologix"/>. </remarks>
''' <returns>   The number of characters that were sent. </returns>
Public Function SelectiveDeviceClear() As Integer
    
    If Me.ViSession.UsingPrologix Then
        SelectiveDeviceClear = Me.WriteLine("++clr")
    Else
        SelectiveDeviceClear = 0
    End If
    
End Function

''' <summary> Reads the status byte. </summary>
''' <param name="canQuery">   [Optional, False] True to send <c>*STB?</c>
''' if not <see cref="UsingPrologix"/>. </param>
''' <returns>   The status byte. </returns>
Public Function ReadStatusByte(Optional ByVal canQuery As Boolean = False) As Integer
    
    If Me.ViSession.UsingPrologix Then
        ReadStatusByte = CInt(Me.QueryLine("++spoll"))
    Else
        ' querying the service request status could cause
        ' a query unterminated error.
        ReadStatusByte = 0
    End If
    
End Function

''' <summary> Query service request status. </summary>
''' <param name="canQuery">   [Optional, False] True to send <c>*STB?</c>
''' if not <see cref="UsingPrologix"/>. </param>
''' <returns>   The status byte. </returns>
Public Function SRQ(Optional ByVal canQuery As Boolean = False) As Boolean
    
    If Me.ViSession.UsingPrologix Then
        SRQ = 1 = CInt(Me.QueryLine("++srq"))
    Else
        SRQ = ServiceRequests.RequestingService = (ServiceRequests.RequestingService And Me.QueryServiceRequestStatus)
    End If
    
End Function

''' <summary> Issues a wait (*WAI) command. </summary>
''' <returns>   The number of characters that were sent. </returns>
Public Function Wait() As Integer
    
    Wait = Me.ViSession.WriteLine(Ieee488Syntax.WaitCommand)
    
End Function

''' <summary> Enables standard events using the (*ESE {0}) command. </summary>
''' <param name="standardEvents">   The stadnard event bits to enable. <param>
''' <returns>   The number of characters that were sent. </returns>
Public Function EnableStandardEvents(ByVal StandardEvents As Integer) As Integer
    
    Dim command As String
    command = StringExtensions.StringFormat(StandardEventEnableCommandFormat, StandardEvents)
    EnableStandardEvents = Me.ViSession.WriteLine(command)
    
End Function

''' <summary>   Returns the standard events enable byte using the *ESE? query command. </summary>
''' <returns>   The standard event register enable status value. </returns>
Public Function QueryStandardEventsEnable() As Integer
    
    Dim StandardEvents As String
    StandardEvents = Me.ViSession.QueryLine(Ieee488Syntax.StandardEventEnableQueryCommand)
    QueryStandardEventsEnable = CInt(StandardEvents)

End Function

''' <summary>   Returns the standard events status byte using the *ESR? query command. </summary>
''' <returns>   The standard event register status value. </returns>
Public Function QueryStandardEventsStatus() As Integer
    
    Dim StandardEvents As String
    StandardEvents = Me.ViSession.QueryLine(Ieee488Syntax.StandardEventStatusQueryCommand)
    QueryStandardEventsStatus = CInt(StandardEvents)
    
End Function

''' <summary>   Enables Service Request using the (*SRE {0}) command. </summary>
''' <param name="ServiceRequest">   The stadnard event bits to enable. <param>
''' <returns>   The number of characters that were sent. </returns>
Public Function EnableServiceRequest(ByVal serviceRequest As Integer) As Integer
    
    Dim command As String
    command = StringExtensions.StringFormat(ServiceRequestEnableCommandFormat, serviceRequest)
    EnableServiceRequest = Me.ViSession.WriteLine(command)
    
End Function

''' <summary>   Returns the Service Request enable byte using the *SRE? query command. </summary>
''' <returns>   The service request event register enable status value. </returns>
Public Function QueryServiceRequestEnable() As Integer
    
    Dim serviceRequest As String
    serviceRequest = Me.ViSession.QueryLine(Ieee488Syntax.ServiceRequestEnableQueryCommand)
    QueryServiceRequestEnable = CInt(serviceRequest)
    
End Function

''' <summary>   Returns the Service Request status byte using the *STB? query command. </summary>
''' <returns>   The service request event register status value. </returns>
Public Function QueryServiceRequestStatus() As Integer
    
    Dim serviceRequest As String
    serviceRequest = Me.ViSession.QueryLine(Ieee488Syntax.ServiceRequestQueryCommand)
    QueryServiceRequestStatus = CInt(serviceRequest)
    
End Function

''' <summary> Reset to known state (*RST) command. </summary>
''' <returns>   The number of characters that were sent. </returns>
Public Function ResetKnownState() As Integer
    
    ResetKnownState = Me.ViSession.WriteLine(Ieee488Syntax.ResetKnownStateCommand)
    
End Function


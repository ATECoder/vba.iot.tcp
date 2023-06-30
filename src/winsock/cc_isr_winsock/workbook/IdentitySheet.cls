VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdentitySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const PROLOGIX_PORT_NO As Integer = 1234

Private Const m_repeatCountCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_receiveTimeoutCell As String = "B4"
Private Const m_errorMessageCell As String = "B5"

Private Const m_socketAddressCell As String = "C2"
Private Const m_connectedSocketIdCell As String = "D2"
Private Const m_connectedCell As String = "E2"
Private Const m_sentMessageCell As String = "F2"
Private Const m_receivedMessageLengthCell As String = "G2"
Private Const m_receivedMessageCell As String = "H2"
Private Const m_statusByteCell As String = "F3"
Private Const m_averageElapsedTimeCell As String = "H3"
Private Const m_elapsedTimeCell As String = "H4"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_identityCell As String = "H2"
Private m_delayStopper As cc_isr_Core.StopWatch

''' <summary>   Reads the identity form the instrument with the
''' IP address and port defined in the parent sheet. </summary>
Private Sub ReadIdentityButton_Click()

    Dim host As String
    Dim port As Integer
    Dim repeatCount As Integer
    Dim timeout As Integer
    
    Set m_delayStopper = cc_isr_Core.Constructor.CreateStopWatch
    timeout = Range(m_receiveTimeoutCell).value
    repeatCount = Range(m_repeatCountCell).value
    host = Range(m_inputHostCell).value
    port = Range(m_inputPortCell).value
    
    If Me.UsingWinsockReadRawButton.value Then
        ' read identity using basic winsock.
        ReadIdentity host, port, repeatCount, timeout
    ElseIf Me.UsingTcpClientOptionButton Then
        ' read identity using the Tcp Client.
        TcpClientReadIdentity host, port, repeatCount, timeout
    ElseIf Me.UseWinsockWithBufferOptionButton Then
        ' read identity using winsock with buffer.
        ReadIdentityUsingBuffer host, port, repeatCount, timeout
    End If
    
End Sub

''' <summary>   Sends the identity query and reads back the identity from the instrument. </summary>
''' <param name="host">   The host IPv4 dotted address. </param>
''' <param name="port">   the host port number. </param>
''' <param name="repeatCount">   The number of time to repeat the query. </param>
Private Sub ReadIdentity(ByVal host As String, ByVal port As Long, ByVal repeatCount As Integer, ByVal timeout As Integer)

    On Error GoTo Finally
    
    Range(m_socketAddressCell) = ""
    Range(m_connectedSocketIdCell).value = ""
    Range(m_connectedCell).value = ""
    Range(m_sentMessageCell).value = ""
    Range(m_receivedMessageLengthCell).value = ""
    Range(m_receivedMessageCell).value = ""
    Range(m_elapsedTimeCell).value = ""
    Range(m_errorMessageCell).value = ""

    Dim maximumLength As Integer
    maximumLength = 1024
    Dim buffer As String * 1024
    
    Dim sentCount As Integer
    Dim readCount As Integer
    
    Dim stopper As StopWatch: Set stopper = cc_isr_Core.Constructor.CreateStopWatch
    Dim socket As New IPv4StreamSocket
    Dim totalMilliseconds As Double
    
    ' report the socket number
    Range(m_connectedSocketIdCell).value = socket.SocketId
    
    socket.OpenConnection host, port
    
    ' report the connection state
    Range(m_connectedCell).value = socket.Connected
    
    If repeatCount > 0 And socket.Connected Then
    
        Range(m_socketAddressCell).value = socket.SocketAddress
    
        Dim command As String
    
        If port = PROLOGIX_PORT_NO Then
        
            ' set auto read after write
            ' Prologix GPIB-ETHERNET controller can be configured to automatically address
            ' instruments to talk after sending them a command in order to read their response. The
            ' feature called, Read-After-Write, saves the user from having to issue read commands
            ' repeatedly.
            command = "++auto 1"

            ' send the command, which may cause Query Unterminated because we are setting the device to talk
            ' where there is nothing to talk.
            sentCount = socket.SendMessage(command & vbLf)
            m_delayStopper.Wait 5

            ' disables front panel operation of the currently addressed instrument.
            
            sentCount = socket.SendMessage("++llo" & vbLf)
            m_delayStopper.Wait 5

            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            sentCount = socket.SendMessage("*CLS; *OPC?" & vbLf)
            socket.ReceiveRaw
            m_delayStopper.Wait 5

        End If
        
        command = "*IDN?"
        Range(m_sentMessageCell).value = command
       
        Dim loopCount As Integer
        Do While loopCount < repeatCount
            
            DoEvents
            
            loopCount = loopCount + 1
            stopper.Restart
            
            ' send the command
            sentCount = socket.SendMessage(command & vbLf)
            m_delayStopper.Wait 5
            
            ' receive the reading
            Dim Identity As String
            Identity = socket.ReceiveRaw()
    
            totalMilliseconds = totalMilliseconds + stopper.ElapsedMilliseconds
            
            Range(m_receivedMessageLengthCell).value = Len(Identity)
            Range(m_receivedMessageCell).value = Identity
    
        Loop
  
        Range(m_averageElapsedTimeCell).value = Format(totalMilliseconds / loopCount, m_elapsedTimeFormat) + " ms"
        Range(m_elapsedTimeCell).value = Format(totalMilliseconds, m_elapsedTimeFormat) + " ms"
    
    ElseIf repeatCount <= 0 Then
    
        Range(m_receivedMessageCell).value = "testing connect and disconnect; disconnected."
        
    Else
        Range(m_errorMessageCell).value = "connection failed without reporting an exception"
    End If
    
Finally:

    Dim errorMessage As String
    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            errorMessage = "Reading identity failed " & UserDefinedErrors.FormattedErr & _
                           " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
        End If
    End If
    
    On Error Resume Next
    If Not socket Is Nothing Then
    
        If socket.Connected Then
        
            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            sentCount = socket.SendMessage("*CLS; *OPC?" & vbLf)
            socket.ReceiveRaw
            m_delayStopper.Wait 5
            
            ' enables front panel operation of the currently addressed instrument.
            
            sentCount = socket.SendMessage("++loc" & vbLf)
            m_delayStopper.Wait 5
            
            socket.CloseConnection
        End If
    
    End If
    
    ' terminate the socket, which should release the Winsock dll.
    Set socket = Nothing
    
    If Len(errorMessage) > 0 Then
        Range(m_errorMessageCell).value = errorMessage
    End If

End Sub


''' <summary>   Sends the identity query and reads back the identity from the instrument. </summary>
''' <param name="host">   The host IPv4 dotted address. </param>
''' <param name="port">   the host port number. </param>
''' <param name="repeatCount">   The number of time to repeat the query. </param>
Private Sub ReadIdentityUsingBuffer(ByVal host As String, ByVal port As Long, ByVal repeatCount As Integer, ByVal timeout As Integer)

    On Error GoTo Finally
    
    Range(m_socketAddressCell) = ""
    Range(m_connectedSocketIdCell).value = ""
    Range(m_connectedCell).value = ""
    Range(m_sentMessageCell).value = ""
    Range(m_receivedMessageLengthCell).value = ""
    Range(m_receivedMessageCell).value = ""
    Range(m_elapsedTimeCell).value = ""
    Range(m_errorMessageCell).value = ""

    Dim maximumLength As Integer
    maximumLength = 1024
    Dim buffer As String * 1024
    
    Dim sentCount As Integer
    Dim readCount As Integer
    
    Dim stopper As StopWatch: Set stopper = cc_isr_Core.Constructor.CreateStopWatch
    Dim socket As New IPv4StreamSocket
    Dim totalMilliseconds As Double
    
    ' report the socket number
    Range(m_connectedSocketIdCell).value = socket.SocketId
    
    socket.OpenConnection host, port
    
    ' report the connection state
    Range(m_connectedCell).value = socket.Connected
    
    If repeatCount > 0 And socket.Connected Then
    
        ' display the socket address
        Range(m_socketAddressCell).value = socket.SocketAddress
        
        Dim command As String
    
        If port = PROLOGIX_PORT_NO Then
        
            ' set auto read after write
            ' Prologix GPIB-ETHERNET controller can be configured to automatically address
            ' instruments to talk after sending them a command in order to read their response. The
            ' feature called, Read-After-Write, saves the user from having to issue read commands
            ' repeatedly.
            
            command = "++auto 1"

            ' send the command, which may cause Query Unterminated because we are setting the device to talk
            ' where there is nothing to talk.
            
            sentCount = socket.SendMessage(command & vbLf)
            m_delayStopper.Wait 5

            ' disables front panel operation of the currently addressed instrument.
            
            sentCount = socket.SendMessage("++llo" & vbLf)
            m_delayStopper.Wait 5

            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            sentCount = socket.SendMessage("*CLS; *OPC?" & vbLf)
            socket.ReceiveRaw
            m_delayStopper.Wait 5

        End If
        
        command = "*IDN?"
        Range(m_sentMessageCell).value = command
        
        Dim loopCount As Integer
        Do While loopCount < repeatCount
            
            DoEvents
            
            loopCount = loopCount + 1
            stopper.Restart
            
            ' send the command
            socket.SendMessage command & vbLf
        
            ' receive the reading
            readCount = socket.ReceiveTerminatedMessage(buffer, maximumLength, vbLf)
    
            totalMilliseconds = totalMilliseconds + stopper.ElapsedMilliseconds
            
            Range(m_receivedMessageLengthCell).value = readCount
            Range(m_receivedMessageCell).value = buffer
    
        Loop
  
        Range(m_averageElapsedTimeCell).value = Format(totalMilliseconds / loopCount, m_elapsedTimeFormat) + " ms"
        Range(m_elapsedTimeCell).value = Format(totalMilliseconds, m_elapsedTimeFormat) + " ms"
    
    ElseIf repeatCount <= 0 Then
    
        Range(m_receivedMessageCell).value = "testing connect and disconnect; disconnected."
        
    Else
        Range(m_errorMessageCell).value = "connection failed without reporting an exception"
    End If
    
Finally:

    Dim errorMessage As String
    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            errorMessage = "Reading identity failed " & UserDefinedErrors.FormattedErr & _
                           " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
        End If
    End If
    
    On Error Resume Next
    
    If Not socket Is Nothing Then
    
        If socket.Connected Then
        
            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            sentCount = socket.SendMessage("*CLS; *OPC?" & vbLf)
            socket.ReceiveRaw
            m_delayStopper.Wait 5
            
            ' enables front panel operation of the currently addressed instrument.
            
            sentCount = socket.SendMessage("++loc" & vbLf)
            m_delayStopper.Wait 5
            
            socket.CloseConnection
        End If
    
    End If
    
    ' terminate the socket, which should release the Winsock dll.
    Set socket = Nothing
    
    If Len(errorMessage) > 0 Then
        Range(m_errorMessageCell).value = errorMessage
    End If

End Sub

''' <summary>   Uses the Tcp Client to send the identity query and reads back the identity
''' from the instrument. </summary>
''' <param name="host">   The host IPv4 dotted address. </param>
''' <param name="port">   the host port number. </param>
''' <param name="repeatCount">   The number of time to repeat the query. </param>
Private Sub TcpClientReadIdentity(ByVal host As String, ByVal port As Long, ByVal repeatCount As Integer, ByVal timeout As Integer)

    On Error GoTo Finally
    
    Range(m_socketAddressCell) = ""
    Range(m_connectedSocketIdCell).value = ""
    Range(m_connectedCell).value = ""
    Range(m_sentMessageCell).value = ""
    Range(m_receivedMessageLengthCell).value = ""
    Range(m_receivedMessageCell).value = ""
    Range(m_elapsedTimeCell).value = ""
    Range(m_errorMessageCell).value = ""

    Dim stopper As StopWatch: Set stopper = cc_isr_Core.Constructor.CreateStopWatch
    Dim totalMilliseconds As Double
    
    Dim sentCount As Integer
    Dim client As TcpClient
    Set client = New TcpClient
    
    client.ConnectTo host, port
    
    ' report the socket number
    Range(m_connectedSocketIdCell).value = client.TcpClientSocket.SocketId
   
    ' report the connection state
    Range(m_connectedCell).value = client.Connected
    
    If repeatCount > 0 And client.Connected Then
    
        ' display the socket address.
        Range(m_socketAddressCell).value = client.TcpClientSocket.SocketAddress
        
        Dim command As String
    
        If port = PROLOGIX_PORT_NO Then
        
            ' set auto read after write
            ' Prologix GPIB-ETHERNET controller can be configured to automatically address
            ' instruments to talk after sending them a command in order to read their response. The
            ' feature called, Read-After-Write, saves the user from having to issue read commands
            ' repeatedly.
            
            command = "++auto 1"

            ' send the command, which may cause Query Unterminated because we are setting the device to talk
            ' where there is nothing to talk.
            
            sentCount = client.SendMessage(command & vbLf)
            m_delayStopper.Wait 5

            
            ' disables front panel operation of the currently addressed instrument.
            
            sentCount = client.SendMessage("++llo" & vbLf)
            m_delayStopper.Wait 5

            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            sentCount = client.SendMessage("*CLS; *OPC?" & vbLf)
            client.ReceiveRaw
            m_delayStopper.Wait 5
            

        End If
        
        command = "*IDN?"
        Range(m_sentMessageCell).value = command
        
        Dim loopCount As Integer
        Do While loopCount < repeatCount
            
            DoEvents
            
            loopCount = loopCount + 1
            stopper.Restart
            
            ' send the command
            client.SendMessage command & vbLf
            m_delayStopper.Wait 5
        
            ' receive the reading
            Dim Identity As String
            Identity = client.ReceiveRaw()
    
            totalMilliseconds = totalMilliseconds + stopper.ElapsedMilliseconds
            
            Range(m_receivedMessageLengthCell).value = Len(Identity)
            Range(m_receivedMessageCell).value = Identity
    
        Loop
  
        Range(m_averageElapsedTimeCell).value = Format(totalMilliseconds / loopCount, m_elapsedTimeFormat) + " ms"
        Range(m_elapsedTimeCell).value = Format(totalMilliseconds, m_elapsedTimeFormat) + " ms"
    
    ElseIf repeatCount <= 0 Then
    
        Range(m_receivedMessageCell).value = "testing connect and disconnect; disconnected."
        
    Else
        Range(m_errorMessageCell).value = "connection failed without reporting an exception"
    End If
    
Finally:

    Dim errorMessage As String
    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            errorMessage = "Reading identity failed using Tcp Client " & UserDefinedErrors.FormattedErr & _
                           " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
        End If
    End If
    
    On Error Resume Next
    If Not client Is Nothing Then
        If client.Connected Then
            
            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            sentCount = client.SendMessage("*CLS; *OPC?" & vbLf)
            client.ReceiveRaw
            m_delayStopper.Wait 5
            
            ' enables front panel operation of the currently addressed instrument.
            
            sentCount = client.SendMessage("++loc" & vbLf)
            m_delayStopper.Wait 5

            
            client.Disconnect
        End If
    End If
    
    ' terminate the client
    Set client = Nothing
    
    If Len(errorMessage) > 0 Then
        Range(m_errorMessageCell).value = errorMessage
    End If

End Sub



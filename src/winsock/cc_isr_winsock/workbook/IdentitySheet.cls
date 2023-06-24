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

''' <summary>   Reads the identity form the instrument with the
''' IP address and port defined in the parent sheet. </summary>
Private Sub ReadIdentityButton_Click()

    Dim host As String
    Dim port As Integer
    Dim repeatCount As Integer
    
    repeatCount = Range("B1").value
    host = Range("B2").value
    port = Range("B3").value
    
    Range("C2").value = host
    Range("D2").value = port
    
#If True Then
    ' read identity using basic winsock.
    ReadIdentity host, port, repeatCount
#ElseIf False Then
    ' read identity using the Tcp Client.
    TcpClientReadIdentity host, port, repeatCount
#Else
    ' read identity using winsock with buffer.
    ReadIdentityUsingBuffer host, port, repeatCount
#End If
    
End Sub

''' <summary>   Sends the identity query and reads back the identity from the instrument. </summary>
''' <param name="host">   The host IPv4 dotted address. </param>
''' <param name="port">   the host port number. </param>
''' <param name="repeatCount">   The number of time to repeat the query. </param>
Private Sub ReadIdentity(ByVal host As String, ByVal port As Long, ByVal repeatCount As Integer)

    On Error GoTo Finally
    
    Range("E2").value = ""
    Range("F2").value = ""
    Range("G2").value = ""
    Range("H2").value = ""
    Range("I2").value = ""

    Dim maximumLength As Integer
    maximumLength = 1024
    Dim buffer As String * 1024
    
    Dim readCount As Integer
    
    Dim sw As StopWatch: Set sw = cc_isr_Core.Constructor.CreateStopWatch
    Dim socket As New IPv4StreamSocket
    Dim totalMilliseconds As Double
    
    ' report the socket number
    Range("E2").value = socket.SocketId
    
    socket.OpenConnection host, port
    
    ' report the connection state
    Range("F2").value = socket.Connected
    
    If socket.Connected Then
    
        Dim command As String
        command = "*IDN?"
        Range("G2").value = command
        
        Dim loopCount As Integer
        Do While loopCount < repeatCount
            
            DoEvents
            
            loopCount = loopCount + 1
            sw.Restart
            
            ' send the command
            socket.SendMessage command & vbLf
        
            ' receive the reading
            Dim Identity As String
            Identity = socket.ReceiveRaw()
    
            totalMilliseconds = totalMilliseconds + sw.ElapsedMilliseconds
            
            Range("H2").value = Len(Identity)
            Range("I2").value = Identity
    
        Loop
  
        ' command = ":SYST:PRES"
        ' count = SendMessage(command)
        ' Call OPC
        
        Range("I3").value = Format(totalMilliseconds / loopCount, "0.0") + " ms"
        Range("I4").value = Format(totalMilliseconds, "0.0") + " ms"
    
    Else
        MsgBox "connection failed without reporting an exception"
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
    If socket.Connected Then
        socket.CloseConnection
    End If
    
    ' terminate the socket, which should release the Winsock dll.
    Set socket = Nothing
    
    If Len(errorMessage) > 0 Then
        MsgBox errorMessage
    End If

End Sub


''' <summary>   Sends the identity query and reads back the identity from the instrument. </summary>
''' <param name="host">   The host IPv4 dotted address. </param>
''' <param name="port">   the host port number. </param>
''' <param name="repeatCount">   The number of time to repeat the query. </param>
Private Sub ReadIdentityUsingBuffer(ByVal host As String, ByVal port As Long, ByVal repeatCount As Integer)

    On Error GoTo Finally
    
    Range("E2").value = ""
    Range("F2").value = ""
    Range("G2").value = ""
    Range("H2").value = ""
    Range("I2").value = ""

    Dim maximumLength As Integer
    maximumLength = 1024
    Dim buffer As String * 1024
    
    Dim readCount As Integer
    
    Dim sw As StopWatch: Set sw = cc_isr_Core.Constructor.CreateStopWatch
    Dim socket As New IPv4StreamSocket
    Dim totalMilliseconds As Double
    
    ' report the socket number
    Range("E2").value = socket.SocketId
    
    socket.OpenConnection host, port
    
    ' report the connection state
    Range("F2").value = socket.Connected
    
    If socket.Connected Then
    
        Dim command As String
        command = "*IDN?"
        Range("G2").value = command
        
        Dim loopCount As Integer
        Do While loopCount < repeatCount
            
            DoEvents
            
            loopCount = loopCount + 1
            sw.Restart
            
            ' send the command
            socket.SendMessage command & vbLf
        
            ' receive the reading
            readCount = socket.ReceiveTerminatedMessage(buffer, maximumLength, vbLf)
    
            totalMilliseconds = totalMilliseconds + sw.ElapsedMilliseconds
            
            Range("H2").value = readCount
            Range("I2").value = buffer
    
        Loop
  
        ' command = ":SYST:PRES"
        ' count = SendMessage(command)
        ' Call OPC
        
        Range("I3").value = Format(totalMilliseconds / loopCount, "0.0") + " ms"
        Range("I4").value = Format(totalMilliseconds, "0.0") + " ms"
    
    Else
        MsgBox "connection failed without reporting an exception"
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
    If socket.Connected Then
        socket.CloseConnection
    End If
    
    ' terminate the socket, which should release the Winsock dll.
    Set socket = Nothing
    
    If Len(errorMessage) > 0 Then
        MsgBox errorMessage
    End If

End Sub

''' <summary>   Uses the Tcp Client to send the identity query and reads back the identity
''' from the instrument. </summary>
''' <param name="host">   The host IPv4 dotted address. </param>
''' <param name="port">   the host port number. </param>
''' <param name="repeatCount">   The number of time to repeat the query. </param>
Private Sub TcpClientReadIdentity(ByVal host As String, ByVal port As Long, ByVal repeatCount As Integer)

    On Error GoTo Finally
    
    Range("E2").value = ""
    Range("F2").value = ""
    Range("G2").value = ""
    Range("H2").value = ""
    Range("I2").value = ""

    Dim sw As StopWatch: Set sw = cc_isr_Core.Constructor.CreateStopWatch
    Dim totalMilliseconds As Double
    
    Dim client As TcpClient
    Set client = New TcpClient
    
    client.ConnectTo host, port
    
    ' report the socket number
    Range("E2").value = client.TcpClientSocket.SocketId
   
    ' report the connection state
    Range("F2").value = client.TcpClientSocket.Connected
    
    If client.TcpClientSocket.Connected Then
    
        Dim command As String
        command = "*IDN?"
        Range("G2").value = command
        
        Dim loopCount As Integer
        Do While loopCount < repeatCount
            
            DoEvents
            
            loopCount = loopCount + 1
            sw.Restart
            
            ' send the command
            client.SendMessage command & vbLf
        
            ' receive the reading
            Dim Identity As String
            Identity = client.ReceiveRaw()
    
            totalMilliseconds = totalMilliseconds + sw.ElapsedMilliseconds
            
            Range("H2").value = Len(Identity)
            Range("I2").value = Identity
    
        Loop
  
        ' command = ":SYST:PRES"
        ' count = SendMessage(command)
        ' Call OPC
        
        Range("I3").value = Format(totalMilliseconds / loopCount, "0.0") + " ms"
        Range("I4").value = Format(totalMilliseconds, "0.0") + " ms"
    
    Else
        MsgBox "connection failed without reporting an exception"
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
            client.Disconnect
        End If
    End If
    
    ' terminate the client
    Set client = Nothing
    
    If Len(errorMessage) > 0 Then
        MsgBox errorMessage
    End If

End Sub



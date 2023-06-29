VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_tcpClientSocket As IPv4StreamSocket

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>] event args. </param>
Public Event ConnectionChanged(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before changing the connection. </summary>
''' <param name="eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>] event args. </param>
Public Event ConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Handles the Class Initialize event; constructs the <see cref="TcpClientSocket"/>. </summary>
Private Sub Class_Initialize()
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="TcpClientSocket"/>. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Close the connection and releases the socket. </summary>
Public Sub Dispose()

    On Error Resume Next
    
    Dim disconnected As Boolean: disconnected = Not Me.Connected
    
    If Not disconnected Then disconnected = Me.Disconnect()
    
    If disconnected And Not m_tcpClientSocket Is Nothing Then Set m_tcpClientSocket = Nothing
    
End Sub


''' <summary>   Gets the Tcp Client socket. </summary>
''' <value>   [<see cref="IPv4StreamSocket"/>. </value>
Public Property Get TcpClientSocket() As IPv4StreamSocket
    Set TcpClientSocket = m_tcpClientSocket
End Property

''' <summary>   Opens a socket connection and sets and returns the socket id. </summary>
''' <param name="host">   An IPv4 dotted-decimal host address. </param>
''' <param name="port">   The port that the server is listening on. </param>
''' <returns>   True if the socket is connected. </returns>
Public Function ConnectTo(ByVal host As String, ByVal port As Long) As Boolean
    
    If host = "localhost" Then host = "127.0.0.1"
    
    Dim eventArgs As New cc_isr_Winsock.ConnectionChangingEventArgs
    eventArgs.Initialize Me.Connected
    OnConnectionChanging eventArgs
    
    If Not eventArgs.Cancel Then
    
        ' a new socket must be created on each connection because closing a socket
        ' releases the socket, which can then be used by the system as soon as it is released.
        Set m_tcpClientSocket = New IPv4StreamSocket
    
        Me.TcpClientSocket.OpenConnection host, port
        
        OnConnectionChanged Me.Connected
    
    End If
    
    ConnectTo = Me.Connected
    
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the socket is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    If m_tcpClientSocket Is Nothing Then
        Connected = False
    Else
        Connected = Me.TcpClientSocket.Connected
    End If
End Property

''' <summary>   Closes the connection. </summary>
''' <returns>   True if disconnection succeeded; otherwise, false. </returns>
Public Function Disconnect() As Boolean

    Dim eventArgs As New cc_isr_Winsock.ConnectionChangingEventArgs
    eventArgs.Initialize Me.Connected
    OnConnectionChanging eventArgs
    
    If eventArgs.Cancel Then
    
        Me.TcpClientSocket.CloseConnection
        
        ' the socket must be disposed here becuase it cannot be used after
        ' the connection is closed.
        Set m_tcpClientSocket = Nothing
        
        OnConnectionChanged Me.Connected
            
    End If
    
    Disconnect = Not Me.Connected
    
End Function

''' <summary>   Initializes this TCP Client using the provided connected socket. </summary>
''' <remarks>   Raises an invalid operation exception is a socket is already connected. </remarks>
''' <param name="connectedSocket ">   A connected INET stream socket. </param>
Public Sub Initialize(ByVal connectedSocket As IPv4StreamSocket)
    
    If Me.Connected Then
        Err.Raise UserDefinedErrors.InvalidOperationError.Code, "cc_isr_Winsock.TcpClient.WinsockApi.Initialize", _
                  UserDefinedErrors.InvalidOperationError.Description & _
                  " This Tc Client is connected to " & Me.TcpClientSocket.SocketAddress & _
                  " A new Tcp Client must be used when initialized using a connected socket at " & connectedSocket.SocketAddress & "."
    End If
    
    ' remove this socket
    If Not m_tcpClientSocket Is Nothing Then Set m_tcpClientSocket = Nothing
    
    Set m_tcpClientSocket = connectedSocket
End Sub

''' <summary>   Sends a terminated message to the server. </summary>
''' <param name="command">   The message. </param>
''' <returns>   The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal message As String) As Integer
    SendMessage = Me.TcpClientSocket.SendMessage(message)
End Function

''' <summary>   Receives a message from the server. </summary>
''' <param name="maxLength">   The maximum number of characters to read. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveBytes(ByVal maxLength As Long) As String
    
    ReceiveBytes = Me.TcpClientSocket.ReceiveBytes(maxLength)

End Function

''' <summary>   Receives a message from the server. </summary>
''' <remarks>   Data is read until exhausted. </remarks>
''' <param name="bufferSize">   [Optional, 1024] The number of character to read on each trial. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveString(Optional ByVal bufferSize As Integer = 1024) As String
    
    ReceiveString = Me.TcpClientSocket.ReceiveString(bufferSize)

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <param name="termination">   [Optional, 10 ASCII] The character the signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function ReceiveRaw(Optional ByVal maxLength As Long = &H7FFF, _
                           Optional ByVal trimEnd As Boolean = True, _
                           Optional ByVal termination As String = vbLf) As String
    
    ReceiveRaw = Me.TcpClientSocket.ReceiveRaw(maxLength, trimEnd, termination)
  
End Function


''' <summary>   Sends a query message and reads back from the server until reaching the specified
''' termination or reading the specified number of characters. </summary>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <param name="termination">   [Optional, 10 ASCII] The character the signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function QueryRaw(ByVal message As String, _
                         Optional ByVal maxLength As Long = &H7FFF, _
                         Optional ByVal trimEnd As Boolean = True, _
                         Optional ByVal termination As String = vbLf) As String
    
    If Me.SendMessage(message) > 0 Then
        QueryRaw = Me.TcpClientSocket.ReceiveRaw(maxLength, trimEnd, termination)
    Else
        QueryRaw = ""
    End If
  
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Connection events
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





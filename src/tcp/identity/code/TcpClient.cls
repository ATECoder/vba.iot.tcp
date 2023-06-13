VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_tcpClientSocket As IPv4StreamSocket

''' <summary> Gets the Tcp Client socket. </summary>
''' <value> A IPv4StreamSocket. </value>
Public Property Get TcpClientSocket() As IPv4StreamSocket
    Set TcpClientSocket = m_tcpClientSocket
End Property

''' <summary> Opens a socket connection and sets and returns the socket id. </summary>
''' <param name="host"> An IPv4 dotted-decimal host address. </param>
''' <param name="port"> The port that the server is listening on. </param>
''' <returns> True if the socket is connected. </returns>
Public Function ConnectTo(host As String, port As Long) As Boolean
    
    If host = "localhost" Then
        host = "127.0.0.1"
    End If

    Set m_tcpClientSocket = New IPv4StreamSocket

    m_tcpClientSocket.OpenConnection host, port
    
End Function

''' <summary> Initializes this TCP Client using the provided connected socket. </summary>
''' <param name="connectedSocket "> A connected Inte stream socket. </param>
Public Sub Initialize(ByVal connectedSocket As IPv4StreamSocket)
    Set m_tcpClientSocket = connectedSocket
End Sub

''' <summary> Sends a terminated message to the server. </summary>
''' <param name="command"> The message. </param>
''' <returns> The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal message As String) As Integer
    SendMessage = Me.TcpClientSocket.SendMessage(message)
End Function

''' <summary> Receives a message from the server. </summary>
''' <param name="maxLength"> The maximum number of maxLength to read. </param>
''' <returns> The received characters. </returns>
Public Function ReceiveBytes(ByVal maxLength As Long) As String
    ReceiveBytes = Me.TcpClientSocket.ReceiveBytes(maxLength)
End Function

Public Function ReceiveString() As String
    Dim buffer As String
    Dim message As String
    message = ""
    
    Do
        buffer = Trim(ReceiveBytes(1024))

        If Len(buffer) > 0 Then
            message = message & buffer
        End If
    Loop While Len(buffer) > 0
    
    ReceiveString = Trim(message)
End Function


''' <summary> Closese and releases the socket. </summary>
Public Sub Dispose()
    If Not m_tcpClientSocket Is Nothing And m_tcpClientSocket.Connected Then
        m_tcpClientSocket.CloseConnection
    End If
    m_tcpClientSocket = Nothing
End Sub


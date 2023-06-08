VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const COMMAND_ERROR = -1
Private Const RECV_ERROR = -1
Private Const NO_ERROR = 0
Private socketId As Long

''' <summary> Opens a socket connection and sets and returns the socket id. </summary>
''' <param name=""> </param>
''' <param name=""> </param>
''' <returns> The socket identity. </returns>
Function OpenConnection(ByVal host As String, ByVal port As Integer) As Integer
   
    If host = "localhost" Then
        host = "127.0.0.1"
    End If
    
    ' see that the singleton windock class is instantiated
    On Error Resume Next
    
    Dim Winsock As Winsock
    Set Winsock = New Winsock
    
    On Error GoTo 0

    ' Create a new socket
    
    socketId = Socket(wsock32.AF_INET, wsock32.SOCK_STREAM, 0)
    If socketId < 0 Then
        MsgBox ("ERROR: open socket = " + Str$(socketId))
        OpenSocket = COMMAND_ERROR
        Exit Function
    End If

    ' Open a connection to a server
    
    Dim address As wsock32.sockaddr_in
    address.sin_addr.s_addr = wsock32.inet_addr(host)
    address.sin_family = wsock32.AF_INET
    address.sin_port = wsock32.htons(port)
    
    Dim connectResult As Long
    connectResult = wsock32.connect(socketId, address, Len(address))
    If connectResult < 0 Then
        MsgBox ("ERROR: connection failed = " + Str$(connectResult))
        OpenSocket = COMMAND_ERROR
        Exit Function
    End If
    
    OpenSocket = socketId

End Function

''' <summary> Closes the socket connection and zeros the socket id. </summary>
Function CloseConnection() As Long

    Dim result As Long
    result = closesocket(socketId)
    socketId = 0
    If result < 0 Then
        MsgBox ("ERROR: closing connection = " + Str$(result))
    End If
    CloseConnection = result

End Function


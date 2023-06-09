Attribute VB_Name = "Framework"
Public Const COMMAND_ERROR = -1
Public Const RECV_ERROR = -1
Public Const NO_ERROR = 0

Private Const sheet = "IdentitySheet"
Private Const versionCell = "B1"

Public SocketId As Long

Sub CloseConnection()

    Dim result As Long
    result = closesocket(SocketId)
    
    If result < 0 Then
        MsgBox ("ERROR: closing connection = " + Str$(result))
        Exit Sub
    End If

End Sub

Sub EndIt()

    ' Shutdown Winsock DLL
    x = Winsock.Cleanup()

End Sub

Function StartIt() As Boolean

    ' Initialize Winsock DLL
    x = Winsock.Initialize()
    
    ' Dim startUpInfo As WSADATA
    
    ' Version 1.1 (1*256 + 1) = 257
    ' version 2.0 (2*256 + 0) = 512
    
    ' Get WinSock version
    ' Sheets(sheet).Select
    ' Range(versionCell).Select
    ' Version = ActiveCell.FormulaR1C1
    ' x = wsock32.WSAStartup(Version, startUpInfo)
    
    If x <> 0 Then
        MsgBox ("ERROR starting winsock")
    End If
    StartIt = (x = 0)

End Function
 
Function OpenSocket(ByVal host As String, ByVal port As Integer) As Integer
   
    If host = "localhost" Then
        host = "127.0.0.1"
    End If

    ' Create a new socket
    
    SocketId = wsock32.CreateSocket(AF_INET, SOCK_STREAM, 0)
    If SocketId < 0 Then
        MsgBox ("ERROR: open socket = " + Str$(SocketId))
        OpenSocket = Framework.COMMAND_ERROR
        Exit Function
    End If

    ' Open a connection to a server
    
    Dim address As wsock32.sockaddr_in
    address.sin_addr.s_addr = wsock32.inet_addr(host)
    address.sin_family = wsock32.AF_INET
    address.sin_port = wsock32.htons(port)
    
    Dim connectResult As Long
    connectResult = wsock32.connect(SocketId, address, Len(address))
    If connectResult < 0 Then
        MsgBox ("ERROR: connection failed = " + Str$(connectResult))
        OpenSocket = Framework.COMMAND_ERROR
        Exit Function
    End If
    
    OpenSocket = SocketId

End Function

Function SendCommand(ByVal command As String) As Integer

    Dim strSend As String
    
    strSend = command + vbCrLf
    
    count = send(SocketId, ByVal strSend, Len(strSend), 0)
    
    If count < 0 Then
        MsgBox ("ERROR: sending command = " + Str$(count))
        SendCommand = COMMAND_ERROR
        Exit Function
    End If
    
    SendCommand = count

End Function

Function Receive(dataBuf As String, ByVal maxLength As Integer) As Integer

    Dim c As String * 1
    Dim length As Integer
    
    dataBuf = ""
    While length < maxLength
        DoEvents
        
        c = ""
        Dim l As Long
        l = Len(c)
        count = recv(SocketId, c, l, 0)
        
        If count < 1 Then
            Receive = RECV_ERROR
            dataBuf = Chr$(0)
            Exit Function
        End If
        
        If c = Chr$(10) Then
           dataBuf = dataBuf + Chr$(0)
           Receive = length
           Exit Function
        End If
        
        length = length + count
        dataBuf = dataBuf + c
    Wend
    
    Receive = RECV_ERROR
    
End Function


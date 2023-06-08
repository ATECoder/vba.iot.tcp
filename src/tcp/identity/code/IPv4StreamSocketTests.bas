Attribute VB_Name = "IPv4StreamSocketTests"

''' <summary> Tests creating a socket. </summary>
''' <returns> An instance of the <see cref="Assert"/> class. </returns>
Public Function TestCreateSocket() As Assert

    Dim sock As IPv4StreamSocket
    Set sock = New IPv4StreamSocket
    
    ' check if socket has a valid id
    Set TestCreateSocket = Assert.IsTrue(sock.SocketId <> wsock32.INVALID_SOCKET, _
        "Failed creating socket; socket id " & Str$(sock.SocketId) & _
        " must not equal to wsock32.INVALID_SOCKET=" & wsock32.INVALID_SOCKET)
    
    If Not TestCreateSocket.AssertSuccessful Then
        Set sock = Nothing
        Exit Function
    End If
    
    Set TestCreateSocket = Assert.IsTrue(Winsock.Initiated, "Winsock should be initiated when a socket is created")
    
    If Not TestCreateSocket.AssertSuccessful Then
        Set sock = Nothing
        Exit Function
    End If
    
    Set TestCreateSocket = Assert.IsFalse(Winsock.Disposed, "Winsock should not be disposed when a socket is created")
    
    If Not TestCreateSocket.AssertSuccessful Then
        Set sock = Nothing
        Exit Function
    End If
    
    Set TestCreateSocket = Assert.AreEqual(Winsock.SocketCount, 1, _
        "Winsock socket count should be 1 after registering a single socket but is " & Str$(Winsock.SocketCount))
    
    If Not TestCreateSocket.AssertSuccessful Then
        Set sock = Nothing
        Exit Function
    End If

    ' test terminating the socket, which should dispose of the winsock class.
    Set sock = Nothing
    
    Set TestCreateSocket = Assert.AreEqual(Winsock.SocketCount, 0, _
        "Winsock socket count should be 0 after nulling single socket but is " & Str$(Winsock.SocketCount))
    
    If Not TestCreateSocket.AssertSuccessful Then
        Set sock = Nothing
        Exit Function
    End If
    
    Set TestCreateSocket = Assert.IsFalse(Winsock.Initiated, "Winsock should no longer be initiated after the last socket was set to nothing")
    
    If Not TestCreateSocket.AssertSuccessful Then
        Set sock = Nothing
        Exit Function
    End If
    
    Set TestCreateSocket = Assert.IsTrue(Winsock.Disposed, "Winsock should be disposed after the last socket was set to nothing")
    
    If Not TestCreateSocket.AssertSuccessful Then
        Set sock = Nothing
        Exit Function
    End If
    
End Function





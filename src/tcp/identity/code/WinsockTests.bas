Attribute VB_Name = "WinsockTests"
''' <summary> Unit test. Asserts instatiating and disposing of the Winsock framework. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestInitializeAndDispose() As Assert

    ' this is required to initialize Winsock.  It will only ran once.
    Winsock.Initialize
    
    Set TestInitializeAndDispose = Assert.IsTrue(Winsock.Initiated, "Winsock should be initiated when a socket is created")
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If
    
    Set TestInitializeAndDispose = Assert.IsFalse(Winsock.Disposed, "Winsock should not be disposed when a socket is created")
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If
    
    Set TestInitializeAndDispose = Assert.AreEqual(Winsock.SocketCount, 0, _
        "Winsock socket count should be 0 as no sockets are registered but is " & Str$(Winsock.SocketCount))
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If

    ' test disposing of Winsock.
    Winsock.Dispose
    
    Set TestInitializeAndDispose = Assert.IsFalse(Winsock.Initiated, "Winsock should no longer be initiated after the last socket was set to nothing")
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If
    
    Set TestInitializeAndDispose = Assert.IsTrue(Winsock.Disposed, "Winsock should be disposed after the last socket was set to nothing")
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If
    
    Winsock.Dispose
    
End Function







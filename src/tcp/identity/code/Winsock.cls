VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Winsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const COMMAND_ERROR = -1
Private Const RECV_ERROR = -1
Private Const NO_ERROR = 0

' true if winsock was initiated
Private m_initiated As Boolean

' true if winsock was disposed
Private m_disposed As Boolean

' counts the number of sockets
Private m_socketCount As Integer

''' <summary> Initiates use of the Winsock DLL by a process. </summary>
''' <param name="was"> A pointer to the WSADATA data structure that is to receive
''' details of the Windows Sockets implementation. </param>
''' <returns> If successful, the WSAStartup function returns zero. Otherwise, it returns one of
''' the error codes listed below. The WSAStartup function directly returns the extended error code
''' in the return value for this function. A call to the WSAGetLastError function is not needed and should not be used.
''' </returns>
Private Declare PtrSafe Function WSAStartup Lib "wsock32.dll" (ByVal versionRequired As Long, wsa As WSADATA) As Long

Private Declare PtrSafe Function WSAGetLastError Lib "wsock32.dll" () As Long

''' <summary> terminates use of the Winsock dll. </summary>
''' <remarks> In a multithreaded environment, WSACleanup terminates Windows Sockets operations
'''   for all threads. </remarks>
''' <param name=""> </param>
''' <returns>
'''   The return value is zero if the operation was successful. Otherwise, the value
'''   SOCKET_ERROR is returned, and a specific error number can be retrieved by calling WSAGetLastError.
''' <returns>
Private Declare PtrSafe Function WSACleanup Lib "wsock32.dll" () As Long

''' <summary> Initializes winsock. </summary>
'''< remarks> This is a singleton class.
''' <see href="https://stackoverflow.com/questions/396117/class-static-methods-in-vba"/>
''' <see href="https://web.archive.org/web/20160418020100/http://www.utteraccess.com/wiki/Index.php/Singleton"/>
''' This method is call on the first method (not property) access of this class.
''' </remarks>
Private Sub Class_Initialize()
    Debug.Print "initializing winsock"
End Sub

''' <summary> Initializes winsock. </summary>
'''< remarks> This is a singleton class.
''' <see href="https://stackoverflow.com/questions/396117/class-static-methods-in-vba"/>
''' <see href="https://web.archive.org/web/20160418020100/http://www.utteraccess.com/wiki/Index.php/Singleton"/>
''' </remarks>
''' <returns> If successful, returns zero. Otherwise, it returns the extended error code
''' in the return value for this function.
''' </returns>
Public Function Initialize()

    Dim result As Long
    result = 0

    ' skip if the class was already initialized.
    If m_initiated Then
        Initialize = result
        Exit Function
    End If
    
    Dim startUpInfo As WSADATA
    
    ' Initialize Winsock DLL
    
    result = WSAStartup(wsock32.WINSOCK_VERSION, startUpInfo)
    
    If result <> 0 Then
        err.Raise ErrorCodes.WinsockStartingError.Code, "IdentityReader.Winsock.Cleanup", _
                  ErrorCodes.WinsockStartingError.Description & " Last error: " & CStr(result)
    End If
    m_initiated = (result = 0)
    m_disposed = False
    m_socketCount = 0
    Initialize = result

End Function

''' <summary> Registers this socket with the Winsock class. </summary>
''' <remarks> Call this method to register the socket with Winsock thus managing when
''' Winsock can be disposed.</remarks>
''' <param name="activeSocket"> The socket that is registring. </param>
Public Sub Register(activeSocket As IPv4StreamSocket)
    If IsObject(activeSocket) Then
        m_socketCount = m_socketCount + 1
    End If
End Sub

''' <summary> Un-Registers this socket from the Winsock class. </summary>
''' <remarks> Call this method to remove the socket from the class so it can be disposed when having
''' no registered sockets.</remarks>
''' <param name="activeSocket"> The socket that is unregistring. </param>
Public Sub UnRegister(activeSocket As IPv4StreamSocket)
    If IsObject(activeSocket) And (Me.SocketCount > 0) Then
        m_socketCount = m_socketCount - 1
        If Me.SocketCount = 0 Then
            Me.Dispose
        End If
    End If
End Sub

''' <summary> Returns true if Winsock was initiated. </summary>
Public Property Get Initiated() As Boolean
    Initiated = m_initiated
End Property

''' <summary> Returns true if Winsock was Disposed. </summary>
Public Property Get Disposed() As Boolean
    Disposed = m_disposed
End Property

''' <summary> returns the last Winsock 32 error . </summary>
Public Property Get LastError() As Long
    LastError = WSAGetLastError()
End Property

''' <summary> Terminates use of the Winsock dll. </summary>
Public Function Cleanup()
    ' Shutdown Winsock DLL
    Dim result As Long
    result = WSACleanup()
    If result = wsock32.SOCKET_ERROR Then
        err.Raise ErrorCodes.WinsockCleanupError.Code, "IdentityReader.Winsock.Cleanup", _
                  ErrorCodes.WinsockCleanupError.Description & " Last error: " & CStr(Winsock.LastError)
    End If
    Cleanup = result
End Function

''' <summary> Returns the number of sockets associated with this Winsock instance. </summary>
Public Property Get SocketCount() As Integer
    SocketCount = m_socketCount
End Property

''' <summary> Terminates Winsock if all sockets released their association with this instance. </summary>
Public Sub Dispose()
    On Error Resume Next
    If m_initiated And Not m_disposed And (m_socketCount = 0) Then
        m_disposed = True
        m_initiated = False
        ' Shutdown Winsock DLL
        Dim result As Long
        result = WSACleanup()
        If result <> 0 Then
            err.Raise ErrorCodes.WinsockCleanupError.Code, "IdentityReader.Winsock.Dispose", _
                      ErrorCodes.WinsockCleanupError.Description & " Last error: " & CStr(Winsock.LastError)
        End If
    End If
End Sub


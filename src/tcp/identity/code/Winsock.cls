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
Public Sub Initialize()

    ' skip if the class was already initialized.
    If m_initiated Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Dim startUpInfo As WSADATA
    
    ' Initialize Winsock DLL
    Dim result As Long
    result = wsock32.WSAStartup(wsock32.WINSOCK_VERSION, startUpInfo)
    
    If result <> 0 Then
        MsgBox ("ERROR starting winsock " & CStr(result))
    End If
    m_initiated = (result = 0)
    m_disposed = False
    m_socketCount = 0
    
    Exit Sub
    
ErrorHandler:
   MsgBox ("ERROR starting winsock class (" & CStr(err.Number) & "): " & err.Description)

End Sub

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
        result = wsock32.WSACleanup()
        If result <> 0 Then
            err.Raise vbObjectError + 513, "IdentityReader.Winsock.Dispose", _
                      "ERROR disposing winsock. Winsock.WSACleanup() returned " & CStr(result)
        End If
    End If
End Sub


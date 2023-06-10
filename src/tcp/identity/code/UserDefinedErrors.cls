VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_errorsDix As Dictionary
Private m_winsockStartingError As UserDefinedError
Private m_winsockCleanupError As UserDefinedError
Private m_socketCreationError As UserDefinedError
Private m_socketConnectionError As UserDefinedError
Private m_socketDisconnectionError As UserDefinedError
Private m_socketOptionSettingError As UserDefinedError
Private m_socketSendError As UserDefinedError
Private m_socketReceiveError As UserDefinedError

Private Sub Class_Initialize()

    Set m_errorsDix = New Dictionary
    Dim errorCode As Long
    Dim dummyError As UserDefinedError
    Set dummyError = New UserDefinedError
    errorCode = vbObjectError + dummyError.UserErrorNumberBase
    
    errorCode = errorCode + 1
    Set m_winsockStartingError = New UserDefinedError
    m_winsockStartingError.InitiateProperties errorCode, "WinsockStartingError", _
        "Failed starting Winsock using winsock WSAStartup()."
    AddUserDefinedError Me.WinsockStartingError
    
    errorCode = errorCode + 1
    Set m_winsockCleanupError = New UserDefinedError
    m_winsockCleanupError.InitiateProperties errorCode, "WinsockCleanupError", _
        "Failed disposing Winsock using winsock WSACleanup()."
    AddUserDefinedError Me.WinsockCleanupError
    
    errorCode = errorCode + 1
    Set m_socketCreationError = New UserDefinedError
    m_socketCreationError.InitiateProperties errorCode, "SocketCreationError", _
        "Failed creating socket using winsock socket()."
    AddUserDefinedError Me.SocketCreationError
    
    errorCode = errorCode + 1
    Set m_socketConnectionError = New UserDefinedError
    m_socketConnectionError.InitiateProperties errorCode, "SsocketConnectionError", _
        "Failed connecting socket using winsock connect()."
    AddUserDefinedError Me.SocketConnectionError
    
    errorCode = errorCode + 1
    Set m_socketDisconnectionError = New UserDefinedError
    m_socketDisconnectionError.InitiateProperties errorCode, "SocketDisconnectionError", _
        "Failed disconnecting socket using winsock closesocket()."
    AddUserDefinedError Me.SocketDisconnectionError
    
    errorCode = errorCode + 1
    Set m_socketOptionSettingError = New UserDefinedError
    m_socketOptionSettingError.InitiateProperties errorCode, "SocketOptionSettingError", _
        "Failed setting socket options using winsock setsockopt()."
    AddUserDefinedError Me.SocketOptionSettingError
    
    errorCode = errorCode + 1
    Set m_socketSendError = New UserDefinedError
    m_socketSendError.InitiateProperties errorCode, "SocketSendError", _
        "Failed send a socket message using winsock send()."
    AddUserDefinedError Me.SocketSendError

    errorCode = errorCode + 1
    Set m_socketReceiveError = New UserDefinedError
    m_socketReceiveError.InitiateProperties errorCode, "SocketReceiveError", _
        "Failed send a socket message using winsock send()."
    AddUserDefinedError Me.SocketReceiveError

End Sub

''' <summary> Initializes the error codes and descriptions. </summary>
Public Sub Initialize()
    Class_Initialize
End Sub

''' <summary> Adds a user defined error to the error dictionary. </summary>
''' <para name"userError"> The error of type <see cref="UserDefinedError"/> </param>
Private Sub AddUserDefinedError(userError As UserDefinedError)
    m_errorsDix.Add userError.Code, userError
End Sub

''' <summary> Returns true if the specified error exists. </summary>
''' <para name"userError"> The error of type <see cref="UserDefinedError"/> </param>
Public Function UserDefinedErrorExists(ByVal userError As UserDefinedError) As Boolean
    UserDefinedErrorExists = m_errorsDix.Exists(userError.Code)
End Function

''' <summary> Returns true if the specified error exists. </summary>
''' <para name"userErrorCode"> The error code </param>
Public Function UserDefinedErrorCodeExists(ByVal userErrorCode As Long) As Boolean
    UserDefinedErrorCodeExists = m_errorsDix.Exists(userErrorCode)
End Function

''' <summary> Returns a user defined error. </summary>
''' <param name="key"> The user defined error code. </param>
Public Function GetUserDefinedError(ByVal key As Long) As UserDefinedError
    If m_errorsDix.Exists(key) Then
        Set GetUserDefinedError = m_errorsDix.Item(key)
    Else
        Set GetUserDefinedError = New UserDefinedError
        GetUserDefinedError.InitiateProperties key, "UnknwonUserDefinedError", "Unknown error."
    End If
End Function

Public Property Get WinsockStartingError() As UserDefinedError
    Set WinsockStartingError = m_winsockStartingError
End Property

Public Property Get WinsockCleanupError() As UserDefinedError
    Set WinsockCleanupError = m_winsockCleanupError
End Property

Public Property Get SocketCreationError() As UserDefinedError
    Set SocketCreationError = m_socketCreationError
End Property

Public Property Get SocketConnectionError() As UserDefinedError
    Set SocketConnectionError = m_socketConnectionError
End Property

Public Property Get SocketDisconnectionError() As UserDefinedError
    Set SocketDisconnectionError = m_socketDisconnectionError
End Property

Public Property Get SocketOptionSettingError() As UserDefinedError
    Set SocketOptionSettingError = m_socketOptionSettingError
End Property

Public Property Get SocketSendError() As UserDefinedError
    Set SocketSendError = m_socketSendError
End Property

Public Property Get SocketReceiveError() As UserDefinedError
    Set SocketReceiveError = m_socketReceiveError
End Property


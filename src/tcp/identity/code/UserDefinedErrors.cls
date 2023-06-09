VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_userErrors As Collection
Private m_winsockStartingError As UserError
Private m_winsockCleanupError As UserError
Private m_socketCreationError As UserError
Private m_socketConnectionError As UserError
Private m_socketDisconnectionError As UserError
Private m_socketOptionSettingError As UserError
Private m_socketSendError As UserError
Private m_socketReceiveError As UserError

Private Sub Class_Initialize()

    Set m_userErrors = New Collection
    Dim baseNumber As Integer
    baseNumber = 512
    
    baseNumber = baseNumber + 1
    Set m_winsockStartingError = New UserError
    m_winsockStartingError.InitiateProperties vbObjectError + baseNumber, _
        "Failed starting Winsock using winsock WSAStartup()."
    m_userErrors.Add m_winsockStartingError, baseNumber
    
    baseNumber = baseNumber + 1
    Set m_winsockCleanupError = New UserError
    m_winsockCleanupError.InitiateProperties vbObjectError + baseNumber, _
        "Failed disposing Winsock using winsock WSACleanup()."
    m_userErrors.Add m_winsockCleanupError, baseNumber
    
    baseNumber = baseNumber + 1
    Set m_socketCreationError = New UserError
    m_socketCreationError.InitiateProperties vbObjectError + baseNumber, _
        "Failed creating socket using winsock CreateSocket() alias socket."
    m_userErrors.Add m_socketCreationError, baseNumber
    
    baseNumber = baseNumber + 1
    Set m_socketConnectionError = New UserError
    m_socketConnectionError.InitiateProperties vbObjectError + baseNumber, _
        "Failed connecting socket using winsock connect()."
    m_userErrors.Add m_socketConnectionError, baseNumber
    
    baseNumber = baseNumber + 1
    Set m_socketDisconnectionError = New UserError
    m_socketDisconnectionError.InitiateProperties vbObjectError + baseNumber, _
        "Failed disconnecting socket using winsock closesocket()."
    m_userErrors.Add m_socketDisconnectionError, baseNumber
    
    baseNumber = baseNumber + 1
    Set m_socketOptionSettingError = New UserError
    m_socketOptionSettingError.InitiateProperties vbObjectError + baseNumber, _
        "Failed setting socket options using winsock setsockopt()."
    m_userErrors.Add m_socketOptionSettingError, baseNumber
    
    baseNumber = baseNumber + 1
    Set m_socketSendError = New UserError
    m_socketSendError.InitiateProperties vbObjectError + baseNumber, _
        "Failed send a socket message using winsock send()."
    m_userErrors.Add m_socketSendError, baseNumber

    baseNumber = baseNumber + 1
    Set m_socketReceiveError = New UserError
    m_socketReceiveError.InitiateProperties vbObjectError + baseNumber, _
        "Failed send a socket message using winsock send()."
    m_userErrors.Add m_socketReceiveError, baseNumber

End Sub

''' <summary> Initializes the error codes and descriptions. </summary>
Public Sub Initialize()
End Sub

''' <summary> Returns a user error. </summary>
''' <param name="key"> The user error code. </param>
Public Function GetUserError(ByVal key As Object) As UserError
    If CollectionExtensions.ContainsKey(m_userErrors, key) Then
        GetUserError = m_userErrors.Item(key)
    Else
        Set GetUserError = New UserError
        GetUserError.InitiateProperties key, "Unknown error."
    End If
End Function

Public Property Get WinsockStartingError() As UserError
    WinsockStartingError = m_winsockStartingError
End Property

Public Property Get WinsockCleanupError() As UserError
    WinsockCleanupError = m_winsockCleanupError
End Property

Public Property Get SocketCreationError() As UserError
    SocketCreationError = m_socketCreationError
End Property

Public Property Get SocketConnectionError() As UserError
    SocketConnectionError = m_socketConnectionError
End Property

Public Property Get SocketDisconnectionError() As UserError
    FailedDisconnecting = m_socketDisconnectionError
End Property

Public Property Get SocketOptionSettingError() As UserError
    SocketOptionSettingError = m_socketOptionSettingError
End Property

Public Property Get SocketSendError() As UserError
    SocketSendError = m_socketSendError
End Property

Public Property Get SocketReceiveError() As UserError
    SocketReceiveError = m_socketReceiveError
End Property



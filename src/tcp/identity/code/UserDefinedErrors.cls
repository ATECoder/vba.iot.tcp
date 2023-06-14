VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_errorsDix As Dictionary
Private m_winsockStartingError As UserDefinedError
Private m_winsockCleanupError As UserDefinedError
Private m_socketCreationError As UserDefinedError
Private m_socketConnectionError As UserDefinedError
Private m_socketDisconnectionError As UserDefinedError
Private m_socketOptionSettingError As UserDefinedError
Private m_socketSendError As UserDefinedError
Private m_socketReceiveError As UserDefinedError
Private m_socketBindError As UserDefinedError
Private m_socketListenError As UserDefinedError
Private m_socketAcceptError As UserDefinedError
Private m_socketSetStatusError As UserDefinedError
Private m_invalidArgumentError As UserDefinedError

''' <summary>   Initializes the user defined errors. </summary>
Private Sub Class_Initialize()

    Set m_errorsDix = New Dictionary
    Dim errorCode As Long
    Dim dummyError As UserDefinedError
    Set dummyError = New UserDefinedError
    errorCode = VBA.vbObjectError + dummyError.UserErrorNumberBase
    
    errorCode = errorCode + 1
    Set m_winsockStartingError = New UserDefinedError
    m_winsockStartingError.InitiateProperties errorCode, "WinsockStartingError", _
        "Failed starting Winsock using Winsock WSAStartup()."
    AddUserDefinedError Me.WinsockStartingError
    
    errorCode = errorCode + 1
    Set m_winsockCleanupError = New UserDefinedError
    m_winsockCleanupError.InitiateProperties errorCode, "WinsockCleanupError", _
        "Failed disposing Winsock using Winsock WSACleanup()."
    AddUserDefinedError Me.WinsockCleanupError
    
    errorCode = errorCode + 1
    Set m_socketCreationError = New UserDefinedError
    m_socketCreationError.InitiateProperties errorCode, "SocketCreationError", _
        "Failed creating socket using Winsock socket()."
    AddUserDefinedError Me.SocketCreationError
    
    errorCode = errorCode + 1
    Set m_socketConnectionError = New UserDefinedError
    m_socketConnectionError.InitiateProperties errorCode, "SsocketConnectionError", _
        "Failed connecting socket using Winsock connect()."
    AddUserDefinedError Me.SocketConnectionError
    
    errorCode = errorCode + 1
    Set m_socketDisconnectionError = New UserDefinedError
    m_socketDisconnectionError.InitiateProperties errorCode, "SocketDisconnectionError", _
        "Failed disconnecting socket using Winsock closesocket()."
    AddUserDefinedError Me.SocketDisconnectionError
    
    errorCode = errorCode + 1
    Set m_socketOptionSettingError = New UserDefinedError
    m_socketOptionSettingError.InitiateProperties errorCode, "SocketOptionSettingError", _
        "Failed setting socket options using Winsock setsockopt()."
    AddUserDefinedError Me.SocketOptionSettingError
    
    errorCode = errorCode + 1
    Set m_socketSendError = New UserDefinedError
    m_socketSendError.InitiateProperties errorCode, "SocketSendError", _
        "Failed send a socket message using Winsock send()."
    AddUserDefinedError Me.SocketSendError

    errorCode = errorCode + 1
    Set m_socketReceiveError = New UserDefinedError
    m_socketReceiveError.InitiateProperties errorCode, "SocketReceiveError", _
        "Failed send a socket message using Winsock send()."
    AddUserDefinedError Me.SocketReceiveError

    errorCode = errorCode + 1
    Set m_socketBindError = New UserDefinedError
    m_socketBindError.InitiateProperties errorCode, "SocketBindError", _
        "Failed send a socket message using Winsock send()."
    AddUserDefinedError Me.SocketBindError

    errorCode = errorCode + 1
    Set m_socketListenError = New UserDefinedError
    m_socketListenError.InitiateProperties errorCode, "SocketListenError", _
        "Failed send a socket message using Winsock send()."
    AddUserDefinedError Me.SocketListenError

    errorCode = errorCode + 1
    Set m_socketAcceptError = New UserDefinedError
    m_socketAcceptError.InitiateProperties errorCode, "SocketAcceptError", _
        "Failed send a socket message using Winsock send()."
    AddUserDefinedError Me.SocketAcceptError

    errorCode = errorCode + 1
    Set m_socketSetStatusError = New UserDefinedError
    m_socketSetStatusError.InitiateProperties errorCode, "SocketSetStatusError", _
        "Failed send a socket message using Winsock send()."
    AddUserDefinedError Me.SocketSetStatusError

    errorCode = errorCode + 1
    Set m_invalidArgumentError = New UserDefinedError
    m_invalidArgumentError.InitiateProperties errorCode, "InvalidArgumentError", _
        "Invalid argument error."
    AddUserDefinedError Me.InvalidArgumentError

End Sub

''' <summary>   Initializes the error codes and descriptions. </summary>
Public Sub Initialize()
    Class_Initialize
End Sub

''' <summary>   Builds a an error message from the VBA <see cref="Err"/>. </summary>
Public Function FormattedErr() As String
    
    Dim ErrNumberDetails As String
    ErrNumberDetails = " (" & (Err.Number - VBA.vbObjectError) & " / 0x" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")"
    FormattedErr = "[Error Number=" & Err.Number & ErrNumberDetails & ", Source=" & Err.source & ", Description=" & Err.Description & "]"

End Function

''' <summary>   Adds a user defined error to the error dictionary. </summary>
''' <para name"userError">   The error of type <see cref="UserDefinedError"/>   </param>
Private Sub AddUserDefinedError(ByRef userError As UserDefinedError)
    m_errorsDix.Add userError.Code, userError
End Sub

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"userError">   The error of type <see cref="UserDefinedError"/>   </param>
Public Function UserDefinedErrorExists(ByVal userError As UserDefinedError) As Boolean
    UserDefinedErrorExists = m_errorsDix.Exists(userError.Code)
End Function

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"userErrorCode">   The error code </param>
Public Function UserDefinedErrorCodeExists(ByVal userErrorCode As Long) As Boolean
    UserDefinedErrorCodeExists = m_errorsDix.Exists(userErrorCode)
End Function

''' <summary>   Returns a user defined error. </summary>
''' <param name="key">   The user defined error code. </param>
Public Function GetUserDefinedError(ByVal key As Long) As UserDefinedError
    If m_errorsDix.Exists(key) Then
        Set GetUserDefinedError = m_errorsDix.Item(key)
    Else
        Set GetUserDefinedError = New UserDefinedError
        GetUserDefinedError.InitiateProperties key, "UnknwonUserDefinedError", "Unknown error."
    End If
End Function

''' <summary>   Get the error that occurred starting Winsock. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get WinsockStartingError() As UserDefinedError
    Set WinsockStartingError = m_winsockStartingError
End Property

''' <summary>   Get the error that occurred ending Winsock. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get WinsockCleanupError() As UserDefinedError
    Set WinsockCleanupError = m_winsockCleanupError
End Property

''' <summary>   Get the error that occurred creating a socket. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketCreationError() As UserDefinedError
    Set SocketCreationError = m_socketCreationError
End Property

''' <summary>   Get the error that occurred connecting a socket. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketConnectionError() As UserDefinedError
    Set SocketConnectionError = m_socketConnectionError
End Property

''' <summary>   Get the error that occurred disconnecting a socket. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketDisconnectionError() As UserDefinedError
    Set SocketDisconnectionError = m_socketDisconnectionError
End Property

''' <summary>   Get the error that occurred setting a socket option. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketOptionSettingError() As UserDefinedError
    Set SocketOptionSettingError = m_socketOptionSettingError
End Property

''' <summary>   Get the error that occurred sending a socket message. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketSendError() As UserDefinedError
    Set SocketSendError = m_socketSendError
End Property

''' <summary>   Get the error that occurred receiving a socket message. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketReceiveError() As UserDefinedError
    Set SocketReceiveError = m_socketReceiveError
End Property

''' <summary>   Get the error that occurred binding a socket to an address. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketBindError() As UserDefinedError
    Set SocketBindError = m_socketBindError
End Property

''' <summary>   Get the error that occurred starting to listen on a socket. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketListenError() As UserDefinedError
    Set SocketListenError = m_socketListenError
End Property

''' <summary>   Get the error that occurred accepting an incoming connection. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketAcceptError() As UserDefinedError
    Set SocketAcceptError = m_socketAcceptError
End Property

''' <summary>   Get the error that occurred determining the status of a socket set. </summary>
''' <values>   Returns a user defined error. </value>
Public Property Get SocketSetStatusError() As UserDefinedError
    Set SocketSetStatusError = m_socketSetStatusError
End Property

Public Property Get InvalidArgumentError() As UserDefinedError
    Set InvalidArgumentError = m_invalidArgumentError
End Property


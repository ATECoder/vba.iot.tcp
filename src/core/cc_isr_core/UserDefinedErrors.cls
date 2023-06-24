VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
Private m_invalidOperationError As UserDefinedError


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
    m_socketConnectionError.InitiateProperties errorCode, "SocketConnectionError", _
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
        "Failed receiving a socket message using Winsock recv()."
    AddUserDefinedError Me.SocketReceiveError

    errorCode = errorCode + 1
    Set m_socketBindError = New UserDefinedError
    m_socketBindError.InitiateProperties errorCode, "SocketBindError", _
        "Failed binding a socket using Winsock bind()."
    AddUserDefinedError Me.SocketBindError

    errorCode = errorCode + 1
    Set m_socketListenError = New UserDefinedError
    m_socketListenError.InitiateProperties errorCode, "SocketListenError", _
        "Failed stating a socket listening using Winsock listen()."
    AddUserDefinedError Me.SocketListenError

    errorCode = errorCode + 1
    Set m_socketAcceptError = New UserDefinedError
    m_socketAcceptError.InitiateProperties errorCode, "SocketAcceptError", _
        "Failed permitting an incoming connection attempt on a socket using Winsock accept()."
    AddUserDefinedError Me.SocketAcceptError

    errorCode = errorCode + 1
    Set m_socketSetStatusError = New UserDefinedError
    m_socketSetStatusError.InitiateProperties errorCode, "SocketSetStatusError", _
        "Failed determining the socket status using Winsock select()."
    AddUserDefinedError Me.SocketSetStatusError

    errorCode = errorCode + 1
    Set m_invalidArgumentError = New UserDefinedError
    m_invalidArgumentError.InitiateProperties errorCode, "InvalidArgumentError", _
        "Invalid argument error."
    AddUserDefinedError Me.InvalidArgumentError

    errorCode = errorCode + 1
    Set m_invalidOperationError = New UserDefinedError
    m_invalidOperationError.InitiateProperties errorCode, "InvalidOperationError", _
        "Invalid operation error."
    AddUserDefinedError Me.InvalidOperationError

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
        Set GetUserDefinedError = m_errorsDix.item(key)
    Else
        Set GetUserDefinedError = New UserDefinedError
        GetUserDefinedError.InitiateProperties key, "UnknwonUserDefinedError", "Unknown error."
    End If
End Function

''' <summary>   Get the error that occurred starting Winsock. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get WinsockStartingError() As UserDefinedError
    Set WinsockStartingError = m_winsockStartingError
End Property

''' <summary>   Get the error that occurred ending Winsock. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get WinsockCleanupError() As UserDefinedError
    Set WinsockCleanupError = m_winsockCleanupError
End Property

''' <summary>   Get the error that occurred creating a socket. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketCreationError() As UserDefinedError
    Set SocketCreationError = m_socketCreationError
End Property

''' <summary>   Get the error that occurred connecting a socket. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketConnectionError() As UserDefinedError
    Set SocketConnectionError = m_socketConnectionError
End Property

''' <summary>   Get the error that occurred disconnecting a socket. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketDisconnectionError() As UserDefinedError
    Set SocketDisconnectionError = m_socketDisconnectionError
End Property

''' <summary>   Get the error that occurred setting a socket option. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketOptionSettingError() As UserDefinedError
    Set SocketOptionSettingError = m_socketOptionSettingError
End Property

''' <summary>   Get the error that occurred sending a socket message. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketSendError() As UserDefinedError
    Set SocketSendError = m_socketSendError
End Property

''' <summary>   Get the error that occurred receiving a socket message. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketReceiveError() As UserDefinedError
    Set SocketReceiveError = m_socketReceiveError
End Property

''' <summary>   Get the error that occurred binding a socket to an address. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketBindError() As UserDefinedError
    Set SocketBindError = m_socketBindError
End Property

''' <summary>   Get the error that occurred starting to listen on a socket. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketListenError() As UserDefinedError
    Set SocketListenError = m_socketListenError
End Property

''' <summary>   Get the error that occurred accepting an incoming connection. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketAcceptError() As UserDefinedError
    Set SocketAcceptError = m_socketAcceptError
End Property

''' <summary>   Get the error that occurred determining the status of a socket set. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get SocketSetStatusError() As UserDefinedError
    Set SocketSetStatusError = m_socketSetStatusError
End Property

''' <summary>   Get the error that occurred when the method argument is invalid. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get InvalidArgumentError() As UserDefinedError
    Set InvalidArgumentError = m_invalidArgumentError
End Property

''' <summary>   Get the error that occurred when the method operation is invalid. </summary>
''' <value>   Returns a user defined error. </value>
Public Property Get InvalidOperationError() As UserDefinedError
    Set InvalidOperationError = m_invalidOperationError
End Property

''' <summary>   Appends the procedure, module, and application to the Err object source string. </summary>
''' <remarks>
''' We use special markers to prefix the application
''' and module names.  This way we can locate the name
''' of the last application or module and determine if
''' we need to append module or application names or
''' just the procedure name.
'''
''' Moreover, we can determine the first source because
''' when an error first occurs in a procedure, the error
''' source string is the application title (App.Title).
''' </remarks>
''' <param name="procedureName">   Specifies the name of the procedure which to append to the source string. </param>
''' <param name="moduleName">      Specifies the module name to append to the source string. </param>
Public Sub AppendErrSource(ByVal procedureName As String, ByVal moduleName As String)

    Const appSeparator = ":"
    Const moduleSeparator = ";"
    Const procSeperator = "."
    
    ' this procedure must not trap errors because it must
    ' not alter the error object.
      
    ' thus we assume that this code is robust and will
    ' not cause errors.
  
    ' get the current source string
    Dim errorSource As String: errorSource = Err.source
    
    ' check if the error occurred in this procedure
    If (errorSource = ActiveWorkbook.name) Or (Len(errorSource) = 0) Then
    
        ' if the source string is the application title,
        ' the error occurred in the procedure.  Because
        ' this is the lowest error level, append the
        ' application and module to the source string
        errorSource = procedureName & _
                  moduleSeparator & moduleName & _
                  appSeparator & ThisWorkbook.name
  
    ' check if the error was raised from the same
    ' re-entrant procedure
    ElseIf InStr(1, errorSource, procedureName) = 1 Then
  
        ' if the same procedure, leave the error
        ' source as is.
      
    Else
    
        ' otherwise, the error was raised from another
        ' procedure.
        
        ' check if the error occurred in the same module
        
        ' get the position of the last module name
        Dim moduleNamePos As Long: moduleNamePos = InStr(1, errorSource, moduleSeparator) + 1
        
        ' check if this is the same as this module
        If InStr(moduleNamePos, errorSource, moduleName) = moduleNamePos Then
      
            ' if the error was in the same module, just
            ' append the procedure.
            errorSource = procedureName & procSeperator & _
                      errorSource
        
        Else
        
            ' if the last error was raised from another
            ' module, check if this module is in this or
            ' another application
            
            ' get the position of the last application in
            ' the string
            moduleNamePos = InStr(1, errorSource, appSeparator) + 1
            
            ' check if this is the same as this application
            If InStr(moduleNamePos, errorSource, ThisWorkbook.name) = moduleNamePos Then
        
                ' if the source is in this application than
                ' just append the module and procedure
                errorSource = procedureName & moduleSeparator & moduleName & _
                              procSeperator & errorSource
            Else
                ' if the source is another application than
                ' append the procedure, module, and application
                errorSource = procedureName & _
                         moduleSeparator & moduleName & _
                         appSeparator & ThisWorkbook.name & _
                         procSeperator & errorSource
            End If
        
        End If
      
    End If
  
    ' Update the current source string
    Err.source = errorSource
 
End Sub

''' <summary>   Builds a standard error message. </summary>
''' <param name="vbWarning">   [Optional, false] A Boolean expression
'''                            that is set to true to display a warning rather than
'''                            an error message.</param>
''' <returns>   A Standard error message string in the form: <para>
''' Error # (0x#) occurred in <c>Source</c>: Description </para><para>
''' or  </para><para>
''' Warning # (0x#) occurred in <c>Source</c>: Description  </para>
''' </returns>
Public Function BuildStandardErrorMessage(Optional ByVal vbWarning As Boolean = False) As String

    Dim builder As StringBuilder: builder = New StringBuilder
  
    ' check if we have an error
    If Err.Number <> 0 Then
    
        builder.Append IIf(vbWarning, "Warning ", "Error ")
        builder.Append Format$(Err.Number)
        
        Dim lErrNumber As Long: lErrNumber = Err.Number - vbObjectError
        
        If Abs(lErrNumber) < &HFFFF& Then
            lErrNumber = lErrNumber - &H200
            builder.Append " (+0x"
        Else
            lErrNumber = Err.Number
            builder.Append " (0x"
        End If
        builder.Append Hex$(lErrNumber)
        builder.Append ") "
        
        builder.Append ") occurred in "
        builder.Append Err.source
        builder.Append vbCrLf
        builder.Append vbCrLf
        builder.Append Err.Description
        
    End If

    BuildStandardErrorMessage = builder.ToString

End Function



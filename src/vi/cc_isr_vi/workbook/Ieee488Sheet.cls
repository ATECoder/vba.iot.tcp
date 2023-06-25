VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ieee488Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_receiveTimeoutCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_connectedHostCell As String = "C2"
Private Const m_connectedPortCell As String = "D2"
Private Const m_connectedSocketIdCell As String = "E2"
Private Const m_sentMessageCell As String = "G2"
Private Const m_receivedMessageLengthCell As String = "H2"
Private Const m_receivedMessageCell As String = "I2"
Private Const m_statusByteCell As String = "G3"
Private Const m_elapsedTimeCell As String = "J2"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_connectedCell As String = "F2"
Private Const m_identityCell As String = "I2"
Private Const m_errorMessageCell As String = "B4"

Private WithEvents m_session As cc_isr_Ieee488.Ieee488Session
Attribute m_session.VB_VarHelpID = -1
Private m_stopWatch As StopWatch

''' <summary>   Gets the singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </summary>
''' <value>   The singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </value>
Public Property Get Session() As cc_isr_Ieee488.Ieee488Session
    If m_session Is Nothing Then
        
        Set m_session = cc_isr_Ieee488.Constructor.CreateIeee488Session
    
        ' populate the combo box commands
        PopulateCommandsComboBox
    End If
    Set Session = m_session
End Property

''' <summary>   Gets the singleton instance of the <see cref="cc_isr_Core.StopWatch"/>. </summary>
''' <value>   The singleton instance of the <see cref="cc_isr_Core.StopWatch"/>. </value>
Public Property Get StopWatch() As cc_isr_Core.StopWatch
    If m_stopWatch Is Nothing Then
        Set m_stopWatch = cc_isr_Core.Constructor.CreateStopWatch
    End If
    Set StopWatch = m_stopWatch
End Property

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()
    Set m_stopWatch = Nothing
    If Not m_session Is Nothing Then m_session.Dispose
    Set m_session = Nothing
End Sub

''' <summary>   Initialize this worksheet. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Public Sub Initialize(ByRef eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Dim affirmative As Boolean: affirmative = eventArgs.Connected
    
    ' report the connection state
    Range(m_connectedCell).value = affirmative
        
    Me.RSTButton.Enabled = affirmative
    Me.SDCButton.Enabled = affirmative
    Me.CLSButton.Enabled = affirmative
    Me.QueryButton.Enabled = affirmative
    Me.ReadButton.Enabled = affirmative
    Me.ReadStatusByteButton.Enabled = affirmative
    Me.WriteButton.Enabled = affirmative

End Sub

''' <summary>   Populates the Commands combo box. </summary>
Friend Sub PopulateCommandsComboBox()
    Me.CommandsComboBox.Clear
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ClearExecutionStateCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.IdentityQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OperationCompleteCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OptionsQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ResetKnownStateCommand
    Me.CommandsComboBox.AddItem Replace(cc_isr_Ieee488.StandardEventEnableCommandFormat, "{0}", &H7F)
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.StandardEventEnableQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.StandardEventStatusQueryCommand
    Me.CommandsComboBox.AddItem Replace(cc_isr_Ieee488.ServiceRequestEnableCommandFormat, "{0}", &H7F)
    Me.CommandsComboBox.AddItem Replace(Replace(cc_isr_Ieee488.StandardServiceEnableCommandFormat, _
                                        "{0}", &H7F), "{1}", &H7F)
    Me.CommandsComboBox.AddItem Replace(Replace(cc_isr_Ieee488.StandardServiceEnableCompleteCommandFormat, _
                                        "{0}", &H7F), "{1}", &H7F)
    Me.CommandsComboBox.AddItem Replace(cc_isr_Ieee488.OperationCompleteEnableCommandFormat, "{0}", &H7F)
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ServiceRequestEnableQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ServiceRequestQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.WaitCommand
   
End Sub

''' <summary>   Restarts the stop watch and clears the elapsed time cell. </summary>
Friend Sub RestartStopWatch()
    Range(m_elapsedTimeCell).value = ""
    Me.StopWatch.Restart
End Sub

''' <summary>   Reads the stop watch and updates the elapsed time cell. </summary>
Friend Sub ReadStopWatch()
    Range(m_elapsedTimeCell).value = Format(Me.StopWatch.ElapsedMilliseconds, m_elapsedTimeFormat)
End Sub

''' <summary>   Write line handling an error messages. </summary>
''' <param name="message">   [String] to send to the instrument. </param>
Private Sub WriteLine(ByVal message As String)

    Const thisProcedureName = "WriteLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.Session.WriteLine message

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Query a message from the instrument handling any errors. </summary>
''' <param name="message">   [String] to send to the instrument. </param>
Private Function QueryLine(ByVal message As String) As String

    Const thisProcedureName = "QueryLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    QueryLine = Me.Session.QueryLine(message)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Read from the instrument handling any errors. </summary>
Private Function Read() As String

    Const thisProcedureName = "Read"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Read = Me.Session.Read()

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Handles the Toggle connection button click event. </summary>
Private Sub ConnectToggleButton_Click()

    Me.RestartStopWatch

    Dim host As String
    Dim port As Integer
    Dim timeout As Integer
    
    timeout = Me.ReceiveTimeout
    host = Me.host
    port = Me.port
    
    If Not ConnectToggleButton.value And Me.Session.Connected Then
        
        Me.Session.CloseConnection
        
    ElseIf ConnectToggleButton.value And Not Me.Session.Connected Then
    
        Me.Session.OpenConnection Me.host, Me.port, timeout
    
    End If
    
    Me.ReadStopWatch
    
End Sub

''' <summary>   Handles the Clear Know State button click event. </summary>
Private Sub CLSButton_Click()
    
    Const thisProcedureName = "CLSButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    Me.Session.ClearExecutionState
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()
    
    Const thisProcedureName = "RSTButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    Me.Session.ResetKnownState
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>  Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    
    Const thisProcedureName = "SDCButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    Me.Session.SelectiveDeviceClear
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Friend Sub ReadStatusByteButton_Click()
    
    Const thisProcedureName = "ReadStatusByteButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    Me.StatusByte = Me.Session.ReadStatusByte
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()
    Me.SentMessage = ""
    Dim messageToSend As String: messageToSend = Me.CommandsComboBox.value
    Me.RestartStopWatch
    Me.WriteLine messageToSend
    Me.ReadStopWatch
    Me.SentMessage = messageToSend
    If Me.AutoStatusReadCheckBox.value Then
        Me.ReadStatusByteButton_Click
    End If
End Sub

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()
    Me.ReceivedMessage = ""
    Dim messageReceived As String
    Me.RestartStopWatch
    messageReceived = Me.Read
    Me.ReadStopWatch
    Me.ReceivedMessage = messageReceived
End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    Me.ReceivedMessage = ""
    Me.SentMessage = ""
    Dim messageToSend As String: messageToSend = Me.CommandsComboBox.value
    Dim messageReceived As String
    Me.RestartStopWatch
    messageReceived = Me.QueryLine(messageToSend)
    Me.ReadStopWatch
    Me.ReceivedMessage = messageReceived
    Me.SentMessage = messageToSend
End Sub

''' <summary>   Gets the host dotted IP Address. </summary>
Friend Property Get host() As String
    host = Range(m_inputHostCell).value
End Property

''' <summary>   Gets the port number. </summary>
Friend Property Get port() As Integer
    port = CInt(Range(m_inputPortCell).value)
End Property

''' <summary> Get the receive timeout. </summary>
Friend Property Get ReceiveTimeout() As Integer
    ReceiveTimeout = CInt(Range(m_receiveTimeoutCell).value)
End Property

''' <summary>   Gets the received message. </summary>
Friend Property Get ReceivedMessage() As String
    ReceivedMessage = Range(m_receivedMessageCell).value
End Property

''' <summary>   Sets the received message. </summary>
Friend Property Let ReceivedMessage(ByVal value As String)
    Range(m_receivedMessageCell).value = value
    Range(m_receivedMessageLengthCell).value = Len(value)
End Property

''' <summary>   Gets the sent message. </summary>
Friend Property Get SentMessage() As String
    SentMessage = Range(m_sentMessageCell).value
End Property

''' <summary>   Sets the sent message. </summary>
Friend Property Let SentMessage(ByVal value As String)
    Range(m_sentMessageCell).value = value
End Property

''' <summary>   Gets the status byte. </summary>
Friend Property Get StatusByte() As Integer
    StatusByte = Range(m_statusByteCell).value
End Property

''' <summary>   Sets the status byte. </summary>
Friend Property Let StatusByte(ByVal value As Integer)
    Range(m_statusByteCell).value = value
End Property

''' <summary>   Handles the session <see cref="cc_isr_VI.ConnectionChanged"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_session_ConnectionChanged(ByRef eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Me.Initialize eventArgs

    If eventArgs.Connected Then
    
        ' report the socket number
        Range(m_connectedHostCell).value = Me.host
        Range(m_connectedPortCell).value = Me.port
        Range(m_connectedSocketIdCell).value = Me.Session.SocketId
    
    Else
    
        Range(m_connectedHostCell).value = ""
        Range(m_connectedPortCell).value = ""
        Range(m_connectedSocketIdCell).value = ""
        
    End If

End Sub



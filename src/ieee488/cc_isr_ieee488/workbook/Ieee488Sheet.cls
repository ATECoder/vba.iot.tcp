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
Private Const m_socketAddressCell As String = "C2"
Private Const m_socketIdCell As String = "E2"
Private Const m_connectedCell As String = "F2"
Private Const m_sentMessageCell As String = "G2"
Private Const m_receivedMessageLengthCell As String = "K2"
Private Const m_receivedMessageCell As String = "L2"
Private Const m_elapsedTimeCell As String = "M2"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_lastErrorMessageCell As String = "B4"

Private Const m_serialPollCell As String = "F12"
Private Const m_statusByteCell As String = "G12"
Private Const m_StandardByteCell As String = "I12"
Private Const m_srqCell As String = "C10"
Private Const m_gpibAddressCell As String = "C12"
Private Const m_readTimeoutCell As String = "C15"
Private Const m_readAfterWriteEnabledCell As String = "C18"

Private WithEvents m_session As cc_isr_Ieee488.Ieee488Session
Attribute m_session.VB_VarHelpID = -1
Private m_stopWatch As StopWatch

''' <summary>   Gets the singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </summary>
''' <value>   The singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </value>
Public Property Get Session() As cc_isr_Ieee488.Ieee488Session
    If m_session Is Nothing Then
        Set m_session = cc_isr_Ieee488.Constructor.CreateIeee488Session
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

    Dim disconnected As Boolean
    
    If m_session Is Nothing Then
        disconnected = True
    Else
        disconnected = Not m_session.Connected
    End If
    
    If Not disconnected Then disconnected = m_session.CloseConnection
    
    If disconnected And Not m_session Is Nothing Then
        m_session.Dispose
        Set m_session = Nothing
    End If

    Set m_stopWatch = Nothing

End Sub

''' <summary>   Initialize this worksheet. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Public Sub Initialize(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Const p_procedureName = "PopulateCommandsComboBox"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    
    On Error Resume Next
    
    Dim p_connected As Boolean: p_connected = eventArgs.Connected
    
    Me.ConnectToggleButton.value = p_connected
    
    ' report the connection state
    Range(m_connectedCell).value = p_connected
        
    Me.RSTButton.Enabled = p_connected
    Me.CLSButton.Enabled = p_connected
    Me.QueryButton.Enabled = p_connected
    Me.ReadButton.Enabled = p_connected
    Me.ReadStatusByteButton.Enabled = p_connected
    Me.WriteButton.Enabled = p_connected
    
    If p_connected Then
        EnableGpibLan m_session.ViSession.UsingGpibLan
    Else
        EnableGpibLan False
    End If
    
    ' populate the combo box commands
    PopulateCommandsComboBox

    If Err.Number <> 0 Then
        ' append the error source
        UserDefinedErrors.SetErrSource p_procedureName, Me.name
        
        ' display the error message
        Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
        
        On Error GoTo 0
    End If

End Sub

''' <summary>   Populates the Commands combo box. </summary>
Friend Sub PopulateCommandsComboBox()
    
    Const p_procedureName = "PopulateCommandsComboBox"
    
    On Error Resume Next
    
    Me.CommandsComboBox.Clear
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ClearExecutionStateCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ClearExecutionStateCommand & _
            ";" & cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.IdentityQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OperationCompleteCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OptionsQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ResetKnownStateCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ResetKnownStateCommand & _
            ";" & cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
            cc_isr_Ieee488.StandardEventEnableCommand, &H7F) & _
            ";" & cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.StandardEventEnableQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.StandardEventStatusQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
            cc_isr_Ieee488.ServiceRequestEnableCommand, &H7F) & _
            ";" & cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
            cc_isr_Ieee488.StandardServiceEnableCommand, &H7F, &H7F) & _
            ";" & cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
            cc_isr_Ieee488.StandardServiceEnableCommand, &H7F, &H7F) & _
            ";" & cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ServiceRequestEnableQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ServiceRequestQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.WaitCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.WaitCommand & _
            ";" & cc_isr_Ieee488.OperationCompletedQueryCommand
   
    If Err.Number <> 0 Then
        ' append the error source
        UserDefinedErrors.SetErrSource p_procedureName, Me.name
        
        ' display the error message
        Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
        
        On Error GoTo 0
    End If
    
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
Friend Sub WriteLine(ByVal message As String)

    Const p_procedureName = "WriteLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.WriteLine message
    
    Me.ReadStopWatch
    
    If Me.Session.ViSession.UsingGpibLan Then
        Me.ReadAfterWriteEnabled = Me.Session.ViSession.GpibLan.ReadAfterWriteEnabled
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler
    
End Sub

''' <summary>   Query a message from the instrument handling any errors. </summary>
''' <param name="message">   [String] to send to the instrument. </param>
Friend Function QueryLine(ByVal message As String) As String

    Const p_procedureName = "QueryLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    QueryLine = Me.Session.QueryLine(message)
    
    Me.ReadStopWatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler
    
End Function

''' <summary>   Read from the instrument handling any errors. </summary>
Friend Function Read() As String

    Const p_procedureName = "Read"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Read = Me.Session.Read()
    
    Me.ReadStopWatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler
    
End Function

''' <summary>   Handles the Toggle connection button click event. </summary>
Friend Sub ConnectToggleButton_Click()

    Const p_procedureName = "ConnectToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
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
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Clear Know State button click event. </summary>
Private Sub CLSButton_Click()
    
    Const p_procedureName = "CLSButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ClearExecutionState
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()
    
    Const p_procedureName = "RSTButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ResetKnownState
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    
    Const p_procedureName = "SDCButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ViSession.GpibLan.SelectiveDeviceClear
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Private Sub ReadStandardEventStatusButton_Click()

    Const p_procedureName = "ReadStandardEventStatusButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.StandardByte = Me.Session.QueryStandardEventsStatus
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler
    
End Sub


''' <summary>   Handles the Read Status Byte button click event. </summary>
Friend Sub ReadStatusByteButton_Click()
    
    Const p_procedureName = "ReadStatusByteButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.StatusByte = Me.Session.QueryServiceRequestStatus
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()

    Const p_procedureName = "WriteButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.SentMessage = ""
    Dim messageToSend As String: messageToSend = Me.CommandsComboBox.value
    
    Me.RestartStopWatch
    
    Me.WriteLine messageToSend
    
    Me.ReadStopWatch
    
    Me.SentMessage = messageToSend
    If Me.Session.ViSession.UsingGpibLan And Me.AutoStatusReadCheckBox.value Then
        Range(m_serialPollCell).value = Me.Session.ViSession.GpibLan.SerialPoll
    End If
    
    Dim isQuery As Boolean: isQuery = cc_isr_Core.StringExtensions.EndsWith(messageToSend, "?")
    If Not isQuery And Me.AutoStatusReadCheckBox.value Then
        Me.StatusByte = Me.Session.QueryServiceRequestStatus
        Me.StandardByte = Me.Session.QueryStandardEventsStatus
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()

    Const p_procedureName = "ReadButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReceivedMessage = ""
    Dim messageReceived As String
    
    Me.RestartStopWatch
    
    messageReceived = Me.Read
    
    Me.ReadStopWatch
    
    Me.ReceivedMessage = messageReceived
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    
    Const p_procedureName = "QueryButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReceivedMessage = ""
    Me.SentMessage = ""
    Dim messageToSend As String: messageToSend = Me.CommandsComboBox.value
    Dim messageReceived As String
    
    Me.RestartStopWatch
    
    messageReceived = Me.QueryLine(messageToSend)
    
    Me.ReadStopWatch
    
    Me.ReceivedMessage = messageReceived
    Me.SentMessage = messageToSend
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

''' <summary>   Gets the host dotted IP Address. </summary>
Friend Property Get host() As String
    host = Range(m_inputHostCell).value
End Property

''' <summary>   Gets the port number. </summary>
Friend Property Get port() As Integer
    port = CInt(Range(m_inputPortCell).value)
End Property

''' <summary>   Get the receive timeout. </summary>
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

''' <summary>   Gets the Last Error Message. </summary>
Friend Property Get LastErrorMessage() As String
    LastErrorMessage = Range(m_lastErrorMessageCell).value
End Property

''' <summary>   Sets the Last Error Message. </summary>
Friend Property Let LastErrorMessage(ByVal value As String)
    Range(m_lastErrorMessageCell).value = value
End Property

''' <summary>   Gets the status byte. </summary>
Friend Property Get StatusByte() As Integer
    StatusByte = Range(m_statusByteCell).value
End Property

''' <summary>   Sets the status byte. </summary>
Friend Property Let StatusByte(ByVal value As Integer)
    Range(m_statusByteCell).value = value
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
Friend Property Get ReadAfterWriteEnabled() As Boolean
    ReadAfterWriteEnabled = CBool(Range(m_readAfterWriteEnabledCell).value)
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
Friend Property Let ReadAfterWriteEnabled(ByVal value As Boolean)
    Range(m_readAfterWriteEnabledCell).value = value
End Property

''' <summary>   Gets the Standard byte. </summary>
Friend Property Get StandardByte() As Integer
    StandardByte = Range(m_StandardByteCell).value
End Property

''' <summary>   Sets the Standard byte. </summary>
Friend Property Let StandardByte(ByVal value As Integer)
    Range(m_StandardByteCell).value = value
End Property

' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +
'
' GPIB LAN
'
' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +

Friend Sub EnableGpibLan(ByVal usingGpibLanController As Boolean)
    Me.SDCButton.Enabled = usingGpibLanController
    Me.GoToLocalButton.Enabled = usingGpibLanController
    Me.ReadAfterWriteGetButton.Enabled = usingGpibLanController
    Me.ReadAfterWriteSetButton.Enabled = usingGpibLanController
    Me.LocalLockoutButton.Enabled = usingGpibLanController
    Me.SerialPollButton.Enabled = usingGpibLanController
    Me.SRQButton.Enabled = usingGpibLanController
    Me.GpibAddressGetButton.Enabled = usingGpibLanController
    Me.GpibAddressSetButton.Enabled = usingGpibLanController
    Me.ReadTimeoutGetButton.Enabled = usingGpibLanController
    Me.ReadTimeoutSetButton.Enabled = usingGpibLanController
End Sub

Private Sub GoToLocalButton_Click()

    Const p_procedureName = "GoToLocalButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ViSession.GpibLan.GoToLocal
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

Private Sub LocalLockoutButton_Click()

    Const p_procedureName = "LocalLockoutButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ViSession.GpibLan.LocalLockout
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

Private Sub ReadAfterWriteSetButton_Click()

    Const p_procedureName = "ReadAfterWriteSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ViSession.GpibLan.ReadAfterWriteEnabledSetter Me.ReadAfterWriteEnabled
    
    Me.ReadStopWatch
    
    ' if Read-After-Write is enabled, make sure to set the GPIB-Lan class to
    ' turn it off on write.
    
    Me.Session.ViSession.GpibLan.DisableReadAfterWriteOnWrite = Me.ReadAfterWriteEnabled
    

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

Private Sub ReadAfterWriteGetButton_Click()

    Const p_procedureName = "ReadAfterWriteGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.ReadAfterWriteEnabled = Me.Session.ViSession.GpibLan.ReadAfterWriteEnabledGetter
    
    Me.ReadStopWatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler


End Sub

Private Sub SerialPollButton_Click()

    Const p_procedureName = "SerialPollButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_serialPollCell).value = Me.Session.ViSession.GpibLan.SerialPoll
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

Private Sub SRQButton_Click()

    Const p_procedureName = "SRQButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_srqCell).value = Me.Session.ViSession.GpibLan.ServiceRequested
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

Private Sub SetGibAddressButton_Click()

    Const p_procedureName = "GpibAddressSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ViSession.GpibLan.GpibAddressSetter CInt(Range(m_gpibAddressCell).value)

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

Private Sub GpibAddressGetButton_Click()

    Const p_procedureName = "GpibAddressGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_gpibAddressCell).value = Me.Session.ViSession.GpibLan.GpibAddressGetter

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

Private Sub ReadTimeoutGetButton_Click()

    Const p_procedureName = "ReadTimeoutGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_readTimeoutCell).value = Me.Session.ViSession.GpibLan.ReadTimeoutGetter

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub

Private Sub ReadTimeoutSetButton_Click()

    Const p_procedureName = "ReadTimeoutSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ViSession.GpibLan.ReadTimeoutSetter (CInt(Range(m_readTimeoutCell).value))
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error GoTo 0
    GoTo exit_Handler

End Sub


' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +
'
' IEEE 488 Session Events
'
' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +

''' <summary>   Handles the <see cref="cc_isr_Ieee488.Ieee488Session.ConnectionChanged"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_session_ConnectionChanged(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Me.Initialize eventArgs

    If eventArgs.Connected Then
    
        ' report the socket number
        Range(m_socketAddressCell).value = Me.Session.ViSession.SocketAddress
        Range(m_socketIdCell).value = Me.Session.SocketId
    
    Else
    
        Range(m_socketAddressCell).value = ""
        Range(m_socketIdCell).value = ""
        
    End If
    
    If Me.Session.ViSession.UsingGpibLan Then
    
        Me.ReadAfterWriteEnabled = Me.Session.ViSession.GpibLan.ReadAfterWriteEnabled
    
    End If

End Sub

''' <summary>   Handles the <see cref="cc_isr_Ieee488.Ieee488Session.ConnectionChanging"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_session_ConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

End Sub





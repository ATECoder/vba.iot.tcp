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

Private Const m_sessionReadTimeoutCell As String = "B1"
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
Private Const m_gpibLanReadTimeoutCell  As String = "C15"
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

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
        Dim p_disconnected As Boolean
    
    If m_session Is Nothing Then
        p_disconnected = True
    Else
        p_disconnected = Not m_session.Connected
    End If
    
    If Not p_disconnected Then p_disconnected = m_session.CloseConnection
    
    If p_disconnected And Not m_session Is Nothing Then
        m_session.Dispose
        Set m_session = Nothing
    End If

    Set m_stopWatch = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Initialize this worksheet. </summary>
Public Sub Initialize()

    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.EnableControls False
    
    ' populate the combo box commands
    PopulateCommandsComboBox
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Toggle the enable state of the controls. </summary>
''' <param name="a_connected">   [Boolean] True if connected; otehrwise, false. </param>
Friend Sub EnableControls(ByVal a_connected As Boolean)

    Const p_procedureName = "EnableControls"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    
    Me.ConnectToggleButton.value = a_connected
    
    ' report the connection state
    Range(m_connectedCell).value = a_connected
        
    Me.RSTButton.Enabled = a_connected
    Me.CLSButton.Enabled = a_connected
    Me.QueryButton.Enabled = a_connected
    Me.ReadButton.Enabled = a_connected
    Me.ReadStatusByteButton.Enabled = a_connected
    Me.ReadStandardEventStatusButton.Enabled = a_connected
    Me.WriteButton.Enabled = a_connected
    
    If a_connected Then
        EnableGpibLanControls m_session.ViSession.UsingGpibLan
    Else
        EnableGpibLanControls False
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Populates the Commands combo box. </summary>
Friend Sub PopulateCommandsComboBox()
    
    Const p_procedureName = "PopulateCommandsComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
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

    Me.CommandsComboBox.ListIndex = 1

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Restarts the stop watch and clears the elapsed time cell. </summary>
Friend Sub RestartStopWatch()
    
    Range(m_elapsedTimeCell).value = vbNullString
    Me.StopWatch.Restart

End Sub

''' <summary>   Reads the stop watch and updates the elapsed time cell. </summary>
Friend Sub ReadStopWatch()
    
    Range(m_elapsedTimeCell).value = Format(Me.StopWatch.ElapsedMilliseconds, m_elapsedTimeFormat)

End Sub

''' <summary>   Write line handling an error messages. </summary>
''' <param name="a_message">   [String] to send to the instrument. </param>
Friend Sub WriteLine(ByVal a_message As String)

    Const p_procedureName = "WriteLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.WriteLine a_message
    
    Me.ReadStopWatch
    
    If Me.Session.ViSession.UsingGpibLan Then
        Me.ReadAfterWriteEnabled = Me.Session.ViSession.GpibLan.ReadAfterWriteEnabled
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Query a message from the instrument handling any errors. </summary>
''' <param name="a_message">   [String] to send to the instrument. </param>
Friend Function QueryLine(ByVal a_message As String) As String

    Const p_procedureName = "QueryLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    QueryLine = Me.Session.QueryLine(a_message)
    
    Me.ReadStopWatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Handles the Toggle connection button click event. </summary>
Friend Sub ConnectToggleButton_Click()

    Const p_procedureName = "ConnectToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch

    Dim p_host As String: p_host = Me.host
    Dim p_port As Long: p_port = Me.port
    Dim p_sessionReadTimeout As Long: p_sessionReadTimeout = Me.SessionReadTimeout
    
    If Not ConnectToggleButton.value And Me.Session.Connected Then
        
        Me.Session.CloseConnection
        
    ElseIf ConnectToggleButton.value And Not Me.Session.Connected Then
    
        ' the socket read timeout default is used here.
        
        Me.Session.OpenConnection p_host, p_port
        
        Me.Session.ViSession.SessionReadTimeout = p_sessionReadTimeout
    
    End If
    
    Me.ReadStopWatch
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()

    Const p_procedureName = "WriteButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.SentMessage = vbNullString
    Dim p_messageToSend As String: p_messageToSend = Me.CommandsComboBox.value
    
    Me.RestartStopWatch
    
    Me.WriteLine p_messageToSend
    
    Me.ReadStopWatch
    
    Me.SentMessage = p_messageToSend
    If Me.Session.ViSession.UsingGpibLan And Me.AutoStatusReadCheckBox.value Then
        Range(m_serialPollCell).value = Me.Session.ViSession.GpibLan.SerialPoll
    End If
    
    Dim p_isQuery As Boolean: p_isQuery = cc_isr_Core.StringExtensions.EndsWith(p_messageToSend, "?")
    If Not p_isQuery And Me.AutoStatusReadCheckBox.value Then
        Me.StatusByte = Me.Session.QueryServiceRequestStatus
        Me.StandardByte = Me.Session.QueryStandardEventsStatus
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()

    Const p_procedureName = "ReadButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReceivedMessage = vbNullString
    Dim p_messageReceived As String
    
    Me.RestartStopWatch
    
    p_messageReceived = Me.Read
    
    Me.ReadStopWatch
    
    Me.ReceivedMessage = p_messageReceived
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    
    Const p_procedureName = "QueryButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReceivedMessage = vbNullString
    Me.SentMessage = vbNullString
    Dim p_messageToSend As String: p_messageToSend = Me.CommandsComboBox.value
    Dim p_messageReceived As String
    
    Me.RestartStopWatch
    
    p_messageReceived = Me.QueryLine(p_messageToSend)
    
    Me.ReadStopWatch
    
    Me.ReceivedMessage = p_messageReceived
    Me.SentMessage = p_messageToSend
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

''' <summary>   Get the session read timeout. </summary>
Friend Property Get SessionReadTimeout() As Integer
    SessionReadTimeout = CInt(Range(m_sessionReadTimeoutCell).value)
End Property

''' <summary>   Get the GPIB-Lan Read timeout. </summary>
Friend Property Get GpibLanReadTimeout() As Integer
    GpibLanReadTimeout = CInt(Range(m_gpibLanReadTimeoutCell).value)
End Property

''' <summary>   Set the GPIB-Lan Read timeout. </summary>
Friend Property Let GpibLanReadTimeout(ByVal a_value As Integer)
    Range(m_gpibLanReadTimeoutCell).value = a_value
End Property


''' <summary>   Gets the received message. </summary>
Friend Property Get ReceivedMessage() As String
    ReceivedMessage = Range(m_receivedMessageCell).value
End Property

''' <summary>   Sets the received message. </summary>
Friend Property Let ReceivedMessage(ByVal a_value As String)
    Range(m_receivedMessageCell).value = a_value
    Range(m_receivedMessageLengthCell).value = Len(a_value)
End Property

''' <summary>   Gets the sent message. </summary>
Friend Property Get SentMessage() As String
    SentMessage = Range(m_sentMessageCell).value
End Property

''' <summary>   Sets the sent message. </summary>
Friend Property Let SentMessage(ByVal a_value As String)
    Range(m_sentMessageCell).value = a_value
End Property

''' <summary>   Gets the Last Error Message. </summary>
Friend Property Get LastErrorMessage() As String
    LastErrorMessage = Range(m_lastErrorMessageCell).value
End Property

''' <summary>   Sets the Last Error Message. </summary>
Friend Property Let LastErrorMessage(ByVal a_value As String)
    Range(m_lastErrorMessageCell).value = a_value
End Property

''' <summary>   Gets the status byte. </summary>
Friend Property Get StatusByte() As Integer
    StatusByte = Range(m_statusByteCell).value
End Property

''' <summary>   Sets the status byte. </summary>
Friend Property Let StatusByte(ByVal a_value As Integer)
    Range(m_statusByteCell).value = a_value
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
Friend Property Get ReadAfterWriteEnabled() As Boolean
    ReadAfterWriteEnabled = CBool(Range(m_readAfterWriteEnabledCell).value)
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
Friend Property Let ReadAfterWriteEnabled(ByVal a_value As Boolean)
    Range(m_readAfterWriteEnabledCell).value = a_value
End Property

''' <summary>   Gets the Standard byte. </summary>
Friend Property Get StandardByte() As Integer
    StandardByte = Range(m_StandardByteCell).value
End Property

''' <summary>   Sets the Standard byte. </summary>
Friend Property Let StandardByte(ByVal a_value As Integer)
    Range(m_StandardByteCell).value = a_value
End Property

' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +
'
' GPIB LAN
'
' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +

''' <summary>   Toggle the enabled state of the GPIB Lan controls. </summary>
''' <param name="a_usingGpibLanController">   [Boolean] True if using the GPIB-Lan controller. </param>
Friend Sub EnableGpibLanControls(ByVal a_usingGpibLanController As Boolean)
    Me.SDCButton.Enabled = a_usingGpibLanController
    Me.GoToLocalButton.Enabled = a_usingGpibLanController
    Me.ReadAfterWriteGetButton.Enabled = a_usingGpibLanController
    Me.ReadAfterWriteSetButton.Enabled = a_usingGpibLanController
    Me.LocalLockoutButton.Enabled = a_usingGpibLanController
    Me.SerialPollButton.Enabled = a_usingGpibLanController
    Me.SRQButton.Enabled = a_usingGpibLanController
    Me.GpibAddressGetButton.Enabled = a_usingGpibLanController
    Me.GpibAddressSetButton.Enabled = a_usingGpibLanController
    Me.ReadTimeoutGetButton.Enabled = a_usingGpibLanController
    Me.ReadTimeoutSetButton.Enabled = a_usingGpibLanController
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

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
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

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub ReadTimeoutGetButton_Click()

    Const p_procedureName = "ReadTimeoutGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_gpibLanReadTimeoutCell).value = Me.Session.ViSession.GpibLan.ReadTimeoutGetter

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub ReadTimeoutSetButton_Click()

    Const p_procedureName = "ReadTimeoutSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ViSession.GpibLan.ReadTimeoutSetter (CInt(Range(m_gpibLanReadTimeoutCell).value))
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +
'
' IEEE 488 Session Events
'
' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +

''' <summary>   Handles the <see cref="cc_isr_Ieee488.Ieee488Session.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_session_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
        
    Const p_procedureName = "m_session_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.EnableControls a_eventArgs.Connected

    If a_eventArgs.Connected Then
    
        ' report the socket number
        Range(m_socketAddressCell).value = Me.Session.ViSession.SocketAddress
        Range(m_socketIdCell).value = Me.Session.SocketId
    
    Else
    
        Range(m_socketAddressCell).value = vbNullString
        Range(m_socketIdCell).value = vbNullString
        
    End If
    
    If Me.Session.ViSession.UsingGpibLan Then
    
        Me.ReadAfterWriteEnabled = Me.Session.ViSession.GpibLan.ReadAfterWriteEnabled
    
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    Me.LastErrorMessage = ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
   
End Sub

''' <summary>   Handles the <see cref="cc_isr_Ieee488.Ieee488Session.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_session_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub





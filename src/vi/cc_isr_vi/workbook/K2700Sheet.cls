VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_receiveTimeoutCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_socketAddressCell As String = "C2"
Private Const m_socketIdCell As String = "D2"
Private Const m_connectedCell As String = "E2"
Private Const m_sentMessageCell As String = "F2"
Private Const m_receivedMessageLengthCell As String = "J2"
Private Const m_receivedMessageCell As String = "K2"
Private Const m_elapsedTimeCell As String = "L2"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_errorMessageCell As String = "B4"

Private Const m_serialPollCell As String = "E11"
Private Const m_statusByteCell As String = "F11"
Private Const m_standardByteCell As String = "H11"

Private Const m_srqCell As String = "C10"
Private Const m_gpibAddressCell As String = "C12"
Private Const m_readTimeoutCell As String = "C15"

Private Const m_timerIntervalCell As String = "C19"
Private Const m_card1Cell As String = "C20"
Private Const m_card2Cell As String = "C21"
Private Const m_card1ScanListCell As String = "C22"
Private Const m_card2ScanListCell As String = "C23"
Private Const m_inputsCell As String = "C24"
Private Const m_resistanceNoCell As String = "C25"
Private Const m_resistanceCell As String = "C26"
Private Const m_resistanceCountCell As String = "C27"
Private Const m_externalTriggerStatusCell As String = "C28"

Private Const m_resistanceNoTopRow As Integer = 2
Private Const m_resistanceNoCol As String = "M"
Private Const m_resistanceCol As String = "N"
Private Const m_resistanceFormat As String = "0.00"

' defines the standard event timer.
Private WithEvents m_eventTimer As cc_isr_Core.stdTimer
Attribute m_eventTimer.VB_VarHelpID = -1

' true if a pause is requested so that the timer event is not handled
' and the timer is paused on the next event.
Private m_pauseRequested As Boolean

' true when handling the timer event to prevent timer
' event handler overrun.
Private m_handlingTimerEvent As Boolean

Private m_extTrigInitiated  As Boolean

Private m_currentResistanceNo As Integer

Private WithEvents m_k2700 As K2700
Attribute m_k2700.VB_VarHelpID = -1

Private m_stopWatch As StopWatch

''' <summary>   Gets the singleton instance of the 2700 virtual instrument. </summary>
''' <value>   The singleton instance of the 2700 virtual instrument. </value>
Public Property Get K2700() As cc_isr_VI.K2700
    If m_k2700 Is Nothing Then
        Set m_k2700 = cc_isr_VI.Constructor.K2700
        ' populate the combo box commands
        PopulateCommandsComboBox
    End If
    Set K2700 = m_k2700
End Property

''' <summary>   Gets the singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </summary>
''' <value>   The singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </value>
Public Property Get Session() As cc_isr_Ieee488.Ieee488Session
    Set Session = Me.K2700.Session
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
    If Not m_k2700 Is Nothing Then m_k2700.Dispose
    Set m_k2700 = Nothing
End Sub

''' <summary>   Initialize this worksheet. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Public Sub Initialize(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Dim affirmative As Boolean: affirmative = eventArgs.Connected
    
    Me.ConnectToggleButton.value = affirmative
    
    ' report the connection state
    Range(m_connectedCell).value = affirmative
        
    Me.RSTButton.Enabled = affirmative
    Me.CLSButton.Enabled = affirmative
    Me.QueryButton.Enabled = affirmative
    Me.ReadButton.Enabled = affirmative
    Me.ReadStatusByteButton.Enabled = affirmative
    Me.WriteButton.Enabled = affirmative
    
    Me.ReadCardsButton.Enabled = affirmative
    Me.SetScansButton.Enabled = affirmative
    Me.QueryInputsButton.Enabled = affirmative
    Me.ReadResistanceButton.Enabled = affirmative
    
    If affirmative Then
        EnableGpibLan Me.Session.VISession.UsingGpibLan
    Else
        EnableGpibLan False
    End If

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
   
    Me.CommandsComboBox.AddItem ":FUNC 'RES',(@101:120)"
    Me.CommandsComboBox.AddItem ":FUNC 'RES',(@201:220)"
    Me.CommandsComboBox.AddItem ":SYST:BEEP"
    Me.CommandsComboBox.AddItem ":SYST:CLE"
    Me.CommandsComboBox.AddItem ":SYST:ERR?"
    Me.CommandsComboBox.AddItem ":SYST:FRSW?"
   
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE 488 Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Write line handling an error messages. </summary>
''' <param name="message">   [String] to send to the instrument. </param>
Friend Sub WriteLine(ByVal message As String)

    Const thisProcedureName = "WriteLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.WriteLine message
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Query a message from the instrument handling any errors. </summary>
''' <param name="message">   [String] to send to the instrument. </param>
Friend Function QueryLine(ByVal message As String) As String

    Const thisProcedureName = "QueryLine"
    
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
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Read from the instrument handling any errors. </summary>
Friend Function Read() As String

    Const thisProcedureName = "Read"
    
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
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Handles the Toggle connection button click event. </summary>
Friend Sub ConnectToggleButton_Click()

    Const thisProcedureName = "ConnectToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch

    Dim host As String
    Dim port As Integer
    Dim timeout As Integer
    
    timeout = Me.ReceiveTimeout
    host = Me.host
    port = Me.port
    
    If Not ConnectToggleButton.value And Me.K2700.Connected Then
        
        Me.K2700.CloseConnection
        
    ElseIf ConnectToggleButton.value And Not Me.K2700.Connected Then
    
        Me.K2700.OpenConnection Me.host, Me.port, timeout
        
    End If
    
    Me.ReadStopWatch
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
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

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Private Sub ReadStandardEventStatusButton_Click()

    Const thisProcedureName = "ReadStandardEventStatusButton"
    
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
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
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

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    
    Const thisProcedureName = "SDCButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.SelectiveDeviceClear
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
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
    
    If Me.Session.VISession.UsingGpibLan Then
    
        Me.RestartStopWatch
        
        Range(m_serialPollCell).value = Me.Session.VISession.GpibLan.SerialPoll
        
        Me.ReadStopWatch
        
    End If
    
    Me.RestartStopWatch
    
    Me.StatusByte = Me.Session.QueryServiceRequestStatus
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()

    Const thisProcedureName = "WriteButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.SentMessage = ""
    Dim messageToSend As String: messageToSend = Me.CommandsComboBox.value
    
    Me.RestartStopWatch
    
    Me.WriteLine messageToSend
    
    Me.ReadStopWatch
    
    Me.SentMessage = messageToSend
    If Me.Session.VISession.UsingGpibLan And Me.AutoStatusReadCheckBox.value Then
        Range(m_serialPollCell).value = Me.Session.VISession.GpibLan.SerialPoll
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
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()

    Const thisProcedureName = "ReadButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.receivedMessage = ""
    Dim messageReceived As String
    
    Me.RestartStopWatch
    
    messageReceived = Me.Read
    
    Me.ReadStopWatch
    
    Me.receivedMessage = messageReceived
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    
    Const thisProcedureName = "QueryButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.receivedMessage = ""
    Me.SentMessage = ""
    Dim messageToSend As String: messageToSend = Me.CommandsComboBox.value
    Dim messageReceived As String
    
    Me.RestartStopWatch
    
    messageReceived = Me.QueryLine(messageToSend)
    
    Me.ReadStopWatch
    
    Me.receivedMessage = messageReceived
    Me.SentMessage = messageToSend
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' SETTINGS
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

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
Friend Property Get receivedMessage() As String
    receivedMessage = Range(m_receivedMessageCell).value
End Property

''' <summary>   Sets the received message. </summary>
Friend Property Let receivedMessage(ByVal value As String)
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

''' <summary>   Gets the Standard byte. </summary>
Friend Property Get StandardByte() As Integer
    StandardByte = Range(m_standardByteCell).value
End Property

''' <summary>   Sets the Standard byte. </summary>
Friend Property Let StandardByte(ByVal value As Integer)
    Range(m_standardByteCell).value = value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB LAN
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +


Friend Sub EnableGpibLan(ByVal usingGpibLanController As Boolean)
    Me.SDCButton.Enabled = usingGpibLanController
    Me.GoToLocalButton.Enabled = usingGpibLanController
    Me.ListenTalkToggleButton.Enabled = usingGpibLanController
    Me.LocalLockoutButton.Enabled = usingGpibLanController
    Me.SerialPollButton.Enabled = usingGpibLanController
    Me.SRQButton.Enabled = usingGpibLanController
    Me.GetGpibAddressButton.Enabled = usingGpibLanController
    Me.SetGpibAddressButton.Enabled = usingGpibLanController
    Me.ReadTimeoutGetButton.Enabled = usingGpibLanController
    Me.ReadTimeoutSetButton.Enabled = usingGpibLanController
End Sub

Private Sub GoToLocalButton_Click()

    Const thisProcedureName = "GoToLocalButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.GoToLocal
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub LocalLockoutButton_Click()

    Const thisProcedureName = "LocalLockoutButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.LocalLockout
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub ListenTalkToggleButton_Click()

    Const thisProcedureName = "ListenTalkToggleButton"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.ReadAfterWriteEnabledSetter Not Me.Session.VISession.GpibLan.ReadAfterWriteEnabled
    
    Me.ReadStopWatch
    
    Me.ListenTalkToggleButton.Caption = IIf(Me.Session.VISession.GpibLan.ReadAfterWriteEnabledGetter, "Talk", "Listen")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub SerialPollButton_Click()

    Const thisProcedureName = "SerialPollButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_serialPollCell).value = Me.Session.VISession.GpibLan.SerialPoll
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub SRQButton_Click()

    Const thisProcedureName = "SRQButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_srqCell).value = Me.Session.VISession.GpibLan.ServiceRequested
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub GetGpibAddressButton_Click()

    Const thisProcedureName = "GetGpibAddressButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_gpibAddressCell).value = Me.Session.VISession.GpibLan.GpibAddressGetter

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub SetGpibAddressButton_Click()

    Const thisProcedureName = "SetGpibAddressButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.GpibAddresssetter CInt(Range(m_gpibAddressCell).value)

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


Private Sub ReadTimeoutGetButton_Click()

    Const thisProcedureName = "ReadTimeoutGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_readTimeoutCell).value = Me.Session.VISession.GpibLan.ReadTimeoutGetter

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub ReadTimeoutSetButton_Click()

    Const thisProcedureName = "ReadTimeoutSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.ReadTimeoutSetter (CInt(Range(m_readTimeoutCell).value))
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' 2700 Measurments Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary> Get the timer interval. </summary>
Friend Property Get TimerInerval() As Integer
    TimerInerval = CInt(Range(m_timerIntervalCell).value)
End Property

''' <summary> Get the resistance count. </summary>
Friend Property Get ResistanceCount() As Integer
    ResistanceCount = CInt(Range(m_resistanceCountCell).value)
End Property

''' <summary>   Indicates that front inputs are required for this mode. </summary>
Friend Property Get FrontReadRequired() As Boolean
    FrontReadRequired = Me.FrontRearToggleButton.value
End Property

''' <summary>   Gets the external trigger status. </summary>
Friend Property Get ExternalTrigEnabled() As Boolean
    ExternalTrigEnabled = Me.ExtTrigToggleButton.value
End Property

''' <summary>   Enables or disables external triggering. </summary>
Friend Property Let ExternalTrigEnabled(ByVal value As Boolean)
    Me.ExtTrigToggleButton.value = value
End Property

''' <summary>   Gets the auto increment resistance number status. </summary>
Friend Property Get AutoIncrementResistanceNoEnabled() As Boolean
    AutoIncrementResistanceNoEnabled = Me.AutoIncrementToggleButton.value
End Property

''' <summary>   Gets the single read enabled. </summary>
Friend Property Get SingeReadEnabled() As Boolean
    SingeReadEnabled = Me.SingleAllToggleButton.value
End Property

''' <summary>   Gets the resistance number. </summary>
Friend Property Get SelectedResistanceNumber() As Integer
    SelectedResistanceNumber = CInt(Range(m_resistanceNoCell).value)
End Property

''' <summary>   Sets the resistance display. </summary>
''' <param name="resistanceNo">      [single] the resistance number. </param>
''' <param name="resistanceValue">   [double] the resistance. </param>
Private Sub SetResistance(ByVal resistanceNo As Integer, ByVal resistanceValue As Double)

    Range(m_resistanceNoCell).value = resistanceNo
    Range(m_resistanceCell).value = resistanceValue
    
    Dim row As Integer: row = m_resistanceNoTopRow + resistanceNo - 1
    Range(m_resistanceNoCol & CStr(row)).value = "R" & CStr(resistanceNo)
    Range(m_resistanceCol & CStr(row)).value = Format(resistanceValue, m_resistanceFormat)
    
End Sub

Private Sub FrontRearToggleButton_Click()
    FrontRearToggleButton.Caption = IIf(FrontRearToggleButton.value, "Front", "Rear")
End Sub

Private Sub AutoIncrementToggleButton_Click()
    AutoIncrementToggleButton.Caption = IIf(AutoIncrementToggleButton.value, "Auto", "Manual")
End Sub

Private Sub SingleAllToggleButton_Click()
    SingleAllToggleButton.Caption = IIf(SingleAllToggleButton.value, "Single", "All")
End Sub

''' <summary> Queries the Inputs status. </summary>
Private Sub QueryInputsButton_Click()

    Const thisProcedureName = "QueryInputsButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_inputsCell).value = IIf(Me.K2700.IsFrontInputs, "Front", "Rear")
    
    Me.ReadStopWatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary> </summary>
Private Sub ReadCardsButton_Click()
    
    Const thisProcedureName = "ReadCardsButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.K2700.ReadCards
    
    Me.ReadStopWatch
    
    Range(m_card1Cell).value = Me.K2700.Card1
    Range(m_card2Cell).value = Me.K2700.Card2

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Read a single or all resistances from the front or read panel. </summary>
Private Sub ReadResistanceButton_Click()
    
    Const thisProcedureName = "ReadResistanceButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReadResistanceButton.Enabled = False
    Me.ExtTrigToggleButton.Enabled = False
    
    ' reinitialize if device not defined
    If Not Me.K2700.Connected Then
        Me.K2700.Initialize Me.host, Me.port, Me.ReceiveTimeout
    End If
    
    If Me.FrontReadRequired And Not Me.K2700.IsFrontInputs Then
        Range(m_errorMessageCell).value = "Please release the front panel INPUTS button to its FRONT (F) position"
        GoTo exit_Handler
    ElseIf Not Me.FrontReadRequired And Me.K2700.IsFrontInputs Then
        Range(m_errorMessageCell).value = "Please depress the front panel INPUTS button to its REAR (R) position"
        GoTo exit_Handler
    End If

    ' configure measuring a single resistance on each read.
    Me.K2700.ConfigureSingleResistanceReading

    Dim resistanceNo As Integer

    If Me.AutoIncrementResistanceNoEnabled Then
        
        ' Clear all cells
        Me.ClearReadings 0

        ' scan all the resistance
        For resistanceNo = 1 To Me.ResistanceCount
        
            ' allow events to occur
            DoEvents
            
            ' set the current resistance
            m_currentResistanceNo = resistanceNo
            
            Me.RestartStopWatch
            If Me.FrontReadRequired Then
                Me.K2700.ReadFrontResistance resistanceNo
            Else
                Me.K2700.ReadRearResistance resistanceNo
            End If
            Me.ReadStopWatch
        
        Next resistanceNo

    Else
        resistanceNo = Me.SelectedResistanceNumber
        Me.RestartStopWatch
        If Me.FrontReadRequired Then
            Me.K2700.ReadFrontResistance resistanceNo
        Else
            Me.K2700.ReadRearResistance resistanceNo
        End If
        ' Range("C22").value = Me.K2700.ReadRearResistance(ResistanceNumber)
        Me.ReadStopWatch
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.ReadResistanceButton.Enabled = True
    Me.ExtTrigToggleButton.Enabled = True
    
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Sets and displays the scan list for the cards. </summary>
Private Sub SetScansButton_Click()

    Const thisProcedureName = "SetScansButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.K2700.DefineResistanceScanLists
    
    Me.ReadStopWatch
    
    Range(m_card1ScanListCell).value = Me.K2700.Card1ScanList
    Range(m_card2ScanListCell).value = Me.K2700.Card2ScanList
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   clears the data acquisition readings. </summary>
''' <param name="offsetResistance">    The offset resistance value to use for clearing the readings. </param>
Public Sub ClearReadings(ByVal offsetResistance As Single)

    Dim ResistanceNumber As Integer
    For ResistanceNumber = 1 To Me.ResistanceCount
        SetResistance ResistanceNumber, offsetResistance
    Next ResistanceNumber
  
End Sub

''' <summary> Sets the instrument and program to get data using external trigger. </summary>
''' <param name="timerInterval"> The polling timer interval. </param>
Private Sub SetExternalTriggerMode(Optional ByVal timerInterval As Integer = 100)

    Const thisProcedureName = "SetExternalTriggerMode"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' check if doing an external trigger reading with or without auto increment
    If ExternalTrigEnabled Then
  
        ' clear the triggered flag
        m_extTrigInitiated = False
        
        ' reinitialize if device not defined
        If Not Me.K2700.Connected Then
            Me.K2700.Initialize Me.host, Me.port, Me.ReceiveTimeout
        End If
        
        ' make sure that the instrument is ready to accept
        ' input from the front panel
     
        ' check if front or read
        If Not Me.K2700.IsFrontInputs Then
        
            ' if rear, toggle back the button
            Me.ExternalTrigEnabled = False
            
            ' display the proper message
            Range(m_errorMessageCell).value = "Please release the front panel INPUTS button to its FRONT (F) position"
              
            ' and get out
            GoTo exit_Handler
      
        End If
  
        Range(m_externalTriggerStatusCell) = "External Trigger On"

        ' set the instrument for externally triggered measurements
        Me.K2700.SetExternalTriggerResistanceMode
    
        If Me.AutoIncrementResistanceNoEnabled Then
        
            ' set the first resistance for manual scan
            m_currentResistanceNo = 1
        
        Else
        
            ' set the single resistance for manual scan
            m_currentResistanceNo = Me.SelectedResistanceNumber
        
        End If
       
        ' start the timer
        Me.StartTimer timerInterval
        
    Else
  
        ' once external triggering is turned off, the timer is allowed to
        ' to turn itself off
  
    End If
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the external trigger toggle button. </summary>
Private Sub ExtTrigToggleButton_Click()

    Const thisProcedureName = "ExtTrigToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Me.ExternalTrigEnabled Then
    
        Me.ExtTrigToggleButton.Enabled = False
        Me.ReadResistanceButton.Enabled = False
        
        If Me.AutoIncrementResistanceNoEnabled Then
            
            ' Clear all readings
            Me.ClearReadings 0
        
        End If
    
        Range(m_externalTriggerStatusCell) = "External Trigger On"
        
        ' set triggered mode
        SetExternalTriggerMode Me.TimerInerval
    
    Else
    
      ' one external triggering is turned off, the timer
      ' is allowed to turn itself off
    
    End If
 
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' stop the timer if the timer started.
    Me.PauseTimer
  
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' TIMER CONTROL
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the external trigger monitoring timer event using the
''' device to handle the timing events. </summary>
''' <remarks>
''' Use this method to check if turn on or off manual triggering mode of the instrument. <para>
''' 10/16/00  David Hary  1.00.00 Add a beep </para>
''' 2023-06-23 David Hary 1.0.8574 Use the device event handler to update the resistances.
''' </remarks>
Private Sub HandleTimerEvent()

    Const thisProcedureName = "HandleTimerEvent"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If m_pauseRequested Or m_handlingTimerEvent Then Exit Sub
   
    ' prevent reenty
    m_handlingTimerEvent = True
    
    Dim resistanceNo As Integer
    
    ' check if scan or single
    If AutoIncrementResistanceNoEnabled Then
      
        ' select the current resistance
        resistanceNo = m_currentResistanceNo
          
    Else
    
        ' get the cell number
        resistanceNo = Me.SelectedResistanceNumber
    
    End If
      
    ' allow other events to occur
    DoEvents
    
    Dim isDone As Boolean
    isDone = Not Me.ExternalTrigEnabled
    
    ' stop the timer.
    If isDone Then Me.PauseTimer
    
    Me.K2700.MonitorExternalTrigger resistanceNo, isDone
    
    ' if we ended manual mode, enable the rest of the buttons
    If isDone Then
    
        ' EnableButtons
        Range(m_externalTriggerStatusCell) = "External Trigger off"
        
        Me.ExtTrigToggleButton.Enabled = True
        Me.ReadResistanceButton.Enabled = True
        
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release reenty prevention
    m_handlingTimerEvent = False

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' stop the timer.
    Me.PauseTimer
    
    ' enable all buttons
    Me.ExtTrigToggleButton.Enabled = True
    Me.ReadResistanceButton.Enabled = True
    
    ' append the error source
    UserDefinedErrors.SetErrSource thisProcedureName, Me.name
    
    ' display the error message
    Range(m_errorMessageCell).value = UserDefinedErrors.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Starts the event timer. </summary>
''' <param name="timerInterval">   Specifies the timer interval in milliseconds. </param>
Friend Sub StartTimer(Optional ByVal timerInterval As Long = 100)
    m_handlingTimerEvent = False
    m_pauseRequested = False
    If m_eventTimer Is Nothing Then
        Set m_eventTimer = cc_isr_Core.Constructor.CreateStdTimer(timerInterval)
        m_eventTimer.StartTimer
    Else
        m_eventTimer.ResumeTimer
    End If
End Sub

''' <summary>   Requests pausing the timer. </summary>
Friend Sub PauseTimer()
    m_pauseRequested = True
End Sub

''' <summary>   Aborts the timer. </summary>
Friend Sub AbortTimer()
  ' we need to figure out how to recreate the timer
  m_eventTimer.StopTimer
End Sub

''' <summary>   Handles a timer tick event. </summary>
Private Sub m_eventTimer_Tick()
    On Error Resume Next
    If m_pauseRequested Then
        m_eventTimer.PauseTimer
    Else
        HandleTimerEvent
    End If
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the K2700 <see cref="cc_isr_VI.ResistanceNumberMeasured"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_VI.ResistanceNumberEventArgs"/> event arguments. </param>
Private Sub m_k2700_ResistanceNumberMeasured(ByVal eventArgs As cc_isr_VI.ResistanceNumberEventArgs)

    SetResistance eventArgs.ResistanceNumber, eventArgs.Resistance
    
    If Me.AutoIncrementResistanceNoEnabled Then
    
        ' increment the resistance for the next reading
        m_currentResistanceNo = eventArgs.ResistanceNumber + 1
        If m_currentResistanceNo > Me.ResistanceCount Then
            m_currentResistanceNo = 1
        End If
    
    End If
    

End Sub

''' <summary>   Handles the K2700 <see cref="cc_isr_VI.K2700.ConnectionChanged"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_k2700_ConnectionChanged(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Me.Initialize eventArgs

    If eventArgs.Connected Then
    
        ' report the socket number
        Range(m_socketAddressCell).value = Me.Session.VISession.socketAddress
        Range(m_socketIdCell).value = Me.Session.SocketId
        
        Me.Session.VISession.GpibLan.SelectiveDeviceClear
        Me.Session.ResetKnownState
        Me.Session.ClearExecutionState
        
        Me.ListenTalkToggleButton.Caption = IIf(Me.Session.VISession.GpibLan.ReadAfterWriteEnabledGetter, "Talk", "Listen")
    
    Else
    
        Range(m_socketAddressCell).value = ""
        Range(m_socketIdCell).value = ""
        
    End If

End Sub

''' <summary>   Handles the <see cref="cc_isr_VI.K2700.ConnectionChanging"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_k2700_ConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

End Sub






VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700SheetBroken"
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
Private Const m_connectedCell As String = "F2"
Private Const m_identityCell As String = "I2"
Private Const m_card1Cell As String = "C15"
Private Const m_card2Cell As String = "C16"
Private Const m_card1ScanListCell As String = "C17"
Private Const m_card2ScanListCell As String = "C18"
Private Const m_inputsCell As String = "C20"
Private Const m_resistanceNoCell As String = "C21"
Private Const m_resistanceCell As String = "D21"
Private Const m_maxResistanceCount As Integer = 20

Private Const m_resistanceNoTopRow As Integer = 2
Private Const m_resistanceNoCol As String = "K"
Private Const m_resistanceCol As String = "L"
Private Const m_resistanceFormat As String = "0.00"

' defines the standard event timer.
Private WithEvents m_eventTimer As cc_isr_Core.stdTimer
Attribute m_eventTimer.VB_VarHelpID = -1
Private m_pauseRequested As Boolean

Private m_extTrigInitiated  As Boolean
Private m_currentResistanceNo As Integer
Private WithEvents m_device As K2700
Private m_session As Ieee488Session
Private m_sw As StopWatch

Friend Sub Worksheet_Activate()
    Set m_device = Constructor.CreateK2700()
    Set m_session = m_device.Session
    Set m_sw = cc_isr_Core.Constructor.CreateStopWatch()
    ' this toggles the enabled status of the controls.
    Me.ConnectedCheckBox_Click
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
    Me.CommandsComboBox.AddItem "SYST:FRSW?"
   
End Sub

Private Sub Worksheet_Deactivate()
    Set m_sw = Nothing
    Set m_session = Nothing
    Set m_device = Nothing
End Sub

Friend Sub RestartStopWatch()
    Range(m_elapsedTimeCell).value = ""
    If m_sw Is Nothing Then
        Me.Worksheet_Activate
    End If
    m_sw.Restart
End Sub

Friend Sub ReadStopWatch()
    Range(m_elapsedTimeCell).value = Format(m_sw.ElapsedMilliseconds, "0.0")
End Sub

Private Sub CLSButton_Click()
    Me.RestartStopWatch
    m_session.ClearExecutionState
    Me.ReadStopWatch
End Sub

Friend Sub ConnectedCheckBox_Click()
    
    Dim affirmative As Boolean
    
    If m_device Is Nothing Then
        affirmative = False
    Else
        affirmative = m_device.Connected
    End If
    
    Me.ConnectedCheckBox.value = affirmative
    
    ' report the connection state
    Range(m_connectedCell).value = affirmative
        
    Me.QueryIdentityButton.Enabled = affirmative
    Me.RSTButton.Enabled = affirmative
    Me.SDCButton.Enabled = affirmative
    Me.CLSButton.Enabled = affirmative
    Me.ReadCardsButton.Enabled = affirmative
    Me.QueryButton.Enabled = affirmative
    Me.ReadButton.Enabled = affirmative
    Me.ReadStatusByteButton.Enabled = affirmative
    Me.WriteButton.Enabled = affirmative
    Me.SetScansButton.Enabled = affirmative
    Me.QueryInputsButton.Enabled = affirmative
    Me.ReadRearResistanceButton.Enabled = affirmative
    
End Sub

''' <summary>   Queries the device Identity string. </summary>
Private Sub QueryIdentityButton_Click()
    Me.RestartStopWatch
    Range(m_identityCell).value = m_session.QueryIdentity
    Me.ReadStopWatch
End Sub

''' <summary> Queries the INputs status. </summary>
Private Sub QueryInputsButton_Click()
    Me.RestartStopWatch
    Range(m_inputsCell).value = IIf(m_device.IsFrontInputs, "Front", "Rear")
    Me.ReadStopWatch
End Sub

''' <summary> </summary>
Private Sub ReadCardsButton_Click()
    Me.RestartStopWatch
    m_device.ReadCards
    Me.ReadStopWatch
    Range(m_card1Cell).value = m_device.Card1
    Range(m_card2Cell).value = m_device.Card2
End Sub

''' <summary> </summary>
Private Sub RSTButton_Click()
    Me.RestartStopWatch
    m_session.ResetKnownState
    Me.ReadStopWatch
End Sub

''' <summary> </summary>
Private Sub ReadRearResistanceButton_Click()
    Dim ResistanceNumber As Integer: ResistanceNumber = Range("C21").value
    Me.RestartStopWatch
    Range("C22").value = m_device.ReadRearResistance(ResistanceNumber)
    Me.ReadStopWatch
End Sub


''' <summary> </summary>
Private Sub SDCButton_Click()
    Me.RestartStopWatch
    m_session.SelectiveDeviceClear
    Me.ReadStopWatch
End Sub

''' <summary> </summary>
Private Sub SetScansButton_Click()
    Me.RestartStopWatch
    m_device.DefineResistanceScanLists
    Me.ReadStopWatch
    Range(m_card1ScanListCell).value = m_device.Card1ScanList
    Range(m_card2ScanListCell).value = m_device.Card2ScanList
End Sub

''' <summary> </summary>
Private Sub ToggleConnectionButton_Click()

    Me.RestartStopWatch

    Dim Host As String
    Dim Port As Integer
    Dim timeout As Integer
    
    timeout = Me.ReceiveTimeout
    Host = Me.Host
    Port = Me.Port
    
    If m_device.Connected Then
        
        m_device.CloseConnection
        Set m_session = m_device.Session
        
        Range(m_connectedHostCell).value = ""
        Range(m_connectedPortCell).value = ""
        Range(m_connectedSocketIdCell).value = ""
        
    Else
    
        Range(m_connectedHostCell).value = Host
        Range(m_connectedPortCell).value = Port
        
        m_device.OpenConnection Host, Port
        Set m_session = m_device.Session
        
        ' report the socket number
        Range(m_connectedSocketIdCell).value = m_session.ViSession.SocketId
    End If
    
    Me.ReadStopWatch
    Me.ConnectedCheckBox_Click
    
End Sub

''' <summary> </summary>
Private Sub WriteButton_Click()
    Range(m_sentMessageCell).value = ""
    Dim msg As String: msg = Me.CommandsComboBox.value
    Me.RestartStopWatch
    m_session.WriteLine msg
    Me.ReadStopWatch
    Range(m_sentMessageCell).value = msg
    If Me.AutoStatusReadCheckBox.value Then
        Me.ReadStatusByteButton_Click
    End If
End Sub

''' <summary> </summary>
Private Sub ReadButton_Click()
    Range(m_receivedMessageCell).value = ""
    Me.RestartStopWatch
    Dim msg As String: msg = m_session.Read
    Me.ReadStopWatch
    Range(m_receivedMessageLengthCell).value = msg
    Range(m_receivedMessageCell).value = Len(msg)
End Sub

''' <summary> </summary>
Friend Sub ReadStatusByteButton_Click()
    Me.RestartStopWatch
    Range(m_statusByteCell).value = m_session.ReadStatusByte
    Me.ReadStopWatch
End Sub

''' <summary> </summary>
Private Sub QueryButton_Click()
    Me.SentMessage = ""
    Dim msg As String: msg = Me.CommandsComboBox.value
    Me.RestartStopWatch
    Me.ReceivedMessage = m_session.QueryLine(msg)
    Me.ReadStopWatch
    Me.SentMessage = msg
End Sub

''' <summary>   Performs a single scan on the given resistance from the rear panel. </summary>
''' <remarks>
''' Use this method to get a single value from the instrument using the back panel switches
''' </remarks>
''' <param name="resistanceNo">   The resistance number. </param>
Private Sub ReadRearResistance(ByVal resistanceNo As Integer)

    Const thisProcedureName = "ExecuteReadScan"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' reinitialize if device not connected
    If Not m_device.Connected Then
        m_device.Initialize Me.Host, Me.Port, Me.ReceiveTimeout
    End If
  
    ' check if we have a valid number
    If resistanceNo > 0 Then
  
        ' query switch mode
        
        If m_device.IsFrontInputs Then
            MsgBox "Please depress the front panel INPUTS button to its READ (R) position"
            GoTo exit_Handler
        End If
        
        Dim resistanceValue As Double: resistanceValue = m_device.ReadRearResistance(resistanceNo)
        
        ' the reistance value might be negative if the resistance number is
        ' out of range, that is no scan card is available for this resistance.
        
        If resistanceValue >= 0 Then
        
            ' add the measured resistance to the data sheet
            SetResistance resistanceNo, resistanceValue
        
        End If
        
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   clears the data acquisition readings. </summary>
''' <param name="offsetResistance">    The offset resistance value to use for clearing the readings. </param>
Public Sub ClearReadings(ByVal offsetResistance As Single)

    Dim ResistanceNumber As Integer
    For ResistanceNumber = 1 To m_maxResistanceCount
        SetResistance ResistanceNumber, offsetResistance
    Next ResistanceNumber
  
End Sub


''' <summary> Handles the click event of the Auto Scan button.
''' Starts a scan of 48 resistances. </summary>
''' <remarks> 10/16/00 Clear all values and set time and date. </remarks>
Private Sub cmdAutoScan_Click()

    Const thisProcedureName = "cmdAutoScan_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' reinitialize if device not defined
    If Not m_device.Connected Then
        m_device.Initialize Me.Host, Me.Port, Me.ReceiveTimeout
    End If
  
    ' disable all buttons
    ' Me.' cmdAutoScan.Enabled = False
    ' Me.' cmdAutoSingle.Enabled = False
    ' Me.toggleManualScan.Enabled = False
    ' Me.toggleManualSingle.Enabled = False
    
    ' Clear all cells
    Me.ClearReadings 0
    
    Dim resistanceNo As Integer
   
    ' scan all the resistance
    For resistanceNo = 1 To m_maxResistanceCount
  
        ' allow events to occur
        DoEvents
        
        ' set the current resistance
        m_currentResistanceNo = resistanceNo
        
        ' get a single scan
        ReadRearResistance resistanceNo
    
    Next resistanceNo
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' enable all buttons
    ' EnableButtons
    
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' enable all buttons
    ' EnableButtons
    
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Reads a single value from the instrument using the back panel scan cards. </summary>
Private Sub cmdAutoSingle_Click()

    Const thisProcedureName = "cmdAutoSingle_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' reinitialize if device not defined
    If Not m_device.Connected Then
        m_device.Initialize Me.Host, Me.Port, Me.ReceiveTimeout
    End If
    
    ' get the resistance number that corresponds to the active cell
    
    ' get a single scan
    ReadRearResistance SelectedResistanceNumber
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' enable all buttons
    ' EnableButtons
    
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary> Sets the instrument and program to get data using external trigger. </summary>
''' <param name="timerInterval"> The polling timer interval. </param>
Private Sub SetExternalTriggerMode(Optional ByVal timerInterval As Integer = 100)

    Const thisProcedureName = "SetExternalTriggerMode"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim sData As String
    Dim iRno As Integer
    
    ' check if doing an external trigger reading with or without auto increment
    If ExternalTrigEnabled Then
  
        ' clear the triggered flag
        m_extTrigInitiated = False
        
        ' reinitialize if device not defined
        If Not m_device.Connected Then
            m_device.Initialize Me.Host, Me.Port, Me.ReceiveTimeout
        End If
        
        ' make sure that the instrument is ready to accept
        ' input from the front panel
     
        ' check if front or read
        If Not m_device.IsFrontInputs Then
        
            ' if rear, toggle back the button
            ToggleManualSingle = False
            
            ' display the proper message
            MsgBox "Please release the front panel INPUTS button to its FRONT (F) position"
              
            ' and get out
            GoTo exit_Handler
      
        End If
  
        Range("C4") = "Extrent Trigger On"

        ' set the instrument for externally triggered measurements
        m_device.SetExternalTriggerResistanceMode
    
        ' set the first resistance for manual scan
        m_currentResistanceNo = 1
       
        ' start the timer
        StartTimer timerInterval
        
        ' now set the timer in motion
        ' motmrAction.Interval = 100
        ' motmrAction.Enabled = True
  
    ElseIf Not ToggleManualSingle Then
  
    ' one we turn this guy off, we allow the timer
    ' to turn itself off
  
    End If
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the toggle manual scan click event.
''' Turns on external trigger mode with reading from the front pannel
''' and with resistance numbers automatically incrementing after each
''' external trigger.
''' </summary>
''' <remarks>
''' 10/16/00  David Hary  1.00.00 Clear all values and set time and date
''' </remarks>
Private Sub toggleManualScan_Click()

    Const thisProcedureName = "toggleManualScan_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim sData As String
    Dim ResistanceNumber As Integer
    
    If ToggleManualScan Then
    
        ' Clear all readings
        Me.ClearReadings 0
       
        Range("C4") = "Extrent Trigger On"
        
        ' if we start manual mode, disable the rest of the buttons
        ' Me.' cmdAutoSingle.Enabled = False
        ' Me.' cmdAutoScan.Enabled = False
        ' Me.toggleManualSingle.Enabled = False
        
        ' set triggered mode
        SetExternalTriggerMode
    
    ElseIf Not ToggleManualScan Then
    
      ' one we turn this guy off, we allow the timer
      ' to turn itself off
    
    End If
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' stop the timer if the timer started.
    PauseTimer
  
    ' enable all buttons
    ' EnableButtons
    
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary> Handles the Single toggle button click event.
''' This toggle manual triggering mode of the instrument.
''' </summary>
Private Sub toggleManualSingle_Click()

    Const thisProcedureName = "toggleManualSingle_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim ResistanceNumber As Integer
    
    If ToggleManualSingle Then
  
        Range("C4") = "Extrent Trigger On"

        ' if we start manual mode, disable the rest of the buttons
        ' cmdAutoSingle.Enabled = False
        ' cmdAutoScan.Enabled = False
        
        ' Me.toggleManualScan.Enabled = False
        
        ' set triggered mode
        SetExternalTriggerMode
  
    ElseIf Not ToggleManualSingle Then
  
        ' once we turn this guy off, we allow the timer
        ' to turn itself off
       
        
    End If
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' stop the timer if the timer started.
    PauseTimer
  
    ' enable all buttons
    ' EnableButtons
    
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the external trigger monitoring timer event using the
''' device to handle the timing events. </summary>
''' <remarks>
''' Use this method to check if turn on or off manual triggering mode of the instrument. <para>
''' 10/16/00  David Hary  1.00.00 Add a beep </para>
''' 2023-06-23 David Hary 1.0.8574 Use the device revent to update the resistances.
''' </remarks>
Private Sub HandleTimerEvent()

    Const thisProcedureName = "HandleTimerEvent"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If m_pauseRequested Then Exit Sub
   
    Dim sData As String
    Dim resistanceNo As Integer
      
    ' check if scan or single
    If AutoIncrementResistanceNoEnabled Then
      
        ' select the current resistance
        resistanceNo = m_currentResistanceNo
          
    ElseIf ToggleManualScan Then
    
        ' get the cell number
        resistanceNo = Me.SelectedResistanceNumber
    
    End If
      
    ' allow other events to occur
    DoEvents
    
    Dim isDone As Boolean
    isDone = Not Me.ExternalTrigEnabled
    
    ' stop the timer.
    If isDone Then PauseTimer
    
    m_device.MonitorExternalTrigger resistanceNo, isDone
    
    ' if we ended manual mode, enable the rest of the buttons
    If isDone Then
        ' EnableButtons
        Range("C4") = "Extrent Trigger off"
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' stop the timer.
    Me.PauseTimer
    
    ' enable all buttons
    ' EnableButtons
    
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Gets the host dotted IP Address. </summary>
Friend Property Get Host() As String
    Host = Range(m_inputHostCell).value
End Property

''' <summary>   Gets the port number. </summary>
Friend Property Get Port() As Integer
    Port = CInt(Range(m_inputPortCell).value)
End Property

''' <summary> Get the receive timeout. </summary>
Friend Property Get ReceiveTimeout() As Integer
    ReceiveTimeout = CInt(Range(m_receiveTimeoutCell).value)
End Property

''' <summary> </summary>
Friend Property Get ReceivedMessage() As String
    ReceivedMessage = Range(m_receivedMessageCell).value
End Property

''' <summary> </summary>
Friend Property Let ReceivedMessage(ByVal value As String)
    Range(m_receivedMessageCell).value = value
    Range(m_receivedMessageLengthCell).value = Len(value)
End Property

''' <summary> </summary>
Friend Property Get SentMessage() As String
    SentMessage = Range(m_sentMessageCell).value
End Property

''' <summary> </summary>
Friend Property Let SentMessage(ByVal value As String)
    Range(m_sentMessageCell).value = value
End Property
Friend Property Get ExternalTrigEnabled() As Boolean
    ExternalTrigEnabled = True
End Property
Friend Property Get AutoIncrementResistanceNoEnabled() As Boolean
    AutoIncrementResistanceNoEnabled = True
End Property

Friend Property Get ToggleManualSingle() As Boolean
    ToggleManualSingle = True ' ToggleManualSingleButton.Value
End Property
Friend Property Let ToggleManualSingle(ByVal value As Boolean)
    ' ToggleManualSingleButton.Value = Value
End Property

Friend Property Get ToggleManualScan() As Boolean
    ToggleManualScan = True ' ToggleManualScanButton.Value
End Property

Friend Property Let ToggleManualScan(ByVal value As Boolean)
    ' ToggleManualScanButton.Value = Value
End Property

''' <summary>   Starts the event timer. </summary>
''' <param name="timerInterval">   Specifies the timer interval in milliseconds. </param>
Friend Sub StartTimer(Optional ByVal timerInterval As Long = 100)
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
    
    Dim row As Integer: row = m_maxResistanceCount + resistanceNo - 1
    Range(m_resistanceNoCol & CStr(row)).value = "R" & CStr(resistanceNo)
    Range(m_resistanceCol & CStr(row)).value = Format(resistanceValue, m_resistanceFormat)
    
End Sub

''' <summary>   Handles the K2700 <see cref="cc_isr_VI.ResistanceNumberMeasured"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_VI.ResistanceNumberEventArgs"/> event arguments. </param>
Private Sub m_k2700_ResistanceNumberMeasured(ByRef eventArgs As cc_isr_VI.ResistanceNumberEventArgs)

    SetResistance eventArgs.ResistanceNumber, eventArgs.resistance
    
    ' increment the resistnace for the next reading
    m_currentResistanceNo = m_currentResistanceNo + 1
    If m_currentResistanceNo > m_maxResistanceCount Then
        m_currentResistanceNo = 1
    End If

End Sub

Private Sub m_device_ConnectionChanged(ByRef eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

End Sub



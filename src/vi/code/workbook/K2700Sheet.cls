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

Private Const inputRepeatCountCell As String = "B1"
Private Const inputHostCell As String = "B2"
Private Const inputPortCell As String = "B3"
Private Const connectedHostCell As String = "C2"
Private Const connectedPortCell As String = "D2"
Private Const connectedSocketIdCell As String = "E2"
Private Const sentMessageCell As String = "G2"
Private Const receivedMessageLengthCell As String = "H2"
Private Const receivedMessageCell As String = "I2"
Private Const statusByteCell As String = "G3"
Private Const elapsedTimeCell As String = "J2"
Private Const connectedCell As String = "F2"
Private Const identityCell As String = "I2"
Private Const card1Cell As String = "C15"
Private Const card2Cell As String = "C16"
Private Const card1ScanListCell As String = "C17"
Private Const card2ScanListCell As String = "C18"
Private Const inputsCell As String = "C20"
Private Const resistanceNoRow As Integer = 21
Private Const resistanceNoCol As String = "C"
Private Const resistanceNoCell As String = "C21"
Private Const resistanceCol As String = "D"
Private Const resistanceCell As String = "D21"
Private Const maxResistanceCount As Integer = 20

' defines the standard event timer.
Private WithEvents m_eventTimer As cc_isr_Core.stdTimer
Attribute m_eventTimer.VB_VarHelpID = -1

Private WithEvents m_device As K2700
Private m_session As Ieee488Session
Private sw As StopWatch

Friend Sub Worksheet_Activate()
    Set m_device = Constructor.CreateK2700()
    Set m_session = m_device.Session
    Set sw = cc_isr_Core.Constructor.CreateStopWatch()
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
    Set sw = Nothing
    Set m_session = Nothing
    Set m_device = Nothing
End Sub

Friend Sub RestartStopWatch()
    Range(elapsedTimeCell).value = ""
    If sw Is Nothing Then
        Me.Worksheet_Activate
    End If
    sw.Restart
End Sub

Friend Sub ReadStopWatch()
    Range(elapsedTimeCell).value = Format(sw.ElapsedMilliseconds, "0.0")
End Sub

Private Sub CLSButton_Click()
    Me.RestartStopWatch
    m_session.ClearExecutionState
    Me.ReadStopWatch
End Sub

Friend Sub ConnectedCheckBox_Click()
    
    Me.ConnectedCheckBox.value = m_device.Connected
    
    ' report the connection state
    Range(connectedCell).value = m_device.Connected
        
    Me.QueryIdentityButton.Enabled = m_device.Connected
    Me.RSTButton.Enabled = m_device.Connected
    Me.SDCButton.Enabled = m_device.Connected
    Me.CLSButton.Enabled = m_device.Connected
    Me.ReadCardsButton.Enabled = m_device.Connected
    Me.QueryButton.Enabled = m_device.Connected
    Me.ReadButton.Enabled = m_device.Connected
    Me.ReadStatusByteButton.Enabled = m_device.Connected
    Me.WriteButton.Enabled = m_device.Connected
    Me.SetScansButton.Enabled = m_device.Connected
    Me.QueryInputsButton.Enabled = m_device.Connected
    Me.ScanReadResistanceButton.Enabled = m_device.Connected
    
End Sub

''' <summary>   Queries the device Identity string. </summary>
Private Sub QueryIdentityButton_Click()
    Me.RestartStopWatch
    Range(identityCell).value = m_session.QueryIdentity
    Me.ReadStopWatch
End Sub

''' <summary> Queries the INputs status. </summary>
Private Sub QueryInputsButton_Click()
    Me.RestartStopWatch
    Range(inputsCell).value = IIf(m_device.IsFrontInputs, "Front", "Rear")
    Me.ReadStopWatch
End Sub

''' <summary> </summary>
Private Sub ReadCardsButton_Click()
    Me.RestartStopWatch
    m_device.ReadCards
    Me.ReadStopWatch
    Range(card1Cell).value = m_device.Card1
    Range(card2Cell).value = m_device.Card2
End Sub

''' <summary> </summary>
Private Sub RSTButton_Click()
    Me.RestartStopWatch
    m_session.ResetKnownState
    Me.ReadStopWatch
End Sub

''' <summary> </summary>
Private Sub ScanReadResistanceButton_Click()
    Dim ResistanceNumber As Integer: ResistanceNumber = Range("C21").value
    Me.RestartStopWatch
    Range("C22").value = m_device.ScanReadResistance(ResistanceNumber)
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
    Range(card1ScanListCell).value = m_device.Card1ScanList
    Range(card2ScanListCell).value = m_device.Card2ScanList
End Sub

''' <summary> </summary>
Private Sub ToggleConnectionButton_Click()

    Me.RestartStopWatch

    Dim host As String
    Dim port As Integer
    Dim repeatCount As Integer
    
    repeatCount = Range(inputRepeatCountCell).value
    host = Range(inputHostCell).value
    port = Range(inputPortCell).value
    
    If m_device.Connected Then
        
        m_device.CloseConnection
        Set m_session = m_device.Session
        
        Range(connectedHostCell).value = ""
        Range(connectedPortCell).value = ""
        Range(connectedSocketIdCell).value = ""
        
    Else
    
        Range(connectedHostCell).value = host
        Range(connectedPortCell).value = port
        
        m_device.OpenConnection host, port
        Set m_session = m_device.Session
        
        ' report the socket number
        Range(connectedSocketIdCell).value = m_session.ViSession.SocketId
    End If
    
    Me.ReadStopWatch
    Me.ConnectedCheckBox_Click
    
End Sub

''' <summary> </summary>
Private Sub WriteButton_Click()
    Range(sentMessageCell).value = ""
    Dim msg As String: msg = Me.CommandsComboBox.value
    Me.RestartStopWatch
    m_session.WriteLine msg
    Me.ReadStopWatch
    Range(sentMessageCell).value = msg
    If Me.AutoStatusReadCheckBox.value Then
        Me.ReadStatusByteButton_Click
    End If
End Sub

''' <summary> </summary>
Private Sub ReadButton_Click()
    Range(receivedMessageCell).value = ""
    Me.RestartStopWatch
    Dim msg As String: msg = m_sesssion.Read
    Me.ReadStopWatch
    Range(receivedMessageLengthCell).value = msg
    Range(receivedMessageCell).value = Len(msg)
End Sub

''' <summary> </summary>
Friend Sub ReadStatusByteButton_Click()
    Me.RestartStopWatch
    Range(statusByteCell).value = m_session.ReadStatusByte
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
Private Sub ExecuteReadScan(ByVal resistanceNo As Integer)

    Const thisProcedureName = "ExecuteReadScan"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' reinitialize if device not connected
    If Not K2700.Connected Then
        K2700.Initialize sheetDataAcq.host, sheetDataAcq.port, sheetDataAcq.ReceiveTimeout
    End If
  
    ' check if we have a valid number
    If resistanceNo > 0 Then
  
        ' query switch mode
        
        If Me.K2700.IsFrontInputs Then
            MsgBox "Please depress the front panel INPUTS button to its READ (R) position"
            GoTo exit_Handler
        End If
        
        Dim resistanceValue As Double: resistanceValue = K2700.ScanReadResistance(resistanceNo)
        
        ' the reistance value might be negative if the resistance number is
        ' out of range, that is no scan card is available for this resistance.
        
        If resistanceValue >= 0 Then
        
            ' add the measured resistance to the data sheet
            sheetDataAcq.SetResistance resistanceValue
        
        End If
        
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' enable all buttons
    EnableButtons
    
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary> Handles the click event of the Auto Scan button.
''' Starts a scan of 48 resistances. </summary>
''' <remarks> 10/16/00 Clear all values and set time and date. </remarks>
Private Sub cmdAutoScan_Click()

    Const thisProcedureName = "cmdAutoScan_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' reinitialize if device not defined
    If Not K2700.Connected Then
        K2700.Initialize sheetDataAcq.host, sheetDataAcq.port, sheetDataAcq.ReceiveTimeout
    End If
  
    ' disable all buttons
    Me.cmdAutoScan.Enabled = False
    Me.cmdAutoSingle.Enabled = False
    Me.toggleManualScan.Enabled = False
    Me.toggleManualSingle.Enabled = False
    
    ' Clear all cells
    sheetDataAcq.ClearReadings Range(slopeCell).value * Range(offsetCell).value
    
    ' Set the data and time
    Range(dateCell).value = Now
    Range(timeCell).value = Format$(Now, hourMinuteTimeFormat)
    
    Dim resistanceNo As Integer
   
    ' scan all the resistance
    For resistanceNo = 1 To sheetDataAcq.maxResistanceCount
  
        ' allow events to occur
        DoEvents
        
        ' set the current resistance
        m_currentResistanceNo = resistanceNo
        
        ' get a single scan
        ScanReadResistance resistanceNo
    
    Next resistanceNo
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' enable all buttons
    EnableButtons
    
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' enable all buttons
    EnableButtons
    
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
    If Not K2700.Connected Then
        K2700.Initialize sheetDataAcq.host, sheetDataAcq.port, sheetDataAcq.ReceiveTimeout
    End If
    
    ' get the resistance number that corresponds to the active cell
    
    ' get a single scan
    ExecuteReadScan GetActiveCellResistanceNumber()
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' enable all buttons
    EnableButtons
    
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
    
    ' check if the manual single mode is toggled on or off:
    If toggleManualSingle.value Or toggleManualScan.value Then
  
        ' clear the triggered flag
        m_extTrigInitiated = False
        
        ' reinitialize if device not defined
        If Not K2700.Connected Then
            K2700.Initialize sheetDataAcq.host, sheetDataAcq.port, sheetDataAcq.ReceiveTimeout
        End If
        
        ' make sure that the instrument is ready to accept
        ' input from the front panel
     
        ' check if front or read
        If Not K2700.IsFrontInputs Then
        
            ' if rear, toggle back the button
            toggleManualSingle.value = False
            
            ' display the proper message
            MsgBox "Please release the front panel INPUTS button to its FRONT (F) position"
              
            ' and get out
            GoTo exit_Handler
      
        End If
  
        ' set the instrument for externally triggered measurements
        K2700.SetExternalTriggerResistanceMode
    
        ' set the first resistance for manual scan
        m_currentResistanceNo = 1
       
        ' start the timer
        StartTimer timerInterval
        
        ' now set the timer in motion
        ' motmrAction.Interval = 100
        ' motmrAction.Enabled = True
  
    ElseIf Not toggleManualSingle.value Then
  
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
''' Turn on or off manual triggering mode of the instrument.
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
    
    If toggleManualScan.value Then
    
        ' Clear all readings
        sheetDataAcq.ClearReadings Range(slopeCell).value * Range(offsetCell).value
        
        ' Set the data and time
        Range(dateCell).value = Now
        Range(timeCell).value = Format$(Now, hourMinuteTimeFormat)
        
        ' if we start manual mode, disable the rest of the buttons
        Me.cmdAutoSingle.Enabled = False
        Me.cmdAutoScan.Enabled = False
        Me.toggleManualSingle.Enabled = False
        
        ' set triggered mode
        SetExternalTriggerMode
    
    ElseIf Not toggleManualScan.value Then
    
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
    EnableButtons
    
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
    
    If toggleManualSingle.value Then
  
        ' if we start manual mode, disable the rest of the buttons
        cmdAutoSingle.Enabled = False
        cmdAutoScan.Enabled = False
        Me.toggleManualScan.Enabled = False
        
        ' set triggered mode
        SetExternalTriggerMode
  
    ElseIf Not toggleManualSingle.value Then
  
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
    EnableButtons
    
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
    
    If pauseRequested Then Exit Sub
   
    Dim sData As String
    Dim resistanceNo As Integer
      
    ' check if scan or single
    If toggleManualSingle.value Then
      
        ' get the cell number
        resistanceNo = Me.GetActiveCellResistanceNo()
    
    ElseIf toggleManualScan.value Then
    
        ' select the current resistance
        resistanceNo = m_currentResistanceNo
          
    End If
      
    ' allow other events to occur
    DoEvents
    
    Private isDone As Boolean
    isDone = Not (toggleManualSingle.value Or toggleManualScan.value)
    
    ' stop the timer.
    If isDone Then PauseTimer
    
    K2700.MonitorExternalTrigger resistanceNo, isDone
    
    ' if we ended manual mode, enable the rest of the buttons
    If isDone Then EnableButtons
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' stop the timer.
    PauseTimer
    
    ' enable all buttons
    EnableButtons
    
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, Me.name
    
    ' display the error message
    MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary> </summary>
Friend Property Get host() As String
    host = Range("B2").value
End Property

''' <summary> </summary>
Friend Property Get port() As Integer
    port = CInt(Range("B3").value)
End Property

''' <summary> </summary>
Friend Property Get ReceivedMessage() As String
    ReceivedMessage = Range("I2").value
End Property

''' <summary> </summary>
Friend Property Let ReceivedMessage(ByVal value As String)
    Range("I2").value = value
    Range("H2").value = Len(value)
End Property

''' <summary> </summary>
Friend Property Get SentMessage() As String
    SentMessage = Range("G2").value
End Property

''' <summary> </summary>
Friend Property Let SentMessage(ByVal value As String)
    Range("G2").value = value
End Property

''' <summary>   Starts the event timer. </summary>
''' <param name="timerInterval">   Specifies the timer interval in milliseconds. </param>
Private Sub StartTimer(Optional ByVal timerInterval As Long = 100)
    pauseRequested = False
    If m_eventTimer Is Nothing Then
        Set m_eventTimer = cc_isr_Core.Constructor.CreateStdTimer(timerInterval)
        m_eventTimer.StartTimer
    Else
        m_eventTimer.ResumeTimer
    End If
End Sub

''' <summary>   Requests pausing the timer. </summary>
Private Sub PauseTimer()
    pauseRequested = True
End Sub

''' <summary>   Aborts the timer. </summary>
Private Sub AbortTimer()
  ' we need to figure out how to recreate the timer
  m_eventTimer.StopTimer
End Sub

''' <summary>   Handles a timer tick event. </summary>
Private Sub m_eventTimer_Tick()
    On Error Resume Next
    If pauseRequested Then
        m_eventTimer.PauseTimer
    Else
        HandleTimerEvent
    End If
End Sub

''' <summary>   Handles the K2700 <see cref="cc_isr_VI.ResistanceNumberMeasured"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_VI.ResistanceNumberEventArgs"/> event arguments. </param>
Private Sub m_k2700_ResistanceNumberMeasured(ByRef eventArgs As cc_isr_VI.ResistanceNumberEventArgs)

    sheetDataAcq.SetResistance eventArgs.ResistanceNumber, eventArgs.resistance
    
    ' increment the resistnace for the next reading
    m_currentResistanceNo = m_currentResistanceNo + 1
    If m_currentResistanceNo > sheetDataAcq.maxResistanceCount Then
        m_currentResistanceNo = 1
    End If

End Sub

Private Sub m_device_ConnectionChanged(ByRef eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

End Sub



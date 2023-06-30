VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GpibLanController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE_NAME As String = "GpibLanController"
Private Const GPIB_LAN_PORT_NO As Integer = 1234

Private m_controllerMode As Boolean

Private m_disableReadAfterWriteOnWrite As Boolean

Private m_enabled As Boolean

Private m_readAfterWriteEnabled As Boolean

Private m_readTermination As String

Private m_readTimeout As Integer

Private m_readAfterWriteDelay As Integer

Private m_statusByte As Integer

Private m_stopper As StopWatch

Private WithEvents m_tcpClient As TcpClient

Private m_writeTermination As String

''' <summary>   Handles the Class Initialize event; constructs the <see cref="TcpClient"/>. </summary>
''' <remarks>   This GPIB Lan class uses the Prologix GPIB-Ethernet Controller. </remarks>
Private Sub Class_Initialize()
    Set m_stopper = cc_isr_Core.Constructor.CreateStopWatch
    m_readTermination = vbLf
    m_writeTermination = vbLf
    m_controllerMode = True
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="TcpClient"/>. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Initializes the <see cref="TcpClient"/> for the GpibLan I/O. </summary>
''' <param name="client">                  Reference to the <see cref="TcpClient"/>. </param>
''' <param name="readTerm">                [Optiona, String, VbLf] the read termination. </param>
''' <param name="writeTerm">               [Optiona, String, VbLf] the write termination. </param>
''' <param name="readAfterWriteDelayMs">   [Optional, Integer, 5] The read after write delay. </param>
Public Sub Initialize(ByVal client As TcpClient, _
                      Optional ByVal readTerm As String = vbLf, _
                      Optional ByVal writeTerm As String = vbLf, _
                      Optional ByVal readAfterWriteDelayMs As Integer = 5)
    
    Set m_tcpClient = client
    m_readTermination = readTerm
    m_writeTermination = writeTerm
    Me.ReadAfterWriteDelay = readAfterWriteDelayMs
    
    ' this needs to be turned on if using the Keithley 2700 with
    ' enabled read-after-write
    Me.DisableReadAfterWriteOnWrite = False
   
End Sub

''' <summary>   Closes and releases the TCP Client. </summary>
Public Sub Dispose()
    
    ' because the tcp client is passed by value we can null this
    m_tcpClient = Nothing

    Set m_stopper = Nothing

End Sub

''' <summary>   Gets the enabled status. </summary>
''' <remarks>
''' The GPIB-Lan controller is enabled only if the Tcp client is
''' connected to a GPIB-Lan controller such as the Prologix GPIB-Lan controller.
''' </remarks>
''' <value>   [Boolean] True if enabled.  </value>
Public Property Get Enabled() As Boolean
    Enabled = m_enabled
End Property

''' <summary>   Sets the enabled status. </summary>
''' <remarks>
''' The GPIB-Lan controller is enabled only if the Tcp client is
''' connected to a GPIB-Lan controller such as the Prologix GPIB-Lan controller.
''' </remarks>
''' <param name=""value>   [Boolean] True if enabled.  </value>
Public Property Let Enabled(ByVal value As Boolean)
    m_enabled = value
End Property

''' <summary>   Gets the Read After Write Delay. </summary>
''' <value>   An Integer. </value>
Public Property Get ReadAfterWriteDelay() As Integer
    ReadAfterWriteDelay = m_readAfterWriteDelay
End Property

''' <summary>   Sets the Read After Write Delay. </summary>
Public Property Let ReadAfterWriteDelay(ByVal value As Integer)
    m_readAfterWriteDelay = value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Controller mode Device IO
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Disable Read After Write On Write. </summary>
''' <remarks>   Read-After-Write should be disabled on non-query commands
''' with instruments that throw Query Unterminated errors on such commands if
''' Read-After-Write is on. </remarks>
''' <value>   A Boolean. </value>
Public Property Get DisableReadAfterWriteOnWrite() As Boolean
    DisableReadAfterWriteOnWrite = m_disableReadAfterWriteOnWrite
End Property

''' <summary>   Sets the read termination. </summary>
''' <param name="value">   True to Disable Read After Write On Write. </param>
Public Property Let DisableReadAfterWriteOnWrite(value As Boolean)
    m_disableReadAfterWriteOnWrite = value
End Property

''' <summary>   Configure the Read-After-Write mode of the controller. </summary>
''' <remarks>
''' The controller can be configured to automatically address
''' instruments to talk after sending them a command in order to read their response. The
''' feature called, Read-After-Write, saves the user from having to issue read commands
''' repeatedly. This command enabled or disabled the Read-After-Write feature.
'''
''' In addition, auto command also addresses the instrument at the currently specified
''' address to TALK or LISTEN. ++auto 0 addresses the instrument to LISTEN and
''' ++auto 1 addresses the instrument to TALK.
''' If the command is issued without any arguments it returns the current state of the
''' readafter-write feature.
'''
''' NOTE:
''' Some instruments generate “Query Unterminated” or “-420” error if they are addressed
''' to talk after sending a command that does not generate a response (often called non
''' query commands). In effect the instrument is saying, I have been asked to talk but I have
''' nothing to say. The error is often benign and may be ignored. Otherwise, use the
''' ++read command to read the instrument response. For example:
''' <code>
''' ++auto 0 — Turn off read-after-write and address instrument to listen
''' SET VOLT 1.0 — Non-query command
''' *idn? — Query command
''' ++read eoi — Read until EOI asserted by instrument
''' "HP54201A" — Response from instrument
''' </code>
''' <param name="enable">   True to enable read-after-write. </param>
''' </remarks>
Public Sub ReadAfterWriteEnabledSetter(ByVal enable As Boolean)
    
    If enable <> Me.ReadAfterWriteEnabled Then
        Dim output As String: output = IIf(enable, "1", "0")
        Me.SendToController "++auto " & output
        m_readAfterWriteEnabled = enable
    End If

End Sub

''' <summary>   Query the Read-After-Write enabled mode of the controller. </summary>
Public Function ReadAfterWriteEnabledGetter() As Boolean
    m_readAfterWriteEnabled = "1" = Me.QueryController("++auto")
    ReadAfterWriteEnabledGetter = m_readAfterWriteEnabled
End Function

''' <summary>   Sends a message to the device. </summary>
''' <remarks>   If using the controller at port 1234 and auto read-after-write is on,
''' this method first sets the controller to auto off (++auto 0) to prevent it from
''' setting the device to talk prematurely which might cause the device
''' (e.g., the Keithley 2700 scanning multimeter) to issue error -420 Query Unterminated. </remarks>
''' <param name="message">   The message. </param>
''' <returns>   [Long] The number of sent characters. </returns>
Public Function SendToDevice(ByVal message As String, _
        Optional ByVal appendTermination As Boolean = True) As Long
    
    If appendTermination Then message = message & m_writeTermination
    
    ' if auto read after write and write control of read after write is enabled,
    ' turn on listen to prevent Query Unterminated error on the 2700.
    
    If Me.ReadAfterWriteEnabled And Me.DisableReadAfterWriteOnWrite Then Me.ReadAfterWriteEnabledSetter False
    
    ' send the message to the device.
    
    SendToDevice = m_tcpClient.SendMessage(message)
    
    ' a read after write delay is necessary for proper operations.
    
    m_stopper.Wait Me.ReadAfterWriteDelay

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <remarks>   If <see cref="ReadAfterWriteEnabled"/> is not enabled then
''' this method uses the <c>++read</c> command to first read the data
''' from the device to the controller. </remarks>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <returns>   [String] The received message. </returns>
Public Function ReceiveFromDevice(Optional ByVal maxLength As Long = &H7FFF, _
                                  Optional ByVal trimEnd As Boolean = True) As String
    
    If Not Me.ReadAfterWriteEnabled Then
        ' if auto read after write is disabled, we can either turn on ' read after write,
        ' but it seems that getting the controller to read from the instrumet is much faster.
        
        Me.ReadFromDeviceToController
    
    End If
    
    ReceiveFromDevice = m_tcpClient.ReceiveRaw(maxLength, trimEnd, m_readTermination)

End Function
    
''' <summary>   Sends a message to the device and receives a reply. </summary>
''' <param name="message">   The message. </param>
''' <returns>   [String] The received message. </returns>
Public Function QueryDevice(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    If Me.SendToDevice(message, appendTermination) > 0 Then
        DoEvents
        QueryDevice = Me.ReceiveFromDevice()
    Else
        QueryDevice = ""
    End If
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' CONTROLLER I/O
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sends a message to the controller. </summary>
''' <remarks>   This method does not alter the auto Read-After-Write condition. </remarks>
''' <param name="message">             [String] The message. </param>
''' <param name="appendTermination">   [Optional, Boolean, True] True to append the
'''                                    <see cref="WriteTermination"/> to the message. </param>
''' <returns>   [Long] The number of sent characters. </returns>
Public Function SendToController(ByVal message As String, _
        Optional ByVal appendTermination As Boolean = True) As Long
    
    If appendTermination Then message = message & m_writeTermination
    SendToController = m_tcpClient.SendMessage(message)
    m_stopper.Wait Me.ReadAfterWriteDelay

End Function

''' <summary>   Receives a message from the controller. </summary>
''' <param name="trimEnd">   [Optional, Boolean, True] True to trim the end termination. </param>
''' <returns>   [String] The received message. </returns>
Public Function ReceiveFromController(Optional trimEnd As Boolean = True) As String

    Dim receivedMessage As String: receivedMessage = m_tcpClient.ReceiveRaw(trimEnd:=trimEnd, _
                                                                            termination:=m_readTermination)
    
    ' messages that come from the controller have a CR + LF termination. The Receive Raw method
    ' may stop at the LF termination, in which case the CR termination needs to be trimmed.
    If trimEnd Then
        ReceiveFromController = cc_isr_Core.StringExtensions.TrimRight(receivedMessage, vbCr)
    End If
    
End Function

''' <summary>   Sends a message to the controller and read back the reply. </summary>
''' <param name="message">   The message. </param>
''' <returns>   The received message. </returns>
Public Function QueryController(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    If Me.SendToController(message, appendTermination) > 0 Then
        DoEvents
        QueryController = Me.ReceiveFromController()
    Else
        QueryController = ""
    End If
End Function

''' <summary>   Sends the '++read eoi' copmmand to have the controller set the instrument to talk
''' and get the data from the instrument into the controller until EOI character is recevied.
''' </summary>
''' <remarks>
''' This command can be used to read data from an instrument until:
''' <code>
''' EOI is detected or timeout expires, or
''' A specified character is read, or
''' timeout expires,
''' </code>
''' Timeout is set using the read_tmo_ms command and applies to inter-character delay, i.e.,
''' the delay since the last character was read. Timeout is not be confused with the total
''' time for which data is read.
''' SYNTAX: ++read [eoi|<char>] where <char> is a decimal value less than 256
''' Examples:
''' <code>
'''  ++read     - Read until timeout
'''  ++read eoi - Read until EOI detected or timeout
'''  ++read 10  - Read until LF (ASCII 10) is received or timeout
''' </code>
''' </remarks>
''' <param name="message">   The message. </param>
''' <returns>   [Long] The number of characters that were sent. </returns>
Public Function ReadFromDeviceToController() As Integer
    ReadFromDeviceToController = Me.SendToController("++read eoi")
End Function

''' <summary>   Gets the read-after-write enabled status. </summary>
''' <value>   [Boolean] True if read-after-write is enabled.  </value>
Public Property Get ReadAfterWriteEnabled() As Boolean
    ReadAfterWriteEnabled = m_readAfterWriteEnabled
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB METHODS
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Issues a Go To Local (GTL) (++loc). </summary>
''' <remarks>   Valid if <see cref="UsingGpibLan"/>. </remarks>
Public Sub GoToLocal()
    
    Me.SendToController "++loc"
    
End Sub

''' <summary>   Query the GPIB address. </summary>
''' <value>    An string including the primary and secondary address separated by a space. </value>
Public Function GpibAddressGetter() As String
    
    GpibAddressGetter = "-1"
    Dim receivedMessage As String: receivedMessage = Me.QueryController("++addr")
    Dim addresses() As String: addresses = Split(receivedMessage, " ")
    If 0 = UBound(addresses) Then
        GpibAddressGetter = receivedMessage
    ElseIf 1 = UBound(addresses) - LBound(addresses) + 1 Then
        GpibAddressGetter = addresses(1)
        If 1 < UBound(addresses) - LBound(addresses) + 1 Then
            GpibAddressGetter = GpibAddressGetter & " " & addresses(2)
        End If
    End If
    
End Function

''' <summary>   Configure the GPIB address of the GPIB Lan Controller.
''' The meaning of the GPIB address depends on the operating mode of the controller.
'''
''' In Controller mode, it refers to the GPIB address of the instrument being controlled.
''' In DEVICE mode, it is the address of the GPIB peripheral that the controller is emulating. </summary>
''' <remarks>
''' An optional secondary address may also be specified.
'''
''' Internally, the secondary address, which is offset by 96, must be separated from the primary
''' address by a space character. Specifying secondary address has no effect in DEVICE mode.
''' </remarks>
''' <para name="primaryAddress">     [Integer] Specifies the primary GPIB address
'''                                   between 0 and 30. </param>
''' <para name="secondaryAddress)>   [Optional, Integer] Specifies the secondard GPIB Address
'''                                  between 0 and 30. </param>
''' <returns>   The status byte. </returns>
Public Sub GpibAddressSetter(ByVal primaryAddress As Integer, _
                             Optional ByVal secondaryAddress As Integer = -1)
    
    If primaryAddress >= 0 And secondaryAddress < 0 Then
    
        primaryAddress = Me.DelimitInteger(primaryAddress, 0, 30)
        Me.SendToController "++addr " & CStr(primaryAddress)
        
    ElseIf primaryAddress >= 0 And secondaryAddress >= 0 Then
        
        primaryAddress = Me.DelimitInteger(primaryAddress, 0, 30)
        secondaryAddress = Me.DelimitInteger(secondaryAddress, 0, 30)
        Me.SendToController "++addr " & CStr(primaryAddress) & _
            " " & CStr(secondaryAddress + 96)
    End If
    
End Sub

''' <summary>   Issues a local lockout (++llo). </summary>
''' <remarks>   Valid if <see cref="UsingGpibLan"/>. </remarks>
Public Sub LocalLockout()
    
    Me.SendToController "++llo"
    
End Sub

''' <summary>   Gets the device read timeout in Milliseconds. </summary>
''' <value>   An Integer. </value>
Public Property Get ReadTimeout() As Integer
    ReadTimeout = m_readTimeout
End Property

''' <summary>   Sets the device read timeout, in milliseconds, to be used in the
'''  controller read and spoll commands. Timeout may be set to any value between 1 and 3000
''' milliseconds.
''' </summary>
''' <param name="timeoutMs">   The timeout interval in milliseconds. </param>
Public Sub ReadTimeoutSetter(ByVal timeoutMs As Integer)
    
    timeoutMs = Me.DelimitInteger(timeoutMs, 1, 3000)
    
    If timeoutMs <> m_readTimeout Then
        Me.SendToController "++read_tmo_ms " & CStr(timeoutMs)
    End If
    m_readTimeout = timeoutMs
    
End Sub

''' <summary>   Gets the device read timeout, in milliseconds, that is used in the
'''  controller read and spoll commands.
''' </summary>
Public Function ReadTimeoutGetter() As Integer
    
    m_readTimeout = CInt(Me.QueryController("++read_tmo_ms"))
    ReadTimeoutGetter = m_readTimeout
    
End Function

''' <summary>   Issues an SDC. </summary>
''' <remarks>   Valid if <see cref="UsingGpibLan"/>. </remarks>
Public Sub SelectiveDeviceClear()
    
    Me.SendToController "++clr"
    
End Sub

''' <summary>   Performs a serial poll of the instrument at the current or the specified address.
''' If no address is specified then this command serial polls the currently addressed instrument (as
''' set by a previous ++addr command). This command uses the time-out value specified
''' by the read_tmo_ms command.
''' </summary>
''' <remarks>
''' Serial polling obtains specific information from a device. When you serial poll, the Controller
''' sends a special command message—Serial Poll Enable (SPE)—to the device, directing it to return
''' its serial poll status byte. The SPE message sets the IEEE 488.1 serial poll mode in the device,
''' so when the device is addressed to talk, it returns a single 8-bit status byte. This serial poll
''' status byte is different for each type of instrument; except for one bit, you must refer to the
''' instrument user manual for information on the other bits. Bit 6 (hex 40) of any serial poll status
''' byte indicates whether a device requested service by asserting the SRQ line. The device uses the
''' other seven bits of the status byte to specify why it needs attention.
'''
''' After the Controller reads the status byte, it sends another command message, Serial Poll Disable (SPD),
''' to the device. The SPD message terminates the serial poll mode, thus returning the device to its normal
''' Talker/Listener state. Once a device requesting service is serial polled, it usually unasserts the SRQ
''' line.
''' </remarks>
''' <para name="primaryAddress">     [Optional, Integer] Specifies the primary GPIB address
'''                                   between 0 and 30. </param>
''' <para name="secondaryAddress)>   [Optional, Integer] Specifies the secondard GPIB Address
'''                                  between 0 and 30. </param>
''' <returns>   The status byte. </returns>
Public Function SerialPoll(Optional ByVal primaryAddress As Integer = -1, _
                           Optional ByVal secondaryAddress As Integer = -1) As Integer
    
    If primaryAddress >= 0 And secondaryAddress < 0 Then
        SerialPoll = CInt(Me.QueryController("++spoll " & CStr(primaryAddress)))
    ElseIf primaryAddress >= 0 And secondaryAddress >= 0 Then
        SerialPoll = CInt(Me.QueryController("++spoll " & CStr(primaryAddress) & _
                     " " & CStr(secondaryAddress + 96)))
    Else
        SerialPoll = CInt(Me.QueryController("++spoll"))
    End If
    
End Function

''' <summary>   Queries the GPIB SRQ signal status. </summary>
''' <remarks>   This command returns the current state of the GPIB SRQ signal. The controler returns
''' <c>1</c> if the SRQ signal is asserted (low) and <c>0</c> if the signal is not asserted (high). </remarks>
''' <returns>   True if Server requested. </returns>
Public Function ServiceRequested() As Boolean
    
    ServiceRequested = 1 = CInt(Me.QueryController("++srq"))
    
End Function


' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Device Mode Functions
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the controller mode status. </summary>
''' <value>   [Boolean] True if the GPIB-Lan is set as a controller. </value>
Public Property Get ControllerMode() As Boolean
    ControllerMode = m_controllerMode
End Property

''' <summary>   Gets the device Status Byte. </summary>
''' <value>   An Integer. </value>
Public Property Get StatusByte() As Integer
    StatusByte = m_statusByte
End Property

''' <summary>   Sets the device Status Byte to be returned on a serial poll. </summary>
''' <remarks>
''' This command works only if the controller is set as a device using the ++mode command.
''' The status command is used to specify the device status byte to be returned when serial
''' polled by a GPIB controller. If the RQS bit (bit #6) of this status byte is set then the
''' SRQ signal is asserted (low). After a serial poll, SRQ line is de-asserted and status byte is
''' set to 0. Status byte is initialized to 0 on power up.
''' SRQ is also de-asserted and status byte is cleared if DEVICE CLEAR (DCL) message,
''' or SELECTED DEVICE CLEAR (SDC) message, is received from the GPIB controller.
''' If the command is issued without any arguments it returns the currently specified status
''' byte.
''' SYNTAX: ++status [0-255]
''' Example:
''' ++status 72 Specify serial poll status byte as 72. Since bit #6 is set, this
'''             command will assert SRQ.
''' ++status Query current serial poll status byte.
''' </remarks>
''' <param name="value">   The status byte mask. </param>
Public Sub StatusByteSetter(ByVal value As Integer)
    
    If Not Me.ControllerMode Then
        value = Me.DelimitInteger(value, 0, 255)
        
        If value <> m_statusByte Then
            Me.SendToController "++status " & CStr(value)
        End If
    End If
    m_statusByte = value
    
End Sub

''' <summary>   Gets the device Status Byte. </summary>
Public Function StatusByteGetter() As Integer
    If Me.ControllerMode Then
        m_statusByte = 0
    Else
        m_statusByte = CInt(Me.QueryController("++status"))
    End If
    StatusByteGetter = m_statusByte
End Function

''' <summary>   Returns the delimited integer value. </summary>
''' <param name="value">   [Integer] the value to delimit.</param>
''' <param name="minValue">   [Integer] the minimum. </param>
''' <param name="maxValue">   [Integer] the maximum. </param>
Public Function DelimitInteger(ByVal value As Integer, ByVal minValue As Integer, ByVal maxValue As Integer) As Integer
    If value < minValue Then value = minValue
    If value > maxValue Then value = maxValue
    DelimitInteger = value
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  TCP Client Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Winsock.TcpClient.ConnectionChanged"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_tcpClient_ConnectionChanged(ByVal eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Const thisProcedureName As String = "m_tcpClient_ConnectionChanging"

    On Error Resume Next
    
    ' the GPIB-Lan controller is enabled only if the Tcp client is
    ' connected to a GPIB-Lan controller such as the Prologix GPIB-Lan controller.
    If Not Me.Enabled Then Return

    ' when connected, the read after write is turned off by default.
    ' this is done to prevent Query Unterminated errors on instruments such as the Keithley 2700.
    ' Where Read-After-Write is disabled, upon read, the controller is commanded to read from
    ' the instrument using the <c>++read</c> command.
    
    If eventArgs.Connected Then
    
        ' from experiments it seems this needs to be set first.
        Dim isEnabled As Boolean: isEnabled = Me.ReadAfterWriteEnabledGetter
    
        Me.ReadAfterWriteEnabledSetter True
    
        ' defaults to turning off auto read-after-write
        
        Me.ReadAfterWriteEnabledSetter False
        
    End If
    
    If Err.Number <> 0 Then
        ' report any error when debugging this code.
        UserDefinedErrors.SetErrSource thisProcedureName, MODULE_NAME
        Debug.Print UserDefinedErrors.BuildStandardErrorMessage()
    End If
    
End Sub

''' <summary>   Handles the <see cref="cc_isr_Winsock.TcpClient.ConnectionChanging"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_tcpClient_ConnectionChanging(ByVal eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

    Const thisProcedureName As String = "m_tcpClient_ConnectionChanging"
    
    On Error Resume Next
    
    ' the GPIB-Lan controller is enabled only if the Tcp client is
    ' connected to a GPIB-Lan controller such as the Prologix GPIB-Lan controller.
    If Not Me.Enabled Then Return
    
    If eventArgs.Connected Then
    
        ' leave the instrument with auto read-after-write off to prevent
        ' query unterminated errors.
        
        Dim isEnabled As Boolean: isEnabled = Me.ReadAfterWriteEnabledGetter
        Me.ReadAfterWriteEnabledSetter False
        
        ' send the instrument back to local.
        
        Me.GoToLocal
        
    End If

    If Err.Number <> 0 Then
        ' report any error when debugging this code.
        UserDefinedErrors.SetErrSource thisProcedureName, MODULE_NAME
        Debug.Print UserDefinedErrors.BuildStandardErrorMessage()
    End If
    
End Sub

''' <summary>   Experiments with options for setting read-after-write. </summary>
''' <remarks>   good options: 1, 2 ?, 5; bad options: 4, 3. </remarks>
''' <para name="opt">   [Integer] selects an configuration option. </param>
Private Sub InitReadAfterWrite(ByVal opt As Integer)

    If opt = 1 Then
    
        ' Identity query time:         40-50 ms
        ' Query Unterminated on write: YES
        ' Crash on power toggle/Local: no
    
        Me.ReadAfterWriteEnabledGetter
        
        Me.ReadAfterWriteEnabledSetter True
    
    ElseIf opt = 2 Then
    
        ' Identity query time:         40-50 ms
        ' Query Unterminated on write: YES
        ' Crash on power toggle/Local: only if toggling when program already started.
        m_readAfterWriteEnabled = False
        Me.ReadAfterWriteEnabledSetter True
    
    ElseIf opt = 3 Then
        
        ' Identity query time:         74 ms
        ' Query Unterminated on write: no
        ' Crash on power toggle/Local: no
        Me.ReadAfterWriteEnabledGetter
        Me.ReadAfterWriteEnabledSetter False
        
        ' this takes it from 380 to 74! go figure!!!
        ' Crash on power toggle/Local: no
        DoEvents
        
        ' this takes it to 24 ms! go figure!!!
        ' Crash on power toggle/Local: YES
    
    ElseIf opt = 4 Then
    
        ' Identity query time:         22 ms
        ' Query Unterminated on write: no
        ' Crash on power toggle:       YES
        ' reading fails after the instrument power is toggled.
        ' apparently, ++read works only if auto read is set first.
        
        m_readAfterWriteEnabled = True
        Me.ReadAfterWriteEnabledSetter False
    
    ElseIf opt = 5 Then
    
        ' Identity query time:         48 ms
        ' Query Unterminated on write: no
        ' Crash on power toggle/local: no
        
        Me.ReadAfterWriteEnabledSetter True
        
        Me.ReadAfterWriteEnabledSetter False
        
    End If

End Sub



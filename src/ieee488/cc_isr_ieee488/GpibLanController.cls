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
Private Const GPIB_LAN_PORT_NO As Integer = 1234
Private m_readTermination As String
Private m_writeTermination As String
Private m_tcpClient As TcpClient
Private m_readAfterWriteEnabled As Boolean
Private m_disableReadAfterWriteOnWrite As Boolean

''' <summary>   Handles the Class Initialize event; constructs the <see cref="TcpClient"/>. </summary>
''' <remarks>   This GPIB Lan class uses the Prologix GPIB-Ethernet Controller. </remarks>
Private Sub Class_Initialize()
    Set m_tcpClient = cc_isr_Winsock.Constructor.CreateTcpClient
    m_readTermination = vbLf
    m_writeTermination = vbLf
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="TcpClient"/>. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Initializes the <see cref="TcpClient"/> for the GpibLan I/O. </summary>
''' <param name="client">                     Reference to the <see cref="TcpClient"/>. </param>
''' <param name="readTerm">                   [Optiona, String, VbLf] the read termination. </param>
''' <param name="writeTerm">                  [Optiona, String, VbLf] the write termination. </param>
''' <param name="preventQueryUnterminated">   [Optional, Boolea, True] prevent Query Unterminated
'''                                           error by disabling read-after-write before write. </param>
Public Sub Initialize(ByRef client As TcpClient, Optional ByVal readTerm As String = vbLf, _
                      Optional ByVal writeTerm As String = vbLf, _
                      Optional ByVal preventQueryUnterminated As Boolean = True)
    Set m_tcpClient = client
    m_readTermination = readTerm
    m_writeTermination = writeTerm
    Me.DisableReadAfterWriteOnWrite = preventQueryUnterminated
    
    Dim opt As Integer: opt = 5
    
    If opt = 1 Then
    
        ' Identity query time:         120 - 160 ms
        ' Query Unterminated on write: YES
        ' Crash on power toggle/Local: no
    
        Me.ReadAfterWriteGetter
        DoEvents
        Me.ReadAfterWriteEnabledSetter True
        DoEvents
    
    ElseIf opt = 2 Then
    
        ' Identity query time:         70 ms
        ' Query Unterminated on write: YES
        ' Crash on power toggle/Local: no
        m_readAfterWriteEnabled = False
        Me.ReadAfterWriteEnabledSetter True
        DoEvents
    
    ElseIf opt = 3 Then
        
        ' Identity query time:         74 ms
        ' Query Unterminated on write: no
        ' Crash on power toggle/Local: no
        Me.ReadAfterWriteGetter
        DoEvents
        Me.ReadAfterWriteEnabledSetter False
        
        ' this takes it from 380 to 74! go figure!!!
        DoEvents
    
    ElseIf opt = 4 Then
    
        ' Identity query time:         22 ms
        ' Query Unterminated on write: no
        ' Crash on power toggle:       YES
        ' but reading fails on the second read after the
        ' instrument power is toggled.
        m_readAfterWriteEnabled = True
        Me.ReadAfterWriteEnabledSetter False
        DoEvents
    
    ElseIf opt = 5 Then
    
        ' Identity query time:         48 ms
        ' Query Unterminated on write: no
        ' Crash on power toggle/local: no
        Me.ReadAfterWriteEnabledSetter True
        DoEvents
        Me.ReadAfterWriteEnabledSetter False
        DoEvents
        
        
    End If
    
End Sub

''' <summary>   Closes and releases the TCP Client. </summary>
Public Sub Dispose()
    If m_tcpClient.Connected Then
        ' leave the instrument in it's default condition to prevent
        ' timeouts on the next read.
        ' Me.ReadAfterWriteEnabledSetter True
    End If
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' DEVICE IO

''' <summary>   Sends a message to the device. </summary>
''' <remarks>   If using the controller at port 1234, this
''' method first sets the controller to auto off (++auto 0) to prevent it from
''' setting the device to talk prematurely which might cause the device
''' (e.g., the Keithley 2700 scanning multimeter) to issue error -420 Query Unterminated. </remarks>
''' <param name="message">   The message. </param>
''' <returns>   The number of sent characters. </returns>
Public Function SendToDevice(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    
    If appendTermination Then message = message & Me.WriteTermination
    
    ' if auto read after write, turn on listen to prevent Query Unterminated error on the 2700.
    If Me.ReadAfterWriteEnabled And Me.DisableReadAfterWriteOnWrite Then Me.ReadAfterWriteEnabledSetter False
    
    ' send the message to the device.
    SendToDevice = m_tcpClient.SendMessage(message)

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <remarks>   If using the controller at port 1234, this
''' method first sets the controller to auto on (++auto 1) assuming auto was turned off on
''' the previous read on <see cref="SendToDevice"/>. </remarks>
''' <param name="maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <returns>   The received string. </returns>
Public Function ReceiveFromDevice(Optional ByVal maxLength As Long = &H7FFF, _
                                  Optional ByVal trimEnd As Boolean = True) As String
    If Not Me.ReadAfterWriteEnabled Then
        ' if auto read after write is disabled, we can either turn on ' read after write,
        ' but it seems that getting the controller to read from the instrumet is much faster.
        Me.ReadFromDeviceToController
    End If
    ReceiveFromDevice = m_tcpClient.ReceiveRaw(maxLength, trimEnd, Me.ReadTermination)

End Function
    
''' <summary>   Sends a message to the device and receives a reply. </summary>
''' <param name="message">   The message. </param>
''' <returns>   The received string. </returns>
Public Function QueryDevice(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    If Me.SendToDevice(message, appendTermination) > 0 Then
        DoEvents
        QueryDevice = Me.ReceiveFromDevice()
    Else
        QueryDevice = ""
    End If
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' CONTROLLER I/O

''' <summary>   Gets the Disable Read After Write On Write. </summary>
''' <value>   A Boolean. </value>
Public Property Get DisableReadAfterWriteOnWrite() As Boolean
    DisableReadAfterWriteOnWrite = m_disableReadAfterWriteOnWrite
End Property

''' <summary>   Sets the read termination. </summary>
''' <param name="value">   True to Disable Read After Write On Write. </param>
Public Property Let DisableReadAfterWriteOnWrite(value As Boolean)
    m_disableReadAfterWriteOnWrite = value
End Property

''' <summary>   Sends a message to the controller. </summary>
''' <param name="message">   The message. </param>
''' <returns>   The number of sent characters. </returns>
Public Function SendToController(ByVal message As String, Optional ByVal appendTermination As Boolean = True)
    
    If appendTermination Then message = message & Me.WriteTermination
    SendToController = m_tcpClient.SendMessage(message)

End Function

''' <summary>   Receives a message from the controller. </summary>
Public Function ReceiveFromController() As String
    ReceiveFromController = m_tcpClient.ReceiveRaw(termination:=Me.ReadTermination)
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
''' <returns>   The number of characters that were sent. </returns>
Public Function ReadFromDeviceToController() As Integer
    ReadFromDeviceToController = Me.SendToController("++read eoi")
End Function

''' <summary>   Gets the read-after-write enabled status. </summary>
''' <value>   True if wrte-after-write is enabled.  </value>
Public Property Get ReadAfterWriteEnabled() As Boolean
    ReadAfterWriteEnabled = m_readAfterWriteEnabled
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

''' <summary>   Query the Read-After-Write mode of the controller. </summary>
Public Function ReadAfterWriteGetter() As Boolean
    m_readAfterWriteEnabled = "1" = Me.QueryController("++auto")
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB METHODS

''' <summary> Issues a Go To Local (GTL) (++loc). </summary>
''' <remarks>   Valid if <see cref="UsingGpibLan"/>. </remarks>
Public Sub GoToLocal()
    
    Me.SendToController "++loc"
    
End Sub

''' <summary>   Query the GPIB address. </summary>
''' <value>    An string including the primary and secondary address separated by a space. </value>
Public Function GpibAddressGetter() As String
    
    GpibAddressGetter = "-1"
    Dim addresses() As String: addresses = Split(Me.QueryController("++addr"), " ")
    If 1 = UBound(addresses) - LBound(addresses) + 1 Then
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
        Me.SendToController "++addr " & CStr(primaryAddress)
    ElseIf primaryAddress >= 0 And secondaryAddress >= 0 Then
        Me.SendToController "++addr " & CStr(primaryAddress) & _
            " " & CStr(secondaryAddress + 96)
    End If
    
End Sub

''' <summary> Issues a local lockout (++llo). </summary>
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
    
    If timeoutMs < 1 Then timeoutMs = 1
    If timeoutMs > 3000 Then timeoutMs = 3000
    
    If timeoutMs <> m_readTimeout Then
        Me.SendToController "++read_tmo_ms " & CStr(timeoutMs)
    End If
    m_readTimeout = timeoutMs
    
End Sub

''' <summary>   Sets the device read timeout, in milliseconds, that is used in the
'''  controller read and spoll commands.
''' </summary>
Public Function ReadTimeoutGetter() As Integer
    
    m_readTimeout = CInt(Me.QueryController("++read_tmo_ms"))
    
End Function

''' <summary> Issues an SDC. </summary>
''' <remarks>   Valid if <see cref="UsingGpibLan"/>. </remarks>
Public Sub SelectiveDeviceClear()
    
    Me.SendToController "++clr"
    
End Sub

''' <summary> Performs a serial poll of the instrument at the current or the specified address.
''' If no address is specified then this command serial polls the currently addressed instrument (as
''' set by a previous ++addr command). This command uses the time-out value specified
''' by the read_tmo_ms command.
''' </summary>
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

''' <summary>   Query controller service request status. </summary>
''' <returns>   True if Server requested. </returns>
Friend Function IsSRQ() As Boolean
    
    IsSRQ = 1 = CInt(Me.QueryController("++srq"))
    
End Function




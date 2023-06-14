Attribute VB_Name = "Ieee488Syntax"
Option Explicit

''' <summary> Gets the Clear Status (CLS) command. </summary>
Public Const ClearExecutionStateCommand As String = "*CLS"

''' <summary> Gets the Identity query (*IDN?) command. </summary>
Public Const IdentityQueryCommand As String = "*IDN?"

''' <summary> Gets the operation complete (*OPC) command. </summary>
Public Const OperationCompleteCommand As String = "*OPC"

''' <summary> Gets the operation complete query (*OPC?) command. </summary>
Public Const OperationCompletedQueryCommand As String = "*OPC?"

''' <summary> Gets the options query (*OPT?) command. </summary>
Public Const OptionsQueryCommand As String = "*OPT?"

''' <summary> Gets the Wait (*WAI) command. </summary>
Public Const WaitCommand As String = "*WAI"

''' <summary> Gets the Standard Event Enable (*ESE {0}) command. </summary>
Public Const StandardEventEnableCommandFormat = "*ESE {0}"

''' <summary> Gets the Standard Event Enable query (*ESE?) command. </summary>
Public Const StandardEventEnableQueryCommand As String = "*ESE?"

''' <summary> Gets the Standard Event Status (*ESR?) command. </summary>
Public Const StandardEventStatusQueryCommand As String = "*ESR?"

''' <summary> Gets the Service Request Enable (*SRE) command. </summary>
Public Const ServiceRequestEnableCommandFormat = "*SRE {0}"

''' <summary> Gets the Standard Event and Service Request Enable '*CLS; *ESE {0}; *SRE {1}' command format. </summary>
Public Const StandardServiceEnableCommandFormat = "*CLS; *ESE {0}; *SRE {1}"

''' <summary> Gets the Standard Event and Service Request Enable '*CLS; *ESE {0}; *SRE {1}; *OPC' command format. </summary>
Public Const StandardServiceEnableCompleteCommandFormat = "*CLS; *ESE {0}; *SRE {1}; *OPC"

''' <summary> Gets the Operation Complete Enable '*CLS; *ESE {0}; *OPC' command format. </summary>
Public Const OperationCompleteEnableCommandFormat = "*CLS; *ESE {0}; *OPC"

''' <summary> Gets the Service Request Enable query (*SRE?) command. </summary>
Public Const ServiceRequestEnableQueryCommand As String = "*SRE?"

''' <summary> Gets the Service Request Status query (*STB?) command. </summary>
Public Const ServiceRequestQueryCommand As String = "*STB?"

''' <summary> Gets the reset to known state (*RST) command. </summary>
Public Const ResetKnownStateCommand As String = "*RST"

''' <summary> Enumerates the status byte flags of the standard event register. </summary>
''' <remarks>
''' Enumerates the Standard Event Status Register Bits. Read this information using ESR? or
''' status.standard.event. Use *ESE or status.standard.enable or event status enable to enable
''' this register.''' These values are used when reading or writing to the standard event
''' registers. Reading a status register returns a value. The binary equivalent of the returned
''' value indicates which register bits are set. The least significant bit of the binary number
''' is bit 0, and the most significant bit is bit 15. For example, assume value 9 is returned for
''' the enable register. The binary equivalent is
''' 0000000000001001. This value indicates that bit 0 (OPC) and bit 3 (DDE)
''' are set.
''' </remarks>
Public Enum standardEvents

    ''' <summary> The None option. </summary>
    None = 0

    ''' <summary>
    ''' Bit B0, Operation Complete (OPC). Set bit indicates that all
    ''' pending selected device operations are completed and the unit is ready to
    ''' accept new commands. The bit is set in response to an *OPC command.
    ''' The ICL function OPC() can be used in place of the *OPC command.
    ''' </summary>
    OperationComplete = 1

    ''' <summary>
    ''' Bit B1, Request Control (RQC). Set bit indicates that....
    ''' </summary>
    RequestControl = &H2

    ''' <summary>
    ''' Bit B2, Query Error (QYE). Set bit indicates that you attempted
    ''' to read data from an empty Output Queue.
    ''' </summary>
    QueryError = &H4

    ''' <summary>
    ''' Bit B3, Device-Dependent Error (DDE). Set bit indicates that a
    ''' device operation did not execute properly due to some internal
    ''' condition.
    ''' </summary>
    DeviceDependentError = &H8

    ''' <summary>
    ''' Bit B4 (16), Execution Error (EXE). Set bit indicates that the unit
    ''' detected an error while trying to execute a command.
    ''' This is used by QUATECH to report No Contact.
    ''' </summary>
    ExecutionError = &H10

    ''' <summary>
    ''' Bit B5 (32), Command Error (CME). Set bit indicates that a
    ''' command error has occurred. Command errors include:<p>
    ''' IEEE-488.2 syntax error — unit received a message that does not follow
    ''' the defined syntax of the IEEE-488.2 standard.  </p><p>
    ''' Semantic error — unit received a command that was misspelled or received
    ''' an optional IEEE-488.2 command that is not implemented.  </p><p>
    ''' The device received a Group Execute Trigger (GET) inside a program
    ''' message.  </p>
    ''' </summary>
    CommandError = &H20

    ''' <summary>
    ''' Bit B6 (64), User Request (URQ). Set bit indicates that the LOCAL
    ''' key on the SourceMeter front panel was pressed.
    ''' </summary>
    UserRequest = &H40

    ''' <summary>
    ''' Bit B7 (128), Power ON (PON). Set bit indicates that the device
    ''' has been turned off and turned back on since the last time this register
    ''' has been read.
    ''' </summary>
    PowerToggled = &H80

    ''' <summary>
    ''' Unknown value due to, for example, error trying to get value from the device.
    ''' </summary>
    Unknown = &H100

    ''' <summary>Includes all bits. </summary>
    All = &HFF ' 255

End Enum

''' <summary> Gets or sets the status byte bits of the service request register. </summary>
''' <remarks>
''' Enumerates the Status Byte Register Bits. Use *STB? or status.request_event to read this
''' register. Use *SRE or status.request_enable to enable these services. This attribute is used
''' to read the status byte, which is returned as a numeric value. The binary equivalent of the
''' returned value indicates which register bits are set. <para>
''' (c) 2005 Integrated Scientific Resources, Inc. All rights reserved. </para><para>
''' Licensed under The MIT License. </para>
''' </remarks>
Public Enum ServiceRequests

    ''' <summary> The None option. </summary>
    None = 0

    ''' <summary>
    ''' Bit B0, Measurement Summary Bit (MSB). Set summary bit indicates
    ''' that an enabled measurement event has occurred.
    ''' </summary>
    MeasurementEvent = &H1

    ''' <summary>
    ''' Bit B1, System Summary Bit (SSB). Set summary bit indicates
    ''' that an enabled system event has occurred.
    ''' </summary>
    SystemEvent = &H2

    ''' <summary>
    ''' Bit B2, Error Available (EAV). Set summary bit indicates that
    ''' an error or status message is present in the Error Queue.
    ''' </summary>
    ErrorAvailable = &H4

    ''' <summary>
    ''' Bit B3, Questionable Summary Bit (QSB). Set summary bit indicates
    ''' that an enabled questionable event has occurred.
    ''' </summary>
    QuestionableEvent = &H8

    ''' <summary>
    ''' Bit B4 (16), Message Available (MAV). Set summary bit indicates that
    ''' a response message is present in the Output Queue.
    ''' </summary>
    MessageAvailable = &H10

    ''' <summary>Bit B5, Event Summary Bit (ESB). Set summary bit indicates
    ''' that an enabled standard event has occurred.
    ''' </summary>
    standardEvent = &H20 ' (32) ESB

    ''' <summary>
    ''' Bit B6 (64), Request Service (RQS)/Master Summary Status (MSS).
    ''' Set bit indicates that an enabled summary bit of the Status Byte Register
    ''' is set. Depending on how it is used, Bit B6 of the Status Byte Register
    ''' is either the Request for Service (RQS) bit or the Master Summary Status
    ''' (MSS) bit: When using the GPIB serial poll sequence of the unit to obtain
    ''' the status byte (serial poll byte), B6 is the RQS bit. When using
    ''' status.condition or the *STB? common command to read the status byte,
    ''' B6 is the MSS bit.
    ''' </summary>
    RequestingService = &H40

    ''' <summary>
    ''' Bit B7 (128), Operation Summary (OSB). Set summary bit indicates that
    ''' an enabled operation event has occurred.
    ''' </summary>
    OperationEvent = &H80

    ''' <summary>
    ''' Includes all bits.
    ''' </summary>
    All = &HFF ' 255

    ''' <summary>
    ''' Unknown value due to, for example, error trying to get value from the device.
    ''' </summary>
    Unknown = &H100

End Enum


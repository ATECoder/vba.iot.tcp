# Testing the [cc.isr.Ieee488] Workbook

[cc.isr.Ieee488] is an Excel workbook for interacting with an instrument that supports the IEEE 488.2 standard with commands such as `*IDN?` and `*CLS`.

## Dependencies

The [cc.isr.Ieee488] workbook depends on the following Workbooks:

* [cc.isr.Core]
* [cc.isr.Winsock]

## References

The following object libraries are used as references:

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]

## Worksheets

The [cc.isr.ieee488] workbook includes two worksheets: _IEEE488_ and _Identity_.

### Identity Worksheet

* The Identity Worksheet is used to query the instrument identity using the `*IDN?` command.

#### Identity Worksheet Testing

Follow this procedure for reading the instrument identity string:

* Select the Identity sheet.
* Enter the instrument dotted IP address, such as `192.168.252`;
* Enter the instrument port:
  * `5025` for an LXI instrument or
  * `1234` for a GPIB instrument connected via a GPIB-Lan controller such as the [Prologix GPIB-Lan controller].
* Check either ___Use VI Session___ or ___Use IEEE 488 Session___ and click __Read Identity___ to read the instrument identity.
	* _VI session_ is the underlying `session` for communicating with the instrument that the _IEEE488 Session_ use for sending commands and query the instrument.

### IEEE 488 Worksheet

* The IEEE 488 Worksheet is used to command and query the instrument using IEEE488.2 commands and queries.

#### Query Unterminated Errors and the GPIB-Lan controller

The GPIB-Lan controller _Read-After-Write_ feature addresses the instrument to talk after sending messages to the instrument.
Instruments such as the Keithley 2700 Scanning Multimeter throw Query Unterminated errors when addressed to talk when not 
having data to send. This can be addressed by turning off _Read-After-Write_ and using the controller's `++read` command for reading from the instrument. 

Turning _Read-After-Write_ on addresses the instrument to talk and, therefore, could could cause a Query Unterminated error. 

Here are some issues to keep in mind when using the IEEE488 test sheet:

* By default, the Controller is initialized with _Read-After-Write_ turned off.
	* Thus, the _Read-After-Write_ state is `False` upon connecting to the instrument.
	* Internally, the program uses the controller's `++read` command to get the readings from the instrument. 
* Toggling _Read-After-Write_ may cause Query Unterminated errors.
* Following instrument errors, commands, which check the status byte for errors, would fail to run because of the error status of the instrument.
* Issuing the `*CLS` command clears this error condition provided the command is appended with `*OPC?`, which turns the command into a query thus avoiding the Query Unterminated error on the bare `*CLS`.
* By default, as implemented by the __CLS__ and __RST__ buttons, the program appends `*OPC?` to its implementation of the `*CLS` and `*RST` commands thus keeping the program in sync with the instrument and avoiding Query Unterminated errors even if the instrument is set for _Read-After-Write_.
* When _Read-After-Write_ is turned on from the test sheet __SET__ command button:
	* The program is set to turn off _Read-After-Write_ on the next `Write` to prevent the Query Unterminated error.
	* The program then updates the state of the _Read-After-Write_ value on the sheet.
	* In other words, with this implementation, instrument communication is largely aimed at avoiding Query Unterminated by turning _Read-After-Write_ off.

#### IEEE 488 Worksheet Testing 

Follow the procedures before connecting, disconnecting and controlling the instrument using the IEEE488 session:

* Enter the instrument dotted IP address such as `192.168.0.252`.
* Enter the instrument port:
  * `5025` for an LXI instrument or
  * `1234` for a GPIB instrument connected via a GPIB-Len controller such as the Prologix controller.
* Depress the ___Toggle Connection___ button to connect the instrument.
	* The instrument connection information such as the _Socket Address_ and _Id_ display at the top row;
	* Control buttons are enabled.
* Release the ___Toggle Connection___ button to disconnect the instrument.
	* Control buttons are diabled.

#### Errors

The last error is displayed to the right of the _Last Error_ row heading.  

Commands issued after an error will be sent to the instrument after clearing the instrument to its known state using the ___CLS___ button.

#### Testing IEEE 488.2 Commands

Follow this procedure to exercise the IEEE 488.2 command:

* Connect the instrument as described above;
* Click the ___RST___ to reset the instrument to its known state. Notice that the reset takes over a second. 
	* Some query commands take a bit longer to execute. The extended time is handled by awaiting for the result for a timeout specified by the session timeout interval, which is different from the socket receive timeout and the GPIB-Lan timeout. 
* Click the ___CLS___ button to clear the instrument to its know state clearing any existing errors;
* Select a command from the ___Command___ drop down list;
	* If a query command, ending with a _?_ is selected, click ___Write___ and then ___Read___ or ___Query___, otherwise click _Read_.
* For example, select the _*IDN?_ command and click ___Query___. The instrument identity should display under the _Received_ heading. 
* The elapsed time for each command is displayed under the _ms_ heading.
* Check the ___Read Status After Write___ check box to automatically read the instrument status byte. 
	* With Tcp control of LXI instruments, the status byte can be queried only after non-query commands. 
	* The GPIB-Lan controller is capable of reading the status byte using _Serial Poll_ even after a query write.
	* The _Read Status After Write_ uses the GPIB-Lan to query the status byte when using the GPIB-Lan controller. 
	* The serial polled value is displayed under _Spoll_ and the value read using ___*STB?___ is displayed under _SRQ_.
	* ___*ESR?___ reads the standard event status which helps determine which event turned on the Requesting Status (RQS) bit of the status register.
	
#### Testing the GPIB_Lan Controller

Follow this procedure to exercise the GPIB-Lan controller:

* Connect the instrument as described above;
* The GPIB-Lan controller buttons are enabled if connecting with the controller on port 1234.
* Once enabled, the command buttons can be used to:
	* ___GTL___: Go to Local sending the instrument to local. The instrument automatically switches to remote on the next command.
	* ___LLO___: Local Lockout to lock the _Local_ instrument button;
	* ___SDC___: Selective device clear;
	* Toggle _Read-After-Write_;
		* Note that if _Read-After-Write_ is `True`, it directs the instrument to 'talk', which automatically sets the instrument to talk after any command. With some instruments (e.g., the Keithley 2700), this causes an instrument Query Unterminated error. This error state lingers until the next `*CLS` command.
	* ___SPOLL___: Serial poll to read the status byte;
	* ___SRQ___: to tell if the Requesting Service signal (Bit 6) of the service request register of the instrument is set;
	* Get or set the _GPIB Address_;
	* Get and set the controller _Read Timeout_ for reading the instrument.
		* Note that the `Ieee488Session` class commands the controller to read the message from the instrument if auto Read-After-Write is turned of. This timeout affects such reading.

[cc.isr.Core]: ./cc.isr.Core.xlsm
[cc.isr.Winsock]: ./cc.isr.Winsock.xlsm
[cc.isr.Ieee488]: ./cc.isr.Ieee488.xlsm
[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
[Prologix GPIB-Lan controller]: https://prologix.biz/product/GPIB-ethernet-controller/

### Testing the cc.isr.Ieee488 Workbook

[cc.isr.Ieee488] is an Excel workbook for interacting with an instrument that supports that IEEE 488.2 standard with commands such as *IDN? and *CLS.

#### Dependencies

The [cc.isr.Ieee488] workbook depends on the following Workbooks:

* [cc.isr.Core]
* [cc.isr.Winsock]

### Worksheets

The [cc.isr.ieee488] workbook includes two worksheets: IEEE488 and Identity.

#### Identity Worksheet

* The Identity Worksheet is used to query the instrument identity using the *IDN? command.

##### Identity Worksheet Testing

Follow this procedure for reading the instrument identity string:

* Select the Identity sheet.
* Enter the instrument dotted IP address.
* Enter the instrument port:
  * 5025 for an LXI instrument or
  * 1234 for a GPIB instrument connected via a GPIB-Len controller such as the Prologix controller.
* Click _Read Identity_ to read the instrument identity.
* Check either _Use VI Session_ or _Use IEEE 488 Session_ and click _Read Identity_.
	* While the VI session is the underlying 'session' for communicating with the instrument, the IEEE488 uses the VI Session to send standard IEEE 488.2 commands and queries.

#### IEEE 488 Worksheet

* The IEEE 488 Worksheet is used to command and query the instrument using IEEE488.2 commands and queries.

##### Query Unterminated Errors and the GPIB-Lan controller

The GPIB-Lan controller Read-After-Write feature addresses the instrument to talk after sending messages to the instrument.
Instruments such as the Keithley 2700 Scanning Multimeter throw Query Unterminated errors when addressed to talk when not 
having data to send. This can be addressed by turning off Read-After-Write and using the controller's `++read` command for reading from the instrument. 
Thus, setting the Read-After-Write from the IEEE488 test sheet could cause Query Unterminated errors. 

Here are some issues to keep in mind when using the IEEE488 test sheet:

* By default, the Controller is initialized with Read-After-Write turned off.
	* Thus, the Read-After-Write state is `False` upon connecting to the instrument.
	* Internally, the program uses the controller's `++read` command to get the readings from the instrument. 
* Upon turning on Read-After-Write, the instrument may issue a Query Unterminated error.
* Following instrument errors, commands, which check the status byte for errors, would fail to run because of the error status of the instrument.
* Issuing the `*CLS` command clears this error condition provided the command is appended with `*OPC?`, which turns the command into a query thus avoiding the Query Unterminated error on the bare `*OPC`.
* By Default, the program appends `*OPC?` to its implementation of the `*CLS` and `*RST` commands thus keeping the program in sync with the instrument.
* If the program is switched to enable Read-After_write from the test sheet:
	* the program is set to turn off Read-After_write on the next `Write` to prevent the Query Unterminated error.
	* The program then updates the state of the Read-After-Write value on the sheet.

##### IEEE 488 Worksheet Testing 

Follow this procedure for reading the instrument identity string:

* Enter the instrument dotted IP address.
* Enter the instrument port:
  * 5025 for an LXI instrument or
  * 1234 for a GPIB instrument connected via a GPIB-Len controller such as the Prologix controller.
* Depress the _Toggle Connection_ button to connect the instrument.
	* The instrument connection information such as the Socket Address and Id display at the top row;
	* Control buttons are enabled.
* click the _RST_ and then the _CLS_ buttons to bring the instrument to its known state clearing any existing errors.
	
###### Errors

The last error is displayed to the right iof the _Last Error_row heading.  

Commands issued after an error will be sent to the instrument after clearing the instrument to its known state using the _CLS_ button.

###### Testing IEEE 488.2 Commands

* Click the RST and then the CLS button to reset the instrument to its know state clearing any existing errors;
* Select a command from the _Command_ drop down list
	* If a query command, ending with a _?_, click _Write_ and then _Read_ or _Query_, otherwise click _Read_.
* For example, select the *IDN? command and click _Query_. The instrument identity should display under the _Received_ heading. 
* The elapsed time for each command is displayed under the _ms_ heading.
* Check the _Read Status After Write_ check box to automatically read the instrument status byte. 
	* With Tcp control of LXI instruments, the status byte can be queried only after non-query commands. 
	* The GPIB-Lan controller is capable of reading the status byte using _Serial Poll_ even after a query write.
	* The _Read Status After Write_ uses the GPIB-Lan to query the status byte when using the GPIB-Lan controller. 
	
###### Testing the GPIB_Lan Controller

The GPIB-Lan controller buttons are enabled if connecting with the controller on port 1234.

Once enabled, the command buttons can be used to:

* Go to Local sending the instrument to local. The instrument automatically switches to remote on the next command.
* Local Lockout to lock the Local instrument button;
* Selective device clear (SDC);
* Toggle Listen and Talk;
	* Note that if the controller is set to Talk, it turns on an internal Read-After_write mode, which automatically sets the instrument to talk after any command. With some instruments (e.g., the Keithley 2700), this causes an instrument Query Unterminated error which lingers until the next `*CLS` command.
* Serial poll to read the status byte;
* SRQ to tell if the Requesting Service signal (Bit 6) of the service request register of the instrument is set;
* Get or set the GPIB address;
* Get and set the controller read timeout for reading the instrument.
	* Note that the Ieee488Session class commands the controller to read the message from the instrument if auto Read-After-Write is turned of. This timeout affect such reading.

[cc.isr.Ieee488]: ./cc.isr.Ieee488.xlsm
[cc.isr.Core]: ./cc.isr.Core.xlsm
[cc.isr.Winsock]: ./cc.isr.Winsock.xlsm

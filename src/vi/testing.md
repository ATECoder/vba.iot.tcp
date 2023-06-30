### Testing the cc.isr.Ieee488 Workbook

[cc.isr.VI] is an Excel workbook for interacting with an instrument that supports that IEEE 488.2 standard with commands such as *IDN? and *CLS.

#### Dependencies

The [cc.isr.VI] workbook depends on the following Workbooks:

* [cc.isr.Core]
* [cc.isr.Winsock]
* [cc.isr.Ieee488]

### Worksheets

The [cc.isr.VI] workbook includes two worksheets: K2700 and Identity.

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

#### K2700 Worksheet

* The K2700 Worksheet is used to command and query the 2700 instrument using SCPI commands and queries.

##### K2700 Worksheet Testing 

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

###### Testing the Resistance Measurements

The K2700 worksheet includes a set of commands and data for measuring resistances that are numbered to match the scan card channels. 

* Click _Read Cards_ to read and display the scanning cards installed in the instrument.
* Click _Set Scans_ to set the scan lists that are used when scanning all cards in sequence.
* Click _Query Inputs_ to read the status of the _Inputs_ button ont he instrument front panel.

The resistance measurements are controlled by command and option buttons as follows:

####### Single read of a specific resistance from the front or rear scanning cards.

* Toggle the Single/Scan button to Single;
* Toggle the Front/Rear button to either Front or Rear;
* Toggle the Manual/Auto button to either manual or auto to select the specified resistance number and optionally automatically increment the resistance after each reading.
* Click _Read R_ to take a single reading from either the front panel or via the scanning cards.

####### Multi-resistance reading from the rear scanning cards.

* Toggle the Single/Scan button to Scan;
* Toggle the Front/Rear button to either Front or Rear;
* Toggle the Manual/Auto button to auto to automatically increment the resistance after each reading.
* Click _Read R_ to take readings from all resistances via the scanning cards.

####### Multi-resistance reading from the rear scanning cards controlled by external triggering.

* Toggle the Single/Scan button to Scan;
* Toggle the Front/Rear button to either Front or Rear;
* Toggle the Manual/Auto button to auto to automatically increment the resistance after each reading.
* Depress _Ext Trig_ to start monitoring the external trigger event.
	* Externally trigger to take sequential measurements on all resistances;
	* Release the _Ext Trig_ button to end monitoring the external trigger event.

[cc.isr.VI]: ./cc.isr.VI.xlsm
[cc.isr.Ieee488]: ./cc.isr.Ieee488.xlsm
[cc.isr.Winsock]: ./cc.isr.Winsock.xlsm
[cc.isr.Core]: ./cc.isr.Core.xlsm

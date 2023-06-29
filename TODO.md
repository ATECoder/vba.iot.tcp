### TODO

All: change by ref to by val.
Core: Add cancel  event arguments class. 
Winsock: Tcp Client: Add Disconnecting events. 
Ieee488: working on initialization.
Default to read after write off.
Add  Connection Closing with cancel event arguments.
On close connection turn off read after write.

#### Tests

#### Fixes
* Send to Device and Send to Controller
	? change to SendMessage and SendMessage ignore read-after-write
	or just leave as is and send RST, SDC, and CLS using send to controller.
* get direct access to the *RST, SDC, and *CLS commands. Turn off RAW with this commands.
* vI: handle error on toggle connection. 

#### Updates
* add test power shell scripts
	* deploy to a bin folder.
	* then run build using relative folders
	* then run test.
	

	

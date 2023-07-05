# Testing the [cc.isr.winsock] Workbook

[cc.isr.winsock] is an Excel workbook implementing TCP Client and Server classes with Windows Winsock API and support higher level [ISR] workbooks.

## Dependencies

The [cc.isr.Winsock] workbook depends on the following Workbook:

* [cc.isr.Core] - Includes core Visual Basic for Applications classes and modules.

## References

The following object libraries are used as references:

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]

## Worksheets

The [cc.isr.Winsock] workbook includes two worksheets: Identity and TestSheet.

* TestSheet - To run unit tests.
* Identity - To query the instrument identity using the *IDN? command.

## Unit Testing

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

### Unit testing with the TestSheet Worksheet

Use the following procedure to run unit tests:
1) Click the ___List Tests___ button.
2) The drop down list now includes the list of available test suites;
3) Select a test from the list;
4) Click ___Run Selected Tests___;
   * The list of tests included in the test suite will display.
   * Passed tests display Passed with a green background;
   * Failed tests display Fail with a red background and a message describing the failure.

## Integration Testing

### Identity Worksheet Testing

Follow this procedure for reading the instrument identity string:

* Select the Identity sheet.
* Enter the instrument dotted IP address, such as `192.168.252`;
* Enter the instrument port:
  * `5025` for an LXI instrument or
  * `1234` for a GPIB instrument connected via a GPIB-Lan controller such as the [Prologix GPIB-Lan controller].
* Click ___Read Identity___ to read the instrument identity using the `*IDN?` query command:
  * Check the following options:
	* ___Using Winsock Read Raw___ -- reads one character at a time till the default termination;
	* ___Using Winsock Buffer Read___ -- reads a buffer of up to 1024 characters at a time;
	* ___Using Tcp Client___ -- reads using the TCP Client class.

[cc.isr.winsock]: https://github.com/ATECoder/vba.iot.tcp/src/winsock
[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>


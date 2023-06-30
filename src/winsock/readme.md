### About

[cc.isr.winsock] is an Excel workbook implementing TCP Client and Server classes with Windows Wisock API.

#### Dependencies

The [cc.isr.Winsock] workbook depends on the following Workbook:

##### [cc.isr.Core]

* [cc.isr.Core] is an Excel workbook with core Visual Basic for Applications code.

#### Worksheets

The [cc.isr.Winsock] workbook includes two worksheets: Identity and TestSheet.

##### Identity Worksheet

* Allows the operator to query the instrument identity using the *IDN? command.

Click _Read Identity_ to read the instrument identity using the *IDN? query command.

Select the options for using socket raw read, which reads one character at a time till the default termination, buffered reading, of reading using the Tcp Client class.

##### TestSheet Worksheet

Allows the operator to run unit tests.

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

Use the following procedure to run unit tests:
1) Click the _List Tests_ button.
2) The drop down list now includes the list of available test suites;
3) Select a test from the list;
4) Click _Run Tests_;
	* The list of tests included in the test suite will display.
	* Passed tests display Passed with a green background;
	* Failed tests display Fail with a red background and a message describing the failure.

### Key Features

* Encapsulates the Windows API to construct the basic objects for Tcp/IP communication.
* Using Windows Winsock32 calls to construct sockets for communicating with the instrument.

### Main Types

The main types provided by this library are:

* _Winsock_ initiates a Winsock session.
* _IPv4StreamSocket_ opens an IPv4 streaming socket to the instrument.
* _TcpCllient_ Encapsulates the _IPv4StreamSocket_.

### Feedback

[cc.isr.winsock] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.winsock] repository.

[cc.isr.winsock]: https://github.com/ATECoder/vba.iot.tcp/src/winsock
[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core

### About

[cc.isr.ieee488] is an Excel workbook for interacting with an instrument that supports that IEEE 488.2 standard with commands such as *IDN? and *CLS.

#### Dependencies

The [cc.isr.Ieee488] workbook depends on the following Workbooks:

##### [cc.isr.Core]

* [cc.isr.Core] is an Excel workbook with core Visual Basic for Applications code.

##### [cc.isr.Winsock]

* [cc.isr.winsock] is an Excel workbook implementing TCP Client and Server classes with Windows Wisock API.

#### Worksheets

The [cc.isr.Ieee488] workbook includes two worksheets: IEEE488 and Identity.

##### Identity Worksheet

* Allows the operator to query the instrument identity using the *IDN? command.

##### IEEE 488 Worksheet

* Allows the operator to command and query the instrument using IEEE488.2 commands and queries.

#### Key Features

* Provides IEEE 488.2 commands and queries for communicating with the instrument.
* Uses Windows Winsock32 calls to construct sockets for communicating with the instrument by way of a GPIB-Lan controller.
* Provides GPIB-Lan commands and queries for communicating with the GPIB-Lan controller.

#### Main Types

The main types provided by this library are:

* _GpibLanController_: Communicates with the instrument by way of a GPIB-Lan controller.
* _ViSession_ Uses a _TcpCllient_ to communicate with the instrument by sending and receiving messages by way of the GPIB-Lan controller.
* _IEEE488Session_ implement the core IEEE488 methods for communicating with an LXI Instrument.

#### [Testing]

Testing information is included in the [Testing] document.

#### Feedback

[cc.isr.Ieee488] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.Ieee488] repository.

[cc.isr.Ieee488]: https://github.com/ATECoder/vba.iot.tcp/src/ieee488
[cc.isr.Winsock]: https://github.com/ATECoder/vba.iot.tcp/src/Winsock
[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core
[Testing]: ./testing.md

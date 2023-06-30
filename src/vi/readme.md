### About

[cc.isr.VI] is an Excel workbook for interacting with SCPI based instruments.

Presently supported is the Keithley 2700 instrument either as an LXI instrument or a GPIB instrument by way of a GPIB-Lan controller such as the [Prologix] GPIB to LAN device.

#### Dependencies

The [cc.isr.VI] workbook depends on the following Workbooks:

##### [cc.isr.Core]

* [cc.isr.Core] is an Excel workbook with core Visual Basic for Applications code.

##### [cc.isr.Winsock]

* [cc.isr.winsock] is an Excel workbook implementing TCP Client and Server classes with Windows Wisock API.

##### [cc.isr.Ieee488]

* [cc.isr.ieee488] is an Excel workbook for interacting with an instrument that supports that IEEE 488.2 standard with commands such as *IDN? and *CLS.

#### Worksheets

The [cc.isr.VI] workbook includes two worksheets: K2700 and Identity.

##### Identity Worksheet

* Allows the operator to query the instrument identity using the *IDN? command.

##### K2700 Worksheet

* Allows the operator to command and query the instrument using IEEE488.2 commands and queries.

### Key Features

* Provides IEEE 488.2 commands and queries for communicating with the instrument.
* Uses Windows Winsock32 calls to construct sockets for communicating with the instrument by way of a GPIB-Lan controller.
* Provides GPIB-Lan commands and queries for communicating with the GPIB-Lan controller.
* Provides an extended sets of methods to control the Keithley 2700 instrument.
* Provides a custom sets of methods to control the Keithley 2700 instrument for measuring 4-wire resistances from the front or read panel using internal or external triggers.

### Main Types

The main types provided by this library are:

* _K2700_ Implements some basic 2700 scanning multimeter functionality.

### Testing

Testing information is included in the [Testing] document.

### Feedback

[cc.isr.vi] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.vi] repository.

[cc.isr.vi]: https://github.com/ATECoder/vba.iot.tcp/src/vi
[cc.isr.Ieee488]: https://github.com/ATECoder/vba.iot.tcp/src/ieee488
[cc.isr.Winsock]: https://github.com/ATECoder/vba.iot.tcp/src/Winsock
[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core
[Testing]: ./testing.md
[Prologix]: https://prologix.biz/product/gpib-ethernet-controller/


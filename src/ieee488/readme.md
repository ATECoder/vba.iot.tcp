# About

[cc.isr.Ieee488] is an Excel workbook for interacting with an instrument that supports the IEEE 488.2 standard with commands such as `*IDN?` and `*CLS` and supporting higher level [ISR] workbooks

## Dependencies

The [cc.isr.Ieee488] workbook depends on the following Workbooks:

* [cc.isr.Core] - Includes core Visual Basic for Applications classes and modules.
* [cc.isr.Winsock] - Implements TCP Client and Server classes with Windows Winsock API.

## References

The following object libraries are used as references:

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]

## Worksheets

The [cc.isr.Ieee488] workbook includes two worksheets:

* Identity -- To query the instrument identity using the *IDN? command.
* IEEE488  -- To command and query an IEEE488.2 instrument.

## Key Features

* Provides commands and queries for communicating with IEEE488.2 instrument.
* Uses Windows Winsock32 calls to construct sockets for communicating with the instrument by way of a GPIB-Lan controller such as the [Prologix GPIB-Lan controller].
* Provides GPIB-Lan commands and queries for communicating with the GPIB-Lan controller.

## Main Types

The main types provided by this library are:

* _GpibLanController_ -- Communicates with the instrument by way of a GPIB-Lan controller.
* _ViSession_ -- Uses a _TcpCllient_ to communicate with the instrument by sending and receiving messages by way of the GPIB-Lan controller.
* _IEEE488Session_ -- Implements the core methods for communicating with an IEEE488.2 Instrument.

## [Testing]

Testing information is included in the [Testing] document.

## Feedback

[cc.isr.Ieee488] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.Ieee488] repository.

[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core
[cc.isr.Winsock]: https://github.com/ATECoder/vba.iot.tcp/src/Winsock
[cc.isr.Ieee488]: https://github.com/ATECoder/vba.iot.tcp/src/ieee488
[Testing]: ./cc.isr.ieee488.testing.md
[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
[Prologix GPIB-Lan controller]: https://prologix.biz/product/GPIB-ethernet-controller/
[ISR]: https://www.integratedscientificresources.com
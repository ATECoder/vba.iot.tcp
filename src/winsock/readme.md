# About

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

# Key Features

* Encapsulates the Windows API to construct the basic objects for Tcp/IP communication.
* Using Windows Winsock32 calls to construct sockets for communicating with the instrument.

# Main Types

The main types provided by this library are:

* _Winsock_ - initiates a Winsock session.
* _IPv4StreamSocket_ - opens an IPv4 streaming socket to the instrument.
* _TcpCllient_ - Encapsulates the _IPv4StreamSocket_.

## [Testing]

Testing information is included in the [Testing] document.

## Scripts

* Build: copies files to the build folder and remove the existing references.
* Deploy: copies files to the build folder.

# Feedback

[cc.isr.winsock] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.winsock] repository.

[cc.isr.winsock]: https://github.com/ATECoder/vba.iot.tcp/src/winsock
[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core
[Testing]: ./cc.isr.winsock.testing.md

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>


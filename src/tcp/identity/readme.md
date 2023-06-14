### About

[vba.iot.tcp.identity] is an Excel workbook for reading the identity string from an LXI instrument.

### How to Use

* Open the Excel file.
* Select the Identity sheet.
* Enter the instrument IP address.
* Enter the instrument port
  * 5025 for an LXI instrument or
  * 1234 for a GOPIB instrument connected via a Prologix controller.
* Click _Read Identity_ to read the instrument identity.  


### Key Features

* Provides rudimentary SCPI methods for reading the instrument identity.
* Using WIndows Winsock32 calls to constrack sockets for communicating witht he instrument.

### Main Types

The main types provided by this library are:

* _Winsock_ initiates a Winsock session.
* _IPv4StreamSocket_ opens an IPv4 streaming socket to the instrument.
* _TcpCllient_ Encapsulates the _IPv4StreamSocket_.
* _ViSession_ Encapsulates the _TcpCllient_ and defines termination and timeout.
* _IEEE488Session_ implement the core IEEE488 methods for communicating with an LXI Instrument.

### Testing

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

### Feedback

[vba.iot.tcp.identity] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [vba.iot.tcp.identity] repository.

[vba.iot.tcp.identity]: https://github.com/ATECoder/vba.iot.tcp

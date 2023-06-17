### About

[cc.isr.winsock] is an Excel workbook implementing TCP Client and Server with Windows WIsock API.

### How to Use

A test sheet provides an example for reading an identity string from an LXI instrument:
* Open the Excel file.
* Select the Identity sheet.
* Enter the instrument IP address.
* Enter the instrument port
  * 5025 for an LXI instrument or
  * 1234 for a GPIB instrument connected via a Prologix controller.
* Click _Read Identity_ to read the instrument identity.  


### Key Features

* Encapsulates the Windows API to construct the basic objects for Tcp/IP communication.
* Using Windows Winsock32 calls to construct sockets for communicating with the instrument.

### Main Types

The main types provided by this library are:

* _Winsock_ initiates a Winsock session.
* _IPv4StreamSocket_ opens an IPv4 streaming socket to the instrument.
* _TcpCllient_ Encapsulates the _IPv4StreamSocket_.

### Testing

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

### Feedback

[cc.isr.winsock] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.winsock] repository.

[cc.isr.winsock]: https://github.com/ATECoder/vba.iot.tcp/src/winsock

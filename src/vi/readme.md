### About

[cc.isr.vi] is an Excel workbook for reading the identity string from a set of LXI instruments. 

Presently supported is the Keithley 2700 instrument either as an LXI instrument of a GPIB instrument by way of the [Prologix] GPIB to LAN device.

### How to Use

* Open the Excel file.
* Select the Identity sheet.
* Enter the instrument IP address.
* Enter the instrument port
  * 5025 for an LXI instrument or
  * 1234 for a GPIB instrument connected via a Prologix controller.
* Click _Read Identity_ to read the instrument identity.  


### Key Features

* Provides rudimentary SCPI methods for reading the instrument identity.
* Using Windows Winsock32 calls to construct sockets for communicating with the instrument.

### Main Types

The main types provided by this library are:

* _K2700_ Implements some basic 2700 scanning multimeter functionality.

### Testing

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

### Feedback

[cc.isr.vi] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.vi] repository.

[cc.isr.vi]: https://github.com/ATECoder/vba.iot.tcp/src/vi
[Prologix]: https://prologix.biz/product/gpib-ethernet-controller/

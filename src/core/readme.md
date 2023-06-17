### About

[cc.isr.Core] is an Excel workbook with core Visual Basic for Applications code.

### How to Use

Use this workbook as a reference in other workbooks.

### Key Features

* Provides core classes.
* Provides a rudimentary test executive.

### Main Types

The main types provided by this library are:

* _Assert_ returns results from unit tests.
* _CollectionExtensions_ Singleton. Collection extensions.
* _Marshal_ Singleton. Supports Endianess.
* _PathExtensions_ Singleton. Build builder with  file and folder deletion and existence methods.
* _StopWatch_ high resolution stop watch using the Windows API.
* _StringBuilder_ A fast string builder.
* _StringExtensions_ Singleton. String extensions.
* _stdTimer_ A thread aware timer.
* _TestExecutive_ Singleton. A rudimentary unit test executive.
* _UserDefinedError_ A user defined error class.
* _WorkbookUnititiels_ Singleton. Exports code files and enumerates test methods.

### Testing

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

### Feedback

[cc.isr.Core] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.Core] repository.

[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core

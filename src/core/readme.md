### About

[cc.isr.Core] is an Excel workbook with core Visual Basic for Applications code.

#### Dependencies

The [cc.isr.Core] has no additional dependencies.

#### Worksheets

The [cc.isr.Core] workbook includes two worksheets: TestSheet and Countdown Timer.

##### TestSheet Worksheet

* Allows the operator to run unit tests.

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

Use the following procedure to run unit tests:
1) Click the _List Tests_ button.
2) The drop down list now includes the list of available test suites;
3) Select a test from the list;
4) Click _Run Tests_;
	* The list of tests included in the test suite will display.
	* Passed tests display Passed with a green background;
	* Failed tests display Fail with a red background and a message describing the failure.

##### Countdown Timer Worksheet

* Allows the operator to test the `clsTimer` class.

#### How to Use

Typically, the [cc.isr.Core] Workbook is added as a reference in other workbooks.

#### Key Features

* Provides core classes such as `clsTimer` and 'StringBuilder'.
* Provide extension classes such as `StringExtensions` and`PathExtensions`.
* Provides a rudimentary test executive.

### Main Types

The main types provided by this library are:

* _Assert_ returns results from unit tests.
* _CanceEventArg_ event arguments for canceling event handlers.
* _CollectionExtensions_ Singleton. Collection extensions.
* _MacroInfo_ holds information such as name and module name about Excel Macro methods.
* _Marshal_ Singleton. Supports Endianess.
* _ModuleInfo_ holds information such as name and project name about Excel modules.
* _PathExtensions_ Singleton. Build builder with  file and folder deletion and existence methods.
* _stdTimer_ a timer class capable of issuing events with millisecond time resolution.
* _StopWatch_ high resolution stop watch using the Windows API.
* _StringBuilder_ A fast string builder.
* _StringExtensions_ Singleton. String extensions.
* _TestExecutive_ Singleton. A rudimentary unit test executive.
* _UserDefinedError_ A user defined error class.
* _WorkbookUnititiels_ Singleton. Exports code files and enumerates test methods.

### Testing

The project `TestSheet` includes commands buttons for running build-in unit tests of the core classes.

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

### Feedback

[cc.isr.Core] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.Core] repository.

[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core

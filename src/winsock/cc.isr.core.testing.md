# Testing the [cc.isr.Core] Workbook

[cc.isr.Core] is an Excel workbook with code Visual Basic for Applications modules and classes that support [ISR] workbooks.

## Dependencies

The [cc.isr.Core] has no additional dependencies.

## References

The following object libraries are used as references:

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]

## Worksheets

The [cc.isr.Core] workbook includes two worksheets: 

* TestSheet -- To run unit tests.
* Countdown Timer -- To test the `EventTimer` class.

## Unit Testing

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

### Unit testing with the TestSheet Worksheet

Use the following procedure to run unit tests:
1) Click the ___List Tests___ button.
2) The drop down list now includes the list of available test suites;
3) Select a test from the list;
4) Click ___Run Selected Tests___;
   * The list of tests included in the test suite will display.
   * Passed tests display Passed with a green background;
   * Failed tests display Fail with a red background and a message describing the failure.

## Integration Testing

### Testing the EventTimer using the Countdown Timer Worksheet

Use the following procedure to run the `EventTimer` tests:

1) Select the _Countdown Timer_ Worksheet;
2) Click ___Reset Timer___ button to initialize the timer duration to 15;
3) Click ___Start Timer___ to start the countdown. after a short pause, the display with decrement at fractions of a second intervals;
4) Click ___Stop Timer___ to pause the timer;
5) Click ___Dispose___ to stop and terminate the timer.
6) Close the Excel workbook and check the task manager to ensure that all Excel instances are closes.

To validate item 6, make sure to start the test session with only a single Excel instance.

[cc.isr.Core] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.Core] repository.

[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>

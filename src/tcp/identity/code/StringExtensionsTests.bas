Attribute VB_Name = "StringExtensionsTests"
Option Explicit

''' <summary> Unit test. Asserts trim left. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestTrimLeft() As Assert
    Set TestTrimLeft = Assert.AreEqual("bar", StringExtensions.TrimLeft("oobar", "o"), "left-trims strings")
End Function

''' <summary> Unit test. Asserts trim right. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestTrimRight() As Assert
    Set TestTrimRight = Assert.AreEqual("f", StringExtensions.TrimRight("foo", "o"), "right-trims strings")
End Function

''' <summary> Unit test. Asserts start with. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestStartsWith() As Assert
    Set TestStartsWith = Assert.IsTrue(StringExtensions.StartsWith("foobar", "foo"), "detects string starts")
End Function

''' <summary> Unit test. Asserts end width. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestEndsWith() As Assert
    Set TestEndsWith = Assert.IsTrue(StringExtensions.EndsWith("foobar", "bar"), "detects string ends")
End Function

''' <summary> Unit test. Asserts character at an index position. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestCharAt() As Assert
    Set TestCharAt = Assert.AreEqual("a", StringExtensions.CharAt("foobar", 5), "gets chars from strings")
End Function

''' <summary> Unit test. Asserts sub-string. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestSubstring() As Assert
    Set TestSubstring = Assert.AreEqual("oo", StringExtensions.Substring("foobar", 1, 2), "gets parts from strings")
End Function

''' <summary> Unit test. Asserts creating a repeated string. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestRepeat() As Assert
    Set TestRepeat = Assert.AreEqual("aaa", StringExtensions.Repeat("a", 3), "repeats strings")
End Function



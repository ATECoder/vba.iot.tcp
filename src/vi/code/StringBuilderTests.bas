Attribute VB_Name = "StringBuilderTests"
Option Explicit

''' <summary>   Unit test. Tests appending items to string builder. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestAppendingToEmptyBuilder() As Assert
    Dim builder As StringBuilder
    Set builder = New StringBuilder
    Dim expected As String
    expected = "a"
    builder.Append expected
    Set TestAppendingToEmptyBuilder = Assert.AreEqual(expected, builder.ToString, _
            "Appended value should equal expected value")
End Function

''' <summary>   Unit test. Tests appending an empty string to the string builder. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestAppendingEmptyString() As Assert
    Dim builder As StringBuilder
    Set builder = New StringBuilder
    Dim expected As String
    expected = ""
    builder.Append expected
    Set TestAppendingEmptyString = Assert.AreEqual(expected, builder.ToString, _
            "Appended empty value should equal expected value")
End Function

''' <summary>   Unit test. Tests appending a long string to the string builder. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestAppendingLongString() As Assert
    Dim builder As StringBuilder
    Set builder = New StringBuilder
    Dim expected As String
    expected = StringExtensions.Repeat("a", 1000)
    builder.Append expected
    Set TestAppendingLongString = Assert.AreEqual(expected, builder.ToString, _
            "Appended a long value should equal expected value")
End Function

Public Function TestAppendingLineFeed() As Assert
    Dim builder As StringBuilder
    Set builder = New StringBuilder
    Dim expected As String
    expected = "a" & vbLf
    builder.Append expected
    Set TestAppendingLineFeed = Assert.AreEqual(expected, builder.ToString, _
            "Appended value with line feed should equal expected value")
End Function




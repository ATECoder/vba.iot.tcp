VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' <summary>   Indicates whether the assertion was successful. </summary>
Public AssertSuccessful As Boolean

''' <summary>   The outcome message of this assertion. </summary>
Public AssertMessage As String

''' <summary>   Asserts that values are equal. If not, this will return a failed assertion. </summary>
''' <param name="expected">   The value the caller expects to have. </param
''' <param name="actual">   The actual value. </param
''' <param name="message">   A message to show in case the assertion failed. </param
''' <returns>   An <see cref="Assert"/>   instant of True if values are equal; otherwise, False. </returns>
Public Function AreEqual(ByRef expected, ByRef actual, ByRef message As String) As Assert
    Set AreEqual = IsTrue(Not expected <> actual, message)
    
    If Not AreEqual.AssertSuccessful Then
        AreEqual.AssertMessage = AreEqual.AssertMessage & vbCrLf & _
                "    Expected: " & expected & vbCrLf & _
                "    Actual:   " & actual
    End If
End Function

''' <summary>   Asserts that values are not equal. If not, this will return a failed assertion. </summary>
''' <param name="expected">   The value the caller expects to have. </param
''' <param name="actual">   The actual value. </param
''' <param name="message">   A message to show in case the assertion failed. </param
''' <returns>   True if values are not equal; otherwise, false. </returns>
Public Function AreNotEqual(ByRef expected, ByRef actual, ByRef message As String) As Assert
    Set AreNotEqual = IsTrue(expected <> actual, message)
    
    If Not AreNotEqual.AssertSuccessful Then
        AreNotEqual.AssertMessage = AreNotEqual.AssertMessage & vbCrLf & _
                "    Expected: " & expected & vbCrLf & _
                "    Actual:   " & actual
    End If
End Function

''' <summary>   Asserts that the given value is null. </summary>
''' <param name="actual">   The actual value. </param
''' <param name="message">   A message to show in case the assertion failed. </param
''' <returns>   True if <paramref name="actual"/>   is null; otherwise, false. </returns>
Public Function IsNull(ByRef actual, ByRef message As String) As Assert
    Set IsNull = IsFalse(IsObject(actual), message)
End Function

''' <summary>   Asserts that the given value is not null. </summary>
''' <param name="actual">   The actual value. </param
''' <param name="message">   A message to show in case the assertion failed. </param
''' <returns>   True if <paramref name="actual"/>   is not null; otherwise, false. </returns>
Public Function IsNotNull(ByRef actual, ByRef message As String) As Assert
    Set IsNotNull = IsTrue(IsObject(actual), message)
End Function

''' <summary>   Asserts that the given value is true. </summary>
''' <param name="truthy">   A value that can be seen as truthy. Very likely should be boolean,
'''                       but maybe you like to live dangerously. </param
''' <param name="message">   A message to show in case the assertion failed. </param
''' <returns>   True if <paramref name="truthy"/>   is true; otherwise, false. </returns>
Public Function IsTrue(ByRef truthy, ByRef message As String) As Assert
    Set IsTrue = New Assert
    IsTrue.AssertSuccessful = True
    IsTrue.AssertMessage = message

    If Not truthy Then
        IsTrue.AssertSuccessful = False
        IsTrue.AssertMessage = "Assert failed: " & message
    End If
End Function

''' <summary>   Asserts that the given value is False. </summary>
''' <param name="falsy">   A value that can be seen as falsy. Very likely should be boolean,
'''                       but maybe you like to live dangerously. </param
''' <param name="message">   A message to show in case the assertion failed. </param
''' <returns>   False if <paramref name="falsy"/>   is False; otherwise, false. </returns>
Public Function IsFalse(ByRef falsy, ByRef message As String) As Assert
    Set IsFalse = New Assert
    IsFalse.AssertSuccessful = True
    IsFalse.AssertMessage = message

    If falsy Then
        IsFalse.AssertSuccessful = False
        IsFalse.AssertMessage = "Assert failed: " & message
    End If
End Function



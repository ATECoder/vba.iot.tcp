VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' <summary> Indicates whether the assertion was successful. </summary>
Public AssertSuccessful As Boolean

''' <summary> The outcome message of this assertion. </summary>
Public AssertMessage As String

''' <summary> Asserts that both values are equal. If not, this will return a failed assertion. </summary>
''' <param name="expected"> The value the caller expects to have. </param
''' <param name="actual"> The actual value. </param
''' <param name="message"> A message to show in case the assertion failed. </param
''' <returns> True if values are equal; otherwise, false. </returns>
Public Function AreEqual(ByRef expected, ByRef actual, ByRef message As String) As Assert
        Set AreEqual = IsTrue(Not expected <> actual, message)
        
        If Not AreEqual.AssertSuccessful Then
            AreEqual.AssertMessage = AreEqual.AssertMessage & vbCrLf & "    Expected: " & expected & vbCrLf & "    Actual:   " & actual
        End If
End Function


''' <summary> Asserts that the given value is true. </summary>
''' <param name="truthy"> A value that can be seen as truthy. Very likely should be boolean,
'''                       but maybe you like to live dangerously. </param
''' <param name="message"> A message to show in case the assertion failed. </param
''' <returns> True if <paramref name="truthy"/> is true; otherwise, false. </returns>
Public Function IsTrue(ByRef truthy, ByRef message As String) As Assert
    Set IsTrue = New Assert
    IsTrue.AssertSuccessful = True
    IsTrue.AssertMessage = message

    If Not truthy Then
        IsTrue.AssertSuccessful = False
        IsTrue.AssertMessage = "  Assert failed: " & message
    End If
End Function



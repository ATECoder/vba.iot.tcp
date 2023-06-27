Attribute VB_Name = "UserDefinedErrorsTests"
Option Explicit

''' <summary>   Unit test. Asserts the existing of a user defined error. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestUserDefinedErrorShouldExist() As Assert
    ' this should be added to the activate event of the workbook
    ' UserDefinedErrors.Initialize
    Dim ude As UserDefinedError
    Set ude = UserDefinedErrors.SocketConnectionError
    Set TestUserDefinedErrorShouldExist = Assert.IsTrue(UserDefinedErrors.UserDefinedErrorExists(ude), _
                                                        ude.ToString(" should exist"))
End Function

Public Function TestErrorMessageShouldBuild() As Assert

    Const thisProcedureName = "TestErrorMessageShouldBuild"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' create an error
    Dim zero As Double: zero = 0
    Dim value As Double: value = 1 / zero
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, "UserDefinedErrorsTests"
    
    ' display the error message
    Dim errorMessage As String: errorMessage = UserDefinedErrors.BuildStandardErrorMessage()
    
    Set TestErrorMessageShouldBuild = Assert.IsTrue(Len(errorMessage) > 0, "error message should build")
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function




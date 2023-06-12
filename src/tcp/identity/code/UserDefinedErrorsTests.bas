Attribute VB_Name = "UserDefinedErrorsTests"
Option Explicit

''' <summary> Unit test. Asserts the existing of a user defined error. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestUserDefinedErrorShouldExist() As Assert
    ' this should be added to the activate event of the workbook
    ' UserDefinedErrors.Initialize
    Dim ude As UserDefinedError
    Set ude = UserDefinedErrors.SocketConnectionError
    Set TestUserDefinedErrorShouldExist = Assert.IsTrue(UserDefinedErrors.UserDefinedErrorExists(ude), _
                                                        ude.ToString(" should exist"))
End Function


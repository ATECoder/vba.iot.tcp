Attribute VB_Name = "UserDefinedErrorsTests"
Public Function TestUserDefinedErrorShouldExist() As Assert
    UserDefinedErrors.Initialize
    Dim ude As UserDefinedError
    Set ude = UserDefinedErrors.SocketConnectionError
    Set TestUserDefinedErrorShouldExist = Assert.IsTrue(UserDefinedErrors.UserDefinedErrorExists(ude), _
                                                        ude.ToString(" should exist"))
End Function


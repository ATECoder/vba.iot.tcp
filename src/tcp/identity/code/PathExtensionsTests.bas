Attribute VB_Name = "PathExtensionsTests"
Option Explicit

''' <summary>   Unit test. Asserts that the path elements should join. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestPathElementsShouldJoin() As Assert
    Dim element1 As String: element1 = ActiveWorkbook.path
    Dim element2 As String: element2 = "folder2"
    Dim element3 As String: element3 = "folder3"
    Dim expected As String: expected = element1 & "\" & element2 & "\" & element3
    Set TestPathElementsShouldJoin = Assert.AreEqual(expected, PathExtensions.PathJoin(element1, element2, element3), _
            "The path element should be joined")
End Function


Attribute VB_Name = "PathExtensionsTests"
Option Explicit

''' <summary>   Unit test. Asserts that the path elements should join and create the directory. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestPathElementsShouldJoin() As Assert

    Dim element1 As String: element1 = ActiveWorkbook.path
    Dim element2 As String: element2 = "dummy"
    Dim element3 As String: element3 = "workbook"
    Dim FileName As String: FileName = "filename.txt"
    
    ' test joining without creating
    
    Dim expectedDummyPath As String: expectedDummyPath = element1 & "\" & element2
    Dim actualDummyPath As String: actualDummyPath = PathExtensions.Join(element1, element2)
    Set TestPathElementsShouldJoin = Assert.AreEqual(expectedDummyPath, actualDummyPath, "The path elements should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    Dim expectedPath As String: expectedPath = element1 & "\" & element2 & "\" & element3
    Dim expectedFilePath As String: expectedFilePath = expectedPath & "\" & FileName
   
    ' test joining without creating
    
    Dim actualPath As String: actualPath = PathExtensions.JoinAll(False, element1, element2, element3)
    Set TestPathElementsShouldJoin = Assert.AreEqual(expectedPath, actualPath, "The path elements should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test joining a file.
    
    Dim actualFilePath As String: actualFilePath = PathExtensions.JoinFile(actualPath, FileName)
    Set TestPathElementsShouldJoin = Assert.AreEqual(expectedFilePath, actualFilePath, _
        "The path path should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the folder if it exists.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.DeleteFolder(actualPath), _
        "The path " & actualPath & " should no longer exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test joining and creating.
    
    actualPath = PathExtensions.JoinAll(True, element1, element2, element3)
    
    Set TestPathElementsShouldJoin = Assert.AreEqual(expectedPath, actualPath, _
        "The path element should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test detecting the created folder.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.FolderExists(actualPath), _
        "The path " & actualPath & " should exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test creating the file.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.CreateTextFile(actualFilePath), _
        "The file " & actualFilePath & " should exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test checking if a file exists.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.FileExists(actualFilePath), _
        "The file " & actualFilePath & " should exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the file if it exists.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.DeleteFile(actualFilePath), _
        "The file " & actualFilePath & " should no longer exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the folder.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.DeleteFolder(actualPath), _
        "The path " & actualPath & " should no longer exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the dummy folder.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.DeleteFolder(actualDummyPath), _
        "The path " & actualDummyPath & " should no longer exist")

End Function


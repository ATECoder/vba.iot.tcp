Attribute VB_Name = "TestExecutiveTests"
''' <summary> Defines a test file handle. </summary>
Public Type TestFileHandle
    TestFilename As String
    TestFileStream As TextStream
End Type

''' <summary> Unit test. Asserts creating a list of test modules. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestModuleList() As Assert
    Dim modules As Collection
    Set modules = WorkbookUtilities.ListTestModules
    
    ' this includes all modules that start with test.
    Dim knownTestModules As Collection
    Set knownTestModules = New Collection
    knownTestModules.Add "StopWatchTests"
    knownTestModules.Add "StringExtensionsTests"
    knownTestModules.Add "TestExecutiveTests"
    
    Set TestElapsedTime = Assert.AreEqual(knownTestModules.count, modules.count, _
        "Expecting " & CStr(knownTestModules.count) & " but found  " & CStr(modules.count) & " test modules")
    
    If Not TestElapsedTime.AssertSuccessful Then
        Return
    End If
    
    Set TestElapsedTime = Assert.IsTrue(CollectionExtensions.ContainsAll(modules, knownTestModules), _
        "listed test modules do not contain all the know test modules")
  
End Function

''' <summary> Unit test. Asserts creating a list of test macros. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestMacroList() As Assert
    Dim modules As Collection
    Set modules = WorkbookUtilities.ListTestModules()
    
    ' this includes all modules that start with test.
    Dim knownTestModules As Collection
    Set knownTestModules = New Collection
    knownTestModules.Add "StopWatchTests"
    knownTestModules.Add "StringExtensionsTests"
    knownTestModules.Add "TestExecutiveTests"
    
    Set TestElapsedTime = Assert.AreEqual(knownTestModules.count, modules.count, _
        "Expecting " & CStr(knownTestModules.count) & " but found  " & CStr(modules.count) & " test modules")
    
    If Not TestElapsedTime.AssertSuccessful Then
        Return
    End If
    
    Set TestElapsedTime = Assert.IsTrue(CollectionExtensions.ContainsAll(modules, knownTestModules), _
        "listed test modules do not contain all the know test modules")
  
End Function





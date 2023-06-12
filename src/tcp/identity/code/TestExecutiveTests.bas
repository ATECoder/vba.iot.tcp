Attribute VB_Name = "TestExecutiveTests"
Option Explicit

''' <summary> Defines a test file handle. </summary>
Public Type TestFileHandle
    TestFilename As String
    TestFileStream As TextStream
End Type

''' <summary> Adds the test modules. </summary>
Private Sub AddTestModules(knownTestModules As VBA.collection)
    knownTestModules.Add "IPv4StreamSocketTests"
    knownTestModules.Add "StopWatchTests"
    knownTestModules.Add "StringExtensionsTests"
    knownTestModules.Add "TestExecutiveTests"
    knownTestModules.Add "UserDefinedErrorsTests"
    knownTestModules.Add "WinsockTests"
End Sub

''' <summary> Unit test. Asserts creating a list of test modules. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestModuleList() As Assert
    Dim modules As VBA.collection
    Set modules = WorkbookUtilities.ListTestModules
    
    ' this includes all modules that start with test.
    Dim knownTestModules As VBA.collection
    Set knownTestModules = New VBA.collection
    AddTestModules knownTestModules
    
    Set TestModuleList = Assert.AreEqual(knownTestModules.count, modules.count, _
        "Expecting " & CStr(knownTestModules.count) & " but found  " & CStr(modules.count) & " test modules")
    
    If Not TestModuleList.AssertSuccessful Then
        Exit Function
    End If
    
    Set TestModuleList = Assert.IsTrue(CollectionExtensions.ContainsAll(modules, knownTestModules), _
        "listed test modules do not contain all the know test modules")
  
End Function

''' <summary> Unit test. Asserts creating a list of test macros. </summary>
''' <returns> An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestMacroList() As Assert
    Dim modules As VBA.collection
    Set modules = WorkbookUtilities.ListTestModules()
    
    ' this includes all modules that start with test.
    Dim knownTestModules As VBA.collection
    Set knownTestModules = New VBA.collection
    AddTestModules knownTestModules
    
    Set TestMacroList = Assert.AreEqual(knownTestModules.count, modules.count, _
        "Expecting " & CStr(knownTestModules.count) & " but found  " & CStr(modules.count) & " test modules")
    
    If Not TestMacroList.AssertSuccessful Then
        Exit Function
    End If
    
    Set TestMacroList = Assert.IsTrue(CollectionExtensions.ContainsAll(modules, knownTestModules), _
        "listed test modules do not contain all the know test modules")
  
End Function





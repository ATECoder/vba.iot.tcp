Attribute VB_Name = "WorkbookUtilitiesTests"
Option Explicit

''' <summary>   Adds the test modules. </summary>
Private Sub AddTestModules(ByVal knownTestModules As VBA.Collection)
    knownTestModules.Add "cc_isr_core.CollectionExtensionsTests"
    knownTestModules.Add "cc_isr_core.MarshalTests"
    knownTestModules.Add "cc_isr_core.PathExtensionsTests"
    knownTestModules.Add "cc_isr_core.StopWatchTests"
    knownTestModules.Add "cc_isr_core.StringBuilderTests"
    knownTestModules.Add "cc_isr_core.StringExtensionsTests"
    knownTestModules.Add "cc_isr_core.UserDefinedErrorsTests"
    knownTestModules.Add "cc_isr_core.WorkbookUntilitiesTests"
End Sub

''' <summary>   Unit test. Asserts creating a list of test modules. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestModuleList() As Assert

    Dim modules As VBA.Collection
    Set modules = WorkbookUtilities.EnumerateProjectModules(Application.ActiveWorkbook.VBProject)
    
    ' this includes all modules that start with test.
    Dim knownTestModules As VBA.Collection
    Set knownTestModules = New VBA.Collection
    AddTestModules knownTestModules
    
    Set TestModuleList = Assert.AreEqual(knownTestModules.count, modules.count, _
        "Expecting " & CStr(knownTestModules.count) & " but found  " & CStr(modules.count) & " test modules")
    
    If Not TestModuleList.AssertSuccessful Then
        Exit Function
    End If
    
    Set TestModuleList = Assert.IsTrue(CollectionExtensions.ContainsAll(modules, knownTestModules), _
        "listed test modules do not contain all the known test modules")
  
End Function


Attribute VB_Name = "WorkbookUtilitiesTests"
Option Explicit

Private Sub AddModule(ByVal col As VBA.Collection, ByVal moduleFullName As String)
    Dim module As ModuleInfo
    Set module = Constructor.CreateModuleInfo
    module.FromModuleFullName moduleFullName
    col.Add module
End Sub

Public Function ContainsModule(ByVal col As VBA.Collection, ByVal findModule As ModuleInfo) As Boolean
    Dim found As Boolean
    found = False
    Dim colItem As ModuleInfo
    For Each colItem In col
        DoEvents
        If colItem.Equals(findModule) Then
            found = True
            Exit For
        End If
    Next colItem
    ContainsModule = found
End Function

Private Function ContainsAllModules(ByVal leftCol As VBA.Collection, ByVal rightCol As VBA.Collection)

    Dim result As Boolean: result = False
    Dim rightModule As ModuleInfo
    For Each rightModule In rightCol
        DoEvents
        If Not ContainsModule(leftCol, rightModule) Then
            result = False
            Exit Function
        End If
    Next rightModule
    ContainsAllModules = result

End Function


''' <summary>   Adds the test modules. </summary>
Private Sub AddTestModules(ByVal knownTestModules As VBA.Collection)
    Dim projectName As String: projectName = Application.ActiveWorkbook.VBProject.name
    AddModule knownTestModules, projectName & ".CollectionExtensionsTests"
    AddModule knownTestModules, projectName & ".MarshalTests"
    AddModule knownTestModules, projectName & ".PathExtensionsTests"
    AddModule knownTestModules, projectName & ".StopWatchTests"
    AddModule knownTestModules, projectName & ".StringBuilderTests"
    AddModule knownTestModules, projectName & ".StringExtensionsTests"
    AddModule knownTestModules, projectName & ".UserDefinedErrorsTests"
    AddModule knownTestModules, projectName & ".WorkbookUtilitiesTests"
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
    
    Dim missingItem As Variant: Set missingItem = Nothing
    Set missingItem = CollectionExtensions.FindMissingItem(modules, knownTestModules)
    
    If Not missingItem Is Nothing Then
        Set TestModuleList = Assert.IsTrue(CollectionExtensions.ContainsAll(modules, knownTestModules), _
            "item " & CStr(missingItem) & " from the expected test module is not found in the actual collection of test modules")
        Exit Function
    End If
  
    Set missingItem = CollectionExtensions.FindMissingItem(knownTestModules, modules)
    
    If Not missingItem Is Nothing Then
        Set TestModuleList = Assert.IsTrue(CollectionExtensions.ContainsAll(modules, knownTestModules), _
            "item " & CStr(missingItem) & " from the actual test module is not found in the exected collection of test modules")
        Exit Function
    End If
  
  
End Function


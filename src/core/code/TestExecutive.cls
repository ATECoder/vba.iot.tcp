VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestExecutive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' <summary>   Executes all test macros in the module specified in the test sheet. </summary>
''' <para name="testDataSheet">   Specifies the work sheet were test data are set and logged. </param>
Public Sub Execute(ByRef testDataSheet As Worksheet)
    
    On Error GoTo Finally
    
    Dim moduleFullName As String
   
    Dim messages As New VBA.Collection
    Dim passedCount As Integer
    Dim failedCount As Integer
    Dim row As Integer
    row = 1
    
    testDataSheet.Rows("2:" & testDataSheet.Rows.count).ClearContents
    
    moduleFullName = testDataSheet.Range("B" & row).value
    
    ' get the module info
    Dim module As ModuleInfo: Set module = New ModuleInfo: module.FromModuleFullName moduleFullName
    
    row = row + 1
    testDataSheet.Range("A" & row).value = "Project"
    testDataSheet.Range("B" & row).value = module.projectName
    
    row = row + 1
    testDataSheet.Range("A" & row).value = "Module"
    testDataSheet.Range("B" & row).value = module.moduleName
   
    row = row + 1
    testDataSheet.Range("A" & row).value = "Test Name"
    testDataSheet.Range("B" & row).value = "Outcome"
    
    Dim macros As Collection: Set macros = WorkbookUtilities.EnumerateMacros(module)
    Dim macro As MacroInfo
    
    Dim sw As StopWatch: Set sw = New StopWatch
    
    For Each macro In macros
        
        Set Assert = Application.Run(macro.FullName)
        
        row = row + 1
        
        If IsObject(Assert) Then
            testDataSheet.Range("A" & row).value = macro.macroName
            If Assert.AssertSuccessful Then
                passedCount = passedCount + 1
                testDataSheet.Range("B" & row).value = "passed"
            Else
                failedCount = failedCount + 1
                testDataSheet.Range("B" & row).value = Assert.AssertMessage
            End If
        Else
            testDataSheet.Range("B" & row).value = "assert not defined"
        End If
        
    Next macro
    
    sw.StopCounter
    row = row + 1
    row = row + 1
    testDataSheet.Range("A" & row).value = "Summary:"
    row = row + 1
    testDataSheet.Range("A" & row).value = "Passed"
    testDataSheet.Range("B" & row).value = passedCount
    row = row + 1
    testDataSheet.Range("A" & row).value = "Failed"
    testDataSheet.Range("B" & row).value = failedCount
    row = row + 1
    testDataSheet.Range("A" & row).value = "Duration"
    testDataSheet.Range("B" & row).value = CStr(sw.ElapsedMilliseconds) & " ms"
    
    Exit Sub

Finally:
    MsgBox "Test execution failed " & UserDefinedErrors.FormattedErr & _
           " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
End Sub



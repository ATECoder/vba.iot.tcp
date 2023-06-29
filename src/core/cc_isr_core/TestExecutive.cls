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
''' <para name="testDataSheet">         Specifies the work sheet were test data are set and logged. </param>
''' <para name="usingRemoteWorkbook">   True if using a remote workbook where workbook file needs to be added
''' to the application run call. </param>
Public Sub Execute(ByVal testDataSheet As Worksheet, Optional usingRemoteWorkbook As Boolean = False)
    
    On Error GoTo Finally
    
    If usingRemoteWorkbook Then
        Dim wb As Workbook: Set wb = testDataSheet.Parent
        Dim filePath As String: filePath = wb.FullName
    End If

    Dim moduleFullName As String
   
    Dim messages As New VBA.Collection
    Dim passedCount As Integer
    Dim failedCount As Integer
    Dim row As Integer
    row = 1
    
    testDataSheet.Rows("2:" & testDataSheet.Rows.count).ClearContents
    testDataSheet.Range("B1", "B" & testDataSheet.Rows.count).Interior.ColorIndex = 0
    
    moduleFullName = testDataSheet.Range("B" & row).value
    
    ' get the module info
    Dim module As ModuleInfo
    Set module = cc_isr_Core.Constructor.CreateModuleInfo
    module.FromModuleFullName moduleFullName
    
    row = row + 1
    testDataSheet.Range("A" & row).value = "Project"
    testDataSheet.Range("B" & row).value = module.projectName
    
    row = row + 1
    testDataSheet.Range("A" & row).value = "Module"
    testDataSheet.Range("B" & row).value = module.moduleName
   
    row = row + 1
    Dim passCountRow As Integer: passCountRow = row
    testDataSheet.Range("A" & passCountRow).value = "Passed"
    
    row = row + 1
    Dim failCountRow As Integer: failCountRow = row
    testDataSheet.Range("A" & failCountRow).value = "Failed"
    
    row = row + 1
    Dim durationRow As Integer: durationRow = row
    testDataSheet.Range("A" & durationRow).value = "Duration"
   
    row = row + 1
    testDataSheet.Range("A" & row).value = "Test Name"
    testDataSheet.Range("B" & row).value = "Outcome"
    
    Dim macros As Collection: Set macros = WorkbookUtilities.EnumerateMacros(module)
    Dim macro As MacroInfo
    
    Dim sw As StopWatch: Set sw = cc_isr_Core.Constructor.CreateStopWatch
    
    For Each macro In macros
        
        DoEvents
        
        Dim arg As String: arg = macro.FullName
        If usingRemoteWorkbook Then
            arg = "'" & filePath & "'!" & arg
        End If
        
        Set Assert = Application.Run(arg)
        
        row = row + 1
        
        If IsObject(Assert) Then
            testDataSheet.Range("A" & row).value = macro.macroName
            If Assert.AssertSuccessful Then
                passedCount = passedCount + 1
                testDataSheet.Range("B" & row).value = "passed"
                testDataSheet.Range("B" & row).Interior.ColorIndex = 4
            Else
                failedCount = failedCount + 1
                testDataSheet.Range("B" & row).value = Assert.AssertMessage
                testDataSheet.Range("B" & row).Interior.ColorIndex = 3
            End If
        Else
            testDataSheet.Range("B" & row).value = "assert not defined"
        End If
        
    Next macro
    
    sw.StopCounter
    
    testDataSheet.Range("B" & passCountRow).value = passedCount
    testDataSheet.Range("B" & failCountRow).value = failedCount
    testDataSheet.Range("B" & durationRow).value = CStr(sw.ElapsedMilliseconds) & " ms"
    
    Exit Sub

Finally:
    MsgBox "Test execution failed " & UserDefinedErrors.FormattedErr & _
           " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
End Sub





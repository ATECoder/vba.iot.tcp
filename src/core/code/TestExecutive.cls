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

''' <summary>   Creates a temporary test file. </summary>
''' <returns>   A <see cref="TestFileHandle"/>. <returns>
Public Function CreateTestFile() As TestFileHandle
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    CreateTestFile.TestFilename = fso.BuildPath(ActiveWorkbook.path, fso.GetTempName())
    
    ' Creates a specified file name and returns a TextStream object that can be used to read from or write to the file.
    Set CreateTestFile.TestFileStream = fso.CreateTextFile(CreateTestFile.TestFilename, True)
End Function

''' <summary>   Deletes the given test file. </summary>
''' <param name="fileHandle">   The <see cref="TestFileHandle"/>   to delete. </param>
Public Sub DeleteTestFile(ByRef fileHandle As TestFileHandle)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFile fileHandle.TestFilename
End Sub

''' <summary>   Executes all test macros in the module specified in the test sheet. </summary>
''' <para name="testDataSheet">   Specifies the work sheet were test data are set and logged. </param>
Public Sub Execute(ByRef testDataSheet As Worksheet)
    
    On Error GoTo Finally
    
    Dim moduleName As String
    Dim messages As New VBA.collection
    Dim passedCount As Integer
    Dim failedCount As Integer
    Dim row As Integer
    row = 1
    
    testDataSheet.Rows("2:" & testDataSheet.Rows.count).ClearContents
    moduleName = testDataSheet.Range("B" & row).value
    row = row + 1
    testDataSheet.Range("A" & row).value = "Testing"
    testDataSheet.Range("B" & row).value = moduleName
   
    row = row + 1
    testDataSheet.Range("A" & row).value = "Test Name"
    testDataSheet.Range("B" & row).value = "Outcome"
    
    Dim AppArray() As String
    
    Dim delimiter As String
    delimiter = " "
    AppArray() = Split(WorkbookUtilities.ListAllMacroNames(moduleName, delimiter), delimiter)
    
    Dim sw As StopWatch
    Set sw = New StopWatch
    Dim temp As String
    Dim i As Integer
    For i = LBound(AppArray) To UBound(AppArray)
    
        temp = AppArray(i)
        
        If temp <> "" Then
        
            If temp <> "execute" And temp <> "ListAllMacroNames" And temp <> "AddTestModules" Then
        
                Set Assert = Application.Run(temp)
                
                row = row + 1
                
                If IsObject(Assert) Then
                    testDataSheet.Range("A" & row).value = temp
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
        
           End If
        
        End If
    
    Next i
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



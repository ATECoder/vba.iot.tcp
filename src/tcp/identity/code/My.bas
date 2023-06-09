Attribute VB_Name = "My"
''' <summary> Exports all code files to the active workbook path. </summary>
Public Sub ExportCodeFiles()
    WorkbookUtilities.ExportCodeFiles
End Sub

''' <summary> Execute the tests defined in the TestSheet. </summary>
Public Sub ExecuteTestSheetTests()
    TestExecutive.Execute (TestSheet)
End Sub


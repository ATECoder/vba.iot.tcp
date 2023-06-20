VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' <summary>   Populates a combo box with the list of tests. </summary>
Private Sub ListTestsButton_Click()
    Dim module As ModuleInfo
    testNamesCombo.Clear
    For Each module In WorkbookUtilities.EnumerateProjectModules(Application.ActiveWorkbook.VBProject)
        DoEvents
        testNamesCombo.AddItem module.FullName
    Next module
End Sub

''' <summary>   Runs the selected test. </summary>
Private Sub RunTestsButton_Click()
'   Call cc_isr_Core.StringExtensionsTests.TestDelimitedStringElementShouldPop
    TestExecutive.Execute TestSheet
End Sub

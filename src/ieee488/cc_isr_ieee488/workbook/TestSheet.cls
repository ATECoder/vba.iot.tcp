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
    For Each module In WorkbookUtilities.EnumerateProjectModules(cc_isr_Ieee488.ThisWorkbook.VBProject)
        DoEvents
        testNamesCombo.AddItem module.FullName
    Next module
End Sub

''' <summary>   Runs the selected test. </summary>
Private Sub RunTestsButton_Click()
    cc_isr_Core.TestExecutive.Execute TestSheet, True
End Sub


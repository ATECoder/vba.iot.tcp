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
    Dim moduleName As Variant
    testNamesCombo.Clear
    For Each moduleName In WorkbookUtilities.ListTestModules
        testNamesCombo.AddItem moduleName
    Next moduleName
End Sub

''' <summary>   Runs the selected test. </summary>
Private Sub RunTestsButton_Click()
    My.ExecuteTestSheetTests
End Sub

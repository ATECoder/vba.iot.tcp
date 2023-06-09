VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub ListTestsButton_Click()
    Dim moduleName As Variant
    testNamesCombo.Clear
    For Each moduleName In WorkbookUtilities.ListTestModules
        testNamesCombo.AddItem moduleName
    Next moduleName
End Sub

Private Sub RunTestsButton_Click()
    My.ExecuteTestSheetTests
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub ListTestsButton_Click()
    Dim moduleName As String
    testNamesCombo.Clear
    For Each moduleName In TestRunner.ListTestModules
        testNamesCombo.AddItem moduleName
    Next moduleName
End Sub

Private Sub RunTestsButton_Click()
    TestRunner.Execute
End Sub

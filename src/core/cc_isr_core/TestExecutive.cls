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
''' <para name="a_testSheet">             [Worksheet] Specifies the work sheet were test data are set and logged. </param>
''' <para name="a_usingRemoteWorkbook">   [Optional, Boolean, True] True if using a remote workbook
'''                                       where workbook file needs to be added
''' to the application run call. </param>
Public Sub Execute(ByVal a_testSheet As Worksheet, Optional a_usingRemoteWorkbook As Boolean = False)
    
    On Error GoTo Finally
    
    If a_usingRemoteWorkbook Then
        Dim p_workbook As Workbook: Set p_workbook = a_testSheet.Parent
        Dim p_filePath As String: p_filePath = p_workbook.FullName
    End If

    Dim p_moduleFullName As String
   
    Dim p_passedCount As Integer
    Dim p_failedCount As Integer
    Dim p_rowNo As Integer
    p_rowNo = 1
    
    a_testSheet.Rows("2:" & a_testSheet.Rows.count).ClearContents
    a_testSheet.Range("B1", "B" & a_testSheet.Rows.count).Interior.ColorIndex = 0
    
    p_moduleFullName = a_testSheet.Range("B" & p_rowNo).value
    
    ' get the module info
    Dim p_moduleInfo As ModuleInfo
    Set p_moduleInfo = cc_isr_Core.Constructor.CreateModuleInfo
    p_moduleInfo.FromModuleFullName p_moduleFullName
    
    p_rowNo = p_rowNo + 1
    a_testSheet.Range("A" & p_rowNo).value = "Project"
    a_testSheet.Range("B" & p_rowNo).value = p_moduleInfo.ProjectName
    
    p_rowNo = p_rowNo + 1
    a_testSheet.Range("A" & p_rowNo).value = "module"
    a_testSheet.Range("B" & p_rowNo).value = p_moduleInfo.ModuleName
   
    p_rowNo = p_rowNo + 1
    Dim p_passCountRow As Integer: p_passCountRow = p_rowNo
    a_testSheet.Range("A" & p_passCountRow).value = "Passed"
    
    p_rowNo = p_rowNo + 1
    Dim p_failCountRow As Integer: p_failCountRow = p_rowNo
    a_testSheet.Range("A" & p_failCountRow).value = "Failed"
    
    p_rowNo = p_rowNo + 1
    Dim p_durationRow As Integer: p_durationRow = p_rowNo
    a_testSheet.Range("A" & p_durationRow).value = "Duration"
   
    p_rowNo = p_rowNo + 1
    a_testSheet.Range("A" & p_rowNo).value = "Test Name"
    a_testSheet.Range("B" & p_rowNo).value = "Outcome"
    
    Dim p_macros As Collection: Set p_macros = WorkbookUtilities.EnumerateMacros(p_moduleInfo)
    Dim p_macroInfo As MacroInfo
    
    Dim p_stopper As StopWatch: Set p_stopper = cc_isr_Core.Constructor.CreateStopWatch
    
    For Each p_macroInfo In p_macros
        
        DoEvents
        
        Dim p_macroFullName As String: p_macroFullName = p_macroInfo.FullName
        If a_usingRemoteWorkbook Then
            p_macroFullName = "'" & p_filePath & "'!" & p_macroFullName
        End If
        
        Set Assert = Application.Run(p_macroFullName)
        
        p_rowNo = p_rowNo + 1
        
        If IsObject(Assert) Then
            a_testSheet.Range("A" & p_rowNo).value = p_macroInfo.MacroName
            If Assert.AssertSuccessful Then
                p_passedCount = p_passedCount + 1
                a_testSheet.Range("B" & p_rowNo).value = "passed"
                a_testSheet.Range("B" & p_rowNo).Interior.ColorIndex = 4
            Else
                p_failedCount = p_failedCount + 1
                a_testSheet.Range("B" & p_rowNo).value = Assert.AssertMessage
                a_testSheet.Range("B" & p_rowNo).Interior.ColorIndex = 3
            End If
        Else
            a_testSheet.Range("B" & p_rowNo).value = "assert not defined"
        End If
        
    Next p_macroInfo
    
    p_stopper.StopCounter
    
    a_testSheet.Range("B" & p_passCountRow).value = p_passedCount
    a_testSheet.Range("B" & p_failCountRow).value = p_failedCount
    a_testSheet.Range("B" & p_durationRow).value = CStr(p_stopper.ElapsedMilliseconds) & " ms"
    
    Exit Sub

Finally:
    MsgBox "Test execution failed " & cc_isr_Core.UserDefinedErrors.FormattedErr & _
           " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
End Sub





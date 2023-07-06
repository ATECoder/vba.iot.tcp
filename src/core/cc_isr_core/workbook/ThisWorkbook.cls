VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private m_exportCodeAfterSave As Boolean

''' <summary>   Gets the option for exporting code files after save. </summary>
''' <value> </value>
Public Property Get ExportCodeAfterSave() As Boolean
    ExportCodeAfterSave = m_exportCodeAfterSave
End Property

''' <summary>   Enables or disables exporting code files after save. </summary>
''' <para name="a_value">   [Boolean] True to enable exporting code files after save. </param>
Public Property Let ExportCodeAfterSave(ByVal a_value As Boolean)
    m_exportCodeAfterSave = a_value
End Property

''' <summary>   Handles the Workbook activate event. </summary>
''' <remarks>   Initializes the user defined errors. </remarks>
Private Sub Workbook_Activate()
    
    Const p_procedureName As String = "Workbook_Activate"
    
    On Error Resume Next
    
    ' init the user defined errors.
    cc_isr_Core.UserDefinedErrors.Initialize
    
    If Err.Number <> 0 Then
        
        ' append the error source
        UserDefinedErrors.SetErrSource p_procedureName, Me.name
        
        ' display the error message
        MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
        
    End If
    
    On Error GoTo 0
    
End Sub

''' <summary>   Handles the workbook after save event. </summary>
''' <remarks>   Exports source code files. </remarks>
''' <para name="a_success">   [Boolean] True if saved successfully. </param>
Private Sub Workbook_AfterSave(ByVal a_success As Boolean)
    
    Const p_procedureName As String = "Workbook_AfterSave"
    
    On Error Resume Next
    
    If a_success And Me.ExportCodeAfterSave Then
        cc_isr_Core.WorkbookUtilities.ExportProjectCodeFiles ThisWorkbook.VBProject
    End If
    
    If Err.Number <> 0 Then
        
        ' append the error source
        UserDefinedErrors.SetErrSource p_procedureName, Me.name
        
        ' display the error message
        MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
        
    End If
    
    On Error GoTo 0

End Sub

''' <summary>   Handles the workbook before close event. </summary>
''' <remarks>   Disposes the disposable workseets. </remarks>
''' <para name="a_cancel">   [Boolean] Set to true to cancel closing. </param>
Private Sub Workbook_BeforeClose(ByRef a_cancel As Boolean)
    
    Const p_procedureName As String = "Workbook_BeforeClose"
    
    On Error Resume Next
    
    ' dispose the disposable worksheets
    TimerSheet.Dispose
    
    If Err.Number <> 0 Then
        
        ' append the error source
        UserDefinedErrors.SetErrSource p_procedureName, Me.name
        
        ' display the error message
        MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
        
    End If
    
    On Error GoTo 0

End Sub

''' <summary>   Handles the workbook open event. </summary>
''' <remarks>   Toggles visibility of referenced workbooks. </remarks>
Private Sub Workbook_Open()
    
    Const p_procedureName As String = "Workbook_Open"
    
    On Error Resume Next
    
    ' turn on code exporting for this work book
    ThisWorkbook.ExportCodeAfterSave = True
    
    Application.Windows(ThisWorkbook.name).Visible = True
    
    If Err.Number <> 0 Then
        
        ' append the error source
        UserDefinedErrors.SetErrSource p_procedureName, Me.name
        
        ' display the error message
        MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
        
    End If
    
    On Error GoTo 0

End Sub


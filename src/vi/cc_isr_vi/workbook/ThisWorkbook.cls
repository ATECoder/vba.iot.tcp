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
    
    ' init the user defined errors.
    cc_isr_Core.UserDefinedErrors.Initialize
    
    ' activate the sheets to initialize values.
    K2700Sheet.Initialize cc_isr_Winsock.Constructor.CreateConnectionChangedEventArgs()
    IdentitySheet.Initialize
    TestSheet.Initialize
    
End Sub

''' <summary>   Handles the workbook after save event. </summary>
''' <para name="a_success">   [Boolean] True if save succeeded. </param>
''' <remarks>   Exports source code files. </remarks>
Private Sub Workbook_AfterSave(ByVal a_success As Boolean)
    
    On Error Resume Next
    If a_success And Me.ExportCodeAfterSave Then
        cc_isr_Core.WorkbookUtilities.ExportProjectCodeFiles ThisWorkbook.VBProject
    End If
    On Error GoTo 0

End Sub

''' <summary>   Handles the workbook open event. </summary>
''' <remarks>   Toggles visibility of referenced workbooks. </remarks>
Private Sub Workbook_Open()

    ' hide and turn off code saving on the referenced workbooks
    cc_isr_Core.ThisWorkbook.ExportCodeAfterSave = True
    Application.Windows(Workbooks("cc.isr.core.xlsm").name).Visible = False
    
    cc_isr_Winsock.ThisWorkbook.ExportCodeAfterSave = True
    Application.Windows(Workbooks("cc.isr.winsock.xlsm").name).Visible = False
    
    cc_isr_Ieee488.ThisWorkbook.ExportCodeAfterSave = True
    Application.Windows(Workbooks("cc.isr.ieee488.xlsm").name).Visible = False
    
    ' turn on code exporting for this work book
    ThisWorkbook.ExportCodeAfterSave = True
    Application.Windows(ThisWorkbook.name).Visible = True

End Sub

''' <summary>   Handles the workbook before close event. </summary>
''' <remarks>   Disposes of worksheet objects. </remarks>
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Const p_procedureName As String = "Workbook_BeforeClose"
    On Error Resume Next

    ' disposes of sheet objects.
    K2700Sheet.Dispose

    If Err.Number <> 0 Then
        ' append the error source
        UserDefinedErrors.SetErrSource p_procedureName, Me.name
        
        ' display the error message
        MsgBox UserDefinedErrors.BuildStandardErrorMessage(), vbExclamation
        
    End If

    On Error GoTo 0
End Sub



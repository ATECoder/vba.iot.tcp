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
''' <para name="value">   True to enable exporting code files after save. </param>
Public Property Let ExportCodeAfterSave(ByVal value As Boolean)
    m_exportCodeAfterSave = value
End Property

''' <summary>   Handles the Workbook activate event. </summary>
''' <remarks>   Initializes the user defined errors. </remarks>
Private Sub Workbook_Activate()
    
    On Error Resume Next
    
    ' init the user defined errors.
    cc_isr_Core.UserDefinedErrors.Initialize
    
    On Error GoTo 0

End Sub

''' <summary>   Handles the workbook after save event. </summary>
''' <remarks>   Exports source code files. </remarks>
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error Resume Next
    If Me.ExportCodeAfterSave Then
        cc_isr_Core.WorkbookUtilities.ExportProjectCodeFiles ThisWorkbook.VBProject
    End If
    On Error GoTo 0
End Sub

''' <summary>   Handles the workbook open event. </summary>
''' <remarks>   Toggles visibility of referenced workbooks. </remarks>
Private Sub Workbook_Open()

    ' turn on code exporting for this work book
    ThisWorkbook.ExportCodeAfterSave = True
    Application.Windows(ThisWorkbook.name).Visible = True

End Sub


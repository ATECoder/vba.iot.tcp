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

Private m_exportCodebBeforeClosing As Boolean

''' <summary>   Gets the option for exporting code files when closing. </summary>
Public Property Get ExportCodeBeforeClosing() As Boolean
    ExportCodeBeforeClosing = m_exportCodebBeforeClosing
End Property

''' <summary>   Enables or disables exporting code files when closing. </summary>
''' <para name="value">   True to enable exporting code files when closing. </param>
Public Property Let ExportCodeBeforeClosing(ByVal value As Boolean)
    m_exportCodebBeforeClosing = value
End Property

''' <summary>   Handles the Workbook activate event. </summary>
''' <remarks>   Initializes the user defined errors. </remarks>
Private Sub Workbook_Activate()
    
    ' init the user defined errors.
    cc_isr_Core.UserDefinedErrors.Initialize
    
End Sub

''' <summary>   Handles the workbook after save event. </summary>
''' <remarks>   Exports source code files. </remarks>
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error Resume Next
    If Me.ExportCodeBeforeClosing Then
        cc_isr_Core.WorkbookUtilities.ExportProjectCodeFiles ThisWorkbook.VBProject
    End If
End Sub

''' <summary>   Handles the workbook open event. </summary>
''' <remarks>   Toggles visibility of referenced workbooks. </remarks>
Private Sub Workbook_Open()

    ' hide and turn off code saving on the referenced workbooks
    cc_isr_Core.ThisWorkbook.ExportCodeBeforeClosing = False
    Application.Windows(Workbooks("cc.isr.core.xlsm").name).Visible = False
    
    cc_isr_Winsock.ThisWorkbook.ExportCodeBeforeClosing = False
    Application.Windows(Workbooks("cc.isr.winsock.xlsm").name).Visible = False
    
    ' turn on code exporting for this work book
    ThisWorkbook.ExportCodeBeforeClosing = True
    Application.Windows(ThisWorkbook.name).Visible = True

End Sub


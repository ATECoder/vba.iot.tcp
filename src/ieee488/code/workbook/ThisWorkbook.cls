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

''' <summary>   Handles the Workbook activate event. </summary>
''' <remarks>   Initializes the user defined errors. </remarks>
Private Sub Workbook_Activate()
    Workbooks("cc.isr.core.xlsm").Windows(1).Visible = False
    cc_isr_Core.UserDefinedErrors.Initialize
End Sub

''' <summary>   Handles the workbook before close event. </summary>
''' <remarks>   Exports source code files. </remarks>
Private Sub Workbook_BeforeClose(ByRef cancel As Boolean)
    cc_isr_Core.WorkbookUtilities.ExportCodeFiles
End Sub



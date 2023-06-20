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
    On Error Resume Next
    cc_isr_Core.UserDefinedErrors.Initialize
    Workbooks("cc.isr.core.xlsm").Windows(1).Visible = True
    
    ' turn on exporting when this book is activated.
    Me.ExportCodeBeforeClosing = True
End Sub

''' <summary>   Handles the workbook before close event. </summary>
''' <remarks>   Exports source code files. </remarks>
Private Sub Workbook_BeforeClose(ByRef cancel As Boolean)
    On Error Resume Next
    If Me.ExportCodeBeforeClosing Then
        cc_isr_Core.WorkbookUtilities.ExportCodeFiles
    End If
End Sub



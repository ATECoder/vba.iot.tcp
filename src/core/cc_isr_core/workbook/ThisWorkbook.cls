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

Private Type ThisData
    ExportCodeAfterSave As Boolean
    DeployMode As Boolean
End Type

Private This As ThisData

''' <summary>   Gets the option for exporting code files after save. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ExportCodeAfterSave() As Boolean
    ExportCodeAfterSave = This.ExportCodeAfterSave
End Property

''' <summary>   Enables or disables exporting code files after save. </summary>
''' <para name="value">   True to enable exporting code files after save. </param>
Public Property Let ExportCodeAfterSave(ByVal a_value As Boolean)
    This.ExportCodeAfterSave = a_value
End Property

''' <summary>   Gets the deloy mode. </summary>
''' <remarks>   Code is not saved when in deploy mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get DeployMode() As Boolean
    DeployMode = This.DeployMode
End Property

''' <summary>   Sets the deploy mode. </summary>
''' <para name="value">   True to set the deploy mode. </param>
Public Property Let DeployMode(ByVal a_value As Boolean)
    This.DeployMode = a_value
End Property

''' <summary>   Sets the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [Optional, String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
Public Sub SetErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = vbNullString)

    ' this procedure must not trap errors because it must
    ' not alter the error object.
      
    ' thus we assume that this code is robust and will
    ' not cause errors.
    
    If StringExtensions.IsNullOrEmpty(a_errProjectName) Then
        a_errProjectName = ActiveWorkbook.VBProject.name
    End If
    
    ' get the current source string
    
    Dim p_errorSource As String: p_errorSource = Err.Source
    
    ' build the error source.
    
    p_errorSource = a_errProjectName & "." & a_errModuleName & "." & a_errProcedureName
  
    ' Update the Err.Source
    
    Err.Source = p_errorSource
 
End Sub

''' <summary>   Builds a standard error message. </summary>
''' <param name="a_displayWarning">         [Optional, Boolean, false] True
'''                                         to display a warning rather than
'''                                         an error message. </param>
''' <param name="a_descriptionDelimiter">   [Optional, String, ': '] Specify
'''                                         the delimiter preceding the description. </param>
''' <returns>   A Standard error message string in the form: <para>
''' Error # (0x#) occurred in <c>Source</c>: Description </para><para>
''' or  </para><para>
''' Warning # (0x#) occurred in <c>Source</c>: Description  </para>
''' </returns>
Public Function BuildStandardErrorMessage(Optional ByVal a_displayWarning As Boolean = False, _
                Optional a_descriptionDelimiter As String = ": ") As String

    Dim p_builder As String
  
    ' check if we have an error
    If Err.Number <> 0 Then
    
        p_builder = p_builder & IIf(a_displayWarning, "Warning ", "Error ")
        p_builder = p_builder & Format$(Err.Number)
        
        Dim p_errNumber As Long: p_errNumber = Err.Number - vbObjectError
        
        If Abs(p_errNumber) < &HFFFF& Then
            p_errNumber = p_errNumber - &H200
            p_builder = p_builder & " (+0x"
        Else
            p_errNumber = Err.Number
            p_builder = p_builder & " (0x"
        End If
        p_builder = p_builder & Hex$(p_errNumber)
        p_builder = p_builder & ") "
        p_builder = p_builder & "occurred in "
        p_builder = p_builder & Err.Source
        p_builder = p_builder & a_descriptionDelimiter
        p_builder = p_builder & Err.Description
        
    End If

    BuildStandardErrorMessage = p_builder

End Function

''' <summary>   Handles the Workbook activate event. </summary>
''' <remarks>   Initializes the user defined errors. </remarks>
Private Sub Workbook_Activate()
    
    Const p_procedureName As String = "Workbook_Activate"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    MsgBox Me.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the workbook after save event. </summary>
''' <remarks>   Exports source code files. </remarks>
''' <para name="a_success">   [Boolean] True if saved successfully. </param>
Private Sub Workbook_AfterSave(ByVal a_success As Boolean)
    
    Const p_procedureName As String = "Workbook_AfterSave"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_success And Me.ExportCodeAfterSave Then
        cc_isr_Core.WorkbookUtilities.ExportProjectCodeFiles ThisWorkbook.VBProject
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    MsgBox Me.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the workbook before close event. </summary>
''' <remarks>   Disposes the disposable workseets. </remarks>
''' <para name="a_cancel">   [Boolean] Set to true to cancel closing. </param>
Private Sub Workbook_BeforeClose(ByRef a_cancel As Boolean)
    
    Const p_procedureName As String = "Workbook_BeforeClose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' dispose the disposable worksheets
    TimerSheet.Dispose
    
    Workbooks(ThisWorkbook.name).Saved = Me.DeployMode Or ThisWorkbook.ReadOnly
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    MsgBox Me.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the workbook open event. </summary>
''' <remarks>   Toggles visibility of referenced workbooks. </remarks>
Private Sub Workbook_Open()
    
    Const p_procedureName As String = "Workbook_Open"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' turn on code exporting if not deploy or read only
    ThisWorkbook.ExportCodeAfterSave = Not (Me.DeployMode Or ThisWorkbook.ReadOnly)
    
    Application.Windows(ThisWorkbook.name).Visible = True
    
    ' init the user defined errors.
    cc_isr_Core.UserDefinedErrors.Initialize
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    MsgBox Me.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub


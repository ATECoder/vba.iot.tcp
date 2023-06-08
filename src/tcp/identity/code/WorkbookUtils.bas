Attribute VB_Name = "WorkbookUtils"
''' <summary> Exports all code files to the active workbook path. </summary>
Public Sub ExportCode()
    ExportCodeFiles
End Sub

''' <summary> Exports all code files to the active workbook path. </summary>
''' <para name="subFolder"> Specifies the sub-folder were the files are to be stopred. </param>
' Public Sub ExportCodeFiles(Optional subFolder As String = "code")
Public Sub ExportCodeFiles(Optional subFolder As String = "code")

    Dim component As VBIDE.VBComponent
    Dim extension As String
    Dim fileCount As Integer
    
    ' set and, optionally, create the folder.
    Dim path As String
    path = ActiveWorkbook.path & "\" & subFolder & "\"
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    If Not fso.FolderExists(path) Then
        fso.CreateFolder (path)
    End If
    
    For Each component In ActiveWorkbook.VBProject.VBComponents
        extension = GetFileExtension(component)
        If extension <> "" Then
            fileCount = fileCount + 1
            component.Export Filename:=path & component.name & extension
        End If
    Next component
    
    Debug.Print "Exported " & CStr(fileCount) & " files to " & path
    
End Sub

''' <summary> Get the extension if the component is a file </summary>
''' <param name="component"> The <see cref="VBComponent"/> </param>
''' <returns> The file extension or an empty string if not a file. </returns>
Public Function GetFileExtension(component As VBComponent)
    Select Case component.Type
        Case vbext_ct_ClassModule, vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ""
    End Select
End Function

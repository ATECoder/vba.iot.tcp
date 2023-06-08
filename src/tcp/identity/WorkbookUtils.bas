Attribute VB_Name = "WorkbookUtils"


''' <summary> Exports all code files to the active workbook path. </summary>
Public Sub ExportCodeFiles()
    Dim VBComp As VBIDE.VBComponent
    Dim Sfx As String
    Dim fileCount As Integer
    Dim path As String
    path = ActiveWorkbook.path & "\"
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComp.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select
        If Sfx <> "" Then
            fileCount = fileCount + 1
            VBComp.Export Filename:=path & VBComp.name & Sfx
        End If
    Next VBComp
    Debug.Print "Exported " & CStr(fileCount) & " files to " & path
End Sub



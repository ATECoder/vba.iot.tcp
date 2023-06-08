Attribute VB_Name = "WorkbookUtils"


Public Sub ExportCodeFiles()
    Dim VBComp As VBIDE.VBComponent
    Dim Sfx As String
    
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
            VBComp.Export Filename:=ActiveWorkbook.Path & "\" & VBComp.name & Sfx
        End If
    Next VBComp
End Sub

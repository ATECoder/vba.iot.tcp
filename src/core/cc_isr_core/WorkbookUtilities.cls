VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' <summary>   Exports all code files to the active workbook path. </summary>
''' <para name="subFolder">   [Optional, 'code'] Specifies the sub-folder were the files are to be stored. </param>
Public Sub ExportCodeFiles(Optional ByVal subFolder As String = "code")

    Me.ExportProjectCodeFiles ActiveWorkbook.VBProject, subFolder
    
End Sub

''' <summary>   Exports code files of the specified project to the active workbook path. </summary>
''' <para name="project">          [VBProject, The VB Project which code files to export. </param>
''' <para name="useProjectName">   [Optional, Boolean, True] true to use the project name for the subfolder. </param>
''' <para name="subFolder">        [Optional, String, 'code'] Specifies the sub-folder were the files
'''                                are to be stored. </param>
''' <para name="mustExist">        [Optional, Boolean, True] true if the folder must exists in order to prevent
'''                                unintended storage into the default folder. </param>
Public Sub ExportProjectCodeFiles(ByVal project As VBProject, _
        Optional ByVal useProjectName As Boolean = True, _
        Optional ByVal subFolder As String = "code", _
        Optional ByVal mustExist As Boolean = True)

    Dim p_vbComponent As VBIDE.VBComponent
    Dim p_extension As String
    Dim p_fileCount As Integer
    
    ' set the subfolder tot he project name in lower case
    If useProjectName Then subFolder = LCase$(project.name)
    
    ' set and, optionally, create the folder.
    Dim p_path As String: p_path = PathExtensions.GetParentFolderName(project.FileName)
    
    ' abort if the folder must exist but does not exist.
    If mustExist And Not PathExtensions.FolderExists(PathExtensions.Join(p_path, subFolder, False)) Then
        Exit Sub
    End If
    
    p_path = PathExtensions.Join(p_path, subFolder, True)
    
    Dim p_filePath As String
    Dim p_destPath As String
    
    For Each p_vbComponent In project.VBComponents
        DoEvents
        Dim p_name As String: p_name = p_vbComponent.name
        p_extension = GetFileExtension(p_vbComponent)
        If Not StringExtensions.IsNullOrEmpty(p_extension) Then
            p_destPath = Me.GetWorkbookPath(p_vbComponent, p_path, "workbook")
            p_filePath = PathExtensions.JoinFile(p_destPath, p_name & p_extension)
            p_fileCount = p_fileCount + 1
            p_vbComponent.Export FileName:=p_filePath
        End If
    Next p_vbComponent
    
    Debug.Print "Exported " & CStr(p_fileCount) & " files to " & p_path
    
End Sub

''' <summary>   Gets the destination path for this component; append work book is a workbook component. <summary>
''' <param name="component">   A reference to a VB Component. </param>
''' <param name="path">        The top path. </param>
''' <param name="workbookPath">   The workbook and sheets sub path. </param>
Public Function GetWorkbookPath(ByVal component As VBComponent, ByVal path As String, ByVal subPath As String)
    If Me.IsSheet(component) Or Me.IsWorkbook(component) Or Me.IsChart(component) Or _
        StringExtensions.StartsWith(component.name, "This", vbTextCompare) Then
        GetWorkbookPath = PathExtensions.Join(path, subPath, True)
    Else
        GetWorkbookPath = path
    End If
End Function

''' <summary>   Detects if a component is a workbook. <summary>
''' <param name="component">   A reference to a VB Component. </param>
''' <returns> True if the component is a work book. </returns>
Public Function IsWorkbook(ByVal component As VBComponent)
    On Error Resume Next
    Dim p_property As Property: Set p_property = component.Properties("AutoSaveOn")
    On Error GoTo 0
    IsWorkbook = Not p_property Is Nothing
End Function

''' <summary>   Detects if a component is a sheet. <summary>
''' <param name="component">   A reference to a VB Component. </param>
''' <returns> True if the component is a sheet. </returns>
Public Function IsSheet(ByVal component As VBComponent)
    On Error Resume Next
    Dim p_property As Property: Set p_property = component.Properties("Cells")
    On Error GoTo 0
    IsSheet = Not p_property Is Nothing
End Function

''' <summary>   Detects if a component is a chart. <summary>
''' <param name="component">   A reference to a VB Component. </param>
''' <returns> True if the component is a chart. </returns>
Public Function IsChart(ByVal component As VBComponent)
    On Error Resume Next
    Dim p_property As Property: Set p_property = component.Properties("AutoScaling")
    On Error GoTo 0
    IsChart = Not p_property Is Nothing
End Function

''' <summary>   Get the extension if the component is a file </summary>
''' <param name="component">   The <see cref="VBComponent"/>   </param>
''' <returns>   The file extension or an empty string if not a file. </returns>
Public Function GetFileExtension(ByVal component As VBComponent)
    Select Case component.Type
        Case vbext_ct_ClassModule, vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = vbNullString
    End Select
End Function

''' <summary>   List all modules that end with the specified suffix and
''' contain methods that start with the specified prefix. <summary>
''' <param name="macroNamePrefix">    [Optional, 'Test'] The macro name prefix. </param>
''' <param name="moduleNameSuffix">   [Optional, 'Tests'] The module name suffix </param>
Public Function EnumerateModules(Optional ByVal macroNamePrefix As String = "Test", _
                                 Optional ByVal moduleNameSuffix As String = "Tests") As VBA.Collection

    Dim p_modules As VBA.Collection: Set p_modules = New VBA.Collection
    Dim p_vbProject As VBProject
    For Each p_vbProject In Application.VBE.VBProjects
        DoEvents
        CollectionExtensions.AddItems EnumerateProjectModules(p_vbProject, macroNamePrefix, moduleNameSuffix), _
                                      p_modules
    Next p_vbProject
    Set EnumerateModules = p_modules

End Function

''' <summary>   List all project modules that end with the specified suffix and
''' contain methods that start with the specified prefix. <summary>
''' <param name="project">            The project which modules are to be added. </param>
''' <param name="macroNamePrefix">    [Optional, 'Test'] The macro name prefix. </param>
''' <param name="moduleNameSuffix">   [Optional, 'Tests'] The module name suffix </param>
Public Function EnumerateProjectModules(ByVal project As VBProject, _
                                        Optional ByVal macroNamePrefix As String = "Test", _
                                        Optional ByVal moduleSuffix As String = "Tests") As VBA.Collection
    
    On Error Resume Next
    Dim p_modules As VBA.Collection: Set p_modules = New VBA.Collection
    Dim p_currentModule As ModuleInfo: Set p_currentModule = New ModuleInfo
    Dim p_newModule As ModuleInfo: Set p_newModule = New ModuleInfo
    
    Dim p_vbComponent As VBComponent
    For Each p_vbComponent In project.VBComponents
        DoEvents
        If IsObject(p_vbComponent) Then
            p_newModule.Initialize project.name, p_vbComponent.CodeModule
            If Not p_newModule.Equals(p_currentModule) Then
                Set p_currentModule = p_newModule.Clone
                If StringExtensions.EndsWith(p_currentModule.ModuleName, "Tests") Then
                    If HasMacros(p_vbComponent) Then
                        p_modules.Add p_currentModule
                    End If
                End If
            End If
        End If
    Next
    Set EnumerateProjectModules = p_modules
End Function

''' <summary>   Checks if the Lists all macros in the specified module. <summary>
''' <param name="component">   The component to check for macro methods. </param>
''' <param name="prefix">      [Optional, 'Test'] The prefix of the macro method name. </param>
Function HasMacros(ByVal component As VBComponent, Optional ByVal prefix As String = "Test") As Boolean

    Dim p_currentName As String, p_newName As String
    HasMacros = False
    
    On Error Resume Next
    p_currentName = vbNullString
    
    If Not component Is Nothing Then
        Dim p_lineNo As Integer
        For p_lineNo = 1 To component.CodeModule.CountOfLines
            p_newName = component.CodeModule.ProcOfLine(Line:=p_lineNo, prockind:=vbext_pk_Proc)
            If p_currentName <> p_newName Then
                p_currentName = p_newName
                If StringExtensions.StartsWith(p_currentName, prefix) Then
                   HasMacros = True
                   Exit For
                End If
            End If
        Next
    End If

End Function

''' <summary>   Finds a project. <summary>
''' <param name="projectName">   The projectName. </param>
''' <returns>   A reference to the found project or nothing. </returns>
Public Function FindProjectByName(ByVal ProjectName As String) As VBProject
    
    Dim p_vbProject As VBProject
    For Each p_vbProject In Application.VBE.VBProjects
        DoEvents
        If 0 = StrComp(p_vbProject.name, ProjectName, vbTextCompare) Then
            Set FindProjectByName = p_vbProject
            Exit For
        End If
    Next p_vbProject

End Function

''' <summary>   Finds a code module by name. <summary>
''' <param name="project">      The project. </param>
''' <param name="moduleName">   The component (e.g., module) name. </param>
''' <returns>   A reference to the component of found code module or nothing. </returns>
Public Function FindCodeModuleByName(ByVal project As VBProject, ByVal ModuleName As String) As VBComponent
    
    If Not project Is Nothing Then
        Dim p_vbComponent As VBComponent
        For Each p_vbComponent In project.VBComponents
            DoEvents
            If 0 = StrComp(p_vbComponent.CodeModule, ModuleName, vbTextCompare) Then
                Set FindCodeModuleByName = p_vbComponent
                Exit For
            End If
        Next p_vbComponent
    End If

End Function


''' <summary>   Enumerates all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="module">   The module info of the module containing the macros. </param>
''' <param name="prefix">   [Optional, 'Test'] The prefix of the macro name. </param>
''' <returns>   A Collection of macro names. </returns>
Public Function EnumerateMacros(ByVal module As ModuleInfo, Optional ByVal prefix As String = "Test") As VBA.Collection

    Dim p_macros As Collection: Set p_macros = New Collection
    If Not module Is Nothing Then
    
        Dim p_vbProject As VBProject: Set p_vbProject = FindProjectByName(module.ProjectName)
        
        If Not p_vbProject Is Nothing Then
        
            Dim p_vbComponent As VBComponent: Set p_vbComponent = FindCodeModuleByName(p_vbProject, module.ModuleName)
            
            If Not p_vbComponent Is Nothing Then
            
                Set p_macros = EnumerateComponentMacros(module, p_vbComponent, prefix)
            
            End If
        
        End If
    
    End If
    Set EnumerateMacros = p_macros

End Function

''' <summary>   Enumerates all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="module">   [<see cref="ModuleInfo"/> The module info of the module containing the macros. </param>
''' <param name="prefix">   [Optional, 'Test'] The prefix of the macro name. </param>
''' <returns>   A Collection of macro names. </returns>
Public Function EnumerateComponentMacros(ByVal module As ModuleInfo, _
        ByVal component As VBComponent, Optional ByVal prefix As String = "Test") As VBA.Collection

    Dim p_macros As VBA.Collection: Set p_macros = New VBA.Collection
    Dim p_currentName As String: p_currentName = vbNullString
    Dim p_newName As String
    Dim p_lineNo As Integer
    For p_lineNo = 1 To component.CodeModule.CountOfLines
        p_newName = component.CodeModule.ProcOfLine(Line:=p_lineNo, prockind:=vbext_pk_Proc)
        If p_currentName <> p_newName Then
            p_currentName = p_newName
            If StringExtensions.StartsWith(p_currentName, prefix) Then
                Dim p_macroInfo As MacroInfo: Set p_macroInfo = Constructor.CreateMacroInfo()
                p_macroInfo.InitializeModuelMacro module, p_currentName
                p_macros.Add p_macroInfo
            End If
        End If
    Next
    Set EnumerateComponentMacros = p_macros

End Function

''' <summary>   Removes the specified reference. <summary>
''' <param name="project">         A reference to the VBProject. </param>
''' <param name="referenceName">   The nae of the reference. </param>
''' <returns>   True if the reference was removed. </returns>
Public Function RemoveReference(ByVal project As VBProject, ByVal referenceName As String) As Boolean
    
    RemoveReference = False
    Dim p_reference As Reference
    For Each p_reference In project.References
        If 0 = StrComp(p_reference.name, referenceName, vbTextCompare) Then
            project.References.Remove p_reference
            RemoveReference = True
            Exit For
        End If
    Next p_reference

End Function


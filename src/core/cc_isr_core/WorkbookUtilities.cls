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

    Dim component As VBIDE.VBComponent
    Dim extension As String
    Dim fileCount As Integer
    
    ' set the subfolder tot he project name in lower case
    If useProjectName Then subFolder = LCase$(project.name)
    
    ' set and, optionally, create the folder.
    Dim path As String: path = PathExtensions.GetParentFolderName(project.FileName)
    
    ' abort if the folder must exist but does not exist.
    If mustExist And Not PathExtensions.FolderExists(PathExtensions.Join(path, subFolder, False)) Then
        Exit Sub
    End If
    
    path = PathExtensions.Join(path, subFolder, True)
    
    Dim filePath As String
    Dim destPath As String
    
    For Each component In project.VBComponents
        DoEvents
        Dim name As String: name = component.name
        extension = GetFileExtension(component)
        If Not StringExtensions.IsNullOrEmpty(extension) Then
            destPath = Me.GetWorkbookPath(component, path, "workbook")
            filePath = PathExtensions.JoinFile(destPath, component.name & extension)
            fileCount = fileCount + 1
            component.Export FileName:=filePath
        End If
    Next component
    
    Debug.Print "Exported " & CStr(fileCount) & " files to " & path
    
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
    Dim op As Property: Set op = component.Properties("AutoSaveOn")
    On Error GoTo 0
    IsWorkbook = Not op Is Nothing
End Function

''' <summary>   Detects if a component is a sheet. <summary>
''' <param name="component">   A reference to a VB Component. </param>
''' <returns> True if the component is a sheet. </returns>
Public Function IsSheet(ByVal component As VBComponent)
    On Error Resume Next
    Dim op As Property: Set op = component.Properties("Cells")
    On Error GoTo 0
    IsSheet = Not op Is Nothing
End Function

''' <summary>   Detects if a component is a chart. <summary>
''' <param name="component">   A reference to a VB Component. </param>
''' <returns> True if the component is a chart. </returns>
Public Function IsChart(ByVal component As VBComponent)
    On Error Resume Next
    Dim op As Property: Set op = component.Properties("AutoScaling")
    On Error GoTo 0
    IsChart = Not op Is Nothing
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
            GetFileExtension = ""
    End Select
End Function

''' <summary>   List all modules that end with the specified suffix and
''' contain methods that start with the specified prefix. <summary>
''' <param name="macroNamePrefix">    [Optional, 'Test'] The macro name prefix. </param>
''' <param name="moduleNameSuffix">   [Optional, 'Tests'] The module name suffix </param>
Public Function EnumerateModules(Optional ByVal macroNamePrefix As String = "Test", _
                                 Optional ByVal moduleNameSuffix As String = "Tests") As VBA.Collection

    Dim modules As VBA.Collection: Set modules = New VBA.Collection

    Dim project As VBProject
    For Each project In Application.VBE.VBProjects
        DoEvents
        CollectionExtensions.AddItems EnumerateProjectModules(project, macroNamePrefix, moduleNameSuffix), _
                                      modules
    Next project
    Set EnumerateModules = modules
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
    Dim modules As VBA.Collection: Set modules = New VBA.Collection
    Dim currentModule As ModuleInfo: Set currentModule = New ModuleInfo
    Dim newModule As ModuleInfo: Set newModule = New ModuleInfo
    
    Dim component As VBComponent
    For Each component In project.VBComponents
        DoEvents
        If IsObject(component) Then
            newModule.Initialize project.name, component.CodeModule
            If Not newModule.Equals(currentModule) Then
                Set currentModule = newModule.Clone
                If StringExtensions.EndsWith(currentModule.moduleName, "Tests") Then
                    If HasMacros(component) Then
                        modules.Add currentModule
                    End If
                End If
            End If
        End If
    Next
    Set EnumerateProjectModules = modules
End Function

''' <summary>   Checks if the Lists all macros in the specified module. <summary>
''' <param name="component">   The component to check for macro methods. </param>
''' <param name="prefix">      [Optional, 'Test'] The prefix of the macro method name. </param>
Function HasMacros(ByVal component As VBComponent, Optional ByVal prefix As String = "Test") As Boolean

    Dim currentName As String, newName As String
    HasMacros = False
    
    On Error Resume Next
    currentName = ""
    
    If Not component Is Nothing Then
        Dim i As Integer
        For i = 1 To component.CodeModule.CountOfLines
            newName = component.CodeModule.ProcOfLine(Line:=i, prockind:=vbext_pk_Proc)
            If currentName <> newName Then
                currentName = newName
                If StringExtensions.StartsWith(currentName, prefix) Then
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
Public Function FindProjectByName(ByVal projectName As String) As VBProject
    Dim project As VBProject
    For Each project In Application.VBE.VBProjects
        DoEvents
        If 0 = StrComp(project.name, projectName, vbTextCompare) Then
            Set FindProjectByName = project
            Exit For
        End If
    Next project
End Function

''' <summary>   Finds a code module by name. <summary>
''' <param name="project">      The project. </param>
''' <param name="moduleName">   The component (e.g., module) name. </param>
''' <returns>   A reference to the component of found code module or nothing. </returns>
Public Function FindCodeModuleByName(ByVal project As VBProject, ByVal moduleName As String) As VBComponent
    If Not project Is Nothing Then
        Dim component As VBComponent
        For Each component In project.VBComponents
            DoEvents
            If 0 = StrComp(component.CodeModule, moduleName, vbTextCompare) Then
                Set FindCodeModuleByName = component
                Exit For
            End If
        Next component
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

    Dim macros As Collection: Set macros = New Collection
    If Not module Is Nothing Then
    
        Dim project As VBProject: Set project = FindProjectByName(module.projectName)
        
        If Not project Is Nothing Then
        
            Dim component As VBComponent: Set component = FindCodeModuleByName(project, module.moduleName)
            
            If Not component Is Nothing Then
            
                Set macros = EnumerateComponentMacros(module, component, prefix)
            
            End If
        
        End If
    
    End If
    Set EnumerateMacros = macros

End Function

''' <summary>   Enumerates all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="module">   The module info of the module containing the macros. </param>
''' <param name="prefix">   [Optional, 'Test'] The prefix of the macro name. </param>
''' <returns>   A Collection of macro names. </returns>
Public Function EnumerateComponentMacros(ByVal module As ModuleInfo, _
        ByVal component As VBComponent, Optional ByVal prefix As String = "Test") As VBA.Collection

    Dim macros As VBA.Collection: Set macros = New VBA.Collection
    Dim currentName As String: currentName = ""
    Dim newName As String
    Dim i As Integer
    For i = 1 To component.CodeModule.CountOfLines
        newName = component.CodeModule.ProcOfLine(Line:=i, prockind:=vbext_pk_Proc)
        If currentName <> newName Then
            currentName = newName
            If StringExtensions.StartsWith(currentName, prefix) Then
                Dim macro As MacroInfo: Set macro = Constructor.CreateMacroInfo()
                macro.InitializeModuelMacro module, currentName
                macros.Add macro
            End If
        End If
    Next
    Set EnumerateComponentMacros = macros

End Function

''' <summary>   Removes the specified reference. <summary>
''' <param name="project">         A reference to the VBProject. </param>
''' <param name="referenceName">   The nae of the reference. </param>
''' <returns>   True if the reference was removed. </returns>
Public Function RemoveReference(ByVal project As VBProject, ByVal referenceName As String) As Boolean
    RemoveReference = False
    Dim ref As Reference
    For Each ref In project.References
        If 0 = StrComp(ref.name, referenceName, vbTextCompare) Then
            project.References.Remove ref
            RemoveReference = True
            Exit For
        End If
    Next ref
End Function


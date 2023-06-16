VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <summary>   Exports all code files to the active workbook path. </summary>
''' <para name="subFolder">   [Optional, 'code'] Specifies the sub-folder were the files are to be stored. </param>
Public Sub ExportCodeFiles(Optional ByVal subFolder As String = "code")

    Dim component As VBIDE.VBComponent
    Dim extension As String
    Dim fileCount As Integer
    
    ' set and, optionally, create the folder.
    Dim path As String
    path = PathExtensions.PathJoin(ActiveWorkbook.path, subFolder) & "\"
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    If Not fso.FolderExists(path) Then
        fso.CreateFolder (path)
    End If
    
    Dim filePath As String
    
    For Each component In ActiveWorkbook.VBProject.VBComponents
        Dim name As String: name = component.name
        extension = GetFileExtension(component)
        If Not StringExtensions.IsNullOrEmpty(extension) Then
            filePath = PathExtensions.PathJoin(Me.GetPath(component, path, "workbook"), component.name & extension)
            fileCount = fileCount + 1
            component.Export Filename:=filePath
        End If
    Next component
    
    Debug.Print "Exported " & CStr(fileCount) & " files to " & path
    
End Sub

''' <summary>   Gets the destination path for this component. <summary>
''' <param name="component">   A reference to a VB Component. </param>
''' <param name="path">        The top path. </param>
''' <param name="workbookPath">   The workbook and sheets sub path. </param>
Public Function GetPath(ByVal component As VBComponent, ByVal path As String, ByVal subPath As String)
    If Me.IsSheet(component) Or Me.IsWorkbook(component) Or Me.IsChart(component) Or _
        StringExtensions.StartsWith(component.name, "This", vbTextCompare) Then
        GetPath = PathExtensions.PathJoin(path, subPath)
    Else
        GetPath = path
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

''' <summary>   Detects if a component is a shart. <summary>
''' <param name="component">   A reference to a VB Component. </param>
''' <returns> True if the component is a shart. </returns>
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

''' <summary>   List all test modules that start with the 'Test' prefix and
''' contain methods that start with the test prefix. <summary>
''' <param name="macroNamePrefix">   [Optional, 'Test'] The macro name prefix. </param>
''' <param name="moduleNameSuffix">   [Optional, 'Tests'] The module name suffix </param>
Public Function ListTestModules(Optional ByVal macroNamePrefix As String = "Test", _
                                Optional ByVal moduleNameSuffix As String = "Tests") As VBA.collection

    Dim pj As VBProject
    Dim component As VBComponent
    On Error Resume Next
    Dim moduleNames As VBA.collection
    Set moduleNames = New VBA.collection
    Dim currentModuleName As String
    Dim newModuleName As String
    newModuleName = ""
    currentModuleName = ""
    
    For Each pj In Application.VBE.VBProjects
        For Each component In pj.VBComponents
            If IsObject(component) Then
                newModuleName = component.CodeModule
                If Not StringExtensions.IsNullOrEmpty(newModuleName) And currentModuleName <> newModuleName Then
                    currentModuleName = newModuleName
                    If StringExtensions.EndsWith(currentModuleName, "Tests") Then
                        If HasMacros(component) Then
                            moduleNames.Add currentModuleName
                        End If
                    End If
                End If
            End If
        Next
    Next
    Set ListTestModules = moduleNames
End Function

''' <summary>   List all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="moduleName">   The module name where the macros reside. </param>
''' <param name="delimiter">   [Optional, ' '] The delimiter to use when joining the names. </param>
''' <returns>   Space separated list of macros delimited by <paramref name="delimiter"/>. </returns>
Function ListAllMacroNames(moduleName As String, Optional ByVal delimiter As String = " ") As String

    Dim pj As VBProject
    Dim component As VBComponent
    Dim curMacro As String, newMacro As String
    Dim x As String
    Dim y As String
    Dim macros As String
    
    On Error Resume Next
    curMacro = ""
    
    For Each pj In Application.VBE.VBProjects
        For Each component In pj.VBComponents
            If Not component Is Nothing Then
                If component.CodeModule = moduleName Then
                    Dim i As Integer
                    For i = 1 To component.CodeModule.CountOfLines
                       newMacro = component.CodeModule.ProcOfLine(Line:=i, prockind:=vbext_pk_Proc)

                       If curMacro <> newMacro Then
                          curMacro = newMacro

                            If curMacro <> "" And curMacro <> "app_NewDocument" Then
                                macros = curMacro + delimiter + macros
                            End If

                       End If
                    Next
                End If
            End If
        Next
    Next

    ListAllMacroNames = macros

End Function

''' <summary>   Checks if the Lists all macros in the specified module. <summary>
''' <param name="component">   The component to check for macro methods. </param>
''' <param name="prefix">      [Optional, 'Test'] The prefix of the macro method name. </param>
Function HasMacros(ByRef component As VBComponent, Optional ByVal prefix As String = "Test") As Boolean

    Dim curMacro As String, newMacro As String
    
    On Error Resume Next
    curMacro = ""
    
    If Not component Is Nothing Then
        Dim i As Integer
        For i = 1 To component.CodeModule.CountOfLines
            newMacro = component.CodeModule.ProcOfLine(Line:=i, prockind:=vbext_pk_Proc)
            If curMacro <> newMacro Then
                curMacro = newMacro
                If StringExtensions.StartsWith(curMacro, prefix) Then
                   HasMacros = True
                   Exit Function
                End If
            End If
        Next
    End If
    HasMacros = False

End Function


''' <summary>   Lists all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="moduleName">   The module name where the macros reside. </param>
''' <param name="prefix">       [Optional, 'Test'] The prefix of the macro name. </param>
''' <returns>   A Collection of macro names. </returns>
Function EnumerateMacroNames(moduleName As String, Optional ByVal prefix As String = "Test") As VBA.collection

    Dim pj As VBProject
    Dim component As VBComponent
    Dim curMacro As String, newMacro As String
    Dim macros As VBA.collection
    Set macros = New VBA.collection
    
    On Error Resume Next
    curMacro = ""
    
    For Each pj In Application.VBE.VBProjects
        For Each component In pj.VBComponents
            If Not component Is Nothing Then
                If component.CodeModule = moduleName Then
                    Dim i As Integer
                    For i = 1 To component.CodeModule.CountOfLines
                        newMacro = component.CodeModule.ProcOfLine(Line:=i, prockind:=vbext_pk_Proc)
                        If curMacro <> newMacro Then
                            curMacro = newMacro
                            If StringExtensions.StartsWith(curMacro, prefix) Then
                                macros.Add curMacro
                            End If
                        End If
                    Next
                End If
            End If
        Next
    Next
    Set EnumerateMacroNames = macros

End Function



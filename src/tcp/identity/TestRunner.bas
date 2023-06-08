Attribute VB_Name = "TestRunner"

''' <summary> List all test modules that start with the 'Test' prefix and
''' contain methods that start with the test prefix. <summary>
Function ListTestModules() As Collection

    Dim pj As VBProject
    Dim VBComp As VBComponent
    On Error Resume Next
    Dim moduleNames As Collection
    Set moduleNames = New Collection
    Dim currentModuleName As String
    
    For Each pj In Application.VBE.VBProjects
        For Each VBComp In pj.VBComponents
            If Not VBComp Is Nothing And currentModuleName <> VBComp.CodeModule Then
                currentModuleName = VBComp.CodeModule
                If StringExtensions.StartsWith(moduleName, "Test") And HasMacros(VBComp) Then
                    moduleNames.Add (moduleName)
                End If
            End If
        Next
    Next
    ListTestModules = moduleNames
End Function

''' <summary> List all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="moduleName"> The module name where the macros reside. </param>
''' <param name="delimiter"> optional delimiter </param>
''' <returns> Space separated list of macros delimited by <paramref name="delimiter"/>. </returns>
Function ListAllMacroNames(moduleName As String, Optional ByVal delimiter As String = " ") As String

    Dim pj As VBProject
    Dim VBComp As VBComponent
    Dim curMacro As String, newMacro As String
    Dim x As String
    Dim y As String
    Dim macros As String
    
    On Error Resume Next
    curMacro = ""
    Documents.Add
    
    For Each pj In Application.VBE.VBProjects
        For Each VBComp In pj.VBComponents
            If Not VBComp Is Nothing Then
                If VBComp.CodeModule = moduleName Then
                    For i = 1 To VBComp.CodeModule.CountOfLines
                       newMacro = VBComp.CodeModule.ProcOfLine(Line:=i, prockind:=vbext_pk_Proc)

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

''' <summary> Checks if the Lists all macros in the specified module. <summary>
''' <param name="VBComp"> The component to check for macro methods. </param>
''' <param name="previx"> optional prefix for the macro method name. </param>
Function HasMacros(VBComp As VBComponent, Optional ByVal prefix As String = "Test") As Collection

    Dim curMacro As String, newMacro As String
    
    On Error Resume Next
    curMacro = ""
    
    If Not VBComp Is Nothing Then
        If VBComp.CodeModule = moduleName Then
            For i = 1 To VBComp.CodeModule.CountOfLines
                newMacro = VBComp.CodeModule.ProcOfLine(Line:=i, prockind:=vbext_pk_Proc)
                If curMacro <> newMacro Then
                    curMacro = newMacro
                    If StringExtensions.StartsWith(curMacro, prefix) Then
                       HasMacros = True
                       Exit Function
                    End If
                End If
            Next
        End If
    End If
    HasMacros = False

End Function


''' <summary> Lists all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="moduleName"> The module name where the macros reside. </param>
''' <param name="previx"> optional prefix </param>
''' <returns> A collection of macro names. </returns>
Function EnumerateMacroNames(moduleName As String, Optional ByVal prefix As String = "Test") As Collection

    Dim pj As VBProject
    Dim VBComp As VBComponent
    Dim curMacro As String, newMacro As String
    Dim macros As Collection
    Set macros = New Collection
    
    On Error Resume Next
    curMacro = ""
    
    For Each pj In Application.VBE.VBProjects
        For Each VBComp In pj.VBComponents
            If Not VBComp Is Nothing Then
                If VBComp.CodeModule = moduleName Then
                    For i = 1 To VBComp.CodeModule.CountOfLines
                        newMacro = VBComp.CodeModule.ProcOfLine(Line:=i, prockind:=vbext_pk_Proc)
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
    EnumerateMacroNames = macros

End Function

''' <summary> Execture all test macros in the module specified in the test sheet. </summary>
Public Sub Execute()
    
    Dim moduleName As String
    Dim messages As New Collection
    Dim passedCount As Integer
    Dim failedCount As Integer
    Dim row As Integer
    row = 1
    Set TestSheet = Sheets("TestSheet")
    TestSheet.Rows("2:" & TestSheet.Rows.count).ClearContents
    moduleName = TestSheet.Range("B" & row).Value
    row = row + 1
    TestSheet.Range("A" & row).Value = "Testing"
    TestSheet.Range("B" & row).Value = moduleName
   
    row = row + 1
    TestSheet.Range("A" & row).Value = "Test Name"
    TestSheet.Range("B" & row).Value = "Outcome"
    
    Dim AppArray() As String
    
    Dim delimiter As String
    delimiter = " "
    AppArray() = Split(ListAllMacroNames(moduleName, delimiter), delimiter)
    
    Dim sw As StopWatch
    Set sw = New StopWatch
    For i = 0 To UBound(AppArray)
    
        temp = AppArray(i)
        
        If temp <> "" Then
        
            If temp <> "execute" And temp <> "ListAllMacroNames" Then
        
                Set Assert = Application.Run(temp)
        
                row = row + 1
                TestSheet.Range("A" & row).Value = temp
                If Assert.AssertSuccessful Then
                    passedCount = passedCount + 1
                    TestSheet.Range("B" & row).Value = "passed"
                Else
                    failedCount = failedCount + 1
                    TestSheet.Range("B" & row).Value = Assert.AssertMessage
                End If
            End If
        
        End If
    
    Next i
    sw.StopCounter
    row = row + 1
    row = row + 1
    TestSheet.Range("A" & row).Value = "Summary:"
    row = row + 1
    TestSheet.Range("A" & row).Value = "Passed"
    TestSheet.Range("B" & row).Value = passedCount
    row = row + 1
    TestSheet.Range("A" & row).Value = "Failed"
    TestSheet.Range("B" & row).Value = failedCount
    row = row + 1
    TestSheet.Range("A" & row).Value = "Duration"
    TestSheet.Range("B" & row).Value = CStr(sw.ElapsedMilliseconds) & " ms"

End Sub

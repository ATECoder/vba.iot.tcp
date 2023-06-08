Attribute VB_Name = "TestRunner"

''' <summary> List all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="moduleName"> The module name where the macros reside. </param>
''' <param name="delimiter"> optional delimiter </param>
''' <returns> Space separated list of macros delimited by <paramref name="delimiter"/>. </returns>
Function ListAllMacroNames(moduleName As String, Optional ByVal delimiter As String = " ") As String

    Dim pj As VBProject
    Dim vbcomp As VBComponent
    Dim curMacro As String, newMacro As String
    Dim x As String
    Dim y As String
    Dim macros As String
    
    On Error Resume Next
    curMacro = ""
    Documents.Add
    
    For Each pj In Application.VBE.VBProjects
    
         For Each vbcomp In pj.VBComponents
                If Not vbcomp Is Nothing Then
                    If vbcomp.CodeModule = moduleName Then
                        For i = 1 To vbcomp.CodeModule.CountOfLines
                           newMacro = vbcomp.CodeModule.ProcOfLine(Line:=i, _
                              prockind:=vbext_pk_Proc)
    
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

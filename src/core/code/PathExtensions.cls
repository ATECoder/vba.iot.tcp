VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' <summary> Trims the path from '\' and '/' suffixes. </summary>
''' <param name="path">   The path. </param>
''' <returns> A string. </returns>
Public Function TrimPath(ByVal path As String) As String

    path = Trim(path)
    path = StringExtensions.TrimRight(path, "/")
    path = StringExtensions.TrimRight(path, "\")
    TrimPath = path
    
End Function

''' <summary> Returns true if the folder was deleted and not longer exists. </summary>
''' <param name="path">   The path. </param>
''' <returns> A string. </returns>
Public Function DeleteFolder(ByVal path As String) As Boolean

    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    If fso.FolderExists(path) Then
        fso.DeleteFolder path
    End If
    DeleteFolder = Not fso.FolderExists(path)
    
End Function

''' <summary> Returns true if the folder was deleted and not longer exists. </summary>
''' <param name="path">   The path. </param>
''' <returns> A string. </returns>
Public Function DeleteFile(ByVal path As String) As Boolean

    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    If fso.FileExists(path) Then
        fso.DeleteFile path
    End If
    DeleteFile = Not fso.FileExists(path)
    
End Function

''' <summary> Returns true if the folder exists. </summary>
''' <param name="path">   The path. </param>
''' <returns> A string. </returns>
Public Function FolderExists(ByVal path As String) As Boolean

    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    FolderExists = fso.FolderExists(path)
    
End Function

''' <summary> Returns true if the file exists. </summary>
''' <param name="path">   The path. </param>
''' <returns> A string. </returns>
Public Function FileExists(ByVal path As String) As Boolean

    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    FileExists = fso.FileExists(path)
    
End Function

''' <summary> Returns true if the test file was created. </summary>
''' <param name="path">   The path. </param>
''' <returns> A string. </returns>
Public Function CreateTextFile(ByVal path As String) As Boolean

    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    If Not fso.FileExists(path) Then
        fso.CreateTextFile path
    End If
    CreateTextFile = fso.FileExists(path)
    
End Function

''' <summary> Creates the directory if it does not exist. </summary>
''' <param name="path">   The path. </param>
''' <returns> A string. </returns>
Public Function CreateFolder(ByVal path As String) As String

    If Not StringExtensions.IsNullOrEmpty(path) Then
        path = TrimPath(path)
        Dim fso As FileSystemObject: Set fso = New FileSystemObject
        If Not fso.FolderExists(path) Then
            fso.CreateFolder path
        End If
    End If
    CreateFolder = path
    
End Function

''' <summary> Joins the file to the path; creates the path if it does not exist. </summary>
''' <param name="path">       The path. </param>
''' <param name="fileName">   The file name. </param>
''' <returns> A string. </returns>
Public Function JoinFile(ByVal path As String, ByVal FileName As String, _
        Optional ByVal Create As Boolean = False) As String

    path = TrimPath(path)
    If Create Then path = CreateFolder(path)
    If Not StringExtensions.IsNullOrEmpty(path) Then path = path & "\"
    JoinFile = path & FileName

End Function

''' <summary> Joins the sub folder to the path and creates the path if it does not exist. </summary>
''' <param name="path">        The path. </param>
''' <param name="subFolder">   The sub folder to join to the path. </param>
''' <param name="create">      [optional, False] If true, ereates the folder for the path. </param>
''' <returns> A string. </returns>
Public Function Join(ByVal path As String, ByVal subFolder As String, Optional ByVal Create As Boolean = False) As String

    path = JoinFile(path, subFolder)
    If Create Then
        path = CreateFolder(path)
    End If
    Join = path

End Function

''' <summary> Joins elements to a path creating the paths as needed. </summary>
''' <param name="pathParts">   The path parts to be joined. </param>
''' <returns> A string. </returns>
Public Function JoinAll(ByVal Create As Boolean, ParamArray pathParts() As Variant) As String
    
    Dim joinedPath As String: joinedPath = ""
    If Not IsEmpty(pathParts) Then
        Dim path As Variant
        For Each path In pathParts
            joinedPath = Join(joinedPath, path, Create)
        Next
    End If

    JoinAll = joinedPath
        
End Function

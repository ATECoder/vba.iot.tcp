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
''' <returns>   [String]. </returns>
Public Function TrimPath(ByVal path As String) As String

    path = Trim(path)
    path = StringExtensions.TrimRight(path, "/")
    path = StringExtensions.TrimRight(path, "\")
    TrimPath = path
    
End Function

''' <summary> Returns true if the folder was deleted and not longer exists. </summary>
''' <param name="path">   The path. </param>
''' <returns>   [String]. </returns>
Public Function DeleteFolder(ByVal path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    If p_fileInfo.FolderExists(path) Then
        p_fileInfo.DeleteFolder path
    End If
    DeleteFolder = Not p_fileInfo.FolderExists(path)
    
End Function

''' <summary> Returns true if the folder was deleted and not longer exists. </summary>
''' <param name="path">   The path. </param>
''' <returns>   [String]. </returns>
Public Function DeleteFile(ByVal path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    If p_fileInfo.FileExists(path) Then
        p_fileInfo.DeleteFile path
    End If
    DeleteFile = Not p_fileInfo.FileExists(path)
    
End Function

''' <summary> Returns true if the folder exists. </summary>
''' <param name="path">   The path. </param>
''' <returns>   [String]. </returns>
Public Function FolderExists(ByVal path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    FolderExists = p_fileInfo.FolderExists(path)
    
End Function

''' <summary> Returns true if the file exists. </summary>
''' <param name="path">   The path. </param>
''' <returns>   [String]. </returns>
Public Function FileExists(ByVal path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    FileExists = p_fileInfo.FileExists(path)
    
End Function

''' <summary> Returns true if the test file was created. </summary>
''' <param name="path">   The path. </param>
''' <returns>   [String]. </returns>
Public Function CreateTextFile(ByVal path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    If Not p_fileInfo.FileExists(path) Then
        p_fileInfo.CreateTextFile path
    End If
    CreateTextFile = p_fileInfo.FileExists(path)
    
End Function

''' <summary> Creates the directory if it does not exist. </summary>
''' <param name="path">   The path. </param>
''' <returns>   [String]. </returns>
Public Function CreateFolder(ByVal path As String) As String

    If Not StringExtensions.IsNullOrEmpty(path) Then
        path = TrimPath(path)
        Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
        If Not p_fileInfo.FolderExists(path) Then
            p_fileInfo.CreateFolder path
        End If
    End If
    CreateFolder = path
    
End Function

''' <summary> Joins the file to the path; creates the path if it does not exist. </summary>
''' <param name="path">       [String] The path. </param>
''' <param name="fileName">   [String] The file name. </param>
''' <param name="create">      [optional, Boolean, False] If true, creates the folder for the path. </param>
''' <returns>   [String]. </returns>
Public Function JoinFile(ByVal path As String, ByVal FileName As String, _
        Optional ByVal Create As Boolean = False) As String

    path = TrimPath(path)
    If Create Then path = CreateFolder(path)
    If Not StringExtensions.IsNullOrEmpty(path) Then path = path & "\"
    JoinFile = path & FileName

End Function

''' <summary> Joins the sub folder to the path and creates the path if it does not exist. </summary>
''' <param name="path">        [String] The path. </param>
''' <param name="subFolder">   [String] The sub folder to join to the path. </param>
''' <param name="create">      [optional, Boolean, False] If true, creates the folder for the path. </param>
''' <returns>   [String]. </returns>
Public Function Join(ByVal path As String, ByVal subFolder As String, Optional ByVal Create As Boolean = False) As String

    path = JoinFile(path, subFolder)
    If Create Then
        path = CreateFolder(path)
    End If
    Join = path

End Function

''' <summary> Joins elements to a path creating the paths as needed. </summary>
''' <param name="create">      [Boolean] If true, creates the folder for the path. </param>
''' <param name="pathParts">   [Variant] The path parts to be joined. </param>
''' <returns>   [String]. </returns>
Public Function JoinAll(ByVal Create As Boolean, ParamArray pathParts() As Variant) As String
    
    Dim p_joinedPath As String: p_joinedPath = vbNullString
    If Not IsEmpty(pathParts) Then
        Dim p_pathPart As Variant
        For Each p_pathPart In pathParts
            DoEvents
            p_joinedPath = Join(p_joinedPath, p_pathPart, Create)
        Next
    End If

    JoinAll = p_joinedPath
        
End Function

''' <summary> Gets the parent folder path from the specified path. </summary>
''' <param name="path">   [String] The path. </param>
''' <returns>   [String]. </returns>
Public Function GetParentFolderName(ByVal path As String) As String
    
    GetParentFolderName = path
    If Not StringExtensions.IsNullOrEmpty(path) Then
        path = TrimPath(path)
        Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
        GetParentFolderName = p_fileInfo.GetParentFolderName(path)
    End If
    
End Function

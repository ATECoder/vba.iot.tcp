VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <summary> Joins elements to a path. </summary>
''' <param name="pathParts">   The path parts to be joined. </param>
''' <returns> A string. </returns>
Public Function PathJoin(ParamArray pathParts() As Variant) As String
    
    Dim joinedPath As String
    joinedPath = ""
        If IsEmpty(pathParts) Then
                PathJoin = joinedPath
        Exit Function
    End If

    Dim partIndex As Integer
    Dim pathPart As String

    For partIndex = LBound(pathParts) To UBound(pathParts) - 1
        pathPart = pathParts(partIndex)
        pathPart = StringExtensions.TrimRight(pathPart, "/")
        pathPart = StringExtensions.TrimRight(pathPart, "\")
        
        joinedPath = joinedPath & pathPart & "\"
    Next

    pathPart = pathParts(UBound(pathParts))

    If UBound(pathParts) > 1 Then
        pathPart = StringExtensions.TrimLeft(pathPart, "/")
        pathPart = StringExtensions.TrimLeft(pathPart, "\")
    End If

    joinedPath = joinedPath & pathPart
    PathJoin = joinedPath
        
End Function

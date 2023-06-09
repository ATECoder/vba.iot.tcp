VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''' <summary> Returns true if the collection contains the specified key. </summary>
Public Function ContainsKey(col As Collection, ByVal key As Object) As Boolean
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    Contains = (err.Number = 0)
    err.Clear
End Function

''' <summary> returns true if the object is containd in the collection. </summary>
Public Function ContainsAll(col As Collection, contained As Collection) As Boolean
    
    Dim result As Boolean
    Dim key As Object
    For Each key In contained
        If Not ContainsKey(col, key) Then
            ContainsAll = False
            Exit Function
        End If
    Next key
    ContainsAll = True
    
End Function




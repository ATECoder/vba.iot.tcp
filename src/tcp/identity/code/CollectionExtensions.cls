VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''' <summary> returns true if the object is contained in the collection. </summary>
Public Function Contains(col As Collection, key As Variant) As Boolean
    On Error GoTo ErrorHandler
    Contains = True
    IsObject (col.Item(key))
    Exit Function
ErrorHandler:
    Contains = False
End Function

''' <summary> returns true if the object is containd in the collection. </summary>
Public Function ContainsAll(col As Collection, contained As Collection) As Boolean
    
    Dim result As Boolean
    Dim key As Variant
    For Each key In contained
        If Not Contains(col, key) Then
            ContainsAll = False
            Exit Function
        End If
    Next key
    ContainsAll = True
    
End Function




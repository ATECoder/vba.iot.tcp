VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''' <summary>   Returns true if the collection contains the specified key. </summary>
''' <param name="collection">   The subject collection. </param>
''' <param name="key">          The key to check for in the collection. </param>
''' <returns>   </returns>
Public Function ContainsKey(ByVal col As VBA.collection, ByVal key As Variant) As Boolean
    Dim found As Boolean
    found = False
    For Each colItem In col
        If colItem = key Then
            found = True
            Exit For
        End If
    Next colItem
#If False Then
    ' this causes error 5 event if doing this:
    Print col.Item(1)
    StringExtensionsTests
    Print col.Item("StringExtensionsTests")
    ' or even this!
    Print col.Item(col.Item(1))
    On Error Resume Next
    col key ' Just try it. If it fails, Err.Number will be nonzero.
    ContainsKey = (Err.Number = 0)
#End If
    ContainsKey = found
End Function

''' <summary>   returns true if the object is containd in the collection. </summary>
''' <param name="collection">   The subject collection. </param>
''' <param name="contained">   The collection which to check for being contained in the
'''                           subject collection. </param>
''' <returns>   </returns>
Public Function ContainsAll(ByVal col As VBA.collection, ByVal contained As VBA.collection) As Boolean
    
    Dim result As Boolean
    Dim key As Variant
    For Each key In contained
        If Not ContainsKey(col, key) Then
            ContainsAll = False
            Exit Function
        End If
    Next key
    ContainsAll = True
    
End Function




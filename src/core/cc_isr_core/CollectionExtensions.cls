VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

''' <summary>   Returns true if the collection contains the specified key. </summary>
''' <param name="collection">   The subject collection. </param>
''' <param name="key">          The key to check for in the collection. </param>
''' <returns>   True if the key is contained in the collection. </returns>
Public Function ContainsKey(ByVal col As VBA.Collection, ByVal key As Variant) As Boolean
    
    Dim p_found As Boolean
    p_found = False
    Dim colItem As Variant
    For Each colItem In col
        DoEvents
        If colItem = key Then
            p_found = True
            Exit For
        End If
    Next colItem
    ContainsKey = p_found

#If False Then
    ' this causes error 5 event if doing this:
    Print col.item(1)
    StringExtensionsTests
    Print col.item("StringExtensionsTests")
    ' or even this!
    Print col.item(col.item(1))
    On Error Resume Next
    col key ' Just try it. If it fails, Err.Number will be nonzero.
    ContainsKey = (Err.Number = 0)
#End If

End Function

''' <summary>   Returns true if the object is contained in the collection. </summary>
''' <param name="col">         The subject collection. </param>
''' <param name="contained">   The collection which to check for being contained in the
'''                            subject collection. </param>
''' <returns>   True if the contained collection is fully contained in the collection. </returns>
Public Function ContainsAll(ByVal col As VBA.Collection, ByVal contained As VBA.Collection) As Boolean
    
    Dim p_result As Boolean: p_result = True
    Dim p_key As Variant
    For Each p_key In contained
        DoEvents
        If Not ContainsKey(col, p_key) Then
            p_result = False
            Exit For
        End If
    Next p_key
    ContainsAll = p_result
    
End Function

''' <summary>   Returns the first item that exists in <paramref name="contained"/>
''' not existing in <paramref name="col"/>. </summary>
''' <param name="col">         The subject collection. </param>
''' <param name="contained">   The collection which to check for being contained in the
'''                            subject collection. </param>
Public Function FindMissingItem(ByVal col As VBA.Collection, ByVal contained As VBA.Collection) As Variant
    
    Dim p_result As Variant: Set p_result = Nothing
    Dim p_key As Variant
    For Each p_key In contained
        DoEvents
        If Not ContainsKey(col, p_key) Then
            p_result = p_key
            Exit For
        End If
    Next p_key
    Set FindMissingItem = p_result
    
End Function

''' <summary>   Adds items from the source to the destination collection. </summary>
''' <param name="source">        The source collection. </param>
''' <param name="destination">   The destination collection which receives the added items. </param>
''' <returns>   The number of items that were added. </returns>
Public Function AddItems(ByVal source As VBA.Collection, ByVal destination As VBA.Collection) As Integer
    
    Dim p_count As Integer: p_count = 0
    Dim p_item As Variant
    For Each p_item In source
        DoEvents
        destination.Add p_item
        p_count = p_count + 1
    Next p_item
    AddRange = p_count
    
End Function




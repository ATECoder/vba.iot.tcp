Attribute VB_Name = "CollectionExtensionsTests"
Option Explicit

''' <summary>   Unit test. Asserts that the collection contains an expected value. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldContain() As Assert
    Dim col As collection
    Set col = New collection
    Dim expected As Variant: expected = "a"
    col.Add expected
    Set TestCollectionShouldContain = Assert.IsTrue(CollectionExtensions.ContainsKey(col, expected), "The collection should contain the value")
End Function

''' <summary>   Unit test. Asserts that the collection does not contain a value. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldNotContain() As Assert
    Dim col As collection
    Set col = New collection
    Dim expected As Variant: expected = "a"
    Dim notExpected As Variant: notExpected = "b"
    col.Add expected
    Set TestCollectionShouldNotContain = Assert.IsFalse(CollectionExtensions.ContainsKey(col, notExpected), "The collection should not contain a value")
End Function

''' <summary>   Unit test. Asserts that the collection contains an expected value. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldContainItself() As Assert
    Dim col As collection
    Set col = New collection
    col.Add "a"
    col.Add "b"
    Set TestCollectionShouldContainItself = Assert.IsTrue(CollectionExtensions.ContainsAll(col, col), "The collection should contain itself")
End Function


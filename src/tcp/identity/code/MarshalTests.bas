Attribute VB_Name = "TestMarshal"
Option Explicit

''' <summary> Tests converting an int8 to a big endian byte string 
''' and back from a big endian byte string to an int8. </summary>
Public Function TestShouldMarshalInt8() As Assert
    Dim value As Byte
    value = 10
   
    Set TestMarshalInt8 = Assert.AreEqual(value, Marshal.BytesToInt8(Marshal.Int8ToBytes(value)), "marshals int8")
End Function

''' <summary> Tests converting an int16 to a big endian byte string 
''' and back from a big endian byte string to an int16. </summary>
Public Function TestShouldMarshalInt16() As Assert
    Dim value As Long
    value = 10
    
    Set TestMarshalInt16 = Assert.AreEqual(value, Marshal.BytesToInt16(Marshal.Int16ToBytes(value)), "marshals int16")
End Function

''' <summary> Tests converting an int32 to a big endian byte string 
''' and back from a big endian byte string to an int32. </summary>
Public Function TestShouldMarshalInt32() As Assert
    Dim value As Long
    value = 10
    
    Set TestMarshalInt32 = Assert.AreEqual(value, Marshal.BytesToInt32(Marshal.Int32ToBytes(value)), "marshals int32")
End Function

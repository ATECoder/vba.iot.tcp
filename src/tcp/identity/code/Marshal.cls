VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Marshal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const integer0xFF As Integer = 256
Private Const long0xFF As Long = 256

''' <summary> Converts the given int8 to a big endian byte string. </summary>
''' <param name="value"> The byte to convert to bytes. </param>
''' <returns> A string. </returns>
Public Function Int8ToBytes(ByVal value As Byte) As String
    Int8ToBytes = Chr$(value)
End Function

''' <summary> Converts the given big endian byte string to an int8. </summary>
''' <param name="bytes"> The bytes to convert. </param>
''' <returns> A byte. </returns>
Public Function BytesToInt8(ByVal bytes As String) As Byte
    BytesToInt8 = Asc(StringExtensions.CharAt(bytes, 1))
End Function

''' <summary> Converts the given int16 to a big endian byte string. </summary>
''' <param name="value"> The byte to convert. </param>
''' <returns> </returns>
Public Function Int16ToBytes(ByVal value As Integer) As String
    Dim bytes As String * 2
    Dim rest As Integer

    rest = value Mod integer0xFF
    Mid(bytes, 2) = Chr$(rest)

    value = (value - rest) / integer0xFF
    rest = value Mod integer0xFF
    Mid(bytes, 1) = Chr$(rest)

    Int16ToBytes = bytes
End Function

''' <summary> Converts the given big endian byte string to an int16. </summary>
''' <param name="bytes"> The bytes to convert. </param>
''' <returns> A Long. </returns>
Public Function BytesToInt16(ByVal bytes As String) As Long
    BytesToInt16 = Asc(StringExtensions.CharAt(bytes, 1)) * long0xFF + _
                   Asc(StringExtensions.CharAt(bytes, 2))
End Function

''' <summary> Converts the given int32 to a big endian byte string. </summary>
''' <param name="value"> The byte to convert. </param>
''' <returns> </returns>
Public Function Int32ToBytes(ByVal value As Long) As String
    Dim bytes As String * 4
    Dim rest As Long

    rest = value Mod long0xFF
    Mid(bytes, 4) = Chr$(rest)

    value = (value - rest) / long0xFF
    rest = value Mod long0xFF
    Mid(bytes, 3) = Chr$(rest)

    value = (value - rest) / long0xFF
    rest = value Mod long0xFF
    Mid(bytes, 2) = Chr$(rest)

    value = (value - rest) / long0xFF
    rest = value Mod long0xFF
    Mid(bytes, 1) = Chr$(rest)

    Int32ToBytes = bytes
End Function

''' <summary> Converts the given big endian byte string to an int32. </summary>
''' <param name="bytes"> The bytes to convert. </param>
''' <returns> A long. </returns>
Public Function BytesToInt32(ByVal bytes As String) As Long
    BytesToInt32 = Asc(StringExtensions.CharAt(bytes, 1)) * long0xFF * long0xFF * long0xFF + _
                   Asc(StringExtensions.CharAt(bytes, 2)) * long0xFF * long0xFF + _
                   Asc(StringExtensions.CharAt(bytes, 3)) * long0xFF + _
                   Asc(StringExtensions.CharAt(bytes, 4))
End Function

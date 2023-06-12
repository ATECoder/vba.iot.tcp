VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <remarks> Trims the given character from the given text starting from the left. </remarks>
''' <param name="text"> The text to trim. </param>
''' <param name="c">    The character to trim. </param>
''' <returns> A string. <returns>
Public Function TrimLeft(ByVal text As String, c As String) As String
    Dim textLength As Long
    textLength = Len(text)
    
    Dim firstCharacter As String
    
    Do While textLength > 0
        firstCharacter = Left(text, 1)
        
        If firstCharacter <> c Then
            Exit Do
        End If
        
        text = Right(text, textLength - 1)
        textLength = Len(text)
    Loop
    
    TrimLeft = text
End Function

''' <summary> Trims the given character from the given text starting from the right. </summary>
''' <remarks> </remarks>
''' <param name="text"> The text to trim. </param>
''' <param name="c">    The character to trim. </param>
''' <returns> A string. <returns>
Public Function TrimRight(ByVal text As String, c As String) As String
    Dim textLength As Long
    textLength = Len(text)
    
    Dim lastCharacter As String
        
    Do While textLength > 0
        lastCharacter = Right(text, 1)
        
        If lastCharacter <> c Then
            Exit Do
        End If
        
        text = Left(text, textLength - 1)
        textLength = Len(text)
    Loop
    
    TrimRight = text
End Function

''' <summary> Gets the substring from the given text. </summary>
''' <remarks> </remarks>
''' <param name="text">       The text to get the substring from. </param>
''' <param name="startIndex"> The index of the first character of the substring. </param>
''' <param name="length">     The number of characters to take from the original string. </param>
''' <returns> A string. <returns>
Public Function Substring(ByVal text As String, ByVal startIndex As Integer, Optional ByVal length As Variant) As String
    If startIndex > Len(text) Then
        startIndex = Len(text)
    End If

    If IsMissing(length) Then
        length = Len(text) - startIndex
    End If
    
    If length > Len(text) Then
        length = Len(text) - startIndex
    End If

    Substring = Left(Right(text, Len(text) - startIndex), length)
End Function

''' <summary> Checks whether the given text starts with the given sequence. </summary>
''' <remarks> </remarks>
''' <param name="text">      The text to check for the sequence. </param>
''' <param name="startText"> The text to be located at the start. </param>
''' <returns> A Boolean <returns>
Public Function StartsWith(ByVal text As String, ByVal startText As String) As Boolean
    StartsWith = InStr(text, startText) = 1
End Function

''' <summary> Checks whether the given text ends with the given sequence. </summary>
''' <remarks> </remarks>
''' <param name="text">    The text to check for the sequence. </param>
''' <param name="endText"> The text to be located at the end. </param>
''' <returns> A Boolean. <returns>
Public Function EndsWith(ByVal text As String, ByVal endText As String) As Boolean
    EndsWith = Right(text, Len(endText)) = endText
End Function

''' <summary> Gets the character at the given index from the given string. </summary>
''' <remarks> </remarks>
''' <param name="text">  The text to get the character from. </param>
''' <param name="index"> The index of the character to get. </param>
''' <returns> A string. <returns>
Public Function CharAt(ByVal text As String, ByVal index As Integer) As String
    CharAt = Mid(text, index, 1)
End Function

''' <summary> Repeats the given string the given amount of times. </summary>
''' <remarks> </remarks>
''' <param name="text">  The text to repeat. </param>
''' <param name="count"> The amount of times to repeat the given string. </param>
''' <returns> A string. <returns>
Public Function Repeat(ByVal text As String, ByVal count As Long) As String
    Repeat = ""
    
    Dim i As Long
    For i = 1 To count
        Repeat = Repeat & text
    Next
End Function

''' <summary> Converts a regular string to a unicode string,
''' e.g., "foo" to a L"foo" string.. </summary>
''' <param name="text"> The string to convert. </param>
''' <returns> A string. <returns>
Public Function StringToWideString(ByVal text As String) As String
    StringToWideString = StrConv(text, vbUnicode)
End Function

''' <summary> Returns true if the value is null or empty. </summary>
''' <returns> true if the value is null or empty. </returns>
Public Function IsNullOrEmpty(ByVal text As String) As Boolean
    IsNullOrEmpty = (text = "")
End Function



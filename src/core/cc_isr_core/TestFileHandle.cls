VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestFileHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_fileName As String
Private m_fileStream As TextStream

''' <summary>   Creates a temporary test file. </summary>
''' <returns>   A <see cref="TestFileHandle"/>. <returns>
Public Function CreateTestFile() As TestFileHandle
    
    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    ' Creates a specified file name and returns a TextStream object that can be used to read from
    ' or write to the file.
    Dim p_handle As TestFileHandle: Set p_handle = New TestFileHandle
    p_handle.FileName = p_fileInfo.BuildPath(ActiveWorkbook.path, p_fileInfo.GetTempName())
    Set p_handle.FileStream = p_fileInfo.CreateTextFile(p_handle.FileName, True)
    Set CreateTestFile = p_handle
End Function

''' <summary>   Deletes the given test file. </summary>
Public Sub DeleteTestFile()
    
    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    p_fileInfo.DeleteFile Me.FileName

End Sub

''' <summary>   Deletes the file handle test file. </summary>
''' <param name="fileHandle">   The <see cref="TestFileHandle"/>   to delete. </param>
Public Sub Delete(ByVal fileHandle As TestFileHandle)
    
    Dim p_fileInfo
    Set p_fileInfo = CreateObject("Scripting.FileSystemObject")
    p_fileInfo.DeleteFile fileHandle.FileName

End Sub

''' <summary> </summary>
''' <value>   [String]. </value>
Public Property Get FileName() As String
    FileName = m_fileName
End Property

''' <summary> </summary>
Public Property Let FileName(ByVal value As String)
    m_fileName = value
End Property

''' <summary> </summary>
''' <value>   [TextStream]. </value>
Public Property Get FileStream() As TextStream
    Set FileStream = m_fileStream
End Property

''' <summary>   Sets the file stream. </summary>
''' <param name="value">   [TextStream]. </param>
Public Property Set FileStream(ByVal value As TextStream)
    Set m_fileStream = value
End Property


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
    
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    ' Creates a specified file name and returns a TextStream object that can be used to read from
    ' or write to the file.
    Dim handle As TestFileHandle: Set handle = New TestFileHandle
    handle.FileName = fso.BuildPath(ActiveWorkbook.path, fso.GetTempName())
    Set handle.FileStream = fso.CreateTextFile(handle.FileName, True)
    Set CreateTestFile = handle
End Function

''' <summary>   Deletes the given test file. </summary>
Public Sub DeleteTestFile()
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    fso.DeleteFile Me.FileName
End Sub

''' <summary>   Deletes the file handle test file. </summary>
''' <param name="fileHandle">   The <see cref="TestFileHandle"/>   to delete. </param>
Public Sub Delete(ByVal fileHandle As TestFileHandle)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile fileHandle.FileName
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


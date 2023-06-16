VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_code As Long
Private m_name As String
Private m_description As String

''' <summary>   Initializes a user defined error. </summary>
''' <param name="errorCode">   The error code. </param>
''' <param name="errorName">   The error name. </param>
''' <param name="errorDescription">   The error description. </param>
Public Sub InitiateProperties(ByVal errorCode As Long, ByVal errorName As String, ByVal errorDescription As String)
    m_name = errorName
    m_code = errorCode
    m_description = errorDescription
End Sub

''' <summary>   Gets the error code. </summary>
''' <value>   An error code of Long type. </value>
Public Property Get Code() As Long
    Code = m_code
End Property

''' <summary>   Gets the error number base. Defaults to 512. </summary>
''' <value>   The error number base value of Long type. </value>
Public Property Get UserErrorNumberBase() As Long
    UserErrorNumberBase = 512
End Property

''' <summary>   Gets the user defined error number after removing the vbObjectError
''' and error number base. </summary>
''' <value>   The user error number of Long type. </value>
Public Property Get UserErrorNumber() As Long
    UserErrorNumber = Me.Code - (VBA.vbObjectError + Me.UserErrorNumberBase)
End Property

''' <summary>   Gets the error name. </summary>
''' <value>   The error name of String type. </value>
Public Property Get name() As String
    name = m_name
End Property

''' <summary>   The error description. </summary>
''' <value>   The error description of String type. </value>
Public Property Get Description() As String
    Description = m_description
End Property

''' <summary>   Returns a string representation of the error. </summary>
''' <param name="additionaInfo">   [Optional, empty] additional information to append to the error message. </param>
''' <returns>   A String representation of the error. </returns>
Public Function ToString(Optional ByVal additionalInfo As String = "") As String
    ToString = "User Error " & Me.name & " (0x" & VBA.LCase$(VBA.Hex$(Me.Code)) & " 0x" & VBA.LCase$(VBA.Hex$(Me.UserErrorNumber)) & "): " & _
                Me.Description & additionalInfo
End Function


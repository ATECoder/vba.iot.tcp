VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_code As Long
Private m_name As String
Private m_description As String

Public Sub InitiateProperties(ByVal errorCode As Long, ByVal errorName As String, ByVal errorDescription As String)
    m_name = errorName
    m_code = errorCode
    m_description = errorDescription
End Sub

Public Property Get Code() As Long
    Code = m_code
End Property

Public Property Get UserErrorNumberBase() As Long
    UserErrorNumberBase = 512
End Property

Public Property Get UserErrorNumber() As Long
    UserErrorNumber = Me.Code - (vbObjectError + Me.UserErrorNumberBase)
End Property

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Get Description() As String
    Description = m_description
End Property

Public Function ToString(Optional ByVal additionaInfo As String = "") As String
    ToString = "User Error " & Me.Name & " (0x" & Hex(Me.Code) & " 0x" & Hex(Me.UserErrorNumber) & "): " & _
                Me.Description & additionalInfo
End Function

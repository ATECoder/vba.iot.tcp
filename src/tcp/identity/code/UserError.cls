VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_code As Long
Private m_description As String

Public Sub InitiateProperties(ByVal errorCode As Long, ByVal errorDescription As String)
    m_code = errorCode
    m_description = errorDescription
End Sub

Public Property Get Code() As Long
    Code = m_code
End Property

Public Property Get Description() As String
    Description = m_description
End Property

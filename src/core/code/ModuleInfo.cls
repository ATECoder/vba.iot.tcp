VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModuleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_projectName As String
Private m_moduleName As String

Public Sub Initialize(ByVal project As VBProject, ByVal component As VBComponent)
    m_projectName = project.Name
    m_moduleName = component.CodeModule
End Sub

Public Property Get ProjectName() As String
    ProjectName = m_projetName
End Property

Public Property Get ModuleName() As String
    ModuleName = m_moduleName
End Property

Public Property Get FullName() As String
    FullName = m_projetName & "." & m_moduleName & "." & m_name
End Property





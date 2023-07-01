VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModuleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_projectName As String
Private m_moduleName As String

''' <summary>   Initializes this class. </summary>
''' <param name="project">     The project name. </param>
''' <param name="component">   The component code module name. </param>
Public Sub Initialize(ByVal ProjectName As String, ByVal codeModuleName As String)
    m_projectName = ProjectName
    m_moduleName = codeModuleName
End Sub

''' <summary>   Initializes this class from a module full name. </summary>
''' <param name="moduleFullName">     the module full name. </param>
Public Sub FromModuleFullName(ByVal moduleFullName As String)
    
    Dim p_names() As String: p_names() = Split(moduleFullName, ".")
    m_projectName = p_names(LBound(p_names))
    m_moduleName = p_names(LBound(p_names) + 1)

End Sub

''' <summary>   Initializes this class from a macro full name. </summary>
''' <param name="macroFullName">     the macro full name. </param>
Public Sub FromMacroFullName(ByVal macroFullName As String)
    
    Dim p_names() As String: p_names() = Split(macroFullName, ".")
    m_projectName = p_names(LBound(p_names))
    m_moduleName = p_names(LBound(p_names) + 1)

End Sub

''' <summary>   Creates a clone of the this module info. </summary>
Public Function Clone() As ModuleInfo
    Set Clone = New ModuleInfo
    Clone.Initialize Me.ProjectName, Me.ModuleName
End Function

''' <summary>   Gets the project name. </summary>
Public Property Get ProjectName() As String
    ProjectName = m_projectName
End Property

''' <summary>   Gets the module name. </summary>
Public Property Get ModuleName() As String
    ModuleName = m_moduleName
End Property

''' <summary>   Gets the full name. </summary>
Public Property Get FullName() As String
Attribute FullName.VB_Description = "This is a description of a member variable"
Attribute FullName.VB_UserMemId = 0
    FullName = m_projectName & "." & m_moduleName
End Property

''' <summary>   Checks if the two modules are equal referring to the same
''  component in the project. </summary>
''' <param name="module">   The module to check for equality. </param>
Public Function Equals(ByVal module As ModuleInfo) As Boolean
    If IsObject(module) Then
        Equals = (0 = StrComp(Me.ProjectName, module.ProjectName)) And _
                 (0 = StrComp(Me.ModuleName, module.ModuleName))
    Else
        Equals = False
    End If
End Function



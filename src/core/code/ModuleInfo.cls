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
Public Sub Initialize(ByVal projectName As String, ByVal codeModuleName As String)
    m_projectName = projectName
    m_moduleName = codeModuleName
End Sub

''' <summary>   Initializes this class from a module full name. </summary>
''' <param name="moduleFullName">     the module full name. </param>
Public Sub FromModuleFullName(ByVal moduleFullName As String)
    Dim names() As String: names() = Split(moduleFullName, ".")
    m_projectName = names(LBound(names))
    m_moduleName = names(LBound(names) + 1)
End Sub

''' <summary>   Initializes this class from a macro full name. </summary>
''' <param name="macroFullName">     the macro full name. </param>
Public Sub FromMacroFullName(ByVal macroFullName As String)
    Dim names() As String: names() = Split(macroFullName, ".")
    m_projectName = names(LBound(names))
    m_moduleName = names(LBound(names) + 1)
End Sub

''' <summary>   Creates a clone of the this module info. </summary>
Public Function Clone() As ModuleInfo
    Dim newModule As ModuleInfo: Set newModule = New ModuleInfo
    newModule.Initialize Me.projectName, Me.moduleName
End Function

''' <summary>   Gets the project name. </summary>
Public Property Get projectName() As String
    projectName = m_projectName
End Property

''' <summary>   Gets the module name. </summary>
Public Property Get moduleName() As String
    moduleName = m_moduleName
End Property

''' <summary>   Gets the full name. </summary>
Public Property Get FullName() As String
    FullName = m_projectName & "." & m_moduleName
End Property

''' <summary>   Checks if the two modules are equal refering to the same
''  component in the project. </summary>
''' <param name="module">   The module to check for equality. </param>
Public Function Equals(ByRef module As ModuleInfo) As Boolean
    If IsObject(module) Then
        Equals = StrComp(Me.projectName, module.projectName) And _
                 StrComp(Me.moduleName, module.moduleName)
    Else
        Equals = False
    End If
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacroInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_projectName As String
Private m_moduleName As String
Private m_macroName As String

''' <summary>   Initializes this class. </summary>
''' <param name="project">     The project name. </param>
''' <param name="component">   The component code module name. </param>
''' <param name="macroName">   The name of the macro. </param>
Public Sub Initialize(ByVal ProjectName As String, ByVal codeModuleName As String, ByVal MacroName As String)
    m_projectName = ProjectName
    m_moduleName = codeModuleName
    m_macroName = MacroName
End Sub

''' <summary>   Initializes this class. </summary>
''' <param name="project">     A reference to the module info. </param>
''' <param name="macroName">   The name of the macro. </param>
Public Sub InitializeModuelMacro(ByVal module As ModuleInfo, ByVal MacroName As String)
    m_projectName = module.ProjectName
    m_moduleName = module.ModuleName
    m_macroName = MacroName
End Sub

''' <summary>   Initializes this class from a macro full name. </summary>
''' <param name="macroFullName">     the macro full name. </param>
Public Sub FromMacroFullName(ByVal macroFullName As String)
    
    Dim p_names() As String: p_names() = Split(macroFullName, ".")
    m_projectName = p_names(LBound(p_names))
    m_moduleName = p_names(LBound(p_names) + 1)
    m_macroName = p_names(LBound(p_names) + 2)

End Sub

''' <summary>   Creates a clone of the this macro info. </summary>
Public Function Clone() As MacroInfo
    
    Set Clone = New MacroInfo
    Clone.Initialize Me.ProjectName, Me.ModuleName, Me.MacroName

End Function

''' <summary>   Gets the macro name. </summary>
Public Property Get MacroName() As String
    
    MacroName = m_macroName

End Property

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
Attribute FullName.VB_Description = "Get the macro full name"
Attribute FullName.VB_UserMemId = 0
    
    FullName = m_projectName & "." & m_moduleName & "." & m_macroName

End Property

''' <summary>   Checks if the two macros are equal referring to the same
''  macro in the project. </summary>
''' <param name="macro">   The macro to check for equality. </param>
Public Function Equals(ByVal macro As MacroInfo) As Boolean
    
    If IsObject(macro) Then
        Equals = (0 = StrComp(Me.ProjectName, macro.ProjectName)) And _
                 (0 = StrComp(Me.ModuleName, macro.ModuleName)) And _
                 (0 = StrComp(Me.MacroName, macro.MacroName))
    Else
        Equals = False
    End If

End Function


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
Public Sub Initialize(ByVal projectName As String, ByVal codeModuleName As String, ByVal macroName As String)
    m_projectName = projectName
    m_moduleName = codeModuleName
    m_macroName = macroName
End Sub

''' <summary>   Initializes this class. </summary>
''' <param name="project">     A reference to the module info. </param>
''' <param name="macroName">   The name of the macro. </param>
Public Sub InitializeModuelMacro(ByRef module As ModuleInfo, ByVal macroName As String)
    m_projectName = module.projectName
    m_moduleName = module.moduleName
    m_macroName = macroName
End Sub

''' <summary>   Initializes this class from a macro full name. </summary>
''' <param name="macroFullName">     the macro full name. </param>
Public Sub FromMacroFullName(ByVal macroFullName As String)
    Dim names() As String: names() = Split(macroFullName, ".")
    m_projectName = names(LBound(names))
    m_moduleName = names(LBound(names) + 1)
    m_macroName = names(LBound(names) + 2)
End Sub

''' <summary>   Creates a clone of the this macro info. </summary>
Public Function Clone() As MacroInfo
    Dim newMacro As MacroInfo: Set newMacro = New MacroInfo
    newMacro.Initialize Me.projectName, Me.moduleName, Me.macroName
End Function

''' <summary>   Gets the macro name. </summary>
Public Property Get macroName() As String
    macroName = m_macroName
End Property

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
    FullName = m_projectName & "." & m_moduleName & "." & m_macroName
End Property

''' <summary>   Checks if the two macros are equal refering to the same
''  macro in the project. </summary>
''' <param name="macro">   The macro to check for equality. </param>
Public Function Equals(ByRef macro As MacroInfo) As Boolean
    If IsObject(macro) Then
        Equals = StrComp(Me.projectName, macro.projectName) And _
                 StrComp(Me.moduleName, macro.moduleName) And _
                 StrComp(Me.macroName, macro.macroName)
    Else
        Equals = False
    End If
End Function



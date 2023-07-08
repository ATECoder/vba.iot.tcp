VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Constructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private m_k2700 As K2700

''' <summary>   Creates a new instance of the <see cref="K2700"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="K2700"/> class. </returns>
Public Function CreateK2700() As K2700
    Set CreateK2700 = New K2700
End Function

''' <summary>   Gets an instance of the <see cref="ResistanceNumberEventArgs"/> class. </summary>
''' <para name="a_resistanceNo">   [integer] The resistance number. </param>
''' <para name="a_resistance">     [double] The resistance. </param>
''' <value>   An instance of the <see cref="ResistanceNumberEventArgs"/> class. </returns>
Public Property Get CreateResistanceNumberEventArgs(ByVal a_resistanceNo As Integer, ByVal a_resistance As Double) As ResistanceNumberEventArgs
    Set CreateResistanceNumberEventArgs = New ResistanceNumberEventArgs
    CreateResistanceNumberEventArgs.Initialize a_resistanceNo, a_resistance
End Property



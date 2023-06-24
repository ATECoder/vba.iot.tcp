VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResistanceNumberEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_resistanceNumber As Integer
Private m_resistance As Double

''' <summary>   Initializes the event arguments. </summary>
''' <para name="resistanceNo">   [integer] The resistance number. </param>
''' <para name="resistance">     [double] The resistance. </param>
Public Sub Initialize(ByVal resistanceNo As Integer, ByVal resistanceValue As Double)
    m_resistanceNumber = resistanceNo
    m_resistance = Resistance
End Sub

''' <summary>   Gets the resistance number. </summary>
''' <value>   A integer. </value>
Public Property Get ResistanceNumber() As Integer
    ResistanceNumber = m_resistanceNumber
End Property

''' <summary>   Gets the resistance. </summary>
''' <value>   A double. </value>
Public Property Get Resistance() As Double
    Resistance = m_resistance
End Property


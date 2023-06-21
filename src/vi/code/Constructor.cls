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
Private m_K2700 As K2700

''' <summary>   Creates a new instance of the <see cref="K2700"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="K2700"/> class. </returns>
Public Function CreateK2700() As K2700
    Set CreateK2700 = New K2700
End Function

''' <summary>   Gets a singleton instance of the <see cref="K2700"/> class. </summary>
''' <value>   A singleton instance of the <see cref="K2700"/> class. </returns>
Public Property Get K2700() As K2700
    If m_K2700 Is Nothing Then Set m_K2700 = Me.CreateK2700()
    Set K2700 = m_K2700
End Property

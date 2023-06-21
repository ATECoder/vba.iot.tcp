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

''' <summary>   Creates a new instance of the <see cref="Ieee488Session"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="Ieee488Session"/> class. </returns>
Public Function CreateIeee488Session() As Ieee488Session
    Set CreateIeee488Session = New Ieee488Session
End Function

''' <summary>   Creates a new instance of the <see cref="ViSession"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="ViSession"/> class. </returns>
Public Function CreateViSession() As ViSession
    Set CreateViSession = New ViSession
End Function


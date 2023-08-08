VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' <summary>   Instantiates the<see cref="Ieee488Session"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="Ieee488Session"/>]. </returns>
Public Function NewIeee488Session() As Ieee488Session
    Set NewIeee488Session = New Ieee488Session
End Function

''' <summary>   Instantiates the<see cref="ViSession"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="ViSession"/>]. </returns>
Public Function NewViSession() As ViSession
    Set NewViSession = New ViSession
End Function

''' <summary>   Instantiates the<see cref="GpibLanController"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="GpibLanController"/>]. </returns>
Public Function NewGpibLanController() As GpibLanController
    Set NewGpibLanController = New GpibLanController
End Function


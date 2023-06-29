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

''' <summary>   A public creatable singleton class for creating instances of this workbook
''' classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Creates a new instance of the <see cref="MacroInfo"/> class to share
''' with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="MacroInfo"/> class. </returns>
Public Function CreateMacroInfo() As MacroInfo
    Set CreateMacroInfo = New MacroInfo
End Function

''' <summary>   Creates a new instance of the <see cref="ModuleInfo"/> class to share
''' with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="ModuleInfo"/> class. </returns>
Public Function CreateModuleInfo() As ModuleInfo
    Set CreateModuleInfo = New ModuleInfo
End Function

''' <summary>   Creates an instance of the <see cref="stdTimer"/> class. </summary>
''' <param name="iMilliseconds">   [Long] Number of milliseconds to raise Tick Event between </param>
''' <param name="callback">        [stdICallable=Nothing] Callback object. Must contain a Run() method. </param>
''' <returns> {stdTimer} The frequency isn't guaranteed but it will be as close as humanly possible. </returns>
''' <returns>   A new instance of the <see cref="stdTimer"/> class. </returns>
Public Function CreateStdTimer(ByVal iMilliseconds As Long, Optional ByVal callback As Object = Nothing) As stdTimer
  Set CreateStdTimer = New stdTimer
  Call CreateStdTimer.protInit(iMilliseconds, callback)
End Function

''' <summary>   Creates a new instance of the <see cref="StopWatch"/> class to share
''' with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="StopWatch"/> class. </returns>
Public Function CreateStopWatch() As StopWatch
    Set CreateStopWatch = New StopWatch
End Function

''' <summary>   Creates a new instance of the <see cref="StringBuilder"/> class to share
''' with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="StringBuilder"/> class. </returns>
Public Function CreateStringBuilder() As StringBuilder
    Set CreateStringBuilder = New StringBuilder
End Function

''' <summary>   Creates a new instance of the <see cref="CancelEventArgs"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="CancelEventArgs"/> class. </returns>
Public Function CreateCancelEventArgs() As CancelEventArgs
    Set CreateCancelEventArgs = New CancelEventArgs
End Function




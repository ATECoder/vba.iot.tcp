VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "User32Api"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <summary>   Displays a modal dialog box that contains a system icon, a set of buttons,
''' and a brief application-specific message, such as status or error information. The message
'''  box returns an integer value that indicates which button the user clicked. </summary>
''' <param name="a_handle">    in, optional] A handle to the owner window of the message box to be created.
'''                            If this parameter is NULL, the message box has no owner window. </param>
''' <param name="a_message">   [in, optional] The message to be displayed. If the string consists of more than
'''                            one line, you can separate the lines using a carriage return and/or linefeed character between each line. </param>
''' <param name="a_Caption">   [in, optional] The dialog box title. If this parameter is NULL, the default
'''                            title is Error. </param>
''' <param name="a_style">     [in] The contents and behavior of the dialog box. This parameter can be a combination
'''                            of flags from the following groups of flags. </param>
''' <returns> An integer value that indicates which button the user clicked. </returns>
Private Declare PtrSafe Function MessageBox_ Lib "user32" Alias "MessageBoxA" ( _
        ByVal a_handle As Long, ByVal a_message As String, _
        ByVal a_caption As String, ByVal a_style As Long) As Long

''' <summary>   The contents and behavior of the dialog box. This parameter can be a combination of flags from the following groups of flags. </summary>
''' <param name="a_handle">      [in, optional] A handle to the window to be associated with the timer.
'''                              This window must be owned by the calling thread. If a NULL value for hWnd is passed in along with an nIDEvent
'''                              of an existing timer, that timer will be replaced in the same way that an existing non-NULL hWnd timer will be. </param>
''' <param name="a_id">          [in] A nonzero timer identifier. If the handle parameter is NULL, and the nIDEvent does not match an
'''                              existing timer then it is ignored and a new timer ID is generated. If the hWnd parameter is not NULL
'''                              and the window specified by hWnd already has a timer with the value nIDEvent, then the existing timer
'''                              is replaced by the new timer. When SetTimer replaces a timer, the timer is reset. Therefore, a message
'''                              will be sent after the current time-out value elapses, but the previously set time-out value is ignored.
'''                              If the call is not intended to replace an existing timer, nIDEvent should be 0 if the hWnd is NULL. </param>
''' <param name="a_timeoutMs">   [in] The time-out value, in milliseconds.
'''                              If uElapse is less than USER_TIMER_MINIMUM (0x0000000A), the timeout is set to USER_TIMER_MINIMUM.
'''                              If uElapse is greater than USER_TIMER_MAXIMUM (0x7FFFFFFF), the timeout is set to USER_TIMER_MAXIMUM. </param>
''' <param name="a_callback">    A pointer to the function to be notified when the time-out value elapses. For more information about the function,
'''                              see TimerProc. If lpTimerFunc is NULL, the system posts a WM_TIMER message to the application queue.
'''                              The hwnd member of the message's MSG structure contains the value of the hWnd parameter. </param>
''' <returns>  If the function succeeds and the hWnd parameter is NULL, the return value is an integer identifying the new timer.
''' An application can pass this value to the KillTimer function to destroy the timer.
''' If the function succeeds and the hWnd parameter is not NULL, then the return value is a nonzero integer.
''' An application can pass the value of the nIDEvent parameter to the KillTimer function to destroy the timer.
''' If the function fails to create a timer, the return value is zero. To get extended error information, call GetLastError. </returns>
Private Declare PtrSafe Function SetTimer_ Lib "user32" Alias "SetTimer" ( _
    ByVal a_handle As Long, _
    ByVal a_id As Long, _
    ByVal a_timeoutMs As Long, _
    ByVal a_callback As Any) As Long

''' <summary>   lpTimerFunc. </summary>
''' <param name="a_handle">   [in,optional]A handle to the window associated with the specified timer.
'''                           This value must be the same as the hWnd value passed to the SetTimer function that created the timer. </param>
''' <param name="a_id">       [in] The timer to be destroyed. If the window handle passed to SetTimer is valid,
'''                           this parameter must be the same as the nIDEvent value passed to SetTimer.
'''                           If the application calls SetTimer with hWnd set to NULL, this parameter must be the timer identifier
'''                           returned by SetTimer. </param>
''' <returns>  A Boolean. If the function succeeds, the return value is nonzero.
''' If the function fails, the return value is zero. To get extended error information, call GetLastError. </returns>
Private Declare PtrSafe Function KillTimer_ Lib "user32" Alias "KillTimer" ( _
    ByVal a_handle As Long, _
    ByVal s_id As Long) As Long


''' <summary>   Displays a modal dialog box that contains a system icon, a set of buttons,
''' and a brief application-specific message, such as status or error information. The message
'''  box returns an integer value that indicates which button the user clicked. </summary>
''' <param name="a_ownerWindow">   [in, optional] A handle to the owner window of the message box to be created.
'''                                If this parameter is NULL, the message box has no owner window. </param>
''' <param name="a_message">       [in, optional] The message to be displayed. If the string consists of more than
'''                                one line, you can separate the lines using a carriage return and/or linefeed character between each line. </param>
''' <param name="a_caption">       [in, optional] The dialog box title. If this parameter is NULL, the default
'''                                title is Error. </param>
''' <param name="a_behavior">      [in] The contents and behavior of the dialog box. This parameter can be a combination
'''                                 of flags from the following groups of flags. </param>
''' <returns> An integer value that indicates which button the user clicked. </returns>
Public Function MessageBox(ByVal a_ownerWindow As Long, ByVal a_message As String, ByVal a_caption As String, _
        ByVal a_behavior As Long) As Long
    MessageBox = MessageBox_(a_ownerWindow, a_message, a_caption, a_behavior)
End Function
    
''' <summary>   Displays a modeless message box allowing other events to proceed. </summary>
''' <param name="a_message">   [String] A String expression that specifies the message. </param>
''' <param name="a_caption">   [String] The dialog box title. </param>
''' <param name="a_style">     [Optional, VbMsgBoxStyle, VbOKOnly] The message dialog
'''                            style that specifies the buttons and icons which to
'''                            display with the message. </param>
''' <returns> A message dialog result that indicates which button the user clicked. </returns>
Public Function DisplayMessageBox(ByVal a_message As String, a_caption, _
                                  Optional ByVal a_style As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult

    Dim p_ownerWindow As Long
    
    ' get the handle assigned by Windows to the main Microsoft Access window.
    
    p_ownerWindow = Application.HWnd
    
    ' display the message box using the Windows API.
    
    DisplayMessageBox = MessageBox(p_ownerWindow, a_message, a_caption, a_style)

End Function




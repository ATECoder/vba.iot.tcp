VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const thisModuleName As String = "K2700"
Private Const Mux40Cards As String = "7702,7708"
Private Const Mux20Cards As String = "7700,7710"

' reference to the IEEE488 session
Private m_session As Ieee488Session

Private m_hasMux40A As Boolean
Private m_hasMux40B As Boolean
Private m_hasMux20A As Boolean
Private m_hasMux20B As Boolean

''' <summary>   Initializes the k2700 . </summary>
Private Sub Class_Initialize()
    Debug.Print "initializing k2700"
End Sub

''' <summary>   Closes and releases the VI Session. </summary>
Public Sub Dispose()
    
    If Not m_session Is Nothing Then
        m_session.Dispose
    End If
    Set m_session = Nothing

End Sub

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the TCP Client socket is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If m_session Is Nothing Then
        Connected = False
    Else
        Connected = m_session.Connected
    End If

End Property

''' <summary>   Returns a reference to the IEEE488 Session. </summary>
''' <value>   An IEEE488Session. </returns>
Public Property Get Session() As Ieee488Session
    Set Session = m_session
End Property

''' <summary>   Initializes the k2700. </summary>
'''< remarks>   This is a singleton class.
''' </remarks>
''' <returns>   If successful, returns zero. Otherwise, it returns the extended error code
''' in the return value for this function.
''' </returns>
Public Function Initialize(ByVal host As String, ByVal port As Long, _
        Optional ByVal timeoutMs As Long = 500) As Boolean

    Set m_session = New Ieee488Session

    m_session.OpenConnection host, port, timeoutMs

    ' this only works with the Prologix device
    Me.Session.SelectiveDeviceClear
    Me.Session.ResetKnownState
    Me.Session.ClearExecutionState
    
    Dim availableCards As String: availableCards = Me.Session.QueryOptions
    
    Dim card1 As String: card1 = StringExtensions.Pop(availableCards, ",")
    Dim card2 As String: card2 = StringExtensions.Pop(availableCards, ",")
    
    Dim card As Variant
    For Each card In Split(Mux20Cards, ",")
        m_hasMux20A = InStr(1, card1, CStr(card), vbTextCompare)
        m_hasMux20B = InStr(1, card2, CStr(card), vbTextCompare)
    Next card
    
    For Each card In Split(Mux40Cards, ",")
        m_hasMux40A = InStr(1, card1, CStr(card), vbTextCompare)
        m_hasMux40B = InStr(1, card2, CStr(card), vbTextCompare)
    Next card
    
    If Me.HasMux40A Then
       Me.Session.WriteLine ":FUNC 'RES',(@101:140)"
    End If
    If Me.HasMux40B Then
      Me.Session.WriteLine ":FUNC 'RES',(@201:240)"
    End If
    If Me.HasMux20A Then
      Me.Session.WriteLine ":FUNC 'RES',(@101:120)"
    End If
    If Me.HasMux20B Then
      Me.Session.WriteLine ":FUNC 'RES',(@201:220)"
    End If

    ' set resistance defaults
    Me.Session.WriteLine ":RES:RANG:AUTO ON"
    Me.Session.WriteLine ":RES:NPLC 1"
    
   ' set reading format
    Me.Session.WriteLine ":FORM:ELEM READ"
    
    ' turn on continuous mode
    Me.Session.WriteLine ":FUNC 'FRES'"
    Me.Session.WriteLine ":FRES:RANG:AUTO ON"
    Me.Session.WriteLine ":FRES:NPLC 1"
    Me.Session.WriteLine ":INIT:CONT On"
    
    Initialize = True
    
End Function


''' <summary>   Returns true if the instrument as a 40 channel multiplexer
''' such as the 7702 board in the first position. </summary>
''' <value> A Boolean. </value>
Public Property Get HasMux40A() As Boolean
        
        HasMux40A = m_hasMux40A

End Property

''' <summary>   Returns true if the instrument as a 40 channel multiplexer
''' such as the 7702 board in the second position. </summary>
''' <value> A Boolean. </value>
Public Property Get HasMux40B() As Boolean

        HasMux40B = m_hasMux40B

End Property

''' <summary>   Returns true if the instrument as a 20 channel multiplexer
''' such as the 7700 or 7710 boards in the first position. </summary>
''' <value> A Boolean. </value>
Public Property Get HasMux20A() As Boolean

        HasMux20A = m_hasMux20A

End Property

''' <summary>   Returns true if the instrument as a 20 channel multiplexer
''' such as the 7700 or 7710 boards in the second position. </summary>
''' <value> A Boolean. </value>
Public Property Get HasMux20B() As Boolean

        HasMux20B = m_hasMux20B

End Property


''' <summary>   Returns the channel list. </summary>
''' <param name="resistanceNumber">   The resistance number. </param>
''' <returns>  The channel list to close. </returns>
Public Function BuildChannelList(ByVal resistanceNumber As Integer) As String

    Const thisProcedureName = "BuildChannelList"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim routeCommand As String
    
    ' Check if we need to close the bottom or top
    ' channle.
    If resistanceNumber <= 20 Then
    
        If Me.HasMux40A Then
          
            routeCommand = ":ROUT:MULT:CLOS (@144,145)"
        
        ElseIf Me.HasMux20A Then
          
            routeCommand = ":ROUT:MULT:CLOS (@124,125)"
        
        End If
    
    ElseIf resistanceNumber <= 40 Then
      
        If Me.HasMux40A Then
        
            ' Get channel string
            routeCommand = ":ROUT:MULT:CLOS (@144,145)"
        
        ElseIf Me.HasMux20B Then
          
            ' Get channel string
            routeCommand = ":ROUT:MULT:CLOS (@224,225)"
        
        End If
      
    ElseIf resistanceNumber > 40 Then
    
        If Me.HasMux40B Then
            routeCommand = ":ROUT:MULT:CLOS (@244,245)"
        End If
    
    End If
        
    ' set the return value
    BuildChannelList = routeCommand

' ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` `
exit_Handler:
    Exit Function
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, thisModuleName
    
    ' raise error
    Err.Raise Err.Number

End Function

''' <summary>   Returns the scan list for a specific resistance. </summary>
''' <param name="resistanceNumber">   The resistance number. </param>
''' <returns>  The scan list. </returns>
Public Function BuildScanList(ByVal resistanceNumber As Integer) As String

    Const thisProcedureName = "BuildScanList"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim scanList As String
    Dim resistaceChannel As String
    
    ' Check if we need to close the bottom or top
    ' channle.
    If resistanceNumber <= 20 Then
    
        ' Get channel string
        If Me.HasMux40A Or Me.HasMux20A Then
        
            resistaceChannel = "1" & Right$("0" & Format$(resistanceNumber), 2)
            scanList = "(@" & resistaceChannel & ")"
        
        End If
      
    ElseIf resistanceNumber <= 40 Then
      
        If Me.HasMux40A Then
        
            ' Get channel string
            resistaceChannel = "1" & Right$("0" & Format$(resistanceNumber), 2)
            scanList = "(@" & resistaceChannel & ")"
        
        ElseIf Me.HasMux20B Then
          
            ' Get channel string
            resistaceChannel = "2" & Right$("0" & Format$(resistanceNumber - 20), 2)
            
            scanList = "(@" & resistaceChannel & ")"
        
        End If
      
    ElseIf resistanceNumber > 40 Then
    
        ' Get channel string
        resistaceChannel = "2" & Right$("0" & Format$(resistanceNumber - 40), 2)
        
        If Me.HasMux40B Then
            scanList = "(@" & resistaceChannel & ")"
        End If
    
    End If
        
    ' set the return value
    BuildScanList = scanList

' ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` ` `
exit_Handler:
    Exit Function
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    UserDefinedErrors.AppendErrSource thisProcedureName, thisModuleName
    
    ' raise error
    Err.Raise Err.Number

End Function



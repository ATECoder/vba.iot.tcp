VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const thisModuleName As String = "K2700"
Private Const Mux40Cards As String = "7702,7708"
Private Const Mux20Cards As String = "7700,7710"

' reference to the IEEE488 session
Private m_session As Ieee488Session

Private m_options As String
Private m_hasMux40A As Boolean
Private m_hasMux40B As Boolean
Private m_hasMux20A As Boolean
Private m_hasMux20B As Boolean
Private m_card1 As String
Private m_card2 As String
Private m_card1ScanList As String
Private m_card2ScanList As String

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

''' <summary>   Opens a TCP Client connection to the specified host at the specified port
''' as sets the connection timeout. </summary>
''' <param name="host">        An IPv4 dotted-decimal host address. </param>
''' <param name="port">        The port that the server is listening on. </param>
''' <param name="timeoutMs">   [Optional, 500] The receive timeout in milliseconds. </param>
''' <returns>   True if the TCP Client is connected. </returns>
Public Function OpenConnection(ByVal host As String, ByVal port As Long, _
        Optional ByVal timeoutMs As Long = 500) As Boolean
       
    Set m_session = cc_isr_Ieee488.Constructor.CreateIeee488Session

    OpenConnection = m_session.OpenConnection(host, port, timeoutMs)

End Function

''' <summary>   Closes the TCP Client connection. </summary>
''' <returns>   True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    CloseConnection = m_session.CloseConnection
    
End Function

''' <summary>   Initializes the k2700. </summary>
'''< remarks>   This is a singleton class.
''' </remarks>
''' <param name="host">        An IPv4 dotted-decimal host address. </param>
''' <param name="port">        The port that the server is listening on. </param>
''' <param name="timeoutMs">   [Optional, 500] The receive timeout in milliseconds. </param>
''' <returns>   True if the device was initialized. </returns>
Public Function Initialize(ByVal host As String, ByVal port As Long, _
        Optional ByVal timeoutMs As Long = 500) As Boolean

    Me.OpenConnection host, port, timeoutMs

    ' this only works with the Prologix device
    Me.Session.SelectiveDeviceClear
    Me.Session.ResetKnownState
    Me.Session.ClearExecutionState
    
    ' read the instrument options and parse the card names and types.
    
    Me.ReadCards
    
    ' define scan list for resistance measurments on all cards.
   
    Me.DefineResistanceScanLists
   
    ' turn on continuous resistance auto range
    Me.ContinuousResistanceAutoRange
    
    Initialize = True
    
End Function

''' <summary>   Gets the options value and sets the cards. </summary>
Public Function ReadCards()

    m_options = Me.Session.QueryOptions
    
    Me.Card1 = StringExtensions.Pop(m_options, ",")
    Me.Card2 = StringExtensions.Pop(m_options, ",")
    
End Function

''' <summary>   Sets the complete scan lists for resistance measurements on all cards. </summary>
Public Function DefineResistanceScanLists()

    If Me.HasMux40A Then
        m_card1ScanList = ":FUNC 'RES',(@101:140)"
        Me.Session.WriteLine m_card1ScanList
    End If
    
    If Me.HasMux40B Then
        m_card2ScanList = ":FUNC 'RES',(@201:240)"
        Me.Session.WriteLine m_card2ScanList
    End If
    
    If Me.HasMux20A Then
        m_card1ScanList = ":FUNC 'RES',(@101:120)"
        Me.Session.WriteLine m_card1ScanList
    End If
    
    If Me.HasMux20B Then
        m_card2ScanList = ":FUNC 'RES',(@201:220)"
        Me.Session.WriteLine m_card2ScanList
    End If

End Function

''' <summary>   Sets the instrument to continous auto range resistance measurement. </summary>
Public Function ContinuousResistanceAutoRange()

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

End Function

''' <summary>   Gets the options value. </summary>
''' <value> A String. </value>
Public Property Get Options() As String
        
    Options = m_options

End Property

''' <summary>   Gets the top card. </summary>
''' <value> A String. </value>
Public Property Get Card1() As String
        
    Card1 = m_card1

End Property

''' <summary>   Sets the top card. </summary>
''' <param name="value">   The name of the top card. </param>
''' <value> A String. </value>
Public Property Let Card1(ByVal value As String)
    m_card1 = value
    m_hasMux20A = 0 < InStr(1, Mux20Cards, value, vbTextCompare)
    m_hasMux40A = 0 < InStr(1, Mux40Cards, value, vbTextCompare)
End Property

''' <summary>   Gets the top card scan list. </summary>
''' <value> A String. </value>
Public Property Get Card1ScanList() As String
        
    Card1ScanList = m_card1ScanList

End Property

''' <summary>   Gets the bottom card. </summary>
''' <value> A String. </value>
Public Property Get Card2() As String
        
    Card2 = m_card2

End Property

''' <summary>   Sets the bottom card. </summary>
''' <param name="value">   The name of the bottom card. </param>
''' <value> A String. </value>
Public Property Let Card2(ByVal value As String)
    m_card2 = value
    m_hasMux40B = 0 < InStr(1, Mux40Cards, value, vbTextCompare)
    m_hasMux20B = 0 < InStr(1, Mux20Cards, value, vbTextCompare)
End Property

''' <summary>   Gets the bottom card scan list. </summary>
''' <value> A String. </value>
Public Property Get Card2ScanList() As String
        
    Card2ScanList = m_card2ScanList

End Property

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

